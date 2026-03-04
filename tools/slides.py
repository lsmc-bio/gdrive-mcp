"""Google Slides tools for gdrive-mcp."""

import uuid
from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from services import get_slides


def register(mcp):
    """Register all Google Slides tools with the MCP server."""

    # ── Helper: resolve slide number to objectId ──────────────────────────

    def _get_presentation(file_id: str) -> dict:
        """Fetch presentation metadata."""
        return get_slides().presentations().get(presentationId=file_id).execute()

    def _slide_number_to_id(presentation: dict, slide_number: int) -> str | None:
        """Convert 1-based slide number to objectId."""
        slides = presentation.get("slides", [])
        if 1 <= slide_number <= len(slides):
            return slides[slide_number - 1].get("objectId")
        return None

    def _batch_update(file_id: str, requests: list) -> dict:
        """Execute a batchUpdate on a presentation."""
        return get_slides().presentations().batchUpdate(
            presentationId=file_id,
            body={"requests": requests},
        ).execute()

    # ── Tool: Read Slides ────────────────────────────────────────────────

    class ReadSlidesInput(BaseModel):
        """Input for reading a Google Slides presentation."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_numbers: Optional[List[int]] = Field(
            default=None,
            description="Specific slide numbers to read (1-based). If omitted, reads all slides.",
        )

    def extract_slide_text(slide: dict) -> str:
        """Extract text content from a slide's page elements."""
        texts = []
        for element in slide.get("pageElements", []):
            shape = element.get("shape", {})
            text_content = shape.get("text", {})
            for text_element in text_content.get("textElements", []):
                text_run = text_element.get("textRun", {})
                content = text_run.get("content", "")
                if content.strip():
                    style = text_run.get("style", {})
                    if style.get("bold"):
                        content = f"**{content.strip()}**"
                    if style.get("italic"):
                        content = f"*{content.strip()}*"
                    if style.get("link", {}).get("url"):
                        content = f"[{content.strip()}]({style['link']['url']})"
                    texts.append(content)

            # Handle tables in slides
            table = element.get("table", {})
            if table:
                md_rows = []
                for row in table.get("tableRows", []):
                    cells = []
                    for cell in row.get("tableCells", []):
                        cell_text = ""
                        for tc in cell.get("text", {}).get("textElements", []):
                            cell_text += tc.get("textRun", {}).get("content", "")
                        cells.append(cell_text.strip().replace("|", "\\|"))
                    md_rows.append("| " + " | ".join(cells) + " |")
                if md_rows:
                    num_cols = len(table.get("tableRows", [{}])[0].get("tableCells", []))
                    md_rows.insert(1, "| " + " | ".join(["---"] * num_cols) + " |")
                    texts.append("\n".join(md_rows))

        return "\n".join(texts)

    @mcp.tool(
        name="gdrive_read_slides",
        annotations={"title": "Read Google Slides", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_read_slides(params: ReadSlidesInput) -> str:
        """Read a Google Slides presentation and return content as markdown.

        Extracts text, formatting, tables, and speaker notes from each slide.
        Optionally read specific slides by number.

        Args:
            params: File ID and optional slide numbers.

        Returns:
            Presentation content as markdown with slide separators.
        """
        try:
            presentation = get_slides().presentations().get(presentationId=params.file_id).execute()
        except Exception as e:
            return f"Error reading presentation: {e}"

        title = presentation.get("title", "Untitled Presentation")
        slides = presentation.get("slides", [])
        total = len(slides)

        if not slides:
            return f"Presentation '{title}' has no slides."

        output = f"# {title}\n\n*{total} slides*\n\n"

        for i, slide in enumerate(slides, 1):
            if params.slide_numbers and i not in params.slide_numbers:
                continue

            output += f"---\n\n## Slide {i}\n\n"

            # Slide content
            content = extract_slide_text(slide)
            if content.strip():
                output += content + "\n\n"
            else:
                output += "*[Empty slide]*\n\n"

            # Speaker notes
            notes_page = slide.get("slideProperties", {}).get("notesPage", {})
            for element in notes_page.get("pageElements", []):
                shape = element.get("shape", {})
                if shape.get("shapeType") == "TEXT_BOX":
                    notes_text = ""
                    for te in shape.get("text", {}).get("textElements", []):
                        notes_text += te.get("textRun", {}).get("content", "")
                    if notes_text.strip():
                        output += f"**Speaker Notes:** {notes_text.strip()}\n\n"

        return output

    # ── Tool: Get Slide Elements ──────────────────────────────────────────

    class GetElementsInput(BaseModel):
        """Input for getting detailed element info from a slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_number: int = Field(..., description="1-based slide number to inspect.", ge=1)

    @mcp.tool(
        name="gdrive_slide_get_elements",
        annotations={"title": "Get Slide Elements", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_slide_get_elements(params: GetElementsInput) -> str:
        """Get detailed element info (IDs, types, text, positions) for a specific slide.

        Returns each element's objectId, type, position, size, and text content.
        Use this to discover element IDs before editing them.

        Args:
            params: File ID and slide number.

        Returns:
            Detailed element listing for the slide.
        """
        try:
            presentation = _get_presentation(params.file_id)
        except Exception as e:
            return f"Error reading presentation: {e}"

        slide_id = _slide_number_to_id(presentation, params.slide_number)
        if not slide_id:
            total = len(presentation.get("slides", []))
            return f"Slide {params.slide_number} not found. Presentation has {total} slides."

        slide = presentation["slides"][params.slide_number - 1]
        elements = slide.get("pageElements", [])

        if not elements:
            return f"Slide {params.slide_number} (objectId: {slide_id}) has no elements."

        output = f"## Slide {params.slide_number} (objectId: `{slide_id}`)\n\n"
        output += f"{len(elements)} elements:\n\n"

        for elem in elements:
            eid = elem.get("objectId", "unknown")
            transform = elem.get("transform", {})
            size = elem.get("size", {})

            # Determine element type and extract text
            if "shape" in elem:
                shape = elem["shape"]
                shape_type = shape.get("shapeType", "SHAPE")
                placeholder = shape.get("placeholder", {})
                ph_type = placeholder.get("type", "")

                text_parts = []
                for te in shape.get("text", {}).get("textElements", []):
                    run = te.get("textRun", {})
                    if run.get("content", "").strip():
                        text_parts.append(run["content"].strip())
                text = " | ".join(text_parts) if text_parts else "(empty)"

                type_label = f"Shape/{shape_type}"
                if ph_type:
                    type_label += f" [placeholder: {ph_type}]"

                output += f"- **`{eid}`** — {type_label}\n"
                output += f"  Text: {text[:200]}{'...' if len(text) > 200 else ''}\n"

            elif "table" in elem:
                table = elem["table"]
                rows = len(table.get("tableRows", []))
                cols = len(table.get("tableRows", [{}])[0].get("tableCells", [])) if rows > 0 else 0
                output += f"- **`{eid}`** — Table ({rows}x{cols})\n"

            elif "image" in elem:
                output += f"- **`{eid}`** — Image\n"

            elif "line" in elem:
                output += f"- **`{eid}`** — Line\n"

            elif "group" in elem:
                children = len(elem["group"].get("children", []))
                output += f"- **`{eid}`** — Group ({children} children)\n"

            else:
                output += f"- **`{eid}`** — Unknown element\n"

            # Position info
            if size:
                w = size.get("width", {})
                h = size.get("height", {})
                w_pt = w.get("magnitude", 0)
                h_pt = h.get("magnitude", 0)
                output += f"  Size: {w_pt:.0f}x{h_pt:.0f} {w.get('unit', 'EMU')}\n"

        return output

    # ── Tool: Replace Text in Slides ──────────────────────────────────────

    class SlideReplaceTextInput(BaseModel):
        """Input for find-and-replace text in slides."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        find_text: str = Field(..., description="The text to find.", min_length=1)
        replace_text: str = Field(..., description="The text to replace it with. Use empty string to delete.")
        slide_numbers: Optional[List[int]] = Field(
            default=None,
            description="Specific slide numbers (1-based) to scope the replacement. If omitted, replaces across ALL slides.",
        )
        match_case: bool = Field(default=True, description="If true (default), search is case-sensitive.")

    @mcp.tool(
        name="gdrive_slide_replace_text",
        annotations={"title": "Replace Text in Slides", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_slide_replace_text(params: SlideReplaceTextInput) -> str:
        """Find and replace text in a presentation, optionally scoped to specific slides.

        Args:
            params: File ID, find/replace text, optional slide numbers, case sensitivity.

        Returns:
            Number of occurrences replaced.
        """
        try:
            page_object_ids = None
            if params.slide_numbers:
                presentation = _get_presentation(params.file_id)
                page_object_ids = []
                for num in params.slide_numbers:
                    sid = _slide_number_to_id(presentation, num)
                    if sid:
                        page_object_ids.append(sid)
                    else:
                        return f"Slide {num} not found."

            request = {
                "replaceAllText": {
                    "containsText": {
                        "text": params.find_text,
                        "matchCase": params.match_case,
                    },
                    "replaceText": params.replace_text,
                }
            }

            if page_object_ids:
                request["replaceAllText"]["pageObjectIds"] = page_object_ids

            result = _batch_update(params.file_id, [request])
            replies = result.get("replies", [{}])
            count = replies[0].get("replaceAllText", {}).get("occurrencesChanged", 0) if replies else 0
            scope = f"slides {params.slide_numbers}" if params.slide_numbers else "all slides"
            return f"Replaced {count} occurrence(s) of '{params.find_text}' → '{params.replace_text}' in {scope}."

        except Exception as e:
            return f"Error replacing text: {e}"

    # ── Tool: Set Element Text ────────────────────────────────────────────

    class SetElementTextInput(BaseModel):
        """Input for setting text on a specific element."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        element_id: str = Field(..., description="The objectId of the element to update (get from gdrive_slide_get_elements).", min_length=1)
        text: str = Field(..., description="The new text content to set. Replaces all existing text in the element.")
        bold: bool = Field(default=False, description="Apply bold formatting to the text.")
        font_size: Optional[float] = Field(default=None, description="Font size in points (e.g., 14.0).")

    @mcp.tool(
        name="gdrive_slide_set_element_text",
        annotations={"title": "Set Element Text", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_slide_set_element_text(params: SetElementTextInput) -> str:
        """Clear all text in an element and replace it with new text.

        Use gdrive_slide_get_elements first to find the element objectId.

        Args:
            params: File ID, element ID, new text, optional formatting.

        Returns:
            Confirmation of the text update.
        """
        try:
            requests = [
                {
                    "insertText": {
                        "objectId": params.element_id,
                        "insertionIndex": 0,
                        "text": params.text,
                    }
                },
            ]

            # Try delete-then-insert; if element is empty, deleteText fails — just insert
            try:
                _batch_update(params.file_id, [
                    {"deleteText": {"objectId": params.element_id, "textRange": {"type": "ALL"}}},
                ] + requests)
            except Exception as delete_err:
                if "startIndex" in str(delete_err) and "endIndex" in str(delete_err):
                    # Element was empty — just insert
                    _batch_update(params.file_id, requests)
                else:
                    raise delete_err

            # Optional formatting
            if params.bold or params.font_size:
                style = {}
                fields = []
                if params.bold:
                    style["bold"] = True
                    fields.append("bold")
                if params.font_size:
                    style["fontSize"] = {"magnitude": params.font_size, "unit": "PT"}
                    fields.append("fontSize")

                _batch_update(params.file_id, [{
                    "updateTextStyle": {
                        "objectId": params.element_id,
                        "textRange": {"type": "ALL"},
                        "style": style,
                        "fields": ",".join(fields),
                    }
                }])

            return f"Set text on element `{params.element_id}`: '{params.text[:100]}{'...' if len(params.text) > 100 else ''}'"

        except Exception as e:
            return f"Error setting element text: {e}"

    # ── Tool: Duplicate Slide ─────────────────────────────────────────────

    class DuplicateSlideInput(BaseModel):
        """Input for duplicating a slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_number: int = Field(..., description="1-based number of the slide to duplicate.", ge=1)
        insert_at: Optional[int] = Field(
            default=None,
            description="1-based position to insert the duplicate. If omitted, inserts right after the original.",
        )

    @mcp.tool(
        name="gdrive_duplicate_slide",
        annotations={"title": "Duplicate Slide", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_duplicate_slide(params: DuplicateSlideInput) -> str:
        """Duplicate a slide within the presentation.

        Creates an exact copy of the specified slide, preserving all formatting,
        elements, and speaker notes. The copy can then be edited independently.

        Args:
            params: File ID, slide number to duplicate, optional insertion position.

        Returns:
            The objectId of the new slide and its position.
        """
        try:
            presentation = _get_presentation(params.file_id)
            source_id = _slide_number_to_id(presentation, params.slide_number)
            if not source_id:
                total = len(presentation.get("slides", []))
                return f"Slide {params.slide_number} not found. Presentation has {total} slides."

            new_id = f"slide_{uuid.uuid4().hex[:12]}"

            requests = [
                {
                    "duplicateObject": {
                        "objectId": source_id,
                        "objectIds": {source_id: new_id},
                    }
                }
            ]

            # If a specific position is requested, move it there after duplication
            insert_pos = params.insert_at
            if insert_pos is not None:
                # duplicateObject places copy right after original.
                # We'll need to move it to the desired position.
                requests.append({
                    "updateSlidesPosition": {
                        "slideObjectIds": [new_id],
                        "insertionIndex": insert_pos - 1,  # API uses 0-based
                    }
                })

            result = _batch_update(params.file_id, requests)
            final_pos = params.insert_at if params.insert_at else params.slide_number + 1
            return f"Duplicated slide {params.slide_number} → new slide objectId: `{new_id}` (position {final_pos})."

        except Exception as e:
            return f"Error duplicating slide: {e}"

    # ── Tool: Create Slide ────────────────────────────────────────────────

    class CreateSlideInput(BaseModel):
        """Input for creating a new slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        insert_at: Optional[int] = Field(
            default=None,
            description="1-based position to insert the new slide. If omitted, appends at the end.",
        )
        layout: str = Field(
            default="BLANK",
            description="Predefined layout: BLANK, TITLE, TITLE_AND_BODY, TITLE_AND_TWO_COLUMNS, TITLE_ONLY, SECTION_HEADER, ONE_COLUMN_TEXT, MAIN_POINT, BIG_NUMBER, CAPTION_ONLY.",
        )

    @mcp.tool(
        name="gdrive_create_slide",
        annotations={"title": "Create Slide", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_create_slide(params: CreateSlideInput) -> str:
        """Create a new slide with a predefined layout.

        The new slide will inherit the presentation's theme/master styling.
        Use gdrive_slide_get_elements after creation to discover placeholder IDs,
        then gdrive_slide_set_element_text to populate them.

        Args:
            params: File ID, optional position, layout type.

        Returns:
            The objectId of the new slide and its position.
        """
        try:
            new_id = f"slide_{uuid.uuid4().hex[:12]}"

            request = {
                "createSlide": {
                    "objectId": new_id,
                    "slideLayoutReference": {
                        "predefinedLayout": params.layout,
                    },
                }
            }

            if params.insert_at is not None:
                request["createSlide"]["insertionIndex"] = params.insert_at - 1  # API uses 0-based

            _batch_update(params.file_id, [request])

            pos = params.insert_at if params.insert_at else "end"
            return f"Created new {params.layout} slide → objectId: `{new_id}` (position: {pos}). Use gdrive_slide_get_elements to find placeholder IDs."

        except Exception as e:
            return f"Error creating slide: {e}"

    # ── Tool: Delete Slide ────────────────────────────────────────────────

    class DeleteSlideInput(BaseModel):
        """Input for deleting a slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_number: int = Field(..., description="1-based number of the slide to delete.", ge=1)

    @mcp.tool(
        name="gdrive_delete_slide",
        annotations={"title": "Delete Slide", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_delete_slide(params: DeleteSlideInput) -> str:
        """Delete a slide from the presentation.

        WARNING: This is destructive and cannot be undone via the API.
        Use version history in Google Slides to recover if needed.

        Args:
            params: File ID and slide number.

        Returns:
            Confirmation of deletion.
        """
        try:
            presentation = _get_presentation(params.file_id)
            slide_id = _slide_number_to_id(presentation, params.slide_number)
            if not slide_id:
                total = len(presentation.get("slides", []))
                return f"Slide {params.slide_number} not found. Presentation has {total} slides."

            _batch_update(params.file_id, [{"deleteObject": {"objectId": slide_id}}])
            return f"Deleted slide {params.slide_number} (objectId: {slide_id})."

        except Exception as e:
            return f"Error deleting slide: {e}"

    # ── Tool: Reorder Slides ──────────────────────────────────────────────

    class ReorderSlidesInput(BaseModel):
        """Input for reordering slides."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_numbers: List[int] = Field(..., description="1-based slide numbers to move.", min_length=1)
        insert_before: int = Field(..., description="1-based position to insert the slides before. Use total+1 to move to end.", ge=1)

    @mcp.tool(
        name="gdrive_reorder_slides",
        annotations={"title": "Reorder Slides", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_reorder_slides(params: ReorderSlidesInput) -> str:
        """Move one or more slides to a new position in the presentation.

        Args:
            params: File ID, slide numbers to move, target position.

        Returns:
            Confirmation of the reorder.
        """
        try:
            presentation = _get_presentation(params.file_id)
            slide_ids = []
            for num in params.slide_numbers:
                sid = _slide_number_to_id(presentation, num)
                if not sid:
                    return f"Slide {num} not found."
                slide_ids.append(sid)

            _batch_update(params.file_id, [{
                "updateSlidesPosition": {
                    "slideObjectIds": slide_ids,
                    "insertionIndex": params.insert_before - 1,  # API uses 0-based
                }
            }])

            return f"Moved slide(s) {params.slide_numbers} to position {params.insert_before}."

        except Exception as e:
            return f"Error reordering slides: {e}"

    # ── Tool: Add Text Box ────────────────────────────────────────────────

    class AddTextBoxInput(BaseModel):
        """Input for adding a text box to a slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_number: int = Field(..., description="1-based slide number to add the text box to.", ge=1)
        text: str = Field(..., description="The text content for the text box.")
        left: float = Field(default=100.0, description="Left position in points from slide edge.")
        top: float = Field(default=100.0, description="Top position in points from slide edge.")
        width: float = Field(default=400.0, description="Width of the text box in points.")
        height: float = Field(default=50.0, description="Height of the text box in points.")
        font_size: Optional[float] = Field(default=None, description="Font size in points (e.g., 14.0).")
        bold: bool = Field(default=False, description="Apply bold formatting.")

    @mcp.tool(
        name="gdrive_slide_add_text_box",
        annotations={"title": "Add Text Box to Slide", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_slide_add_text_box(params: AddTextBoxInput) -> str:
        """Add a new text box with content to a specific slide.

        Positions are in points (72 points = 1 inch). A standard slide is
        720x405 points (10x5.625 inches).

        Args:
            params: File ID, slide number, text, position, size, formatting.

        Returns:
            The objectId of the new text box.
        """
        try:
            presentation = _get_presentation(params.file_id)
            slide_id = _slide_number_to_id(presentation, params.slide_number)
            if not slide_id:
                total = len(presentation.get("slides", []))
                return f"Slide {params.slide_number} not found. Presentation has {total} slides."

            box_id = f"textbox_{uuid.uuid4().hex[:12]}"

            # Points to EMU conversion (1 point = 12700 EMU)
            PT_TO_EMU = 12700

            requests = [
                {
                    "createShape": {
                        "objectId": box_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {
                                "width": {"magnitude": params.width * PT_TO_EMU, "unit": "EMU"},
                                "height": {"magnitude": params.height * PT_TO_EMU, "unit": "EMU"},
                            },
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": params.left * PT_TO_EMU,
                                "translateY": params.top * PT_TO_EMU,
                                "unit": "EMU",
                            },
                        },
                    }
                },
                {
                    "insertText": {
                        "objectId": box_id,
                        "insertionIndex": 0,
                        "text": params.text,
                    }
                },
            ]

            # Optional formatting
            if params.bold or params.font_size:
                style = {}
                fields = []
                if params.bold:
                    style["bold"] = True
                    fields.append("bold")
                if params.font_size:
                    style["fontSize"] = {"magnitude": params.font_size, "unit": "PT"}
                    fields.append("fontSize")

                requests.append({
                    "updateTextStyle": {
                        "objectId": box_id,
                        "textRange": {"type": "ALL"},
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })

            _batch_update(params.file_id, requests)
            return f"Added text box `{box_id}` to slide {params.slide_number}: '{params.text[:80]}{'...' if len(params.text) > 80 else ''}'"

        except Exception as e:
            return f"Error adding text box: {e}"

    # ── Tool: Set Speaker Notes ───────────────────────────────────────────

    class SetSpeakerNotesInput(BaseModel):
        """Input for setting speaker notes on a slide."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
        slide_number: int = Field(..., description="1-based slide number.", ge=1)
        notes: str = Field(..., description="The speaker notes text to set. Replaces any existing notes.")

    @mcp.tool(
        name="gdrive_slide_set_notes",
        annotations={"title": "Set Speaker Notes", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_slide_set_notes(params: SetSpeakerNotesInput) -> str:
        """Set speaker notes on a slide. Replaces any existing notes.

        Speaker notes appear in presenter view and are useful for talking points.

        Args:
            params: File ID, slide number, notes text.

        Returns:
            Confirmation of the notes update.
        """
        try:
            presentation = _get_presentation(params.file_id)
            slides = presentation.get("slides", [])
            if params.slide_number < 1 or params.slide_number > len(slides):
                return f"Slide {params.slide_number} not found. Presentation has {len(slides)} slides."

            slide = slides[params.slide_number - 1]
            notes_page = slide.get("slideProperties", {}).get("notesPage", {})

            # Find the notes text box element
            notes_element_id = None
            for elem in notes_page.get("pageElements", []):
                shape = elem.get("shape", {})
                if shape.get("shapeType") == "TEXT_BOX":
                    ph = shape.get("placeholder", {})
                    if ph.get("type") == "BODY":
                        notes_element_id = elem.get("objectId")
                        break

            if not notes_element_id:
                return f"Could not find notes element on slide {params.slide_number}."

            insert_req = [{"insertText": {"objectId": notes_element_id, "insertionIndex": 0, "text": params.notes}}]

            # Try delete-then-insert; if notes are empty, deleteText fails — just insert
            try:
                _batch_update(params.file_id, [
                    {"deleteText": {"objectId": notes_element_id, "textRange": {"type": "ALL"}}},
                ] + insert_req)
            except Exception as delete_err:
                if "startIndex" in str(delete_err) and "endIndex" in str(delete_err):
                    _batch_update(params.file_id, insert_req)
                else:
                    raise delete_err

            return f"Set speaker notes on slide {params.slide_number}: '{params.notes[:80]}{'...' if len(params.notes) > 80 else ''}'"

        except Exception as e:
            return f"Error setting speaker notes: {e}"
