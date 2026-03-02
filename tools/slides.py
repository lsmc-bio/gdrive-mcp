"""Google Slides tools for gdrive-mcp."""

from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from services import get_slides


def register(mcp):
    """Register all Google Slides tools with the MCP server."""

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
