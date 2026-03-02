"""Apps Script tools for gdrive-mcp."""

import json
from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from services import get_scripts


def register(mcp):
    """Register all Apps Script tools with the MCP server."""

    # ── Tool: Run Script ─────────────────────────────────────────────────

    class RunScriptInput(BaseModel):
        """Input for running an Apps Script function."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(
            ...,
            description="The Apps Script project ID. Find this in the script editor URL or project settings.",
            min_length=1,
        )
        function_name: str = Field(
            ...,
            description="The function name to execute.",
            min_length=1,
        )
        parameters: Optional[List] = Field(
            default=None,
            description="Parameters to pass to the function as a JSON array.",
        )

    @mcp.tool(
        name="gdrive_run_script",
        annotations={"title": "Run Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_run_script(params: RunScriptInput) -> str:
        """Execute a function in a Google Apps Script project.

        The script must be deployed as an API executable. Use this for advanced
        operations like complex formatting, chart creation, or custom automations
        that aren't possible through the Sheets/Docs APIs directly.

        Prerequisites:
        1. Open the script project in the Apps Script editor
        2. Deploy > New deployment > API Executable
        3. Use the script project ID (not the deployment ID)

        Args:
            params: Script project ID, function name, and optional parameters.

        Returns:
            The function's return value or error details.
        """
        body = {
            "function": params.function_name,
            "devMode": True,
        }
        if params.parameters:
            body["parameters"] = params.parameters

        try:
            response = get_scripts().scripts().run(
                scriptId=params.script_id,
                body=body,
            ).execute()

            if "error" in response:
                error = response["error"]
                details = error.get("details", [{}])
                error_msg = details[0].get("errorMessage", str(error)) if details else str(error)
                error_type = details[0].get("errorType", "UNKNOWN") if details else "UNKNOWN"
                return f"Script error ({error_type}): {error_msg}"

            result = response.get("response", {}).get("result")
            if result is None:
                return f"Function `{params.function_name}` executed successfully (no return value)."

            if isinstance(result, (dict, list)):
                return f"Function `{params.function_name}` returned:\n\n```json\n{json.dumps(result, indent=2)}\n```"

            return f"Function `{params.function_name}` returned: {result}"

        except Exception as e:
            error_str = str(e)
            if "not been deployed as an API Executable" in error_str:
                return (
                    f"Error: Script is not deployed as an API Executable.\n\n"
                    f"To fix:\n"
                    f"1. Open: https://script.google.com/d/{params.script_id}/edit\n"
                    f"2. Click Deploy > New deployment\n"
                    f"3. Select 'API Executable'\n"
                    f"4. Click Deploy\n"
                    f"5. Try again"
                )
            return f"Error running script: {e}"

    # ── Tool: Create Script ──────────────────────────────────────────────

    class CreateScriptInput(BaseModel):
        """Input for creating an Apps Script project."""
        model_config = ConfigDict(str_strip_whitespace=True)

        title: str = Field(..., description="Name for the Apps Script project.", min_length=1)
        parent_id: Optional[str] = Field(
            default=None,
            description="File ID of a Google Doc/Sheet/Slides to bind the script to. If omitted, creates a standalone script.",
        )
        code: str = Field(
            ...,
            description="The JavaScript/Apps Script code to add to the project.",
            min_length=1,
        )

    @mcp.tool(
        name="gdrive_create_script",
        annotations={"title": "Create Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_create_script(params: CreateScriptInput) -> str:
        """Create a new Google Apps Script project, optionally bound to a file.

        Creates the script project and pushes the provided code. The script
        can then be deployed and executed via gdrive_run_script.

        Args:
            params: Title, optional parent file ID, and the script code.

        Returns:
            Script project ID and link.
        """
        try:
            # Create project
            body = {"title": params.title}
            if params.parent_id:
                body["parentId"] = params.parent_id

            project = get_scripts().projects().create(body=body).execute()
            script_id = project["scriptId"]

            # Push the code
            content = {
                "files": [
                    {
                        "name": "Code",
                        "type": "SERVER_JS",
                        "source": params.code,
                    },
                    {
                        "name": "appsscript",
                        "type": "JSON",
                        "source": json.dumps({
                            "timeZone": "America/New_York",
                            "dependencies": {},
                            "exceptionLogging": "STACKDRIVER",
                            "runtimeVersion": "V8",
                            "oauthScopes": [
                                "https://www.googleapis.com/auth/spreadsheets",
                                "https://www.googleapis.com/auth/documents",
                                "https://www.googleapis.com/auth/drive",
                            ],
                        }),
                    },
                ]
            }

            get_scripts().projects().updateContent(
                scriptId=script_id,
                body=content,
            ).execute()

            link = f"https://script.google.com/d/{script_id}/edit"
            bound = f" (bound to file `{params.parent_id}`)" if params.parent_id else " (standalone)"

            return (
                f"Created Apps Script project: **{params.title}**{bound}\n\n"
                f"Script ID: `{script_id}`\n"
                f"Editor: {link}\n\n"
                f"To execute via API, deploy as 'API Executable' from the editor."
            )

        except Exception as e:
            return f"Error creating script: {e}"
