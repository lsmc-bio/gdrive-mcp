"""Apps Script tools for gdrive-mcp."""

import json
from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from helpers import drive_query_files, escape_drive_query
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
        dev_mode: bool = Field(
            default=True,
            description="If true (default), runs the latest saved code (HEAD). If false, runs the most recently deployed version.",
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
            "devMode": params.dev_mode,
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

    # ── Tool: List Scripts ────────────────────────────────────────────────

    class ListScriptsInput(BaseModel):
        """Input for listing Apps Script projects."""
        model_config = ConfigDict(str_strip_whitespace=True)

        max_results: int = Field(default=20, description="Maximum projects to return.", ge=1, le=100)
        bound_to: Optional[str] = Field(
            default=None,
            description="File ID to filter scripts bound to a specific doc/sheet/slides.",
        )
        query: Optional[str] = Field(default=None, description="Search script projects by name.")

    @mcp.tool(
        name="gdrive_list_scripts",
        annotations={"title": "List Apps Script Projects", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_list_scripts(params: ListScriptsInput) -> str:
        """List Apps Script projects, optionally filtered by bound document or name.

        Args:
            params: Filters and result limit.

        Returns:
            List of script projects with IDs and links.
        """
        try:
            q = "mimeType='application/vnd.google-apps.script'"
            if params.query:
                q += f" and name contains '{escape_drive_query(params.query)}'"
            if params.bound_to:
                q += f" and '{params.bound_to}' in parents"

            files = drive_query_files(q, max_results=params.max_results)

            if not files:
                return "No Apps Script projects found."

            output = f"## Apps Script Projects ({len(files)})\n\n"
            for f in files:
                name = f.get("name", "untitled")
                fid = f.get("id", "?")
                modified = f.get("modifiedTime", "")[:10]
                link = f"https://script.google.com/d/{fid}/edit"
                output += f"**{name}**\n  ID: `{fid}` | Modified: {modified}\n  Editor: {link}\n\n"

            return output

        except Exception as e:
            return f"Error listing scripts: {e}"

    # ── Tool: Get Script ──────────────────────────────────────────────────

    class GetScriptInput(BaseModel):
        """Input for reading an Apps Script project."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        file_name: Optional[str] = Field(
            default=None,
            description="Specific file to read (e.g., 'Code.gs'). If omitted, returns all files.",
        )
        include_manifest: bool = Field(default=True, description="Include appsscript.json manifest.")

    @mcp.tool(
        name="gdrive_get_script",
        annotations={"title": "Read Apps Script Project", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_get_script(params: GetScriptInput) -> str:
        """Read the contents of an Apps Script project.

        Can read a specific file or all files. Includes the manifest by default.

        Args:
            params: Script ID and optional file filter.

        Returns:
            Script file contents.
        """
        try:
            content = get_scripts().projects().getContent(scriptId=params.script_id).execute()
            files = content.get("files", [])

            if params.file_name:
                files = [f for f in files if f.get("name") == params.file_name.replace(".gs", "").replace(".html", "").replace(".json", "")]
                if not files:
                    return f"File '{params.file_name}' not found in project."

            if not params.include_manifest:
                files = [f for f in files if f.get("name") != "appsscript"]

            if not files:
                return "No files in project."

            output = f"## Apps Script Project `{params.script_id}`\n\n"
            for f in files:
                name = f.get("name", "?")
                ftype = f.get("type", "?")
                ext = {"SERVER_JS": ".gs", "HTML": ".html", "JSON": ".json"}.get(ftype, "")
                source = f.get("source", "")
                output += f"### {name}{ext} ({ftype})\n\n```{'javascript' if ftype == 'SERVER_JS' else 'html' if ftype == 'HTML' else 'json'}\n{source}\n```\n\n"

            return output

        except Exception as e:
            return f"Error reading script: {e}"

    # ── Tool: Update Script ───────────────────────────────────────────────

    class ScriptFileInput(BaseModel):
        name: str = Field(..., description="File name (e.g., 'Code.gs', 'utils.gs', 'sidebar.html').")
        source: str = Field(..., description="File source code.")
        type: Optional[str] = Field(
            default=None,
            description="File type: 'SERVER_JS', 'HTML', or 'JSON'. Auto-detected from extension if omitted.",
        )

    class UpdateScriptInput(BaseModel):
        """Input for updating an Apps Script project."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        files: List[ScriptFileInput] = Field(..., description="Files to create or update.", min_length=1)
        merge: bool = Field(
            default=True,
            description="If true (default), only update specified files — keep existing files untouched. If false, replace ALL files.",
        )

    @mcp.tool(
        name="gdrive_update_script",
        annotations={"title": "Update Apps Script Project", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_update_script(params: UpdateScriptInput) -> str:
        """Update files in an Apps Script project.

        Default merge mode only updates the files you specify — existing files
        are preserved. Set merge=False to replace all files.

        Args:
            params: Script ID, files to update, and merge mode.

        Returns:
            Confirmation of files updated.
        """
        try:
            svc = get_scripts()

            # Auto-detect file types from extensions
            type_map = {".gs": "SERVER_JS", ".js": "SERVER_JS", ".html": "HTML", ".json": "JSON"}

            new_files = []
            for f in params.files:
                ftype = f.type
                if not ftype:
                    for ext, t in type_map.items():
                        if f.name.endswith(ext):
                            ftype = t
                            break
                    if not ftype:
                        ftype = "SERVER_JS"

                # Strip extension for API (it uses bare names)
                name = f.name
                for ext in type_map:
                    if name.endswith(ext):
                        name = name[:-len(ext)]
                        break

                new_files.append({"name": name, "type": ftype, "source": f.source})

            if params.merge:
                # Fetch existing files and merge
                existing = svc.projects().getContent(scriptId=params.script_id).execute()
                existing_files = existing.get("files", [])

                new_names = {f["name"] for f in new_files}
                merged = [f for f in existing_files if f["name"] not in new_names]
                merged.extend(new_files)
                final_files = merged
            else:
                final_files = new_files

            svc.projects().updateContent(
                scriptId=params.script_id,
                body={"files": final_files},
            ).execute()

            names = [f["name"] for f in new_files]
            mode = "merged into" if params.merge else "replaced all files in"
            return f"Updated {len(new_files)} file(s) ({', '.join(names)}) — {mode} project `{params.script_id}`."

        except Exception as e:
            return f"Error updating script: {e}"

    # ── Tool: Deploy Script ───────────────────────────────────────────────

    class DeployScriptInput(BaseModel):
        """Input for managing Apps Script deployments."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        action: str = Field(
            ...,
            description="Action: 'create', 'list', 'update', or 'delete'.",
        )
        deployment_id: Optional[str] = Field(default=None, description="Deployment ID (required for 'update' and 'delete').")
        description: Optional[str] = Field(default=None, description="Deployment description.")
        version: Optional[int] = Field(default=None, description="Script version number. If omitted, uses latest.")

    @mcp.tool(
        name="gdrive_deploy_script",
        annotations={"title": "Deploy Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_deploy_script(params: DeployScriptInput) -> str:
        """Create, list, update, or delete Apps Script deployments.

        Args:
            params: Script ID, action, and deployment details.

        Returns:
            Deployment details or listing.
        """
        try:
            svc = get_scripts()

            if params.action == "list":
                result = svc.projects().deployments().list(scriptId=params.script_id).execute()
                deployments = result.get("deployments", [])
                if not deployments:
                    return "No deployments found."

                output = f"## Deployments ({len(deployments)})\n\n"
                for d in deployments:
                    did = d.get("deploymentId", "?")
                    config = d.get("deploymentConfig", {})
                    output += f"**{config.get('description', 'No description')}**\n"
                    output += f"  ID: `{did}`\n"
                    output += f"  Version: {config.get('versionNumber', 'HEAD')}\n"
                    output += f"  Type: {config.get('scriptId', 'API_EXECUTABLE')}\n\n"
                return output

            elif params.action == "create":
                # Create a version first if needed
                if params.version is None:
                    ver = svc.projects().versions().create(
                        scriptId=params.script_id,
                        body={"description": params.description or "Deployed via gdrive-mcp"},
                    ).execute()
                    version_number = ver.get("versionNumber")
                else:
                    version_number = params.version

                config = {"versionNumber": version_number}
                if params.description:
                    config["description"] = params.description

                deployment = svc.projects().deployments().create(
                    scriptId=params.script_id,
                    body={"deploymentConfig": config},
                ).execute()

                did = deployment.get("deploymentId", "?")
                return f"Created deployment `{did}` (version {version_number})."

            elif params.action == "update":
                if not params.deployment_id:
                    return "Error: deployment_id is required for 'update'."

                config = {}
                if params.version:
                    config["versionNumber"] = params.version
                if params.description:
                    config["description"] = params.description

                svc.projects().deployments().update(
                    scriptId=params.script_id,
                    deploymentId=params.deployment_id,
                    body={"deploymentConfig": config},
                ).execute()

                return f"Updated deployment `{params.deployment_id}`."

            elif params.action == "delete":
                if not params.deployment_id:
                    return "Error: deployment_id is required for 'delete'."

                svc.projects().deployments().delete(
                    scriptId=params.script_id,
                    deploymentId=params.deployment_id,
                ).execute()

                return f"Deleted deployment `{params.deployment_id}`."

            else:
                return "Error: action must be 'create', 'list', 'update', or 'delete'."

        except Exception as e:
            return f"Error managing deployment: {e}"
