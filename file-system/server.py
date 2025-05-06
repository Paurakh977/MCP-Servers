from mcp.server import Server
import mcp.types as types
import asyncio
import os,sys
from typing import Optional, Dict, Any,List
from mcp.server.stdio import stdio_server
from pydantic import AnyUrl

server=Server(
    name="file-system",
    instructions="You are a file system MCP server . You can create, read, update, and delete files and directories. You can also list the contents of a directory. You can only access files and directories that are in the base working directory. You cannot access files and directories outside of the current base directory.",
    version="0.1",
)


def check_path_security(allowed_base_path: str, target_path: str) -> dict:
    """
    Checks if a target_path is allowed based on the allowed_base_path.

    This involves:
    1. Normalizing and resolving real paths for both inputs.
    2. Checking if the normalized paths exist on the filesystem.
    3. On Windows, checking if paths are on the same drive.
    4. Checking if the target path is a child path (or the same) as the base path
       using a secure relative path comparison.

    Args:
        allowed_base_path: The root directory path that is permitted.
        
        target_path: The path to a file or directory to check.
                     

    Returns:
        A dictionary containing detailed results of the checks:
        - 'original_base': The input allowed_base_path string.
        - 'original_target': The input target_path string.
        - 'normalized_base': The normalized and real path of the base, or None if normalization failed.
        - 'normalized_target': The normalized and real path of the target, or None if normalization failed.
        - 'normalization_successful': bool, True if normalization succeeded.
        - 'base_exists': bool, True if normalized_base_path exists.
        - 'target_exists': bool, True if normalized_target_path exists.
        - 'same_drive_check': bool, True if paths are on the same drive (Windows only, always True otherwise).
        - 'relative_path_from_base': str, The path from normalized_base to normalized_target, or None if not applicable.
        - 'is_within_base': bool, True if the relative path doesn't go outside the base.
        - 'is_allowed': bool, The final decision (True if all checks pass).
        - 'message': str, A human-readable summary of the result.
    """
    results = {
        'original_base': allowed_base_path,
        'original_target': target_path,
        'normalized_base': None,
        'normalized_target': None,
        'normalization_successful': False,
        'base_exists': False,
        'target_exists': False,
        'same_drive_check': True,
        'relative_path_from_base': None,
        'is_within_base': False,
        'is_allowed': False,
        'message': ''
    }

    try:
        normalized_base = os.path.realpath(os.path.normpath(allowed_base_path))
        normalized_target = os.path.realpath(os.path.normpath(target_path))
        results['normalized_base'] = normalized_base
        results['normalized_target'] = normalized_target
        results['normalization_successful'] = True
    except Exception as e:
        results['message'] = f"Normalization error: {e}"
        return results

    if not os.path.exists(normalized_base) or not os.path.isdir(normalized_base):
        results['message'] = f"Base path does not exist or not a directory: {normalized_base}"
        return results
    results['base_exists'] = True

    if not os.path.exists(normalized_target):
        results['message'] = f"Target does not exist: {normalized_target}"
        return results
    results['target_exists'] = True

    # Windows drive check
    if sys.platform == 'win32':
        base_drive = os.path.splitdrive(normalized_base)[0].lower()
        tgt_drive = os.path.splitdrive(normalized_target)[0].lower()
        if base_drive != tgt_drive:
            results['same_drive_check'] = False
            results['message'] = f"Different drive: {tgt_drive} vs {base_drive}"
            return results

    # Relative path
    try:
        rel = os.path.relpath(normalized_target, start=normalized_base)
        results['relative_path_from_base'] = rel
        if rel == os.pardir or rel.startswith(os.pardir + os.sep):
            results['message'] = f"Outside base: {rel}"
            return results
        results['is_within_base'] = True
        results['is_allowed'] = True
        results['message'] = "Path is allowed"
    except Exception as e:
        results['message'] = f"Relative path error: {e}"
        return results

    return results



@server.list_resources()
async def list_resources() -> list[types.Resource]:
    """
    Expose two resource schemas:
      1. file:///{path} for reading text files
      2. filesystem://list?path={path} for directory listings
    """
    
async def list_resources() -> list[types.Resource]:
    """
    Expose two resource schemas:
      1. file:///{path} for reading text files
      2. filesystem://list?path={path} for directory listings
    """
    
    return [
        types.Resource(
            uri="file:///{path}",
            name="Read File",
            mimeType="text/plain",
            description="Read any UTF-8 text file by absolute path",
            schema={
                "type": "object",
                "properties": {"path": {"type": "string"}},
                "required": ["path"],
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="filesystem://list?path={path}",
            name="Directory Listing",
            mimeType="application/json",
            description="List contents of a directory",
            schema={
                "type": "object",
                "properties": {"path": {"type": "string"}},
                "required": ["path"],
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
    ]
    
    
@server.read_resource()
async def read_resource(uri: AnyUrl) -> str:
    """
    Handle file and directory-listing URIs, returning text or JSON.
    """
    s = str(uri)
    if s.startswith("file:///"):
        path = s[len("file:///") :]
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")
        return open(path, "r", encoding="utf-8", errors="ignore").read()

    if s.startswith("filesystem://list"):
        # e.g. filesystem://list?path=/tmp
        q = s.split("?", 1)[1]
        path = dict(qc.split("=") for qc in q.split("&"))["path"]
        if not os.path.isdir(path):
            raise NotADirectoryError(f"Directory not found: {path}")
        return types.TextContent(type="json", text=str(os.listdir(path))).text

    raise ValueError(f"Unsupported resource URI: {uri}")



async def main() -> None:
    async with stdio_server() as streams:
        await server.run(
            read_stream=streams[0],
            write_stream=streams[1],
            initialization_options=server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())