"""Resource, Prmopts and tool definitions for the MCP file-system server."""

from .resource_definitions import get_resource_definitions
from .tool_definitions import get_tool_definitions
from .prompt_definitions import PROMPTS

__all__ = [
    'get_resource_definitions',
    'get_tool_definitions',
    'PROMPTS',
] 