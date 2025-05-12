"""Prompt definitions for the MCP file-system server."""

import mcp.types as types

PROMPTS = {
    "find-and-read-file": types.Prompt(
        name="find-and-read-file",
        description="Find a file by its name (even with partial or misspelled names) and read its content",
        arguments=[
            types.PromptArgument(
                name="filename",
                description="The full or partial name of the file you want to read",
                required=True,
            ),
            types.PromptArgument(
                name="summarize",
                description="Whether to summarize large content (true/false)",
                required=False,
            ),
            types.PromptArgument(
                name="max_length",
                description="Maximum length for summary if summarizing",
                required=False,
            ),
        ],
    ),
    "read-excel-as-table": types.Prompt(
        name="read-excel-as-table",
        description="Locate an Excel/CSV file, read the specified sheet, and output it as a markdown table.",
        arguments=[
            types.PromptArgument(
                name="filename",
                description="Partial or full filename (e.g. 'sales.xlsx' or '*.csv') to locate in allowed directories",
                required=True,
            ),
            types.PromptArgument(
                name="sheet_name",
                description="Name of the Excel sheet to read. If the file is CSV or omitted, the assistant should read the first/default sheet.",
                required=True,
            ),
            types.PromptArgument(
                name="summarize",
                description="Whether to append a textual summary after the table (true/false)",
                required=False,
            ),
            types.PromptArgument(
                name="max_length",
                description="Maximum length (in characters or tokens) of the summary, if summarizing",
                required=False,
            ),
        ],
    ),
    "find-files": types.Prompt(
        name="find-files",
        description="Find files matching a pattern and get a list of matching files",
        arguments=[
            types.PromptArgument(
                name="pattern",
                description="The pattern to search for in filenames",
                required=True,
            )
        ],
    ),
    "find-file-info": types.Prompt(
        name="find-file-info",
        description="Find a file by name and get detailed information about it",
        arguments=[
            types.PromptArgument(
                name="filename",
                description="The full or partial name of the file to get info about",
                required=True,
            )
        ],
    ),
}
