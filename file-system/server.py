import base64
from mimetypes import guess_type
from pathlib import Path
from mcp.server import Server
from mcp.server.stdio import stdio_server
from typing import Dict, Any, List
import mcp.types as types
import asyncio
from pydantic import AnyUrl
import os
from urllib.parse import urlparse, parse_qs
import subprocess

app = Server(
    name="Coding Server", 
    version="0.0.1", 
)

@app.list_resources()
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


@app.read_resource()
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



PROMPTS = {
    "python": types.Prompt(
        name="Python",
        description="Python code boilerplate",
        arguments=[
            types.PromptArgument(
                name="code_description",
                description="Description of the code to be generated",
                required=True,
            )
        ]
    ),
    "javascript": types.Prompt(
        name="JavaScript",
        description="JavaScript code boilerplate",
        arguments=[
            types.PromptArgument(
                name="code_description",
                description="Description of the code to be generated",
                required=True,
            )
        ]
    ),
    "html": types.Prompt(
        name="HTML",
        description="HTML code boilerplate",
        arguments=[
            types.PromptArgument(
                name="code_description",
                description="Description of the code to be generated",
                required=True,
            )
        ]
    ),
    "css": types.Prompt(
        name="CSS",
        description="CSS code boilerplate",
        arguments=[
            types.PromptArgument(
                name="code_description",
                description="Description of the code to be generated",
                required=True,
            )
        ]
    ),
    "rust": types.Prompt(
        name="Rust",
        description="Rust code boilerplate",
        arguments=[
            types.PromptArgument(
                name="code_description",
                description="Description of the code to be generated",
                required=True,
            )
        ]
    ),
    "summarize-output": types.Prompt(
        name="summarize-output",
        description="Summarize the stdout/stderr from a program execution",
        arguments=[
            types.PromptArgument(
                name="output",
                description="Captured program stdout and stderr",
                required=True
            )
        ],
        idempotentHint=True
    )
}


@app.list_prompts()
async def list_prompts() -> list[types.Prompt]:
    """
    Returns a list of available prompts.
    """
    return list(PROMPTS.values())
    

@app.get_prompt()
async def get_prompt(prompt_name: str, arguments: Dict[str, str]) -> types.GetPromptResult:
    """
    Returns a prompt by name.
    """
    if prompt_name not in PROMPTS:
        raise ValueError(f"Prompt not found: {prompt_name}")
    elif prompt_name == "python":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a Python developer and expert."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Generate a Python code for {arguments['code_description']}"),
                ),
            ]
        )
    elif prompt_name == "javascript":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a JavaScript developer and expert."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Generate a JavaScript code for {arguments['code_description']}"),
                ),
            ]
        )
    elif prompt_name == "html":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a HTML developer and expert."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Generate a HTML code for {arguments['code_description']}"),
                ),
            ]
        )
    elif prompt_name == "css":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a CSS developer and expert."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Generate a CSS code for {arguments['code_description']}"),
                ),
            ]
        )
    elif prompt_name == "rust":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a Rust developer and expert."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Generate a Rust code for {arguments['code_description']}"),
                ),
            ]
        )
    elif prompt_name == "summarize-output":
        return types.GetPromptResult(
            messages=[
                types.PromptMessage(
                    role="system",
                    content=types.TextContent(type="text", text=f"Your are a helpful assistant that will summarize the code's output properly."),
                ),
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=f"Summarize the output of the program: {arguments['output']}"),
                ),
            ]
        )
    

@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """
    Returns a list of available tools.
    """
    return [
        types.Tool(
            name="write-python-file",
            description="Write a new python file to the filesystem",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute python filesystem path to the file"
                    },
                    "content": {
                        "type": "string",
                        "description": "Content to write to the python file"
                    }
                },
                "required": ["path", "content"],
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="write-javascript-file",
            description="Write a new javascript file to the filesystem",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute javascript filesystem path to the file"
                    },
                    "content": {
                        "type": "string",
                        "description": "Content to write to the javascript file"
                    }
                },
                "required": ["path", "content"],
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="write-html-file",
            description="Write a new html file to the filesystem",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute html filesystem path to the file"
                    },
                    "content": {
                        "type": "string",
                        "description": "Content to write to the html file"
                    }
                },
                "required": ["path", "content"], 
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="write-css-file",
            description="Write a new css file to the filesystem",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute css filesystem path to the file"
                    },
                    "content": {
                        "type": "string",
                        "description": "Content to write to the css file"
                    }
                },
                "required": ["path", "content"],
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="delete-file",
            description="Delete a file with file-path",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute filesystem path to the file"
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="Explain-code",
            description="Explain a code snippet",
            inputSchema={
                "type": "object",
                "properties": {
                    "language": {
                        "type": "string",
                        "description": "Programming language of the code snippet"
                    },
                    "path": {
                        "type": "string",
                        "description": "Code snippet path to explain"
                    }
                },
                "required": ["language", "path"],
                "additionalProperties": False
            },
        ),
        types.Tool(
            name="exec-program",
            description="Execute any type of program file and return its output",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Absolute path to an executable script or program file"
                    },
                    "args": {
                        "type": "array",
                        "items": {"type":"string"},
                        "description": "Command-line arguments to pass (optional)"
                    },
                    "language": {
                        "type": "string",
                        "description": "Programming language of the file (e.g., python, node, php, ruby, etc.)"
                    }
                },
                "required": ["path", "language"],
                "additionalProperties": False
            },
            idempotentHint=False,
            destructiveHint=False
        )
    ]


@app.call_tool()
async def call_tool(tool_name: str, arguments: Dict[str, Any]) -> list[types.TextContent]:
    """
    Calls a tool by name with validated arguments.
    """
    if tool_name not in ["write-python-file", "write-javascript-file", "write-html-file", 
                         "write-css-file", "exec-and-summarize", "delete-file", "Explain-code"]:
        raise ValueError(f"Tool not found: {tool_name}")
    
    if tool_name == "write-python-file":
        if not arguments.get("path", "").endswith(".py"):
           raise ValueError("File path must end with .py")
        
        path = arguments["path"]
        content = arguments["content"]
        with open(path, "w", encoding="utf-8") as file:
            file.write(content)
        return [types.TextContent(
            type="text",
            text=f"Python file written to {path}"
        )]
        
    elif tool_name == "write-javascript-file":
        if not arguments.get("path", "").endswith(".js"):
            raise ValueError("File path must end with .js")
        
        path = arguments["path"]
        content = arguments["content"]
        with open(path, "w", encoding="utf-8") as file:
            file.write(content)
        return [types.TextContent(
            type="text",
            text=f"JavaScript file written to {path}"
        )]
        
    elif tool_name == "write-html-file":
        if not arguments.get("path", "").endswith(".html"):
            raise ValueError("File path must end with .html")
        
        path = arguments["path"]
        content = arguments["content"]
        with open(path, "w", encoding="utf-8") as file:
            file.write(content)
        return [types.TextContent(
            type="text",
            text=f"HTML file written to {path}"
        )]
        
    elif tool_name == "write-css-file":
        if not arguments.get("path", "").endswith(".css"):
            raise ValueError("File path must end with .css")
        
        path = arguments["path"]
        content = arguments["content"]
        with open(path, "w", encoding="utf-8") as file:
            file.write(content)
        return [types.TextContent(
            type="text",
            text=f"CSS file written to {path}"
        )]
        
    elif tool_name == "delete-file":
        path = arguments["path"]
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")
        
        os.remove(path)
        return [types.TextContent(
            type="text",
            text=f"File deleted: {path}"
        )]
        
    elif tool_name == "explain-code":
        path = arguments["path"]
        
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")
            
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as file:
                code_content = file.read()
            
            # Determine language based on file extension
            extension = os.path.splitext(path)[1].lower()
            language_map = {
                ".py": "Python",
                ".js": "JavaScript",
                ".html": "HTML",
                ".css": "CSS",
                ".java": "Java",
                ".c": "C",
                ".cpp": "C++",
                ".cs": "C#",
                ".go": "Go",
                ".php": "PHP",
                ".rb": "Ruby",
                ".rs": "Rust",
                ".sh": "Shell",
                ".ts": "TypeScript",
                ".swift": "Swift",
                ".kt": "Kotlin",
                ".pl": "Perl",
                ".r": "R",
                ".sql": "SQL"
            }
            language = language_map.get(extension, "Unknown")
            
            return [
                types.TextContent(
                    type="text",
                    text=f"Code explanation for {language} file ({path}):\n\n{code_content}"
                )
            ]
        except Exception as e:
            raise ValueError(f"Error reading file: {e}")
        
    elif tool_name == "exec-program":
        path = arguments["path"]
        args = arguments.get("args", [])
        language = arguments["language"].lower()
        
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")
        
        # Dictionary mapping language to interpreter/command
        interpreters = {
            "python": "python",
            "python2": "python2",
            "python3": "python3",
            "node": "node",
            "javascript": "node",
            "php": "php",
            "ruby": "ruby",
            "perl": "perl",
            "bash": "bash",
            "sh": "sh",
            "r": "Rscript",
            "java": "java",
            "c": "cc",
            "cpp": "g++",
            "go": "go run",
            "rust": "rustc",
            "swift": "swift",
            "typescript": "ts-node",
            "kotlin": "kotlin"
        }
        
        # Special handling for compiled languages
        compiled_languages = ["c", "cpp", "java", "rust", "go"]
        
        try:
            if language in compiled_languages:
                # Handle compilation and execution for compiled languages
                if language == "c":
                    output_path = path.replace(".c", ".out")
                    compile_cmd = await asyncio.create_subprocess_shell(
                        f"gcc {path} -o {output_path}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    _, compile_stderr = await compile_cmd.communicate()
                    
                    if compile_cmd.returncode != 0:
                        return [types.TextContent(
                            type="text",
                            text=f"Compilation error:\n{compile_stderr.decode()}"
                        )]
                    
                    exec_cmd = await asyncio.create_subprocess_exec(
                        output_path, *args,
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    
                elif language == "cpp":
                    output_path = path.replace(".cpp", ".out")
                    compile_cmd = await asyncio.create_subprocess_shell(
                        f"g++ {path} -o {output_path}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    _, compile_stderr = await compile_cmd.communicate()
                    
                    if compile_cmd.returncode != 0:
                        return [types.TextContent(
                            type="text",
                            text=f"Compilation error:\n{compile_stderr.decode()}"
                        )]
                    
                    exec_cmd = await asyncio.create_subprocess_exec(
                        output_path, *args,
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    
                elif language == "java":
                    # Extract class name (assuming it matches the filename)
                    class_name = os.path.basename(path).replace(".java", "")
                    dir_path = os.path.dirname(path)
                    
                    # Compile Java file
                    compile_cmd = await asyncio.create_subprocess_shell(
                        f"javac {path}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    _, compile_stderr = await compile_cmd.communicate()
                    
                    if compile_cmd.returncode != 0:
                        return [types.TextContent(
                            type="text",
                            text=f"Java compilation error:\n{compile_stderr.decode()}"
                        )]
                    
                    # Execute Java class
                    exec_cmd = await asyncio.create_subprocess_shell(
                        f"java -cp {dir_path} {class_name} {' '.join(args)}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    
                elif language == "rust":
                    output_path = path.replace(".rs", ".out")
                    compile_cmd = await asyncio.create_subprocess_shell(
                        f"rustc {path} -o {output_path}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    _, compile_stderr = await compile_cmd.communicate()
                    
                    if compile_cmd.returncode != 0:
                        return [types.TextContent(
                            type="text",
                            text=f"Rust compilation error:\n{compile_stderr.decode()}"
                        )]
                    
                    exec_cmd = await asyncio.create_subprocess_exec(
                        output_path, *args,
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                    
                elif language == "go":
                    # Go can be run directly with "go run"
                    exec_cmd = await asyncio.create_subprocess_shell(
                        f"go run {path} {' '.join(args)}",
                        stdout=asyncio.subprocess.PIPE,
                        stderr=asyncio.subprocess.PIPE
                    )
                
                # Get output from the executed program
                stdout, stderr = await exec_cmd.communicate()
                output = stdout.decode() + ("\n--- STDERR ---\n" + stderr.decode() if stderr else "")
                
            else:
                # Handle interpreted languages
                if language not in interpreters:
                    raise ValueError(f"Unsupported language: {language}")
                
                interpreter = interpreters[language]
                
                # Make sure script is executable for shell scripts
                if language in ["bash", "sh"]:
                    os.chmod(path, 0o755)
                
                # Execute the script with the appropriate interpreter
                exec_cmd = await asyncio.create_subprocess_shell(
                    f"{interpreter} {path} {' '.join(args)}",
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE
                )
                
                stdout, stderr = await exec_cmd.communicate()
                output = stdout.decode() + ("\n--- STDERR ---\n" + stderr.decode() if stderr else "")
            
            return [types.TextContent(
                type="text",
                text=f"Program executed ({language}). Exit code: {exec_cmd.returncode}\n\nOutput:\n\n{output}"
            )]
            
        except Exception as e:
            return [types.TextContent(
                type="text",
                text=f"Error executing {language} program: {str(e)}"
            )]
 
    elif tool_name == "Explain-code":
        language = arguments["language"]
        path = arguments["path"]
        
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")
            
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as file:
                code_content = file.read()
            
            return [
                types.TextContent(
                    type="text",
                    text=f"Code explanation for {language}:\n\n{code_content}"
                )
            ]
        except Exception as e:
            raise ValueError(f"Error reading file: {e}")

async def main() -> None:
    async with stdio_server() as streams:
        await app.run(
            read_stream=streams[0],
            write_stream=streams[1],
            initialization_options=app.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())