import asyncio
from typing import Optional
from contextlib import AsyncExitStack
import os
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from dotenv import load_dotenv
import  google.generativeai as genai

load_dotenv()

API_KEY = os.getenv("GOOGLE_API_KEY")
MODEL=os.getenv("MODEL", "gemini-2.0-flash")


class MCPClient:
    def __init__(self,):
        self.session: Optional[ClientSession] = None
        self.model= MODEL
        self.exit_stack = AsyncExitStack()
        if API_KEY:
            genai.configure(api_key=API_KEY)
            self.llm = genai.GenerativeModel(self.model)
        else:
            raise ValueError("API_KEY is not set in the environment variables.")
        
    async def connect_to_server(self,server_script_path):
        """Connect to an MCP server
        Args:
            server_script_path: Path to the server script (.py or .js)
        """
        
        # check if the server script exists
        
        if not os.path.exists(server_script_path):
            raise FileNotFoundError(f"Server script {server_script_path} does not exist.")
        
        is_python = server_script_path.endswith('.py')
        is_js = server_script_path.endswith('.js')
        if not (is_python or is_js):
            raise ValueError("Server script must be a .py or .js file")
            
        command = "python" if is_python else "node"
        server_params = StdioServerParameters(
            command=command,
            args=[server_script_path],
            env=None
        )
        
        stdio_transport = await self.exit_stack.enter_async_context(stdio_client(server_params))
        self.stdio, self.write = stdio_transport
        self.session = await self.exit_stack.enter_async_context(ClientSession(self.stdio, self.write))
        
        await self.session.initialize()
        
        # List available tools
        response = await self.session.list_tools()
        tools = response.tools
        print("\nConnected to server with tools:", [tool.name for tool in tools])
