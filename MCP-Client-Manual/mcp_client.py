#!/usr/bin/env python3

import asyncio
import json
import os
import logging
import sys
from contextlib import AsyncExitStack
from typing import Dict, List, Optional, Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.client.sse import sse_client

try:
    from colorama import init, Fore, Style
    init()  # Initialize colorama
    COLOR_SUPPORT = True
except ImportError:
    COLOR_SUPPORT = False
    print("For colorful output, install colorama: pip install colorama")

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Define colors if colorama is available
if COLOR_SUPPORT:
    INFO_COLOR = Fore.CYAN
    SUCCESS_COLOR = Fore.GREEN
    ERROR_COLOR = Fore.RED
    WARNING_COLOR = Fore.YELLOW
    HIGHLIGHT_COLOR = Fore.MAGENTA
    RESET = Style.RESET_ALL
else:
    INFO_COLOR = ""
    SUCCESS_COLOR = ""
    ERROR_COLOR = ""
    WARNING_COLOR = ""
    HIGHLIGHT_COLOR = ""
    RESET = ""

class MCPClient:
    def __init__(self, config_path: str = "config.json"):
        """Initialize the MCP client with the configuration file.
        
        Args:
            config_path: Path to the configuration JSON file
        """
        self.config_path = config_path
        self.config = self._load_config()
        self.sessions = {}
        self.exit_stack = AsyncExitStack()
    
    def _load_config(self) -> Dict:
        """Load configuration from JSON file."""
        try:
            with open(self.config_path, "r") as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"{ERROR_COLOR}Error: Configuration file '{self.config_path}' not found.{RESET}")
            print(f"Please create a '{self.config_path}' file with your MCP server configurations.")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"{ERROR_COLOR}Error: Configuration file '{self.config_path}' is invalid JSON.{RESET}")
            print(f"JSON error: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"{ERROR_COLOR}Error loading configuration: {e}{RESET}")
            sys.exit(1)
    
    async def connect_to_server(self, server_name: str):
        """Connect to an MCP server specified in the configuration.
        
        Args:
            server_name: Name of the server in the configuration
            
        Returns:
            bool: True if connection successful, False otherwise
        """
        if server_name not in self.config.get("mcpServers", {}):
            print(f"{ERROR_COLOR}Error: Server '{server_name}' not found in configuration.{RESET}")
            return False
        
        server_config = self.config["mcpServers"][server_name]
        command = server_config.get("command")
        args = server_config.get("args", [])
        
        if not command:
            print(f"{ERROR_COLOR}Error: No command specified for server '{server_name}'.{RESET}")
            return False
        
        try:
            print(f"{INFO_COLOR}Connecting to {server_name} MCP server...{RESET}")
            
            server_params = StdioServerParameters(
                command=command,
                args=args,
                env=None
            )
            
            stdio_transport = await self.exit_stack.enter_async_context(stdio_client(server_params))
            stdio_reader, stdio_writer = stdio_transport
            session = await self.exit_stack.enter_async_context(ClientSession(stdio_reader, stdio_writer))
            
            await session.initialize()
            
            # List available tools
            response = await session.list_tools()
            tools = response.tools
            
            print(f"{SUCCESS_COLOR}Connected to {server_name} MCP Server{RESET}")
            
            if tools:
                tool_names = [tool.name for tool in tools]
                print(f"{INFO_COLOR}Available tools ({len(tools)}): {HIGHLIGHT_COLOR}{', '.join(tool_names)}{RESET}")
            else:
                print(f"{WARNING_COLOR}No tools found on server '{server_name}'{RESET}")
            
            # Store the session for later use
            self.sessions[server_name] = {
                "session": session,
                "tools": tools
            }
            
            return True
            
        except Exception as e:
            print(f"{ERROR_COLOR}Failed to connect to server '{server_name}': {e}{RESET}")
            logger.error(f"Connection error: {e}", exc_info=True)
            return False
    
    def list_available_servers(self) -> List[str]:
        """List all available server names from configuration."""
        return list(self.config.get("mcpServers", {}).keys())
    
    def list_available_tools(self, server_name: str) -> List[str]:
        """List all available tools for a connected server.
        
        Args:
            server_name: Name of the server
        """
        if server_name not in self.sessions:
            print(f"{ERROR_COLOR}Error: Server '{server_name}' not connected.{RESET}")
            return []
        
        return [tool.name for tool in self.sessions[server_name]["tools"]]
    
    def get_tool_details(self, server_name: str, tool_name: str = None):
        """Get detailed information about tools on a server.
        
        Args:
            server_name: Name of the server
            tool_name: Optional name of specific tool to get details for
        """
        if server_name not in self.sessions:
            print(f"{ERROR_COLOR}Error: Server '{server_name}' not connected.{RESET}")
            return None
        
        tools = self.sessions[server_name]["tools"]
        
        if tool_name:
            for tool in tools:
                if tool.name == tool_name:
                    return tool
            print(f"{ERROR_COLOR}Error: Tool '{tool_name}' not found on server '{server_name}'.{RESET}")
            return None
        
        return tools
    
    async def call_tool(self, server_name: str, tool_name: str, arguments: Dict[str, Any]):
        """Call a tool on the specified server.
        
        Args:
            server_name: Name of the server
            tool_name: Name of the tool to call
            arguments: Arguments to pass to the tool
        """
        if server_name not in self.sessions:
            print(f"{ERROR_COLOR}Error: Server '{server_name}' not connected.{RESET}")
            return None
            
        session = self.sessions[server_name]["session"]
        
        try:
            # Find tool definition
            tool = self.get_tool_details(server_name, tool_name)
            if not tool:
                return None
                
            # Print call information
            print(f"{INFO_COLOR}Calling tool: {HIGHLIGHT_COLOR}{tool_name}{RESET}")
            print(f"{INFO_COLOR}Arguments:{RESET}")
            print(json.dumps(arguments, indent=2))
            
            # Confirm execution
            if not self._confirm_execution():
                print(f"{WARNING_COLOR}Tool execution cancelled.{RESET}")
                return None
            
            # Execute the tool
            print(f"{INFO_COLOR}Executing tool...{RESET}")
            start_time = asyncio.get_event_loop().time()
            response = await session.call_tool(name=tool_name, arguments=arguments)
            end_time = asyncio.get_event_loop().time()
            execution_time = end_time - start_time
            
            # Handle the response
            if response and hasattr(response, 'tool_result'):
                result = response.tool_result
                is_error = result.get('isError', False) if isinstance(result, dict) else False
                
                if is_error:
                    print(f"{ERROR_COLOR}Tool execution failed ({execution_time:.2f}s):{RESET}")
                    return result
                else:
                    print(f"{SUCCESS_COLOR}Tool execution successful ({execution_time:.2f}s){RESET}")
                    return result
            else:
                print(f"{WARNING_COLOR}Tool returned no result ({execution_time:.2f}s){RESET}")
                return None
                
        except asyncio.CancelledError:
            print(f"{WARNING_COLOR}Tool execution cancelled.{RESET}")
            return None
        except Exception as e:
            print(f"{ERROR_COLOR}Failed to call tool '{tool_name}': {e}{RESET}")
            logger.error(f"Tool execution error: {e}", exc_info=True)
            return None
    
    def _confirm_execution(self) -> bool:
        """Ask for confirmation before executing a tool.
        
        Returns:
            bool: True if confirmed, False otherwise
        """
        confirm = input(f"{WARNING_COLOR}Execute this tool? (y/n): {RESET}").lower()
        return confirm in ('y', 'yes')
    
    async def close(self):
        """Close all connections and clean up resources."""
        print(f"{INFO_COLOR}Closing all connections...{RESET}")
        try:
            await self.exit_stack.aclose()
            print(f"{SUCCESS_COLOR}All connections closed.{RESET}")
        except Exception as e:
            print(f"{ERROR_COLOR}Error closing connections: {e}{RESET}")

async def display_banner():
    """Display a welcome banner for the MCP client."""
    banner = [
        "┌────────────────────────────────────────────┐",
        "│ MCP Client - Model Context Protocol Client │",
        "│                                            │",
        "│ Type 'help' for available commands         │",
        "│ Type 'exit' or 'quit' to exit              │",
        "└────────────────────────────────────────────┘"
    ]
    
    for line in banner:
        print(f"{HIGHLIGHT_COLOR}{line}{RESET}")

async def interactive_session(client: MCPClient):
    """Run an interactive session with the connected MCP servers."""
    await display_banner()
    
    servers = client.list_available_servers()
    if not servers:
        print(f"{ERROR_COLOR}No servers found in configuration.{RESET}")
        return
    
    while True:
        try:
            command = input(f"\n{HIGHLIGHT_COLOR}mcp>{RESET} ").strip().lower()
            
            if command == "quit" or command == "exit":
                print(f"{INFO_COLOR}Exiting interactive mode...{RESET}")
                break
                
            elif command == "help":
                print(f"\n{HIGHLIGHT_COLOR}Available commands:{RESET}")
                print(f"  {INFO_COLOR}servers              {RESET}- List available servers")
                print(f"  {INFO_COLOR}connect <server>     {RESET}- Connect to a server")
                print(f"  {INFO_COLOR}tools <server>       {RESET}- List tools for a server")
                print(f"  {INFO_COLOR}info <server> <tool> {RESET}- Get detailed info about a tool")
                print(f"  {INFO_COLOR}call <server> <tool> {RESET}- Call a tool with JSON arguments")
                print(f"  {INFO_COLOR}exit/quit            {RESET}- Exit interactive mode")
                print(f"  {INFO_COLOR}help                 {RESET}- Show this help message")
                
            elif command == "servers":
                servers = client.list_available_servers()
                if servers:
                    print(f"\n{INFO_COLOR}Available servers:{RESET}")
                    for idx, server in enumerate(servers, 1):
                        connected = server in client.sessions
                        status = f"{SUCCESS_COLOR}[Connected]{RESET}" if connected else f"{WARNING_COLOR}[Not Connected]{RESET}"
                        print(f"  {idx}. {HIGHLIGHT_COLOR}{server:<20}{RESET} {status}")
                else:
                    print(f"{WARNING_COLOR}No servers defined in configuration.{RESET}")
                
            elif command.startswith("connect "):
                server_name = command[8:].strip()
                connected = await client.connect_to_server(server_name)
                if connected:
                    tools = client.list_available_tools(server_name)
                    if tools:
                        print(f"{SUCCESS_COLOR}Connected to {server_name} with {len(tools)} tools available.{RESET}")
            
            elif command.startswith("tools "):
                server_name = command[6:].strip()
                tools = client.list_available_tools(server_name)
                if tools:
                    print(f"\n{INFO_COLOR}Tools for {HIGHLIGHT_COLOR}{server_name}{INFO_COLOR}:{RESET}")
                    for idx, tool in enumerate(tools, 1):
                        print(f"  {idx}. {HIGHLIGHT_COLOR}{tool}{RESET}")
                else:
                    print(f"{WARNING_COLOR}No tools available or server not connected.{RESET}")
                        
            elif command.startswith("info "):
                parts = command[5:].strip().split()
                if len(parts) < 1:
                    print(f"{ERROR_COLOR}Usage: info <server> [tool]{RESET}")
                    continue
                    
                server_name = parts[0]
                tool_name = parts[1] if len(parts) > 1 else None
                
                tools = client.get_tool_details(server_name, tool_name)
                if tools:
                    if tool_name:
                        # Display single tool details
                        print(f"\n{HIGHLIGHT_COLOR}Tool Details:{RESET}")
                        print(f"{INFO_COLOR}Name:        {RESET}{tools.name}")
                        print(f"{INFO_COLOR}Description: {RESET}{tools.description}")
                        
                        # Pretty print the input schema
                        if hasattr(tools, 'input_schema'):
                            print(f"{INFO_COLOR}Input Schema:{RESET}")
                            print(json.dumps(tools.input_schema, indent=2))
                    else:
                        # Display all tools
                        print(f"\n{INFO_COLOR}Tools for {HIGHLIGHT_COLOR}{server_name}{RESET}:")
                        for idx, tool in enumerate(tools, 1):
                            print(f"\n{HIGHLIGHT_COLOR}Tool #{idx}: {tool.name}{RESET}")
                            print(f"{INFO_COLOR}Description: {RESET}{tool.description}")
                            
                            if hasattr(tool, 'input_schema'):
                                print(f"{INFO_COLOR}Input Schema:{RESET}")
                                print(json.dumps(tool.input_schema, indent=2))
            
            elif command.startswith("call "):
                parts = command[5:].strip().split(maxsplit=2)
                if len(parts) < 2:
                    print(f"{ERROR_COLOR}Usage: call <server> <tool> [JSON arguments]{RESET}")
                    continue
                    
                server_name = parts[0]
                tool_name = parts[1]
                arguments = {}
                
                if len(parts) > 2:
                    try:
                        arguments = json.loads(parts[2])
                    except json.JSONDecodeError:
                        print(f"{ERROR_COLOR}Invalid JSON arguments. Please provide valid JSON.{RESET}")
                        continue
                else:
                    # Interactive argument input
                    tool = client.get_tool_details(server_name, tool_name)
                    if not tool or not hasattr(tool, 'input_schema'):
                        print(f"{ERROR_COLOR}Tool {tool_name} not found or has no schema.{RESET}")
                        continue
                    
                    if 'properties' in tool.input_schema:
                        properties = tool.input_schema['properties']
                        required = tool.input_schema.get('required', [])
                        
                        print(f"\n{INFO_COLOR}Enter arguments for {HIGHLIGHT_COLOR}{tool_name}{RESET}:")
                        for prop_name, prop_details in properties.items():
                            is_required = prop_name in required
                            prop_type = prop_details.get('type', 'string')
                            default = prop_details.get('default', None)
                            description = prop_details.get('description', '')
                            
                            if description:
                                print(f"  {INFO_COLOR}{description}{RESET}")
                                
                            prompt = f"  {HIGHLIGHT_COLOR}{prop_name}{RESET} ({prop_type}"
                            if is_required:
                                prompt += f", {ERROR_COLOR}required{RESET}"
                            if default is not None:
                                prompt += f", default: {default}"
                            prompt += "): "
                            
                            value = input(prompt)
                            
                            if value:
                                try:
                                    if prop_type == 'integer':
                                        arguments[prop_name] = int(value)
                                    elif prop_type == 'number':
                                        arguments[prop_name] = float(value)
                                    elif prop_type == 'boolean':
                                        arguments[prop_name] = value.lower() in ('true', 'yes', 'y', '1')
                                    elif prop_type == 'array':
                                        arguments[prop_name] = json.loads(value)
                                    elif prop_type == 'object':
                                        arguments[prop_name] = json.loads(value)
                                    else:
                                        arguments[prop_name] = value
                                except (ValueError, json.JSONDecodeError) as e:
                                    print(f"{ERROR_COLOR}Invalid value for {prop_name}: {e}{RESET}")
                                    print(f"{WARNING_COLOR}Tool call aborted.{RESET}")
                                    continue
                            elif is_required:
                                print(f"{ERROR_COLOR}Error: {prop_name} is required.{RESET}")
                                continue
                            elif default is not None:
                                arguments[prop_name] = default
                
                result = await client.call_tool(server_name, tool_name, arguments)
                if result:
                    print(f"\n{HIGHLIGHT_COLOR}Result:{RESET}")
                    print(json.dumps(result, indent=2))
                
            else:
                print(f"{WARNING_COLOR}Unknown command: {command}{RESET}")
                print("Type 'help' for available commands")
                
        except KeyboardInterrupt:
            print(f"\n{WARNING_COLOR}Interrupted. Type 'exit' to quit.{RESET}")
        except Exception as e:
            print(f"{ERROR_COLOR}Error: {e}{RESET}")
            logger.error(f"Interactive session error: {e}", exc_info=True)

async def main():
    """Main function to demonstrate MCP client usage."""
    # Create client instance
    client = MCPClient()
    
    # List available servers from config
    servers = client.list_available_servers()
    if not servers:
        print(f"{WARNING_COLOR}No servers defined in configuration file.{RESET}")
        sys.exit(1)
        
    print(f"{INFO_COLOR}Found {len(servers)} servers in configuration.{RESET}")
    
    # Connect to all servers in config
    for server_name in servers:
        await client.connect_to_server(server_name)
    
    try:
        # Run interactive mode
        await interactive_session(client)
    except Exception as e:
        logger.error(f"Error in interactive mode: {e}", exc_info=True)
        print(f"{ERROR_COLOR}An error occurred in interactive mode: {e}{RESET}")
    finally:
        await client.close()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print(f"\n{INFO_COLOR}Program interrupted by user.{RESET}")
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        print(f"{ERROR_COLOR}Unexpected error: {e}{RESET}") 