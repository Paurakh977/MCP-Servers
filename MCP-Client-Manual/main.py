#!/usr/bin/env python
import asyncio
import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

from mistralai import Mistral
from mistralai.extra.run.context import RunContext
from mcp import StdioServerParameters
from mistralai.extra.mcp.stdio import MCPClientSTDIO
from mistralai.types import BaseModel

# Configuration
MODEL = "mistral-medium-latest"
CONFIG_FILE = "config.json"

class MCPClientManager:
    """Manages MCP clients from configuration file"""
    
    def __init__(self, config_path: str = CONFIG_FILE):
        self.config_path = config_path
        self.config = self._load_config()
        
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file"""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"Configuration file {self.config_path} not found")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in configuration file: {e}")
    
    def get_server_parameters(self) -> Dict[str, StdioServerParameters]:
        """Convert config to StdioServerParameters objects"""
        server_params = {}
        
        mcp_servers = self.config.get("mcpServers", {})
        for server_name, server_config in mcp_servers.items():
            command = server_config.get("command")
            args = server_config.get("args", [])
            env = server_config.get("env")
            
            if not command:
                print(f"Warning: No command specified for server '{server_name}', skipping...")
                continue
                
            server_params[server_name] = StdioServerParameters(
                command=command,
                args=args,
                env=env,
            )
            
        return server_params

class MistralMCPAgent:
    """Main agent class that integrates Mistral with MCP servers"""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get("MISTRAL_API_KEY")
        if not self.api_key:
            raise ValueError("MISTRAL_API_KEY environment variable is required")
            
        self.client = Mistral(self.api_key)
        self.mcp_manager = MCPClientManager()
        
    async def create_agent(self, 
                          name: str = "MCP Assistant", 
                          instructions: str = "You are a helpful assistant with access to various tools and services.",
                          description: str = "") -> Any:
        """Create a Mistral agent"""
        return self.client.beta.agents.create(
            model=MODEL,
            name=name,
            instructions=instructions,
            description=description,
        )
    
    async def run_with_mcp(self, 
                          query: str, 
                          output_format: Optional[BaseModel] = None,
                          agent_name: str = "MCP Assistant",
                          agent_instructions: str = "You are a helpful assistant with access to file system operations and other tools. Help users with their requests using the available tools.",
                          stream: bool = False) -> Any:
        """Run a query with MCP server integration"""
        
        # Create agent
        agent = await self.create_agent(
            name=agent_name,
            instructions=agent_instructions
        )
        
        # Get server parameters
        server_params = self.mcp_manager.get_server_parameters()
        
        if not server_params:
            raise ValueError("No valid MCP servers found in configuration")
        
        # Create run context
        async with RunContext(
            agent_id=agent.id,
            output_format=output_format,
            continue_on_fn_error=True,
        ) as run_ctx:
            
            # Register all MCP clients
            mcp_clients = []
            for server_name, params in server_params.items():
                try:
                    print(f"Connecting to MCP server: {server_name}")
                    mcp_client = MCPClientSTDIO(stdio_params=params)
                    await run_ctx.register_mcp_client(mcp_client=mcp_client)
                    mcp_clients.append((server_name, mcp_client))
                    print(f"Successfully connected to {server_name}")
                except Exception as e:
                    print(f"Failed to connect to {server_name}: {e}")
                    continue
            
            if not mcp_clients:
                raise RuntimeError("Failed to connect to any MCP servers")
            
            # Run the query
            if stream:
                return await self._run_stream(run_ctx, query)
            else:
                return await self._run_sync(run_ctx, query)
    
    async def _run_sync(self, run_ctx: RunContext, query: str) -> Any:
        """Run synchronously and return complete result"""
        run_result = await self.client.beta.conversations.run_async(
            run_ctx=run_ctx,
            inputs=query,
        )
        
        return run_result
    
    async def _run_stream(self, run_ctx: RunContext, query: str) -> Any:
        """Run with streaming responses"""
        events = await self.client.beta.conversations.run_stream_async(
            run_ctx=run_ctx,
            inputs=query,
        )
        
        run_result = None
        async for event in events:
            if hasattr(event, '__class__') and 'RunResult' in str(event.__class__):
                run_result = event
            else:
                print(f"Event: {event}")
        
        return run_result

# Example usage functions
async def main():
    """Main function demonstrating usage"""
    try:
        # Initialize the agent
        agent = MistralMCPAgent()
        
        # Example query
        query = "List the files in the current directory and tell me about any Python files you find."
        
        print(f"Running query: {query}")
        print("-" * 50)
        
        # Run the query
        result = await agent.run_with_mcp(
            query=query,
            agent_instructions="You are a helpful file system assistant. Use the available file system tools to help users manage and explore their files. Provide clear and detailed responses about file operations.",
            stream=False  # Set to True for streaming
        )
        
        # Print results
        print("Results:")
        print("-" * 30)
        
        if hasattr(result, 'output_entries'):
            for entry in result.output_entries:
                print(f"{entry}")
                print()
        
        if hasattr(result, 'output_as_model') and result.output_as_model:
            print(f"Structured output: {result.output_as_model}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

async def interactive_mode():
    """Interactive mode for continuous queries"""
    agent = MistralMCPAgent()
    
    print("MCP Interactive Mode")
    print("Type 'quit' to exit")
    print("-" * 30)
    
    while True:
        try:
            query = input("\nEnter your query: ").strip()
            
            if query.lower() in ['quit', 'exit', 'q']:
                break
                
            if not query:
                continue
            
            print(f"\nProcessing: {query}")
            print("-" * 20)
            
            result = await agent.run_with_mcp(query=query)
            
            if hasattr(result, 'output_entries'):
                for entry in result.output_entries:
                    print(f"{entry}")
            
        except KeyboardInterrupt:
            print("\nExiting...")
            break
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "interactive":
        asyncio.run(interactive_mode())
    else:
        asyncio.run(main())