import asyncio
from typing import Optional, List, Dict, Any
from contextlib import AsyncExitStack
import os
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

API_KEY = os.getenv("GOOGLE_API_KEY")
MODEL = os.getenv("MODEL", "gemini-2.0-flash")


class MCPClient:
    def __init__(self):
        self.session: Optional[ClientSession] = None
        self.model = MODEL
        self.messages = []  # Store conversation history
        self.tools = []
        self.llm = None
        self.exit_stack = AsyncExitStack()
        if API_KEY:
            genai.configure(api_key=API_KEY)
            self.llm = genai.GenerativeModel(self.model)
        else:
            raise ValueError("API_KEY is not set in the environment variables.")
        
    async def connect_to_server(self, server_script_path):
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
        
        try:
            stdio_transport = await self.exit_stack.enter_async_context(stdio_client(server_params))
            self.stdio, self.write = stdio_transport
            self.session = await self.exit_stack.enter_async_context(ClientSession(self.stdio, self.write))
            
            await self.session.initialize()
            
            # List available tools
            response = await self.session.list_tools()
            self.tools = [
                {
                    "name": tool.name,
                    "description": tool.description,
                    "input_schema": tool.inputSchema
                }
                for tool in response.tools
            ]
            print(f"Connected to MCP server with tools: {[tool['name'] for tool in self.tools]}")
            return True 
        except Exception as e:
            print(f"Error connecting to MCP server: {e}")
            return False
    
    def _convert_tools_to_gemini_format(self):
        """Convert MCP tools to Gemini function calling format"""
        if not self.tools:
            return None
            
        gemini_tools = []
        for tool in self.tools:
            # Clean and convert input_schema for Gemini
            input_schema = tool["input_schema"].copy() if tool["input_schema"] else {}
            
            # Remove fields that Gemini doesn't support
            unsupported_fields = ["title", "$schema", "additionalProperties", "examples"]
            for field in unsupported_fields:
                input_schema.pop(field, None)
            
            # Ensure required fields
            if "type" not in input_schema:
                input_schema["type"] = "object"
            if "properties" not in input_schema:
                input_schema["properties"] = {}
            
            # Clean properties recursively
            if "properties" in input_schema:
                cleaned_properties = {}
                for prop_name, prop_schema in input_schema["properties"].items():
                    if isinstance(prop_schema, dict):
                        cleaned_prop = prop_schema.copy()
                        # Remove unsupported fields from property schemas
                        for field in unsupported_fields:
                            cleaned_prop.pop(field, None)
                        cleaned_properties[prop_name] = cleaned_prop
                    else:
                        cleaned_properties[prop_name] = prop_schema
                input_schema["properties"] = cleaned_properties
                
            gemini_tool = {
                "function_declarations": [{
                    "name": tool["name"],
                    "description": tool["description"],
                    "parameters": input_schema
                }]
            }
            gemini_tools.append(gemini_tool)
        
        return gemini_tools
    
    def _format_messages_for_gemini(self):
        """Format conversation history for Gemini with proper function calling structure"""
        if not self.messages:
            return ""
        
        # For Gemini function calling, we need to be more structured
        # Build the conversation parts properly
        conversation_parts = []
        
        for msg in self.messages:
            role = msg.get("role", "")
            content = msg.get("content", "")
            
            if role == "user":
                conversation_parts.append(f"User: {content}")
            elif role == "assistant":
                conversation_parts.append(f"Assistant: {content}")
            elif role == "function":
                # For function results, include them in a way that's clear to the LLM
                function_name = msg.get("name", "unknown")
                conversation_parts.append(f"Function {function_name} returned: {content}")
        
        full_conversation = "\n".join(conversation_parts)
        
        # Add explicit instruction to prevent loops
        if any(msg.get("role") == "function" for msg in self.messages):
            full_conversation += "\n\nPlease provide a final answer to the user based on the function results above. Do not call the same function again."
        
        return full_conversation
        
    def debug_tools_schema(self):
        """Debug method to inspect tool schemas"""
        print("\n=== DEBUG: Tool Schemas ===")
        for i, tool in enumerate(self.tools):
            print(f"\nTool {i+1}: {tool['name']}")
            print(f"Description: {tool['description']}")
            print(f"Input Schema: {tool['input_schema']}")
        
        print("\n=== DEBUG: Converted Gemini Tools ===")
        gemini_tools = self._convert_tools_to_gemini_format()
        if gemini_tools:
            for i, tool in enumerate(gemini_tools):
                print(f"\nGemini Tool {i+1}:")
                print(f"Function Declaration: {tool['function_declarations'][0]}")
        print("=== END DEBUG ===\n")
        
    async def __call_llm(self):
        """Call the LLM with conversation history and tools"""
        if not self.llm:
            raise RuntimeError("LLM is not configured. Please set the API_KEY environment variable.")
        
        try:
            gemini_tools = self._convert_tools_to_gemini_format()
            
            # Format the conversation for Gemini
            if len(self.messages) > 1:
                # Use full conversation context with explicit instructions
                prompt = self._format_messages_for_gemini()
            else:
                # First message
                prompt = self.messages[-1]["content"] if self.messages else ""
            
            # Configure tool calling
            tool_config = None
            if gemini_tools:
                tool_config = {
                    'function_calling_config': {
                        'mode': 'AUTO'  # Let Gemini decide when to use tools
                    }
                }
            
            print(f"DEBUG: Sending prompt: {prompt[:200]}...")
            print(f"DEBUG: Using {len(gemini_tools) if gemini_tools else 0} tools")
            
            response = self.llm.generate_content(
                prompt,
                tools=gemini_tools,
                tool_config=tool_config,
                # Add generation config to encourage finishing
                generation_config={
                    'temperature': 0.1,  # Lower temperature for more consistent behavior
                    'max_output_tokens': 1000,
                }
            )
            
            return response
            
        except Exception as e:
            print(f"Error calling LLM: {e}")
            # Print more detailed error info
            import traceback
            print(f"Full traceback: {traceback.format_exc()}")
            return None

    
    async def process_query(self, prompt: str):
        """Process a prompt using the LLM and available tools"""
        if not self.session:
            raise RuntimeError("Not connected to any MCP server.")
        
        try:
            # Add user message to conversation history
            user_message = {
                "role": "user",
                "content": prompt
            }
            self.messages.append(user_message)
            
            max_iterations = 5  # Prevent infinite loops
            iteration = 0
            
            while iteration < max_iterations:
                iteration += 1
                response = await self.__call_llm()
                
                if not response:
                    print("No response from LLM")
                    break
                
                # Check if response has candidates
                if not hasattr(response, 'candidates') or not response.candidates:
                    print("No candidates in response")
                    break
                
                candidate = response.candidates[0]
                
                # Check for finish reason
                if hasattr(candidate, 'finish_reason'):
                    print(f"Finish reason: {candidate.finish_reason}")
                
                if not hasattr(candidate, 'content') or not hasattr(candidate.content, 'parts'):
                    print("No content parts in response")
                    break
                
                text_parts = []
                function_calls = []
                
                # Process all parts in the response
                for part in candidate.content.parts:
                    if hasattr(part, 'text') and part.text:
                        text_parts.append(part.text)
                    elif hasattr(part, 'function_call'):
                        function_calls.append(part.function_call)
                
                # Add any text response to conversation
                if text_parts:
                    assistant_message = {
                        "role": "assistant",
                        "content": " ".join(text_parts)
                    }
                    self.messages.append(assistant_message)
                
                # Handle function calls
                if function_calls:
                    print(f"Processing {len(function_calls)} function call(s)")
                    
                    for function_call in function_calls:
                        tool_name = function_call.name
                        tool_args = dict(function_call.args)
                        
                        print(f"Calling tool: {tool_name} with args: {tool_args}")
                        
                        try:
                            # Call the MCP tool
                            result = await self.session.call_tool(tool_name, tool_args)
                            
                            # Add function result to conversation history
                            function_result_message = {
                                "role": "function",
                                "name": tool_name,
                                "content": str(result.content)
                            }
                            self.messages.append(function_result_message)
                            
                            print(f"Tool {tool_name} returned: {result.content}")
                            
                        except Exception as e:
                            print(f"Error calling tool {tool_name}: {e}")
                            # Add error to conversation
                            error_message = {
                                "role": "function",
                                "name": tool_name,
                                "content": f"Error: {str(e)}"
                            }
                            self.messages.append(error_message)
                    
                    # Continue the conversation to get LLM's response to tool results
                    continue
                else:
                    # No function calls, conversation is complete
                    break
                        
            return self.messages
        
        except Exception as e:
            print(f"Error processing query: {e}")
            return None
    
    async def get_mcp_tools(self):
        """Get the list of tools available on the MCP server"""
        if not self.session:
            raise RuntimeError("Not connected to any MCP server.")
        
        response = await self.session.list_tools()
        return response.tools
    
    def clear_conversation(self):
        """Clear the conversation history"""
        self.messages = []
        print("Conversation history cleared")
    
    def get_conversation_history(self):
        """Get the current conversation history"""
        return self.messages.copy()
    
    async def cleanup(self):
        """Clean up resources"""
        await self.exit_stack.aclose()