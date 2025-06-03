import asyncio 
from mcp_client import MCPClient

async def main():
    client = MCPClient()
    try:
        # Connect to server
        print("Connecting to MCP server...")
        success = await client.connect_to_server("D:\\MCP\\file-system\\server.py")
        if not success:
            print("Failed to connect to server")
            return
            
        # List tools
        tools = await client.get_mcp_tools()
        print(f"\nAvailable tools: {[tool.name for tool in tools]}")
        
        # Interactive loop
        print("\n" + "="*50)
        print("MCP Client Ready!")
        print("Commands: 'quit' to exit, 'clear' to clear history, 'debug' to show tool schemas, 'history' to show conversation")
        print("="*50)
        
        while True:
            try:
                user_input = input("\nYour query: ").strip()
                
                if user_input.lower() == 'quit':
                    break
                elif user_input.lower() == 'clear':
                    client.clear_conversation()
                    continue
                elif user_input.lower() == 'debug':
                    client.debug_tools_schema()
                    continue
                elif user_input.lower() == 'history':
                    history = client.get_conversation_history()
                    print("\nConversation History:")
                    for msg in history:
                        role = msg.get('role', 'unknown')
                        content = msg.get('content', '')
                        if role == 'function':
                            name = msg.get('name', 'unknown')
                            print(f"  {role} ({name}): {content[:100]}...")
                        else:
                            print(f"  {role}: {content[:100]}...")
                    continue
                elif not user_input:
                    continue
                
                print("\nProcessing query...")
                
                # Process query using the client's method
                messages = await client.process_query(user_input)
                
                print("\n" + "-"*40)
                print("RESPONSE:")
                print("-"*40)
                
                if messages:
                    # Show only the latest assistant response
                    latest_assistant_msg = None
                    for msg in reversed(messages):
                        if msg.get('role') == 'assistant':
                            latest_assistant_msg = msg
                            break
                    
                    if latest_assistant_msg:
                        print(latest_assistant_msg['content'])
                    else:
                        print("No assistant response found")
                        
                    # Optionally show full conversation
                    show_full = input("\nShow full conversation? (y/n): ").lower() == 'y'
                    if show_full:
                        print("\nFull Conversation:")
                        for i, message in enumerate(messages):
                            role = message.get('role', 'unknown')
                            content = message.get('content', '')
                            if role == 'function':
                                name = message.get('name', 'unknown')
                                print(f"{i+1}. {role} ({name}): {content}")
                            else:
                                print(f"{i+1}. {role}: {content}")
                else:
                    print("No response received")
                    
            except KeyboardInterrupt:
                print("\n\nExiting...")
                break
            except Exception as e:
                print(f"\nError processing query: {e}")
                continue
                
    except Exception as e:
        print(f"Fatal error: {e}")
        return None
    finally:    
        print("\nCleaning up...")
        await client.cleanup()
        print("Goodbye! 👋")
        
if __name__ == "__main__":
    asyncio.run(main())