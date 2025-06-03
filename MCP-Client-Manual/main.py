import asyncio 
from mcp_client import MCPClient

async def main():
    client=MCPClient()
    try:
        await client.connect_to_server(r"D:\MCP\MCP-Client-Manual\test.py")
        response = await client.session.list_tools()
        print("\nConnected to server with tools:", [tool.name for tool in response.tools])
    except Exception as e:
        print(f"Error connecting to server: {e}")
        return  None
    finally:    
        await client.exit_stack.aclose()
        
if __name__ == "__main__":
    asyncio.run(main())