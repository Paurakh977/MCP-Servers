import asyncio
from mcp_client import MCPClient

async def main():
    client = MCPClient()
    server_script_path = "server.py"  # Path to your server script
    
    try:
        # Connect to the server
        await client.connect_to_server(server_script_path)
        
        if client.session:
            print("\n=== Testing MCP Tools ===")
            
            # Test add function
            result = await client.session.call_tool("add", {"a": 10, "b": 5})
            print(f"10 + 5 = {result.content[0].text}")
            
            # Test multiply function
            result = await client.session.call_tool("multiply", {"a": 7, "b": 8})
            print(f"7 * 8 = {result.content[0].text}")
            
            # Test area calculation
            result = await client.session.call_tool("calculate_area", {"length": 12.5, "width": 8.0})
            print(f"Area of rectangle (12.5 x 8.0) = {result.content[0].text}")
            
            # Test weather info
            result = await client.session.call_tool("get_weather_info", {"city": "Kathmandu"})
            print(f"Weather: {result.content[0].text}")
            
            print("\n=== Testing Resources ===")
            
            # Test server info resource
            resources = await client.session.list_resources()
            print(f"Available resources: {[r.uri for r in resources.resources]}")
            
            # Read server info
            server_info = await client.session.read_resource("info://server")
            print(f"Server info: {server_info.contents[0].text}")
            
            # Read greeting resource
            greeting = await client.session.read_resource("greeting://Alice")
            print(f"Greeting: {greeting.contents[0].text}")
            
            print("\n=== Using with Gemini ===")
            
            # Example of using MCP tools with Gemini
            if client.llm:
                # Get calculation result
                calc_result = await client.session.call_tool("add", {"a": 25, "b": 17})
                number = calc_result.content[0].text
                
                # Ask Gemini to explain the result
                prompt = f"The result of 25 + 17 is {number}. Can you explain this calculation and give me a fun fact about the number {number}?"
                response = client.llm.generate_content(prompt)
                print(f"Gemini says: {response.text}")
            
        # Keep running for a bit to see everything
        print("\n=== Connection successful! ===")
        await asyncio.sleep(2)
        
    except KeyboardInterrupt:
        print("\nShutting down...")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Clean up
        await client.exit_stack.aclose()

if __name__ == "__main__":
    asyncio.run(main())