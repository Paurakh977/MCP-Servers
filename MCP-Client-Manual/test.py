# server.py
from mcp.server.fastmcp import FastMCP

# Create an MCP server
mcp = FastMCP("Calculator Server")

@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers together"""
    return a + b

@mcp.tool()
def multiply(a: int, b: int) -> int:
    """Multiply two numbers together"""
    return a * b

@mcp.tool()
def calculate_area(length: float, width: float) -> float:
    """Calculate the area of a rectangle"""
    return length * width

@mcp.tool()
def get_weather_info(city: str) -> str:
    """Get mock weather information for a city"""
    # This is just a mock - in real implementation you'd call a weather API
    return f"The weather in {city} is sunny with 25°C temperature."

@mcp.resource("info://server")
def get_server_info() -> str:
    """Get information about this server"""
    return "This is a calculator and utility server with basic math operations."

@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""   
    return f"Hello, {name}! Welcome to the Calculator Server."

if __name__ == "__main__":
    mcp.run()