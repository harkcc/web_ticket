import asyncio
import sys
import os
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constants
PROJECT_DIR = "/Users/chenminghui/py/lingxing_request/web_ticket"
SERVER_SCRIPT = os.path.join(PROJECT_DIR, "mcp_server.py")

async def run():
    # Prepare environment
    env = os.environ.copy()
    env["PYTHONPATH"] = PROJECT_DIR + os.pathsep + env.get("PYTHONPATH", "")
    # Force development environment to use SSH tunnel
    env["DEPLOY_ENV"] = "development"

    # Define server parameters
    server_params = StdioServerParameters(
        command=sys.executable,
        args=[SERVER_SCRIPT],
        env=env
    )

    logger.info(f"Starting server: {SERVER_SCRIPT} with DEPLOY_ENV=development")
    
    try:
        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                # Initialize the connection
                await session.initialize()
                logger.info("Initialized connection to MCP server")

                # List tools
                tools = await session.list_tools()
                logger.info(f"Available tools: {[t.name for t in tools.tools]}")
                
                # List resources
                resources = await session.list_resources()
                logger.info(f"Available resources: {[r.uri for r in resources.resources]}")

                # Call list_collections tool
                logger.info("Calling list_collections...")
                collections = await session.call_tool("list_collections", {})
                logger.info(f"Result: {collections.content[0].text[:500]}")

    except Exception as e:
        logger.error(f"Test failed with error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(run())
