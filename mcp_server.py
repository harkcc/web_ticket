from mcp.server.fastmcp import FastMCP
from db_connector import MongoDBConnector, get_mongo_config
import logging
from typing import List, Dict, Any, Optional
import json
from bson import json_util

# Initialize FastMCP application
mcp = FastMCP("MongoDB MCP Server")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def parse_json(data):
    """Helper to dump MongoDB documents to JSON format compatible with MCP."""
    return json.loads(json_util.dumps(data))

@mcp.tool()
def list_collections() -> List[str]:
    """
    List all available collections in the configured MongoDB database.
    """
    with MongoDBConnector() as db:
        return db.list_collection_names()

@mcp.tool()
def query_collection(collection_name: str, query: Dict[str, Any] = {}, limit: int = 10) -> List[Dict[str, Any]]:
    """
    Query a specific collection with a MongoDB find query.
    
    Args:
        collection_name: The name of the collection to query.
        query: MongoDB query dictionary (e.g. {"status": "active"}). Defaults to empty dict (find all).
        limit: Maximum number of documents to return. Defaults to 10.
    """
    with MongoDBConnector() as db:
        collection = db[collection_name]
        cursor = collection.find(query).limit(limit)
        return parse_json(list(cursor))

@mcp.tool()
def get_collection_stats(collection_name: str) -> Dict[str, Any]:
    """
    Get statistics for a specific collection, such as document count.
    """
    with MongoDBConnector() as db:
        collection = db[collection_name]
        count = collection.count_documents({})
        return {
            "collection": collection_name,
            "document_count": count
        }

@mcp.resource("mongo://{collection_name}")
def get_collection_resource(collection_name: str) -> str:
    """
    Read the first 50 documents of a collection as a resource.
    """
    with MongoDBConnector() as db:
        collection = db[collection_name]
        # Limit to 50 for resource reading to prevent overwhelming output
        cursor = collection.find({}).limit(50)
        docs = parse_json(list(cursor))
        return json.dumps(docs, indent=2)

if __name__ == "__main__":
    mcp.run()
