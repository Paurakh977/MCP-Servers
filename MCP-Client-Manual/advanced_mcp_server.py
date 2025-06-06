# advanced_mcp_server.py
from mcp.server.fastmcp import FastMCP
import json
import random
import hashlib
import base64
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import re
import uuid
import sys
# Create an advanced MCP server
mcp = FastMCP("Advanced Multi-Tool Server")

# In-memory storage for demonstration
task_storage = {}
note_storage = {}
user_profiles = {}

# =============================================================================
# TEXT & STRING MANIPULATION TOOLS
# =============================================================================

@mcp.tool()
def analyze_text(text: str) -> Dict[str, Any]:
    """Analyze text and return detailed statistics including word count, character count, readability metrics"""
    words = text.split()
    sentences = len([s for s in re.split(r'[.!?]+', text) if s.strip()])
    paragraphs = len([p for p in text.split('\n\n') if p.strip()])
    
    # Basic readability estimate (Flesch formula approximation)
    avg_sentence_length = len(words) / max(sentences, 1)
    
    return {
        "word_count": len(words),
        "character_count": len(text),
        "character_count_no_spaces": len(text.replace(' ', '')),
        "sentence_count": sentences,
        "paragraph_count": paragraphs,
        "average_words_per_sentence": round(avg_sentence_length, 2),
        "estimated_reading_time_minutes": round(len(words) / 200, 1),  # 200 WPM average
        "most_common_words": get_most_common_words(words, 5)
    }

def get_most_common_words(words: List[str], top_n: int = 5) -> List[Dict[str, Any]]:
    """Helper function to get most common words"""
    word_freq = {}
    for word in words:
        clean_word = re.sub(r'[^\w]', '', word.lower())
        if len(clean_word) > 2:  # Skip short words
            word_freq[clean_word] = word_freq.get(clean_word, 0) + 1
    
    sorted_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)
    return [{"word": word, "count": count} for word, count in sorted_words[:top_n]]

@mcp.tool()
def generate_password(length: int = 12, include_symbols: bool = True, include_numbers: bool = True) -> Dict[str, str]:
    """Generate a secure random password with customizable options"""
    import string
    
    characters = string.ascii_letters
    if include_numbers:
        characters += string.digits
    if include_symbols:
        characters += "!@#$%^&*()_+-=[]{}|;:,.<>?"
    
    password = ''.join(random.choice(characters) for _ in range(length))
    
    # Calculate strength
    strength_score = 0
    if any(c.islower() for c in password):
        strength_score += 1
    if any(c.isupper() for c in password):
        strength_score += 1
    if any(c.isdigit() for c in password):
        strength_score += 1
    if any(c in "!@#$%^&*()_+-=[]{}|;:,.<>?" for c in password):
        strength_score += 1
    if length >= 12:
        strength_score += 1
    
    strength_levels = ["Very Weak", "Weak", "Fair", "Good", "Strong", "Very Strong"]
    strength = strength_levels[min(strength_score, 5)]
    
    return {
        "password": password,
        "length": length,
        "strength": strength,
        "entropy_bits": round(length * 4.7, 1)  # Rough estimate
    }

@mcp.tool()
def encode_decode_text(text: str, operation: str, encoding_type: str = "base64") -> Dict[str, str]:
    """Encode or decode text using various encoding methods (base64, hex, url, etc.)"""
    try:
        if operation.lower() == "encode":
            if encoding_type.lower() == "base64":
                result = base64.b64encode(text.encode('utf-8')).decode('utf-8')
            elif encoding_type.lower() == "hex":
                result = text.encode('utf-8').hex()
            elif encoding_type.lower() == "url":
                import urllib.parse
                result = urllib.parse.quote(text)
            else:
                return {"error": f"Unsupported encoding type: {encoding_type}"}
        
        elif operation.lower() == "decode":
            if encoding_type.lower() == "base64":
                result = base64.b64decode(text.encode('utf-8')).decode('utf-8')
            elif encoding_type.lower() == "hex":
                result = bytes.fromhex(text).decode('utf-8')
            elif encoding_type.lower() == "url":
                import urllib.parse
                result = urllib.parse.unquote(text)
            else:
                return {"error": f"Unsupported encoding type: {encoding_type}"}
        else:
            return {"error": "Operation must be 'encode' or 'decode'"}
        
        return {
            "operation": operation,
            "encoding_type": encoding_type,
            "input": text,
            "result": result
        }
    except Exception as e:
        return {"error": f"Failed to {operation}: {str(e)}"}

# =============================================================================
# DATA GENERATION & MOCK DATA TOOLS
# =============================================================================

@mcp.tool()
def generate_mock_user_data(count: int = 1) -> List[Dict[str, Any]]:
    """Generate realistic mock user profiles for testing"""
    first_names = ["Emma", "Liam", "Olivia", "Noah", "Ava", "Ethan", "Sophia", "Mason", "Isabella", "William"]
    last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez"]
    domains = ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "example.com"]
    
    users = []
    for _ in range(count):
        first_name = random.choice(first_names)
        last_name = random.choice(last_names)
        user_id = str(uuid.uuid4())
        
        user = {
            "id": user_id,
            "first_name": first_name,
            "last_name": last_name,
            "full_name": f"{first_name} {last_name}",
            "email": f"{first_name.lower()}.{last_name.lower()}@{random.choice(domains)}",
            "age": random.randint(18, 80),
            "phone": f"+1-{random.randint(200, 999)}-{random.randint(100, 999)}-{random.randint(1000, 9999)}",
            "address": {
                "street": f"{random.randint(1, 9999)} {random.choice(['Main', 'Oak', 'Pine', 'Elm', 'Cedar'])} St",
                "city": random.choice(["New York", "Los Angeles", "Chicago", "Houston", "Phoenix"]),
                "state": random.choice(["NY", "CA", "IL", "TX", "AZ"]),
                "zip_code": f"{random.randint(10000, 99999)}"
            },
            "created_at": (datetime.now() - timedelta(days=random.randint(1, 365))).isoformat(),
            "is_active": random.choice([True, True, True, False])  # 75% active
        }
        users.append(user)
    
    return users

@mcp.tool()
def generate_lorem_ipsum(paragraphs: int = 1, words_per_paragraph: int = 50) -> Dict[str, Any]:
    """Generate Lorem Ipsum placeholder text"""
    lorem_words = [
        "lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing", "elit",
        "sed", "do", "eiusmod", "tempor", "incididunt", "ut", "labore", "et", "dolore",
        "magna", "aliqua", "enim", "ad", "minim", "veniam", "quis", "nostrud",
        "exercitation", "ullamco", "laboris", "nisi", "aliquip", "ex", "ea", "commodo",
        "consequat", "duis", "aute", "irure", "in", "reprehenderit", "voluptate",
        "velit", "esse", "cillum", "fugiat", "nulla", "pariatur", "excepteur", "sint",
        "occaecat", "cupidatat", "non", "proident", "sunt", "culpa", "qui", "officia",
        "deserunt", "mollit", "anim", "id", "est", "laborum"
    ]
    
    generated_paragraphs = []
    for _ in range(paragraphs):
        paragraph_words = []
        for _ in range(words_per_paragraph):
            paragraph_words.append(random.choice(lorem_words))
        
        # Capitalize first word and add proper punctuation
        paragraph = " ".join(paragraph_words)
        paragraph = paragraph.capitalize()
        
        # Add some periods randomly
        words = paragraph.split()
        for i in range(5, len(words), random.randint(8, 15)):
            if i < len(words):
                words[i] += "."
                if i + 1 < len(words):
                    words[i + 1] = words[i + 1].capitalize()
        
        paragraph = " ".join(words)
        if not paragraph.endswith('.'):
            paragraph += "."
        
        generated_paragraphs.append(paragraph)
    
    return {
        "paragraphs": generated_paragraphs,
        "word_count": paragraphs * words_per_paragraph,
        "paragraph_count": paragraphs,
        "full_text": "\n\n".join(generated_paragraphs)
    }

# =============================================================================
# TASK MANAGEMENT TOOLS
# =============================================================================

@mcp.tool()
def create_task(title: str, description: str = "", priority: str = "medium", due_date: str = "") -> Dict[str, Any]:
    """Create a new task with title, description, priority, and optional due date"""
    task_id = str(uuid.uuid4())
    
    task = {
        "id": task_id,
        "title": title,
        "description": description,
        "priority": priority.lower(),
        "status": "pending",
        "created_at": datetime.now().isoformat(),
        "due_date": due_date,
        "completed_at": None
    }
    
    task_storage[task_id] = task
    
    return {
        "message": "Task created successfully",
        "task": task,
        "total_tasks": len(task_storage)
    }

@mcp.tool()
def list_tasks(status: str = "all", priority: str = "all") -> Dict[str, Any]:
    """List all tasks with optional filtering by status and priority"""
    filtered_tasks = []
    
    for task in task_storage.values():
        if status != "all" and task["status"] != status.lower():
            continue
        if priority != "all" and task["priority"] != priority.lower():
            continue
        filtered_tasks.append(task)
    
    # Sort by priority and creation date
    priority_order = {"high": 0, "medium": 1, "low": 2}
    filtered_tasks.sort(key=lambda x: (priority_order.get(x["priority"], 3), x["created_at"]))
    
    return {
        "tasks": filtered_tasks,
        "total_count": len(filtered_tasks),
        "filters_applied": {"status": status, "priority": priority}
    }

@mcp.tool()
def complete_task(task_id: str) -> Dict[str, Any]:
    """Mark a task as completed"""
    if task_id not in task_storage:
        return {"error": f"Task with ID {task_id} not found"}
    
    task = task_storage[task_id]
    task["status"] = "completed"
    task["completed_at"] = datetime.now().isoformat()
    
    return {
        "message": "Task marked as completed",
        "task": task
    }

# =============================================================================
# FILE & DATA PROCESSING TOOLS
# =============================================================================

@mcp.tool()
def create_json_template(template_name: str, fields: List[str]) -> Dict[str, Any]:
    """Create a JSON template with specified fields"""
    template = {}
    
    for field in fields:
        # Guess field type based on name
        field_lower = field.lower()
        if any(word in field_lower for word in ['id', 'uuid']):
            template[field] = "{{uuid}}"
        elif any(word in field_lower for word in ['email', 'mail']):
            template[field] = "{{email}}"
        elif any(word in field_lower for word in ['name', 'title']):
            template[field] = "{{string}}"
        elif any(word in field_lower for word in ['date', 'time']):
            template[field] = "{{datetime}}"
        elif any(word in field_lower for word in ['age', 'count', 'number']):
            template[field] = "{{integer}}"
        elif any(word in field_lower for word in ['price', 'amount', 'cost']):
            template[field] = "{{float}}"
        elif any(word in field_lower for word in ['active', 'enabled', 'flag']):
            template[field] = "{{boolean}}"
        else:
            template[field] = "{{string}}"
    
    return {
        "template_name": template_name,
        "template": template,
        "pretty_json": json.dumps(template, indent=2),
        "field_count": len(fields)
    }

@mcp.tool()
def hash_data(data: str, algorithm: str = "sha256") -> Dict[str, str]:
    """Generate hash of input data using various algorithms (md5, sha1, sha256, sha512)"""
    try:
        if algorithm.lower() == "md5":
            hash_obj = hashlib.md5(data.encode('utf-8'))
        elif algorithm.lower() == "sha1":
            hash_obj = hashlib.sha1(data.encode('utf-8'))
        elif algorithm.lower() == "sha256":
            hash_obj = hashlib.sha256(data.encode('utf-8'))
        elif algorithm.lower() == "sha512":
            hash_obj = hashlib.sha512(data.encode('utf-8'))
        else:
            return {"error": f"Unsupported algorithm: {algorithm}"}
        
        return {
            "input": data,
            "algorithm": algorithm.lower(),
            "hash": hash_obj.hexdigest(),
            "hash_length": len(hash_obj.hexdigest())
        }
    except Exception as e:
        return {"error": f"Failed to generate hash: {str(e)}"}

# =============================================================================
# UTILITY & CONVERSION TOOLS
# =============================================================================

@mcp.tool()
def convert_units(value: float, from_unit: str, to_unit: str, unit_type: str) -> Dict[str, Any]:
    """Convert between different units (length, weight, temperature, etc.)"""
    conversions = {
        "length": {
            "mm": 1, "cm": 10, "m": 1000, "km": 1000000,
            "inch": 25.4, "ft": 304.8, "yard": 914.4, "mile": 1609344
        },
        "weight": {
            "mg": 1, "g": 1000, "kg": 1000000,
            "oz": 28349.5, "lb": 453592, "stone": 6350293
        },
        "temperature": {
            # Special handling needed for temperature
        }
    }
    
    if unit_type.lower() == "temperature":
        # Handle temperature conversion separately
        if from_unit.lower() == "c" and to_unit.lower() == "f":
            result = (value * 9/5) + 32
        elif from_unit.lower() == "f" and to_unit.lower() == "c":
            result = (value - 32) * 5/9
        elif from_unit.lower() == "c" and to_unit.lower() == "k":
            result = value + 273.15
        elif from_unit.lower() == "k" and to_unit.lower() == "c":
            result = value - 273.15
        elif from_unit.lower() == "f" and to_unit.lower() == "k":
            result = (value - 32) * 5/9 + 273.15
        elif from_unit.lower() == "k" and to_unit.lower() == "f":
            result = (value - 273.15) * 9/5 + 32
        else:
            return {"error": f"Unsupported temperature conversion: {from_unit} to {to_unit}"}
    else:
        if unit_type.lower() not in conversions:
            return {"error": f"Unsupported unit type: {unit_type}"}
        
        unit_conversions = conversions[unit_type.lower()]
        
        if from_unit.lower() not in unit_conversions or to_unit.lower() not in unit_conversions:
            return {"error": f"Unsupported units for {unit_type}: {from_unit} or {to_unit}"}
        
        # Convert to base unit, then to target unit
        base_value = value * unit_conversions[from_unit.lower()]
        result = base_value / unit_conversions[to_unit.lower()]
    
    return {
        "input_value": value,
        "from_unit": from_unit,
        "to_unit": to_unit,
        "unit_type": unit_type,
        "result": round(result, 6),
        "formatted": f"{value} {from_unit} = {round(result, 6)} {to_unit}"
    }

@mcp.tool()
def generate_qr_data(text: str, size: str = "medium") -> Dict[str, Any]:
    """Generate QR code data and information (mock implementation - would need qrcode library for real QR generation)"""
    # This is a mock implementation - in reality you'd use the qrcode library
    size_pixels = {"small": 100, "medium": 200, "large": 400}.get(size.lower(), 200)
    
    return {
        "input_text": text,
        "text_length": len(text),
        "size": size,
        "estimated_pixels": f"{size_pixels}x{size_pixels}",
        "data_capacity": "up to 4,296 alphanumeric characters",
        "note": "This is mock QR data - integrate with qrcode library for actual QR generation",
        "suggested_url": f"https://api.qrserver.com/v1/create-qr-code/?size={size_pixels}x{size_pixels}&data={text}"
    }

# =============================================================================
# RESOURCES
# =============================================================================

@mcp.resource("info://server")
def get_server_info() -> str:
    """Get comprehensive information about this advanced server"""
    return json.dumps({
        "name": "Advanced Multi-Tool Server",
        "version": "2.0.0",
        "description": "A comprehensive MCP server with text analysis, data generation, task management, and utility tools",
        "capabilities": [
            "Text analysis and manipulation",
            "Password generation and security tools",
            "Mock data generation for testing",
            "Task and project management",
            "Data encoding/decoding",
            "Unit conversions",
            "Hash generation",
            "JSON template creation"
        ],
        "total_tools": 12,
        "storage": {
            "tasks": len(task_storage),
            "notes": len(note_storage),
            "users": len(user_profiles)
        }
    }, indent=2)

@mcp.resource("stats://usage")
def get_usage_stats() -> str:
    """Get current usage statistics"""
    return json.dumps({
        "timestamp": datetime.now().isoformat(),
        "storage_stats": {
            "active_tasks": len([t for t in task_storage.values() if t["status"] != "completed"]),
            "completed_tasks": len([t for t in task_storage.values() if t["status"] == "completed"]),
            "total_tasks": len(task_storage)
        },
        "server_uptime": "Session-based (resets on restart)",
        "memory_usage": "In-memory storage only"
    }, indent=2)

if __name__ == "__main__":
    print("🚀 Starting Advanced Multi-Tool MCP Server...")
    print("📋 Available tools: Text Analysis, Password Gen, Mock Data, Task Management, and more!")
    try:
        # Add a debug print to confirm we reach this point
        print("DEBUG: advanced_mcp_server.py: Attempting to call mcp.run()", file=sys.stderr)
        mcp.run()
        # This print should ideally not be reached if the server is running properly
        print("DEBUG: advanced_mcp_server.py: mcp.run() finished (server probably exited early)", file=sys.stderr)
    except Exception as e:
        # Crucial: print any exception to stderr
        print(f"FATAL ERROR in advanced_mcp_server.py during mcp.run(): {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr) # Also print the full traceback
    finally:
        # Indicate that the server process is definitely exiting
        print("DEBUG: advanced_mcp_server.py: Server process exiting.", file=sys.stderr)