# DocuForge - test_phase1.py
# Run from project root: python test_phase1.py

import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

# Load .env manually (no BOM issues)
env_path = os.path.join(os.path.dirname(__file__), '.env')
with open(env_path, 'rb') as f:
    content = f.read().replace(b'\xef\xbb\xbf', b'').decode('utf-8')
for line in content.splitlines():
    if '=' in line:
        k, v = line.strip().split('=', 1)
        os.environ[k] = v

print("=== DocuForge Phase 1 Smoke Test ===\n")

# Test 1: state.py imports
print("1. Testing state.py import...")
from src.graph.state import AgentState
print("   ✅ AgentState imported\n")

# Test 2: Groq LLM
print("2. Testing Groq (Llama 3.3 70B)...")
from src.llm.groq_client import get_groq_llm
llm = get_groq_llm("default")
response = llm.invoke("Say 'Groq OK' and nothing else.")
print(f"   ✅ Groq response: {response.content}\n")

# Test 3: Google Gemini fallback
print("3. Testing Google Gemini fallback...")
from src.llm.google_client import get_google_llm
llm = get_google_llm()
response = llm.invoke("Say 'Gemini OK' and nothing else.")
print(f"   ✅ Gemini response: {response.content}\n")

# Test 4: Router
print("4. Testing LLM router...")
from src.llm.router import get_llm
llm = get_llm(task="chat")
response = llm.invoke("Say 'Router OK' and nothing else.")
print(f"   ✅ Router response: {response.content}\n")

# Test 5: Supabase connection
print("5. Testing Supabase connection...")
from supabase import create_client
url = os.environ["SUPABASE_URL"]
key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
sb = create_client(url, key)
result = sb.table("chat_sessions").select("id").limit(1).execute()
print(f"   ✅ Supabase connected. chat_sessions table reachable.\n")

print("=== All Phase 1 checks passed ✅ ===")
