from dotenv import load_dotenv
load_dotenv()

from src.memory.faiss_store import FAISSStore

store = FAISSStore(user_id="testuser")
store.add("The project deadline is Friday", {"role": "user", "session_id": "s1", "timestamp": ""})
results = store.search("when is the deadline")
print(results)