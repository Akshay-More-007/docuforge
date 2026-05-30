from dotenv import load_dotenv
load_dotenv()

import asyncio
from src.llm.router import get_llm

async def test():
    llm = get_llm(task="chat")
    r = await llm.ainvoke("Say hello")
    print(r.content)

asyncio.run(test())