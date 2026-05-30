# DocuForge - src/llm/google_client.py
# Google Gemini 2.5 Flash — fallback when Groq rate limit hit.

import os
from langchain_google_genai import ChatGoogleGenerativeAI


def get_google_llm() -> ChatGoogleGenerativeAI:
    api_key = os.environ["GOOGLE_API_KEY"]

    return ChatGoogleGenerativeAI(
        google_api_key=api_key,
        model="gemini-2.5-flash",
        temperature=0.2,
        max_tokens=4096,
    )
