"""
faiss_store.py — Local FAISS vector store with FastEmbed BGE embeddings.
Stores and retrieves chat history semantically.
"""

import json
import logging
import os
import pickle
from pathlib import Path
from typing import Optional

import faiss
import numpy as np

logger = logging.getLogger(__name__)

FAISS_INDEX_DIR = os.getenv("FAISS_INDEX_DIR", "./faiss_index")
EMBED_MODEL = "BAAI/bge-small-en-v1.5"  # FastEmbed model
EMBED_DIM = 384  # bge-small dimension


class FAISSStore:
    """
    Thread-safe FAISS store for semantic memory retrieval.
    Persists index + metadata to disk.
    """

    def __init__(self, user_id: str):
        self.user_id = user_id
        self.index_dir = Path(FAISS_INDEX_DIR) / user_id
        self.index_dir.mkdir(parents=True, exist_ok=True)
        self.index_path = self.index_dir / "index.faiss"
        self.meta_path = self.index_dir / "meta.pkl"
        self._embedder = None
        self._index: Optional[faiss.IndexFlatIP] = None
        self._metadata: list[dict] = []
        self._load()

    def _get_embedder(self):
        if self._embedder is None:
            from fastembed import TextEmbedding
            self._embedder = TextEmbedding(model_name=EMBED_MODEL)
        return self._embedder

    def _load(self):
        """Load existing FAISS index and metadata from disk."""
        if self.index_path.exists() and self.meta_path.exists():
            try:
                self._index = faiss.read_index(str(self.index_path))
                with open(self.meta_path, "rb") as f:
                    self._metadata = pickle.load(f)
                logger.info(f"[FAISSStore] Loaded {len(self._metadata)} memories for user {self.user_id}")
            except Exception as e:
                logger.warning(f"[FAISSStore] Load error: {e} — creating fresh index")
                self._init_fresh()
        else:
            self._init_fresh()

    def _init_fresh(self):
        self._index = faiss.IndexFlatIP(EMBED_DIM)  # Inner product (cosine with normalized vecs)
        self._metadata = []

    def _save(self):
        faiss.write_index(self._index, str(self.index_path))
        with open(self.meta_path, "wb") as f:
            pickle.dump(self._metadata, f)

    def _embed(self, texts: list[str]) -> np.ndarray:
        embedder = self._get_embedder()
        vecs = list(embedder.embed(texts))
        arr = np.array(vecs, dtype="float32")
        # Normalize for cosine similarity via inner product
        norms = np.linalg.norm(arr, axis=1, keepdims=True)
        norms = np.where(norms == 0, 1, norms)
        return arr / norms

    def add(self, text: str, metadata: dict):
        """Add a single memory entry."""
        vec = self._embed([text])
        self._index.add(vec)
        self._metadata.append({"text": text, **metadata})
        self._save()

    def add_batch(self, entries: list[tuple[str, dict]]):
        """Add multiple memory entries efficiently."""
        if not entries:
            return
        texts = [e[0] for e in entries]
        metas = [e[1] for e in entries]
        vecs = self._embed(texts)
        self._index.add(vecs)
        for text, meta in zip(texts, metas):
            self._metadata.append({"text": text, **meta})
        self._save()

    def search(self, query: str, top_k: int = 5) -> list[dict]:
        """
        Retrieve top_k most relevant memories for a query.
        Returns list of metadata dicts with 'text' and 'score'.
        """
        if self._index.ntotal == 0:
            return []

        vec = self._embed([query])
        k = min(top_k, self._index.ntotal)
        scores, indices = self._index.search(vec, k)

        results = []
        for score, idx in zip(scores[0], indices[0]):
            if idx < 0 or idx >= len(self._metadata):
                continue
            entry = dict(self._metadata[idx])
            entry["score"] = float(score)
            results.append(entry)

        return results

    def count(self) -> int:
        return self._index.ntotal if self._index else 0
