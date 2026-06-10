# Deploying DocuForge

DocuForge ships as a **single container**: the Streamlit UI runs the LangGraph agent
graph in-process. There is no separate API service to deploy. External dependencies
are all SaaS: Supabase (auth), Groq + Gemini (LLMs), Tavily (search).

## Image

- `docker/Dockerfile.docuforge` — python:3.10-slim, deps from `requirements-deploy.txt`
  (exact pins matching the verified dev environment), FastEmbed BGE model pre-baked,
  build-time smoke test that imports the compiled agent graph.
- `compose.docuforge.yaml` — one service on port 8501, `.env` for secrets, named
  volume `docuforge_faiss` for persistent semantic memory.

## Run anywhere with Docker

```bash
# 1. Put real secrets in .env (see .env.docuforge.example)
# 2. Build + start
docker compose -f compose.docuforge.yaml up -d --build
# 3. Verify
curl -f http://localhost:8501/healthz
```

## Deploy to a Linux VM (EC2 / Oracle / any box)

1. Install Docker (`curl -fsSL https://get.docker.com | sh`).
2. Copy the repo (git clone) and a real `.env` to the box
   (`scp .env user@host:~/docuforge/.env` — `.env` is git-ignored, never commit it).
3. `docker compose -f compose.docuforge.yaml up -d --build`
4. Open port 8501 in the security group / firewall (or put nginx/caddy with TLS in
   front and only expose 443 — recommended since login credentials transit this).

Sizing: ~1 GB RAM for the app (FastEmbed ONNX model + Streamlit). A 2 GB VM is
comfortable; 1 GB works with swap.

## Security notes

- Signup is restricted by `ALLOWED_EMAILS` and **fails closed** — if the variable is
  empty, account creation is disabled entirely.
- The Supabase **service-role key** is in `.env`; treat the box accordingly.
- Streamlit serves plain HTTP. For internet exposure put TLS in front
  (Caddy: two-line config, automatic Let's Encrypt).

## Persistent state

| State            | Where                          | Survives restart?                  |
| ---------------- | ------------------------------ | ---------------------------------- |
| Accounts/auth    | Supabase (SaaS)                | yes                                |
| Semantic memory  | volume `docuforge_faiss`       | yes (named volume)                 |
| Uploads/outputs  | container tmp dir              | no — by design, ephemeral          |
