"""
DocuForge - seed_users.py
Run once to create approved users in Supabase.
Usage: python scripts/seed_users.py
"""
from supabase import create_client
import os

env_path = os.path.join(os.path.dirname(__file__), '..', '.env')
with open(env_path) as f:
    for line in f:
        if '=' in line:
            k, v = line.strip().split('=', 1)
            os.environ[k] = v

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]

# Admin client (service role key — never expose this publicly)
supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)

APPROVED_USERS = [
    {"email": "akshay.more00741@gmail.com", "password": "ChangeMe@123"},
    # Add friend later:
    # {"email": "friend@email.com", "password": "ChangeMe@123"},
]

for user in APPROVED_USERS:
    response = supabase.auth.admin.create_user({
        "email": user["email"],
        "password": user["password"],
        "email_confirm": True,  # Skip email verification
    })
    print(f"Created: {response.user.email} — ID: {response.user.id}")

print("Done.")
