from supabase import create_client

SUPABASE_URL = "YOUR_URL"
SUPABASE_KEY = "YOUR_ANON_KEY"

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
