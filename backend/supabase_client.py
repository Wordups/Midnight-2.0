import os
from supabase import create_client

SUPABASE_URL = os.getenv("SUPABASE_URL", "https://devphpxqlpcvsuuzkwyo.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRlcHZocHhxbHBjdnN1dXprd3lvIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NTQyOTU0OSwiZXhwIjoyMDkxMDA1NTQ5fQ.HDoNM9zDYYVK5BnDW-uZAgLVPIDcq857ijN2-kGzd78")

supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
