import os
from supabase import create_client

SUPABASE_URL = os.getenv("SUPABASE_URL", "https://devphpxqlpcvsuuzkwyo.supabase.co")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRlcHZocHhxbHBjdnN1dXprd3lvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU0Mjk1NDksImV4cCI6MjA5MTAwNTU0OX0.2OzdwVVjbGpdT6LyCioAj8bfqL8ZLWHXWkKeWLRGuWQ")

supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
