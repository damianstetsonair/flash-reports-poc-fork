import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from backend directory
backend_dir = Path(__file__).parent.parent
load_dotenv(backend_dir / ".env")

# API Keys
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY") or os.getenv("SUPABASE_ANON_KEY")

# LibreOffice path (macOS Homebrew default)
SOFFICE_PATH = os.getenv("SOFFICE_PATH", "/opt/homebrew/bin/soffice")

# Temp directory for processing
TEMP_DIR = Path(os.getenv("TEMP_DIR", "/tmp/flash-reports"))
TEMP_DIR.mkdir(parents=True, exist_ok=True)

# Claude model configuration
CLAUDE_MODEL = "claude-opus-4-5-20251101"
CLAUDE_MAX_TOKENS = 32000

# Fast parallel pipeline (uses same model as single-agent)
CLAUDE_MODEL_FAST = "claude-opus-4-5-20251101"
CLAUDE_MAX_TOKENS_FAST = 32000

# Server configuration
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8000"))
DEBUG = os.getenv("DEBUG", "false").lower() == "true"

# CORS
CORS_ORIGINS = os.getenv("CORS_ORIGINS", "http://localhost:5173,http://localhost:3000").split(",")
