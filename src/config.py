"""Secrets/config lookup: .env / environment variables first, then .streamlit/secrets.toml."""
import os
from dotenv import load_dotenv

load_dotenv()


def get_secret(name, default=None):
    value = os.getenv(name)
    if value:
        return value
    try:
        import streamlit as st
        value = st.secrets.get(name)
    except Exception:
        # No secrets.toml (e.g. running a generator script directly)
        value = None
    return value or default
