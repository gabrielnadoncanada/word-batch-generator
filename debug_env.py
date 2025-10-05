# -*- coding: utf-8 -*-
"""
Debug des variables d'environnement
"""
import os
from dotenv import load_dotenv

print("=== Debug Variables d'Environnement ===")

# Charger .env
load_dotenv()

print(f"SMTP_SERVER: '{os.getenv('SMTP_SERVER')}'")
print(f"SMTP_PORT: '{os.getenv('SMTP_PORT')}'")
print(f"SMTP_USERNAME: '{os.getenv('SMTP_USERNAME')}'")
print(f"SMTP_PASSWORD: '{os.getenv('SMTP_PASSWORD')}'")
print(f"FROM_ACCOUNT: '{os.getenv('FROM_ACCOUNT')}'")

# Test direct
print("\n=== Test direct ===")
print(f"SMTP_SERVER depuis config: {os.getenv('SMTP_SERVER', 'NON_TROUVE')}")
