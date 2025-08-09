import os
import json
import hashlib

# ✅ Persistent location on Windows:
APP_FOLDER = os.path.join(os.getenv("LOCALAPPDATA"), "HypeProduction")
os.makedirs(APP_FOLDER, exist_ok=True)

CONFIG_FILE = os.path.join(APP_FOLDER, "config.json")

DEFAULT_CONFIG = {
    "password": ""
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        print("Failed to load config:", e)
        return DEFAULT_CONFIG

def save_config(data):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print("Failed to save config:", e)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def get_password():
    config = load_config()
    return config.get("password", "")

def set_password(new_password):
    config = load_config()
    config["password"] = hash_password(new_password)
    save_config(config)

def verify_password(entered_password):
    return hash_password(entered_password) == get_password()