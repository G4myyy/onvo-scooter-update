import json
import os
from pathlib import Path

class Config:
    _instance = None
    _config = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._config is None:
            self.load()
    
    def load(self):
        config_path = Path(__file__).parent / "config.json"
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                self._config = json.load(f)
        else:
            self._config = {}
        
        if "FIREBASE_URL" in os.environ:
            self._config["firebase_url"] = os.environ["FIREBASE_URL"]
        if "GITHUB_REPO" in os.environ:
            self._config["github_repo"] = os.environ["GITHUB_REPO"]
    
    def get(self, key, default=None):
        return self._config.get(key, default)
    
    @property
    def firebase_url(self):
        return self.get("firebase_url", "https://kaen-onvo-scooter-default-rtdb.firebaseio.com")
    
    @property
    def github_repo(self):
        return self.get("github_repo", "G4myyy/onvo-scooter-update")

config = Config()
