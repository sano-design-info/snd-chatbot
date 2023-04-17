import os
import dotenv
import toml

dotenv.load_dotenv()

if os.getenv("DEBUG") == "True":
    CONFIG = toml.load("dev-config.toml")
else:
    CONFIG = toml.load("config.toml")
