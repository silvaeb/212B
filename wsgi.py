from dotenv import load_dotenv

load_dotenv()

from app import app, init_database

init_database()
application = app
