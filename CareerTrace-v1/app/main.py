from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from .database import init_db
from .api.routes import router

app = FastAPI(title="CareerTrace", version="1.0.0")
init_db()
app.include_router(router, prefix="/api")
app.mount("/", StaticFiles(directory="frontend", html=True), name="frontend")
