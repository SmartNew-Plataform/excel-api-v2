from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import yaml
from app.routers import export, home 

app = FastAPI()

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(home.router)  
app.include_router(export.router)

def get_openapi_spec():
    with open("./openapi.json", "r") as file:
        return yaml.safe_load(file)
app.openapi_schema = get_openapi_spec()
