from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import yaml
from app.routers import export, home  # Verifique se isso está correto

app = FastAPI()

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(home.router)   # Inclui a rota GET "/"
app.include_router(export.router) # Deve incluir as rotas POST "/api/v1/export" e "/api/v1/export-unified"

def get_openapi_spec():
    with open("./openapi.json", "r") as file:
        return yaml.safe_load(file)
app.openapi_schema = get_openapi_spec()