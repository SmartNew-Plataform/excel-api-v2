from fastapi import APIRouter
from fastapi.responses import HTMLResponse

router = APIRouter()

@router.get("/")
async def welcome():
    html_content = """
    <html>
        <head>
            <title>Bem-vindo ao projeto Excel API</title>
        </head>
        <body>
            <h1>Bem-vindo ao projeto Excel API</h1>
            <p>Explore a documentação interativa:</p>
            <ul>
                <li><a href="/docs">Swagger UI</a></li>
                <li><a href="/redoc">Redoc</a></li>
            </ul>
        </body>
    </html>
    """
    return HTMLResponse(content=html_content)