from fastapi import FastAPI

app = FastAPI(
    title="SaaS Demandas Legal API",
    description="Backend para la automatización de documentos jurídicos",
    version="0.1.0"
)

@app.get("/")
def read_root():
    return {"status": "ok", "message": "El motor de la API está en línea y funcionando."}