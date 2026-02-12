from fastapi import FastAPI

from app.api.convert import router as convert_router


app = FastAPI(title="PPT/PPTX to Markdown", version="0.1.0")

# Keep the same routes (/health, /convert) while moving code under app/api
app.include_router(convert_router)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=True)
