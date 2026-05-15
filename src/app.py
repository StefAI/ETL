from fastapi import FastAPI, UploadFile, Request
from fastapi.templating import Jinja2Templates
import asyncio
import uvicorn
import logging
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
from model import Model

app = FastAPI()

app.mount("/tmp", StaticFiles(directory="tmp", check_dir=False), name="images")
templates = Jinja2Templates(directory="templates")
templates.env.policies["json.dumps_kwargs"] = {"ensure_ascii": False}

app_logger = logging.getLogger(__name__)
app_logger.setLevel(logging.INFO)
app_handler = logging.StreamHandler()
app_formatter = logging.Formatter("%(name)s %(asctime)s %(levelname)s %(message)s")
app_handler.setFormatter(app_formatter)
app_logger.addHandler(app_handler)

@app.get("/health")
def health():
    return {"status": "OK"}


@app.post("/process")
def process_request(file: UploadFile, request: Request):
    save_pth = "tmp/" + file.filename
    app_logger.info(f'processing file - {save_pth}')
    with open(save_pth, "wb") as fid:
        fid.write(file.file.read())
    predictor = Model()
    status, result = predictor(save_pth)
    return templates.TemplateResponse("res_form.html",
                                          {"request": request,
                                           "res": status,
                                           "message": "OK" if status else result,
                                           "json": result if status else None})

@app.get("/")
def main(request: Request):
    return templates.TemplateResponse("start_form.html",
                                      {"request": request})

if __name__ == "__main__":
    host = "0.0.0.0"
    port = 8999

    config = uvicorn.Config(app, host=host, port=port, log_level="info")
    server = uvicorn.Server(config)
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        server.run()
    else:
        asyncio.create_task(server.serve())
