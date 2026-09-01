from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from typing import Dict
import os
import uuid

app = FastAPI(
    title="SMAGGE - Gerador de Petições",
    version="1.0.0",
    servers=[
        {
            "url": "https://peti-o-smagge.onrender.com"
        }
    ]
)

PASTA_SAIDA = "arquivos"
os.makedirs(PASTA_SAIDA, exist_ok=True)


class PedidoPeticao(BaseModel):
    tipo_peticao: str
    nome_cliente: str
    substituicoes: Dict[str, str]


def localizar_modelo(tipo: str):
    arquivos = os.listdir(".")

    tipo = tipo.lower()

    if tipo == "juros":
        for arq in arquivos:
            nome = arq.lower()
            if (
                arq.endswith(".docx")
                and "juros" in nome
                and "abusiv" in nome
            ):
                return arq

    if tipo == "descumprimento":
        for arq in arquivos:
            nome = arq.lower()
            if (
                arq.endswith(".docx")
                and "descumprimento" in nome
            ):
                return arq

    return None


def substituir_em_runs(paragrafo, substituicoes):
    if not paragrafo.runs:
        return

    texto = "".join(run.text for run in paragrafo.runs)
    original = texto

    for antigo, novo in substituicoes.items():
        if antigo:
            texto = texto.replace(antigo, novo)

    if texto == original:
        return

    # mantém a formatação principal do parágrafo
    paragrafo.runs[0].text = texto

    for run in paragrafo.runs[1:]:
        run.text = ""


def processar_documento(doc, substituicoes):

    for paragrafo in doc.paragraphs:
        substituir_em_runs(
            paragrafo,
            substituicoes
        )

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    substituir_em_runs(
                        paragrafo,
                        substituicoes
                    )

    for secao in doc.sections:

        for paragrafo in secao.header.paragraphs:
            substituir_em_runs(
                paragrafo,
                substituicoes
            )

        for tabela in secao.header.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for paragrafo in celula.paragraphs:
                        substituir_em_runs(
                            paragrafo,
                            substituicoes
                        )

        for paragrafo in secao.footer.paragraphs:
            substituir_em_runs(
                paragrafo,
                substituicoes
            )


@app.get("/")
def status():
    return {
        "status": "online",
        "sistema": "SMAGGE Gerador de Petições"
    }


@app.post("/gerar")
def gerar_peticao(pedido: PedidoPeticao):

    modelo = localizar_modelo(
        pedido.tipo_peticao
    )

    if not modelo:
        raise HTTPException(
            status_code=404,
            detail="Modelo da petição não encontrado."
        )

    try:
        doc = Document(modelo)
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Erro ao abrir modelo: {str(e)}"
        )

    processar_documento(
        doc,
        pedido.substituicoes
    )

    nome_seguro = "".join(
        c if c.isalnum() else "_"
        for c in pedido.nome_cliente
    )

    codigo = str(uuid.uuid4())

    nome_arquivo = (
        f"{codigo}_PETICAO_{nome_seguro}.docx"
    )

    caminho = os.path.join(
        PASTA_SAIDA,
        nome_arquivo
    )

    doc.save(caminho)

    return {
        "sucesso": True,
        "arquivo_id": codigo,
        "arquivo": nome_arquivo,
        "download": f"/download/{codigo}"
    }


@app.get("/download/{arquivo_id}")
def baixar_peticao(arquivo_id: str):

    for arquivo in os.listdir(PASTA_SAIDA):

        if arquivo.startswith(arquivo_id):

            caminho = os.path.join(
                PASTA_SAIDA,
                arquivo
            )

            return FileResponse(
                caminho,
                filename=arquivo,
                media_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )

    raise HTTPException(
        status_code=404,
        detail="Arquivo não encontrado."
    )
