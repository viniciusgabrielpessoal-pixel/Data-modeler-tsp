import json
import os
from io import BytesIO
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file, session

from processor import (
    DE_PARA_PADRAO,
    carregar_de_para,
    gerar_excel_bytes,
    processar_varios,
)

# Quando rodando como .exe, o launcher aponta UNIFICADOR_TEMPLATES
# para a pasta templates dentro do bundle PyInstaller (_MEIPASS)
_templates_dir = os.environ.get("UNIFICADOR_TEMPLATES", "templates")
app = Flask(__name__, template_folder=_templates_dir)
app.secret_key = "unificador-latam-secret"

DE_PARA_GLOBAL = carregar_de_para()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/de-para", methods=["GET"])
def get_de_para():
    return jsonify(DE_PARA_GLOBAL)


@app.route("/api/de-para", methods=["POST"])
def set_de_para():
    global DE_PARA_GLOBAL
    dados = request.get_json(force=True)
    if not isinstance(dados, dict):
        return jsonify({"erro": "Payload inválido"}), 400
    DE_PARA_GLOBAL = dados
    return jsonify({"ok": True, "total": len(DE_PARA_GLOBAL)})


@app.route("/api/de-para/restaurar", methods=["POST"])
def restaurar_de_para():
    global DE_PARA_GLOBAL
    DE_PARA_GLOBAL = carregar_de_para()
    return jsonify({"ok": True, "de_para": DE_PARA_GLOBAL})


@app.route("/api/processar", methods=["POST"])
def processar():
    arquivos = request.files.getlist("arquivos")
    nome_saida = request.form.get("nome_saida", "Resultado_Unificado_Final.xlsx")
    de_para_raw = request.form.get("de_para", None)

    if not nome_saida.lower().endswith(".xlsx"):
        nome_saida += ".xlsx"

    if not arquivos:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400

    de_para = DE_PARA_GLOBAL
    if de_para_raw:
        try:
            de_para = json.loads(de_para_raw)
        except Exception:
            pass

    arquivos_dict = {f.filename: f.read() for f in arquivos}
    logs = []
    df = processar_varios(arquivos_dict, de_para, logs)

    if df is None or df.empty:
        return jsonify({"erro": "Nenhum dado processado", "logs": logs}), 422

    return jsonify({
        "ok": True,
        "registros": len(df),
        "arquivos": len(arquivos_dict),
        "regras": len(de_para),
        "logs": logs,
        "preview": df.head(200).fillna("").to_dict(orient="records"),
        "nome_saida": nome_saida,
    })


@app.route("/api/download", methods=["POST"])
def download():
    arquivos = request.files.getlist("arquivos")
    nome_saida = request.form.get("nome_saida", "Resultado_Unificado_Final.xlsx")
    de_para_raw = request.form.get("de_para", None)

    if not nome_saida.lower().endswith(".xlsx"):
        nome_saida += ".xlsx"

    if not arquivos:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400

    de_para = DE_PARA_GLOBAL
    if de_para_raw:
        try:
            de_para = json.loads(de_para_raw)
        except Exception:
            pass

    arquivos_dict = {f.filename:    f.read() for f in arquivos}
    logs = []
    df = processar_varios(arquivos_dict, de_para, logs)

    if df is None or df.empty:
        return jsonify({"erro": "Nenhum dado processado", "logs": logs}), 422

    excel_bytes = gerar_excel_bytes(df)

    return send_file(
        BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=nome_saida,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)