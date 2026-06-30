import json
import re
import unicodedata
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path

import pandas as pd


class HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_thead = False
        self.in_tr = False
        self.current_row = []
        self.current_cell = []
        self.headers = []
        self.data = []
        self.capture_data = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
            self.headers = []
            self.data = []
        elif tag == "thead" and self.in_table:
            self.in_thead = True
        elif tag == "tr" and self.in_table:
            self.in_tr = True
            self.current_row = []
        elif tag in ("th", "td") and self.in_tr:
            self.current_cell = []
            self.capture_data = True

    def handle_endtag(self, tag):
        if tag == "table":
            self.in_table = False
        elif tag == "thead":
            self.in_thead = False
        elif tag == "tr":
            self.in_tr = False
            if self.in_thead and self.current_row:
                self.headers = self.current_row.copy()
            elif self.in_table and self.current_row:
                self.data.append(self.current_row.copy())
            self.current_row = []
        elif tag in ("th", "td"):
            self.capture_data = False
            self.current_row.append("".join(self.current_cell).strip())

    def handle_data(self, data):
        if self.capture_data:
            self.current_cell.append(data.strip())

    def get_dataframe(self):
        if not self.data:
            return None
        if self.headers:
            max_cols = max(len(self.headers), max((len(r) for r in self.data), default=0))
            headers = self.headers[:]
            if len(headers) < max_cols:
                headers.extend([f"Col_{i}" for i in range(len(headers), max_cols)])
            linhas = []
            for row in self.data:
                nova = row[:]
                if len(nova) < max_cols:
                    nova.extend([""] * (max_cols - len(nova)))
                linhas.append(nova)
            return pd.DataFrame(linhas, columns=headers[:max_cols])
        return pd.DataFrame(self.data)


COLUNAS_PADRAO = [
    "Codigo", "Nota Fiscal", "Pedido", "Cliente", "Cep Remetente",
    "Destino", "Cidade", "Bairro", "UF", "Destinatario",
    "CNPJ/CPF Destinatario", "Cep Destinatario", "TP", "Status",
    "STATUS DHL", "Dt Emissao", "Dt Evento", "Previsao",
    "DescricaoRecebedorDoc", "Recebedor", "Transportador",
]

COLUNAS_TEXTO = [
    "Codigo", "Nota Fiscal", "Cep Remetente", "Cep Destinatario",
    "Pedido", "CNPJ/CPF Destinatario", "TP",
]

DE_PARA_PADRAO = {
    "ENTREGUE": {"STATUS DHL": "ENTREGUE", "SINAL": "FINALIZADO", "TITULO": "ENTREGUE"},
    "DEVOLVIDO": {"STATUS DHL": "DEVOLVIDO", "SINAL": "FINALIZADO", "TITULO": "DEVOLVIDO"},
    "EXTRAVIO": {"STATUS DHL": "EXTRAVIO", "SINAL": "FINALIZADO", "TITULO": "EXTRAVIO"},
}


def carregar_de_para(json_path=None):
    if json_path and Path(json_path).exists():
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    caminho_json = Path(__file__).with_name("novo_de_para_status.json")
    if caminho_json.exists():
        with open(caminho_json, "r", encoding="utf-8") as f:
            return json.load(f)
    return DE_PARA_PADRAO.copy()


def corrigir_mojibake(texto):
    texto = str(texto)
    corrigido = texto
    for _ in range(2):
        if any(char in corrigido for char in ("Ã", "Â", "â€")):
            try:
                novo = corrigido.encode("latin-1").decode("utf-8")
                if novo != corrigido:
                    corrigido = novo
                    continue
            except Exception:
                pass
        break
    return corrigido


def normalizar_texto_match(texto):
    texto = corrigir_mojibake(texto).upper().strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def formatar_data_valor(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime("%d/%m/%Y")
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nat", "nan", "none"}:
        return ""
    convertido = pd.to_datetime(texto, dayfirst=True, errors="coerce")
    if pd.notna(convertido):
        return convertido.strftime("%d/%m/%Y")
    return texto


def formatar_texto_identificador(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none", "nat"}:
        return ""
    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".")[0]
    return texto


def forcar_colunas_texto(df, colunas=None):
    colunas = colunas or COLUNAS_TEXTO
    for coluna in colunas:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(formatar_texto_identificador)
    return df


def formatar_colunas_data(df, colunas):
    for coluna in colunas:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(formatar_data_valor)
    return df


def normalizar_nome_coluna(nome):
    nome = str(nome).strip()
    return unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("ASCII")


def identificar_tipo_arquivo(nome):
    nome = Path(nome).name
    if nome.startswith("Exp_") and (nome.endswith(".xls") or nome.endswith(".xlsx")):
        return "EXP", "CORREIOS"
    if (nome.startswith("Performance") and nome.endswith(".xlsx")) or nome.endswith(".zip"):
        return "PERFORMANCE", "JADLOG"
    if nome.startswith("relatorio_") and nome.endswith(".csv"):
        return "RELATORIO", "TOTAL EXPRESS"
    return "DESCONHECIDO", "DESCONHECIDO"


def aplicar_mapeamento(status_raw, de_para):
    texto = normalizar_texto_match(status_raw)
    for chave, valores in de_para.items():
        if normalizar_texto_match(chave) in texto:
            return {
                "STATUS DHL": valores["STATUS DHL"],
                "SINAL": valores["SINAL"],
                "TITULO": valores["TITULO"],
            }
    return {"STATUS DHL": "TRANSITO", "SINAL": "OK", "TITULO": "TRANSITO"}


def ler_html_nativo_bytes(conteudo, logs):
    try:
        html_content = conteudo.decode("utf-8", errors="ignore")
        parser = HTMLTableParser()
        parser.feed(html_content)
        df = parser.get_dataframe()
        if df is not None and not df.empty:
            logs.append("  Lido com parser HTML nativo.")
            return df
    except Exception as e:
        logs.append(f"  Parser HTML nativo falhou: {str(e)[:80]}")
    return None


def ler_arquivo_com_pandas(nome_arquivo, conteudo, logs):
    prefixo = conteudo[:2048].lower()

    try:
        if b"<html" in prefixo or b"<table" in prefixo:
            logs.append("  Detectado formato HTML...")
            df = ler_html_nativo_bytes(conteudo, logs)
            if df is not None:
                return df
    except Exception as e:
        logs.append(f"  Erro ao detectar HTML: {str(e)[:80]}")

    try:
        if nome_arquivo.lower().endswith(".xls"):
            logs.append("  Tentando ler com xlrd...")
            df = pd.read_excel(BytesIO(conteudo), engine="xlrd")
            if not df.empty:
                logs.append("  Lido com sucesso via xlrd.")
                return df
    except Exception as e:
        logs.append(f"  Falha no xlrd: {str(e)[:80]}")

    try:
        logs.append("  Tentando ler com openpyxl...")
        df = pd.read_excel(BytesIO(conteudo), engine="openpyxl")
        if not df.empty:
            logs.append("  Lido com sucesso via openpyxl.")
            return df
    except Exception as e:
        logs.append(f"  Falha no openpyxl: {str(e)[:80]}")

    try:
        logs.append("  Tentando ler com engine automático...")
        df = pd.read_excel(BytesIO(conteudo))
        if not df.empty:
            logs.append("  Lido com sucesso com engine automático.")
            return df
    except Exception as e:
        logs.append(f"  Falha no engine automático: {str(e)[:80]}")

    try:
        logs.append("  Tentando ler como CSV UTF-8...")
        df = pd.read_csv(BytesIO(conteudo), encoding="utf-8", on_bad_lines="skip")
        if not df.empty:
            logs.append("  Lido como CSV UTF-8.")
            return df
    except Exception:
        pass

    try:
        logs.append("  Tentando ler como CSV latin-1...")
        df = pd.read_csv(BytesIO(conteudo), encoding="latin-1", on_bad_lines="skip")
        if not df.empty:
            logs.append("  Lido como CSV latin-1.")
            return df
    except Exception:
        pass

    raise Exception("Não foi possível ler o arquivo com nenhum método disponível.")


def processar_arquivo(nome, conteudo_bytes, de_para, logs):
    tipo, transportadora = identificar_tipo_arquivo(nome)

    if tipo == "DESCONHECIDO":
        logs.append(f"Tipo desconhecido ignorado: {nome}")
        return None

    logs.append(f"Processando {tipo}: {nome}")

    if tipo == "RELATORIO":
        try:
            df = pd.read_csv(BytesIO(conteudo_bytes), encoding="utf-8")
        except Exception:
            try:
                df = pd.read_csv(BytesIO(conteudo_bytes), encoding="latin-1")
            except Exception:
                df = pd.read_csv(BytesIO(conteudo_bytes), encoding="iso-8859-1")
    elif tipo == "PERFORMANCE":
        if nome.lower().endswith(".zip"):
            logs.append("Lendo Performance dentro do ZIP...")
            with zipfile.ZipFile(BytesIO(conteudo_bytes)) as z:
                excel = next((n for n in z.namelist() if n.lower().endswith(".xlsx")), None)
                if excel is None:
                    raise Exception("Nenhum arquivo .xlsx encontrado dentro do ZIP.")
                with z.open(excel) as f:
                    df = pd.read_excel(BytesIO(f.read()), header=1)
        else:
            df = ler_arquivo_com_pandas(nome, conteudo_bytes, logs)
    else:
        df = ler_arquivo_com_pandas(nome, conteudo_bytes, logs)

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["_".join(map(str, col)).strip() for col in df.columns.values]

    df.columns = [re.sub(r".*?_+", "", str(col)) for col in df.columns]
    df.columns = df.columns.str.strip()

    if tipo == "EXP":
        if "DataPostagem" in df.columns and "Prazo" in df.columns:
            df["DataPostagem"] = pd.to_datetime(df["DataPostagem"], dayfirst=True, errors="coerce")
            df["Prazo"] = pd.to_numeric(df["Prazo"], errors="coerce").fillna(0).astype(int)
            df["PrevisaoEntrega"] = df.apply(
                lambda x: x["DataPostagem"] + pd.offsets.BusinessDay(x["Prazo"])
                if pd.notnull(x["DataPostagem"]) else None,
                axis=1,
            )

        colunas_map = {
            "Registro": "Codigo", "NumeroNotaFiscal": "Nota Fiscal",
            "Remetente": "Cliente", "Cep Remetente": "Cep Remetente",
            "Destino": "Destino", "Cidade": "Cidade", "Bairro": "Bairro",
            "UF": "UF", "Cep Destinatario": "Cep Destinatario", "Status": "Status",
            "DataPostagem": "Dt Emissao", "DataDoEvento": "Dt Evento",
            "PrevisaoEntrega": "Previsao",
        }

        df_final = pd.DataFrame()
        for col_orig, col_final in colunas_map.items():
            df_final[col_final] = df.get(col_orig, "")

        coluna_status = next((col for col in df.columns if "status" in col.lower()), None)
        if coluna_status:
            mapeado = df[coluna_status].apply(lambda s: aplicar_mapeamento(s, de_para))
            df_map = pd.DataFrame(mapeado.tolist(), index=df.index)
            df_final = pd.concat([df_final, df_map], axis=1)
        else:
            df_final["STATUS DHL"] = ""
            df_final["SINAL"] = ""
            df_final["TITULO"] = ""

        df_final["Transportador"] = transportadora

    elif tipo == "PERFORMANCE":
        df_status_info = df["Status"].apply(lambda s: pd.Series(aplicar_mapeamento(s, de_para)))
        df_final = pd.concat([df, df_status_info], axis=1)
        for col in ["Pedido", "Cep Remetente", "Destino", "Cidade", "Bairro", "UF",
                    "Destinatario", "CNPJ/CPF Destinatario", "Cep Destinatario",
                    "TP", "DescricaoRecebedorDoc", "Recebedor"]:
            if col not in df_final.columns:
                df_final[col] = ""
        df_final["Transportador"] = transportadora

    else:  # RELATORIO / TOTAL EXPRESS
        colunas_esperadas = {
            "Data da encomenda": "Dt Emissao", "Awb": "Codigo",
            "Nota Fiscal": "Nota Fiscal", "Data do último status": "Dt Evento",
            "Descrição do último status": "Status", "Previsão de entrega": "Previsao",
        }
        mapeamento_colunas = {}
        for col_esperada, col_destino in colunas_esperadas.items():
            col_norm = normalizar_nome_coluna(col_esperada)
            for col in df.columns:
                if normalizar_nome_coluna(col) == col_norm:
                    mapeamento_colunas[col] = col_destino
                    break
        if mapeamento_colunas:
            df = df.rename(columns=mapeamento_colunas)

        if "Nota Fiscal" not in df.columns:
            for col in df.columns:
                if any(k in col.lower() for k in ("nota", "fiscal", "nf")):
                    df = df.rename(columns={col: "Nota Fiscal"})
                    break

        if "Nota Fiscal" not in df.columns:
            df["Nota Fiscal"] = ""
        if "Status" not in df.columns:
            df["Status"] = "TRANSITO"

        df["Pedido"] = ""
        df_status_info = df["Status"].apply(lambda s: pd.Series(aplicar_mapeamento(s, de_para)))
        df_final = pd.concat([df, df_status_info], axis=1)

        for col in ["Cliente", "Cep Remetente", "Destino", "Cidade", "Bairro", "UF",
                    "Destinatario", "CNPJ/CPF Destinatario", "Cep Destinatario",
                    "TP", "DescricaoRecebedorDoc", "Recebedor"]:
            if col not in df_final.columns:
                df_final[col] = ""

        df_final["Transportador"] = transportadora

    for col in COLUNAS_PADRAO:
        if col not in df_final.columns:
            df_final[col] = ""

    df_final = forcar_colunas_texto(df_final)
    df_final = formatar_colunas_data(df_final, ["Dt Emissao", "Dt Evento", "Previsao"])
    logs.append(f"{tipo} processado: {len(df_final)} registros.")
    return df_final[COLUNAS_PADRAO]


def processar_varios(arquivos_dict, de_para, logs):
    """
    arquivos_dict: {nome_arquivo: bytes}
    Retorna DataFrame unificado ou None.
    """
    resultados = []
    for nome, conteudo_bytes in arquivos_dict.items():
        try:
            df = processar_arquivo(nome, conteudo_bytes, de_para, logs)
            if df is not None and not df.empty:
                resultados.append(df)
        except Exception as e:
            logs.append(f"Erro ao processar {nome}: {str(e)}")

    if not resultados:
        return None

    df_unificado = pd.concat(resultados, ignore_index=True)[COLUNAS_PADRAO]
    df_unificado = forcar_colunas_texto(df_unificado)
    df_unificado = formatar_colunas_data(df_unificado, ["Dt Emissao", "Dt Evento", "Previsao"])
    return df_unificado


def gerar_excel_bytes(df):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    output.seek(0)
    return output.read()