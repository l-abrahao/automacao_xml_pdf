# -*- coding: utf-8 -*-
"""
RESUMO XML + PDF DE NF-e
===========================

MODO ATUAL: SIMULAÇÃO
- Não move arquivos.
- Analisa XMLs e PDFs encontrados em A_PROCESSAR.
- Gera relatorio_xml_pdf.txt.

XML:
- tenta identificar emitente, destinatário, número e data;
- cruza o CNPJ com banco.xlsx.

PDF:
- tenta extrair texto do DANFE;
- procura número da NF, chave de acesso, CNPJ e data;
- usa o conteúdo para auxiliar a identificação.

IMPORTANTE:
O PDF pode exigir tratamento adicional dependendo do modelo/qualidade.
Por isso, nenhum arquivo é movido nesta etapa ainda.

Dependências:
    python -m pip install openpyxl pypdf
"""

from pathlib import Path
import re
import openpyxl
import xml.etree.ElementTree as ET
from datetime import datetime

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

# ============================================================
# CONFIGURAÇÃO
# ============================================================

MODO_EXECUCAO = False  # False = simulação | True = mover

BASE = Path(__file__).resolve().parent
ARQUIVO_EXCEL = BASE / "banco.xlsx"
PASTA_PROCESSAR = BASE / "A_PROCESSAR"
RELATORIO = BASE / "relatorio_xml_pdf.txt"

# ============================================================
# FUNÇÕES GERAIS
# ============================================================

def normalizar_cnpj(valor):
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def normalizar_numero_nf(valor):
    if not valor:
        return ""
    valor = str(valor).strip()
    # Remove zeros à esquerda apenas para comparação;
    # o número original do XML será usado no nome quando possível.
    try:
        return str(int(valor))
    except Exception:
        return re.sub(r"\D", "", valor)


def localizar_tag(elemento, nome):
    if elemento is None:
        return None

    for filho in elemento.iter():
        tag = filho.tag.split("}")[-1] if isinstance(filho.tag, str) else ""
        if tag == nome:
            return (filho.text or "").strip()

    return None


def encontrar_inf_nfe(root):
    for elemento in root.iter():
        tag = elemento.tag.split("}")[-1] if isinstance(elemento.tag, str) else ""
        if tag == "infNFe":
            return elemento
    return None


# ============================================================
# XML
# ============================================================

def ler_xml(caminho):
    try:
        root = ET.parse(caminho).getroot()
    except Exception as e:
        return None, f"XML inválido: {e}"

    inf = encontrar_inf_nfe(root)

    if inf is None:
        return None, "Não foi encontrado infNFe."

    emit = None
    dest = None

    for filho in inf:
        tag = filho.tag.split("}")[-1] if isinstance(filho.tag, str) else ""
        if tag == "emit":
            emit = filho
        elif tag == "dest":
            dest = filho

    cnpj_emitente = localizar_tag(emit, "CNPJ")
    cnpj_destinatario = localizar_tag(dest, "CNPJ")
    numero = localizar_tag(inf, "nNF")
    data = localizar_tag(inf, "dhEmi") or localizar_tag(inf, "dEmi")

    if not cnpj_emitente:
        return None, "XML sem CNPJ do emitente."

    if not numero:
        return None, "XML sem número da NF."

    if not data:
        return None, "XML sem data de emissão."

    try:
        if "T" in data:
            data_nf = datetime.fromisoformat(data.replace("Z", "+00:00"))
        else:
            data_nf = datetime.strptime(data[:10], "%Y-%m-%d")
    except Exception:
        return None, f"Data inválida: {data}"

    chave = inf.attrib.get("Id", "")
    chave = re.sub(r"^NFe", "", chave)

    return {
        "tipo": "XML",
        "arquivo": caminho.name,
        "emitente": normalizar_cnpj(cnpj_emitente),
        "destinatario": normalizar_cnpj(cnpj_destinatario),
        "numero": str(numero).strip(),
        "numero_cmp": normalizar_numero_nf(numero),
        "data": data_nf,
        "chave": re.sub(r"\D", "", chave),
    }, None


# ============================================================
# PDF
# ============================================================

def extrair_texto_pdf(caminho):
    if PdfReader is None:
        return None, "Biblioteca pypdf não instalada."

    try:
        reader = PdfReader(str(caminho))
        paginas = []

        for pagina in reader.pages:
            try:
                paginas.append(pagina.extract_text() or "")
            except Exception:
                paginas.append("")

        texto = "\n".join(paginas)

        if not texto.strip():
            return None, "PDF sem texto extraível. Pode ser PDF escaneado/imagem."

        return texto, None

    except Exception as e:
        return None, f"Erro ao ler PDF: {e}"


def limpar_cnpj_texto(valor):
    return normalizar_cnpj(valor)


def extrair_dados_pdf(caminho):
    texto, erro = extrair_texto_pdf(caminho)

    if erro:
        return None, erro

    texto_compacto = re.sub(r"[ \t]+", " ", texto)
    texto_maiusculo = texto_compacto.upper()

    # Chave de acesso: 44 dígitos, geralmente impressa no DANFE.
    chaves = re.findall(r"\b\d{44}\b", re.sub(r"\D", " ", texto))
    chave = chaves[0] if chaves else ""

    # Número da NF.
    numero = ""
    padroes_nf = [
        r"N[ÚU]MERO\s*(?:DA\s*)?NF(?:-E)?\s*[:.]?\s*([0-9]{1,12})",
        r"N[ÚU]MERO\s*[:.]?\s*([0-9]{1,12})",
        r"NF(?:-E)?\s*N[º°]?\s*[:.]?\s*([0-9]{1,12})",
    ]

    for padrao in padroes_nf:
        m = re.search(padrao, texto_maiusculo)
        if m:
            numero = m.group(1)
            break

    # CNPJs encontrados no PDF.
    cnpjs = []
    encontrados = re.findall(
        r"\d{2}[.\s]?\d{3}[.\s]?\d{3}[\/\s]?\d{3,4}[-\s]?\d{2}",
        texto_compacto
    )

    for item in encontrados:
        cnpj = limpar_cnpj_texto(item)
        if len(cnpj) == 14 and cnpj not in cnpjs:
            cnpjs.append(cnpj)

    # Datas no padrão brasileiro.
    datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", texto_compacto)

    data_nf = None
    for d in datas:
        try:
            data_nf = datetime.strptime(d, "%d/%m/%Y")
            break
        except Exception:
            pass

    return {
        "tipo": "PDF",
        "arquivo": caminho.name,
        "chave": chave,
        "numero": numero,
        "numero_cmp": normalizar_numero_nf(numero),
        "cnpjs": cnpjs,
        "data": data_nf,
        "texto": texto,
    }, None


# ============================================================
# CLIENTES
# ============================================================

def carregar_clientes():
    if not ARQUIVO_EXCEL.exists():
        raise FileNotFoundError(f"Não encontrei {ARQUIVO_EXCEL}")

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL, data_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(1, col).value
        if valor is not None:
            headers[str(valor).strip().upper()] = col

    obrigatorias = ["COD", "RAZÃO SOCIAL", "CNPJ", "CAMINHO ABSOLUTO"]
    faltantes = [x for x in obrigatorias if x not in headers]

    if faltantes:
        raise ValueError(
            "Colunas faltantes: " + ", ".join(faltantes)
        )

    clientes = {}

    for linha in range(2, ws.max_row + 1):
        cnpj = normalizar_cnpj(ws.cell(linha, headers["CNPJ"]).value)

        if not cnpj:
            continue

        cliente = {
            "cod": ws.cell(linha, headers["COD"]).value,
            "razao": ws.cell(linha, headers["RAZÃO SOCIAL"]).value,
            "cnpj": cnpj,
            "pasta": str(
                ws.cell(linha, headers["CAMINHO ABSOLUTO"]).value or ""
            ).strip(),
        }

        clientes.setdefault(cnpj, []).append(cliente)

    return clientes


# ============================================================
# LOCALIZAÇÃO DO CLIENTE
# ============================================================

def cliente_por_cnpj(cnpj, clientes):
    candidatos = clientes.get(cnpj, [])

    if len(candidatos) == 1:
        return candidatos[0], None

    if len(candidatos) > 1:
        nomes = " / ".join(str(x["razao"]) for x in candidatos)
        return None, f"CNPJ duplicado na planilha: {nomes}"

    return None, "CNPJ não encontrado na planilha."


def determinar_destino_por_xml(dados, clientes):
    cliente_emitente, erro_emitente = cliente_por_cnpj(
        dados["emitente"], clientes
    )

    cliente_dest, erro_dest = cliente_por_cnpj(
        dados["destinatario"], clientes
    )

    if cliente_emitente and cliente_dest:
        return None, None, "Cliente aparece como emitente e destinatário."

    if cliente_emitente:
        return cliente_emitente, "SAIDAS", None

    if cliente_dest:
        return cliente_dest, "ENTRADAS", None

    return None, None, (
        f"Cliente não encontrado. "
        f"Emitente={dados['emitente']} | "
        f"Destinatário={dados['destinatario']}"
    )


def determinar_destino_por_pdf(dados, clientes):
    candidatos = []

    for cnpj in dados["cnpjs"]:
        cliente, erro = cliente_por_cnpj(cnpj, clientes)
        if cliente:
            candidatos.append((cliente, cnpj))

    # Remove duplicidades
    unicos = {}
    for cliente, cnpj in candidatos:
        chave = (cliente["cnpj"], cliente["pasta"])
        unicos[chave] = (cliente, cnpj)

    candidatos = list(unicos.values())

    if len(candidatos) == 1:
        cliente, cnpj = candidatos[0]

        # No PDF sozinho, ainda não temos com segurança
        # se o cliente é emitente ou destinatário.
        return cliente, None, (
            "Cliente encontrado no PDF, mas entrada/saída precisa "
            "ser confirmada pelo conteúdo/posição dos CNPJs ou XML."
        )

    if len(candidatos) > 1:
        nomes = " / ".join(str(x[0]["razao"]) for x in candidatos)
        return None, None, f"Mais de um cliente encontrado no PDF: {nomes}"

    return None, None, "Nenhum CNPJ do PDF correspondeu à planilha."


# ============================================================
# PROCESSAMENTO
# ============================================================

print()
print("=" * 80)
print("  ANÁLISE DE XML + PDF DE NF-e")
print("=" * 80)
print()

PASTA_PROCESSAR.mkdir(parents=True, exist_ok=True)

try:
    clientes = carregar_clientes()
except Exception as e:
    print("ERRO:", e)
    input("\nPressione ENTER para fechar...")
    raise SystemExit(1)

arquivos = sorted(
    [
        p for p in PASTA_PROCESSAR.iterdir()
        if p.is_file() and p.suffix.lower() in (".xml", ".pdf")
    ],
    key=lambda x: x.name.lower()
)

xmls = [p for p in arquivos if p.suffix.lower() == ".xml"]
pdfs = [p for p in arquivos if p.suffix.lower() == ".pdf"]

print(f"Arquivos encontrados: {len(arquivos)}")
print(f"XML: {len(xmls)}")
print(f"PDF: {len(pdfs)}")
print()

resultados = []

# ---------- XML ----------
dados_xml = {}

for arquivo in xmls:
    dados, erro = ler_xml(arquivo)

    if erro:
        resultados.append(
            f"REVISAR | XML | {arquivo.name} | {erro}"
        )
        continue

    cliente, tipo, erro_destino = determinar_destino_por_xml(
        dados, clientes
    )

    if erro_destino:
        resultados.append(
            f"REVISAR | XML | {arquivo.name} | "
            f"NF={dados['numero']} | {erro_destino}"
        )
        continue

    dados["cliente"] = cliente
    dados["tipo_movimento"] = tipo
    dados["destino"] = (
        Path(cliente["pasta"])
        / "NF"
        / tipo
        / str(dados["data"].year)
        / f"{dados['data'].month:02d}-{str(dados['data'].year)[-2:]}"
        / f"NF {dados['numero']}.xml"
    )

    dados_xml[dados["chave"]] = dados

    resultados.append(
        f"OK | XML | {arquivo.name} | "
        f"{tipo} | Cliente={cliente['cod']} - {cliente['razao']} | "
        f"NF={dados['numero']} | Data={dados['data']:%d/%m/%Y} | "
        f"Destino={dados['destino']}"
    )

# ---------- PDF ----------
for arquivo in pdfs:
    dados, erro = extrair_dados_pdf(arquivo)

    if erro:
        resultados.append(
            f"REVISAR | PDF | {arquivo.name} | {erro}"
        )
        continue

    # Primeiro tenta casar com XML pela chave.
    xml_correspondente = None

    if dados["chave"]:
        for xml_dados in dados_xml.values():
            if xml_dados.get("chave") == dados["chave"]:
                xml_correspondente = xml_dados
                break

    # Se não houver chave, tenta pelo número da NF.
    if xml_correspondente is None and dados["numero_cmp"]:
        candidatos = [
            x for x in dados_xml.values()
            if x.get("numero_cmp") == dados["numero_cmp"]
        ]

        if len(candidatos) == 1:
            xml_correspondente = candidatos[0]

    if xml_correspondente:
        destino_pdf = (
            xml_correspondente["destino"].parent
            / f"NF {xml_correspondente['numero']}.pdf"
        )

        resultados.append(
            f"PAR XML/PDF | PDF | {arquivo.name} | "
            f"Cliente={xml_correspondente['cliente']['cod']} - "
            f"{xml_correspondente['cliente']['razao']} | "
            f"NF={xml_correspondente['numero']} | "
            f"{xml_correspondente['tipo_movimento']} | "
            f"Destino={destino_pdf}"
        )
        continue

    # PDF sem XML correspondente.
    cliente, tipo, aviso = determinar_destino_por_pdf(dados, clientes)

    if cliente is None:
        resultados.append(
            f"REVISAR | PDF | {arquivo.name} | "
            f"NF={dados['numero'] or '?'} | {aviso}"
        )
        continue

    # Sem XML, não vamos adivinhar entrada/saída.
    resultados.append(
        f"REVISAR | PDF | {arquivo.name} | "
        f"NF={dados['numero'] or '?'} | "
        f"Cliente={cliente['cod']} - {cliente['razao']} | "
        f"{aviso}"
    )

# ============================================================
# RELATÓRIO
# ============================================================

with open(RELATORIO, "w", encoding="utf-8") as f:
    f.write("RELATÓRIO - ANÁLISE DE XML + PDF DE NF-e\n")
    f.write("=" * 110 + "\n")
    f.write(f"Data/hora: {datetime.now():%d/%m/%Y %H:%M:%S}\n")
    f.write(f"Modo: {'EXECUÇÃO' if MODO_EXECUCAO else 'SIMULAÇÃO'}\n")
    f.write(f"Total de arquivos XML/PDF: {len(arquivos)}\n")
    f.write(f"XML: {len(xmls)}\n")
    f.write(f"PDF: {len(pdfs)}\n\n")

    for linha in resultados:
        f.write(linha + "\n")

print("=" * 80)
print("ANÁLISE CONCLUÍDA")
print("=" * 80)
print(f"Total: {len(arquivos)} arquivos")
print(f"XML:   {len(xmls)}")
print(f"PDF:   {len(pdfs)}")
print()
print(f"Relatório criado em:")
print(RELATORIO)
print()
print("IMPORTANTE: nenhum arquivo foi movido.")
print("Vamos revisar o relatório antes de ativar qualquer movimentação.")

input("\nPressione ENTER para fechar...")
