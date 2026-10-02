# -*- coding: utf-8 -*-
"""
PROCESSAR XML DE NF-e

ETAPA 1 - TESTE / SIMULAÇÃO
----------------------------
Lê os XMLs colocados em A_PROCESSAR, identifica:
- CNPJ do emitente
- CNPJ do destinatário
- número da NF
- data de emissão

Depois consulta banco.xlsx para descobrir o cliente.

REGRA:
- CNPJ do cliente = emitente  -> SAIDAS
- CNPJ do cliente = destinatário -> ENTRADAS

IMPORTANTE:
Por enquanto o script NÃO MOVE os XMLs.
Ele apenas analisa e gera um relatório.

Quando a conferência estiver correta, basta mudar:
MODO_EXECUCAO = True
"""

from pathlib import Path
import openpyxl
import xml.etree.ElementTree as ET
from datetime import datetime
import re

# ============================================================
# CONFIGURAÇÃO
# ============================================================

MODO_EXECUCAO = False   # False = apenas relatório | True = mover os XMLs

ANO_PADRAO = None       # None = usa o ano da NF-e

BASE = Path(__file__).resolve().parent

ARQUIVO_EXCEL = BASE / "banco.xlsx"
PASTA_PROCESSAR = BASE / "A_PROCESSAR"
PASTA_ERROS = BASE / "ERROS_XML"
RELATORIO = BASE / "relatorio_xml.txt"

# ============================================================
# FUNÇÕES
# ============================================================

def normalizar_cnpj(valor):
    """Retorna somente os números do CNPJ."""
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def texto_local(elemento, nome):
    """Procura uma tag ignorando namespace."""
    if elemento is None:
        return None

    for filho in elemento.iter():
        tag = filho.tag
        if isinstance(tag, str):
            tag_sem_ns = tag.split("}")[-1]
            if tag_sem_ns == nome:
                return (filho.text or "").strip()

    return None


def localizar_inf_nfe(root):
    """Localiza infNFe dentro de NFe/nfeProc."""
    for elemento in root.iter():
        tag = elemento.tag.split("}")[-1] if isinstance(elemento.tag, str) else ""
        if tag == "infNFe":
            return elemento
    return None


def ler_nfe(caminho):
    """Extrai os dados principais do XML."""
    try:
        root = ET.parse(caminho).getroot()
    except Exception as e:
        return None, f"XML inválido ou não foi possível ler: {e}"

    inf = localizar_inf_nfe(root)

    if inf is None:
        return None, "Não foi encontrado o elemento infNFe. O arquivo pode não ser uma NF-e."

    # Emitente e destinatário são procurados dentro de infNFe
    emit = None
    dest = None

    for filho in inf:
        tag = filho.tag.split("}")[-1] if isinstance(filho.tag, str) else ""
        if tag == "emit":
            emit = filho
        elif tag == "dest":
            dest = filho

    cnpj_emitente = texto_local(emit, "CNPJ")
    cnpj_destinatario = texto_local(dest, "CNPJ")

    numero = texto_local(inf, "nNF")
    chave = inf.attrib.get("Id", "")
    chave = re.sub(r"^NFe", "", chave)

    data = texto_local(inf, "dhEmi") or texto_local(inf, "dEmi")

    if not cnpj_emitente:
        return None, "XML sem CNPJ do emitente."

    if not numero:
        return None, "XML sem número da NF (nNF)."

    if not data:
        return None, "XML sem data de emissão (dhEmi/dEmi)."

    # Converte a data para datetime
    try:
        if "T" in data:
            data_nf = datetime.fromisoformat(data.replace("Z", "+00:00"))
        else:
            data_nf = datetime.strptime(data[:10], "%Y-%m-%d")
    except Exception:
        return None, f"Data de emissão inválida: {data}"

    return {
        "emitente": normalizar_cnpj(cnpj_emitente),
        "destinatario": normalizar_cnpj(cnpj_destinatario),
        "numero": str(numero).strip(),
        "data": data_nf,
        "chave": chave,
    }, None


def carregar_clientes():
    """Lê COD, RAZÃO SOCIAL, CNPJ e CAMINHO ABSOLUTO."""
    if not ARQUIVO_EXCEL.exists():
        raise FileNotFoundError(
            f"Não encontrei {ARQUIVO_EXCEL}"
        )

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL, data_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(1, col).value
        if valor is not None:
            headers[str(valor).strip().upper()] = col

    obrigatorias = ["COD", "RAZÃO SOCIAL", "CNPJ", "CAMINHO ABSOLUTO"]
    faltantes = [h for h in obrigatorias if h not in headers]

    if faltantes:
        raise ValueError(
            "Colunas faltantes na planilha: " + ", ".join(faltantes)
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


def pasta_mes(data_nf):
    ano = ANO_PADRAO or data_nf.year
    mes = data_nf.month
    return str(ano), f"{mes:02d}-{str(ano)[-2:]}"


# ============================================================
# INÍCIO
# ============================================================

print()
print("=" * 70)
print("  PROCESSAMENTO DE XML DE NF-e")
print("=" * 70)
print()

PASTA_PROCESSAR.mkdir(parents=True, exist_ok=True)

try:
    clientes = carregar_clientes()
except Exception as e:
    print("ERRO ao carregar a planilha:")
    print(e)
    input("\nPressione ENTER para fechar...")
    raise SystemExit(1)

arquivos_xml = sorted(PASTA_PROCESSAR.glob("*.xml"))

print(f"Planilha: {ARQUIVO_EXCEL}")
print(f"Pasta de entrada: {PASTA_PROCESSAR}")
print(f"XML encontrados: {len(arquivos_xml)}")
print()

if not arquivos_xml:
    print("Nenhum XML encontrado em A_PROCESSAR.")
    print("Coloque os XMLs nessa pasta e execute novamente.")
    input("\nPressione ENTER para fechar...")
    raise SystemExit(0)

resultados = []
quantidade_ok = 0
quantidade_erro = 0

for xml in arquivos_xml:
    dados, erro = ler_nfe(xml)

    if erro:
        quantidade_erro += 1
        resultados.append(
            f"ERRO | {xml.name} | {erro}"
        )
        continue

    emitente = dados["emitente"]
    destinatario = dados["destinatario"]

    candidatos_emitente = clientes.get(emitente, [])
    candidatos_destinatario = clientes.get(destinatario, [])

    # Verifica em qual lado está o cliente
    encontrou_saida = len(candidatos_emitente) > 0
    encontrou_entrada = len(candidatos_destinatario) > 0

    # Se o mesmo cliente aparece nos dois lados, não decide sozinho.
    if encontrou_saida and encontrou_entrada:
        quantidade_erro += 1
        resultados.append(
            f"REVISAR | {xml.name} | "
            f"Cliente aparece como emitente E destinatário | "
            f"NF {dados['numero']} | "
            f"Emitente {emitente} | Destinatário {destinatario}"
        )
        continue

    if not encontrou_saida and not encontrou_entrada:
        quantidade_erro += 1
        resultados.append(
            f"REVISAR | {xml.name} | "
            f"Cliente não encontrado na planilha | "
            f"NF {dados['numero']} | "
            f"Emitente {emitente} | Destinatário {destinatario}"
        )
        continue

    if encontrou_saida:
        candidatos = candidatos_emitente
        tipo = "SAIDAS"
    else:
        candidatos = candidatos_destinatario
        tipo = "ENTRADAS"

    # CNPJ duplicado na planilha: não escolhe automaticamente.
    if len(candidatos) != 1:
        quantidade_erro += 1
        nomes = " / ".join(
            str(c["razao"]) for c in candidatos
        )
        resultados.append(
            f"REVISAR | {xml.name} | "
            f"CNPJ cadastrado para mais de um cliente | "
            f"NF {dados['numero']} | "
            f"CNPJ {candidatos[0]['cnpj']} | "
            f"Clientes: {nomes}"
        )
        continue

    cliente = candidatos[0]
    ano, mes = pasta_mes(dados["data"])

    pasta_destino = (
        Path(cliente["pasta"])
        / "NF"
        / tipo
        / ano
        / mes
    )

    novo_nome = f"NF {dados['numero']}.xml"
    destino = pasta_destino / novo_nome

    if not cliente["pasta"]:
        quantidade_erro += 1
        resultados.append(
            f"REVISAR | {xml.name} | Caminho absoluto vazio para "
            f"{cliente['razao']}"
        )
        continue

    if not Path(cliente["pasta"]).exists():
        quantidade_erro += 1
        resultados.append(
            f"REVISAR | {xml.name} | Pasta do cliente não existe: "
            f"{cliente['pasta']}"
        )
        continue

    if destino.exists():
        quantidade_erro += 1
        resultados.append(
            f"CONFLITO | {xml.name} | Destino já existe: {destino}"
        )
        continue

    # Por enquanto, somente simulação.
    if MODO_EXECUCAO:
        try:
            pasta_destino.mkdir(parents=True, exist_ok=True)
            xml.rename(destino)
            acao = "MOVIDO"
        except Exception as e:
            quantidade_erro += 1
            resultados.append(
                f"ERRO | {xml.name} | Falha ao mover: {e}"
            )
            continue
    else:
        acao = "SIMULAÇÃO"

    quantidade_ok += 1

    resultados.append(
        f"{acao} | {xml.name} | "
        f"{tipo} | Cliente: {cliente['cod']} - {cliente['razao']} | "
        f"NF {dados['numero']} | "
        f"Data {dados['data']:%d/%m/%Y} | "
        f"Destino: {destino}"
    )

# ============================================================
# RELATÓRIO
# ============================================================

with open(RELATORIO, "w", encoding="utf-8") as f:
    f.write("RELATÓRIO DE PROCESSAMENTO DE XML - NF-e\n")
    f.write("=" * 100 + "\n")
    f.write(f"Data/hora: {datetime.now():%d/%m/%Y %H:%M:%S}\n")
    f.write(f"Modo de execução: {'MOVER ARQUIVOS' if MODO_EXECUCAO else 'SIMULAÇÃO'}\n")
    f.write(f"XML encontrados: {len(arquivos_xml)}\n")
    f.write(f"Processados com sucesso: {quantidade_ok}\n")
    f.write(f"Revisões/erros: {quantidade_erro}\n\n")

    for linha in resultados:
        f.write(linha + "\n")

print("=" * 70)
print("RESULTADO")
print("=" * 70)
print(f"XML encontrados:        {len(arquivos_xml)}")
print(f"Processados:            {quantidade_ok}")
print(f"Revisões/erros:         {quantidade_erro}")
print()
print(f"Relatório: {RELATORIO}")
print()

if not MODO_EXECUCAO:
    print("MODO SIMULAÇÃO: nenhum XML foi movido.")
    print("Primeiro vamos conferir o relatório.")
else:
    print("MODO EXECUÇÃO: os XMLs aprovados foram movidos.")

input("\nPressione ENTER para fechar...")
