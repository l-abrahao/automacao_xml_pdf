import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

import openpyxl
from pypdf import PdfReader


# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_PROCESSAR = Path("A_PROCESSAR")
ARQUIVO_BANCO = Path("banco.xlsx")
ARQUIVO_RELATORIO = Path("relatorio_xml_pdf_v7.txt")


# ============================================================
# FUNÇÕES BÁSICAS
# ============================================================

def normalizar_cnpj(valor):
    if valor is None:
        return ""

    valor = str(valor).strip()

    # Evita transformar 126.0 em algo estranho quando usado
    # somente para exibição de código.
    return re.sub(r"\D", "", valor)


def formatar_codigo(valor):
    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def normalizar_texto(valor):
    if valor is None:
        return ""

    texto = str(valor)

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )

    texto = texto.upper()

    return texto


def somente_digitos(valor):
    if valor is None:
        return ""

    return re.sub(r"\D", "", str(valor))


# ============================================================
# XML
# ============================================================

def ler_xml(caminho):
    try:
        return ET.parse(caminho)
    except Exception as e:
        return None


def nome_tag(tag):
    """
    Remove namespace:
    {http://...}nNF -> nNF
    """
    if "}" in tag:
        return tag.split("}", 1)[1]

    return tag


def encontrar_elementos(root, nome):
    resultado = []

    for elem in root.iter():
        if nome_tag(elem.tag).lower() == nome.lower():
            resultado.append(elem)

    return resultado


def primeiro_texto(root, nomes):
    nomes = {x.lower() for x in nomes}

    for elem in root.iter():
        if nome_tag(elem.tag).lower() in nomes:
            if elem.text and elem.text.strip():
                return elem.text.strip()

    return ""


def identificar_tipo_xml(root):
    tags = set()

    for elem in root.iter():
        tags.add(nome_tag(elem.tag).lower())

    # NFS-e
    if (
        "nfse" in tags
        or "infnfse" in tags
        or "nnfse" in tags
        or "vissqn" in tags
    ):
        return "NFS-e"

    # NF-e
    if "infnfe" in tags or "nfe" in tags:
        return "NF-e"

    return "DESCONHECIDO"


def extrair_numero_xml(root, tipo):
    if tipo == "NF-e":
        numero = primeiro_texto(root, ["nNF"])

        if numero:
            return numero

        return ""

    if tipo == "NFS-e":
        numero = primeiro_texto(
            root,
            [
                "nNFSe",
                "numero",
                "Numero",
                "NumeroNota",
                "numeroNfse",
            ]
        )

        return numero

    return ""


def extrair_data_xml(root):
    candidatos = []

    for elem in root.iter():
        tag = nome_tag(elem.tag).lower()

        if tag in {
            "dhemi",
            "demi",
            "dataemissao",
            "dataemissao",
            "dhprocessamento",
        }:
            if elem.text and elem.text.strip():
                candidatos.append(elem.text.strip())

    for valor in candidatos:
        # ISO:
        # 2026-08-25T10:20:30-03:00
        m = re.search(
            r"(\d{4})-(\d{2})-(\d{2})",
            valor
        )

        if m:
            try:
                return datetime.strptime(
                    m.group(0),
                    "%Y-%m-%d"
                )
            except Exception:
                pass

        # BR:
        m = re.search(
            r"(\d{2})/(\d{2})/(\d{4})",
            valor
        )

        if m:
            try:
                return datetime.strptime(
                    m.group(0),
                    "%d/%m/%Y"
                )
            except Exception:
                pass

    return None


def extrair_cnpj_por_contexto(root, contexto):
    """
    Procura CNPJ dentro de emit/dest.
    """

    for elem in root.iter():

        if nome_tag(elem.tag).lower() != contexto.lower():
            continue

        for filho in elem.iter():

            if nome_tag(filho.tag).lower() == "cnpj":

                if filho.text:
                    cnpj = normalizar_cnpj(filho.text)

                    if len(cnpj) == 14:
                        return cnpj

    return ""


def extrair_emitente(root):
    return extrair_cnpj_por_contexto(root, "emit")


def extrair_destinatario(root):
    return extrair_cnpj_por_contexto(root, "dest")


def extrair_chave_xml(root, caminho):
    """
    Tenta obter a chave da NF-e de forma segura.
    """

    # --------------------------------------------------------
    # 1. atributo Id de infNFe
    # --------------------------------------------------------

    for elem in root.iter():

        if nome_tag(elem.tag).lower() == "infnfe":

            for chave_attr in ["Id", "id"]:

                valor = elem.attrib.get(chave_attr)

                if valor:

                    digitos = somente_digitos(valor)

                    if len(digitos) == 44:
                        return digitos

    # --------------------------------------------------------
    # 2. conteúdo XML com sequência de 44 dígitos
    # --------------------------------------------------------

    texto_xml = ""

    try:
        texto_xml = ET.tostring(
            root,
            encoding="unicode"
        )
    except Exception:
        pass

    encontrados = re.findall(
        r"(?<!\d)\d{44}(?!\d)",
        texto_xml
    )

    if encontrados:
        return encontrados[0]

    # --------------------------------------------------------
    # 3. nome do arquivo
    # --------------------------------------------------------

    encontrados = re.findall(
        r"(?<!\d)\d{44}(?!\d)",
        caminho.stem
    )

    if encontrados:
        return encontrados[0]

    return ""


# ============================================================
# BANCO DE CLIENTES
# ============================================================

def carregar_clientes():

    clientes = []

    wb = openpyxl.load_workbook(
        ARQUIVO_BANCO,
        data_only=True
    )

    ws = wb.active

    cabecalhos = {}

    for col in range(1, ws.max_column + 1):

        valor = ws.cell(
            row=1,
            column=col
        ).value

        if valor:
            cabecalhos[
                str(valor).strip().upper()
            ] = col

    col_cod = cabecalhos.get("COD")
    col_razao = cabecalhos.get("RAZÃO SOCIAL")
    col_cnpj = cabecalhos.get("CNPJ")
    col_caminho = cabecalhos.get("CAMINHO ABSOLUTO")

    if not col_cnpj or not col_caminho:
        raise Exception(
            "O banco.xlsx precisa conter as colunas "
            "'CNPJ' e 'CAMINHO ABSOLUTO'."
        )

    for linha in range(2, ws.max_row + 1):

        cnpj = ws.cell(
            linha,
            col_cnpj
        ).value

        caminho = ws.cell(
            linha,
            col_caminho
        ).value

        if not cnpj or not caminho:
            continue

        clientes.append({
            "cod": formatar_codigo(
                ws.cell(linha, col_cod).value
                if col_cod else ""
            ),
            "razao": str(
                ws.cell(linha, col_razao).value
                if col_razao else ""
            ).strip(),
            "cnpj": normalizar_cnpj(cnpj),
            "caminho": Path(str(caminho).strip()),
        })

    return clientes


# ============================================================
# IDENTIFICAÇÃO DO CLIENTE
# ============================================================

def identificar_cliente(
    clientes,
    emitente,
    destinatario
):

    candidatos_entrada = []
    candidatos_saida = []

    for cliente in clientes:

        cnpj_cliente = cliente["cnpj"]

        if destinatario == cnpj_cliente:
            candidatos_entrada.append(cliente)

        if emitente == cnpj_cliente:
            candidatos_saida.append(cliente)

    # --------------------------------------------------------
    # Entrada
    # --------------------------------------------------------

    if len(candidatos_entrada) == 1:
        return (
            candidatos_entrada[0],
            "ENTRADAS",
            ""
        )

    if len(candidatos_entrada) > 1:
        return (
            None,
            "",
            "CNPJ do destinatário encontrado em mais de um cliente"
        )

    # --------------------------------------------------------
    # Saída
    # --------------------------------------------------------

    if len(candidatos_saida) == 1:
        return (
            candidatos_saida[0],
            "SAIDAS",
            ""
        )

    if len(candidatos_saida) > 1:
        return (
            None,
            "",
            "CNPJ do emitente encontrado em mais de um cliente"
        )

    return (
        None,
        "",
        "CNPJ do emitente/destinatário não encontrado no banco.xlsx"
    )


# ============================================================
# DESTINO
# ============================================================

def montar_destino(cliente, direcao, data):

    if not cliente or not data:
        return None

    ano = str(data.year)

    mes = f"{data.month:02d}"
    ano_curto = str(data.year)[-2:]

    pasta_mes = f"{mes}-{ano_curto}"

    return (
        cliente["caminho"]
        / "NF"
        / direcao
        / ano
        / pasta_mes
    )


# ============================================================
# PDF
# ============================================================

def extrair_texto_pdf(caminho):

    textos = []

    try:

        reader = PdfReader(str(caminho))

        for pagina in reader.pages:

            try:
                texto = pagina.extract_text()

                if texto:
                    textos.append(texto)

            except Exception:
                continue

    except Exception:
        return ""

    return "\n".join(textos)


def normalizar_texto_pdf(texto):

    if not texto:
        return ""

    texto = texto.replace("\xa0", " ")

    texto = unicodedata.normalize(
        "NFKC",
        texto
    )

    return texto


def extrair_chaves_pdf(texto):

    if not texto:
        return []

    texto = normalizar_texto_pdf(texto)

    encontrados = set()

    # --------------------------------------------------------
    # Forma 1:
    # 44 dígitos contínuos
    # --------------------------------------------------------

    for chave in re.findall(
        r"(?<!\d)\d{44}(?!\d)",
        texto
    ):
        encontrados.add(chave)

    # --------------------------------------------------------
    # Forma 2:
    # chave separada por espaços
    #
    # Exemplo:
    # 35 2608 5824 8352 ...
    # --------------------------------------------------------

    linhas = texto.splitlines()

    for linha in linhas:

        digitos = somente_digitos(linha)

        if len(digitos) == 44:
            encontrados.add(digitos)

    # --------------------------------------------------------
    # Forma 3:
    # grupos separados por espaços
    #
    # Procuramos grupos que, juntos, formem 44 dígitos.
    # --------------------------------------------------------

    grupos = re.findall(
        r"(?:\d[\s]*){44}",
        texto
    )

    for grupo in grupos:

        digitos = somente_digitos(grupo)

        if len(digitos) == 44:
            encontrados.add(digitos)

    return list(encontrados)


def extrair_numeros_pdf(texto):

    if not texto:
        return []

    texto = normalizar_texto_pdf(texto)

    candidatos = []

    padroes = [
        r"\bN[º°o]?\s*F\s*[-.]?\s*E?\s*[:#]?\s*(\d{1,9})\b",
        r"\bNF[-\s]?[Ee]\s*[:#]?\s*(\d{1,9})\b",
        r"\bN[úu]mero\s+(?:da\s+)?(?:NF|NFe|NFS-e|NFS)\s*[:#]?\s*(\d{1,9})\b",
        r"\bNFS[-\s]?e\s*[:#]?\s*(\d{1,9})\b",
        r"\bNFS\s*[:#]?\s*(\d{1,9})\b",
    ]

    for padrao in padroes:

        encontrados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        for numero in encontrados:

            if numero not in candidatos:
                candidatos.append(numero)

    return candidatos


def extrair_cnpj_pdf(texto):

    if not texto:
        return []

    texto = normalizar_texto_pdf(texto)

    encontrados = set()

    # CNPJ formatado
    for cnpj in re.findall(
        r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b",
        texto
    ):
        encontrados.add(
            somente_digitos(cnpj)
        )

    # CNPJ sem máscara
    for cnpj in re.findall(
        r"(?<!\d)\d{14}(?!\d)",
        texto
    ):
        encontrados.add(cnpj)

    return list(encontrados)


def extrair_data_pdf(texto):

    if not texto:
        return []

    texto = normalizar_texto_pdf(texto)

    datas = []

    for data in re.findall(
        r"\b\d{2}/\d{2}/\d{4}\b",
        texto
    ):

        if data not in datas:
            datas.append(data)

    return datas


# ============================================================
# ÍNDICE DOS PDFs
# ============================================================

def indexar_pdfs():

    indice = []

    arquivos = sorted(
        PASTA_PROCESSAR.glob("*.pdf")
    )

    for pdf in arquivos:

        texto = extrair_texto_pdf(pdf)

        chaves = extrair_chaves_pdf(texto)
        numeros = extrair_numeros_pdf(texto)
        cnpjs = extrair_cnpj_pdf(texto)
        datas = extrair_data_pdf(texto)

        # Também procura chave no nome do arquivo
        chaves_nome = re.findall(
            r"(?<!\d)\d{44}(?!\d)",
            pdf.stem
        )

        for chave in chaves_nome:
            if chave not in chaves:
                chaves.append(chave)

        indice.append({
            "arquivo": pdf,
            "texto": texto,
            "chaves": chaves,
            "numeros": numeros,
            "cnpjs": cnpjs,
            "datas": datas,
        })

    return indice


# ============================================================
# ASSOCIAÇÃO XML ↔ PDF
# ============================================================

def associar_pdf(
    registro_xml,
    indice_pdfs,
):

    chave_xml = registro_xml["chave"]
    numero_xml = registro_xml["numero"]
    tipo_xml = registro_xml["tipo"]

    candidatos = []

    # --------------------------------------------------------
    # 1. Correspondência EXATA pela chave
    # --------------------------------------------------------

    if chave_xml:

        for pdf in indice_pdfs:

            if chave_xml in pdf["chaves"]:

                candidatos.append({
                    "pdf": pdf,
                    "metodo": "CHAVE EXATA"
                })

        if len(candidatos) == 1:
            return candidatos[0], ""

        if len(candidatos) > 1:
            return (
                None,
                "Mais de um PDF possui a mesma chave de acesso"
            )

    # --------------------------------------------------------
    # 2. Tentativa pelo número + tipo
    #
    # Só fazemos isso quando houver exatamente UM candidato.
    # --------------------------------------------------------

    if numero_xml:

        numero_limpo = somente_digitos(numero_xml)

        candidatos_numero = []

        for pdf in indice_pdfs:

            for numero_pdf in pdf["numeros"]:

                if somente_digitos(numero_pdf) == numero_limpo:

                    candidatos_numero.append(pdf)
                    break

        if len(candidatos_numero) == 1:

            return (
                {
                    "pdf": candidatos_numero[0],
                    "metodo": "NÚMERO DO DOCUMENTO"
                },
                ""
            )

        if len(candidatos_numero) > 1:

            # Não escolher no chute.
            return (
                None,
                "Número encontrado em mais de um PDF"
            )

    return (
        None,
        "PDF correspondente não identificado"
    )


# ============================================================
# PROCESSAMENTO DOS XMLs
# ============================================================

def processar_xmls(clientes, indice_pdfs):

    registros = []

    arquivos_xml = sorted(
        PASTA_PROCESSAR.glob("*.xml")
    )

    for arquivo in arquivos_xml:

        registro = {
            "arquivo": arquivo,
            "status": "REVISÃO",
            "tipo": "",
            "numero": "",
            "data": None,
            "chave": "",
            "cliente": None,
            "direcao": "",
            "destino": None,
            "nome_xml": "",
            "pdf": None,
            "nome_pdf": "",
            "metodo_pdf": "",
            "motivo": "",
        }

        arvore = ler_xml(arquivo)

        if arvore is None:

            registro["motivo"] = (
                "Não foi possível ler o XML"
            )

            registros.append(registro)
            continue

        root = arvore.getroot()

        tipo = identificar_tipo_xml(root)

        registro["tipo"] = tipo

        numero = extrair_numero_xml(
            root,
            tipo
        )

        registro["numero"] = numero

        registro["data"] = extrair_data_xml(root)

        registro["chave"] = extrair_chave_xml(
            root,
            arquivo
        )

        emitente = extrair_emitente(root)
        destinatario = extrair_destinatario(root)

        if tipo == "DESCONHECIDO":

            registro["motivo"] = (
                "Tipo de documento não identificado"
            )

            registros.append(registro)
            continue

        if not numero:

            registro["motivo"] = (
                f"{tipo} sem número identificável"
            )

            registros.append(registro)
            continue

        cliente, direcao, erro_cliente = identificar_cliente(
            clientes,
            emitente,
            destinatario
        )

        if not cliente:

            registro["motivo"] = erro_cliente

            registros.append(registro)
            continue

        if not registro["data"]:

            registro["motivo"] = (
                "Data de emissão não identificada"
            )

            registros.append(registro)
            continue

        registro["cliente"] = cliente
        registro["direcao"] = direcao

        # ----------------------------------------------------
        # Nome final
        # ----------------------------------------------------

        if tipo == "NF-e":

            registro["nome_xml"] = (
                f"NF {numero}.xml"
            )

        elif tipo == "NFS-e":

            registro["nome_xml"] = (
                f"NFS {numero}.xml"
            )

        registro["destino"] = montar_destino(
            cliente,
            direcao,
            registro["data"]
        )

        # ----------------------------------------------------
        # Associação do PDF
        # ----------------------------------------------------

        associacao, erro_pdf = associar_pdf(
            registro,
            indice_pdfs
        )

        if associacao:

            pdf_info = associacao["pdf"]

            registro["pdf"] = pdf_info["arquivo"]
            registro["metodo_pdf"] = associacao["metodo"]

            if tipo == "NF-e":

                registro["nome_pdf"] = (
                    f"NF {numero}.pdf"
                )

            elif tipo == "NFS-e":

                registro["nome_pdf"] = (
                    f"NFS {numero}.pdf"
                )

            registro["status"] = "OK"

            registro["motivo"] = (
                f"{tipo} identificada | "
                f"PDF associado por {registro['metodo_pdf']}"
            )

        else:

            registro["status"] = "OK"

            registro["motivo"] = (
                f"{tipo} identificada | "
                f"{erro_pdf}"
            )

        registros.append(registro)

    return registros


# ============================================================
# PDF ÓRFÃOS
# ============================================================

def identificar_pdfs_associados(registros):

    associados = set()

    for registro in registros:

        if registro["pdf"]:

            associados.add(
                registro["pdf"].resolve()
            )

    return associados


def listar_pdfs_orfaos(
    indice_pdfs,
    associados
):

    orfaos = []

    for pdf in indice_pdfs:

        caminho = pdf["arquivo"].resolve()

        if caminho not in associados:

            orfaos.append(pdf)

    return orfaos


# ============================================================
# RELATÓRIO
# ============================================================

def escrever_linha(arquivo, texto=""):
    arquivo.write(texto + "\n")


def escrever_relatorio(
    registros,
    indice_pdfs,
    pdfs_orfaos
):

    xml_ok = sum(
        1
        for r in registros
        if r["status"] == "OK"
    )

    xml_revisao = sum(
        1
        for r in registros
        if r["status"] == "REVISÃO"
    )

    pares = sum(
        1
        for r in registros
        if r["pdf"] is not None
    )

    xml_sem_pdf = sum(
        1
        for r in registros
        if r["pdf"] is None
    )

    with open(
        ARQUIVO_RELATORIO,
        "w",
        encoding="utf-8"
    ) as f:

        escrever_linha(
            f,
            "=" * 80
        )

        escrever_linha(
            f,
            "RELATÓRIO DE SIMULAÇÃO - V7"
        )

        escrever_linha(
            f,
            "=" * 80
        )

        escrever_linha(f)

        escrever_linha(
            f,
            "ATENÇÃO: NENHUM ARQUIVO FOI "
            "MOVIDO OU RENOMEADO."
        )

        escrever_linha(f)

        escrever_linha(
            f,
            f"XML processados: {len(registros)}"
        )

        escrever_linha(
            f,
            f"PDF encontrados: {len(indice_pdfs)}"
        )

        escrever_linha(
            f,
            f"XML OK: {xml_ok}"
        )

        escrever_linha(
            f,
            f"XML REVISÃO: {xml_revisao}"
        )

        escrever_linha(
            f,
            f"PARES XML + PDF: {pares}"
        )

        escrever_linha(
            f,
            f"XML SEM PDF: {xml_sem_pdf}"
        )

        escrever_linha(
            f,
            f"PDF ÓRFÃOS: {len(pdfs_orfaos)}"
        )

        escrever_linha(f)

        escrever_linha(
            f,
            "-" * 80
        )

        escrever_linha(
            f,
            "DETALHAMENTO DOS XMLs"
        )

        escrever_linha(
            f,
            "-" * 80
        )

        for r in registros:

            escrever_linha(f)

            escrever_linha(
                f,
                f"XML: {r['arquivo'].name}"
            )

            escrever_linha(
                f,
                f"Status: {r['status']}"
            )

            escrever_linha(
                f,
                f"Tipo: {r['tipo']}"
            )

            escrever_linha(
                f,
                f"Número: {r['numero'] or '?'}"
            )

            if r["data"]:

                data_formatada = (
                    r["data"].strftime("%d/%m/%Y")
                )

            else:

                data_formatada = "?"

            escrever_linha(
                f,
                f"Data: {data_formatada}"
            )

            escrever_linha(
                f,
                f"Chave: {r['chave'] or '?'}"
            )

            if r["cliente"]:

                escrever_linha(
                    f,
                    f"Cliente: {r['cliente']['cod']}"
                )

                escrever_linha(
                    f,
                    f"Razão Social: "
                    f"{r['cliente']['razao']}"
                )

            else:

                escrever_linha(
                    f,
                    "Cliente: ?"
                )

            escrever_linha(
                f,
                f"Direção: {r['direcao'] or '?'}"
            )

            escrever_linha(
                f,
                f"Nome XML: "
                f"{r['nome_xml'] or '?'}"
            )

            escrever_linha(
                f,
                f"PDF associado: "
                f"{r['pdf'].name if r['pdf'] else '?'}"
            )

            escrever_linha(
                f,
                f"Método PDF: "
                f"{r['metodo_pdf'] or '?'}"
            )

            escrever_linha(
                f,
                f"Nome PDF: "
                f"{r['nome_pdf'] or '?'}"
            )

            escrever_linha(
                f,
                f"Destino: "
                f"{r['destino'] if r['destino'] else '?'}"
            )

            escrever_linha(
                f,
                f"Motivo: {r['motivo']}"
            )

        # ----------------------------------------------------
        # PDFs ÓRFÃOS
        # ----------------------------------------------------

        escrever_linha(f)

        escrever_linha(
            f,
            "=" * 80
        )

        escrever_linha(
            f,
            "PDFs ÓRFÃOS / NÃO ASSOCIADOS"
        )

        escrever_linha(
            f,
            "=" * 80
        )

        if not pdfs_orfaos:

            escrever_linha(
                f,
                "Nenhum PDF órfão."
            )

        else:

            for pdf in pdfs_orfaos:

                escrever_linha(f)

                escrever_linha(
                    f,
                    f"PDF: {pdf['arquivo'].name}"
                )

                escrever_linha(
                    f,
                    "Chaves encontradas: "
                    + (
                        ", ".join(pdf["chaves"])
                        if pdf["chaves"]
                        else "nenhuma"
                    )
                )

                escrever_linha(
                    f,
                    "Números encontrados: "
                    + (
                        ", ".join(pdf["numeros"])
                        if pdf["numeros"]
                        else "nenhum"
                    )
                )

                escrever_linha(
                    f,
                    "CNPJs encontrados: "
                    + (
                        ", ".join(pdf["cnpjs"])
                        if pdf["cnpjs"]
                        else "nenhum"
                    )
                )

                escrever_linha(
                    f,
                    "Datas encontradas: "
                    + (
                        ", ".join(pdf["datas"][:10])
                        if pdf["datas"]
                        else "nenhuma"
                    )
                )

        escrever_linha(f)

        escrever_linha(
            f,
            "=" * 80
        )

        escrever_linha(
            f,
            "FIM DA SIMULAÇÃO"
        )

        escrever_linha(
            f,
            "NENHUM ARQUIVO FOI ALTERADO."
        )

        escrever_linha(
            f,
            "=" * 80
        )


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print("PROCESSAMENTO XML + PDF NF-e / NFS-e - V7")
    print("=" * 70)

    if not PASTA_PROCESSAR.exists():

        print(
            f"\nERRO: pasta '{PASTA_PROCESSAR}' não encontrada."
        )

        return

    if not ARQUIVO_BANCO.exists():

        print(
            f"\nERRO: arquivo '{ARQUIVO_BANCO}' não encontrado."
        )

        return

    print("\nCarregando banco.xlsx...")

    try:

        clientes = carregar_clientes()

    except Exception as e:

        print(
            f"\nERRO ao carregar banco.xlsx: {e}"
        )

        return

    print(
        f"Clientes carregados: {len(clientes)}"
    )

    print("\nIndexando PDFs...")

    indice_pdfs = indexar_pdfs()

    print(
        f"PDFs indexados: {len(indice_pdfs)}"
    )

    print("\nProcessando XMLs...")

    registros = processar_xmls(
        clientes,
        indice_pdfs
    )

    associados = identificar_pdfs_associados(
        registros
    )

    pdfs_orfaos = listar_pdfs_orfaos(
        indice_pdfs,
        associados
    )

    escrever_relatorio(
        registros,
        indice_pdfs,
        pdfs_orfaos
    )

    print("\n" + "=" * 70)

    print(
        f"XMLs processados: {len(registros)}"
    )

    print(
        f"PDFs encontrados: {len(indice_pdfs)}"
    )

    print(
        "Pares XML + PDF: "
        + str(
            sum(
                1
                for r in registros
                if r["pdf"]
            )
        )
    )

    print(
        f"PDFs órfãos: {len(pdfs_orfaos)}"
    )

    print(
        f"\nRelatório criado: {ARQUIVO_RELATORIO}"
    )

    print(
        "\nNENHUM ARQUIVO FOI MOVIDO OU RENOMEADO."
    )

    print("=" * 70)


if __name__ == "__main__":
    main()
