import os
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import load_workbook

# ============================================================
# CONFIGURAÇÕES
# ============================================================

PASTA_PROCESSAR = "A_PROCESSAR"
ARQUIVO_EXCEL = "banco.xlsx"
ARQUIVO_RELATORIO = "relatorio_xml_pdf_v6.txt"

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def limpar_cnpj(valor):
    if valor is None:
        return ""

    return re.sub(r"\D", "", str(valor).strip())


def nome_tag(tag):
    if "}" in tag:
        return tag.split("}", 1)[1]

    return tag


def texto(elemento):
    if elemento is None or elemento.text is None:
        return ""

    return elemento.text.strip()


def encontrar_primeiro(root, nomes):
    nomes = set(nomes)

    for elemento in root.iter():
        if nome_tag(elemento.tag) in nomes:
            return elemento

    return None


def encontrar_todos(root, nomes):
    nomes = set(nomes)

    encontrados = []

    for elemento in root.iter():
        if nome_tag(elemento.tag) in nomes:
            encontrados.append(elemento)

    return encontrados


# ============================================================
# DATA
# ============================================================

def extrair_data_valor(valor):
    if not valor:
        return None

    valor = valor.strip()

    # ISO com horário
    match = re.search(
        r"(\d{4})-(\d{2})-(\d{2})",
        valor
    )

    if match:
        try:
            return datetime(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3))
            )
        except ValueError:
            pass

    # Data brasileira
    match = re.search(
        r"(\d{2})/(\d{2})/(\d{4})",
        valor
    )

    if match:
        try:
            return datetime(
                int(match.group(3)),
                int(match.group(2)),
                int(match.group(1))
            )
        except ValueError:
            pass

    return None


def extrair_data_xml(root):
    tags_prioritarias = [
        "dhEmi",
        "dEmi"
    ]

    for tag in tags_prioritarias:
        elementos = encontrar_todos(root, [tag])

        for elemento in elementos:
            data = extrair_data_valor(texto(elemento))

            if data:
                return data

    tags_nfse = [
        "DataEmissao",
        "dataEmissao",
        "DataEmissaoNfse",
        "dataEmissaoNfse",
        "DataEmissaoRps",
        "dataEmissaoRps"
    ]

    for tag in tags_nfse:
        elementos = encontrar_todos(root, [tag])

        for elemento in elementos:
            data = extrair_data_valor(texto(elemento))

            if data:
                return data

    return None


# ============================================================
# TIPO DO DOCUMENTO
# ============================================================

def identificar_tipo_xml(root):
    tags = {
        nome_tag(elemento.tag)
        for elemento in root.iter()
    }

    if "infNFe" in tags or "NFe" in tags:
        return "NF-e"

    if (
        "NFSe" in tags
        or "infNFSe" in tags
        or "nNFSe" in tags
        or "vISSQN" in tags
    ):
        return "NFS-e"

    return "DESCONHECIDO"


# ============================================================
# NÚMERO DA NOTA
# ============================================================

def extrair_numero_nfe(root):
    elemento = encontrar_primeiro(root, ["nNF"])

    if elemento is not None:
        numero = texto(elemento)

        if numero:
            return numero

    return None


def extrair_numero_nfse(root):
    tags = [
        "nNFSe",
        "NumeroNfse",
        "numeroNfse",
        "NumeroNFSe",
        "numeroNFSe",
        "NúmeroNfse"
    ]

    for tag in tags:
        elementos = encontrar_todos(root, [tag])

        for elemento in elementos:
            numero = texto(elemento)

            if numero and re.fullmatch(r"\d+", numero):
                return numero

    return None


# ============================================================
# CNPJ
# ============================================================

def extrair_cnpj_emitente(root):
    emits = encontrar_todos(root, ["emit"])

    for emit in emits:

        elemento = encontrar_primeiro(
            emit,
            ["CNPJ"]
        )

        if elemento is not None:
            cnpj = limpar_cnpj(texto(elemento))

            if cnpj:
                return cnpj

    return None


def extrair_cnpj_destinatario(root):
    dests = encontrar_todos(root, ["dest"])

    for dest in dests:

        elemento = encontrar_primeiro(
            dest,
            ["CNPJ"]
        )

        if elemento is not None:
            cnpj = limpar_cnpj(texto(elemento))

            if cnpj:
                return cnpj

    return None


# ============================================================
# CHAVE DE ACESSO
# ============================================================

def extrair_chave_xml(root, caminho):
    # Primeiro: Id do infNFe
    for elemento in root.iter():

        if nome_tag(elemento.tag) == "infNFe":

            identificador = elemento.attrib.get(
                "Id",
                ""
            )

            chave = re.sub(
                r"\D",
                "",
                identificador
            )

            if len(chave) == 44:
                return chave

    # Segundo: procura chave de 44 dígitos no arquivo
    try:

        with open(
            caminho,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as arquivo:

            conteudo = arquivo.read()

        encontrados = re.findall(
            r"\d{44}",
            conteudo
        )

        if encontrados:
            return encontrados[0]

    except Exception:
        pass

    return None


# ============================================================
# BANCO DE CLIENTES
# ============================================================

def carregar_clientes():

    wb = load_workbook(
        ARQUIVO_EXCEL,
        data_only=True
    )

    ws = wb.active

    cabecalhos = {}

    for coluna in range(
        1,
        ws.max_column + 1
    ):

        valor = ws.cell(
            1,
            coluna
        ).value

        if valor:
            cabecalhos[
                str(valor).strip().upper()
            ] = coluna

    col_cod = cabecalhos.get("COD")
    col_razao = cabecalhos.get("RAZÃO SOCIAL")
    col_cnpj = cabecalhos.get("CNPJ")
    col_caminho = cabecalhos.get(
        "CAMINHO ABSOLUTO"
    )

    if not col_cnpj or not col_caminho:
        raise Exception(
            "O banco.xlsx precisa possuir "
            "CNPJ e CAMINHO ABSOLUTO."
        )

    clientes = {}

    for linha in range(
        2,
        ws.max_row + 1
    ):

        cnpj = limpar_cnpj(
            ws.cell(
                linha,
                col_cnpj
            ).value
        )

        caminho = ws.cell(
            linha,
            col_caminho
        ).value

        if not cnpj or not caminho:
            continue

        cliente = {
            "cod": (
                ws.cell(
                    linha,
                    col_cod
                ).value
                if col_cod else ""
            ),

            "razao": (
                ws.cell(
                    linha,
                    col_razao
                ).value
                if col_razao else ""
            ),

            "cnpj": cnpj,

            "caminho": str(
                caminho
            ).strip()
        }

        clientes.setdefault(
            cnpj,
            []
        ).append(cliente)

    return clientes


# ============================================================
# IDENTIFICAÇÃO DO CLIENTE
# ============================================================

def identificar_cliente(
    cnpj_emitente,
    cnpj_destinatario,
    clientes
):

    candidatos = []

    # Cliente é destinatário → ENTRADA
    if cnpj_destinatario in clientes:

        for cliente in clientes[
            cnpj_destinatario
        ]:

            candidatos.append(
                (
                    cliente,
                    "ENTRADAS"
                )
            )

    # Cliente é emitente → SAÍDA
    if cnpj_emitente in clientes:

        for cliente in clientes[
            cnpj_emitente
        ]:

            candidatos.append(
                (
                    cliente,
                    "SAIDAS"
                )
            )

    # Remove duplicidades
    unicos = []

    for candidato in candidatos:

        if candidato not in unicos:
            unicos.append(candidato)

    if not unicos:
        return (
            None,
            None,
            "CNPJ não encontrado no banco.xlsx"
        )

    if len(unicos) == 1:
        return (
            unicos[0][0],
            unicos[0][1],
            "OK"
        )

    # Se houver exatamente o mesmo cliente/direção
    combinacoes = set()

    for cliente, direcao in unicos:

        combinacoes.add(
            (
                cliente["caminho"],
                direcao
            )
        )

    if len(combinacoes) == 1:

        return (
            unicos[0][0],
            unicos[0][1],
            "OK"
        )

    return (
        None,
        None,
        "Mais de um cliente/direção possível"
    )


# ============================================================
# PROCESSAMENTO XML
# ============================================================

def processar_xml(
    caminho,
    clientes
):

    resultado = {
        "arquivo": os.path.basename(caminho),
        "status": "REVISÃO",
        "tipo": None,
        "numero": None,
        "data": None,
        "chave": None,
        "cliente": None,
        "direcao": None,
        "destino": None,
        "nome_xml": None,
        "pdf": None,
        "nome_pdf": None,
        "motivo": None
    }

    try:

        tree = ET.parse(caminho)
        root = tree.getroot()

    except Exception as erro:

        resultado["motivo"] = (
            f"Erro ao ler XML: {erro}"
        )

        return resultado

    tipo = identificar_tipo_xml(root)

    resultado["tipo"] = tipo

    data = extrair_data_xml(root)

    resultado["data"] = data

    chave = extrair_chave_xml(
        root,
        caminho
    )

    resultado["chave"] = chave

    cnpj_emitente = (
        extrair_cnpj_emitente(root)
    )

    cnpj_destinatario = (
        extrair_cnpj_destinatario(root)
    )

    # --------------------------------------------------------
    # NF-e
    # --------------------------------------------------------

    if tipo == "NF-e":

        numero = extrair_numero_nfe(
            root
        )

        resultado["numero"] = numero

        # Caso sem nNF
        if not numero:

            resultado["motivo"] = (
                "NF-e sem nNF; "
                "será necessário investigar a estrutura"
            )

            return resultado

        if not data:

            resultado["motivo"] = (
                "NF-e sem data de emissão"
            )

            return resultado

        cliente, direcao, motivo = (
            identificar_cliente(
                cnpj_emitente,
                cnpj_destinatario,
                clientes
            )
        )

        if not cliente:

            resultado["motivo"] = motivo

            return resultado

        resultado["cliente"] = cliente
        resultado["direcao"] = direcao

        resultado["nome_xml"] = (
            f"NF {numero}.xml"
        )

        resultado["destino"] = os.path.join(
            cliente["caminho"],
            "NF",
            direcao,
            str(data.year),
            f"{data.month:02d}-{str(data.year)[2:]}"
        )

        resultado["status"] = "OK"

        resultado["motivo"] = (
            "NF-e identificada"
        )

        return resultado

    # --------------------------------------------------------
    # NFS-e
    # --------------------------------------------------------

    if tipo == "NFS-e":

        numero = extrair_numero_nfse(
            root
        )

        resultado["numero"] = numero

        if not numero:

            resultado["motivo"] = (
                "NFS-e sem número"
            )

            return resultado

        if not data:

            resultado["motivo"] = (
                "NFS-e sem data de emissão"
            )

            return resultado

        cliente, direcao, motivo = (
            identificar_cliente(
                cnpj_emitente,
                cnpj_destinatario,
                clientes
            )
        )

        if not cliente:

            resultado["motivo"] = motivo

            return resultado

        resultado["cliente"] = cliente
        resultado["direcao"] = direcao

        resultado["nome_xml"] = (
            f"NFS {numero}.xml"
        )

        resultado["destino"] = os.path.join(
            cliente["caminho"],
            "NF",
            direcao,
            str(data.year),
            f"{data.month:02d}-{str(data.year)[2:]}"
        )

        resultado["status"] = "OK"

        resultado["motivo"] = (
            "NFS-e identificada"
        )

        return resultado

    resultado["motivo"] = (
        "Tipo de XML não identificado"
    )

    return resultado


# ============================================================
# PDF
# ============================================================

def extrair_chave_pdf(
    caminho
):

    try:

        from pypdf import PdfReader

    except ImportError:

        return None

    try:

        reader = PdfReader(
            caminho
        )

        texto_total = ""

        for pagina in reader.pages:

            texto = pagina.extract_text()

            if texto:
                texto_total += "\n" + texto

        # Junta números quebrados por espaços,
        # mantendo também o texto original.
        somente_numeros = re.sub(
            r"\D",
            "",
            texto_total
        )

        encontrados = re.findall(
            r"\d{44}",
            somente_numeros
        )

        if encontrados:
            return encontrados[0]

        # Procura também sequências de 44 dígitos
        # no texto original.
        encontrados = re.findall(
            r"\b\d{44}\b",
            texto_total
        )

        if encontrados:
            return encontrados[0]

    except Exception:
        pass

    return None


def processar_pdfs(
    clientes
):

    pdfs = {}

    arquivos = os.listdir(
        PASTA_PROCESSAR
    )

    for arquivo in arquivos:

        if not arquivo.lower().endswith(
            ".pdf"
        ):
            continue

        caminho = os.path.join(
            PASTA_PROCESSAR,
            arquivo
        )

        chave = extrair_chave_pdf(
            caminho
        )

        registro = {
            "arquivo": arquivo,
            "caminho": caminho,
            "chave": chave
        }

        if chave:

            pdfs.setdefault(
                chave,
                []
            ).append(
                registro
            )

    return pdfs


# ============================================================
# ASSOCIAÇÃO XML + PDF
# ============================================================

def associar_pdf(
    resultado_xml,
    pdfs
):

    chave = resultado_xml["chave"]

    if not chave:

        resultado_xml["motivo"] = (
            "XML sem chave de acesso; "
            "PDF não pode ser associado automaticamente"
        )

        resultado_xml["status"] = "REVISÃO"

        return

    candidatos = pdfs.get(
        chave,
        []
    )

    if not candidatos:

        resultado_xml["motivo"] += (
            " | PDF correspondente não encontrado"
        )

        # XML continua OK individualmente,
        # mas não existe par.
        return

    if len(candidatos) > 1:

        resultado_xml["status"] = "REVISÃO"

        resultado_xml["motivo"] = (
            "Mais de um PDF encontrado "
            "para a mesma chave"
        )

        return

    pdf = candidatos[0]

    resultado_xml["pdf"] = pdf

    if resultado_xml["tipo"] == "NF-e":

        resultado_xml["nome_pdf"] = (
            f"NF {resultado_xml['numero']}.pdf"
        )

    elif resultado_xml["tipo"] == "NFS-e":

        resultado_xml["nome_pdf"] = (
            f"NFS {resultado_xml['numero']}.pdf"
        )

    resultado_xml["motivo"] += (
        " | PDF associado pela chave de acesso"
    )


# ============================================================
# RELATÓRIO
# ============================================================

def gerar_relatorio(
    resultados,
    pdfs
):

    total = len(resultados)

    ok = sum(
        1
        for resultado in resultados
        if resultado["status"] == "OK"
    )

    revisao = total - ok

    pares = sum(
        1
        for resultado in resultados
        if resultado["pdf"] is not None
    )

    sem_pdf = total - pares

    with open(
        ARQUIVO_RELATORIO,
        "w",
        encoding="utf-8"
    ) as arquivo:

        arquivo.write(
            "=" * 80 + "\n"
        )

        arquivo.write(
            "RELATÓRIO DE SIMULAÇÃO - V6\n"
        )

        arquivo.write(
            "=" * 80 + "\n\n"
        )

        arquivo.write(
            "ATENÇÃO: NENHUM ARQUIVO FOI "
            "MOVIDO OU RENOMEADO.\n\n"
        )

        arquivo.write(
            f"XML processados: {total}\n"
        )

        arquivo.write(
            f"PDF encontrados: "
            f"{sum(len(v) for v in pdfs.values())}\n"
        )

        arquivo.write(
            f"XML OK: {ok}\n"
        )

        arquivo.write(
            f"XML REVISÃO: {revisao}\n"
        )

        arquivo.write(
            f"PARES XML + PDF: {pares}\n"
        )

        arquivo.write(
            f"XML SEM PDF: {sem_pdf}\n\n"
        )

        arquivo.write(
            "-" * 80 + "\n"
        )

        arquivo.write(
            "DETALHAMENTO\n"
        )

        arquivo.write(
            "-" * 80 + "\n\n"
        )

        for resultado in resultados:

            arquivo.write(
                f"XML: {resultado['arquivo']}\n"
            )

            arquivo.write(
                f"Status: {resultado['status']}\n"
            )

            arquivo.write(
                f"Tipo: {resultado['tipo'] or '?'}\n"
            )

            arquivo.write(
                f"Número: "
                f"{resultado['numero'] or '?'}\n"
            )

            if resultado["data"]:

                arquivo.write(
                    "Data: "
                    f"{resultado['data'].strftime('%d/%m/%Y')}\n"
                )

            else:

                arquivo.write(
                    "Data: ?\n"
                )

            arquivo.write(
                "Chave: "
                f"{resultado['chave'] or '?'}\n"
            )

            if resultado["cliente"]:

                arquivo.write(
                    "Cliente: "
                    f"{resultado['cliente']['cod']}\n"
                )

                arquivo.write(
                    "Razão Social: "
                    f"{resultado['cliente']['razao']}\n"
                )

            else:

                arquivo.write(
                    "Cliente: ?\n"
                )

            arquivo.write(
                "Direção: "
                f"{resultado['direcao'] or '?'}\n"
            )

            arquivo.write(
                "Nome XML: "
                f"{resultado['nome_xml'] or '?'}\n"
            )

            arquivo.write(
                "PDF associado: "
                f"{resultado['pdf']['arquivo'] if resultado['pdf'] else '?'}\n"
            )

            arquivo.write(
                "Nome PDF: "
                f"{resultado['nome_pdf'] or '?'}\n"
            )

            arquivo.write(
                "Destino: "
                f"{resultado['destino'] or '?'}\n"
            )

            arquivo.write(
                "Motivo: "
                f"{resultado['motivo'] or '?'}\n"
            )

            arquivo.write(
                "\n"
            )

        arquivo.write(
            "=" * 80 + "\n"
        )

        arquivo.write(
            "FIM DA SIMULAÇÃO\n"
        )

        arquivo.write(
            "NENHUM ARQUIVO FOI ALTERADO.\n"
        )

        arquivo.write(
            "=" * 80 + "\n"
        )


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print("PROCESSAMENTO XML + PDF NF - V6")
    print("=" * 70)
    print()
    print("MODO SIMULAÇÃO")
    print("Nenhum arquivo será movido ou renomeado.")
    print()

    if not os.path.exists(
        PASTA_PROCESSAR
    ):

        print(
            f"ERRO: pasta '{PASTA_PROCESSAR}' não encontrada."
        )

        return

    if not os.path.exists(
        ARQUIVO_EXCEL
    ):

        print(
            f"ERRO: '{ARQUIVO_EXCEL}' não encontrado."
        )

        return

    try:

        clientes = carregar_clientes()

    except Exception as erro:

        print(
            f"ERRO ao carregar banco.xlsx: {erro}"
        )

        return

    arquivos = os.listdir(
        PASTA_PROCESSAR
    )

    arquivos_xml = sorted([
        arquivo
        for arquivo in arquivos
        if arquivo.lower().endswith(".xml")
    ])

    arquivos_pdf = sorted([
        arquivo
        for arquivo in arquivos
        if arquivo.lower().endswith(".pdf")
    ])

    print(
        f"XML encontrados: {len(arquivos_xml)}"
    )

    print(
        f"PDF encontrados: {len(arquivos_pdf)}"
    )

    print()

    # --------------------------------------------------------
    # Primeiro: indexa TODOS os PDFs pela chave
    # --------------------------------------------------------

    print("Indexando PDFs pela chave de acesso...")

    pdfs = processar_pdfs(
        clientes
    )

    print(
        f"Chaves de PDF identificadas: {len(pdfs)}"
    )

    print()

    # --------------------------------------------------------
    # Depois: processa XMLs
    # --------------------------------------------------------

    resultados = []

    for arquivo in arquivos_xml:

        caminho = os.path.join(
            PASTA_PROCESSAR,
            arquivo
        )

        print(
            f"Processando XML: {arquivo}"
        )

        resultado = processar_xml(
            caminho,
            clientes
        )

        # Associa PDF somente depois
        # que o XML foi interpretado.
        associar_pdf(
            resultado,
            pdfs
        )

        resultados.append(
            resultado
        )

    # --------------------------------------------------------
    # Relatório
    # --------------------------------------------------------

    gerar_relatorio(
        resultados,
        pdfs
    )

    print()
    print("=" * 70)
    print("SIMULAÇÃO V6 CONCLUÍDA")
    print("=" * 70)
    print()
    print(
        f"Relatório: {ARQUIVO_RELATORIO}"
    )
    print()
    print(
        "NENHUM ARQUIVO FOI MOVIDO OU RENOMEADO."
    )


if __name__ == "__main__":
    main()
