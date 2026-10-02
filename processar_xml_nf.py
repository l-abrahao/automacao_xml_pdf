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
ARQUIVO_RELATORIO = "relatorio_xml_pdf_v5.txt"

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def limpar_cnpj(valor):
    """Remove pontuação do CNPJ."""
    if valor is None:
        return ""

    valor = str(valor).strip()
    return re.sub(r"\D", "", valor)


def texto_elemento(elemento):
    """Retorna o texto de um elemento XML."""
    if elemento is None or elemento.text is None:
        return ""

    return elemento.text.strip()


def nome_tag(tag):
    """Remove namespace da tag XML."""
    if "}" in tag:
        return tag.split("}", 1)[1]

    return tag


def encontrar_primeiro(root, nomes):
    """Procura o primeiro elemento com uma das tags informadas."""
    nomes = set(nomes)

    for elemento in root.iter():
        if nome_tag(elemento.tag) in nomes:
            return elemento

    return None


def encontrar_todos(root, nomes):
    """Retorna todos os elementos com as tags informadas."""
    nomes = set(nomes)

    encontrados = []

    for elemento in root.iter():
        if nome_tag(elemento.tag) in nomes:
            encontrados.append(elemento)

    return encontrados


def extrair_data_valor(valor):
    """Converte diferentes formatos de data para dd/mm/aaaa."""
    if not valor:
        return None

    valor = valor.strip()

    formatos = [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%d",
        "%d/%m/%Y",
    ]

    for formato in formatos:
        try:
            data = datetime.strptime(valor[:25], formato)
            return data
        except ValueError:
            pass

    # Tenta localizar uma data ISO dentro do texto
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", valor)

    if match:
        try:
            return datetime(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3))
            )
        except ValueError:
            pass

    # Tenta localizar dd/mm/yyyy
    match = re.search(r"(\d{2})/(\d{2})/(\d{4})", valor)

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
    """
    Procura a data de emissão em toda a estrutura do XML.

    NF-e:
        dhEmi
        dEmi

    NFS-e:
        procura também por tags relacionadas à emissão.
    """

    # Prioridade para NF-e
    for tag in ["dhEmi", "dEmi"]:
        elemento = encontrar_primeiro(root, [tag])

        if elemento is not None:
            data = extrair_data_valor(texto_elemento(elemento))

            if data:
                return data

    # Possíveis tags de emissão em NFS-e
    for tag in [
        "dhEmi",
        "dEmi",
        "DataEmissao",
        "dataEmissao",
        "DataEmissaoRps",
        "dataEmissaoRps",
        "DataEmissaoNfse",
        "dataEmissaoNfse",
    ]:
        elementos = encontrar_todos(root, [tag])

        for elemento in elementos:
            data = extrair_data_valor(texto_elemento(elemento))

            if data:
                return data

    return None


def extrair_numero_nfe(root):
    """Extrai o número da NF-e pela tag nNF."""
    elemento = encontrar_primeiro(root, ["nNF"])

    if elemento is None:
        return None

    numero = texto_elemento(elemento)

    if numero:
        return numero

    return None


def extrair_numero_nfse(root):
    """
    Extrai o número da NFS-e.

    Prioridade:
    nNFSe
    NumeroNfse
    numeroNfse
    NúmeroNfse
    Número
    numero
    """

    tags = [
        "nNFSe",
        "NumeroNfse",
        "numeroNfse",
        "NúmeroNfse",
        "NumeroNFSe",
        "numeroNFSe",
        "Número",
        "numero",
    ]

    for tag in tags:
        elementos = encontrar_todos(root, [tag])

        for elemento in elementos:
            numero = texto_elemento(elemento)

            if numero:
                # Remove espaços desnecessários
                numero = numero.strip()

                # Aceita números simples
                if re.fullmatch(r"\d+", numero):
                    return numero

    return None


def extrair_cnpj_emitente(root):
    """Extrai CNPJ do emitente."""
    elementos = encontrar_todos(root, ["emit"])

    for emit in elementos:
        cnpj = encontrar_primeiro(emit, ["CNPJ"])

        if cnpj is not None:
            valor = limpar_cnpj(texto_elemento(cnpj))

            if valor:
                return valor

    # Fallback
    elementos_cnpj = encontrar_todos(root, ["CNPJ"])

    for elemento in elementos_cnpj:
        valor = limpar_cnpj(texto_elemento(elemento))

        if valor:
            return valor

    return None


def extrair_cnpj_destinatario(root):
    """Extrai CNPJ do destinatário."""
    elementos = encontrar_todos(root, ["dest"])

    for dest in elementos:
        cnpj = encontrar_primeiro(dest, ["CNPJ"])

        if cnpj is not None:
            valor = limpar_cnpj(texto_elemento(cnpj))

            if valor:
                return valor

    return None


def extrair_chave_acesso(root, caminho_arquivo):
    """
    Tenta extrair a chave de acesso da NF-e.

    Primeiro procura infNFe Id.
    Depois procura chave numérica no conteúdo.
    """

    for elemento in root.iter():
        tag = nome_tag(elemento.tag)

        if tag == "infNFe":
            identificador = elemento.attrib.get("Id", "")

            identificador = re.sub(r"\D", "", identificador)

            if len(identificador) == 44:
                return identificador

    # Fallback: procura 44 dígitos no XML
    try:
        with open(caminho_arquivo, "r", encoding="utf-8", errors="ignore") as arquivo:
            conteudo = arquivo.read()

        chaves = re.findall(r"\b\d{44}\b", conteudo)

        if chaves:
            return chaves[0]

    except Exception:
        pass

    return None


def identificar_tipo_xml(root):
    """
    Identifica:
    NF-e
    NFS-e
    DESCONHECIDO
    """

    tags = {nome_tag(elemento.tag) for elemento in root.iter()}

    # NF-e modelo 55
    if "infNFe" in tags or "NFe" in tags:
        return "NF-e"

    # Estrutura típica de NFS-e
    if (
        "NFSe" in tags
        or "infNFSe" in tags
        or "nNFSe" in tags
        or "vISSQN" in tags
    ):
        return "NFS-e"

    return "DESCONHECIDO"


def obter_modelo(root):
    """Procura a tag mod."""
    elemento = encontrar_primeiro(root, ["mod"])

    if elemento is None:
        return None

    return texto_elemento(elemento)


# ============================================================
# LEITURA DO EXCEL
# ============================================================

def carregar_clientes():
    clientes = {}

    wb = load_workbook(ARQUIVO_EXCEL, data_only=True)
    ws = wb.active

    cabecalhos = {}

    for coluna in range(1, ws.max_column + 1):
        valor = ws.cell(1, coluna).value

        if valor:
            cabecalhos[str(valor).strip().upper()] = coluna

    col_cnpj = cabecalhos.get("CNPJ")
    col_cod = cabecalhos.get("COD")
    col_razao = cabecalhos.get("RAZÃO SOCIAL")
    col_caminho = cabecalhos.get("CAMINHO ABSOLUTO")

    if not col_cnpj or not col_caminho:
        raise Exception(
            "O Excel precisa possuir as colunas CNPJ e CAMINHO ABSOLUTO."
        )

    for linha in range(2, ws.max_row + 1):

        cnpj = limpar_cnpj(ws.cell(linha, col_cnpj).value)
        caminho = ws.cell(linha, col_caminho).value

        if not cnpj or not caminho:
            continue

        cliente = {
            "cnpj": cnpj,
            "cod": ws.cell(linha, col_cod).value if col_cod else "",
            "razao": ws.cell(linha, col_razao).value if col_razao else "",
            "caminho": str(caminho).strip(),
        }

        if cnpj not in clientes:
            clientes[cnpj] = []

        clientes[cnpj].append(cliente)

    return clientes


# ============================================================
# IDENTIFICAÇÃO DO CLIENTE
# ============================================================

def identificar_cliente(cnpj_emitente, cnpj_destinatario, clientes):
    """
    Identifica o cliente.

    Retorna:
        cliente
        ENTRADAS / SAIDAS
        motivo
    """

    candidatos = []

    if cnpj_emitente in clientes:
        for cliente in clientes[cnpj_emitente]:
            candidatos.append((cliente, "SAIDAS"))

    if cnpj_destinatario in clientes:
        for cliente in clientes[cnpj_destinatario]:
            candidatos.append((cliente, "ENTRADAS"))

    # Nenhum
    if not candidatos:
        return None, None, "CNPJ não encontrado no banco.xlsx"

    # Remove duplicidade exata
    unicos = []

    for candidato in candidatos:
        if candidato not in unicos:
            unicos.append(candidato)

    # Apenas um resultado
    if len(unicos) == 1:
        return unicos[0][0], unicos[0][1], "OK"

    # Mesmo cliente / mesma direção
    combinacoes = set()

    for cliente, direcao in unicos:
        combinacoes.add(
            (
                cliente["caminho"],
                direcao
            )
        )

    if len(combinacoes) == 1:
        return unicos[0][0], unicos[0][1], "OK"

    return None, None, "CNPJ corresponde a mais de um cliente/direção"


# ============================================================
# PROCESSAMENTO XML
# ============================================================

def processar_xml(caminho_xml, clientes):
    resultado = {
        "arquivo": os.path.basename(caminho_xml),
        "status": "REVISÃO",
        "tipo": None,
        "numero": None,
        "data": None,
        "cliente": None,
        "direcao": None,
        "chave": None,
        "nome_final": None,
        "destino": None,
        "motivo": None,
    }

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

    except Exception as erro:
        resultado["motivo"] = f"Erro ao ler XML: {erro}"
        return resultado

    tipo = identificar_tipo_xml(root)
    resultado["tipo"] = tipo

    data = extrair_data_xml(root)
    resultado["data"] = data

    cnpj_emitente = extrair_cnpj_emitente(root)
    cnpj_destinatario = extrair_cnpj_destinatario(root)

    resultado["chave"] = extrair_chave_acesso(root, caminho_xml)

    # ========================================================
    # NF-e
    # ========================================================

    if tipo == "NF-e":

        numero = extrair_numero_nfe(root)
        resultado["numero"] = numero

        if not numero:
            resultado["motivo"] = "NF-e sem número nNF"
            return resultado

        modelo = obter_modelo(root)

        if modelo and modelo != "55":
            resultado["motivo"] = f"Modelo diferente de 55: {modelo}"
            return resultado

        if not data:
            resultado["motivo"] = "NF-e sem data de emissão identificável"
            return resultado

        cliente, direcao, motivo = identificar_cliente(
            cnpj_emitente,
            cnpj_destinatario,
            clientes
        )

        if not cliente:
            resultado["motivo"] = motivo
            return resultado

        resultado["cliente"] = cliente
        resultado["direcao"] = direcao

        nome_final = f"NF {numero}.xml"

        destino = os.path.join(
            cliente["caminho"],
            "NF",
            direcao,
            str(data.year),
            f"{data.month:02d}-{str(data.year)[2:]}"
        )

        resultado["nome_final"] = nome_final
        resultado["destino"] = destino
        resultado["status"] = "OK"
        resultado["motivo"] = "NF-e identificada"

        return resultado

    # ========================================================
    # NFS-e
    # ========================================================

    if tipo == "NFS-e":

        numero = extrair_numero_nfse(root)
        resultado["numero"] = numero

        if not numero:
            resultado["motivo"] = "NFS-e identificada, mas número não localizado"
            return resultado

        if not data:
            resultado["motivo"] = "NFS-e sem data de emissão identificável"
            return resultado

        cliente, direcao, motivo = identificar_cliente(
            cnpj_emitente,
            cnpj_destinatario,
            clientes
        )

        if not cliente:
            resultado["motivo"] = motivo
            return resultado

        resultado["cliente"] = cliente
        resultado["direcao"] = direcao

        nome_final = f"NFS {numero}.xml"

        destino = os.path.join(
            cliente["caminho"],
            "NF",
            direcao,
            str(data.year),
            f"{data.month:02d}-{str(data.year)[2:]}"
        )

        resultado["nome_final"] = nome_final
        resultado["destino"] = destino
        resultado["status"] = "OK"
        resultado["motivo"] = "NFS-e identificada"

        return resultado

    # ========================================================
    # DESCONHECIDO
    # ========================================================

    resultado["motivo"] = (
        "XML não identificado como NF-e ou NFS-e"
    )

    return resultado


# ============================================================
# PDF
# ============================================================

def extrair_texto_pdf(caminho_pdf):
    """
    Extrai texto do PDF usando pypdf.
    """

    try:
        from pypdf import PdfReader
    except ImportError:
        return None, "Biblioteca pypdf não instalada"

    try:
        reader = PdfReader(caminho_pdf)

        textos = []

        for pagina in reader.pages:
            texto = pagina.extract_text()

            if texto:
                textos.append(texto)

        return "\n".join(textos), None

    except Exception as erro:
        return None, f"Erro ao ler PDF: {erro}"


def extrair_numero_pdf(texto):
    """
    Tenta localizar número da nota em PDF.
    """

    if not texto:
        return None

    padroes = [
        r"N[úu]mero\s*(?:da\s*)?(?:NF|Nota)?\s*[:\-]?\s*(\d+)",
        r"N[úu]mero\s*[:\-]\s*(\d+)",
        r"N[º°o]\s*[:\-]?\s*(\d+)",
        r"NF(?:-?e)?\s*(?:n[º°o]?|n[uú]mero)?\s*[:\-]?\s*(\d+)",
    ]

    for padrao in padroes:
        encontrados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        if encontrados:
            return encontrados[0]

    return None


def extrair_chave_pdf(texto):
    """Procura chave de acesso de 44 dígitos no PDF."""

    if not texto:
        return None

    numeros = re.sub(r"\D", "", texto)

    encontrados = re.findall(r"\d{44}", numeros)

    if encontrados:
        return encontrados[0]

    return None


def processar_pdf(caminho_pdf, clientes):
    """
    Processa PDF de forma conservadora.

    O PDF será identificado, quando possível, pela chave
    de acesso ou pelos dados textuais.
    """

    resultado = {
        "arquivo": os.path.basename(caminho_pdf),
        "status": "REVISÃO",
        "tipo": "PDF",
        "numero": None,
        "data": None,
        "cliente": None,
        "direcao": None,
        "chave": None,
        "nome_final": None,
        "destino": None,
        "motivo": None,
    }

    texto, erro = extrair_texto_pdf(caminho_pdf)

    if erro:
        resultado["motivo"] = erro
        return resultado

    if not texto:
        resultado["motivo"] = "PDF sem texto extraível"
        return resultado

    chave = extrair_chave_pdf(texto)
    resultado["chave"] = chave

    numero = extrair_numero_pdf(texto)
    resultado["numero"] = numero

    # --------------------------------------------------------
    # Tenta localizar CNPJs no PDF
    # --------------------------------------------------------

    cnpjs = re.findall(
        r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}",
        texto
    )

    cnpjs_normalizados = []

    for cnpj in cnpjs:
        valor = limpar_cnpj(cnpj)

        if valor and valor not in cnpjs_normalizados:
            cnpjs_normalizados.append(valor)

    candidatos = []

    for cnpj in cnpjs_normalizados:
        if cnpj in clientes:
            for cliente in clientes[cnpj]:
                candidatos.append(cliente)

    # Se não encontrou cliente pelo CNPJ
    if not candidatos:
        resultado["motivo"] = (
            "PDF identificado, mas cliente não localizado pelo CNPJ"
        )
        return resultado

    # Remove duplicados
    clientes_unicos = []

    for cliente in candidatos:
        if cliente not in clientes_unicos:
            clientes_unicos.append(cliente)

    if len(clientes_unicos) != 1:
        resultado["motivo"] = (
            "PDF corresponde a mais de um cliente"
        )
        return resultado

    cliente = clientes_unicos[0]
    resultado["cliente"] = cliente

    # --------------------------------------------------------
    # PDF sem número
    # --------------------------------------------------------

    if not numero:
        resultado["motivo"] = (
            "PDF identificado pelo cliente, mas número da nota não localizado"
        )
        return resultado

    # --------------------------------------------------------
    # PDFs com chave de acesso
    # --------------------------------------------------------

    if chave:
        # Modelo 55
        if len(chave) == 44:

            # Se houver identificação clara de NF-e
            resultado["tipo"] = "NF-e"

            nome_final = f"NF {numero}.pdf"

            # A direção ainda precisa ser determinada.
            # Para não correr risco, procura CNPJ do cliente no texto
            # e tenta diferenciar emitente/destinatário.

            resultado["motivo"] = (
                "PDF NF-e identificado; direção precisa ser confirmada "
                "pela associação com XML ou dados do documento"
            )

            # Não marca OK sozinho.
            return resultado

    # --------------------------------------------------------
    # PDF sem chave NF-e
    # --------------------------------------------------------

    resultado["motivo"] = (
        "PDF identificado, mas sem chave de acesso NF-e suficiente "
        "para determinar com segurança o tipo/direção"
    )

    return resultado


# ============================================================
# RELATÓRIO
# ============================================================

def escrever_linha(arquivo, texto=""):
    arquivo.write(texto + "\n")


def gerar_relatorio(resultados_xml, resultados_pdf):
    total_xml = len(resultados_xml)
    total_pdf = len(resultados_pdf)
    total = total_xml + total_pdf

    ok_xml = sum(
        1 for resultado in resultados_xml
        if resultado["status"] == "OK"
    )

    ok_pdf = sum(
        1 for resultado in resultados_pdf
        if resultado["status"] == "OK"
    )

    revisao_xml = total_xml - ok_xml
    revisao_pdf = total_pdf - ok_pdf

    with open(
        ARQUIVO_RELATORIO,
        "w",
        encoding="utf-8"
    ) as arquivo:

        escrever_linha(arquivo, "=" * 80)
        escrever_linha(arquivo, "RELATÓRIO DE SIMULAÇÃO - V5")
        escrever_linha(arquivo, "=" * 80)
        escrever_linha(arquivo)

        escrever_linha(
            arquivo,
            "ATENÇÃO: NENHUM ARQUIVO FOI MOVIDO OU RENOMEADO."
        )

        escrever_linha(arquivo)

        escrever_linha(arquivo, f"Total de arquivos: {total}")
        escrever_linha(arquivo, f"XML: {total_xml}")
        escrever_linha(arquivo, f"PDF: {total_pdf}")
        escrever_linha(arquivo)

        escrever_linha(
            arquivo,
            f"XML OK: {ok_xml}"
        )

        escrever_linha(
            arquivo,
            f"XML REVISÃO: {revisao_xml}"
        )

        escrever_linha(
            arquivo,
            f"PDF OK: {ok_pdf}"
        )

        escrever_linha(
            arquivo,
            f"PDF REVISÃO: {revisao_pdf}"
        )

        escrever_linha(arquivo)
        escrever_linha(arquivo, "-" * 80)
        escrever_linha(arquivo, "XML")
        escrever_linha(arquivo, "-" * 80)

        for resultado in resultados_xml:

            escrever_linha(
                arquivo,
                f"Arquivo: {resultado['arquivo']}"
            )

            escrever_linha(
                arquivo,
                f"Status: {resultado['status']}"
            )

            escrever_linha(
                arquivo,
                f"Tipo: {resultado['tipo'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Número: {resultado['numero'] or '?'}"
            )

            if resultado["data"]:
                escrever_linha(
                    arquivo,
                    f"Data: {resultado['data'].strftime('%d/%m/%Y')}"
                )
            else:
                escrever_linha(
                    arquivo,
                    "Data: ?"
                )

            escrever_linha(
                arquivo,
                f"Cliente: "
                f"{resultado['cliente']['cod'] if resultado['cliente'] else '?'}"
            )

            escrever_linha(
                arquivo,
                f"Razão Social: "
                f"{resultado['cliente']['razao'] if resultado['cliente'] else '?'}"
            )

            escrever_linha(
                arquivo,
                f"Direção: {resultado['direcao'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Nome final: {resultado['nome_final'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Destino: {resultado['destino'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Motivo: {resultado['motivo'] or '?'}"
            )

            escrever_linha(arquivo)

        escrever_linha(arquivo, "-" * 80)
        escrever_linha(arquivo, "PDF")
        escrever_linha(arquivo, "-" * 80)

        for resultado in resultados_pdf:

            escrever_linha(
                arquivo,
                f"Arquivo: {resultado['arquivo']}"
            )

            escrever_linha(
                arquivo,
                f"Status: {resultado['status']}"
            )

            escrever_linha(
                arquivo,
                f"Tipo: {resultado['tipo'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Número: {resultado['numero'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Cliente: "
                f"{resultado['cliente']['cod'] if resultado['cliente'] else '?'}"
            )

            escrever_linha(
                arquivo,
                f"Razão Social: "
                f"{resultado['cliente']['razao'] if resultado['cliente'] else '?'}"
            )

            escrever_linha(
                arquivo,
                f"Nome final: {resultado['nome_final'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Destino: {resultado['destino'] or '?'}"
            )

            escrever_linha(
                arquivo,
                f"Motivo: {resultado['motivo'] or '?'}"
            )

            escrever_linha(arquivo)

        escrever_linha(arquivo, "=" * 80)
        escrever_linha(
            arquivo,
            "FIM DA SIMULAÇÃO - NENHUM ARQUIVO FOI ALTERADO."
        )
        escrever_linha(arquivo, "=" * 80)


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print("PROCESSAMENTO XML + PDF NF - V5")
    print("=" * 70)
    print()
    print("MODO SIMULAÇÃO")
    print("Nenhum arquivo será movido ou renomeado.")
    print()

    if not os.path.exists(PASTA_PROCESSAR):
        print(
            f"ERRO: pasta '{PASTA_PROCESSAR}' não encontrada."
        )
        return

    if not os.path.exists(ARQUIVO_EXCEL):
        print(
            f"ERRO: arquivo '{ARQUIVO_EXCEL}' não encontrado."
        )
        return

    try:
        clientes = carregar_clientes()
    except Exception as erro:
        print(f"ERRO ao carregar Excel: {erro}")
        return

    arquivos = os.listdir(PASTA_PROCESSAR)

    arquivos_xml = [
        arquivo
        for arquivo in arquivos
        if arquivo.lower().endswith(".xml")
    ]

    arquivos_pdf = [
        arquivo
        for arquivo in arquivos
        if arquivo.lower().endswith(".pdf")
    ]

    print(f"XML encontrados: {len(arquivos_xml)}")
    print(f"PDF encontrados: {len(arquivos_pdf)}")
    print()

    resultados_xml = []
    resultados_pdf = []

    # --------------------------------------------------------
    # XML
    # --------------------------------------------------------

    for arquivo in sorted(arquivos_xml):

        caminho = os.path.join(
            PASTA_PROCESSAR,
            arquivo
        )

        print(f"Processando XML: {arquivo}")

        resultado = processar_xml(
            caminho,
            clientes
        )

        resultados_xml.append(resultado)

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    for arquivo in sorted(arquivos_pdf):

        caminho = os.path.join(
            PASTA_PROCESSAR,
            arquivo
        )

        print(f"Processando PDF: {arquivo}")

        resultado = processar_pdf(
            caminho,
            clientes
        )

        resultados_pdf.append(resultado)

    # --------------------------------------------------------
    # RELATÓRIO
    # --------------------------------------------------------

    gerar_relatorio(
        resultados_xml,
        resultados_pdf
    )

    print()
    print("=" * 70)
    print("SIMULAÇÃO CONCLUÍDA")
    print("=" * 70)
    print()
    print(
        f"Relatório gerado: {ARQUIVO_RELATORIO}"
    )
    print()
    print("NENHUM ARQUIVO FOI MOVIDO OU RENOMEADO.")


if __name__ == "__main__":
    main()
