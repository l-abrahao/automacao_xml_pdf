import os
import re
import shutil
import xml.etree.ElementTree as ET
from datetime import datetime
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("ERRO: openpyxl não instalado.")
    print("Execute: python -m pip install openpyxl")
    input("\nPressione ENTER para sair...")
    raise SystemExit

try:
    from pypdf import PdfReader
except ImportError:
    print("ERRO: pypdf não instalado.")
    print("Execute: python -m pip install pypdf")
    input("\nPressione ENTER para sair...")
    raise SystemExit


# ============================================================
# CONFIGURAÇÃO
# ============================================================

MODO_EXECUCAO = False   # False = SIMULAÇÃO | True = MOVIMENTA OS ARQUIVOS

BASE = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(BASE, "banco.xlsx")
PASTA_PROCESSAR = os.path.join(BASE, "A_PROCESSAR")
ARQUIVO_RELATORIO = os.path.join(BASE, "relatorio_xml_pdf_v2.txt")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_cnpj(valor):
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def normalizar_numero_nf(valor):
    if not valor:
        return ""
    numeros = re.sub(r"\D", "", str(valor))
    if not numeros:
        return ""
    return str(int(numeros))


def extrair_data_xml(elemento):
    valor = elemento.attrib.get("dhEmi") or elemento.attrib.get("dEmi") or ""
    if not valor:
        return None

    # ISO: 2026-09-15T10:30:00-03:00
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", valor)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    # Caso venha somente como YYYYMMDD
    m = re.search(r"(\d{4})(\d{2})(\d{2})", valor)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    return None


def encontrar_inf_nfe(root):
    for elem in root.iter():
        if elem.tag.split("}")[-1] == "infNFe":
            return elem
    return None


def texto_tag(parent, nome):
    for elem in parent.iter():
        if elem.tag.split("}")[-1] == nome:
            return (elem.text or "").strip()
    return ""


def extrair_chave_de_texto(texto):
    """
    Procura uma chave NF-e de 44 dígitos mesmo quando o PDF
    quebra a chave em espaços, pontos, hífens ou linhas.
    """
    if not texto:
        return None

    # Primeiro tenta uma sequência contínua de 44 dígitos.
    m = re.search(r"(?<!\d)\d{44}(?!\d)", texto)
    if m:
        return m.group(0)

    # Depois permite separadores entre os dígitos.
    # Ex.: 35 2609 12... ou 35.2609.12...
    padrao = r"(?<!\d)(\d(?:[\s.\-]*)\d){43}(?!\d)"
    candidatos = re.findall(padrao, texto)

    # O findall acima pode retornar apenas a última captura dependendo
    # da implementação; por isso fazemos uma segunda abordagem:
    for inicio in range(len(texto)):
        if not texto[inicio].isdigit():
            continue

        trecho = texto[inicio:inicio + 180]
        somente = re.sub(r"[\s.\-]", "", trecho)

        if len(somente) >= 44 and somente[:44].isdigit():
            # Exige que a região original tenha quantidade plausível de dígitos.
            chave = somente[:44]
            if len(chave) == 44:
                return chave

    # Última tentativa: remove tudo que não é dígito e procura blocos.
    # Evitamos usar isso como primeira opção para não juntar números
    # diferentes do documento.
    compacto = re.sub(r"\D", "", texto)
    m = re.search(r"\d{44}", compacto)
    if m:
        return m.group(0)

    return None


def validar_chave_nfe(chave):
    if not chave or len(chave) != 44 or not chave.isdigit():
        return False

    # Dígito verificador da chave NF-e
    base = chave[:43]
    dv_informado = int(chave[43])

    soma = 0
    peso = 2

    for digito in reversed(base):
        soma += int(digito) * peso
        peso += 1
        if peso > 9:
            peso = 2

    resto = soma % 11
    dv_calculado = 0 if resto in (0, 1) else 11 - resto

    return dv_calculado == dv_informado


def dados_da_chave(chave):
    if not chave or len(chave) != 44:
        return {}

    return {
        "uf": chave[0:2],
        "aamm": chave[2:6],
        "cnpj_emitente": chave[6:20],
        "modelo": chave[20:22],
        "serie": chave[22:25],
        "nf": normalizar_numero_nf(chave[25:34]),
        "tp_emis": chave[34],
        "codigo": chave[35:43],
        "dv": chave[43],
    }


def extrair_cnpjs_pdf(texto):
    if not texto:
        return []

    encontrados = []

    # Formato tradicional 00.000.000/0000-00
    for m in re.findall(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", texto):
        cnpj = normalizar_cnpj(m)
        if len(cnpj) == 14 and cnpj not in encontrados:
            encontrados.append(cnpj)

    # CNPJ pode aparecer sem pontuação
    for m in re.findall(r"(?<!\d)\d{14}(?!\d)", texto):
        if m not in encontrados:
            encontrados.append(m)

    return encontrados


def extrair_cnpj_proximo(texto, termos):
    """
    Procura CNPJ próximo de um rótulo como DESTINATÁRIO/EMITENTE.
    """
    texto_upper = texto.upper()

    for termo in termos:
        pos = texto_upper.find(termo)
        if pos < 0:
            continue

        trecho = texto[pos:pos + 1500]
        cnpjs = extrair_cnpjs_pdf(trecho)

        if cnpjs:
            return cnpjs[0]

    return ""


def extrair_data_pdf(texto):
    if not texto:
        return None

    # Formatos comuns: 15/09/2026
    padroes = [
        r"\b(\d{2})/(\d{2})/(\d{4})\b",
        r"\b(\d{2})-(\d{2})-(\d{4})\b",
    ]

    for padrao in padroes:
        for d, m, a in re.findall(padrao, texto):
            try:
                dt = datetime(int(a), int(m), int(d))
                # Evita datas claramente improváveis.
                if 2000 <= dt.year <= 2100:
                    return dt
            except ValueError:
                pass

    return None


def ler_excel():
    if not os.path.exists(ARQUIVO_EXCEL):
        raise FileNotFoundError(
            f"Arquivo não encontrado: {ARQUIVO_EXCEL}"
        )

    wb = openpyxl.load_workbook(ARQUIVO_EXCEL, data_only=True)
    ws = wb.active

    cabecalhos = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(1, col).value
        if valor:
            cabecalhos[str(valor).strip().upper()] = col

    obrigatorios = ["CNPJ", "CAMINHO ABSOLUTO"]
    for campo in obrigatorios:
        if campo not in cabecalhos:
            raise ValueError(
                f"A coluna obrigatória '{campo}' não foi encontrada no banco.xlsx."
            )

    clientes = defaultdict(list)

    for linha in range(2, ws.max_row + 1):
        cnpj = normalizar_cnpj(
            ws.cell(linha, cabecalhos["CNPJ"]).value
        )
        if not cnpj:
            continue

        caminho = ws.cell(
            linha, cabecalhos["CAMINHO ABSOLUTO"]
        ).value

        if not caminho:
            continue

        codigo = ""
        razao = ""

        if "COD" in cabecalhos:
            codigo = ws.cell(linha, cabecalhos["COD"]).value

        if "RAZÃO SOCIAL" in cabecalhos:
            razao = ws.cell(linha, cabecalhos["RAZÃO SOCIAL"]).value

        clientes[cnpj].append({
            "linha": linha,
            "cnpj": cnpj,
            "codigo": codigo,
            "razao": str(razao or ""),
            "caminho": str(caminho).strip(),
        })

    return clientes


def identificar_cliente(cnpj, clientes):
    cnpj = normalizar_cnpj(cnpj)

    if not cnpj:
        return None, "CNPJ não informado"

    encontrados = clientes.get(cnpj, [])

    if len(encontrados) == 1:
        return encontrados[0], ""

    if len(encontrados) > 1:
        nomes = " | ".join(
            f"{x['codigo']} - {x['razao']}" for x in encontrados
        )
        return None, f"CNPJ duplicado no banco: {nomes}"

    return None, "CNPJ não encontrado no banco"


def determinar_movimento(cliente_cnpj, cnpj_emitente, cnpj_destinatario):
    cliente_cnpj = normalizar_cnpj(cliente_cnpj)
    cnpj_emitente = normalizar_cnpj(cnpj_emitente)
    cnpj_destinatario = normalizar_cnpj(cnpj_destinatario)

    if cliente_cnpj and cliente_cnpj == cnpj_emitente:
        return "SAIDAS"

    if cliente_cnpj and cliente_cnpj == cnpj_destinatario:
        return "ENTRADAS"

    return ""


def pasta_destino(cliente, movimento, data_nf):
    ano = str(data_nf.year)
    mes = f"{data_nf.month:02d}-{str(data_nf.year)[-2:]}"

    return os.path.join(
        cliente["caminho"],
        "NF",
        movimento,
        ano,
        mes
    )


def nome_codigo_cliente(cliente):
    codigo = cliente.get("codigo", "")

    if isinstance(codigo, float) and codigo.is_integer():
        codigo = int(codigo)

    return f"{codigo} - {cliente['razao']}"


def analisar_xml(caminho, clientes):
    resultado = {
        "arquivo": os.path.basename(caminho),
        "tipo": "XML",
        "status": "",
        "motivo": "",
        "cliente": None,
        "movimento": "",
        "numero_nf": "",
        "data": None,
        "chave": "",
        "cnpj_emitente": "",
        "cnpj_destinatario": "",
        "origem": caminho,
    }

    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
    except Exception as e:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = f"Erro ao ler XML: {e}"
        return resultado

    inf = encontrar_inf_nfe(root)

    if inf is None:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "XML sem infNFe; provavelmente não é um XML NF-e padrão."
        return resultado

    chave = inf.attrib.get("Id", "")
    chave = re.sub(r"\D", "", chave)

    if len(chave) == 44:
        resultado["chave"] = chave

    emit = None
    dest = None

    for elem in inf:
        nome = elem.tag.split("}")[-1]
        if nome == "emit":
            emit = elem
        elif nome == "dest":
            dest = elem

    if emit is not None:
        resultado["cnpj_emitente"] = normalizar_cnpj(texto_tag(emit, "CNPJ"))

    if dest is not None:
        resultado["cnpj_destinatario"] = normalizar_cnpj(texto_tag(dest, "CNPJ"))

    resultado["numero_nf"] = normalizar_numero_nf(texto_tag(inf, "nNF"))
    resultado["data"] = extrair_data_xml(inf)

    # Se o XML não tiver CNPJ do emitente, tentamos usar a chave.
    if not resultado["cnpj_emitente"] and len(resultado["chave"]) == 44:
        resultado["cnpj_emitente"] = resultado["chave"][6:20]

    cliente, motivo = identificar_cliente(
        resultado["cnpj_emitente"], clientes
    )

    cliente_dest, motivo_dest = identificar_cliente(
        resultado["cnpj_destinatario"], clientes
    )

    # Prioridade: cliente emitente; depois cliente destinatário.
    if cliente:
        resultado["cliente"] = cliente
        resultado["movimento"] = "SAIDAS"

    elif cliente_dest:
        resultado["cliente"] = cliente_dest
        resultado["movimento"] = "ENTRADAS"

    else:
        resultado["motivo"] = (
            f"Emitente: {motivo}; Destinatário: {motivo_dest}"
        )
        resultado["status"] = "REVISÃO"
        return resultado

    if not resultado["data"]:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "Data de emissão não identificada."
        return resultado

    if not resultado["numero_nf"] and len(resultado["chave"]) == 44:
        resultado["numero_nf"] = normalizar_numero_nf(
            resultado["chave"][25:34]
        )

    if not resultado["numero_nf"]:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "Número da NF não identificado."
        return resultado

    resultado["status"] = "OK"
    return resultado


def extrair_texto_pdf(caminho):
    leitor = PdfReader(caminho)

    partes = []
    for pagina in leitor.pages:
        try:
            partes.append(pagina.extract_text() or "")
        except Exception:
            partes.append("")

    return "\n".join(partes)


def analisar_pdf(caminho, clientes, xml_por_chave):
    resultado = {
        "arquivo": os.path.basename(caminho),
        "tipo": "PDF",
        "status": "",
        "motivo": "",
        "cliente": None,
        "movimento": "",
        "numero_nf": "",
        "data": None,
        "chave": "",
        "cnpj_emitente": "",
        "cnpj_destinatario": "",
        "origem": caminho,
        "xml_par": None,
    }

    try:
        texto = extrair_texto_pdf(caminho)
    except Exception as e:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = f"Erro ao ler PDF: {e}"
        return resultado

    if not texto.strip():
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "PDF sem texto extraível."
        return resultado

    chave = extrair_chave_de_texto(texto)

    # Também tenta pelo nome do arquivo, somente como fallback de pareamento.
    if not chave:
        nome = os.path.basename(caminho)
        m = re.search(r"(?<!\d)\d{44}(?!\d)", nome)
        if m:
            chave = m.group(0)

    if chave and len(chave) == 44:
        resultado["chave"] = chave

    # --------------------------------------------------------
    # 1) Melhor cenário: PDF encontrado pelo XML correspondente
    # --------------------------------------------------------
    xml = xml_por_chave.get(chave) if chave else None

    if xml:
        resultado["xml_par"] = xml

        resultado["cliente"] = xml["cliente"]
        resultado["movimento"] = xml["movimento"]
        resultado["numero_nf"] = xml["numero_nf"]
        resultado["data"] = xml["data"]
        resultado["cnpj_emitente"] = xml["cnpj_emitente"]
        resultado["cnpj_destinatario"] = xml["cnpj_destinatario"]

        resultado["status"] = "PAR XML/PDF"
        resultado["motivo"] = "PDF vinculado ao XML pela chave de acesso."
        return resultado

    # --------------------------------------------------------
    # 2) PDF sem XML: tenta montar a identificação internamente
    # --------------------------------------------------------
    dados_chave = dados_da_chave(chave)

    cnpj_emitente = ""
    cnpj_destinatario = ""

    if dados_chave:
        cnpj_emitente = dados_chave.get("cnpj_emitente", "")

    # Procura explicitamente próximo de DESTINATÁRIO.
    cnpj_destinatario = extrair_cnpj_proximo(
        texto,
        [
            "DESTINATÁRIO",
            "DESTINATARIO",
            "DESTINATÁRIO/REMETENTE",
            "DESTINATARIO/REMETENTE",
        ]
    )

    # Se não achou pelo rótulo, coleta CNPJs do PDF.
    cnpjs = extrair_cnpjs_pdf(texto)

    resultado["cnpj_emitente"] = cnpj_emitente
    resultado["cnpj_destinatario"] = cnpj_destinatario

    cliente_emit, motivo_emit = identificar_cliente(
        cnpj_emitente, clientes
    )

    cliente_dest, motivo_dest = identificar_cliente(
        cnpj_destinatario, clientes
    )

    # Se o destinatário não foi identificado pelo rótulo,
    # verifica CNPJs encontrados no documento.
    if not cliente_dest and cnpjs:
        for cnpj in cnpjs:
            if cnpj == cnpj_emitente:
                continue

            candidato, _ = identificar_cliente(cnpj, clientes)
            if candidato:
                cliente_dest = candidato
                cnpj_destinatario = cnpj
                resultado["cnpj_destinatario"] = cnpj
                break

    movimento = ""

    if cliente_emit and cliente_dest:
        # Se o cliente aparece nos dois lados, não decide sozinho.
        if normalizar_cnpj(cliente_emit["cnpj"]) == normalizar_cnpj(cliente_dest["cnpj"]):
            resultado["status"] = "REVISÃO"
            resultado["motivo"] = "Mesmo cliente identificado como emitente e destinatário."
            return resultado

    if cliente_emit:
        movimento = "SAIDAS"
        resultado["cliente"] = cliente_emit

    elif cliente_dest:
        movimento = "ENTRADAS"
        resultado["cliente"] = cliente_dest

    else:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = (
            f"Cliente não identificado com segurança. "
            f"Emitente: {motivo_emit}; Destinatário: {motivo_dest}"
        )
        return resultado

    resultado["movimento"] = movimento

    if dados_chave:
        resultado["numero_nf"] = dados_chave.get("nf", "")

    if not resultado["numero_nf"]:
        # Não confiar no primeiro número encontrado no PDF.
        # Só usamos padrões próximos de 'Número'.
        padroes_nf = [
            r"N[ÚU]MERO\s*(?:DA\s*NF(?:-E)?|NF(?:-E)?)?\s*[:\-]?\s*(\d{1,9})",
            r"N[ÚU]MERO\s*[:\-]?\s*(\d{1,9})",
        ]

        texto_upper = texto.upper()

        for padrao in padroes_nf:
            m = re.search(padrao, texto_upper)
            if m:
                resultado["numero_nf"] = normalizar_numero_nf(m.group(1))
                break

    resultado["data"] = extrair_data_pdf(texto)

    if not resultado["data"] and dados_chave:
        aamm = dados_chave.get("aamm", "")
        if len(aamm) == 4:
            # A chave só informa ano/mês, não o dia.
            resultado["motivo"] = (
                "PDF permite identificar cliente/NF, mas não foi possível "
                "confirmar o dia da emissão."
            )
            resultado["status"] = "REVISÃO"
            return resultado

    if not resultado["numero_nf"]:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "Número da NF não identificado com segurança."
        return resultado

    if not resultado["data"]:
        resultado["status"] = "REVISÃO"
        resultado["motivo"] = "Data da NF não identificada com segurança."
        return resultado

    resultado["status"] = "OK"
    resultado["motivo"] = "PDF analisado sem XML correspondente."
    return resultado


def localizar_pares_por_chave(resultados):
    """
    Apenas cria informação de vínculo. Não movimenta nada.
    """
    xml_por_chave = {}

    for r in resultados:
        if r["tipo"] == "XML" and r["chave"]:
            if r["chave"] not in xml_por_chave:
                xml_por_chave[r["chave"]] = r

    return xml_por_chave


def destino_nome(resultado, extensao):
    cliente = resultado["cliente"]
    movimento = resultado["movimento"]
    data_nf = resultado["data"]
    numero_nf = resultado["numero_nf"]

    pasta = pasta_destino(cliente, movimento, data_nf)

    nome = f"NF {numero_nf}{extensao}"

    return pasta, nome


def processar():
    print("=" * 70)
    print("ORGANIZADOR NF-e XML + PDF - V2")
    print("=" * 70)

    if not os.path.exists(PASTA_PROCESSAR):
        print(f"\nERRO: pasta não encontrada:\n{PASTA_PROCESSAR}")
        input("\nPressione ENTER para sair...")
        return

    try:
        clientes = ler_excel()
    except Exception as e:
        print(f"\nERRO ao ler banco.xlsx: {e}")
        input("\nPressione ENTER para sair...")
        return

    arquivos = []
    for nome in os.listdir(PASTA_PROCESSAR):
        caminho = os.path.join(PASTA_PROCESSAR, nome)

        if not os.path.isfile(caminho):
            continue

        extensao = os.path.splitext(nome)[1].lower()

        if extensao in [".xml", ".pdf"]:
            arquivos.append(caminho)

    arquivos.sort()

    print(f"\nClientes carregados: {sum(len(v) for v in clientes.values())}")
    print(f"Arquivos encontrados: {len(arquivos)}")

    # Primeiro XMLs.
    resultados_xml = []
    for caminho in arquivos:
        if caminho.lower().endswith(".xml"):
            resultados_xml.append(
                analisar_xml(caminho, clientes)
            )

    xml_por_chave = localizar_pares_por_chave(resultados_xml)

    # Depois PDFs, já tendo os XMLs disponíveis para pareamento.
    resultados_pdf = []
    for caminho in arquivos:
        if caminho.lower().endswith(".pdf"):
            resultados_pdf.append(
                analisar_pdf(
                    caminho,
                    clientes,
                    xml_por_chave
                )
            )

    resultados = resultados_xml + resultados_pdf

    # ========================================================
    # RELATÓRIO
    # ========================================================

    linhas = []
    linhas.append("=" * 90)
    linhas.append("RELATÓRIO - ORGANIZADOR NF-e XML + PDF V2")
    linhas.append("=" * 90)
    linhas.append(
        f"Data/hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    )
    linhas.append(
        f"Modo: {'EXECUÇÃO' if MODO_EXECUCAO else 'SIMULAÇÃO'}"
    )
    linhas.append(f"Total de arquivos: {len(arquivos)}")
    linhas.append(
        f"XML: {sum(1 for x in arquivos if x.lower().endswith('.xml'))}"
    )
    linhas.append(
        f"PDF: {sum(1 for x in arquivos if x.lower().endswith('.pdf'))}"
    )
    linhas.append("")

    contagem = defaultdict(int)
    for r in resultados:
        contagem[r["status"]] += 1

    linhas.append("RESUMO:")
    for status, quantidade in sorted(contagem.items()):
        linhas.append(f"- {status}: {quantidade}")

    linhas.append("")
    linhas.append("-" * 90)
    linhas.append("DETALHAMENTO")
    linhas.append("-" * 90)

    for r in resultados:
        linhas.append("")
        linhas.append(f"Arquivo: {r['arquivo']}")
        linhas.append(f"Tipo: {r['tipo']}")
        linhas.append(f"Status: {r['status']}")

        if r["cliente"]:
            linhas.append(
                f"Cliente: {nome_codigo_cliente(r['cliente'])}"
            )
            linhas.append(
                f"CNPJ cliente: {r['cliente']['cnpj']}"
            )

        linhas.append(
            f"Movimento: {r['movimento'] or '?'}"
        )
        linhas.append(
            f"NF: {r['numero_nf'] or '?'}"
        )

        if r["data"]:
            linhas.append(
                f"Data: {r['data'].strftime('%d/%m/%Y')}"
            )
        else:
            linhas.append("Data: ?")

        linhas.append(
            f"CNPJ emitente: {r['cnpj_emitente'] or '?'}"
        )
        linhas.append(
            f"CNPJ destinatário: {r['cnpj_destinatario'] or '?'}"
        )
        linhas.append(
            f"Chave: {r['chave'] or '?'}"
        )

        if r.get("xml_par"):
            linhas.append(
                f"XML correspondente: {r['xml_par']['arquivo']}"
            )

        if r["motivo"]:
            linhas.append(f"Observação: {r['motivo']}")

        if r["status"] in ("OK", "PAR XML/PDF"):
            # Só calcula o destino quando cliente, movimento, número e data
            # estiverem completos. Nunca deixa um caso incompleto derrubar
            # o relatório inteiro.
            if r.get("cliente") and r.get("movimento") and r.get("data") and r.get("numero_nf"):
                pasta, nome = destino_nome(
                    r,
                    os.path.splitext(r["arquivo"])[1].lower()
                )

                linhas.append(
                    f"Destino proposto: {os.path.join(pasta, nome)}"
                )
            else:
                linhas.append(
                    "Destino proposto: REVISÃO - dados insuficientes para montar a pasta."
                )
                if r["status"] == "OK":
                    r["status"] = "REVISÃO"
                    r["motivo"] = (
                        "Dados insuficientes para montar o destino "
                        "(cliente, movimento, data ou número da NF)."
                    )

    # ========================================================
    # EXECUÇÃO
    # ========================================================

    if MODO_EXECUCAO:
        linhas.append("")
        linhas.append("-" * 90)
        linhas.append("MOVIMENTAÇÕES EXECUTADAS")
        linhas.append("-" * 90)

        # Cria diretórios e move somente itens aprovados.
        # Para PDF/ XML pareados, ambos são processados individualmente.
        for r in resultados:
            if r["status"] not in ("OK", "PAR XML/PDF"):
                continue

            try:
                # Proteção adicional: jamais mover um arquivo sem destino completo.
                if not (
                    r.get("cliente")
                    and r.get("movimento")
                    and r.get("data")
                    and r.get("numero_nf")
                ):
                    linhas.append(
                        f"REVISÃO - não movido: {r['arquivo']} - dados insuficientes."
                    )
                    continue

                pasta, nome = destino_nome(
                    r,
                    os.path.splitext(r["arquivo"])[1].lower()
                )

                os.makedirs(pasta, exist_ok=True)

                destino = os.path.join(pasta, nome)

                # Evita sobrescrever silenciosamente.
                if os.path.exists(destino):
                    linhas.append(
                        f"JÁ EXISTE - não movido: {r['arquivo']} -> {destino}"
                    )
                    continue

                shutil.move(r["origem"], destino)

                linhas.append(
                    f"MOVIDO: {r['arquivo']} -> {destino}"
                )

            except Exception as e:
                linhas.append(
                    f"ERRO AO MOVER {r['arquivo']}: {e}"
                )

    else:
        linhas.append("")
        linhas.append("-" * 90)
        linhas.append("SIMULAÇÃO")
        linhas.append("-" * 90)
        linhas.append(
            "Nenhum arquivo foi movido. O relatório mostra apenas o destino proposto."
        )

    with open(
        ARQUIVO_RELATORIO,
        "w",
        encoding="utf-8"
    ) as f:
        f.write("\n".join(linhas))

    print("\n" + "\n".join(linhas))
    print("\n" + "=" * 70)
    print(f"Relatório salvo em:\n{ARQUIVO_RELATORIO}")
    print("=" * 70)

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    processar()
