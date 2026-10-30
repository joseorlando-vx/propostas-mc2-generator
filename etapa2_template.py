"""
============================================================
GERADOR DE PROPOSTAS MC2 - V2 (4 TIPOS DE PROPOSTA)
CORRIGIDO em 2026-07-06: campo "sistema" adicionado aos tipos
A, C e D (antes só existia no Tipo B), pra acompanhar a correção
do bug "Ref.: hardcoded" nos templates. Ver
gerador-propostas-mc2/references/bugs_conhecidos.md.
============================================================

Suporta 4 tipos de proposta:
    A - SIMPLES         (só fornecimento)
    B - INTERMEDIARIA   (fornecimento + serviço — Osteria)
    C - SERVICO         (manutenção, visita, laudo)
    D - ADITIVO         (aditivo contratual com justificativa)

Como usar:
    1. Edita TIPO_PROPOSTA pra escolher A, B, C ou D
    2. Edita o dicionário DADOS_DA_PROPOSTA correspondente
    3. Roda: python etapa2_template.py
    4. Resultado em propostas_geradas/
============================================================
"""

import re

from docxtpl import DocxTemplate
from pathlib import Path


# -------------------------------------------------------------
# CONFIGURAÇÃO PRINCIPAL
# -------------------------------------------------------------
# Escolha o tipo da proposta. Mude esta linha pra A, B, C ou D:
TIPO_PROPOSTA = "B"


# -------------------------------------------------------------
# MAPEAMENTO TIPO -> TEMPLATE
# -------------------------------------------------------------
# Caminhos ancorados na pasta deste arquivo (não no diretório de trabalho),
# pra funcionar tanto pela CLI quanto pela interface Streamlit (que pode
# rodar de outro cwd). Ver Fase 2 / app.py.
BASE_DIR = Path(__file__).resolve().parent
# v5: painel único (2026-07-08) — funde campos de escopo/objetivo por tipo
# (ver bugs_conhecidos.md, "Debate completo do novo formulário/preview
# único") e corrige um bug até então não percebido: os templates v4_arial
# ainda usavam {{ exclusao_1/2/3 }} soltos em vez de {{ bloco_exclusoes }},
# então as exclusões saíam VAZIAS em silêncio desde a Fase 2 (docxtpl não
# reclama de variável indefinida). v4_arial fica intacta como fallback.
PASTA_TEMPLATES = BASE_DIR / "templates_v5_painel_unico"
PASTA_SAIDA = BASE_DIR / "propostas_geradas"

TEMPLATES = {
    "A": PASTA_TEMPLATES / "template_simples.docx",
    "B": PASTA_TEMPLATES / "template_intermediario.docx",
    "C": PASTA_TEMPLATES / "template_servico.docx",
    "D": PASTA_TEMPLATES / "template_aditivo.docx",
}

# Marcadores/campos esperados de cada tipo (para validação dos INPUTS).
# NOTA 2026-07-06: "sistema" adicionado em A, C e D — esses templates
# tinham o campo "Ref.:" do cabeçalho hardcoded (CFTV / SDAI / RACK) em
# vez de usar marcador. Corrigido pra {{ sistema }} nos 4 tipos.
# NOTA 2026-07-06 (Fase 2 — simplificação do form):
#   - "paragrafo_apresentacao" saiu daqui: agora é DERIVADO (montado
#     automaticamente de sistema+cliente+cidade+tipo em montar_apresentacao),
#     não é mais um campo digitado. O template ainda tem {{ paragrafo_apresentacao }};
#     gerar_proposta preenche sozinho.
#   - "exclusao_1..N" viraram uma lista única "exclusoes" (list[str]),
#     renderizada como bullets reais via Subdoc (ver gerar_proposta). O
#     template agora tem um único {{ bloco_exclusoes }} em vez de N marcadores.
#   - "cidade" adicionada (input curto, alimenta a apresentação; default
#     "Brasília/DF"). No Tipo D, "cliente" também virou input (pra apresentação).
# NOTA 2026-07-08 (v5 — painel único, ver bugs_conhecidos.md):
#   - Tipo B: "escopo_geral"+"escopo_fornecimento"+"escopo_servicos"+
#     "mao_de_obra" fundem num único campo "escopo" (texto corrido). Isso
#     REABRE a decisão anterior de manter os 4 blocos separados — José
#     confirmou explicitamente a fusão nesta data.
#   - Tipo C: "introducao"+"objetivo_intro"+"mao_de_obra"+"cronograma_intro"+
#     "cronograma_detalhe" fundem no campo único "introducao". "objetivo_1..7"
#     vira a lista única "objetivos" (renderizada como bullets via Subdoc,
#     mesmo mecanismo de "exclusoes" — ver gerar_proposta / bloco_objetivos).
#   - Tipo D: "escopo_geral"+"justificativa_1/2/3" fundem no campo único
#     "escopo" (o template não tem mais heading separado de JUSTIFICATIVA
#     TÉCNICA).
MARCADORES_POR_TIPO = {
    "A": {
        "numero_proposta", "revisao", "data_extenso",
        "cliente", "cidade", "sistema",
        "escopo_geral",
        "exclusoes",
        "valor_total", "condicoes_pagamento",
    },
    "B": {
        "numero_proposta", "revisao", "data_extenso",
        "cliente", "cidade", "sistema",
        "escopo",
        "exclusoes",
        "valor_total", "condicoes_pagamento",
    },
    "C": {
        "numero_proposta", "revisao", "data_extenso",
        "cliente", "cidade", "sistema",
        "introducao",
        "objetivos",
        "exclusoes",
        "valor_total", "condicoes_pagamento",
    },
    "D": {
        "numero_proposta", "revisao", "data_extenso",
        "titulo_aditivo", "cliente", "cidade", "sistema",
        "escopo",
        "exclusoes",
        "valor_total", "condicoes_pagamento",
    },
}


# -------------------------------------------------------------
# APRESENTAÇÃO AUTOMÁTICA (Fase 2)
# -------------------------------------------------------------
# Sigla -> nome completo do sistema, pra montar a frase de apresentação.
# SDAI e CFTV confirmados no texto das propostas reais (PS26_009/066, PS25_107).
# SCA/Interfonia/Automação Predial NÃO aparecem na apresentação de nenhum
# real; nomes padrão aprovados por José em 2026-07-06 (ver bugs_conhecidos.md).
# Sigla desconhecida passa direto (devolve o que veio).
SISTEMAS = {
    "SDAI": "Sistema de Detecção e Alarme de Incêndio (SDAI)",
    "CFTV": "Sistema de CFTV",
    "SCA": "Sistema de Controle de Acesso (SCA)",
    "INTERFONIA": "Sistema de Interfonia",
    "AUTOMAÇÃO PREDIAL": "Sistema de Automação Predial",
    # "AUT": sigla curta usada no select do painel único (v5) — mesmo nome
    # completo de "AUTOMAÇÃO PREDIAL", só troca o texto que aparece no Ref.:.
    "AUT": "Sistema de Automação Predial",
}


def nome_sistema(sigla: str) -> str:
    """Sigla -> nome completo. Desconhecida volta como veio (upper/strip)."""
    return SISTEMAS.get((sigla or "").strip().upper(), (sigla or "").strip())


# Frase por tipo. A e B são as frases EXATAS dos reais (PS26_066, PS26_009).
# C e D são PROVISÓRIAS — não há exemplo real desses tipos na pasta de
# referência (ver bugs_conhecidos.md: "não confirmado com real").
_FRASE_POR_TIPO = {
    "A": ("Apresentamos nossa proposta para fornecimento de equipamentos "
          "para o {sistema}, para a {cliente}, em {cidade}."),
    "B": ("Apresentamos a presente proposta para o fornecimento de equipamentos "
          "e materiais, bem como para a aplicação de mão de obra especializada, "
          "destinados ao {sistema}, para a {cliente}, localizada em {cidade}."),
    "C": ("Apresentamos nossa proposta para prestação de serviços com aplicação "
          "de mão de obra especializada para o {sistema}, para a {cliente}, "
          "em {cidade}."),
    "D": ("Apresentamos a presente proposta de aditivo contratual para a "
          "aplicação de mão de obra especializada no {sistema}, para a {cliente}, "
          "em {cidade}."),
}


def montar_apresentacao(tipo: str, dados: dict) -> str:
    """Monta o parágrafo de apresentação a partir de sistema+cliente+cidade."""
    sistema = nome_sistema(dados.get("sistema", ""))
    cliente = dados.get("cliente") or dados.get("titulo_aditivo") or "[cliente]"
    cidade = (dados.get("cidade") or "").strip() or "Brasília/DF"
    return _FRASE_POR_TIPO[tipo].format(sistema=sistema, cliente=cliente, cidade=cidade)


# -------------------------------------------------------------
# CONDIÇÕES DE PAGAMENTO (Fase 2) — dropdown no app
# -------------------------------------------------------------
# Rótulo curto -> frase exata. 100% e 50/50 batem com os reais (PS26_002,
# Osteria). "Mensalidade" é redigida aqui (não há frase exata num real);
# marcada como tal em bugs_conhecidos.md.
# "Outro" (v5) NÃO entra neste dict: o app.py mostra um text_input livre
# quando o usuário escolhe essa opção e manda o texto digitado direto pro
# campo "condicoes_pagamento" — não precisa de preset aqui.
PAGAMENTOS = {
    "100% no aceite": "Pagamento – 100% no aceite da proposta.",
    "50% / 50%": "Pagamento – 50% no aceite da proposta e 50% na entrega dos serviços.",
    "Mensalidade": "Pagamento – mensalidade recorrente, faturada mensalmente conforme contrato.",
}


# =============================================================
# DADOS DAS PROPOSTAS — EDITE A SEÇÃO DO TIPO QUE VAI USAR
# =============================================================

# -------------------------------------------------------------
# TIPO A — só fornecimento
# Exemplo: substituição de câmeras, fornecimento de detector
# -------------------------------------------------------------
DADOS_TIPO_A = {
    "numero_proposta": "PS 26_080",
    "revisao": "00",
    "data_extenso": "30 de abril de 2026",
    "cliente": "EDIFICIO EXEMPLO",
    "cidade": "Brasília/DF",
    "sistema": "CFTV",
    "escopo_geral": (
        "O escopo da MC2/AQ3 Engenharia nesta proposta compreende o "
        "fornecimento de equipamentos para substituição de câmeras queimadas "
        "no Edifício Exemplo.\n\n"
        "- 02 (duas) câmeras digitais compatíveis com o sistema existente;\n"
        "- Cabos e conectores necessários para a instalação."
    ),
    "exclusoes": [
        "Serviços Civis tais como reboco, gesso, pintura etc.;",
        "Instalação e fornecimento de equipamentos que não sejam os contemplados no escopo (item 1);",
    ],
    "valor_total": "R$ 360,00 (Trezentos e sessenta reais).",
    "condicoes_pagamento": "Pagamento – 100% no aceite da proposta.",
}


# -------------------------------------------------------------
# TIPO B — fornecimento + serviço (caso Osteria)
# Exemplo: troca de detector com instalação e mão de obra
# -------------------------------------------------------------
DADOS_TIPO_B = {
    "numero_proposta": "PS 26_081",
    "revisao": "00",
    "data_extenso": "30 de abril de 2026",
    "cliente": "EXEMPLO LTDA",
    "cidade": "Brasília/DF",
    "sistema": "SDAI",
    "escopo": (
        "O escopo da MC2/AQ3 Engenharia contempla a substituição de 02 (dois) "
        "detectores de fumaça, contemplando o fornecimento dos equipamentos, "
        "mão de obra especializada e comissionamento.\n\n"
        "- 02 (dois) detectores de fumaça endereçáveis compatíveis com a central existente;\n"
        "- Materiais de instalação necessários.\n\n"
        "Os serviços contemplam a desinstalação dos dispositivos atuais, "
        "instalação e endereçamento dos novos dispositivos, integração à "
        "central existente, testes funcionais e comissionamento do sistema. "
        "O serviço será executado por equipe técnica especializada da MC2, "
        "com a alocação de profissionais qualificados para a execução em campo."
    ),
    "exclusoes": [
        "Serviços Civis tais como reboco, gesso, pintura etc.;",
        "Serviços fora do horário comercial;",
        "Instalação e fornecimento de equipamentos que não sejam os contemplados no escopo (item 1);",
    ],
    "valor_total": "R$ 2.450,00 (Dois mil quatrocentos e cinquenta reais).",
    "condicoes_pagamento": "Pagamento – 50% no aceite da proposta e 50% na entrega dos serviços.",
}


# -------------------------------------------------------------
# TIPO C — serviço puro (manutenção, visita, laudo)
# Exemplo: manutenção corretiva, visita técnica
# -------------------------------------------------------------
DADOS_TIPO_C = {
    "numero_proposta": "PS 26_082",
    "revisao": "00",
    "data_extenso": "30 de abril de 2026",
    "cliente": "SHOPPING EXEMPLO",
    "cidade": "Brasília/DF",
    "sistema": "SDAI",
    "introducao": (
        "A MC2/AQ3 Engenharia apresenta proposta para execução de manutenção "
        "corretiva no sistema de detecção, contemplando análise técnica, "
        "limpeza e reabilitação dos pontos com falhas identificadas.\n\n"
        "Toda a manutenção será executada por equipe técnica especializada, "
        "utilizando ferramentas adequadas e seguindo procedimentos alinhados "
        "às recomendações do fabricante.\n\n"
        "Os serviços serão realizados conforme a seguinte programação: "
        "execução das atividades no período das 08h00 às 18h00, de segunda a "
        "sexta-feira, garantindo maior flexibilidade e agilidade na execução "
        "dos serviços.\n\n"
        "A manutenção corretiva proposta tem por finalidade:"
    ),
    "objetivos": [
        "Limpeza de todos os detectores convencionais;",
        "Análise da situação do sistema após a limpeza, verificando se as não-conformidades são sanadas;",
        "Listar as falhas que permanecerão;",
        "Identificar com precisão os pontos que permanecem com não conformidades;",
        "Apresentar laudo técnico com a situação remanescente após limpeza;",
        "Apresentar proposta para solucionar pendências remanescentes;",
        "Garantir que o sistema opere com o nível mínimo de segurança e conformidade técnica estabelecido pelas normas vigentes, especialmente a NBR 17240.",
    ],
    "exclusoes": [
        "Serviços Civis tais como reboco, gesso, pintura etc.;",
        "Fornecimento de materiais, cabos ou equipamentos de qualquer natureza;",
        "Ampliação do sistema ou instalação de novos dispositivos;",
        "Serviços fora do horário comercial;",
        "Fornecimento de ESCADA.",
    ],
    "valor_total": "R$ 8.500,00 (Oito mil e quinhentos reais).",
    "condicoes_pagamento": "Pagamento – 30% no aceite da proposta, 70% na entrega dos serviços.",
}


# -------------------------------------------------------------
# TIPO D — aditivo contratual
# Exemplo: serviço extra durante obra em andamento
# -------------------------------------------------------------
DADOS_TIPO_D = {
    "numero_proposta": "PO 25_016.AD04",
    "revisao": "00",
    "data_extenso": "30 de abril de 2026",
    "titulo_aditivo": "ADITIVO 04 — DESCRIÇÃO DO SERVIÇO",
    "cliente": "EMBAIXADA DE MALTA",
    "cidade": "Brasília/DF",
    "sistema": "RACK",
    "escopo": (
        "A presente proposta tem por objeto a formalização de aditivo "
        "contratual para a execução de [DESCRIÇÃO DO SERVIÇO], realizado pela "
        "MC2 Engenharia no âmbito da obra do [CLIENTE], sob responsabilidade "
        "da [CONSTRUTORA].\n\n"
        "[Contexto do serviço original] — descreva o que estava previsto no "
        "escopo original e qual foi o evento que gerou a necessidade do "
        "aditivo.\n\n"
        "[Solução proposta] — descreva o que precisou ser feito tecnicamente "
        "para resolver a situação, preservando funcionalidade e conformidade "
        "com normas técnicas aplicáveis.\n\n"
        "Registra-se que o serviço aqui descrito não integrava o escopo "
        "originalmente contratado, configurando-se como intervenção adicional "
        "decorrente de condição alheia à execução da MC2 Engenharia."
    ),
    "exclusoes": [
        "Serviços Civis tais como reboco, gesso, pintura etc.;",
        "Serviços fora do horário comercial;",
        "Instalação e fornecimento de equipamentos que não sejam os contemplados no escopo (item 1);",
    ],
    "valor_total": "R$ 1.350,00 (Mil trezentos e cinquenta reais).",
    "condicoes_pagamento": "Pagamento – 100% no aceite.",
}


# -------------------------------------------------------------
# DICIONÁRIO DE DADOS por tipo (não precisa mexer aqui)
# -------------------------------------------------------------
DADOS = {
    "A": DADOS_TIPO_A,
    "B": DADOS_TIPO_B,
    "C": DADOS_TIPO_C,
    "D": DADOS_TIPO_D,
}


# =============================================================
# FUNÇÕES (não precisa mexer abaixo)
# =============================================================

# -------------------------------------------------------------
# TEXTO -> PARÁGRAFOS REAIS + SUBLISTA DE MATERIAIS (v5, correção 2026-07-08)
# -------------------------------------------------------------
# Convenção pros campos de texto corrido fundidos (escopo_geral/escopo/
# introducao): linha em branco = novo parágrafo de verdade; dentro de um
# bloco, se TODAS as linhas começam com "- ", esse bloco vira uma sublista
# de materiais em bullet (separada da prosa), não fica embutido no texto
# corrido. NUNCA usar \n dentro de uma string de marcador comum — o
# docxtpl escreve tudo num único <w:t>, então \n não vira quebra de
# parágrafo pro Word (bug #1 do bugs_conhecidos.md). Por isso os
# parágrafos e a sublista são montados como Subdoc (parágrafos <w:p> de
# verdade), igual ao mecanismo já usado pras Exclusões/Objetivo.
def _texto_para_blocos(texto: str) -> list[tuple[str, str]]:
    """Quebra o texto em blocos ("texto" = parágrafo de prosa, "item" =
    linha de uma sublista de materiais). Uma linha em branco separa
    parágrafos; um bloco onde toda linha começa com "- " vira uma
    sequência de itens de lista."""
    blocos = []
    for bruto in re.split(r"\n\s*\n", (texto or "").strip()):
        linhas = [l.strip() for l in bruto.split("\n") if l.strip()]
        if not linhas:
            continue
        if len(linhas) > 1 and all(l.startswith("- ") for l in linhas):
            for linha in linhas:
                blocos.append(("item", linha[2:].strip()))
        else:
            blocos.append(("texto", " ".join(linhas)))
    return blocos


def _construir_subdoc_texto(doc: DocxTemplate, texto: str):
    """Monta um Subdoc com parágrafos reais a partir de _texto_para_blocos:
    prosa vira parágrafo "Normal", sublista de materiais vira bullets
    "Topicos" (bullet simples, sem numeração — não é uma seção do
    documento, não deve ligar ao contador do Heading 1)."""
    sub = doc.new_subdoc()
    for tipo, conteudo in _texto_para_blocos(texto):
        if not conteudo:
            continue
        estilo = "Topicos" if tipo == "item" else "Normal"
        sub.add_paragraph(conteudo, style=estilo)
    return sub


def validar_dados(tipo: str, dados: dict) -> list:
    """Verifica campos faltantes para um tipo. Retorna lista de faltantes."""
    esperados = MARCADORES_POR_TIPO[tipo]
    fornecidos = set(dados.keys())
    return sorted(esperados - fornecidos)


def gerar_proposta(tipo: str, dados: dict) -> Path:
    """Gera DOCX usando o template do tipo escolhido."""
    PASTA_SAIDA.mkdir(exist_ok=True)

    template_path = TEMPLATES[tipo]
    if not template_path.exists():
        raise FileNotFoundError(f"Template não encontrado: {template_path}")

    doc = DocxTemplate(template_path)

    # Monta o contexto de render a partir dos dados de entrada, derivando os
    # campos que não são mais digitados (Fase 2):
    contexto = dict(dados)

    # 1) Apresentação automática (se não vier pronta no dicionário).
    contexto.setdefault("paragrafo_apresentacao", montar_apresentacao(tipo, dados))

    # 2) Escopo/Introdução (texto corrido fundido, v5) -> Subdoc com
    #    parágrafos reais + sublista de materiais quando marcada (ver
    #    _construir_subdoc_texto). O template tem {{p bloco_escopo_geral}}/
    #    {{p bloco_escopo}}/{{p bloco_introducao}} conforme o tipo.
    for campo in ("escopo_geral", "escopo", "introducao"):
        valor = contexto.pop(campo, None)
        if valor is not None:
            contexto[f"bloco_{campo}"] = _construir_subdoc_texto(doc, valor)

    # 3) Exclusões -> Subdoc com numeração LIGADA ao Heading da seção
    #    (estilo "TopicosNumerados": mesmo numId do Heading 1, um nível
    #    abaixo — Word calcula "2.1, 2.2..." ou "3.1, 3.2..." sozinho,
    #    conforme o número que aquela seção recebeu naquele tipo. Não é
    #    mais bullet solto nem número hardcoded — ver bugs_conhecidos.md,
    #    correção de 2026-07-08). NUNCA usar \n/• numa string (bug #1). O
    #    Subdoc precisa ser criado a partir DESTE DocxTemplate, por isso é
    #    montado aqui dentro e não no app. O template tem um único
    #    {{p bloco_exclusoes}}.
    exclusoes = contexto.pop("exclusoes", None)
    if exclusoes is not None:
        sub = doc.new_subdoc()
        for item in exclusoes:
            item = (item or "").strip()
            if item:
                sub.add_paragraph(item, style="TopicosNumerados")
        contexto["bloco_exclusoes"] = sub

    # 4) Objetivo (só Tipo C, v5) -> mesmo mecanismo de numeração ligada ao
    #    Heading das Exclusões acima. O template tem {{p bloco_objetivos}}.
    objetivos = contexto.pop("objetivos", None)
    if objetivos is not None:
        sub_obj = doc.new_subdoc()
        for item in objetivos:
            item = (item or "").strip()
            if item:
                sub_obj.add_paragraph(item, style="TopicosNumerados")
        contexto["bloco_objetivos"] = sub_obj

    doc.render(contexto)

    # Nome do arquivo no padrão real da MC2 (v5, 2026-07-08 — ver
    # bugs_conhecidos.md "Mockup do painel único v2/v3"): José mostrou um
    # arquivo real, "PS26_115_Proposta Comercial rev00JO.docx" — sem
    # espaço depois de "PS", revisão com 2 dígitos, iniciais de quem
    # gerou coladas no fim da revisão. "responsavel" é só metadado do
    # nome do arquivo (José confirmou que NÃO aparece no corpo do
    # documento) — por isso não é marcador do template nem entra em
    # MARCADORES_POR_TIPO/validar_dados; vem opcionalmente em `dados`.
    numero_limpo = dados["numero_proposta"].replace(" ", "")
    revisao_fmt = str(dados.get("revisao", "00")).strip().zfill(2)
    responsavel = (dados.get("responsavel") or "").strip().upper()
    nome_arquivo = f"{numero_limpo}_Proposta Comercial rev{revisao_fmt}{responsavel}.docx"
    caminho_saida = PASTA_SAIDA / nome_arquivo
    doc.save(caminho_saida)
    return caminho_saida


def main():
    print("=" * 60)
    print(f"GERADOR DE PROPOSTAS MC2 — TIPO {TIPO_PROPOSTA}")
    print("=" * 60)

    if TIPO_PROPOSTA not in TEMPLATES:
        print(f"ERRO: Tipo '{TIPO_PROPOSTA}' inválido. Use A, B, C ou D.")
        return

    dados = DADOS[TIPO_PROPOSTA]
    faltantes = validar_dados(TIPO_PROPOSTA, dados)
    if faltantes:
        print(f"ERRO: Campos faltando: {faltantes}")
        return

    try:
        arquivo = gerar_proposta(TIPO_PROPOSTA, dados)
        print(f"\n✓ Proposta gerada com sucesso!")
        print(f"  Tipo:    {TIPO_PROPOSTA}")
        print(f"  Arquivo: {arquivo.absolute()}")
        print(f"  Tamanho: {arquivo.stat().st_size / 1024:.1f} KB")
        print(f"\nAbra o arquivo no Word pra conferir o resultado.")
    except Exception as e:
        print(f"\n✗ ERRO ao gerar a proposta: {e}")


if __name__ == "__main__":
    main()
