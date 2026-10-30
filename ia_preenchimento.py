"""
============================================================
IA — Preenchimento automático a partir de texto livre (Fase 3)
============================================================
José recebe a descrição do serviço de forma informal (técnico explica
pessoalmente ou manda mensagem de WhatsApp) — não um formulário
estruturado. Este módulo pega esse texto cru e devolve um dicionário já
no padrão MC2, usando os marcadores exatos daquele tipo de proposta.

Responsabilidade única: IA preenche o DICIONÁRIO. Quem gera o .docx
continua sendo só o docxtpl, em etapa2_template.py (não misturar as
duas coisas — ver regra do CLAUDE.md do projeto).

Requer a variável de ambiente ANTHROPIC_API_KEY (chave criada em
console.anthropic.com). Sem ela, preencher_com_ia() levanta
RuntimeError com uma mensagem clara — não falha silenciosamente.

NÃO preenche por IA: numero_proposta, revisao, data_extenso, cliente,
titulo_aditivo, cidade, sistema, valor_total, condicoes_pagamento.
São dados de identificação/negócio que José prefere sempre digitar (a
decisão original de deixar a IA tentar cliente/sistema/valor foi revista
em 2026-07-07: José quer controlar esses campos ele mesmo, a IA só
escreve o TEXTO — escopo, exclusões, mão de obra, introdução, objetivos,
cronograma, justificativa — nunca os dados de negócio). Tudo o que a IA
devolve ainda passa pela tela de revisão do Streamlit antes de gerar o
.docx — nunca é usado direto.
============================================================
"""

import json
import os
from pathlib import Path

import etapa2_template as gen

BASE_DIR = Path(__file__).resolve().parent
MANUAL_ESTILO_PATH = (
    BASE_DIR.parent / ".claude" / "skills" / "gerador-propostas-mc2"
    / "references" / "manual_estilo.md"
)

# Campos que a IA NUNCA deve preencher — dados administrativos/de negócio
# que José prefere digitar ele mesmo. A IA só escreve texto/prosa (escopo,
# exclusões, mão de obra, introdução, objetivos, cronograma, justificativa).
CAMPOS_FORA_DA_IA = {
    "numero_proposta", "revisao", "data_extenso",
    "cliente", "titulo_aditivo", "cidade", "sistema",
    "valor_total", "condicoes_pagamento",
}

MODELO = "claude-sonnet-4-5-20250929"


# -------------------------------------------------------------
# Pré-preenchimento DETERMINÍSTICO de exclusões (v5 — painel único)
# -------------------------------------------------------------
# Catálogo copiado de manual_estilo.md ("Catálogo de exclusões padrão").
# Disparado pelo app.py assim que tipo+sistema são escolhidos, ANTES/
# independente da chamada de IA — a IA (se usada depois) só complementa
# em cima do que já veio daqui. Não tem catálogo equivalente pra
# "objetivos" (só Exclusões têm essa lista pronta no manual).
_EXCLUSOES_NUCLEO = [
    "Serviços Civis tais como reboco, gesso, pintura etc.;",
    "Instalação e fornecimento de equipamentos que não sejam os contemplados no escopo (item 1);",
]
_EXCLUSOES_COMUNS = [
    "Serviços fora do horário comercial;",
    "Esta proposta não contempla o fornecimento de projeto nem de ART.",
]
# Específico por TIPO de proposta.
_EXCLUSOES_POR_TIPO = {
    "C": [
        "Fornecimento de materiais, cabos ou equipamentos de qualquer natureza;",
        "Ampliação do sistema ou instalação de novos dispositivos;",
        "Fornecimento de ESCADA.",
    ],
}
# Específico por SISTEMA (sigla em maiúsculo).
_EXCLUSOES_POR_SISTEMA = {
    "SCA": [
        "Fixação das catracas;",
        "Fornecimento de portinhola de PNE;",
        "Mola de porta e demais elementos mecânicos associados às portas;",
    ],
}


def pre_preencher_exclusoes(tipo: str, sistema: str) -> list[str]:
    """Monta a lista de exclusões determinística (núcleo + comuns + específico
    por tipo/sistema) assim que tipo+sistema forem escolhidos no painel —
    sem chamar a IA. Não faz chamada de rede; sempre retorna algo (mesmo com
    tipo/sistema vazios, ao menos núcleo+comuns)."""
    itens = list(_EXCLUSOES_NUCLEO) + list(_EXCLUSOES_COMUNS)
    itens += _EXCLUSOES_POR_TIPO.get((tipo or "").strip().upper(), [])
    itens += _EXCLUSOES_POR_SISTEMA.get((sistema or "").strip().upper(), [])
    return itens


def _carregar_manual_estilo() -> str:
    """Lê o Manual de Estilo MC2 se disponível; senão devolve um resumo mínimo."""
    if MANUAL_ESTILO_PATH.exists():
        return MANUAL_ESTILO_PATH.read_text(encoding="utf-8")
    # Fallback curto, caso o arquivo não esteja acessível neste ambiente.
    return (
        "Use linguagem formal MC2/AQ3 Engenharia: 'V.Sas.', 'o presente/a "
        "presente proposta', 'contemplando', 'em conformidade com'. Exclusões "
        "sempre incluem: 'Serviços Civis tais como reboco, gesso, pintura "
        "etc.;' e 'Instalação e fornecimento de equipamentos que não sejam "
        "os contemplados no escopo (item 1);'. Não invente marca/modelo se "
        "não vier no texto do técnico."
    )


def _campos_para_ia(tipo: str) -> list[str]:
    """Campos esperados do tipo, exceto os administrativos (fora da IA)."""
    return [c for c in gen.MARCADORES_POR_TIPO[tipo] if c not in CAMPOS_FORA_DA_IA]


def _montar_prompt_sistema(tipo: str) -> str:
    manual = _carregar_manual_estilo()
    campos = _campos_para_ia(tipo)
    exemplo = {k: v for k, v in gen.DADOS[tipo].items() if k in campos}

    return f"""Você ajuda a MC2/AQ3 Engenharia (automação predial em Brasília/DF)
a transformar a explicação informal de um técnico (fala transcrita ou
mensagem de WhatsApp) em um dicionário de campos para uma proposta
comercial Tipo {tipo}.

⚠️ ELABORE o texto, não reformule minimamente. Você tem uma conversa de
verdade por trás — a mensagem do técnico é o PONTO DE PARTIDA (os fatos),
não o texto final a ser só reescrito/polido. Escreva parágrafos completos,
com linguagem técnica adequada, no tom das propostas reais MC2 (ver
manual abaixo) — não uma paráfrase curta do que o técnico digitou. Pode
complementar com conhecimento técnico geral do tipo de serviço/sistema
(ex: por que aquele procedimento é recomendado, boas práticas usuais,
o que costuma estar envolvido tecnicamente numa manutenção/instalação
desse tipo) para enriquecer a APRESENTAÇÃO do conteúdo. **Mas nunca
invente FATOS CONCRETOS não confirmados pelo técnico** — marca, modelo,
quantidade exata, valor, prazo, endereço específico etc. continuam saindo
só do que foi relatado (ou em branco, se não veio). A linha é: pode
enriquecer COMO algo é descrito/justificado tecnicamente; não pode
inventar O QUE foi contratado/fornecido.

MANUAL DE ESTILO MC2 (siga rigorosamente o vocabulário e os padrões abaixo):
---
{manual}
---

Devolva SOMENTE um objeto JSON válido (sem markdown, sem \`\`\`, sem texto
antes ou depois), contendo exatamente estas chaves:
{json.dumps(campos, ensure_ascii=False)}

Regras:
- "exclusoes" é uma LISTA de strings curtas, cada uma terminando em ";" —
  comece pelo núcleo padrão do manual e adicione as específicas do
  contexto relatado. O app já pré-preenche essa lista deterministicamente
  a partir do catálogo do manual antes de te chamar (ver
  `pre_preencher_exclusoes`); você pode complementar/ajustar em cima disso,
  não precisa começar do zero.
- "objetivos" (só Tipo C) também é uma LISTA de strings curtas, uma por
  item, no formato de tópico (o que a manutenção/visita/laudo pretende
  alcançar) — não é texto corrido.
- Campos de texto corrido fundidos ("escopo_geral" no Tipo A, "escopo" nos
  Tipos B/D, "introducao" no Tipo C) cobrem o que antes eram várias seções
  separadas (escopo geral + fornecimento + serviços + mão de obra, no caso
  do Tipo B; introdução + mão de obra + cronograma, no caso do Tipo C) —
  mas NÃO é um parágrafo único corrido: use a convenção abaixo pra formatar
  como parágrafos reais e, se aplicável, uma sublista de materiais:
  - Separe parágrafos de verdade com uma LINHA EM BRANCO (`\\n\\n`) entre
    eles — nunca um texto gigante numa linha só. 2 a 4 parágrafos curtos é
    o normal.
  - Quando o texto mencionar equipamentos/materiais A SEREM FORNECIDOS,
    coloque-os em um bloco SEPARADO (linha em branco antes e depois) onde
    CADA LINHA começa com "- " (hífen + espaço) — isso vira uma sublista
    em bullet no documento final, fora da prosa corrida. Não liste
    materiais dentro de uma frase corrida.
  - Não use subtítulos nem numeração manual dentro do texto (isso já é
    tratado pelo template).
  - Exemplo de formato esperado pro campo (Tipo B, ilustrativo):
    "O escopo contempla a substituição de 02 (dois) detectores de fumaça,
    incluindo fornecimento e mão de obra especializada.\\n\\n- 02 (dois)
    detectores de fumaça endereçáveis compatíveis com a central "
    "existente;\\n- Materiais de instalação necessários.\\n\\nOs serviços "
    "contemplam a desinstalação dos dispositivos atuais, instalação e "
    "comissionamento do sistema."
- Campos de texto livre devem soar como as propostas reais MC2: formais,
  diretas, no padrão do manual.
- Se uma informação não estiver clara na mensagem do técnico (ex: valor,
  cliente, cidade), devolva string vazia "" nesse campo em vez de inventar.
- Não invente marca/modelo de equipamento fora da tabela do manual.
- Exemplo de como os campos deste tipo costumam ficar preenchidos (não
  copie o conteúdo, é só referência de tom/tamanho):
{json.dumps(exemplo, ensure_ascii=False, indent=2)}
"""


def _extrair_json(texto: str) -> dict:
    """Extrai o primeiro bloco JSON de uma resposta (tolera cercas ```)."""
    texto = texto.strip()
    if texto.startswith("```"):
        texto = texto.strip("`")
        if texto.lower().startswith("json"):
            texto = texto[4:]
        texto = texto.strip()
    inicio = texto.find("{")
    fim = texto.rfind("}")
    if inicio == -1 or fim == -1:
        raise ValueError(f"Resposta da IA não contém JSON reconhecível: {texto[:200]}")
    return json.loads(texto[inicio:fim + 1])


def preencher_com_ia(tipo: str, texto_bruto: str) -> dict:
    """
    Chama a Claude API pra transformar texto_bruto (mensagem/explicação do
    técnico) num dicionário de campos pro tipo indicado. Levanta
    RuntimeError se faltar a chave, ValueError se a resposta não vier em
    JSON válido.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY não configurada. Crie uma chave em "
            "console.anthropic.com e defina a variável de ambiente antes "
            "de rodar o Streamlit (ex: no PowerShell: "
            "$env:ANTHROPIC_API_KEY='sua-chave')."
        )
    if not texto_bruto or not texto_bruto.strip():
        raise ValueError("Cole a mensagem/descrição do técnico antes de preencher com IA.")

    import anthropic  # import local pra não exigir a lib se a função não for usada

    cliente = anthropic.Anthropic(api_key=api_key)
    resposta = cliente.messages.create(
        model=MODELO,
        max_tokens=2000,
        system=_montar_prompt_sistema(tipo),
        messages=[{"role": "user", "content": texto_bruto.strip()}],
    )
    texto_resposta = "".join(
        bloco.text for bloco in resposta.content if getattr(bloco, "type", None) == "text"
    )
    dados = _extrair_json(texto_resposta)

    # Segurança extra: nunca deixar a IA devolver campo administrativo,
    # mesmo que o modelo "alucine" e inclua um por conta própria.
    for campo in CAMPOS_FORA_DA_IA:
        dados.pop(campo, None)

    return dados
