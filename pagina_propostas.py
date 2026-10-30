"""
============================================================
INTERFACE — Gerador de Propostas MC2 (v5 — painel único)
============================================================
App Streamlit pra José preencher/salvar/gerar propostas sem
editar os dicionários DADOS_TIPO_X no código.

v5 (2026-07-08): interface reconstruída como painel único editável, no
formato do documento final — poucos blocos de texto grandes (escopo,
introdução, exclusões), editáveis diretamente, na ordem real do
documento. Tipo e sistema viram select; Aditivo vira checkbox/flag que
troca o tipo efetivo pra D (não é mais uma 5ª opção paralela no seletor
de tipo — ver bugs_conhecidos.md, "Debate completo do novo
formulário/preview único"). Ver esse mesmo arquivo pra a simplificação
DELIBERADA de o Aditivo substituir 100% o tipo base (sem lista de
Objetivo), em vez de "aditivo por cima" do conteúdo de Produto/Serviço.

REAPROVEITA o etapa2_template.py (não duplica geração):
    - gen.gerar_proposta(tipo, dados)  -> gera o .docx
    - gen.MARCADORES_POR_TIPO          -> validação
    - gen.DADOS[tipo]                  -> ordem dos campos + exemplo
Persistência em SQLite via db.py.

Rodar:  py -m streamlit run app.py   (a partir da pasta propostas-mc2 —
        "streamlit run" sozinho NÃO funciona nesta máquina, streamlit.exe
        não está no PATH)
============================================================
"""

import base64
import zipfile
from pathlib import Path

import mammoth
import streamlit as st
from docx2pdf import convert as _docx2pdf_convert

import db
import etapa2_template as gen
import ia_preenchimento as ia

# -------------------------------------------------------------
# Aparência (visual puro — não mexe em dados/lógica).
# Cores/raio vêm de .streamlit/config.toml (paleta extraída do logo MC2:
# vermelho #C7393E / cinza #98999E). Este CSS só cobre o que o tema do
# Streamlit não alcança: acabamento do título, do card de confirmação da
# IA e da caixa de chat, e remoção do rodapé "Made with Streamlit".
# -------------------------------------------------------------
_CSS = """
<style>
/* Título principal: acabamento tipo "cabeçalho corporativo", sem ícone */
[data-testid="stAppViewContainer"] h1 {
    border-bottom: 3px solid #C7393E;
    padding-bottom: 0.5rem;
    margin-bottom: 0.25rem;
    letter-spacing: -0.01em;
}

/* Cabeçalhos da sidebar em caixa alta discreta, estilo painel de navegação */
[data-testid="stSidebar"] h2 {
    font-size: 0.95rem;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: #26272B;
}

/* Card de confirmação da IA (chave "card_ia_confirmacao"): leve destaque
   na cor da marca em vez de depender de emoji pra sinalizar "aguardando
   ação". Usa a chave do container (não o test-id genérico) pra não vazar
   esse destaque pros outros cartões do painel (identificação/conteúdo,
   exclusões). */
.st-key-card_ia_confirmacao {
    background-color: #FBEEEF;
}

/* Cartão principal do painel único: leve respiro interno além da borda
   nativa do st.container(border=True). */
.st-key-card_painel {
    padding: 0.5rem 0.25rem;
}

/* Caixa de chat: raio consistente com o resto do app (não "bolha de chat") */
[data-testid="stChatInput"] {
    border-radius: 0.4rem;
}

/* Remove o rodapé de marca do Streamlit (mantém o menu de hambúrguer,
   que ainda tem funções úteis como Rerun/Print) */
footer {
    visibility: hidden;
}
</style>
"""

# -------------------------------------------------------------
# Metadados de apresentação
# -------------------------------------------------------------
# Tipo "base" (o que o usuário escolhe) — Aditivo NÃO é uma opção aqui,
# é o checkbox separado logo abaixo (ver _tipo_efetivo). Ordem da lista
# segue a taxonomia de negócio já debatida com José: Produto / Serviço /
# Produto e Serviço.
TIPOS_BASE_DESC = {
    "A": "Produto (só fornecimento)",
    "C": "Serviço (manutenção, visita, laudo)",
    "B": "Produto e Serviço (fornecimento + instalação)",
}
TIPOS_BASE_ORDEM = ["A", "C", "B"]

# Sistema: select fechado + "Outro" com texto livre (v5). Siglas em
# maiúsculo pra bater com as chaves de gen.SISTEMAS (nome por extenso na
# apresentação). "Outro" não é uma chave de gen.SISTEMAS — nesse caso o
# texto livre digitado vai direto pro campo "sistema".
SISTEMA_OPCOES = ["CFTV", "SDAI", "SCA", "AUT", "INTERFONIA", "Outro"]
SISTEMA_ROTULOS = {
    "CFTV": "CFTV",
    "SDAI": "SDAI",
    "SCA": "SCA (Controle de Acesso)",
    "AUT": "AUT (Automação Predial)",
    "INTERFONIA": "Interfonia",
    "Outro": "Outro (digitar)",
}

# Campos de texto corrido fundidos (v5) — ganham um text_area mais alto
# porque agora carregam o que antes eram várias seções separadas.
CAMPOS_GRANDES = {"escopo_geral", "escopo", "introducao"}

# Campos que são lista (list[str]) no dicionário de dados, não string.
# Editados como texto multi-linha (uma linha = um item da lista).
CAMPOS_LISTA = {"exclusoes", "objetivos"}

# Rótulos amigáveis; o que não estiver aqui é gerado automaticamente
ROTULOS = {
    "numero_proposta": "Número da proposta",
    "revisao": "Revisão",
    "data_extenso": "Data (por extenso)",
    "cliente": "Cliente",
    "cidade": "Cidade",
    "titulo_aditivo": "Título do aditivo",
    "sistema": "Sistema (aparece no campo Ref.:)",
    "escopo_geral": "Escopo",
    "escopo": "Escopo",
    "introducao": "Introdução",
    "objetivos": "Objetivo",
    "exclusoes": "Exclusões (pré-preenchidas ao escolher tipo+sistema)",
    "responsavel": "Responsável (iniciais)",
    "valor_total": "Valor total",
    "condicoes_pagamento": "Condições de pagamento",
}


def rotulo(campo: str) -> str:
    return ROTULOS.get(campo) or campo.replace("_", " ").capitalize()


def campos_do_tipo(tipo: str) -> list[str]:
    """Ordem dos campos = ordem das chaves no dicionário de exemplo do tipo
    (já está na ordem do documento final — ver etapa2_template.py)."""
    return list(gen.DADOS[tipo].keys())


def _tipo_efetivo(tipo_base: str, aditivo: bool) -> str:
    """Aditivo é um FLAG, não um tipo paralelo: quando marcado, substitui
    100% o tipo base pelo Tipo D (template/campos de aditivo contratual),
    independente do que estava selecionado em Produto/Serviço/Produto e
    Serviço. Simplificação DELIBERADA (não junta o conteúdo do tipo base
    com o de aditivo) — ver bugs_conhecidos.md, decisão de 2026-07-08."""
    return "D" if aditivo else tipo_base


def _reverse_sistema(valor: str) -> tuple[str, str]:
    """Valor salvo de 'sistema' -> (opção do select, texto de 'Outro')."""
    v = (valor or "").strip()
    if v.upper() in SISTEMA_OPCOES[:-1]:  # tudo exceto "Outro"
        return v.upper(), ""
    return "Outro", v


def _reverse_pagamento(valor: str) -> tuple[str, str]:
    """Valor salvo de 'condicoes_pagamento' -> (rótulo do preset, texto de 'Outro')."""
    v = (valor or "").strip()
    for label, frase in gen.PAGAMENTOS.items():
        if frase == v:
            return label, ""
    return "Outro", v


# -------------------------------------------------------------
# PDF opcional via Word instalado na máquina (docx2pdf, automação COM —
# só funciona no Windows com Microsoft Word instalado). v5
# (2026-07-08, mockup v2/v3): substitui o LibreOffice — José rejeitou
# depender dele (nem estava instalado nesta máquina) e prefere um PDF
# fiel, saído do Word de verdade. Conveniência, não substitui o .docx.
# -------------------------------------------------------------
def _converter_para_pdf(caminho_docx: Path) -> Path | None:
    caminho_pdf = caminho_docx.with_suffix(".pdf")
    # docx2pdf usa win32com.client.Dispatch direto, sem inicializar o COM —
    # funciona rodando um script comum (thread principal, COM implícito),
    # mas o Streamlit executa o script numa thread própria (ScriptRunner),
    # onde o COM NUNCA foi inicializado -> "CoInitialize não foi chamado"
    # (testado ao vivo em 2026-07-09). Precisa inicializar/desinicializar o
    # apartment COM manualmente nesta thread a cada chamada.
    import pythoncom

    pythoncom.CoInitialize()
    try:
        _docx2pdf_convert(str(caminho_docx), str(caminho_pdf))
    except Exception as e:
        st.error(
            f"Falha ao converter pra PDF via Word: {e}. Confira se o "
            "Microsoft Word está instalado e se não há nenhuma janela dele "
            "travada esperando confirmação (feche diálogos abertos do Word "
            "antes de tentar de novo)."
        )
        return None
    finally:
        pythoncom.CoUninitialize()
    return caminho_pdf if caminho_pdf.exists() else None


# -------------------------------------------------------------
# Preview do documento formatado (mammoth: .docx -> HTML direto, sem
# LibreOffice). v5 (2026-07-08, mockup v2/v3): gerado sob demanda (botão
# "Visualizar"), não em tempo real — mammoth não é instantâneo o
# suficiente pra reconverter a cada tecla. Roda sobre o MESMO .docx que
# gen.gerar_proposta produziria (não é uma renderização paralela).
# ⚠️ Fidelidade limitada: mammoth mapeia estilos conhecidos (heading via
# outline level, negrito/itálico) mas pode não reconhecer 100% dos
# estilos customizados do template (ex: numeração ligada ao Heading) —
# é conveniência de revisão rápida, o .docx/Word continua sendo a fonte
# confiável.
# -------------------------------------------------------------
# O mammoth detecta lista via numeração do parágrafo, mas não resolve
# numPr herdado de ESTILO (só o que está direto no parágrafo) — por isso
# "TopicosNumerados" (Exclusões/Objetivo) não vira <ul>/<ol> sozinho. Este
# style_map força isso pra pelo menos aparecer como lista no preview
# (perde a numeração "2.1, 2.2" ligada ao Heading, que só o Word calcula
# de verdade — é uma aproximação visual, não o resultado real).
_MAMMOTH_STYLE_MAP = "p[style-name='TopicosNumerados'] => ul > li:fresh"

# Paleta da marca (mesma do _CSS principal, extraída do logo MC2).
_COR_VERMELHO = "#C7393E"
_COR_CINZA_ESCURO = "#26272B"


def _extrair_marca_dagua_base64(caminho_docx: Path) -> str | None:
    """Extrai a maior imagem de word/media/ do .docx gerado — é a marca
    d'água/timbrado (logo MC2) que o Word aplica no cabeçalho com um efeito
    de "lavado" (gain/blacklevel) que o mammoth não reproduz, porque ele só
    converte o corpo do documento, não cabeçalho/rodapé. Lê do .docx recém
    gerado (não de uma cópia solta no projeto) pra nunca ficar desatualizado
    se o timbrado mudar. Pega a maior imagem, não uma pelo nome, porque o
    template também tem uma imagem minúscula (glifo de marcador de lista)
    que não é o logo — ver numeração dos arquivos em word/media/."""
    try:
        with zipfile.ZipFile(caminho_docx) as z:
            candidatos = [n for n in z.namelist() if n.startswith("word/media/")]
            if not candidatos:
                return None
            maior = max(candidatos, key=lambda n: z.getinfo(n).file_size)
            dados = z.read(maior)
        return base64.b64encode(dados).decode("ascii")
    except (KeyError, zipfile.BadZipFile, FileNotFoundError):
        return None


# CSS do "papel" do preview — aproximação visual (fundo de página, margens,
# fonte Arial batendo com o template, marca d'água centralizada) pra ficar
# parecido com o .docx/PDF real. Não é renderização fiel (isso exigiria
# LibreOffice/Word rodando no servidor, descartado por complexidade/custo
# pro que é "só" uma conveniência de revisão rápida — ver decisão de
# 2026-07-09). Fica num documento HTML separado (não usa o _CSS do app)
# porque roda dentro do próprio iframe do st.iframe, isolado do resto da
# página.
_PREVIEW_CSS_TEMPLATE = """
<style>
  body {{
    margin: 0;
    padding: 24px 0;
    background: #d9d9dc;
    font-family: Arial, Helvetica, sans-serif;
  }}
  .pagina {{
    position: relative;
    max-width: 780px;
    margin: 0 auto;
    background: #ffffff;
    box-shadow: 0 2px 14px rgba(0, 0, 0, 0.25);
    padding: 64px 72px;
    color: #1a1a1a;
    line-height: 1.5;
  }}
  .pagina::before {{
    content: "";
    position: absolute;
    inset: 0;
    background-image: url("data:image/png;base64,{marca_dagua}");
    background-repeat: no-repeat;
    background-position: center 38%;
    background-size: 62% auto;
    opacity: 0.10;
    pointer-events: none;
  }}
  .conteudo {{
    position: relative;
  }}
  .conteudo p {{
    margin: 0 0 10px 0;
  }}
  .conteudo h1, .conteudo h2, .conteudo h3 {{
    color: {vermelho};
    font-weight: bold;
    letter-spacing: -0.01em;
    margin: 20px 0 8px 0;
  }}
  .conteudo strong {{
    color: {cinza};
  }}
  .conteudo ul {{
    margin: 6px 0 14px 0;
    padding-left: 22px;
  }}
  .conteudo li {{
    margin-bottom: 4px;
  }}
</style>
"""


def _gerar_preview_html(caminho_docx: Path) -> str:
    with open(caminho_docx, "rb") as fh:
        resultado = mammoth.convert_to_html(fh, style_map=_MAMMOTH_STYLE_MAP)

    marca_dagua = _extrair_marca_dagua_base64(caminho_docx)
    if not marca_dagua:
        # Sem imagem encontrada -> mostra o conteúdo mesmo assim, só sem
        # marca d'água (silencioso de propósito: não é motivo pra travar
        # o preview inteiro).
        return resultado.value

    css = _PREVIEW_CSS_TEMPLATE.format(
        marca_dagua=marca_dagua, vermelho=_COR_VERMELHO, cinza=_COR_CINZA_ESCURO
    )
    return (
        f"{css}"
        f'<div class="pagina"><div class="conteudo">{resultado.value}</div></div>'
    )


# -------------------------------------------------------------
# Listas editáveis item a item (Exclusões e Objetivo, v5 mockup v2/v3) —
# cada item é um id ESTÁVEL (contador global, nunca reaproveitado), não a
# posição na lista: se a key do widget dependesse da posição, remover um
# item do meio faria o Streamlit reexibir o texto ANTIGO nas posições
# seguintes (ele ignora `value=` quando a key já existe no session_state
# de uma rodada anterior). Com id estável isso não acontece.
# -------------------------------------------------------------
def _proximo_item_id() -> int:
    atual = st.session_state.get("_prox_item_id", 0)
    st.session_state["_prox_item_id"] = atual + 1
    return atual


def _definir_lista(session_key: str, valores: list[str]):
    """Substitui o conteúdo da lista por completo (troca de tipo, carregar
    exemplo/salva, ou preenchimento pela IA) — sempre com ids novos."""
    st.session_state[session_key] = [
        {"id": _proximo_item_id(), "texto": v} for v in valores
    ]


def _lista_valores(session_key: str) -> list[str]:
    return [
        item["texto"].strip()
        for item in st.session_state.get(session_key, [])
        if item["texto"].strip()
    ]


def _lista_editavel(session_key: str, placeholder_add: str):
    """Renderiza a lista como um item por linha (texto editável + botão de
    remover) e um campo no fim que adiciona item ao apertar Enter."""
    itens = st.session_state.setdefault(session_key, [])
    for item in list(itens):
        col_item, col_remover = st.columns([10, 1])
        with col_item:
            item["texto"] = st.text_input(
                f"item {item['id']}", value=item["texto"],
                key=f"{session_key}_txt_{item['id']}",
                label_visibility="collapsed",
            )
        with col_remover:
            if st.button(":material/close:", key=f"{session_key}_del_{item['id']}", help="Remover"):
                st.session_state[session_key] = [
                    i for i in st.session_state[session_key] if i["id"] != item["id"]
                ]
                st.rerun()

    def _adicionar():
        novo = st.session_state.get(f"{session_key}__novo", "").strip()
        if novo:
            st.session_state[session_key].append({"id": _proximo_item_id(), "texto": novo})
        st.session_state[f"{session_key}__novo"] = ""

    st.text_input(
        placeholder_add, key=f"{session_key}__novo",
        placeholder=placeholder_add, on_change=_adicionar,
        label_visibility="collapsed",
    )


# -------------------------------------------------------------
# Callbacks (rodam ANTES do rerun -> pode setar keys de widget)
# -------------------------------------------------------------
def _preencher(tipo: str, dados: dict, editing_id):
    st.session_state.tipo = tipo
    st.session_state.editing_id = editing_id
    if tipo == "D":
        st.session_state.aditivo = True
        # tipo_base não muda (Aditivo é um flag por cima; não afeta a
        # geração enquanto aditivo=True, então preserva a última escolha).
    else:
        st.session_state.aditivo = False
        st.session_state.tipo_base = tipo

    for k, v in dados.items():
        if k == "sistema":
            sigla, outro = _reverse_sistema(v)
            st.session_state["sel_sistema"] = sigla
            st.session_state["f_sistema_outro"] = outro
        elif k == "condicoes_pagamento":
            label, outro = _reverse_pagamento(v)
            st.session_state["sel_pagamento"] = label
            st.session_state["f_pagamento_outro"] = outro
        elif k in CAMPOS_LISTA and isinstance(v, list):
            _definir_lista(f"lista_{k}", v)
        else:
            st.session_state[f"f_{k}"] = v


def _novo_do_exemplo(tipo: str):
    _preencher(tipo, dict(gen.DADOS[tipo]), None)
    _prefill_exclusoes()


def _prefill_exclusoes():
    """Pré-preenchimento DETERMINÍSTICO das exclusões (núcleo + comuns +
    específico por tipo/sistema, do catálogo do manual_estilo.md) — roda
    assim que tipo efetivo ou sistema mudam, ANTES/independente da IA.
    Sobrescreve o campo Exclusões (é o comportamento pedido: escolher
    tipo+sistema já entrega uma lista pronta pra editar)."""
    tipo_atual = st.session_state.get("tipo", "A")
    sistema_valor = _sistema_valor_atual()
    _definir_lista("lista_exclusoes", ia.pre_preencher_exclusoes(tipo_atual, sistema_valor))


def _sistema_valor_atual() -> str:
    sel = st.session_state.get("sel_sistema", "")
    if sel == "Outro":
        return st.session_state.get("f_sistema_outro", "")
    return sel


def _on_tipo_base_change():
    novo_tipo = _tipo_efetivo(st.session_state.tipo_base, st.session_state.aditivo)
    _novo_do_exemplo(novo_tipo)


def _on_aditivo_change():
    novo_tipo = _tipo_efetivo(st.session_state.tipo_base, st.session_state.aditivo)
    _novo_do_exemplo(novo_tipo)


def _on_sistema_change():
    _prefill_exclusoes()


def _carregar_salva(pid: int):
    reg = db.obter(pid)
    if reg:
        _preencher(reg["tipo"], reg["dados"], pid)


def _aplicar_ia():
    """Chama a IA com o texto digitado na caixa "Preencher com IA" e popula
    os campos de texto do tipo atual. Não mexe em dados administrativos/de
    negócio (ver CAMPOS_FORA_DA_IA em ia_preenchimento.py). v5 (mockup
    v2/v3): fluxo de 1 passo só — a caixa deixou de ser um chat_input que
    manda ao apertar Enter, então o clique no botão "Preencher com IA" já
    é a confirmação explícita (não precisa mais do cartão de confirmação
    à parte que existia quando o envio era por Enter)."""
    tipo_atual = st.session_state.tipo
    texto = st.session_state.get("ia_texto_input", "")
    if not texto.strip():
        st.session_state["ia_mensagem"] = ("warning", "Descreva o serviço antes de preencher com IA.")
        return
    try:
        resultado = ia.preencher_com_ia(tipo_atual, texto)
    except Exception as e:
        st.session_state["ia_mensagem"] = ("error", f"Erro ao preencher com IA: {e}")
        return

    campos_validos = set(campos_do_tipo(tipo_atual))
    aplicados = 0
    for k, v in resultado.items():
        if k not in campos_validos:
            continue
        if k in CAMPOS_LISTA and isinstance(v, list):
            _definir_lista(f"lista_{k}", v)
        else:
            st.session_state[f"f_{k}"] = v
        aplicados += 1
    st.session_state["ia_mensagem"] = (
        "success",
        f"{aplicados} campo(s) preenchido(s) pela IA. Revise tudo antes de gerar o .docx.",
    )


# -------------------------------------------------------------
# Inicialização
# -------------------------------------------------------------
st.set_page_config(page_title="Gerador de Propostas MC2", page_icon=":material/description:", layout="wide")
st.markdown(_CSS, unsafe_allow_html=True)

try:
    db.init_db()
except Exception as e:
    st.error(
        f"Não consegui conectar ao banco (Turso): {e}\n\n"
        "Confira se `TURSO_DATABASE_URL` e `TURSO_AUTH_TOKEN` estão "
        "configuradas (Secrets, no Streamlit Community Cloud; variável de "
        "ambiente, localmente)."
    )
    st.stop()

if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.editing_id = None
    st.session_state.tipo_base = "A"
    st.session_state.aditivo = False
    _novo_do_exemplo("A")  # pré-preenche f_*/sel_* com o exemplo do Tipo A


# -------------------------------------------------------------
# SIDEBAR — só navegação (novo + propostas salvas). Tipo/Aditivo/Sistema
# subiram pra linha única no topo do painel principal (ver abaixo).
# -------------------------------------------------------------
with st.sidebar:
    st.header("Propostas")
    st.button(
        ":material/add: Novo (carregar exemplo do tipo)",
        use_container_width=True,
        on_click=lambda: _novo_do_exemplo(
            _tipo_efetivo(st.session_state.tipo_base, st.session_state.aditivo)
        ),
    )

    st.divider()
    st.subheader("Salvas no banco")
    salvas = db.listar()
    if not salvas:
        st.caption("Nenhuma proposta salva ainda.")
    for p in salvas:
        marca_aditivo = " · Aditivo" if p["tipo"] == "D" else ""
        rotulo_item = f"[{p['tipo']}{marca_aditivo}] {p['numero_proposta'] or '(sem nº)'} — {p['cliente']}"
        col1, col2 = st.columns([5, 1])
        col1.button(
            rotulo_item,
            key=f"load_{p['id']}",
            use_container_width=True,
            on_click=_carregar_salva,
            args=(p["id"],),
            help=f"Atualizada em {p['atualizado_em']}",
        )
        if col2.button(":material/delete:", key=f"del_{p['id']}", help="Excluir"):
            db.excluir(p["id"])
            if st.session_state.editing_id == p["id"]:
                st.session_state.editing_id = None
            st.rerun()


# -------------------------------------------------------------
# ÁREA PRINCIPAL — painel único do tipo atual (identificação + blocos de
# texto grandes, na ordem do documento final)
# -------------------------------------------------------------
tipo = st.session_state.tipo
editing_id = st.session_state.editing_id

st.title("Gerador de Propostas MC2/AQ3")
if editing_id:
    st.info(f"Editando proposta salva **#{editing_id}** (Tipo {tipo}). "
            f"Salvar vai atualizar essa mesma; use *Novo* pra começar do zero.")
else:
    st.caption(f"Nova proposta — Tipo {tipo}. Preencha e salve ou gere direto.")

# -------------------------------------------------------------
# Linha única de controles: tipo, aditivo e sistema — as 3 decisões que
# definem QUAL painel aparece embaixo, por isso ficam juntas no topo em
# vez de espalhadas pela sidebar/formulário.
# -------------------------------------------------------------
col_tipo, col_aditivo, col_sistema = st.columns([2, 1.4, 2])
with col_tipo:
    st.selectbox(
        "Tipo de proposta",
        options=TIPOS_BASE_ORDEM,
        format_func=lambda t: TIPOS_BASE_DESC[t],
        key="tipo_base",
        on_change=_on_tipo_base_change,
    )
with col_aditivo:
    st.checkbox(
        "É aditivo?",
        key="aditivo",
        on_change=_on_aditivo_change,
        help="Substitui a estrutura acima pela de aditivo contratual "
             "(ESCOPO + JUSTIFICATIVA fundidos + numeração XX_NNN.ADXX). "
             "Não combina com Produto/Serviço — é uma estrutura própria.",
    )
with col_sistema:
    st.session_state.setdefault("sel_sistema", "CFTV")
    st.selectbox(
        rotulo("sistema"),
        options=SISTEMA_OPCOES,
        format_func=lambda s: SISTEMA_ROTULOS[s],
        key="sel_sistema",
        on_change=_on_sistema_change,
    )
    if st.session_state["sel_sistema"] == "Outro":
        st.text_input("Sistema (texto livre)", key="f_sistema_outro")

# -------------------------------------------------------------
# IA — preenchimento a partir de texto livre (Fase 3). v5 (2026-07-08,
# mockup v2/v3): deixou de ser chat fixo no rodapé (chat_input manda ao
# apertar Enter) e virou uma caixa sempre visível (2-3 linhas) + botão
# ao lado, logo abaixo da linha Tipo/Sistema/Aditivo — o clique no botão
# já é a ação deliberada, não precisa mais de um cartão de confirmação
# à parte.
# -------------------------------------------------------------
if "ia_mensagem" in st.session_state:
    nivel, msg = st.session_state.pop("ia_mensagem")
    getattr(st, nivel)(msg)

col_ia_texto, col_ia_botao = st.columns([5, 1])
with col_ia_texto:
    st.text_area(
        "Preencher com IA",
        key="ia_texto_input",
        height=90,
        placeholder="Descreva o serviço que será realizado, o que foi feito e como...",
        label_visibility="collapsed",
    )
with col_ia_botao:
    st.write("")
    st.button(
        ":material/auto_awesome: Preencher com IA", type="primary",
        use_container_width=True, on_click=_aplicar_ia,
    )
st.caption(
    "A IA não mexe em número da proposta/revisão/data/cliente/sistema/"
    "cidade/valor/pagamento/responsável — isso continua com você. Ela só "
    "escreve os textos (escopo, exclusões etc., complementando o "
    "pré-preenchimento automático). Sempre revise antes de gerar."
)

st.divider()

# -------------------------------------------------------------
# Cartão do painel único: identificação em colunas, conteúdo (escopo/
# introdução/objetivo) na ordem do documento, exclusões destacadas num
# quadro com borda leve, e comercial (valor + pagamento) lado a lado.
# -------------------------------------------------------------
campos_tipo_atual = campos_do_tipo(tipo)
ident_l1 = [c for c in ("numero_proposta", "revisao", "data_extenso") if c in campos_tipo_atual]
# "responsavel" (iniciais de quem gerou) não é marcador do template — José
# confirmou que só entra no NOME do arquivo, não no corpo do documento
# (ver bugs_conhecidos.md). Por isso não vem de campos_tipo_atual/gen.DADOS,
# é sempre adicionado aqui na Identificação.
ident_l2 = [c for c in ("titulo_aditivo", "cliente", "cidade") if c in campos_tipo_atual] + ["responsavel"]
campos_conteudo = [
    c for c in campos_tipo_atual
    if c not in {"sistema", "exclusoes", "valor_total", "condicoes_pagamento", *ident_l1, *ident_l2}
]

dados = {"sistema": _sistema_valor_atual()}

with st.container(border=True, key="card_painel"):
    st.markdown("###### Identificação")
    for grupo in (ident_l1, ident_l2):
        cols = st.columns(len(grupo))
        for col, campo in zip(cols, grupo):
            with col:
                key = f"f_{campo}"
                st.session_state.setdefault(key, gen.DADOS[tipo].get(campo, ""))
                dados[campo] = st.text_input(rotulo(campo), key=key)

    st.divider()
    for campo in campos_conteudo:
        if campo in CAMPOS_LISTA:  # "objetivos" (só Tipo C) — lista item a item
            st.markdown(f"**{rotulo(campo)}**")
            with st.container(border=True, key=f"card_{campo}"):
                _lista_editavel(f"lista_{campo}", "+ Adicionar item e apertar Enter")
            dados[campo] = _lista_valores(f"lista_{campo}")
            continue

        key = f"f_{campo}"
        if key not in st.session_state:
            st.session_state[key] = gen.DADOS[tipo].get(campo, "")

        if campo in CAMPOS_GRANDES:
            dados[campo] = st.text_area(
                rotulo(campo), key=key, height=220,
                help="Deixe uma linha em branco entre parágrafos (vira parágrafo "
                     "de verdade no Word, não um texto corrido). Pra listar "
                     "materiais/equipamentos a fornecer, separe em um bloco à "
                     "parte com uma linha em branco antes e depois, e comece "
                     "cada linha com \"- \" (vira uma lista com marcadores).",
            )
        else:
            dados[campo] = st.text_area(rotulo(campo), key=key, height=90)

    st.divider()
    st.markdown(f"###### {rotulo('exclusoes')}")
    with st.container(border=True, key="card_exclusoes"):
        _lista_editavel("lista_exclusoes", "Adicionar outra exclusão e apertar Enter")
    dados["exclusoes"] = _lista_valores("lista_exclusoes")

    st.divider()
    st.markdown("###### Comercial")
    col_valor, col_pagamento = st.columns(2)
    with col_valor:
        key = "f_valor_total"
        st.session_state.setdefault(key, gen.DADOS[tipo].get("valor_total", ""))
        dados["valor_total"] = st.text_input(rotulo("valor_total"), key=key)
    with col_pagamento:
        opcoes_pagamento = list(gen.PAGAMENTOS.keys()) + ["Outro"]
        st.session_state.setdefault("sel_pagamento", opcoes_pagamento[0])
        st.selectbox(rotulo("condicoes_pagamento"), options=opcoes_pagamento, key="sel_pagamento")
        if st.session_state["sel_pagamento"] == "Outro":
            st.text_input("Condição de pagamento (texto livre)", key="f_pagamento_outro")
            dados["condicoes_pagamento"] = st.session_state.get("f_pagamento_outro", "")
        else:
            dados["condicoes_pagamento"] = gen.PAGAMENTOS[st.session_state["sel_pagamento"]]

st.divider()

gerar_pdf = st.checkbox(
    "Gerar também PDF (via Word, opcional)",
    key="gerar_pdf",
    help="Conveniência pra revisão rápida — usa o Microsoft Word instalado "
         "nesta máquina (automação COM, biblioteca docx2pdf). Feche "
         "diálogos abertos do Word antes de usar. O .docx continua sendo "
         "a fonte confiável pra enviar ao cliente.",
)

col_salvar, col_visualizar, col_gerar, _ = st.columns([1, 1, 1, 2])

if col_salvar.button(":material/save: Salvar no banco", use_container_width=True):
    novo_id = db.salvar(tipo, dados, editing_id)
    st.session_state.editing_id = novo_id
    st.success(f"Proposta salva (#{novo_id}).")
    st.rerun()

if col_visualizar.button(":material/visibility: Visualizar", use_container_width=True):
    faltantes = list(gen.validar_dados(tipo, dados))
    if "exclusoes" in campos_do_tipo(tipo) and not dados.get("exclusoes"):
        faltantes.append("exclusoes (lista vazia — adicione ao menos um item)")
    if faltantes:
        st.error(f"Campos obrigatórios em branco/faltando: {faltantes}")
    else:
        try:
            caminho_preview = gen.gerar_proposta(tipo, dados)
            html = _gerar_preview_html(caminho_preview)
            with st.container(border=True, key="card_preview"):
                st.caption(
                    "Preview aproximado (fundo, marca d'água e cores da marca "
                    "adicionados por cima do HTML do mammoth) — não reflete "
                    "100% da paginação/numeração de seção do Word real. Pra "
                    "conferir de verdade, gere o .docx (ou marque \"Gerar "
                    "também PDF\", se o Word estiver instalado nesta máquina)."
                )
                st.iframe(html, height=700)
        except Exception as e:
            st.error(f"Erro ao gerar preview: {e}")

if col_gerar.button(":material/description: Gerar .docx", type="primary", use_container_width=True):
    faltantes = list(gen.validar_dados(tipo, dados))
    if "exclusoes" in campos_do_tipo(tipo) and not dados.get("exclusoes"):
        faltantes.append("exclusoes (lista vazia — adicione ao menos um item)")
    if faltantes:
        st.error(f"Campos obrigatórios em branco/faltando: {faltantes}")
    else:
        try:
            caminho = gen.gerar_proposta(tipo, dados)
            # Guarda no session_state em vez de só na variável local: um
            # st.download_button também é um botão -> clicar nele já causa
            # um rerun do script, e nesse rerun o `if col_gerar.button(...)`
            # acima volta a ser False. Sem isso, o botão "Baixar PDF" (que
            # ficava DENTRO desse mesmo bloco) sumia assim que o usuário
            # clicava em "Baixar .docx" primeiro — mesmo o arquivo já
            # estando pronto no disco (bug real, José bateu nisso em
            # 2026-07-09). Com o caminho no session_state, os botões abaixo
            # sobrevivem a qualquer rerun até a próxima geração.
            st.session_state["ultimo_docx"] = str(caminho)
            st.session_state.pop("ultimo_pdf", None)
            if gerar_pdf:
                caminho_pdf = _converter_para_pdf(caminho)
                if caminho_pdf:
                    st.session_state["ultimo_pdf"] = str(caminho_pdf)
        except Exception as e:
            st.error(f"Erro ao gerar: {e}")

if st.session_state.get("ultimo_docx"):
    caminho = Path(st.session_state["ultimo_docx"])
    if caminho.exists():
        st.success(f"Gerado: {caminho.name}")
        with open(caminho, "rb") as fh:
            st.download_button(
                ":material/download: Baixar .docx",
                data=fh.read(),
                file_name=caminho.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx",
            )
        st.caption(f"Também salvo em: {caminho}")
        st.warning(
            "Confira no **Word** antes de enviar. Cheque: Ref.:/sistema, "
            "marcadores `{{ }}` soltos, espaçamento e quebra de página nos "
            "preços."
        )
        caminho_pdf_str = st.session_state.get("ultimo_pdf")
        if caminho_pdf_str:
            caminho_pdf = Path(caminho_pdf_str)
            if caminho_pdf.exists():
                with open(caminho_pdf, "rb") as fh:
                    st.download_button(
                        ":material/picture_as_pdf: Baixar PDF",
                        data=fh.read(),
                        file_name=caminho_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                        key="dl_pdf",
                    )
