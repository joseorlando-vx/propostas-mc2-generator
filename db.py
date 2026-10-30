"""
============================================================
BANCO DE DADOS — Gerador de Propostas MC2 (v5 — Turso)
============================================================
Persistência das propostas salvas. Cada proposta é uma linha; os
marcadores `{{ }}` ficam num JSON na coluna `dados_json`, porque cada
tipo (A/B/C/D) tem um conjunto diferente de marcadores e não vale a
pena uma coluna por marcador.

v5 (2026-07-08): trocado de SQLite LOCAL (arquivo `propostas.db`) pra
Turso (libsql hospedado), porque o Streamlit Community Cloud tem
sistema de arquivos EFÊMERO — um `propostas.db` local seria apagado a
cada reinício/redeploy do app (ver bugs_conhecidos.md, item 11 do
handoff "Mockup do painel único v2/v3"). A API pública deste módulo
(init_db/listar/obter/salvar/excluir) não mudou — só a conexão por
baixo. Requer as variáveis de ambiente:
    TURSO_DATABASE_URL  (ex: "libsql://seu-banco.turso.io")
    TURSO_AUTH_TOKEN
Em produção (Streamlit Community Cloud) ficam em "Secrets"; localmente,
defina como variável de ambiente antes de rodar o Streamlit. Não tem
fallback pra SQLite local — é troca, não modo dual (ferramenta interna,
prioriza simplicidade).

Ferramenta interna — prioriza simplicidade. Sem ORM, sem migração: o
schema é criado no init_db() se ainda não existir.
============================================================
"""

import json
import os
from datetime import datetime

import libsql_client


def _conectar() -> libsql_client.ClientSync:
    url = os.environ.get("TURSO_DATABASE_URL")
    token = os.environ.get("TURSO_AUTH_TOKEN")
    if not url:
        raise RuntimeError(
            "TURSO_DATABASE_URL não configurada. Crie um banco em turso.tech "
            "e defina TURSO_DATABASE_URL (e TURSO_AUTH_TOKEN) como variável "
            "de ambiente antes de rodar o Streamlit — em produção, como "
            "Secret do Streamlit Community Cloud."
        )
    # O painel do Turso mostra a URL como "libsql://...", que o libsql_client
    # mapeia pra WebSocket (wss://) por padrão. Testado ao vivo nesta máquina
    # (2026-07-09): o handshake wss:// falhou (WSServerHandshakeError 400),
    # mas https:// (Hrana sobre HTTP) funcionou de primeira contra o MESMO
    # banco/token — por isso convertemos aqui, pra funcionar direto com a URL
    # que o José copia do painel, sem precisar editar manualmente.
    if url.startswith("libsql://"):
        url = "https://" + url[len("libsql://"):]
    return libsql_client.create_client_sync(url, auth_token=token)


def init_db() -> None:
    """Cria a tabela se não existir. Idempotente — pode chamar sempre."""
    with _conectar() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS propostas (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo            TEXT NOT NULL,           -- A / B / C / D
                numero_proposta TEXT,                   -- ex: "PS 26_080" (só rótulo, não é chave)
                cliente         TEXT,                   -- cliente ou titulo_aditivo (só pra listar)
                dados_json      TEXT NOT NULL,          -- dict de marcadores {{ }} serializado
                criado_em       TEXT NOT NULL,
                atualizado_em   TEXT NOT NULL
            )
            """
        )


def listar() -> list[dict]:
    """Retorna todas as propostas salvas (mais recente primeiro), sem o JSON."""
    with _conectar() as conn:
        resultado = conn.execute(
            "SELECT id, tipo, numero_proposta, cliente, atualizado_em "
            "FROM propostas ORDER BY atualizado_em DESC"
        )
    return [row.asdict() for row in resultado.rows]


def obter(proposta_id: int) -> dict | None:
    """Retorna uma proposta completa (com os dados desserializados) ou None."""
    with _conectar() as conn:
        resultado = conn.execute("SELECT * FROM propostas WHERE id = ?", [proposta_id])
    if not resultado.rows:
        return None
    registro = resultado.rows[0].asdict()
    registro["dados"] = json.loads(registro.pop("dados_json"))
    return registro


def salvar(tipo: str, dados: dict, proposta_id: int | None = None) -> int:
    """
    Insere (proposta_id=None) ou atualiza uma proposta.
    Retorna o id. `cliente` é derivado dos dados só pra facilitar a listagem.
    """
    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    numero = dados.get("numero_proposta", "")
    # Tipo D usa titulo_aditivo em vez de cliente
    cliente = dados.get("cliente") or dados.get("titulo_aditivo") or "(sem cliente)"
    dados_json = json.dumps(dados, ensure_ascii=False)

    with _conectar() as conn:
        if proposta_id is None:
            resultado = conn.execute(
                "INSERT INTO propostas "
                "(tipo, numero_proposta, cliente, dados_json, criado_em, atualizado_em) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                [tipo, numero, cliente, dados_json, agora, agora],
            )
            return resultado.last_insert_rowid
        else:
            conn.execute(
                "UPDATE propostas SET tipo=?, numero_proposta=?, cliente=?, "
                "dados_json=?, atualizado_em=? WHERE id=?",
                [tipo, numero, cliente, dados_json, agora, proposta_id],
            )
            return proposta_id


def excluir(proposta_id: int) -> None:
    with _conectar() as conn:
        conn.execute("DELETE FROM propostas WHERE id = ?", [proposta_id])
