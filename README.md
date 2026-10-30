# Gerador de Propostas Comerciais (Streamlit + docxtpl)

Aplicação web que gera propostas comerciais em Word (.docx) a partir de formulário, com preenchimento assistido por IA (API da Anthropic). Foi construída para o fluxo de propostas de uma empresa de automação predial, em substituição ao preenchimento manual de documentos.

> Versão pública: sem templates com timbrado, propostas reais, dados de clientes ou histórico do projeto privado.

## O problema

Cada proposta era montada à mão a partir de uma proposta antiga: copiar, trocar cliente, escopo, exclusões e valores. Isso gerava erros de marcador, inconsistência de linguagem e retrabalho.

## O que a aplicação faz

- 4 tipos de proposta (simples, intermediária, serviço e aditivo), cada um com seu template e marcadores
- Interface em painel único com pré-visualização do documento (mammoth)
- Botão "Preencher com IA" que sugere escopo, exclusões e justificativa seguindo um manual de estilo
- Geração do .docx via `docxtpl` e opção de gerar PDF (via Word, só Windows)
- Histórico de propostas salvo em banco hospedado (Turso/libSQL)

## Stack

Python, Streamlit, docxtpl, python-docx, Anthropic SDK, mammoth, libSQL (Turso).

## Estrutura

```
app.py                  ponto de entrada (st.navigation)
pagina_propostas.py     interface do painel de propostas
etapa2_template.py      dicionário de marcadores e geração do .docx por tipo
ia_preenchimento.py     chamada à API da Anthropic para sugerir textos
db.py                   persistência (Turso)
templates_v5_painel_unico/   coloque aqui seus templates .docx
```

## Como rodar

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=...
export TURSO_DATABASE_URL=libsql://seu-banco.turso.io
export TURSO_AUTH_TOKEN=...
streamlit run app.py
```

Crie os templates `.docx` em `templates_v5_painel_unico/` (veja o README da pasta). Em produção, as variáveis ficam em "Secrets" do Streamlit Community Cloud.

## Meu papel

Desenhei e implementei a aplicação, a taxonomia de tipos de proposta e a integração com IA, trabalhando com Claude Code e Codex como assistentes de desenvolvimento.

## Licença

MIT
