# Commercial Proposal Generator (Streamlit + docxtpl)

A web app that generates commercial proposals as Word (.docx) files from a form, with AI-assisted text filling (Anthropic API). It was built for the proposal workflow of a building-automation company, replacing manual document editing.

> Public version: no letterhead templates, real proposals, client data or history from the private project.

## The problem

Each proposal was assembled by hand from an older one: copy, swap client, scope, exclusions and values. That caused placeholder errors, inconsistent wording and rework.

## What the app does

- 4 proposal types (simple, intermediate, service and amendment), each with its own template and placeholders
- Single-panel interface with a document preview (mammoth)
- "Fill with AI" button that suggests scope, exclusions and justification following a style guide
- .docx generation through `docxtpl`, plus optional PDF output (via Word, Windows only)
- Proposal history stored in a hosted database (Turso/libSQL)

## Stack

Python, Streamlit, docxtpl, python-docx, Anthropic SDK, mammoth, libSQL (Turso).

## Structure

```
app.py                  entry point (st.navigation)
pagina_propostas.py     proposal panel UI
etapa2_template.py      placeholder dictionary and per-type .docx generation
ia_preenchimento.py     Anthropic API call that suggests texts
db.py                   persistence (Turso)
templates_v5_painel_unico/   put your .docx templates here
```

## How to run

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=...
export TURSO_DATABASE_URL=libsql://your-database.turso.io
export TURSO_AUTH_TOKEN=...
streamlit run app.py
```

Create the `.docx` templates in `templates_v5_painel_unico/` (see that folder's README). In production, the variables live in the Streamlit Community Cloud "Secrets".

Note: code comments, UI text and proposal wording are in Brazilian Portuguese, because the proposals are written for Brazilian clients.

## My role

I designed and implemented the app, the proposal-type taxonomy and the AI integration, working with Claude Code and Codex as development assistants.

## License

MIT
