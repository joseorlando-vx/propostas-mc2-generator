"""
============================================================
ROTEADOR MULTIPÁGINA — Gerador de Propostas MC2
============================================================
Este arquivo é só o ponto de entrada (o deploy do Streamlit Community
Cloud aponta pra ele — não renomear). O conteúdo de verdade vive em:

    - pagina_propostas.py          página "Propostas" (o app original, v5)

Usa st.navigation (padrão moderno, Streamlit >= 1.36) em vez da
auto-descoberta por pasta pages/ porque só assim os rótulos da sidebar
ficam com os nomes certos ("Propostas", não "app").

Rodar:  py -m streamlit run app.py   (a partir da pasta propostas-mc2)
============================================================
"""

import streamlit as st

_paginas = [
    st.Page("pagina_propostas.py", title="Propostas",
            icon=":material/description:", default=True),
]

st.navigation(_paginas).run()
