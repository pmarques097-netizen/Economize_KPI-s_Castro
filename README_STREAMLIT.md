# Rede Economize KPI — Viewer Streamlit

Este projeto é somente de exibição.

- Não conecta ao PostgreSQL.
- Não precisa de OpenVPN.
- Lê `data/kpis_mensal.sqlite`.
- A atualização é feita pelo aplicativo local.
- Usuários compradores ficam restritos às telas de premiação.

## Deploy

Main file path: `app.py`

O SQLite possui mais de 25 MB. Portanto, para publicar no GitHub use GitHub Desktop
ou `git push`; o upload pelo navegador do GitHub pode recusar arquivos desse tamanho.
