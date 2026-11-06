# Painel de Compras por Insumo (UAU)

    pip install -r requirements.txt
    streamlit run app.py

## Conexão com o UAU
1. Instale o **ODBC Driver 17 for SQL Server** na máquina que roda o app.
2. Copie `.streamlit/secrets.toml.exemplo` para `.streamlit/secrets.toml` e preencha servidor, banco, usuário e senha.
3. `usuario_uau` é o usuário do UAU aplicado na regra de permissão de obras (ObrUsr), como na consulta original.

No topo do app, **busque o insumo** digitando parte do código ou da descrição (a busca roda no InsumosGeral enquanto você digita; várias palavras = todas precisam aparecer). Marque os insumos na grade — ou use **Marcar todos** para marcar todos os resultados de uma vez —, escolha o **período da data de cotação** e clique em **Carregar**.
A consulta é a mesma da vwHistoricoItensDeCompra, com insumos e datas como parâmetros (sem concatenar texto no SQL).
Resultados ficam em cache por 10 min; a lista de insumos, por 1 h.

## Baixar PDF e Excel
No topo do painel:
- **Baixar PDF** — gera o dashboard em A4 (tema claro), respeitando os filtros aplicados no painel, sem cortar cards entre páginas.
- **Baixar Excel** — .xlsx com os dados da consulta: abas Sobre, Compras (dados tratados), Resumo por UF / insumo / obra / fornecedor e Consulta UAU (bruta).

## Sem banco
Se o `secrets.toml` não existir, o app usa a planilha `diesel.xlsx` ao lado do app.py (ou uma enviada pela barra lateral) e aplica os mesmos filtros.

## Arquivos
- `app.py` — consulta, tratamento dos dados e barra de filtros (Streamlit)
- `assets/painel.html` — painel (HTML/CSS/JS + ECharts)
- `assets/echarts.min.js`, `assets/brasil_uf.json` — embutidos no HTML (funciona sem internet)
