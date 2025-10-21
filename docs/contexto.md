# Dashboard de Stakeholders – Contexto do Projeto

## Visão geral
- Objetivo: construir um dashboard HTML/CSS estático para analisar entrevistas de stakeholders da Onfly a partir da planilha `Base_Stakeholders_DS_v4.xlsx`.
- Stack atual: HTML puro, CSS customizado, Tailwind via CDN, JavaScript vanilla (modularizado em `app.js`).
- Dados: extraídos da planilha para `data/stakeholders.json` e `data/insights.json` usando o script Ruby `scripts/extract_data.rb`.

## Estrutura de páginas (sidebar)
1. **Visão geral**
   - KPIs agregados (stakeholders, insights, tempo médio).
   - Gráficos básicos: distribuição por time (barras horizontais), tipos de insight, insight mais frequente.
   - Filtro global de time (aplica a esta e a outras páginas, exceto Detalhamento/Panorama).

2. **Dimensão**
   - Filtro extra de tipo (ex.: Dor, Observacao etc.).
   - Gráfico interativo de categorias (barras) filtrando demais componentes.
   - Donut por time, tabela de menções, tags relacionadas.

3. **Expectativas & Engajamento**
   - Métricas: contagem de stakeholders/insights de expectativa e engajamento.
   - Gráfico de categorias (tipos Impacto_Esperado, Necessidade, Sugestao).
   - Donut de engajamento por time.
  - Tabela com todas as expectativas e tags agregadas.
   - Ranking dos stakeholders mais engajados (usa somente insights `Engajamento`).

4. **Detalhamento**
   - Filtros combináveis (tipo/categoria/time/stakeholder) independentes do filtro global.
   - Tabela completa de insights (seleção destaca perfil) com:
     - Perfil do stakeholder (avatar, dados, stats, tags frequentes).
     - Timeline mensal do conjunto filtrado.
     - Tags recorrentes no contexto atual.

5. **Panorama final**
   - KPIs executivos (totais, percentuais de expectativa/engajamento).
   - Temas críticos por tipo (resumo das categorias de Dor, Expectativa, Engajamento).
   - Recomendações automáticas baseadas em tags.
   - Mapa de influência: todos os stakeholders ordenados pelo volume de menções (diferenciando dor/expectativa/engajamento).
   - Sem uso do filtro global de time (visão macro).

## Arquivos importantes
- `index.html`: estrutura das páginas e componentes.
- `styles.css`: estilos globais, grids e componentes customizados.
- `app.js`: lógica de carregamento dos dados, filtros, rendering dos gráficos/tabelas.
- `data/*.json`: dados processados da planilha.
- `scripts/extract_data.rb`: script Ruby para atualizar os JSON a partir da planilha.
- `assets/logo onfly.svg`: logo oficial usada na sidebar.

## Convenções de dados
- Expectativas: insights com tipos `Impacto_Esperado`, `Necessidade`, `Sugestao`.
- Engajamento: insights com tipo `Engajamento`.
- Dores: insights com tipo `Dor`.
- Timeline utiliza `Data_da_Entrevista` (convertida para ISO durante a extração).

## Processo para atualizar dados
1. Atualizar a planilha `Base_Stakeholders_DS_v4.xlsx`.
2. Rodar `ruby scripts/extract_data.rb` (gera `data/stakeholders.json` e `data/insights.json`).
3. Recarregar o dashboard no navegador.

## Próximos passos potenciais
- Incluir estado de ações (follow-up) quando houver dados.
- Adicionar modos de exportar gráficos (PNG/PDF) se necessário.
- Criar testes/validações automatizadas para o script de extração.
