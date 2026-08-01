# Metrix-Lab

# Painel de Consumo de Materiais - Produzido por IA

Dashboard web para análise inteligente de planilhas de consumo de materiais. Basta importar um arquivo `.csv`, `.xlsx` ou `.xls` e o sistema gera automaticamente KPIs, gráficos, rankings e alertas de variação — tudo direto no navegador, sem backend e sem banco de dados.

## ✨ Funcionalidades

- **Upload de planilhas** — suporta CSV (com detecção automática de separador `;`/`,` e de encoding UTF‑8/Latin‑1), XLSX e XLS.
- **Reconhecimento automático de colunas** — identifica Produto, Quantidade, Data, Setor, Unidade, Valor etc. mesmo com nomes de coluna diferentes (ex: "Material" → Produto, "Destino" → Setor).
- **KPIs**: consumo total, produtos distintos, setores envolvidos, média mensal, média diária, maior consumidor.
- **Gráficos interativos**:
  - Evolução do consumo (linha, por mês)
  - Participação por setor (pizza)
  - Ranking de produtos — top 10 (barra horizontal)
  - Comparação entre setores (barra)
- **Alertas automáticos** de aumento/queda ≥15% entre meses consecutivos, por produto.
- **Filtros**: produto, setor, período (data inicial/final) e busca livre.
- **Tabela de movimentações** com os dados filtrados.
- **Validação de arquivo** — se colunas essenciais (Produto, Quantidade, Data) não forem encontradas, o sistema informa claramente o que está faltando.

## 🧱 Stack

Projeto **100% client-side** — HTML, CSS e JavaScript puro, sem framework e sem servidor:

- [SheetJS (xlsx)](https://sheetjs.com/) — leitura de planilhas Excel e CSV
- [Chart.js](https://www.chartjs.org/) — gráficos
- Fontes: Space Grotesk (títulos), Inter (corpo), IBM Plex Mono (números)

Ambas as bibliotecas são carregadas via CDN, então o projeto roda em qualquer ambiente com HTML/CSS/JS — incluindo sandboxes online como o OneCompiler — sem precisar instalar nada.

## 📂 Estrutura

```
.
├── index.html   # estrutura da página e áreas do dashboard
├── style.css    # tema visual (paleta escura, tipografia, layout)
└── script.js    # parsing de arquivos, cálculos e renderização dos gráficos
```

## 🚀 Como rodar

### Localmente
Basta abrir o `index.html` em qualquer navegador moderno (não precisa de servidor, build ou instalação de dependências).

### Em um sandbox online (ex: OneCompiler)
1. Crie um projeto do tipo **HTML/CSS/JS**.
2. Cole o conteúdo de `index.html`, `style.css` e `script.js` nos respectivos arquivos.
3. Rode o preview e importe sua planilha pelo botão **"Importar planilha"**.

## 📋 Colunas reconhecidas

| Campo interno | Nomes de coluna aceitos                       |
|---------------|-----------------------------------------------|
| Produto       | Material, Produto, Item, Descrição            |
| Código        | Código, Cod, SKU                              |
| Setor         | Destino, Setor, Departamento, Centro de Custo |
| Origem        | Origem                                        |
| Data          | Data, D. Lançamento, Data Movimento           |
| Quantidade    | Qte, Quantidade, Qtd, Qtde                    |
| Unidade       | Unidade, UN, Unid                             |
| Valor         | Valor                                         |
| Total         | Total, Valor Total                            |
| Responsável   | Responsável                                   |
|---------------|-----------------------------------------------|

> Produto, Quantidade e Data são obrigatórios — os demais campos são opcionais e, se ausentes, o dashboard segue funcionando sem eles.

## ⚠️ Limitações conhecidas

- Não há persistência entre sessões: ao recarregar a página, é preciso reimportar a planilha.
- Não há autenticação, múltiplos usuários ou banco de dados — é uma ferramenta de análise pontual, não um sistema multiusuário.
- Arquivos muito grandes (dezenas de milhares de linhas) podem impactar a performance, já que todo o processamento acontece no navegador.

## 🗺️ Possíveis evoluções futuras

- Exportação de relatórios em PDF/Excel
- Mais tipos de gráfico (heatmap, treemap, radar, waterfall)
- Projeção de consumo com regressão linear/médias móveis
- Modo claro/escuro alternável
