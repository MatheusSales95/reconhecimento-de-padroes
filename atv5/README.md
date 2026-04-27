# Reconhecimento de Padrões — Agrupamento (BSAS / Parzen / KNN)

Projeto LaTeX pronto para Overleaf. Compila em `pdfLaTeX` com a classe `article`,
em português (Babel `brazil`).

## Estrutura

```
.
├── main.tex            ← documento principal (article)
├── run_analysis.py     ← script Python que regenera todas as figuras (PDF)
├── results.csv         ← métricas finais (ARI, NMI, acc*) por método/dataset
├── figs/               ← figuras já geradas, em PDF vetorial
│   ├── dados_pca.pdf
│   ├── bsas_curve_iris.pdf
│   ├── bsas_curve_wine.pdf
│   ├── q1_iris_clusters.pdf
│   ├── q1_wine_clusters.pdf
│   ├── parzen_sensitivity.pdf
│   ├── parzen_iris.pdf
│   ├── parzen_wine.pdf
│   ├── knn_iris.pdf
│   ├── knn_wine.pdf
│   ├── cm_iris.pdf
│   └── cm_wine.pdf
└── README.md
```

## Como usar no Overleaf

1. Crie um novo projeto vazio em https://www.overleaf.com.
2. *Menu* → *Upload Project* (arraste o ZIP inteiro), **ou** crie o projeto e
   arraste todos os arquivos listados acima preservando a pasta `figs/`.
3. Verifique se *Compiler* está em `pdfLaTeX` (Menu → *Settings*).
4. Compile (`Recompile`). O PDF de 10 páginas deve ser gerado sem warnings.

> O script `run_analysis.py` **não precisa rodar no Overleaf** — todas as
> figuras já estão em `figs/` em formato PDF vetorial. Ele está incluído
> apenas para reprodutibilidade caso queira reexecutar localmente
> (`python run_analysis.py`, requer `numpy`, `scikit-learn`, `scipy`,
> `matplotlib`, `pandas`).

## Pacotes LaTeX usados

`inputenc`, `fontenc`, `babel (brazil)`, `geometry`, `amsmath`, `amssymb`,
`graphicx`, `float`, `booktabs`, `multirow`, `caption`, `subcaption`,
`xcolor`, `listings`, `hyperref`, `url`, `microtype`, `enumitem`,
`algorithm`, `algpseudocode`. Todos disponíveis na imagem padrão do Overleaf.

## Resultados

| Dataset | Método           | K  | acc*  | ARI   | NMI   |
|---------|------------------|----|-------|-------|-------|
| Iris    | BSAS → Ward      | 2  | 0,660 | 0,544 | 0,692 |
| Iris    | Parzen mode-seek | 2  | 0,667 | 0,568 | 0,734 |
| Iris    | KNN  mkNN        | 2  | 0,667 | 0,568 | 0,734 |
| Wine    | BSAS → Ward      | 3  | **0,927** | **0,790** | **0,786** |
| Wine    | Parzen mode-seek | 12 | 0,899 | 0,781 | 0,752 |
| Wine    | KNN  mkNN        | 4  | 0,904 | 0,761 | 0,762 |

`acc*` = acurácia após alinhamento ótimo dos rótulos via algoritmo Hungarian.
