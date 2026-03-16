# Marine fuels carbon-intensity workflow

This repository contains the R scripts used to reproduce the data-processing and figure-generation workflow for the marine fuels carbon-intensity study. The code is organised into two main components:

1. **Harmonisation scripts**, which clean, reconcile, and prepare analysis-ready datasets for HSFO, LPG, and LNG.
2. **Figure scripts**, which generate publication-ready figures from the harmonised outputs.

The workflow supports transparent and reproducible analysis of complex marine-fuel supply chains across refinery, voyage, terminal, shipping, and regasification boundaries.

## Scope

The repository currently covers three fuel pathways:

- **HSFO**
- **LPG**
- **LNG**

The harmonisation scripts process model outputs and source datasets into consistent tabular products. The plotting scripts use those harmonised datasets to generate the figures used in the manuscript.

## Repository structure

```text
.
├── harmonisation/
│   ├── hsfo_harmonisation.R
│   ├── lpg_harmonisation.R
│   └── lng_harmonisation.R
├── figures/
│   ├── marine_figures_shared_helpers.R
│   ├── fig3_hsfo.R
│   ├── fig4_lpg.R
│   ├── fig5_lng.R
│   └── fig7_comparison.R
├── data/
│   ├── raw/              # user-supplied raw inputs, not included here unless permitted
│   ├── interim/          # harmonised and intermediate outputs
│   └── figures/          # exported PNG, SVG, CSV, and Excel bundles
└── README.md
