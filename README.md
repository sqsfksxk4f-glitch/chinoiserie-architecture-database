# Chinoiserie Architecture Database

A desktop pipeline for the cross-lingual retrieval, AI-assisted screening and deduplication, verification, and analysis of architectural cases. This repository accompanies the dataset paper *"A dataset of Chinoiserie Architecture cases perceived by publics in 69 countries"* and provides the software used to build, quality-check, and analyse that dataset.

## Overview

The tool automates a full research-data workflow:

1. **Cross-lingual retrieval** — translates core keywords into the official languages of target countries/regions and issues multi-language search queries across heterogeneous search engines.
2. **AI-assisted screening & deduplication** — filters relevant records with an LLM relevance rule, then performs URL-exact and title-similarity deduplication.
3. **Structured extraction** — extracts case metadata (name, location, function, designer, condition, etc.) via LLM.
4. **Multi-source verification** — verifies each field against at least two heterogeneous sources, with cascade fallback on failure.
5. **Analysis & reporting** — aggregates verified cases by dimension and renders statistical charts.

## Patent Notice

This software implements methods covered by a pending Chinese invention patent:

- **Title**: 一种建筑案例的跨语言检索、人工智能筛选去重、核验与分析的方法及系统 (*A method and system for cross-lingual retrieval, AI-assisted screening and deduplication, verification, and analysis of architectural cases*)
- **Application number**: 2026112536760
- **Status**: Published, not yet granted

## Features

- **Case Query (案例查询)** — search and filter the verified case database.
- **Data Organization (数据整理)** — field completion, deduplication, and verification workflows.
- **Data Analysis (数据分析)** — statistical aggregation and charts.
- **AI Chat Assistant** — toggle with `Ctrl+B` / `Cmd+B`.
- **Keyword generation & translation** — LLM-assisted keyword production.
- **API settings** — configure LLM channels via *Settings → API 配置*.

## Repository structure

```
.
├── main.py                 # Application entry point (Tkinter GUI)
├── config.py               # Configuration loader (LLM channels, language)
├── config.local.json       # Local API keys (NOT committed — gitignored)
├── database.py             # SQLite database access layer
├── ui/                     # GUI panels (search / organize / analysis / chat / settings)
├── utils/                  # AI extraction, verification, and helper modules
├── data/                   # SQLite working database
├── LICENSE                 # Apache-2.0
└── README.md
```

## Requirements

Python 3.9+ (tested with 3.11).

```bash
pip install -r requirements.txt
```

## Installation & running

1. Clone or download this repository.
2. Install dependencies (above).
3. Configure API keys via *Settings → API 配置* in the app (or edit `config.local.json`). API keys are **not** hard-coded in the source.
4. Run:

```bash
python main.py
```

## Configuration (API keys)

The software talks to LLMs through four configurable channels (OpenAI-compatible endpoints). Configure them in *Settings → API 配置*; they are persisted to `config.local.json`, which is excluded from version control. You must supply your own API keys and ensure the model name matches the endpoint you use.

## Data

The canonical dataset — **686 cases + 125 zero-case countries, 24 fields** — is published at ScienceDB (DOI: `10.57760/sciencedb.45005`) and included with the paper as `supplementary_table_case_database.xlsx` / `.csv`. The software's working SQLite database lives in `data/architecture.db`.

## Citation

If you use the dataset, cite the ScienceDB record:

> Zhao, X., Wang, Z. (2026). *A dataset of Chinoiserie Architecture cases perceived by publics in 69 countries*. ScienceDB. https://doi.org/10.57760/sciencedb.45005

## License

[Apache License 2.0](LICENSE). See the LICENSE file for full terms.

## Author

Xiaoxi Zhao
