"""
Engineering report generator for the Chinoiserie Architecture Beyond China Database.
Produces a Word (.docx) document suitable for journal supplementary materials.
Supports bilingual output: English (for HSSC submission) and Chinese.

Usage:
    python generate_report.py          # uses LANG setting below
    LANG=en python generate_report.py  # force English
    LANG=zh python generate_report.py  # force Chinese
"""

import os
import sys
from datetime import datetime

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ---------------------------------------------------------------------------
# Language configuration
#   Set to "en" for English (HSSC submission), "zh" for Chinese
#   Override with environment variable LANG=en or LANG=zh
# ---------------------------------------------------------------------------
_env_lang = os.environ.get("REPORT_LANG", "")
if _env_lang in ("en", "zh"):
    LANG = _env_lang
else:
    # Detect from system locale
    _sys_lang = os.environ.get("LANG", "")
    LANG = "zh" if _sys_lang.startswith("zh") else "en"

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(
    PROJECT_DIR,
    "Chinoiserie_Architecture_Database_Engineering_Report.docx" if LANG == "en"
    else "海外中式建筑数据库_工程报告.docx"
)

# ---------------------------------------------------------------------------
# Font configuration
# ---------------------------------------------------------------------------
def _available_font(name, fallback):
    candidates = [
        f"/System/Library/Fonts/{name}.ttc",
        f"/System/Library/Fonts/{name}.ttf",
        f"/Library/Fonts/{name}.ttc",
        f"/Library/Fonts/{name}.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return name
    return fallback

HEADING_FONT = _available_font("PingFang SC", "Arial")
BODY_FONT = _available_font("PingFang SC", "Arial")
CODE_FONT = "Courier New"

# ---------------------------------------------------------------------------
# Bilingual text dictionary
#   All user-facing strings are defined here.  Methods reference keys via
#   self._t("key").  To switch language, change LANG at the top of the file.
#   "Chinoiserie Architecture" is the canonical English term used throughout
#   the HSSC submission.
# ---------------------------------------------------------------------------
T = {
    # ── Cover ──────────────────────────────────────────────────────────
    "cover_title_main": {
        "zh": "海外中式建筑数据库",
        "en": "Chinoiserie Architecture Beyond China Database",
    },
    "cover_title_sub": {
        "zh": "工程设计与技术报告",
        "en": "Software Engineering Report",
    },
    "cover_version": {
        "zh": "版本：1.0",
        "en": "Version: 1.0",
    },
    "cover_date": {
        "zh": "日期：{date}",
        "en": "Date: {date}",
    },
    "cover_env": {
        "zh": "开发环境：Python 3.11 / macOS / VS Code",
        "en": "Environment: Python 3.11 / macOS / VS Code",
    },

    # ── TOC ────────────────────────────────────────────────────────────
    "toc_title": {"zh": "目  录", "en": "Table of Contents"},
    "toc_hint": {
        "zh": "（在 Word 中右键此处 → 更新域 → 更新整个目录）",
        "en": "(In Word, right-click here → Update Field → Update entire table)",
    },

    # ── Chapter titles ─────────────────────────────────────────────────
    "ch1_title": {"zh": "第1章  项目概述", "en": "1. Project Overview"},
    "ch2_title": {"zh": "第3章  系统架构设计", "en": "3. System Architecture"},
    "ch3_title": {"zh": "第4章  数据库设计", "en": "4. Database Design"},
    "ch4_title": {"zh": "第5章  功能模块设计", "en": "5. Functional Module Design"},
    "ch5_title": {"zh": "第6章  核心技术实现", "en": "6. Core Technical Implementation"},
    "ch6_title": {"zh": "第7章  部署与运维", "en": "7. Deployment and Operations"},
    "ch7_title": {"zh": "第8章  总结与展望", "en": "8. Summary and Outlook"},

    # ── Appendix titles ────────────────────────────────────────────────
    "appendix_a": {"zh": "附录A  完整数据字典", "en": "Appendix A  Complete Data Dictionary"},
    "appendix_b": {"zh": "附录B  API集成文档", "en": "Appendix B  API Integration Reference"},
    "appendix_c": {"zh": "附录C  快捷命令与快捷键参考", "en": "Appendix C  Quick Commands and Keyboard Shortcuts"},
    "appendix_d": {"zh": "附录D  数据库迁移历史", "en": "Appendix D  Database Migration History"},

    # ── Ch1 ────────────────────────────────────────────────────────────
    "1.1": {"zh": "1.1  项目背景", "en": "1.1  Background"},
    "1.2": {"zh": "1.2  项目目标", "en": "1.2  Objectives"},
    "1.3": {"zh": "1.3  系统功能概览", "en": "1.3  System Function Overview"},
    "1.4": {"zh": "1.4  技术栈概览", "en": "1.4  Technology Stack"},

    "bg_p1": {
        "zh": (
            "随着全球化进程的深入，中式建筑在世界各地的传播与演变已成为建筑学、"
            "文化研究等领域的重要课题。海外中式建筑不仅是中国文化输出的物质载体，"
            "也反映了不同文化语境下建筑形态的适应与创新。然而，相关研究长期面临数据"
            "分散、信息碎片化、真伪难辨等挑战。研究人员需要耗费大量时间在搜索引擎、"
            "学术数据库和社交平台上检索、筛选和核验信息。"
        ),
        "en": (
            "The global diffusion of Chinoiserie Architecture has become an important "
            "subject in architectural and cultural studies. Buildings perceived and "
            "labelled as Chinese-style outside China are material carriers of cultural "
            "exchange, reflecting the adaptation and innovation of architectural forms "
            "across different cultural contexts. However, research in this area has long "
            "been constrained by fragmented data, scattered information, and the difficulty "
            "of verifying authenticity. Researchers spend considerable time searching, "
            "screening, and verifying information across search engines, academic databases, "
            "and social media platforms."
        ),
    },
    "bg_p2": {
        "zh": (
            "针对上述痛点，本软件\"海外中式建筑数据库\"应运而生。它是一套面向建筑学"
            "研究者的一站式数据采集、整理、核验与分析工具，将AI大模型能力与传统数据"
            "库管理深度整合，旨在提升海外中式建筑研究的效率与数据质量。"
        ),
        "en": (
            "To address these challenges, we developed the Global Chinoiserie Architecture "
            "Database — an integrated data-collection, curation, verification, and analysis "
            "tool for architectural researchers. It combines large-language-model capabilities "
            "with traditional database management to improve the efficiency and data quality "
            "of research on Chinoiserie Architecture beyond China."
        ),
    },

    "obj_1": {
        "zh": "构建一个结构化的海外中式建筑案例数据库，支持灵活的字段扩展；",
        "en": "Build a structured database of Chinoiserie Architecture cases worldwide, with flexible field extension;",
    },
    "obj_2": {
        "zh": "实现多引擎、多语言的自动化网络搜索与信息采集；",
        "en": "Implement multi-engine, multilingual automated web search and information collection;",
    },
    "obj_3": {
        "zh": "集成AI大模型（DeepSeek、Gemini）实现建筑案例的自动提取、真实性核验与字段补全；",
        "en": "Integrate LLMs (DeepSeek, Gemini) for automated case extraction, authenticity verification, and field completion;",
    },
    "obj_4": {
        "zh": "提供交互式工作流引擎，将繁琐的多步骤操作自动化为\"一键\"执行；",
        "en": "Provide an interactive workflow controller that automates multi-step operations into one-click execution;",
    },
    "obj_5": {
        "zh": "支持语音指令与自然语言交互，降低软件使用门槛；",
        "en": "Support voice commands and natural-language interaction to reduce the barrier to entry;",
    },
    "obj_6": {
        "zh": "提供数据统计分析与可视化导出功能，支持学术研究与论文撰写。",
        "en": "Provide statistical analysis and visualisation export to support academic research and paper writing.",
    },

    "func_overview": {
        "zh": "系统围绕\"搜索→提取→整理→核验→入库→分析\"的数据管道设计，主要功能模块包括：",
        "en": "The system is designed around a data pipeline of Search → Extract → Curate → Verify → Ingest → Analyse. The main functional modules are:",
    },

    "func_1_name": {"zh": "案例查询", "en": "Case Search"},
    "func_1_desc": {
        "zh": "网页关键词搜索、AI案例提取、本地数据库查询、文献库管理",
        "en": "Web keyword search, AI case extraction, local database query, literature management",
    },
    "func_2_name": {"zh": "数据整理", "en": "Data Curation"},
    "func_2_desc": {
        "zh": "原始数据导入导出、字段级真实性核验（百度/Google/DeepSeek/Gemini四引擎）、查重合并、数据清洗、正式案例库管理",
        "en": "Raw data import/export, field-level authenticity verification (four engines: Baidu/Google/DeepSeek/Gemini), deduplication and merging, data cleaning, verified-case database management",
    },
    "func_3_name": {"zh": "数据分析", "en": "Data Analysis"},
    "func_3_desc": {
        "zh": "按建筑类型/国家/建造年代/保护等级维度的统计图表、全量数据导出",
        "en": "Statistical charts by building type, country, founding period, and protection status; full-data export",
    },
    "func_4_name": {"zh": "聊天助手", "en": "Chat Assistant"},
    "func_4_desc": {
        "zh": "24条正则快捷命令、50+种AI动作、离线语音识别、自然语言交互",
        "en": "24 regex-based quick commands, 50+ AI actions, offline speech recognition, natural-language interaction",
    },
    "func_5_name": {"zh": "工作流引擎", "en": "Workflow Controller"},
    "func_5_desc": {
        "zh": "一键搜索（195个国家×3种语言×13个关键词）、一键DS/G核验（名称核验→字段补全→图片搜索）、一键核验（DS→G级联回退）",
        "en": "One-click search (195 countries × 3 languages × 13 keywords); one-click DS/G verification (name verify → field complete → image search); one-click cascade verification (DS → G fallback)",
    },

    "tech_tier": {"zh": "层次", "en": "Tier"},
    "tech_choice": {"zh": "技术选型", "en": "Technology"},
    "tech_note": {"zh": "说明", "en": "Notes"},
    "tech_lang": {"zh": "编程语言", "en": "Language"},
    "tech_lang_val": {"zh": "Python 3.11", "en": "Python 3.11"},
    "tech_lang_desc": {"zh": "跨平台、丰富的科学计算与AI生态", "en": "Cross-platform; rich ecosystem for scientific computing and AI"},
    "tech_gui": {"zh": "GUI框架", "en": "GUI framework"},
    "tech_gui_val": {"zh": "Tkinter (标准库)", "en": "Tkinter (standard library)"},
    "tech_gui_desc": {"zh": "零额外依赖、轻量级桌面应用", "en": "Zero extra dependencies; lightweight desktop application"},
    "tech_db": {"zh": "数据库", "en": "Database"},
    "tech_db_val": {"zh": "SQLite3", "en": "SQLite3"},
    "tech_db_desc": {"zh": "嵌入式数据库，无需独立服务，单文件管理", "en": "Embedded database; no standalone server needed; single-file management"},
    "tech_ai": {"zh": "AI模型", "en": "AI models"},
    "tech_ai_val": {"zh": "DeepSeek-Chat / Gemini-2.5-Flash", "en": "DeepSeek-Chat / Gemini-2.5-Flash"},
    "tech_ai_desc": {"zh": "通过OpenAI兼容API调用", "en": "Accessed via OpenAI-compatible API"},
    "tech_voice": {"zh": "语音识别", "en": "Speech recognition"},
    "tech_voice_val": {"zh": "Vosk (离线)", "en": "Vosk (offline)"},
    "tech_voice_desc": {"zh": "无需网络，支持中文离线语音转文字", "en": "Offline Chinese speech-to-text; no network required"},
    "tech_search": {"zh": "搜索引擎", "en": "Search engines"},
    "tech_search_val": {"zh": "Serper.dev (Google) / 百度千帆 / 浏览器直搜", "en": "Serper.dev (Google) / Baidu Qianfan / browser direct"},
    "tech_search_desc": {"zh": "多源采集保障信息覆盖面", "en": "Multi-source collection ensures coverage breadth"},
    "tech_pdf": {"zh": "文档解析", "en": "Document parsing"},
    "tech_pdf_val": {"zh": "PyMuPDF (fitz)", "en": "PyMuPDF (fitz)"},
    "tech_pdf_desc": {"zh": "支持中文PDF文本提取与PUA字符清理", "en": "Chinese PDF text extraction with PUA character cleaning"},
    "tech_analysis": {"zh": "数据分析", "en": "Data analysis"},
    "tech_analysis_val": {"zh": "pandas / matplotlib / openpyxl", "en": "pandas / matplotlib / openpyxl"},
    "tech_analysis_desc": {"zh": "数据处理、图表渲染与Excel导出", "en": "Data processing, chart rendering, and Excel export"},
    "tech_docgen": {"zh": "文档生成", "en": "Document generation"},
    "tech_docgen_val": {"zh": "python-docx", "en": "python-docx"},
    "tech_docgen_desc": {"zh": "本报告的自动生成", "en": "Automated generation of this report"},

    # ── Ch2 ────────────────────────────────────────────────────────────
    "2.1": {"zh": "3.1  整体架构", "en": "2.1  Overall Architecture"},
    "2.2": {"zh": "3.2  模块依赖关系", "en": "2.2  Module Dependencies"},
    "2.3": {"zh": "3.3  技术选型说明", "en": "2.3  Technology Choices"},
    "2.3.1": {"zh": "3.3.1  Tkinter GUI框架", "en": "2.3.1  Tkinter GUI Framework"},
    "2.3.2": {"zh": "3.3.2  SQLite3数据库", "en": "2.3.2  SQLite3 Database"},
    "2.3.3": {"zh": "3.3.3  多AI模型集成", "en": "2.3.3  Multi-LLM Integration"},
    "2.4": {"zh": "3.4  文件组织结构", "en": "2.4  File Organisation"},

    "arch_p1": {
        "zh": (
            "系统采用桌面应用的经典分层架构，自底向上分为数据层、业务逻辑层、UI表现层和AI服务层。"
            "各层之间通过明确的接口调用，实现了关注点分离和模块化设计。"
        ),
        "en": (
            "The system adopts a classic layered architecture for desktop applications, "
            "organised from bottom to top into data layer, business-logic layer, UI "
            "presentation layer, and AI service layer. Each layer communicates through "
            "well-defined interfaces, achieving separation of concerns and modular design."
        ),
    },
    "arch_tier_ui": {"zh": "UI表现层", "en": "UI Presentation Layer"},
    "arch_tier_logic": {"zh": "业务逻辑层", "en": "Business Logic Layer"},
    "arch_tier_ai": {"zh": "AI服务层", "en": "AI Service Layer"},
    "arch_tier_data": {"zh": "数据层", "en": "Data Layer"},
    "arch_tier_util": {"zh": "工具层", "en": "Utility Layer"},
    "arch_col_module": {"zh": "组成模块", "en": "Modules"},
    "arch_col_role": {"zh": "职责", "en": "Responsibilities"},
    "arch_ui_desc": {
        "zh": "用户交互界面，包含主窗口、三个功能标签页和聊天助手面板",
        "en": "User interface: main window, three function tabs, and chat assistant panel",
    },
    "arch_logic_desc": {
        "zh": "工作流编排、关键词生成、数据导入映射、图片管理",
        "en": "Workflow orchestration, keyword generation, data import mapping, image management",
    },
    "arch_ai_desc": {
        "zh": "AI案例提取、真实性核验、字段补全、智能搜索",
        "en": "AI case extraction, authenticity verification, field completion, intelligent search",
    },
    "arch_data_desc": {
        "zh": "SQLite3数据库操作、表结构定义、迁移管理、撤销/重做",
        "en": "SQLite3 database operations, schema definition, migration management, undo/redo",
    },
    "arch_util_desc": {
        "zh": "PDF解析、数据清洗、多语言翻译",
        "en": "PDF parsing, data cleaning, multilingual translation",
    },

    "dep_p1": {
        "zh": (
            "main.py 作为程序入口，创建 MainWindow 实例并启动 Tkinter 主循环。"
            "MainWindow 聚合了 SearchTab、OrganizeTab、AnalysisTab 和 ChatAssistant 四个"
            "核心UI组件，它们共享同一个 database 模块提供的数据库连接。"
            "ChatAssistant 通过 CommandDispatcher 将用户指令路由到具体的UI操作或工作流引擎。"
            "工作流引擎（workflows.py）调用 utils/ai_extractor.py 中的AI函数"
            "（extract_case_names、verify_field_deepseek、auto_complete_with_gemini 等），"
            "这些AI函数再通过 OpenAI SDK 与 DeepSeek API 或 Gemini（via Yunwu）进行通信。"
        ),
        "en": (
            "main.py serves as the program entry point, creating a MainWindow instance and "
            "starting the Tkinter main loop. MainWindow aggregates four core UI components "
            "(SearchTab, OrganizeTab, AnalysisTab, and ChatAssistant), which share a single "
            "database connection provided by the database module. ChatAssistant routes user "
            "commands to specific UI operations or workflow-controller actions via "
            "CommandDispatcher. The workflow controller (workflows.py) calls AI functions "
            "in utils/ai_extractor.py (extract_case_names, verify_field_deepseek, "
            "auto_complete_with_gemini, etc.), which communicate with the DeepSeek API "
            "or Gemini (via Yunwu) through the OpenAI SDK."
        ),
    },

    "tkinter_p1": {
        "zh": (
            "选择Tkinter而非PyQt或Web前端的主要考虑：（1）Tkinter是Python标准库组件，"
            "无需额外安装，部署简单；（2）对于数据管理类桌面应用，Tkinter的功能完全满足需求；"
            "（3）配合ttk主题组件可获得现代化的外观。缺点是缺乏高级控件和动画效果，"
            "但对于以数据操作为核心的工具软件而言影响有限。"
        ),
        "en": (
            "Tkinter was chosen over PyQt or web-based frontends for three reasons: "
            "(1) it is a Python standard-library component requiring no extra installation; "
            "(2) its feature set is fully adequate for a data-management desktop application; "
            "(3) ttk themed widgets provide a modern appearance. The lack of advanced "
            "controls and animation effects has limited impact on a data-centric tool."
        ),
    },
    "sqlite_p1": {
        "zh": (
            "SQLite3的选择理由：（1）嵌入式架构，数据库即单个.db文件，备份和迁移极为便捷；"
            "（2）无需独立数据库服务进程，适合个人研究者使用场景；"
            "（3）支持完整的SQL标准（包括ALTER TABLE动态加列），满足灵活的字段扩展需求；"
            "（4）通过SQLite的WAL模式可在合理范围内支持并发读写。"
            "当前数据库文件约972 MB，包含17张表和811条正式案例记录，运行流畅。"
        ),
        "en": (
            "SQLite3 was chosen because: (1) its embedded architecture stores the entire "
            "database as a single .db file, making backup and migration trivial; "
            "(2) no standalone database server process is needed — ideal for individual "
            "researchers; (3) it supports the full SQL standard (including ALTER TABLE "
            "ADD COLUMN) for flexible field extension; (4) WAL mode provides reasonable "
            "concurrent read-write support. The current database file is approximately "
            "972 MB, containing 17 tables and 811 verified case records, and runs smoothly."
        ),
    },
    "multi_ai_p1": {
        "zh": (
            "系统同时集成了DeepSeek和Gemini两个大模型。DeepSeek作为主力模型处理中文语境下的"
            "建筑信息提取和核验，Gemini作为辅助模型提供多语言能力和结果交叉验证。"
            "两个模型均通过OpenAI兼容API协议调用，底层使用同一套 openai SDK，"
            "便于统一管理和切换。Gemini通过云屋（yunwu.ai）代理访问，解决了国内网络环境"
            "的访问限制问题。"
        ),
        "en": (
            "The system integrates two large language models: DeepSeek serves as the primary "
            "model for architectural information extraction and verification in Chinese-language "
            "contexts; Gemini provides multilingual capability and cross-validation as a "
            "secondary model. Both are accessed through an OpenAI-compatible API protocol "
            "using the same openai SDK, simplifying management and model switching. Gemini "
            "access is proxied through yunwu.ai to address network-access constraints."
        ),
    },

    "file_tree": {
        "zh": (
            "海外中式建筑数据库/编程/\n"
            "├── main.py                  # 程序入口，MainWindow主窗口（109行）\n"
            "├── database.py              # 数据库层（3,157行），17张表CRUD + 迁移v12 + 撤销/重做\n"
            "├── generate_report.py       # 本报告的自动生成脚本\n"
            "├── refill_reasons.py        # AI核验理由补填工具\n"
            "├── test_search_filter.py    # 搜索过滤测试\n"
            "├── requirements.txt         # Python依赖声明\n"
            "├── data/\n"
            "│   ├── architecture.db      # SQLite3主数据库（~972 MB）\n"
            "│   ├── images/              # 建筑图片存储\n"
            "│   ├── img_cache/           # 图片缩略图磁盘缓存\n"
            "│   ├── literature/          # PDF文献存储\n"
            "│   └── vosk-model/          # Vosk离线语音模型\n"
            "├── ui/\n"
            "│   ├── search_tab.py        # 案例查询标签页（6,779行）\n"
            "│   ├── organize_tab.py      # 数据整理标签页（7,196行）\n"
            "│   ├── analysis_tab.py      # 数据分析标签页（3,078行）\n"
            "│   ├── chat_assistant.py    # 聊天助手面板（1,323行）\n"
            "│   ├── workflows.py         # 工作流引擎（1,109行）\n"
            "│   ├── keyword_gen.py       # 多语言关键词生成（1,040行）\n"
            "│   ├── import_dialog.py     # Excel导入对话框（728行）\n"
            "│   └── image_components.py  # 图片缩略图/查看器（719行）\n"
            "├── utils/\n"
            "│   ├── ai_extractor.py      # AI提取/核验/补全核心（69 KB）\n"
            "│   ├── web_search.py        # Serper/Baidu搜索API封装（230行）\n"
            "│   ├── smart_search.py      # 分页多引擎搜索+AI过滤（206行）\n"
            "│   ├── pdf_extractor.py     # PDF文本提取（118行）\n"
            "│   ├── data_cleaning.py     # 数据清洗工具（48行）\n"
            "│   ├── import_export.py     # Excel导入导出对话框（48行）\n"
            "│   ├── keyword_translations.py # 104语言关键词翻译（783 KB）\n"
            "│   ├── perf_monitor.py      # 性能监控工具\n"
            "│   ├── info_verifier.py     # 百度API信息核验（遗留模块）\n"
            "│   └── build_translations.py   # 翻译数据生成脚本\n"
            "├── tests/\n"
            "│   └── test_baidu_image.py  # 百度图片功能测试\n"
            "├── 论文/                    # 论文图表与稿件生成脚本（15个）\n"
            "├── ext/                     # Chrome浏览器扩展（建筑案例提取）\n"
            "├── auto-refresh-extension/  # Chrome扩展（自动刷新防代理掉线）\n"
            "└── 图片下载/                # 图片下载辅助脚本"
        ),
        "en": (
            "Chinoiserie-Architecture-Beyond-China-Database/\n"
            "├── main.py                  # Entry point, MainWindow (109 lines)\n"
            "├── database.py              # Database layer (3,157 lines), 17-table CRUD + migration v12 + undo/redo\n"
            "├── generate_report.py       # Auto-generation script for this report\n"
            "├── refill_reasons.py        # AI verification-reason backfill utility\n"
            "├── test_search_filter.py    # Search-filter tests\n"
            "├── requirements.txt         # Python dependencies\n"
            "├── data/\n"
            "│   ├── architecture.db      # SQLite3 main database (~972 MB)\n"
            "│   ├── images/              # Building image storage\n"
            "│   ├── img_cache/           # Thumbnail disk cache\n"
            "│   ├── literature/          # PDF literature storage\n"
            "│   └── vosk-model/          # Vosk offline speech model\n"
            "├── ui/\n"
            "│   ├── search_tab.py        # Case search tab (6,779 lines)\n"
            "│   ├── organize_tab.py      # Data curation tab (7,196 lines)\n"
            "│   ├── analysis_tab.py      # Data analysis tab (3,078 lines)\n"
            "│   ├── chat_assistant.py    # Chat assistant panel (1,323 lines)\n"
            "│   ├── workflows.py         # Workflow controller (1,109 lines)\n"
            "│   ├── keyword_gen.py       # Multilingual keyword generation (1,040 lines)\n"
            "│   ├── import_dialog.py     # Excel import dialog (728 lines)\n"
            "│   └── image_components.py  # Thumbnail/image viewer (719 lines)\n"
            "├── utils/\n"
            "│   ├── ai_extractor.py      # AI extraction/verification/completion core (69 KB)\n"
            "│   ├── web_search.py        # Serper/Baidu search API wrappers (230 lines)\n"
            "│   ├── smart_search.py      # Paged multi-engine search + AI filtering (206 lines)\n"
            "│   ├── pdf_extractor.py     # PDF text extraction (118 lines)\n"
            "│   ├── data_cleaning.py     # Data cleaning utilities (48 lines)\n"
            "│   ├── import_export.py     # Excel import/export dialogs (48 lines)\n"
            "│   ├── keyword_translations.py # 104-language keyword translations (783 KB)\n"
            "│   ├── perf_monitor.py      # Performance monitoring utility\n"
            "│   ├── info_verifier.py     # Baidu API verification (legacy module)\n"
            "│   └── build_translations.py   # Translation data generation script\n"
            "├── tests/\n"
            "│   └── test_baidu_image.py  # Baidu image function tests\n"
            "├── 论文/                    # Paper figure & manuscript generation scripts (15 files)\n"
            "├── ext/                     # Chrome extension (building-case extraction)\n"
            "├── auto-refresh-extension/  # Chrome extension (auto-refresh, anti-proxy-disconnect)\n"
            "└── 图片下载/                # Image-download helper scripts"
        ),
    },

    # ── Ch3 ────────────────────────────────────────────────────────────
    "3.1": {"zh": "4.1  设计原则", "en": "3.1  Design Principles"},
    "3.2": {"zh": "4.2  核心表结构", "en": "3.2  Core Table Schema"},
    "3.2.1": {"zh": "4.2.1  正式案例表 (cases)", "en": "3.2.1  Verified Cases Table (cases)"},
    "3.2.2": {"zh": "4.2.2  原始案例表 (raw_cases)", "en": "3.2.2  Raw Cases Table (raw_cases)"},
    "3.2.3": {"zh": "4.2.3  辅助表一览", "en": "3.2.3  Auxiliary Tables"},
    "3.3": {"zh": "4.3  数据迁移机制", "en": "3.3  Data Migration Mechanism"},
    "3.4": {"zh": "4.4  撤销/重做机制", "en": "3.4  Undo/Redo Mechanism"},

    "db_principles": {
        "zh": (
            "数据库设计遵循以下原则：（1）分层存储——区分正式案例库（cases）和原始待核验数据"
            "（raw_cases），确保入库数据的质量；（2）来源可追溯——通过 case_sources 和 "
            "field_sources 表记录每条数据、每个字段的来源URL和置信度；（3）灵活扩展——支持"
            "运行时动态添加/删除字段，适应研究需求的变化；（4）操作可撤销——内置撤销/重做栈，"
            "支持最近10次操作的回滚。"
        ),
        "en": (
            "The database design follows four principles: (1) tiered storage — separating "
            "verified cases (cases) from raw, unverified data (raw_cases) to ensure data "
            "quality; (2) provenance tracking — recording the source URL and confidence "
            "level for every record and every field through the case_sources and field_sources "
            "tables; (3) flexible extension — supporting runtime addition and removal of "
            "fields to adapt to evolving research needs; (4) undoable operations — a "
            "built-in undo/redo stack supporting rollback of the last 10 operations."
        ),
    },
    "cases_desc": {
        "zh": (
            "cases 表是系统的核心数据表，存储经过核验确认的正式建筑案例。每条记录对应一个"
            "海外中式建筑实体，包含名称、位置、年代、类型、保护等级等完整信息。"
        ),
        "en": (
            "The cases table is the core data table of the system, storing verified "
            "architectural cases. Each record corresponds to one Chinoiserie Architecture "
            "entity outside China, with full information including name, location, date, "
            "type, and protection status."
        ),
    },
    "col_field": {"zh": "字段名", "en": "Field Name"},
    "col_name_zh": {"zh": "中文名称", "en": "Description"},
    "col_type": {"zh": "类型", "en": "Type"},
    "col_desc": {"zh": "说明", "en": "Notes"},
    "col_pk": {"zh": "主键，自增", "en": "Primary key, auto-increment"},
    "col_name_desc": {"zh": "NOT NULL，中文名称", "en": "NOT NULL"},
    "col_name_en": {"zh": "英文名称", "en": "Name in English"},
    "col_name_official": {"zh": "官方名称", "en": "Name in local official language"},
    "col_country": {"zh": "国家", "en": "Country"},
    "col_city": {"zh": "城市", "en": "City"},
    "col_year_built": {"zh": "建造时间", "en": "Year built"},
    "col_demolition": {"zh": "拆除时间", "en": "Year demolished"},
    "col_restoration": {"zh": "修复时间", "en": "Year restored"},
    "col_quantity": {"zh": "数量", "en": "Quantity"},
    "col_building_type": {"zh": "建筑类型", "en": "Building type (e.g., garden, temple, residence)"},
    "col_architect": {"zh": "建筑师/设计者", "en": "Architect / designer"},
    "col_lat": {"zh": "纬度", "en": "Latitude"},
    "col_lon": {"zh": "经度", "en": "Longitude"},
    "col_material": {"zh": "建筑材料", "en": "Building material"},
    "col_area": {"zh": "面积(平方米)", "en": "Area (m²)"},
    "col_protection": {"zh": "保护等级", "en": "Protection level"},
    "col_description": {"zh": "描述", "en": "Description"},
    "col_image_path": {"zh": "图片路径", "en": "Local image file path"},
    "col_image_urls": {"zh": "图片URL", "en": "Online image URLs"},
    "col_search_time": {"zh": "检索时间", "en": "Search timestamp"},
    "col_search_kw": {"zh": "检索词", "en": "Search keywords used"},
    "col_condition": {"zh": "建筑状况", "en": "Building condition"},
    "col_spread": {"zh": "传播方式", "en": "Diffusion channel"},
    "col_notes": {"zh": "备注", "en": "Notes"},
    "col_verif_status": {"zh": "核验状态", "en": "Verification status"},
    "col_completeness": {"zh": "完整度评分", "en": "Completeness score"},
    "col_credibility": {"zh": "来源可信度", "en": "Source credibility"},
    "col_raw_ids": {"zh": "原始案例ID", "en": "Associated raw_cases record IDs"},
    "col_sort_order": {"zh": "排序序号", "en": "Sort order"},
    "col_created": {"zh": "创建时间", "en": "Created at"},
    "col_updated": {"zh": "更新时间", "en": "Updated at"},

    "raw_desc": {
        "zh": (
            "raw_cases 是数据管道的中间层，存储从网络搜索、AI提取等渠道获得的未核验数据。"
            "该表支持运行时动态添加字段（通过ALTER TABLE ADD COLUMN），字段定义同时记录在"
            "内存的 RAW_FIELDS 列表中，实现元数据驱动的灵活数据管理。"
        ),
        "en": (
            "raw_cases is the intermediate layer of the data pipeline, storing unverified "
            "data obtained from web searches, AI extraction, and other sources. The table "
            "supports runtime addition of custom fields (via ALTER TABLE ADD COLUMN), with "
            "field definitions tracked in the in-memory RAW_FIELDS list for metadata-driven "
            "flexible data management."
        ),
    },
    "col_auth_status": {"zh": "真实性状态", "en": "Authenticity status"},
    "col_auth_score": {"zh": "真实性评分", "en": "Authenticity score"},
    "col_auth_notes": {"zh": "核验备注", "en": "Verification notes"},
    "col_building_status": {"zh": "建筑状态", "en": "Building status"},
    "col_source_urls": {"zh": "来源URL", "en": "Source URLs"},
    "col_function": {"zh": "功能", "en": "Function"},

    "aux_tbl_name": {"zh": "表名", "en": "Table Name"},
    "aux_tbl_purpose": {"zh": "用途", "en": "Purpose"},
    "aux_tbl_keys": {"zh": "关键字段", "en": "Key Fields"},
    "aux_bt": {"zh": "建筑类型字典表", "en": "Building-type dictionary"},
    "aux_cs": {"zh": "案例来源URL（外键关联raw_cases）", "en": "Case source URLs (FK → raw_cases)"},
    "aux_fs": {"zh": "字段级溯源信息", "en": "Field-level provenance"},
    "aux_ci": {"zh": "建筑图片存储", "en": "Building image storage"},
    "aux_it": {"zh": "软删除图片回收站", "en": "Soft-delete image trash"},
    "aux_vl": {"zh": "核验操作审计日志", "en": "Verification audit log"},
    "aux_ol": {"zh": "通用操作审计日志", "en": "General operation audit log"},
    "aux_lit": {"zh": "学术文献库", "en": "Academic literature repository"},
    "aux_lml": {"zh": "文献-案例匹配日志", "en": "Literature–case match log"},
    "aux_fr": {"zh": "AI筛选后的搜索结果", "en": "AI-filtered search results"},
    "aux_rs": {"zh": "原始搜索结果（未筛选）", "en": "Raw search results (unfiltered)"},
    "aux_ssr": {"zh": "学术搜索结果", "en": "Scholarly search results"},
    "aux_vc": {"zh": "已验证案例快照", "en": "Verified-cases snapshot"},
    "aux_dr": {"zh": "去重删除记录", "en": "Deduplication removal records"},
    "aux_cases": {"zh": "正式建筑案例主表（31字段）", "en": "Verified architectural cases main table (31 fields)"},
    "aux_raw_cases": {"zh": "待核验原始案例表（26+字段）", "en": "Unverified raw cases table (26+ fields)"},

    "migration_p1": {
        "zh": (
            "系统通过 migrate_db_v2() 至 migrate_db_v12() 共11个增量迁移函数实现无停机式的数据库升级。"
            "迁移在每次应用启动时自动执行，检测并补全缺失的列和表。关键设计：（1）每个ALTER TABLE "
            "ADD COLUMN语句包裹在try/except中，确保幂等性——重复执行不会报错；"
            "（2）新增列使用DEFAULT值，已有数据不受影响；（3）source_urls字段的数据迁移"
            "（从旧格式的TEXT字段解析为规范化的case_sources表）通过_migrate_source_urls_to_case_sources()完成，"
            "包括URL分类（学术/政府/新闻/博客）和去重。"
        ),
        "en": (
            "The system implements zero-downtime incremental database upgrades through 11 "
            "migration functions, from migrate_db_v2() through migrate_db_v12(), which run "
            "automatically at each application start, detecting and adding missing columns "
            "and tables. Key design features: (1) each ALTER TABLE ADD COLUMN statement is "
            "wrapped in try/except to ensure idempotency — repeated execution will not "
            "error; (2) new columns use DEFAULT values so existing data are unaffected; "
            "(3) data migration of the source_urls field (from legacy TEXT format to the "
            "normalised case_sources table) is handled by _migrate_source_urls_to_case_sources(), "
            "including URL classification (academic/government/news/blog) and deduplication."
        ),
    },
    "undo_p1": {
        "zh": (
            "系统维护一个最大深度为10的撤销/重做内存栈，覆盖6种操作类型：insert（插入）、"
            "delete（删除）、update（修改）、status_change（状态变更）、merge（合并）和"
            "import_batch（批量导入）。每次操作前将旧状态推入栈中，撤销时将数据恢复。"
            "对于涉及关联表（如case_sources、case_images）的操作，同时保存和恢复关联记录。"
            "所有撤销/重做操作返回人类可读的中文描述，在UI中以通知形式展示。"
        ),
        "en": (
            "The system maintains an in-memory undo/redo stack with a maximum depth of 10, "
            "covering six operation types: insert, delete, update, status_change, merge, and "
            "import_batch. Before each operation, the old state is pushed onto the stack; "
            "on undo, the data are restored. For operations involving related tables (e.g., "
            "case_sources, case_images), associated records are saved and restored as well. "
            "All undo/redo operations return human-readable descriptions displayed as "
            "notifications in the UI."
        ),
    },

    # ── Ch4 ────────────────────────────────────────────────────────────
    "4.1": {"zh": "5.1  案例查询模块 (search_tab.py)", "en": "4.1  Case Search Module (search_tab.py)"},
    "4.1.1": {"zh": "5.1.1  网页搜索引擎管理", "en": "4.1.1  Web Search Engine Management"},
    "4.1.2": {"zh": "5.1.2  Chrome扩展HTTP服务器", "en": "4.1.2  Chrome Extension HTTP Server"},
    "4.1.3": {"zh": "5.1.3  AI案例提取", "en": "4.1.3  AI Case Extraction"},
    "4.1.4": {"zh": "5.1.4  文献库管理", "en": "4.1.4  Literature Management"},
    "4.2": {"zh": "5.2  数据整理模块 (organize_tab.py)", "en": "4.2  Data Curation Module (organize_tab.py)"},
    "4.2.1": {"zh": "5.2.1  原始数据面板", "en": "4.2.1  Raw Data Panel"},
    "4.2.2": {"zh": "5.2.2  真实性核验面板", "en": "4.2.2  Authenticity Verification Panel"},
    "4.2.3": {"zh": "5.2.3  正式案例面板", "en": "4.2.3  Verified Cases Panel"},
    "4.3": {"zh": "5.3  数据分析模块 (analysis_tab.py)", "en": "4.3  Data Analysis Module (analysis_tab.py)"},
    "4.4": {"zh": "5.4  聊天助手模块 (chat_assistant.py)", "en": "4.4  Chat Assistant Module (chat_assistant.py)"},
    "4.4.1": {"zh": "5.4.1  快捷命令系统", "en": "4.4.1  Quick Command System"},
    "4.4.2": {"zh": "5.4.2  AI自然语言理解", "en": "4.4.2  AI Natural-Language Understanding"},
    "4.4.3": {"zh": "5.4.3  离线语音识别", "en": "4.4.3  Offline Speech Recognition"},
    "4.5": {"zh": "5.5  工作流引擎 (workflows.py)", "en": "4.5  Workflow Controller (workflows.py)"},
    "4.5.1": {"zh": "5.5.1  交互式工作流有限状态机", "en": "4.5.1  Interactive Workflow Finite-State Machine"},

    "search_p1": {
        "zh": (
            "案例查询模块是系统的信息入口，提供网页搜索、AI提取和库内搜索三大功能。"
            "该模块包含约6,779行代码，核心类为SearchTab（继承自object，通过frame属性"
            "提供ttk.Frame容器）。"
        ),
        "en": (
            "The Case Search module is the information entry point of the system, providing "
            "three main functions: web search, AI extraction, and local database query. "
            "The module contains approximately 6,779 lines of code; its core class is "
            "SearchTab (extending object, exposing a ttk.Frame container via a frame attribute)."
        ),
    },
    "engines_p1": {
        "zh": (
            "系统将搜索引擎分为两组：灰色文献组（Google、百度）和研究文献组"
            "（Web of Science、Google Scholar、中国知网、中国国家图书馆、罗马国家图书馆）。"
            "每组支持全选/全不选操作。用户输入关键词后，点击搜索按钮即可在Chrome浏览器中"
            "为每个关键词×每个搜索引擎打开对应的搜索标签页。搜索URL模板存储在ALL_ENGINES"
            "字典中，便于扩展新的搜索引擎。"
        ),
        "en": (
            "The system organises search engines into two groups: grey literature (Google, "
            "Baidu) and scholarly literature (Web of Science, Google Scholar, CNKI, National "
            "Library of China, National Central Library of Rome). Each group supports "
            "select-all / deselect-all. After the user enters keywords, clicking the search "
            "button opens the corresponding search tabs in Chrome for every keyword × engine "
            "combination. Search URL templates are stored in the ALL_ENGINES dictionary, "
            "making it easy to add new engines."
        ),
    },
    "chrome_ext_p1": {
        "zh": (
            "系统在127.0.0.1:9876端口启动一个轻量HTTP服务器（_ExtractServerHandler），"
            "接收来自Chrome浏览器扩展发送的网页内容。Chrome扩展在用户浏览搜索结果的网页时，"
            "自动抓取页面正文并通过HTTP POST发送到本应用。服务器根据当前活跃的标签页将内容"
            "路由到AI提取文本框或核验面板的粘贴区域。"
        ),
        "en": (
            "The system starts a lightweight HTTP server (_ExtractServerHandler) on "
            "127.0.0.1:9876 to receive web-page content from a Chrome browser extension. "
            "As the user browses search-result pages, the Chrome extension automatically "
            "captures the page body and sends it to the application via HTTP POST. The "
            "server routes content to the AI extraction text area or the verification "
            "panel paste area based on the currently active tab."
        ),
    },
    "ai_extract_p1": {
        "zh": (
            "对于积累的网页文本内容，系统调用DeepSeek API进行案例提取。通过精心设计的提示词"
            "（EXTRACTION_PROMPT），AI被要求从多语言文本中识别建筑案例，并以结构化的JSON"
            "格式返回名称、国家、城市、来源等信息。当检测到中文比例低于5%时，系统自动先"
            "将文本翻译为中文再提取。提取结果显示在结果树中，用户可选择性添加到数据库。"
        ),
        "en": (
            "For accumulated web-page text, the system calls the DeepSeek API for case "
            "extraction. Through a carefully designed prompt (EXTRACTION_PROMPT), the AI is "
            "instructed to identify architectural cases from multilingual text and return "
            "structured JSON with name, country, city, source, and other information. When "
            "the proportion of Chinese characters falls below 5%, the system automatically "
            "translates the text into Chinese before extraction. Results are displayed in a "
            "tree view, and the user can selectively add them to the database."
        ),
    },
    "lit_mgmt_p1": {
        "zh": (
            "系统内置学术文献管理功能，支持：（1）上传PDF文献，自动提取全文文本并生成AI摘要；"
            "（2）文献列表的增删改查和搜索过滤；（3）将文献内容发送到AI提取区域进行分析。"
            "PDF文本提取基于PyMuPDF（fitz）库，针对中文学术PDF中常见的PUA（私人使用区）"
            "字符进行专门的清理映射。"
        ),
        "en": (
            "The system includes built-in academic literature management: (1) upload PDF "
            "documents, auto-extract full text and generate AI summaries; (2) CRUD operations "
            "and search filtering for the literature list; (3) send literature content to the "
            "AI extraction area for analysis. PDF text extraction uses the PyMuPDF (fitz) "
            "library, with dedicated cleaning mappings for PUA (Private Use Area) characters "
            "commonly found in Chinese academic PDFs."
        ),
    },
    "organize_p1": {
        "zh": (
            "数据整理模块是系统最复杂的模块（7,196行），由一个三级标签页容器（OrganizeTab）"
            "管理三个子面板：原始数据面板（RawDataPanel）、真实性核验面板（VerificationPanel）"
            "和正式案例面板（DatabasePanel）。"
        ),
        "en": (
            "The Data Curation module is the most complex module in the system (7,196 lines), "
            "managed by a three-level tab container (OrganizeTab) with three sub-panels: "
            "Raw Data Panel (RawDataPanel), Authenticity Verification Panel "
            "(VerificationPanel), and Verified Cases Panel (DatabasePanel)."
        ),
    },
    "raw_panel_p1": {
        "zh": (
            "原始数据面板管理raw_cases表中的待处理数据。主要功能包括：Excel/CSV文件的智能导入"
            "（支持模糊列名匹配、坐标字段自动拆分、WPS图片公式解析）、数据提交核验队列、"
            "重复检测与合并、数据清洗、以及拖拽排序。面板支持按状态（全部/待处理/待核验/已通过/已拒绝）"
            "和自定义条件进行筛选。"
        ),
        "en": (
            "The Raw Data Panel manages pending data in the raw_cases table. Key features "
            "include: intelligent Excel/CSV import (fuzzy column-name matching, automatic "
            "coordinate-field splitting, WPS image-formula parsing), submission to the "
            "verification queue, duplicate detection and merging, data cleaning, and "
            "drag-and-drop sorting. The panel supports filtering by status (All / Pending / "
            "Awaiting Verification / Approved / Rejected) and custom criteria."
        ),
    },
    "verify_panel_p1": {
        "zh": (
            "真实性核验面板是系统的核心功能，提供字段级的建筑信息核验。左侧队列树显示待核验案例，"
            "右侧详情面板展示13个可核验字段（国家、名称、城市、建造年份、拆除年份、修复年份、"
            "建筑状态、传播方式、数量、功能、设计者、坐标、图片URL）。"
        ),
        "en": (
            "The Authenticity Verification Panel is the core function of the system, providing "
            "field-level architectural-information verification. The left queue tree displays "
            "cases awaiting verification; the right detail panel shows 13 verifiable fields "
            "(country, name, city, year built, year demolished, year restored, building "
            "status, diffusion channel, quantity, function, designer, coordinates, image URLs)."
        ),
    },
    "verify_panel_p2": {
        "zh": (
            "每个字段支持四种核验引擎：（1）百度搜索——在Chrome中打开百度搜索结果，通过"
            "Chrome扩展将页面内容传回核验面板的粘贴区域；（2）Google搜索——同上，使用Google"
            "搜索引擎；（3）DeepSeek核验——直接调用DeepSeek API，基于模型内部知识进行"
            "核验，无需浏览器搜索；（4）Gemini核验——通过Gemini API进行多语言核验，"
            "尤其适合非中文名称的验证。"
        ),
        "en": (
            "Each field supports four verification engines: (1) Baidu Search — opens Baidu "
            "search results in Chrome, with page content sent back to the verification "
            "panel paste area via the Chrome extension; (2) Google Search — same workflow "
            "using Google; (3) DeepSeek Verification — calls the DeepSeek API directly, "
            "leveraging the model's internal knowledge without browser search; (4) Gemini "
            "Verification — multilingual verification through the Gemini API, particularly "
            "suited for non-Chinese names."
        ),
    },
    "verify_panel_p3": {
        "zh": (
            "核验结果通过两个AI分析按钮处理：\"名称字段分析\"解析粘贴区域中所有名称核验结果，"
            "融合多个来源的信息并处理冲突；\"其他字段分析\"从粘贴内容中提取JSON格式的字段"
            "信息并填充到对应字段。核验通过后可一键将案例从raw_cases提升到正式cases表。"
        ),
        "en": (
            "Verification results are processed through two AI analysis buttons: "
            "\"Analyse Name Field\" parses all name-verification results in the paste area, "
            "merging information from multiple sources and resolving conflicts; \"Analyse "
            "Other Fields\" extracts JSON-formatted field information from the paste content "
            "and populates the corresponding fields. Once verified, cases can be promoted "
            "from raw_cases to the verified cases table in one click."
        ),
    },
    "db_panel_p1": {
        "zh": (
            "正式案例面板管理已核验入库的建筑案例。主要功能：中文/英文双语言Excel导出"
            "（包含图片嵌入、核验结果、参考链接）、案例增删改、重复检测、缺失字段检查、"
            "以及丰富的筛选功能（文本搜索+下拉筛选+排序控制）。导出功能支持从verification_log中"
            "提取结构化名称信息（中文名/英文名/官方名/别名等）。"
        ),
        "en": (
            "The Verified Cases Panel manages cases that have been verified and stored. "
            "Main functions: bilingual Excel export (Chinese/English, with embedded images, "
            "verification results, and reference links); case CRUD; duplicate detection; "
            "missing-field checking; and rich filtering (text search + dropdown filters + "
            "sort controls). The export function extracts structured name information "
            "(Chinese/English/official/alias) from verification_log."
        ),
    },
    "analysis_p1": {
        "zh": (
            "数据分析模块提供四个维度的统计图表：建筑类型、国家分布、建造年代和保护等级。"
            "图表使用matplotlib渲染，配置了针对中文字体的支持（PingFang SC、Heiti SC、"
            "Songti SC等），在macOS和Windows上均可正常显示。图表采用大地色系配色，"
            "标签截断为8个字符以保证排版美观。为规避macOS上matplotlib的autorelease pool"
            "崩溃问题，图表渲染延迟到主循环空闲后才执行。模块同时支持全量数据导出和"
            "按当前维度的统计数据导出。"
        ),
        "en": (
            "The Data Analysis module provides statistical charts across four dimensions: "
            "building type, country distribution, founding period, and protection level. "
            "Charts are rendered with matplotlib, configured with Chinese font support "
            "(PingFang SC, Heiti SC, Songti SC, etc.) and display correctly on both macOS "
            "and Windows. A warm earth-tone colour palette is used, and labels are truncated "
            "to 8 characters for visual clarity. To avoid matplotlib autorelease-pool crashes "
            "on macOS, chart rendering is deferred until the main loop is idle. The module "
            "also supports full-data export and dimension-specific statistical-data export."
        ),
    },
    "chat_p1": {
        "zh": (
            "聊天助手是一个可折叠的右侧面板，支持文本输入、语音输入和自然语言交互。"
            "它由三个核心组件构成：ChatAssistant（UI面板）、CommandDispatcher（命令分发）"
            "和VoiceInput（语音输入）。"
        ),
        "en": (
            "The Chat Assistant is a collapsible right-side panel supporting text input, "
            "voice input, and natural-language interaction. It consists of three core "
            "components: ChatAssistant (UI panel), CommandDispatcher (command routing), "
            "and VoiceInput (speech input)."
        ),
    },
    "quick_cmd_p1": {
        "zh": (
            "系统内置24条正则快捷命令，覆盖导航、CRUD操作、核验审批、工作流触发等常见操作。"
            "快捷命令在本地即时匹配，无需网络请求，响应速度极快。命令包括：搜索关键词、"
            "切换标签页/子标签页、刷新列表、导入/导出数据、自动补全字段、删除选中、"
            "添加案例、查看详情、上一条/下一条导航、提交核验、批准/拒绝、全选/取消全选等。"
            "此外还支持闲聊问候（你好、谢谢、再见）。"
        ),
        "en": (
            "The system has 24 built-in regex-based quick commands covering navigation, "
            "CRUD operations, verification approval, workflow triggering, and other common "
            "tasks. Quick commands are matched locally and instantaneously without network "
            "requests. Commands include: search keywords, switch tabs/sub-tabs, refresh "
            "list, import/export data, auto-complete fields, delete selected, add case, "
            "view details, previous/next navigation, submit for verification, approve/reject, "
            "select all/deselect all, and casual greetings."
        ),
    },
    "nl_understanding_p1": {
        "zh": (
            "当用户输入的文本无法匹配任何快捷命令时，系统自动将文本发送给DeepSeek API进行"
            "自然语言理解。通过精心设计的系统提示词（SYSTEM_PROMPT，约90行），AI被要求将"
            "用户的自然语言指令映射为结构化的JSON动作（{\"action\": \"动作名\", \"params\": {...}}）。"
            "系统支持50+种AI动作，覆盖全部功能模块。AI调用使用temperature=0.1的低温度参数，"
            "确保输出的一致性和确定性。"
        ),
        "en": (
            "When the user's input does not match any quick command, the system automatically "
            "sends the text to the DeepSeek API for natural-language understanding. Through "
            "a carefully designed system prompt (SYSTEM_PROMPT, approximately 90 lines), the "
            "AI maps natural-language instructions to structured JSON actions "
            "({\"action\": \"<action_name>\", \"params\": {...}}). The system supports 50+ "
            "AI actions covering all functional modules. AI calls use a low temperature of "
            "0.1 to ensure output consistency and determinism."
        ),
    },
    "voice_p1": {
        "zh": (
            "语音输入使用sounddevice库进行16kHz采样录音（默认3秒），录音内容保存为临时WAV"
            "文件，然后通过Vosk离线语音识别引擎转换为文字。Vosk模型（vosk-model-small-cn）"
            "在应用启动时惰性加载，模型文件约40MB。若模型未安装或Vosk库不可用，系统优雅降级"
            "并提示用户。语音识别结果自动填入输入框并触发发送。"
        ),
        "en": (
            "Voice input uses the sounddevice library for 16 kHz sampling (default 3 seconds); "
            "the recording is saved as a temporary WAV file and converted to text by the Vosk "
            "offline speech-recognition engine. The Vosk model (vosk-model-small-cn) is lazily "
            "loaded at application start and is approximately 40 MB. If the model is not "
            "installed or the Vosk library is unavailable, the system degrades gracefully "
            "and notifies the user. Recognition results are automatically entered into the "
            "input box and sent."
        ),
    },
    "wf_p1": {
        "zh": (
            "工作流引擎将多步骤的复杂操作封装为可一键触发的自动化流程。系统注册了四个工作流，"
            "在WORKFLOW_REGISTRY字典中统一管理，每个工作流包含名称、描述、参数签名和执行函数。"
        ),
        "en": (
            "The Workflow Controller encapsulates multi-step, complex operations into automated "
            "processes that can be triggered with a single click. Four workflows are registered "
            "and managed in the WORKFLOW_REGISTRY dictionary; each workflow includes a name, "
            "description, parameter signature, and execution function."
        ),
    },
    "wf_col_name": {"zh": "工作流", "en": "Workflow"},
    "wf_col_type": {"zh": "类型", "en": "Type"},
    "wf_col_steps": {"zh": "步骤数", "en": "Steps"},
    "wf_col_desc": {"zh": "说明", "en": "Description"},
    "wf1_name": {"zh": "一键搜索", "en": "One-Click Search"},
    "wf1_type": {"zh": "非交互式", "en": "Non-interactive"},
    "wf1_steps": {"zh": "1步", "en": "1 step"},
    "wf1_desc": {
        "zh": (
            "对国家+语言的所有关键词组合打开Chrome搜索标签页。支持195个国家×3种语言"
            "（中文/English/官方语言）×13个核心关键词。中文使用百度，其他语言使用Google。"
            "搜索结果页数超过关键词数×引擎数，适合批量信息采集。"
        ),
        "en": (
            "Opens Chrome search tabs for all keyword combinations for a given country and "
            "language. Supports 195 countries × 3 languages (Chinese / English / official "
            "language) × 13 core keywords. Baidu is used for Chinese queries; Google for "
            "all other languages. The number of search result pages exceeds keywords × "
            "engines, making it suitable for batch information collection."
        ),
    },
    "wf2_name": {"zh": "一键DS核验", "en": "One-Click DS Verification"},
    "wf2_type": {"zh": "交互式", "en": "Interactive"},
    "wf2_steps": {"zh": "9步", "en": "9 steps"},
    "wf2_desc": {
        "zh": (
            "对核验队列中选中案例进行自动化处理：DS名称核验→名称字段分析→DS一键补全→"
            "其他字段分析→百度图片搜索→（可选）谷歌图片搜索。每步通过API异步调用，"
            "在等待用户确认时暂停。"
        ),
        "en": (
            "Automated processing of selected cases in the verification queue: DS name "
            "verification → name-field analysis → DS one-click completion → other-field "
            "analysis → Baidu image search → (optional) Google image search. Each step "
            "uses asynchronous API calls and pauses to wait for user confirmation."
        ),
    },
    "wf3_name": {"zh": "一键G核验", "en": "One-Click G Verification"},
    "wf3_desc": {
        "zh": (
            "与DS核验流程相同，但使用Gemini API替代DeepSeek。Gemini在非中文名称的"
            "核验和多语言搜索方面具有优势。"
        ),
        "en": (
            "Same workflow as DS Verification, but using the Gemini API instead of DeepSeek. "
            "Gemini has advantages in verifying non-Chinese names and multilingual search."
        ),
    },
    "wf4_name": {"zh": "一键核验", "en": "One-Click Cascade Verification"},
    "wf4_desc": {
        "zh": (
            "DS优先、G回退的级联核验流程。先尝试DS核验→补全，询问是否满意；如果不满意，"
            "自动切换到Gemini重新核验→补全。最终打开百度图片搜索，不满意可切换谷歌搜图。"
        ),
        "en": (
            "Cascade verification with DS-first, G-fallback strategy: attempts DS verification "
            "→ completion first, then asks the user whether the result is satisfactory; if "
            "not, it automatically switches to Gemini for re-verification → completion. "
            "Finally opens Baidu image search, with the option to switch to Google if "
            "unsatisfactory."
        ),
    },
    "fsm_p1": {
        "zh": (
            "交互式工作流基于WorkflowSession基类，采用基于步骤的有限状态机模式。每个步骤"
            "对应一个_step_{N}方法，通过next()方法推进状态。状态机支持三种推进方式："
            "（1）自动推进——上一步完成后延迟300ms自动进入下一步；（2）异步等待——API调用"
            "通过_run_async在后台线程执行，完成后在主线程回调；（3）用户交互——通过_ask()"
            "方法向用户提问并等待回复。每步都包裹在try/except中，出错时提供\"重试/跳过/结束\""
            "三种恢复选择。"
        ),
        "en": (
            "Interactive workflows are built on the WorkflowSession base class using a "
            "step-based finite-state-machine pattern. Each step corresponds to a "
            "_step_{N} method, and state is advanced via the next() method. The state "
            "machine supports three progression modes: (1) auto-advance — after a step "
            "completes, the next step is triggered after a 300 ms delay; (2) async wait — "
            "API calls execute on background threads via _run_async, with a main-thread "
            "callback on completion; (3) user interaction — the _ask() method poses a "
            "question to the user and waits for a reply. Every step is wrapped in "
            "try/except, offering three recovery options on error: retry / skip / end."
        ),
    },

    # ── Ch5 ────────────────────────────────────────────────────────────
    "5.1": {"zh": "6.1  AI多模型统一调用", "en": "5.1  Unified Multi-LLM Invocation"},
    "5.2": {"zh": "6.2  多源搜索引擎集成", "en": "5.2  Multi-Source Search-Engine Integration"},
    "5.3": {"zh": "6.3  智能数据导入", "en": "5.3  Intelligent Data Import"},
    "5.4": {"zh": "6.4  图片管理系统", "en": "5.4  Image Management System"},
    "5.5": {"zh": "6.5  PDF文本提取与中文处理", "en": "5.5  PDF Text Extraction and Chinese-Language Processing"},
    "5.6": {"zh": "6.6  线程安全与并发控制", "en": "5.6  Thread Safety and Concurrency Control"},

    "ai_unified_p1": {
        "zh": (
            "系统通过OpenAI兼容API协议实现了DeepSeek和Gemini两个模型的统一调用。"
            "调用链路如下：业务代码 → _call_yunwu() / verify_field_deepseek() 等函数 → "
            "openai.OpenAI.chat.completions.create() → HTTP POST → API端点。"
        ),
        "en": (
            "The system achieves unified invocation of DeepSeek and Gemini through an "
            "OpenAI-compatible API protocol. The call chain is: business logic → "
            "_call_yunwu() / verify_field_deepseek() etc. → "
            "openai.OpenAI.chat.completions.create() → HTTP POST → API endpoint."
        ),
    },
    "ai_unified_p2": {
        "zh": (
            "AI提示词工程是系统的关键设计。系统维护了8套专用提示词模板，分别用于案例提取"
            "（EXTRACTION_PROMPT）、真实性评估（VERIFICATION_PROMPT）、字段核验"
            "（FIELD_VERIFY_PROMPT）、字段补全（COMPLETION_PROMPT、AUTO_COMPLETE_PROMPT）、"
            "交叉验证（CROSS_VALIDATION_PROMPT）等场景。每套提示词经过反复迭代优化，"
            "确保输出格式的一致性和结果的可用性。AI返回的JSON结果经过_parse_response()"
            "函数处理，该函数实现了多级JSON提取策略（直接解析→Markdown代码块提取→"
            "正则匹配→括号平衡修复），能够适应各种非标准输出格式。"
        ),
        "en": (
            "Prompt engineering is a critical design element. The system maintains eight "
            "dedicated prompt templates for case extraction (EXTRACTION_PROMPT), authenticity "
            "assessment (VERIFICATION_PROMPT), field verification (FIELD_VERIFY_PROMPT), "
            "field completion (COMPLETION_PROMPT, AUTO_COMPLETE_PROMPT), cross-validation "
            "(CROSS_VALIDATION_PROMPT), and other scenarios. Each prompt has been iteratively "
            "refined to ensure consistent output formatting and usable results. AI-returned "
            "JSON is processed by _parse_response(), which implements a multi-level JSON "
            "extraction strategy (direct parse → Markdown code-block extraction → regex "
            "matching → bracket-balance repair) capable of handling a wide range of "
            "non-standard output formats."
        ),
    },
    "search_integration_p1": {
        "zh": (
            "系统集成了三种搜索方式：（1）浏览器直接搜索——通过webbrowser模块打开Chrome"
            "标签页，配合Chrome扩展将搜索结果页内容回传；（2）Serper.dev API——Google搜索"
            "的结构化API，通过HTTP代理访问，支持分页和按国家/语言过滤；（3）百度千帆API——"
            "百度搜索的结构化API，单次最多返回50条结果。"
        ),
        "en": (
            "The system integrates three search modalities: (1) direct browser search — "
            "opens Chrome tabs via the webbrowser module, with the Chrome extension "
            "relaying search-result page content back to the application; (2) Serper.dev "
            "API — a structured Google Search API, accessed through an HTTP proxy, "
            "supporting pagination and country/language filtering; (3) Baidu Qianfan API — "
            "a structured Baidu Search API returning up to 50 results per request."
        ),
    },
    "search_integration_p2": {
        "zh": (
            "smart_search.py实现了智能分页搜索策略：每次获取10条结果，通过AI逐批过滤"
            "相关性（判断是否同时涉及中文/中式风格和建筑/建造），当连续N页（empty_stop"
            "参数）无相关结果时提前终止，在保证查全率的同时避免无效翻页。过滤后的结果"
            "存入filtered_results表，原始数据保留在raw_search_results表。"
        ),
        "en": (
            "smart_search.py implements an intelligent paged search strategy: it fetches "
            "10 results at a time and uses AI to batch-filter for relevance (determining "
            "whether each record involves both Chinoiserie/Chinese style and architecture/"
            "building). When N consecutive pages (empty_stop parameter) yield no relevant "
            "results, the search terminates early — preserving recall while avoiding "
            "wasted pagination. Filtered results are stored in the filtered_results table; "
            "raw data are preserved in raw_search_results."
        ),
    },
    "import_p1": {
        "zh": (
            "Excel导入模块（import_dialog.py）实现了灵活的数据映射算法_fuzzy_match()，"
            "采用六级降级匹配策略：完全匹配（1.0置信度）→忽略大小写匹配（1.0）→子串匹配"
            "（0.9）→部分包含匹配（0.8）→图片列检测（1.0）→关键词启发式匹配（0.7）。"
            "关键词映射表涵盖52对常见中英文字段名的对应关系（如\"国家\"→country、"
            "\"建造年代\"→construction_year等）。"
        ),
        "en": (
            "The Excel import module (import_dialog.py) implements a flexible data-mapping "
            "algorithm, _fuzzy_match(), using a six-tier fallback matching strategy: exact "
            "match (confidence 1.0) → case-insensitive match (1.0) → substring match (0.9) "
            "→ partial contain match (0.8) → image-column detection (1.0) → keyword "
            "heuristic match (0.7). The keyword mapping table covers 52 pairs of common "
            "Chinese–English field-name correspondences (e.g., \"国家\" → country, "
            "\"建造年代\" → construction_year)."
        ),
    },
    "import_p2": {
        "zh": (
            "导入模块还专门处理了WPS Office的DISPIMG图片公式格式，通过直接解析xlsx文件的"
            "ZIP内部结构（xl/cellimages.xml和xl/_rels/cellimages.xml.rels）来提取嵌入图片，"
            "解决了openpyxl标准API无法读取WPS图片的问题。"
        ),
        "en": (
            "The import module also handles the WPS Office DISPIMG image-formula format by "
            "directly parsing the internal ZIP structure of xlsx files "
            "(xl/cellimages.xml and xl/_rels/cellimages.xml.rels) to extract embedded "
            "images, solving the problem that the standard openpyxl API cannot read WPS images."
        ),
    },
    "image_mgmt_p1": {
        "zh": (
            "图片管理采用\"文件系统存储+数据库索引\"的混合架构。图片文件以UUID命名保存在"
            "data/images/目录，数据库case_images表存储文件路径、来源URL和创建时间。"
            "删除操作采用软删除策略，图片先移入data/images/trash/回收站，可随时恢复。"
            "缩略图加载使用磁盘缓存（data/img_cache/），以URL的MD5哈希为键，"
            "TTL为1天，通过ThreadPoolExecutor(max_workers=3)进行并行异步加载。"
        ),
        "en": (
            "Image management uses a hybrid architecture of file-system storage + database "
            "index. Image files are saved with UUID filenames in data/images/; the "
            "case_images table stores file paths, source URLs, and creation timestamps. "
            "Deletion uses a soft-delete strategy: images are first moved to the "
            "data/images/trash/ recycle bin and can be restored at any time. Thumbnail "
            "loading uses a disk cache (data/img_cache/) keyed by the MD5 hash of the URL, "
            "with a TTL of 1 day, and is loaded asynchronously in parallel via "
            "ThreadPoolExecutor(max_workers=3)."
        ),
    },
    "pdf_p1": {
        "zh": (
            "PDF文本提取使用PyMuPDF（fitz）库。针对中文学术PDF中常见的PUA（私人使用区，"
            "Unicode U+E000-U+F8FF）字符问题，_clean_pua()函数维护了一个字体特定的字符"
            "映射表，将PUA字符替换为对应的标准Unicode字符。该映射表覆盖了常见的学术PDF"
            "字体（如SimSun、SimHei、FangSong等）中PUA编码的标点符号和生僻字。"
        ),
        "en": (
            "PDF text extraction uses the PyMuPDF (fitz) library. For the PUA (Private Use "
            "Area, Unicode U+E000–U+F8FF) character issue common in Chinese academic PDFs, "
            "the _clean_pua() function maintains a font-specific character mapping table "
            "that replaces PUA characters with their corresponding standard Unicode "
            "characters. This mapping covers PUA-encoded punctuation and rare characters "
            "in common academic PDF fonts (SimSun, SimHei, FangSong, etc.)."
        ),
    },
    "thread_p1": {
        "zh": (
            "由于Tkinter不是线程安全的，所有UI更新操作必须通过root.after(0, callback)从"
            "主线程执行。系统的并发模型为：后台线程执行耗时操作（API调用、图片下载、文件读写），"
            "完成后通过after()回调更新UI。用户在聊天助手中的每次输入都在独立线程中处理，"
            "避免阻塞UI。工作流引擎的_run_async方法将API调用放入后台线程，回调在主线程执行，"
            "确保UI响应性。取消机制通过_cancelled标志实现，在各关键节点检查，支持用户随时"
            "中断正在执行的操作。"
        ),
        "en": (
            "Because Tkinter is not thread-safe, all UI updates must be executed from the "
            "main thread via root.after(0, callback). The system's concurrency model is: "
            "background threads execute time-consuming operations (API calls, image downloads, "
            "file I/O) and update the UI through after() callbacks upon completion. Each user "
            "input in the Chat Assistant is processed in a separate thread to avoid blocking "
            "the UI. The workflow controller's _run_async method places API calls on "
            "background threads, with callbacks executing on the main thread to ensure UI "
            "responsiveness. A cancellation mechanism using a _cancelled flag checked at "
            "key points allows the user to interrupt any ongoing operation at any time."
        ),
    },

    # ── Ch6 ────────────────────────────────────────────────────────────
    "6.1": {"zh": "7.1  运行环境要求", "en": "6.1  System Requirements"},
    "6.2": {"zh": "7.2  安装步骤", "en": "6.2  Installation"},
    "6.3": {"zh": "7.3  Chrome扩展配置", "en": "6.3  Chrome Extension Setup"},
    "6.4": {"zh": "7.4  数据库维护", "en": "6.4  Database Maintenance"},

    "env_col_component": {"zh": "组件", "en": "Component"},
    "env_col_requirement": {"zh": "要求", "en": "Requirement"},
    "env_os": {"zh": "操作系统", "en": "Operating System"},
    "env_os_val": {"zh": "macOS 10.15+ / Windows 10+ / Linux", "en": "macOS 10.15+ / Windows 10+ / Linux"},
    "env_python": {"zh": "Python版本", "en": "Python Version"},
    "env_python_val": {"zh": "3.8+（推荐3.11）", "en": "3.8+ (3.11 recommended)"},
    "env_ram": {"zh": "内存", "en": "RAM"},
    "env_ram_val": {"zh": "≥ 4GB（Vosk语音模型运行时）", "en": "≥ 4 GB (with Vosk speech model)"},
    "env_disk": {"zh": "磁盘空间", "en": "Disk Space"},
    "env_disk_val": {"zh": "≥ 2GB（含Vosk模型和图片存储）", "en": "≥ 2 GB (including Vosk model and image storage)"},
    "env_net": {"zh": "网络", "en": "Network"},
    "env_net_val": {"zh": "需要互联网连接（AI API调用和网络搜索）", "en": "Internet connection required (AI API calls and web search)"},

    "install_p1": {"zh": "（1）安装Python依赖：", "en": "(1) Install Python dependencies:"},
    "install_p2": {"zh": "（2）安装语音识别依赖（可选）：", "en": "(2) Install speech-recognition dependencies (optional):"},
    "install_p3": {
        "zh": (
            "（3）下载Vosk中文语音模型（可选）：从 https://alphacephei.com/vosk/models "
            "下载 vosk-model-small-cn-0.22.zip，解压到 data/vosk-model/vosk-model-small-cn/ 目录。"
        ),
        "en": (
            "(3) Download the Vosk Chinese speech model (optional): from "
            "https://alphacephei.com/vosk/models download vosk-model-small-cn-0.22.zip "
            "and extract it to data/vosk-model/vosk-model-small-cn/."
        ),
    },
    "install_p4": {
        "zh": "（4）配置API密钥：编辑 main.py 或 chat_assistant.py，设置以下密钥：",
        "en": "(4) Configure API keys: edit main.py or chat_assistant.py and set the following keys:",
    },
    "install_keys": {
        "zh": (
            "DeepSeek API Key: 编辑 ui/workflows.py 中的 DS_API_KEY 变量\n"
            "Serper.dev API Key: 编辑 utils/web_search.py 中的 SERPER_API_KEY 变量\n"
            "Gemini API (via Yunwu): 编辑 ui/workflows.py 中的 G_API_KEY 变量"
        ),
        "en": (
            "DeepSeek API Key: edit the DS_API_KEY variable in ui/workflows.py\n"
            "Serper.dev API Key: edit the SERPER_API_KEY variable in utils/web_search.py\n"
            "Gemini API (via Yunwu): edit the G_API_KEY variable in ui/workflows.py"
        ),
    },
    "install_p5": {
        "zh": "（5）启动应用：",
        "en": "(5) Launch the application:",
    },
    "install_p5_cmd": {
        "zh": "cd /path/to/海外中式建筑数据库/编程\npython main.py",
        "en": "cd /path/to/Chinoiserie-Architecture-Beyond-China-Database\npython main.py",
    },
    "chrome_ext_setup": {
        "zh": (
            "系统包含两个Chrome浏览器扩展：（1）ext/ ——\"建筑案例提取\"扩展，通过HTTP"
            "协议向本机127.0.0.1:9876端口发送网页正文内容，用于AI提取和核验功能；"
            "（2）auto-refresh-extension/ ——\"自动刷新\"扩展，定时刷新代理页面防止掉线。"
            "扩展的安装和配置说明见对应目录下的文档。如不使用Chrome扩展，AI提取和核验"
            "功能仍可通过手动粘贴文本内容的方式使用。"
        ),
        "en": (
            "The system includes two Chrome browser extensions: (1) ext/ — \"Building Case "
            "Extraction\" extension, which sends page-body content to 127.0.0.1:9876 via "
            "HTTP for AI extraction and verification; (2) auto-refresh-extension/ — "
            "\"Auto-Refresh\" extension, which periodically refreshes proxy pages to "
            "prevent disconnection. Installation and configuration instructions are in "
            "the respective directories. Without the Chrome extensions, AI extraction and "
            "verification can still be used by manually pasting text content."
        ),
    },
    "db_maintenance": {
        "zh": (
            "数据库文件位于 data/architecture.db（约972 MB，17张表），建议定期备份。"
            "系统首次启动时自动创建所有表结构和种子数据（建筑类型字典）。数据库升级通过"
            "migrate_db_v2() 至 migrate_db_v12() 共11个增量迁移自动进行，对用户透明。"
            "如需手动查看数据，可使用任何SQLite客户端（如DB Browser for SQLite）直接"
            "打开.db文件。"
        ),
        "en": (
            "The database file is located at data/architecture.db (~972 MB, 17 tables); "
            "regular backups are recommended. On first launch, the system automatically "
            "creates all table structures and seed data (building-type dictionary). "
            "Database upgrades are performed automatically and transparently through 11 "
            "incremental migrations from migrate_db_v2() through migrate_db_v12(). To "
            "inspect the data manually, any SQLite client (e.g., DB Browser for SQLite) "
            "can open the .db file directly."
        ),
    },

    # ── Ch7 ────────────────────────────────────────────────────────────
    "7.1": {"zh": "8.1  项目成果", "en": "7.1  Achievements"},
    "7.2": {"zh": "8.2  技术难点与解决方案", "en": "7.2  Technical Challenges and Solutions"},
    "7.3": {"zh": "8.3  未来展望", "en": "7.3  Future Directions"},

    "ach_1": {
        "zh": "完成了海外中式建筑数据库的设计与实现，包含12张数据表，支持灵活的字段扩展和完整的CRUD操作；",
        "en": "Designed and implemented the Chinoiserie Architecture Beyond China Database with 12 data tables, supporting flexible field extension and full CRUD operations;",
    },
    "ach_2": {
        "zh": "集成了百度、Google、DeepSeek、Gemini四种核验引擎，实现了字段级的信息真实性验证；",
        "en": "Integrated four verification engines (Baidu, Google, DeepSeek, Gemini) for field-level authenticity verification;",
    },
    "ach_3": {
        "zh": (
            "构建了\"一键搜索\"、\"一键DS核验\"、\"一键G核验\"、\"一键核验\"四个自动化工作流，"
            "将原本需要数十步手动操作的任务简化为一次指令；"
        ),
        "en": (
            "Built four automated workflows (One-Click Search, One-Click DS Verification, "
            "One-Click G Verification, One-Click Cascade Verification) that reduce tasks "
            "requiring dozens of manual steps to a single command;"
        ),
    },
    "ach_4": {
        "zh": (
            "实现了基于AI的聊天助手，支持24条快捷命令、50+种AI动作和离线语音识别，"
            "为研究人员提供了自然的交互体验；"
        ),
        "en": (
            "Implemented an AI-powered Chat Assistant with 24 quick commands, 50+ AI "
            "actions, and offline speech recognition, providing researchers with a "
            "natural interaction experience;"
        ),
    },
    "ach_5": {
        "zh": "提供了多维度数据统计分析和Excel导出功能，支持学术论文的图表制作。",
        "en": "Provided multi-dimensional statistical analysis and Excel export to support chart and table production for academic papers.",
    },

    "chal_1_title": {"zh": "Tkinter线程安全问题", "en": "Tkinter Thread Safety"},
    "chal_1_sol": {
        "zh": "采用 root.after(0, callback) 模式统一调度UI更新，后台线程执行耗时操作，确保UI永不阻塞。",
        "en": "Adopted the root.after(0, callback) pattern for unified UI-update scheduling; background threads execute time-consuming operations, ensuring the UI never blocks.",
    },
    "chal_2_title": {"zh": "AI输出格式不稳定", "en": "Unstable AI Output Format"},
    "chal_2_sol": {
        "zh": "实现多级JSON提取策略（直接解析→Markdown提取→正则匹配→括号修复），大幅降低AI输出解析失败率。",
        "en": "Implemented a multi-level JSON extraction strategy (direct parse → Markdown extraction → regex matching → bracket repair), substantially reducing AI output parsing failures.",
    },
    "chal_3_title": {"zh": "中文PDF PUA字符乱码", "en": "Chinese PDF PUA Character Corruption"},
    "chal_3_sol": {
        "zh": "针对常见中文字体建立PUA→标准Unicode映射表，在PDF提取时自动清理。",
        "en": "Built a PUA → standard-Unicode mapping table for common Chinese fonts, applying automatic cleanup during PDF extraction.",
    },
    "chal_4_title": {"zh": "WPS Excel图片兼容", "en": "WPS Excel Image Compatibility"},
    "chal_4_sol": {
        "zh": "通过直接解析xlsx ZIP内部结构提取DISPIMG格式的图片，解决openpyxl的兼容性问题。",
        "en": "Extracted DISPIMG-format images by directly parsing the internal ZIP structure of xlsx files, resolving openpyxl compatibility issues.",
    },
    "chal_5_title": {"zh": "多语言关键词生成", "en": "Multilingual Keyword Generation"},
    "chal_5_sol": {
        "zh": "预生成104种语言的关键词翻译数据，配合AI实时翻译作为回退方案。",
        "en": "Pre-generated keyword translation data for 104 languages, with real-time AI translation as a fallback.",
    },

    "future_1": {
        "zh": (
            "Web化部署：将桌面应用迁移为Web应用（如Flask/FastAPI后端 + React前端），"
            "支持多用户协同和远程访问；"
        ),
        "en": (
            "Web-based deployment: migrate the desktop application to a web application "
            "(e.g., Flask/FastAPI backend + React frontend) to support multi-user "
            "collaboration and remote access;"
        ),
    },
    "future_2": {
        "zh": (
            "图像识别集成：引入计算机视觉模型（如CLIP），自动识别建筑图片中的中式元素，"
            "提升图片筛选效率；"
        ),
        "en": (
            "Computer-vision integration: incorporate vision models (e.g., CLIP) to "
            "automatically identify Chinoiserie elements in building images, improving "
            "image-screening efficiency;"
        ),
    },
    "future_3": {
        "zh": (
            "知识图谱构建：基于已积累的案例数据，构建海外中式建筑的知识图谱，"
            "分析传播路径和风格演化规律；"
        ),
        "en": (
            "Knowledge graph construction: build a knowledge graph of Chinoiserie "
            "Architecture based on accumulated case data to analyse diffusion paths "
            "and stylistic evolution;"
        ),
    },
    "future_4": {
        "zh": "地图可视化：集成Leaflet或Mapbox，在GIS地图上展示案例的地理分布；",
        "en": "Map visualisation: integrate Leaflet or Mapbox to display the geographic distribution of cases on GIS maps;",
    },
    "future_5": {
        "zh": "更丰富的分析维度：增加建筑风格、材料用途、建筑规模等细分维度的统计分析；",
        "en": "Richer analytical dimensions: add statistical analysis for sub-dimensions such as architectural style, material use, and building scale;",
    },
    "future_6": {
        "zh": "协作功能：支持多研究者共享数据库、评论标注和版本管理。",
        "en": "Collaboration features: support shared databases, commenting, annotation, and version management for multiple researchers.",
    },

    # ── Ch8 ────────────────────────────────────────────────────────────
    "ch8_title": {"zh": "第2章  核心创新点", "en": "8. Core Innovations"},
    "8.1": {"zh": "2.1  数据处理流程核心创新", "en": "2.1  Core Innovations in Data Processing"},
    "8.1.1": {"zh": "2.1.1  多维系统检索策略", "en": "2.1.1  Multi-dimensional Systematic Search Strategy"},
    "8.1.2": {"zh": "2.1.2  灰色文献优先的数据源策略", "en": "2.1.2  Grey-Literature-First Data Sourcing"},
    "8.1.3": {"zh": "2.1.3  LLM双模型协作筛选与人机验证", "en": "2.1.3  Dual-LLM Collaborative Screening with Human Verification"},
    "8.1.4": {"zh": "2.1.4  全链条一键自动化工作流", "en": "2.1.4  End-to-End One-Click Automated Workflow"},
    "8.1.5": {"zh": "2.1.5  多模态自然语言交互", "en": "2.1.5  Multimodal Natural-Language Interaction"},
    "8.2": {"zh": "2.2  效率提升量化分析", "en": "2.2  Quantified Efficiency Improvements"},
    "8.3": {"zh": "2.3  数据库架构总览", "en": "2.3  Database Architecture Overview"},
    "8.4": {"zh": "2.4  数据处理全流程", "en": "2.4  Data Processing Pipeline"},

    "inno_1_p1": {
        "zh": (
            "传统的海外建筑研究依赖人工搜索，研究人员需要逐个访问搜索引擎，手动组合"
            "关键词并逐一阅读结果，覆盖范围有限且重复劳动严重。本系统创新性地构建了"
            "\"国家×语言×关键词\"三维检索矩阵：第一维覆盖全球194个M49国家和地区；"
            "第二维覆盖各国家/地区的官方语言，共计106种语言；第三维覆盖14个核心建筑"
            "关键词（包括中式建筑、Chinese architecture、Chinoiserie Architecture、"
            "Chinese garden、Chinese temple、Chinese pavilion、Chinatown architecture、"
            "Chinoiserie、中国庭園、中華建築、architecture chinoise、arquitectura china、"
            "chinesische Architektur、中国風建築）。通过关键词翻译引擎将14个核心关键词"
            "自动翻译为106种语言并生成国家-语言-关键词的组合矩阵，最终产生7,787条灰色"
            "文献检索式和5,200条学术文献检索式。这是首个系统性地解决跨语言、跨国家建筑"
            "检索问题的方法，从根本上突破了人工检索的语言和地域瓶颈。"
        ),
        "en": (
            "Traditional architectural research relies on manual searching: researchers "
            "visit search engines individually, manually combine keywords, and read each "
            "result sequentially — limited in coverage and burdened by repetitive labour. "
            "This system innovatively constructs a three-dimensional \"country × language × "
            "keyword\" search matrix: the first dimension covers all 194 M49 countries and "
            "areas; the second dimension covers the official languages of each country/area "
            "(106 languages in total); the third dimension covers 14 core architectural "
            "keywords. A keyword translation engine automatically translates the 14 core "
            "keywords into all 106 languages and generates the country-language-keyword "
            "combination matrix, yielding 7,787 grey literature search expressions and "
            "5,200 scholarly literature search expressions. This is the first method to "
            "systematically solve the problem of cross-lingual, cross-national architectural "
            "search, fundamentally overcoming the linguistic and geographic bottlenecks of "
            "manual retrieval."
        ),
    },
    "inno_1_p2": {
        "zh": (
            "与传统的固定关键词搜索方式相比，本系统的三维矩阵策略具有三个关键优势："
            "（1）覆盖完整性——确保每个国家的每个官方语言都有对应的检索式，不会遗漏"
            "非英语信息源；（2）搜索引擎互补性——同时使用Google搜索（全球多语言优势）"
            "和百度搜索（中文内容优势），两者返回的内容类型不同（Google偏向旅游网站和"
            "多语言新闻，百度偏向中文博客、政府页面和社区媒体），双引擎设计有效弥补了"
            "单一引擎的覆盖盲区；（3）可复现性——检索矩阵的全部组合可追溯、可审查、"
            "可复现，满足了科学研究对透明度和可重复性的要求。"
        ),
        "en": (
            "Compared with conventional fixed-keyword search approaches, the "
            "three-dimensional matrix strategy offers three key advantages: (1) coverage "
            "completeness — ensuring that every country's every official language has a "
            "corresponding search expression, leaving no non-English information source "
            "uncovered; (2) search-engine complementarity — simultaneously using Google "
            "Search (global multi-lingual strength) and Baidu Search (Chinese-language "
            "content strength), whose returned content types differ (Google favours tourism "
            "websites and multi-lingual news; Baidu surfaces Chinese-language blogs, "
            "government pages, and community media), with the dual-engine design effectively "
            "compensating for single-engine coverage gaps; (3) reproducibility — all "
            "combinations in the search matrix are traceable, auditable, and reproducible, "
            "meeting the scientific requirements for transparency and replicability."
        ),
    },

    "inno_2_p1": {
        "zh": (
            "传统建筑学研究以学术文献为主要数据源，但在\"海外中式建筑\"这一研究对象中，"
            "公众感知才是判断建筑是否具有中国风格的关键。学术文献反映的是专家视角，"
            "而大量关于建筑的日常描述和公众评价存在于旅游博客、新闻报道、旅游网站和"
            "地方政府页面等灰色文献中。本系统创新性地提出\"灰色文献优先\"的数据源策略："
            "以灰色文献为主要数据源捕获公众感知，以学术文献为辅助参考提供历史深度和"
            "交叉验证。这一策略从方法论层面解决了\"谁能定义什么是中国风建筑\"的核心问题"
            "——让公众定义，而非专家预设。这一方法论的创新对于任何以公众感知为研究对象的"
            "建筑类型（如乡土建筑、侨民宗教建筑、殖民建筑遗产等）均具有迁移价值。"
        ),
        "en": (
            "Traditional architectural research treats scholarly literature as the primary "
            "data source. However, for the subject of \"Chinoiserie Architecture,\" public "
            "perception — rather than expert judgment — is the key to determining whether "
            "a building qualifies as Chinese-style. Scholarly literature reflects an expert "
            "perspective, whereas the vast body of everyday descriptions and public "
            "evaluations of buildings resides in grey literature — travel blogs, news "
            "reports, tourism websites, and local government pages. This system innovatively "
            "proposes a \"grey-literature-first\" data-sourcing strategy: grey literature "
            "serves as the primary data source for capturing public perception, while "
            "scholarly literature provides historical depth and corroboration as a "
            "supplementary reference. This strategy resolves, at the methodological level, "
            "the core question of \"who defines what counts as Chinoiserie Architecture\" — "
            "letting the public define, rather than experts prescribe. This methodological "
            "innovation is transferable to any architectural category where public perception "
            "is the object of study."
        ),
    },

    "inno_3_p1": {
        "zh": (
            "面对85万条检索结果的筛选挑战，完全人工筛选不可行（估算需要数十人年）。"
            "本系统创新性地设计了\"LLM双模型协作筛选 + 统计抽样验证\"的人机混合筛选方案："
            "使用DeepSeek-Chat模型筛选灰色文献（因其对非结构化网络文本的理解能力更强），"
            "使用Gemini-2.5-Flash模型筛选学术文献（因其对学术语言和引用格式的识别更"
            "准确）。两个模型各自独立运行，筛选结果连同原始数据一并保留，实现全链条可"
            "追溯。"
        ),
        "en": (
            "Faced with the challenge of screening 850,000+ search results, fully manual "
            "screening is infeasible (estimated at dozens of person-years). This system "
            "innovatively designs a \"dual-LLM collaborative screening + statistical "
            "sampling verification\" human-machine hybrid approach: DeepSeek-Chat screens "
            "grey literature (better at understanding unstructured web text), while "
            "Gemini-2.5-Flash screens scholarly literature (more accurate at recognizing "
            "academic language and citation formats). The two models operate independently, "
            "with screening decisions and raw data preserved together for full-chain "
            "traceability."
        ),
    },
    "inno_3_p2": {
        "zh": (
            "为量化验证筛选质量，设计了严格的分层随机抽样方案：从LLM排除的569,533条"
            "记录中抽取3,000条（灰文1,500条、学术1,500条），每种文献类型内随机均分为"
            "三组（每组500条）以控制标注批次效应。两名独立标注员进行双盲标注，标注员"
            "间一致性达κ=0.980（近乎完美）。结果显示灰文假阴性率1.20%、学术假阴性率"
            "0.067%，估计记录级召回率98.2%。这一验证方案不仅证明了LLM筛选的可靠性，"
            "也为同类研究提供了一套完整的筛选质量评估方法论。"
        ),
        "en": (
            "To quantify screening quality, a rigorous stratified random sampling design "
            "was implemented: from the 569,533 records excluded by the LLM, 3,000 were "
            "sampled (1,500 grey literature, 1,500 scholarly literature), with each type "
            "randomly divided into three groups of 500 to control for annotation batch "
            "effects. Two independent annotators performed double-blind labelling, achieving "
            "near-perfect inter-annotator agreement (κ = 0.980). Results showed a false "
            "negative rate of 1.20% for grey literature and 0.067% for scholarly literature, "
            "yielding an estimated record-level recall of 98.2%. This validation design not "
            "only demonstrates the reliability of LLM screening but also provides a complete "
            "methodological template for screening-quality assessment in similar studies."
        ),
    },

    "inno_4_p1": {
        "zh": (
            "传统的数据采集流程中，检索、筛选、提取、去重、验证、评估、入库等步骤各自"
            "独立，需要研究人员手动切换工具并按顺序执行，操作繁琐、容易出错，且无法批量"
            "处理。本系统创新性地设计了基于有限状态机（FSM）的工作流引擎，将上述7个"
            "独立步骤串联为可一键执行的自动化管道。核心设计包括：（1）状态驱动的流程"
            "控制——每个步骤完成后自动触发下一步，步骤间通过共享上下文传递数据；（2）"
            "断点恢复机制——任何步骤失败后可从中断点继续，不丢失已完成工作；（3）全量"
            "操作日志——每一步操作的输入、输出和决策均记录在SQLite数据库中，支持事后"
            "审计和流程复现。"
        ),
        "en": (
            "In traditional data collection workflows, steps such as search, screening, "
            "extraction, deduplication, verification, evaluation, and ingestion are each "
            "independent, requiring researchers to manually switch tools and execute "
            "sequentially — cumbersome, error-prone, and incompatible with batch processing. "
            "This system innovatively designs a workflow engine based on a Finite-State "
            "Machine (FSM) that chains the seven independent steps into an automated pipeline "
            "executable with a single command. Core designs include: (1) state-driven process "
            "control — each step automatically triggers the next upon completion, with data "
            "passed between steps via shared context; (2) checkpoint recovery — any failed "
            "step can resume from the interruption point without losing completed work; "
            "(3) comprehensive operation logging — the input, output, and decision of every "
            "operation are recorded in the SQLite database, supporting post-hoc audit and "
            "process reproduction."
        ),
    },
    "inno_4_p2": {
        "zh": (
            "该工作流引擎定义了四个核心工作流：\"一键搜索工作流\"（关键词生成→API检索→"
            "LLM筛选→去重→入库）、\"一键核验工作流\"（DeepSeek核验→Gemini回退→人工"
            "确认）、\"DS核验工作流\"和\"G核验工作流\"。用户通过聊天助手的快捷命令即可触发"
            "（如\"@一键核验\"），系统自动完成全部流程并返回结构化结果。将传统需要数周"
            "的人工操作压缩至小时级。"
        ),
        "en": (
            "The workflow engine defines four core workflows: \"One-Click Search\" (keyword "
            "generation → API retrieval → LLM screening → deduplication → ingestion), "
            "\"One-Click Verification\" (DeepSeek verification → Gemini fallback → manual "
            "confirmation), \"DS Verification,\" and \"G Verification.\" Users trigger them "
            "through the Chat Assistant's quick commands (e.g., \"@一键核验\"), and the "
            "system automatically completes the entire process and returns structured "
            "results — compressing what traditionally required weeks of manual work into "
            "hours."
        ),
    },

    "inno_6_p1": {
        "zh": (
            "传统数据库软件依赖图形界面操作（鼠标点击、菜单导航），对非技术用户存在"
            "较高的使用门槛。本系统创新性地集成了离线语音识别引擎（Vosk）和AI自然语言"
            "理解模块，使研究人员可以通过语音指令或自然语言文本完成复杂的数据库操作。"
            "核心技术实现包括：（1）24条正则表达式快捷命令——覆盖搜索、核验、导入导出、"
            "导航等高频操作，输入命令名即可一键执行；（2）50+种AI动作——通过自然语言"
            "描述意图，AI自动解析并调用对应的系统功能；（3）离线语音识别——基于Vosk"
            "中文语音模型，无需联网即可将语音指令转为文本并执行。这种多模态交互设计"
            "将数据库操作门槛从\"需要编程技能\"降低至\"会说话就会用\"。"
        ),
        "en": (
            "Traditional database software relies on GUI operations (mouse clicks, menu "
            "navigation), posing a high barrier to entry for non-technical users. This "
            "system innovatively integrates an offline speech recognition engine (Vosk) "
            "and an AI natural-language understanding module, enabling researchers to "
            "perform complex database operations through voice commands or natural-language "
            "text. Core technical implementations include: (1) 24 regex-based quick commands "
            "covering high-frequency operations such as search, verification, import/export, "
            "and navigation — executed with a single command name; (2) 50+ AI actions — "
            "users describe their intent in natural language, and the AI automatically "
            "parses it and invokes the corresponding system function; (3) offline speech "
            "recognition — based on the Vosk Chinese speech model, voice commands are "
            "converted to text and executed without an internet connection. This multimodal "
            "interaction design reduces the database operation threshold from \"requires "
            "programming skills\" to \"speak and use.\""
        ),
    },

    "eff_table_title": {"zh": "表8-1  自动化处理效率对比", "en": "Table 8-1  Automated Processing Efficiency Comparison"},
    "eff_col_stage": {"zh": "处理阶段", "en": "Stage"},
    "eff_col_task": {"zh": "任务规模", "en": "Task Scale"},
    "eff_col_manual": {"zh": "人工估算", "en": "Manual Estimate"},
    "eff_col_auto": {"zh": "自动化处理", "en": "Automated Processing"},
    "eff_col_saving": {"zh": "效率提升", "en": "Efficiency Gain"},
    "eff_search": {"zh": "系统检索", "en": "Systematic Search"},
    "eff_search_task": {"zh": "12,987条检索式", "en": "12,987 search expressions"},
    "eff_search_manual": {"zh": "~650小时（手动逐条输入并阅读首页结果）", "en": "~650 h (manual input and first-page reading)"},
    "eff_search_auto": {"zh": "~3小时（API并发执行，含速率限制等待）", "en": "~3 h (API concurrent execution with rate limiting)"},
    "eff_search_saving": {"zh": "~99.5%", "en": "~99.5%"},
    "eff_screening": {"zh": "LLM初筛", "en": "LLM Pre-Screening"},
    "eff_screening_task": {"zh": "851,795条记录", "en": "851,795 records"},
    "eff_screening_manual": {"zh": "~17,700小时（按每30秒阅读判断一条）", "en": "~17,700 h (at 30 s per record)"},
    "eff_screening_auto": {"zh": "~8小时（两模型并行，含API调用时间）", "en": "~8 h (dual-model parallel execution)"},
    "eff_screening_saving": {"zh": "~99.9%", "en": "~99.9%"},
    "eff_extraction": {"zh": "案例提取", "en": "Case Extraction"},
    "eff_extraction_task": {"zh": "35,800条通过筛选", "en": "35,800 passed screening"},
    "eff_extraction_manual": {"zh": "~2,980小时（按每5分钟阅读提取一条）", "en": "~2,980 h (at 5 min per record)"},
    "eff_extraction_auto": {"zh": "~3小时（DeepSeek批量提取）", "en": "~3 h (DeepSeek batch extraction)"},
    "eff_extraction_saving": {"zh": "~99.9%", "en": "~99.9%"},
    "eff_verification": {"zh": "字段核验", "en": "Field Verification"},
    "eff_verification_task": {"zh": "686案例 × 8字段 = 5,488次核验", "en": "686 cases × 8 fields = 5,488 verifications"},
    "eff_verification_manual": {"zh": "~915小时（按每10分钟核验一个字段）", "en": "~915 h (at 10 min per field)"},
    "eff_verification_auto": {"zh": "~2小时（LLM多源交叉比对 + Geocoding API）", "en": "~2 h (LLM multi-source cross-check + Geocoding API)"},
    "eff_verification_saving": {"zh": "~99.8%", "en": "~99.8%"},
    "eff_total": {"zh": "合计", "en": "Total"},
    "eff_total_task": {"zh": "全流程", "en": "End-to-end pipeline"},
    "eff_total_manual": {"zh": "~22,245小时（≈2.5人年）", "en": "~22,245 h (≈2.5 person-years)"},
    "eff_total_auto": {"zh": "~16小时（含人工复核时间）", "en": "~16 h (including manual review)"},
    "eff_total_saving": {"zh": "~99.9%", "en": "~99.9%"},
    "eff_note": {
        "zh": (
            "注：人工估算基于合理假设——每条记录平均阅读时间30秒（初筛）、每案例提取"
            "5分钟、每字段核验10分钟。实际人工操作中，由于疲劳、注意力分散和工具切换"
            "成本，实际时间可能更长。自动化处理时间含API速率限制等待和人工复核环节。"
        ),
        "en": (
            "Note: Manual estimates are based on reasonable assumptions — 30 s per record "
            "for initial screening, 5 min per case for extraction, 10 min per field for "
            "verification. In actual manual operation, times may be longer due to fatigue, "
            "attention drift, and tool-switching overhead. Automated processing times "
            "include API rate-limit waiting and manual review steps."
        ),
    },

    "arch_schema_title": {"zh": "表8-2  数据库核心表架构", "en": "Table 8-2  Core Database Table Architecture"},
    "arch_col_table": {"zh": "表名", "en": "Table"},
    "arch_col_fields": {"zh": "字段数", "en": "Fields"},
    "arch_col_rows": {"zh": "记录数", "en": "Records"},
    "arch_col_purpose": {"zh": "用途", "en": "Purpose"},
    "arch_col_relations": {"zh": "关联关系", "en": "Relationships"},
    "arch_rel_cases": {
        "zh": "主表。FK引用field_sources、case_sources、case_images",
        "en": "Main table. FK referenced by field_sources, case_sources, case_images",
    },
    "arch_raw_cases_rows": {"zh": "~3,500", "en": "~3,500"},
    "arch_rel_raw": {
        "zh": "FK → cases.id (promote后)。FK引用verification_log",
        "en": "FK → cases.id (after promote). FK referenced by verification_log",
    },
    "arch_cs_rows": {"zh": "~1,200", "en": "~1,200"},
    "arch_cs_purpose": {"zh": "案例来源URL存储", "en": "Case source URL storage"},
    "arch_fs_rows": {"zh": "~5,500", "en": "~5,500"},
    "arch_fs_purpose": {"zh": "字段级溯源信息", "en": "Field-level provenance"},
    "arch_vl_rows": {"zh": "~16,000", "en": "~16,000"},
    "arch_vl_purpose": {"zh": "核验操作审计日志", "en": "Verification audit log"},
    "arch_ol_rows": {"zh": "~50,000", "en": "~50,000"},
    "arch_ol_purpose": {"zh": "通用操作审计日志", "en": "General operation audit log"},
    "arch_lit_rows": {"zh": "~1,200", "en": "~1,200"},
    "arch_lit_purpose": {"zh": "学术文献元数据", "en": "Scholarly literature metadata"},
    "arch_lml_rows": {"zh": "~800", "en": "~800"},
    "arch_lml_purpose": {"zh": "文献-案例匹配记录", "en": "Literature–case match records"},
    "arch_fr_rows": {"zh": "~215,000", "en": "~215,000"},
    "arch_rs_rows": {"zh": "~521,000", "en": "~521,000"},
    "arch_ssr_rows": {"zh": "~330,000", "en": "~330,000"},
    "arch_bt_rows": {"zh": "11", "en": "11"},
    "arch_bt_purpose": {"zh": "建筑类型字典", "en": "Building-type dictionary"},
    "arch_img_rows": {"zh": "~2,500", "en": "~2,500"},
    "arch_img_purpose": {"zh": "图片存储/回收站", "en": "Image storage / recycle bin"},
    "arch_vc_rows": {"zh": "~700", "en": "~700"},
    "arch_dr_rows": {"zh": "~170,000", "en": "~170,000"},

    "pipeline_note": {
        "zh": (
            "数据处理全流程如图1-1所示（详见软件主体界面\"案例查询\"页签）。流程涵盖"
            "四个阶段：数据源识别（灰文+学术双源策略）、原始数据获取（多语言系统检索"
            "矩阵→API自动执行→原始结果入库）、数据筛选（LLM双模型初筛→去重→人工复核"
            "→全文获取）和案例提取（LLM结构化提取→自动去重→人工合并→AHP评估→正式入库）"
            "。每个阶段的输入、输出和中间数据均完整保留在项目数据库中，实现全链条可审计。"
        ),
        "en": (
            "The complete data processing pipeline is illustrated in Figure 1-1 (see the "
            "\"Case Search\" tab of the main application interface). The pipeline comprises "
            "four stages: data source identification (grey + scholarly dual-source strategy), "
            "raw data acquisition (multilingual systematic search matrix → API automated "
            "execution → raw results ingestion), data screening (LLM dual-model pre-screening "
            "→ deduplication → manual review → full-text retrieval), and case extraction "
            "(LLM structured extraction → automated deduplication → manual merging → AHP "
            "evaluation → formal ingestion). The input, output, and intermediate data of "
            "every stage are fully preserved in the project database, enabling end-to-end "
            "auditability."
        ),
    },

    # ── Appendix A ─────────────────────────────────────────────────────
    "a1": {"zh": "A.1  cases 表", "en": "A.1  cases Table"},
    "a2": {"zh": "A.2  raw_cases 表", "en": "A.2  raw_cases Table"},
    "a3": {"zh": "A.3  辅助表", "en": "A.3  Auxiliary Tables"},

    "app_a1_desc": {
        "zh": (
            "cases表存储正式入库的建筑案例。字段包括：id, name, name_en, name_official, "
            "country, city, year_built, demolition_year, restoration_year, quantity, "
            "building_type, architect, latitude, longitude, material, area_sqm, "
            "protection_level, description, image_path, image_urls, search_time, "
            "search_keywords, condition, spread_method, notes, verification_status, "
            "completeness_score, source_credibility, raw_case_ids, sort_order, "
            "created_at, updated_at。总计31个字段。"
        ),
        "en": (
            "The cases table stores verified architectural cases. Fields: id, name, "
            "name_en, name_official, country, city, year_built, demolition_year, "
            "restoration_year, quantity, building_type, architect, latitude, longitude, "
            "material, area_sqm, protection_level, description, image_path, image_urls, "
            "search_time, search_keywords, condition, spread_method, notes, "
            "verification_status, completeness_score, source_credibility, raw_case_ids, "
            "sort_order, created_at, updated_at. Total: 31 fields."
        ),
    },
    "app_a2_desc": {
        "zh": (
            "raw_cases表存储待核验的原始数据。固定字段26个，支持运行时动态添加自定义字段。"
            "字段映射规则（promote_raw_to_case时）：construction_year→year_built, "
            "function→building_type, designer→architect。"
        ),
        "en": (
            "The raw_cases table stores unverified raw data. It has 26 fixed fields and "
            "supports runtime addition of custom fields. Field mapping rules (on "
            "promote_raw_to_case): construction_year → year_built, function → "
            "building_type, designer → architect."
        ),
    },
    "app_a3_desc": {
        "zh": (
            "building_types: 建筑类型字典（园林、民居、寺庙、宫殿、衙署、塔、桥、祠堂、"
            "书院、现代中式、其他）。case_sources: 案例来源URL（外键关联raw_cases）。"
            "field_sources: 字段级溯源信息。case_images / image_trash: 图片存储和回收站。"
            "verification_log: 核验操作日志。operation_log: 通用操作审计日志。"
            "literature: 学术文献。lit_match_log: 文献-案例匹配日志。"
            "filtered_results / raw_search_results / scholar_search_results: 搜索结果存储（三层）。"
            "verified_cases: 已验证案例快照。dedup_removed: 去重记录。"
        ),
        "en": (
            "building_types: building-type dictionary (garden, residence, temple, palace, "
            "government office, pagoda, bridge, ancestral hall, academy, modern "
            "Chinoiserie, other). case_sources: case source URLs (FK → raw_cases). "
            "field_sources: field-level provenance. case_images / image_trash: image "
            "storage and recycle bin. verification_log: verification audit log. "
            "operation_log: general operation audit log. literature: academic literature. "
            "lit_match_log: literature–case match log. filtered_results / raw_search_results "
            "/ scholar_search_results: search-result storage (three tiers). "
            "verified_cases: verified-cases snapshot. dedup_removed: deduplication records."
        ),
    },

    # ── Appendix B ─────────────────────────────────────────────────────
    "api_col_name": {"zh": "API", "en": "API"},
    "api_col_endpoint": {"zh": "端点", "en": "Endpoint"},
    "api_col_model": {"zh": "模型", "en": "Model"},
    "api_col_purpose": {"zh": "用途", "en": "Purpose"},
    "api_deepseek_purpose": {
        "zh": "案例提取、名称核验、字段补全、自然语言理解",
        "en": "Case extraction, name verification, field completion, natural-language understanding",
    },
    "api_gemini_purpose": {
        "zh": "多语言名称核验、字段补全、图片搜索",
        "en": "Multilingual name verification, field completion, image search",
    },
    "api_serper_purpose": {
        "zh": "Google结构化搜索（通过HTTP代理127.0.0.1:51081）",
        "en": "Google structured search (via HTTP proxy 127.0.0.1:51081)",
    },
    "api_baidu_purpose": {
        "zh": "百度结构化搜索",
        "en": "Baidu structured search",
    },
    "api_keys_note": {
        "zh": (
            "API 密钥统一存放在本地配置文件 config.local.json 中（该文件不随代码公开），"
            "可通过「设置 → API 配置」对话框集中编辑；程序运行时经 config.py 读取各通道"
            "的地址、模型与密钥。config.local.json 已被 .gitignore 排除，不会进入公开仓库。"
        ),
        "en": (
            "API keys are stored in a local configuration file, config.local.json "
            "(excluded from the public repository), and can be edited centrally via "
            "the Settings → API Configuration dialog. At runtime, each channel's "
            "endpoint, model, and key are read through config.py. The config.local.json "
            "file is excluded via .gitignore and will not be published."
        ),
    },

    # ── Appendix C ─────────────────────────────────────────────────────
    "cmd_col_shortcut": {"zh": "命令/快捷键", "en": "Command / Shortcut"},
    "cmd_col_function": {"zh": "功能", "en": "Function"},
    "cmd_ctrl_b": {"zh": "切换聊天助手面板显示/隐藏", "en": "Toggle Chat Assistant panel show/hide"},
    "cmd_search": {"zh": "在搜索引擎中搜索", "en": "Search in search engines"},
    "cmd_verify": {"zh": "启动一键核验工作流（DS→G级联）", "en": "Launch One-Click Cascade Verification workflow (DS → G fallback)"},
    "cmd_ds_verify": {"zh": "启动DeepSeek核验工作流", "en": "Launch DeepSeek Verification workflow"},
    "cmd_g_verify": {"zh": "启动Gemini核验工作流", "en": "Launch Gemini Verification workflow"},
    "cmd_oneclick_search": {"zh": "对该国家+语言执行全关键词搜索", "en": "Execute full-keyword search for the specified country + language"},
    "cmd_switch_tab": {"zh": "切换主标签页", "en": "Switch main tab"},
    "cmd_refresh": {"zh": "刷新当前列表", "en": "Refresh current list"},
    "cmd_import_export": {"zh": "导入或导出数据", "en": "Import or export data"},
    "cmd_autocomplete": {"zh": "AI自动补全当前案例缺失字段", "en": "AI auto-complete missing fields for current case"},
    "cmd_delete": {"zh": "删除当前选中的记录", "en": "Delete currently selected record"},
    "cmd_add_case": {"zh": "打开新建案例对话框", "en": "Open New Case dialog"},
    "cmd_view_detail": {"zh": "查看选中案例详细信息", "en": "View details of selected case"},
    "cmd_nav": {"zh": "导航到下一个/上一个案例", "en": "Navigate to next / previous case"},
    "cmd_submit_verify": {"zh": "将选中项提交到核验队列", "en": "Submit selected items to verification queue"},
    "cmd_approve": {"zh": "核验通过，提升到正式库", "en": "Approve verification and promote to verified database"},
    "cmd_reject": {"zh": "核验拒绝", "en": "Reject verification"},
    "cmd_select_all": {"zh": "选中或取消所有列表项", "en": "Select all / deselect all list items"},
    "cmd_cancel": {"zh": "取消当前运行的工作流", "en": "Cancel currently running workflow"},
    "cmd_help": {"zh": "显示可用命令列表", "en": "Display list of available commands"},

    # ── Appendix D ─────────────────────────────────────────────────────
    "migration_v1": {
        "zh": "V1.0（初始版本）：创建cases和building_types表。",
        "en": "V1.0 (initial release): Created cases and building_types tables.",
    },
    "migration_v2": {
        "zh": "V2.0（migrate_db_v2）：新增raw_cases等6张表，cases表增加17列，执行source_urls数据迁移。",
        "en": "V2.0 (migrate_db_v2): Added 6 tables including raw_cases, added 17 columns to cases, performed source_urls data migration.",
    },
    "migration_v3_v12": {
        "zh": (
            "V3-V12（migrate_db_v3 至 migrate_db_v12）："
            "v3新增literature表；v4新增filtered_results表；v5新增raw_search_results表；"
            "v6新增dedup_removed表；v7新增case_images和image_trash表；"
            "v8新增verification_log表；v9新增field_sources表；"
            "v10新增operation_log和lit_match_log表；v11新增scholar_search_results表；"
            "v12新增verified_cases表并优化索引。"
        ),
        "en": (
            "V3–V12 (migrate_db_v3 through migrate_db_v12): "
            "v3 added literature table; v4 added filtered_results; v5 added raw_search_results; "
            "v6 added dedup_removed; v7 added case_images and image_trash; "
            "v8 added verification_log; v9 added field_sources; "
            "v10 added operation_log and lit_match_log; v11 added scholar_search_results; "
            "v12 added verified_cases with index optimisation."
        ),
    },
    "migration_note": {
        "zh": (
            "迁移机制设计为幂等：每个ALTER/DROP/CREATE语句包裹在try/except中，"
            "确保重复执行不会报错，兼容从任意旧版本直接升级到最新版本。"
        ),
        "en": (
            "The migration mechanism is designed to be idempotent: each ALTER/DROP/CREATE "
            "statement is wrapped in try/except, ensuring repeated execution will not "
            "error and enabling direct upgrade from any previous version to the latest."
        ),
    },

    # ── Misc ───────────────────────────────────────────────────────────
    "cmd_search_keyword": {"zh": "搜索 <关键词>", "en": "search <keyword>"},
    "cmd_verify_cmd": {"zh": "核验 / 一键核验", "en": "verify / cascade-verify"},
    "cmd_ds_cmd": {"zh": "一键DS核验 / ds核验", "en": "ds-verify"},
    "cmd_g_cmd": {"zh": "一键G核验 / g核验", "en": "g-verify"},
    "cmd_oneclick_cmd": {"zh": "一键搜索 <国家> <语言>", "en": "one-click-search <country> <language>"},
    "cmd_switch_cmd": {"zh": "切换到搜索页/整理页/分析页", "en": "switch to search/curation/analysis"},
    "cmd_refresh_cmd": {"zh": "刷新 / 更新列表", "en": "refresh / update"},
    "cmd_import_cmd": {"zh": "导入 / 导出", "en": "import / export"},
    "cmd_autocomplete_cmd": {"zh": "自动补全 / 一键补全", "en": "autocomplete / one-click-complete"},
    "cmd_delete_cmd": {"zh": "删除选中 / 删除", "en": "delete"},
    "cmd_add_cmd": {"zh": "添加案例 / 新建案例", "en": "add case / new case"},
    "cmd_detail_cmd": {"zh": "查看详情 / 详情", "en": "view details / details"},
    "cmd_nav_cmd": {"zh": "下一个 / 上一个", "en": "next / previous"},
    "cmd_submit_cmd": {"zh": "提交核验 / 送审", "en": "submit for verification"},
    "cmd_approve_cmd": {"zh": "批准 / 通过 / approve", "en": "approve"},
    "cmd_reject_cmd": {"zh": "拒绝 / reject / 驳回", "en": "reject"},
    "cmd_select_all_cmd": {"zh": "全选 / 取消全选", "en": "select all / deselect all"},
    "cmd_cancel_cmd": {"zh": "取消 / 停止 / 暂停", "en": "cancel / stop / pause"},
    "cmd_help_cmd": {"zh": "帮助 / 你能做什么", "en": "help / what can you do"},

    "done_msg": {
        "zh": "报告已生成：{path}",
        "en": "Report generated: {path}",
    },
}


# ---------------------------------------------------------------------------
# ReportGenerator
# ---------------------------------------------------------------------------

class ReportGenerator:
    def __init__(self, lang=None):
        self.lang = lang or LANG
        self.doc = Document()
        self._setup_styles()
        self._setup_page()

    def _t(self, key, **kwargs):
        """Return the string for `key` in the current language.  Falls back to
        the key itself if not found.  ``**kwargs`` are passed to str.format()."""
        entry = T.get(key, {})
        text = entry.get(self.lang) if entry else None
        if text is None:
            # Fallback: try the other language
            other = "zh" if self.lang == "en" else "en"
            text = entry.get(other, key) if entry else key
        if kwargs:
            text = text.format(**kwargs)
        return text

    def _setup_page(self):
        section = self.doc.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    def _setup_styles(self):
        style = self.doc.styles['Normal']
        style.font.name = BODY_FONT
        style.font.size = Pt(12)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)
        pf = style.paragraph_format
        pf.line_spacing = 1.5
        pf.space_after = Pt(6)
        pf.first_line_indent = Cm(0.74)

        for level, (size, bold) in enumerate([(22, True), (16, True), (14, True)], 1):
            hstyle = self.doc.styles[f'Heading {level}']
            hstyle.font.name = HEADING_FONT
            hstyle.font.size = Pt(size)
            hstyle.font.bold = bold
            hstyle.font.color.rgb = RGBColor(0, 0, 0)
            hstyle.element.rPr.rFonts.set(qn('w:eastAsia'), HEADING_FONT)
            hstyle.paragraph_format.first_line_indent = Cm(0)
            hstyle.paragraph_format.space_before = Pt(18 if level == 1 else 12)
            hstyle.paragraph_format.space_after = Pt(12 if level == 1 else 6)

    def _add_heading(self, text, level=1):
        h = self.doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.name = HEADING_FONT
            run._element.rPr.rFonts.set(qn('w:eastAsia'), HEADING_FONT)
        return h

    def _add_para(self, text, bold=False, indent=True, size=12, align=None):
        p = self.doc.add_paragraph()
        if not indent:
            p.paragraph_format.first_line_indent = Cm(0)
        if align is not None:
            p.alignment = align
        run = p.add_run(text)
        run.font.name = BODY_FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        run._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)
        return p

    def _add_bullet(self, text):
        p = self.doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(text)
        run.font.name = BODY_FONT
        run.font.size = Pt(12)
        run._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)
        return p

    def _add_code(self, code_text):
        p = self.doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F0F0" w:val="clear"/>')
        p.paragraph_format.element.get_or_add_pPr().append(shading_elm)
        run = p.add_run(code_text)
        run.font.name = CODE_FONT
        run.font.size = Pt(9)
        return p

    def _add_table(self, headers, rows, col_widths=None):
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = ''
            p = hdr_cells[i].paragraphs[0]
            p.paragraph_format.first_line_indent = Cm(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(h)
            run.font.name = BODY_FONT
            run.font.size = Pt(10)
            run.font.bold = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9D9D9" w:val="clear"/>')
            hdr_cells[i]._element.get_or_add_tcPr().append(shading)

        for r, row in enumerate(rows):
            cells = table.rows[r + 1].cells
            for c, val in enumerate(row):
                cells[c].text = ''
                p = cells[c].paragraphs[0]
                p.paragraph_format.first_line_indent = Cm(0)
                run = p.add_run(str(val) if val is not None else '')
                run.font.name = BODY_FONT
                run.font.size = Pt(10)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), BODY_FONT)

        if col_widths:
            for i, w in enumerate(col_widths):
                for row in table.rows:
                    row.cells[i].width = Cm(w)

        self.doc.add_paragraph()
        return table

    # -------------------------------------------------------------------
    # Chapter generators
    # -------------------------------------------------------------------

    def build_cover(self):
        for _ in range(6):
            self.doc.add_paragraph()

        self._add_para(self._t("cover_title_main"), bold=True, indent=False, size=28,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
        self._add_para(self._t("cover_title_sub"), bold=True, indent=False, size=20,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        for _ in range(4):
            self.doc.add_paragraph()

        self._add_para(self._t("cover_version"), indent=False, size=14,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
        if self.lang == "zh":
            date_str = datetime.now().strftime('%Y年%m月%d日')
        else:
            date_str = datetime.now().strftime('%d %B %Y')
        self._add_para(self._t("cover_date", date=date_str), indent=False, size=14,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
        self._add_para(self._t("cover_env"), indent=False, size=12,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        self.doc.add_page_break()

    def build_toc(self):
        self._add_heading(self._t("toc_title"), level=1)
        p = self.doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        run = p.add_run()
        fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        run._element.append(fldChar1)
        run2 = p.add_run()
        instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>')
        run2._element.append(instrText)
        run3 = p.add_run()
        fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
        run3._element.append(fldChar2)
        run4 = p.add_run(self._t("toc_hint"))
        run4.font.name = BODY_FONT
        run4.font.size = Pt(10)
        run4.font.color.rgb = RGBColor(128, 128, 128)
        run5 = p.add_run()
        fldChar3 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        run5._element.append(fldChar3)
        self.doc.add_page_break()

    def build_ch1_overview(self):
        self._add_heading(self._t("ch1_title"), level=1)
        self._add_heading(self._t("1.1"), level=2)
        self._add_para(self._t("bg_p1"))
        self._add_para(self._t("bg_p2"))

        self._add_heading(self._t("1.2"), level=2)
        for k in ["obj_1", "obj_2", "obj_3", "obj_4", "obj_5", "obj_6"]:
            self._add_bullet(self._t(k))

        self._add_heading(self._t("1.3"), level=2)
        self._add_para(self._t("func_overview"))
        for name_k, desc_k in [("func_1_name", "func_1_desc"),
                                ("func_2_name", "func_2_desc"),
                                ("func_3_name", "func_3_desc"),
                                ("func_4_name", "func_4_desc"),
                                ("func_5_name", "func_5_desc")]:
            self._add_bullet(f"{self._t(name_k)}：{self._t(desc_k)}")

        self._add_heading(self._t("1.4"), level=2)
        tech_headers = [self._t("tech_tier"), self._t("tech_choice"), self._t("tech_note")]
        tech_rows = [
            [self._t("tech_lang"), self._t("tech_lang_val"), self._t("tech_lang_desc")],
            [self._t("tech_gui"), self._t("tech_gui_val"), self._t("tech_gui_desc")],
            [self._t("tech_db"), self._t("tech_db_val"), self._t("tech_db_desc")],
            [self._t("tech_ai"), self._t("tech_ai_val"), self._t("tech_ai_desc")],
            [self._t("tech_voice"), self._t("tech_voice_val"), self._t("tech_voice_desc")],
            [self._t("tech_search"), self._t("tech_search_val"), self._t("tech_search_desc")],
            [self._t("tech_pdf"), self._t("tech_pdf_val"), self._t("tech_pdf_desc")],
            [self._t("tech_analysis"), self._t("tech_analysis_val"), self._t("tech_analysis_desc")],
            [self._t("tech_docgen"), self._t("tech_docgen_val"), self._t("tech_docgen_desc")],
        ]
        self._add_table(tech_headers, tech_rows, col_widths=[2.5, 5, 8])
        self.doc.add_page_break()

    def build_ch2_architecture(self):
        self._add_heading(self._t("ch2_title"), level=1)

        self._add_heading(self._t("2.1"), level=2)
        self._add_para(self._t("arch_p1"))

        arch_headers = [self._t("tech_tier"), self._t("arch_col_module"), self._t("arch_col_role")]
        arch_rows = [
            [self._t("arch_tier_ui"),
             "main.py, ui/search_tab.py, ui/organize_tab.py, ui/analysis_tab.py, ui/chat_assistant.py",
             self._t("arch_ui_desc")],
            [self._t("arch_tier_logic"),
             "ui/workflows.py, ui/keyword_gen.py, ui/import_dialog.py, ui/image_components.py",
             self._t("arch_logic_desc")],
            [self._t("arch_tier_ai"),
             "utils/ai_extractor.py, utils/smart_search.py, utils/web_search.py",
             self._t("arch_ai_desc")],
            [self._t("arch_tier_data"),
             "database.py",
             self._t("arch_data_desc")],
            [self._t("arch_tier_util"),
             "utils/pdf_extractor.py, utils/data_cleaning.py, utils/keyword_translations.py, utils/build_translations.py",
             self._t("arch_util_desc")],
        ]
        self._add_table(arch_headers, arch_rows, col_widths=[2.5, 5.5, 7.5])

        self._add_heading(self._t("2.2"), level=2)
        self._add_para(self._t("dep_p1"))

        self._add_heading(self._t("2.3"), level=2)
        self._add_heading(self._t("2.3.1"), level=3)
        self._add_para(self._t("tkinter_p1"))
        self._add_heading(self._t("2.3.2"), level=3)
        self._add_para(self._t("sqlite_p1"))
        self._add_heading(self._t("2.3.3"), level=3)
        self._add_para(self._t("multi_ai_p1"))

        self._add_heading(self._t("2.4"), level=2)
        self._add_code(self._t("file_tree"))
        self.doc.add_page_break()

    def build_ch3_database(self):
        self._add_heading(self._t("ch3_title"), level=1)

        self._add_heading(self._t("3.1"), level=2)
        self._add_para(self._t("db_principles"))

        self._add_heading(self._t("3.2"), level=2)
        self._add_heading(self._t("3.2.1"), level=3)
        self._add_para(self._t("cases_desc"))

        case_headers = [self._t("col_field"), self._t("col_name_zh"), self._t("col_type"), self._t("col_desc")]
        case_rows = [
            ["id", "ID", "INTEGER", self._t("col_pk")],
            ["name", self._t("col_name_zh"), "TEXT", self._t("col_name_desc")],
            ["name_en", self._t("col_name_en"), "TEXT", ""],
            ["name_official", self._t("col_name_official"), "TEXT", ""],
            ["country", self._t("col_country"), "TEXT", ""],
            ["city", self._t("col_city"), "TEXT", ""],
            ["year_built", self._t("col_year_built"), "TEXT", ""],
            ["demolition_year", self._t("col_demolition"), "TEXT", ""],
            ["restoration_year", self._t("col_restoration"), "TEXT", ""],
            ["quantity", self._t("col_quantity"), "INTEGER", "DEFAULT 1"],
            ["building_type", self._t("col_building_type"), "TEXT", ""],
            ["architect", self._t("col_architect"), "TEXT", ""],
            ["latitude", self._t("col_lat"), "REAL", ""],
            ["longitude", self._t("col_lon"), "REAL", ""],
            ["material", self._t("col_material"), "TEXT", ""],
            ["area_sqm", self._t("col_area"), "REAL", ""],
            ["protection_level", self._t("col_protection"), "TEXT", ""],
            ["description", self._t("col_description"), "TEXT", ""],
            ["image_path", self._t("col_image_path"), "TEXT", ""],
            ["image_urls", self._t("col_image_urls"), "TEXT", ""],
            ["search_time", self._t("col_search_time"), "TEXT", ""],
            ["search_keywords", self._t("col_search_kw"), "TEXT", ""],
            ["condition", self._t("col_condition"), "TEXT", ""],
            ["spread_method", self._t("col_spread"), "TEXT", ""],
            ["notes", self._t("col_notes"), "TEXT", ""],
            ["verification_status", self._t("col_verif_status"), "TEXT", "DEFAULT 'unverified'"],
            ["completeness_score", self._t("col_completeness"), "REAL", "DEFAULT 0"],
            ["source_credibility", self._t("col_credibility"), "TEXT", "DEFAULT 'unknown'"],
            ["raw_case_ids", self._t("col_raw_ids"), "TEXT", ""],
            ["sort_order", self._t("col_sort_order"), "REAL", "DEFAULT 0"],
            ["created_at", self._t("col_created"), "TIMESTAMP", "DEFAULT CURRENT_TIMESTAMP"],
            ["updated_at", self._t("col_updated"), "TIMESTAMP", "DEFAULT CURRENT_TIMESTAMP"],
        ]
        self._add_table(case_headers, case_rows, col_widths=[3, 2.5, 2.5, 7.5])

        self._add_heading(self._t("3.2.2"), level=3)
        self._add_para(self._t("raw_desc"))

        raw_headers = [self._t("col_field"), self._t("col_name_zh"), self._t("col_type"), self._t("col_desc")]
        raw_rows = [
            ["id", "ID", "INTEGER", self._t("col_pk")],
            ["name", self._t("col_name_zh"), "TEXT", "NOT NULL"],
            ["name_en", self._t("col_name_en"), "TEXT", ""],
            ["name_official", self._t("col_name_official"), "TEXT", ""],
            ["country", self._t("col_country"), "TEXT", ""],
            ["city", self._t("col_city"), "TEXT", ""],
            ["construction_year", self._t("col_year_built"), "TEXT", ""],
            ["demolition_year", self._t("col_demolition"), "TEXT", ""],
            ["restoration_year", self._t("col_restoration"), "TEXT", ""],
            ["quantity", self._t("col_quantity"), "INTEGER", "DEFAULT 1"],
            ["function", self._t("col_function"), "TEXT", ""],
            ["designer", self._t("col_architect"), "TEXT", ""],
            ["latitude", self._t("col_lat"), "REAL", ""],
            ["longitude", self._t("col_lon"), "REAL", ""],
            ["search_time", self._t("col_search_time"), "TEXT", ""],
            ["search_keywords", self._t("col_search_kw"), "TEXT", ""],
            ["image_urls", self._t("col_image_urls"), "TEXT", ""],
            ["source_urls", self._t("col_source_urls"), "TEXT", ""],
            ["authenticity_status", self._t("col_auth_status"), "TEXT", "DEFAULT 'raw'"],
            ["authenticity_score", self._t("col_auth_score"), "REAL", ""],
            ["authenticity_notes", self._t("col_auth_notes"), "TEXT", ""],
            ["building_status", self._t("col_building_status"), "TEXT", ""],
            ["image_path", self._t("col_image_path"), "TEXT", ""],
            ["spread_method", self._t("col_spread"), "TEXT", ""],
            ["sort_order", self._t("col_sort_order"), "REAL", "DEFAULT 0"],
            ["created_at", self._t("col_created"), "TIMESTAMP", "DEFAULT CURRENT_TIMESTAMP"],
            ["updated_at", self._t("col_updated"), "TIMESTAMP", "DEFAULT CURRENT_TIMESTAMP"],
        ]
        self._add_table(raw_headers, raw_rows, col_widths=[3, 2.5, 2.5, 7.5])

        self._add_heading(self._t("3.2.3"), level=3)
        aux_headers = [self._t("aux_tbl_name"), self._t("aux_tbl_purpose"), self._t("aux_tbl_keys")]
        aux_rows = [
            ["building_types", self._t("aux_bt"), "id, type_name (UNIQUE)"],
            ["case_sources", self._t("aux_cs"), "id, raw_case_id, url, credibility, title, fetch_time"],
            ["field_sources", self._t("aux_fs"), "id, raw_case_id, case_id, field_name, field_value, source_url, source_type, confidence"],
            ["case_images", self._t("aux_ci"), "id, raw_case_id, filename, image_data (BLOB), file_path, source_url, author, image_time"],
            ["image_trash", self._t("aux_it"), "id, raw_case_id, filename, image_data, file_path, source_url, author, image_time, deleted_at"],
            ["verification_log", self._t("aux_vl"), "id, raw_case_id, action, details, result, created_at"],
            ["operation_log", self._t("aux_ol"), "id, operation_type, target_table, record_id, details, created_at"],
            ["literature", self._t("aux_lit"), "id, title, authors, pub_date, source_url, content, summary, file_path"],
            ["lit_match_log", self._t("aux_lml"), "id, literature_id, raw_case_id, match_score, match_details, created_at"],
            ["filtered_results", self._t("aux_fr"), "id, query, title, snippet, link, relevance_reason"],
            ["raw_search_results", self._t("aux_rs"), "id, query, engine, title, snippet, link"],
            ["scholar_search_results", self._t("aux_ssr"), "id, query, engine, title, snippet, link, authors, pub_year"],
            ["verified_cases", self._t("aux_vc"), "id, case_id, verified_at, verified_by, verification_method"],
            ["dedup_removed", self._t("aux_dr"), "id, link, query, removed_at"],
        ]
        self._add_table(aux_headers, aux_rows, col_widths=[3, 6, 6.5])

        self._add_heading(self._t("3.3"), level=2)
        self._add_para(self._t("migration_p1"))

        self._add_heading(self._t("3.4"), level=2)
        self._add_para(self._t("undo_p1"))
        self.doc.add_page_break()

    def build_ch4_modules(self):
        self._add_heading(self._t("ch4_title"), level=1)

        self._add_heading(self._t("4.1"), level=2)
        self._add_para(self._t("search_p1"))
        self._add_heading(self._t("4.1.1"), level=3)
        self._add_para(self._t("engines_p1"))
        self._add_heading(self._t("4.1.2"), level=3)
        self._add_para(self._t("chrome_ext_p1"))
        self._add_heading(self._t("4.1.3"), level=3)
        self._add_para(self._t("ai_extract_p1"))
        self._add_heading(self._t("4.1.4"), level=3)
        self._add_para(self._t("lit_mgmt_p1"))

        self._add_heading(self._t("4.2"), level=2)
        self._add_para(self._t("organize_p1"))
        self._add_heading(self._t("4.2.1"), level=3)
        self._add_para(self._t("raw_panel_p1"))
        self._add_heading(self._t("4.2.2"), level=3)
        self._add_para(self._t("verify_panel_p1"))
        self._add_para(self._t("verify_panel_p2"))
        self._add_para(self._t("verify_panel_p3"))
        self._add_heading(self._t("4.2.3"), level=3)
        self._add_para(self._t("db_panel_p1"))

        self._add_heading(self._t("4.3"), level=2)
        self._add_para(self._t("analysis_p1"))

        self._add_heading(self._t("4.4"), level=2)
        self._add_para(self._t("chat_p1"))
        self._add_heading(self._t("4.4.1"), level=3)
        self._add_para(self._t("quick_cmd_p1"))
        self._add_heading(self._t("4.4.2"), level=3)
        self._add_para(self._t("nl_understanding_p1"))
        self._add_heading(self._t("4.4.3"), level=3)
        self._add_para(self._t("voice_p1"))

        self._add_heading(self._t("4.5"), level=2)
        self._add_para(self._t("wf_p1"))

        wf_headers = [self._t("wf_col_name"), self._t("wf_col_type"), self._t("wf_col_steps"), self._t("wf_col_desc")]
        wf_rows = [
            [self._t("wf1_name"), self._t("wf1_type"), self._t("wf1_steps"), self._t("wf1_desc")],
            [self._t("wf2_name"), self._t("wf2_type"), self._t("wf2_steps"), self._t("wf2_desc")],
            [self._t("wf3_name"), self._t("wf2_type"), self._t("wf2_steps"), self._t("wf3_desc")],
            [self._t("wf4_name"), self._t("wf2_type"), self._t("wf2_steps"), self._t("wf4_desc")],
        ]
        self._add_table(wf_headers, wf_rows, col_widths=[3, 1.5, 1, 10])

        self._add_heading(self._t("4.5.1"), level=3)
        self._add_para(self._t("fsm_p1"))
        self.doc.add_page_break()

    def build_ch5_tech(self):
        self._add_heading(self._t("ch5_title"), level=1)

        self._add_heading(self._t("5.1"), level=2)
        self._add_para(self._t("ai_unified_p1"))
        self._add_para(self._t("ai_unified_p2"))

        self._add_heading(self._t("5.2"), level=2)
        self._add_para(self._t("search_integration_p1"))
        self._add_para(self._t("search_integration_p2"))

        self._add_heading(self._t("5.3"), level=2)
        self._add_para(self._t("import_p1"))
        self._add_para(self._t("import_p2"))

        self._add_heading(self._t("5.4"), level=2)
        self._add_para(self._t("image_mgmt_p1"))

        self._add_heading(self._t("5.5"), level=2)
        self._add_para(self._t("pdf_p1"))

        self._add_heading(self._t("5.6"), level=2)
        self._add_para(self._t("thread_p1"))
        self.doc.add_page_break()

    def build_ch6_deployment(self):
        self._add_heading(self._t("ch6_title"), level=1)

        self._add_heading(self._t("6.1"), level=2)
        env_headers = [self._t("env_col_component"), self._t("env_col_requirement")]
        env_rows = [
            [self._t("env_os"), self._t("env_os_val")],
            [self._t("env_python"), self._t("env_python_val")],
            [self._t("env_ram"), self._t("env_ram_val")],
            [self._t("env_disk"), self._t("env_disk_val")],
            [self._t("env_net"), self._t("env_net_val")],
        ]
        self._add_table(env_headers, env_rows, col_widths=[4, 11.5])

        self._add_heading(self._t("6.2"), level=2)
        self._add_para(self._t("install_p1"))
        self._add_code("pip install pandas openpyxl matplotlib pymupdf python-docx openai")
        self._add_para(self._t("install_p2"))
        self._add_code("pip install sounddevice vosk")
        self._add_para(self._t("install_p3"))
        self._add_para(self._t("install_p4"))
        self._add_code(self._t("install_keys"))
        self._add_para(self._t("install_p5"))
        self._add_code(self._t("install_p5_cmd"))

        self._add_heading(self._t("6.3"), level=2)
        self._add_para(self._t("chrome_ext_setup"))

        self._add_heading(self._t("6.4"), level=2)
        self._add_para(self._t("db_maintenance"))
        self.doc.add_page_break()

    def build_ch7_summary(self):
        self._add_heading(self._t("ch7_title"), level=1)

        self._add_heading(self._t("7.1"), level=2)
        for k in ["ach_1", "ach_2", "ach_3", "ach_4", "ach_5"]:
            self._add_bullet(self._t(k))

        self._add_heading(self._t("7.2"), level=2)
        for title_k, sol_k in [("chal_1_title", "chal_1_sol"),
                                ("chal_2_title", "chal_2_sol"),
                                ("chal_3_title", "chal_3_sol"),
                                ("chal_4_title", "chal_4_sol"),
                                ("chal_5_title", "chal_5_sol")]:
            self._add_bullet(f"{self._t(title_k)}：{self._t(sol_k)}")

        self._add_heading(self._t("7.3"), level=2)
        for k in ["future_1", "future_2", "future_3", "future_4", "future_5", "future_6"]:
            self._add_bullet(self._t(k))
        self.doc.add_page_break()

    def build_ch8_innovation(self):
        self._add_heading(self._t("ch8_title"), level=1)

        # ── 8.1 核心创新 ──
        self._add_heading(self._t("8.1"), level=2)

        self._add_heading(self._t("8.1.1"), level=3)
        self._add_para(self._t("inno_1_p1"))
        self._add_para(self._t("inno_1_p2"))

        self._add_heading(self._t("8.1.2"), level=3)
        self._add_para(self._t("inno_2_p1"))

        self._add_heading(self._t("8.1.3"), level=3)
        self._add_para(self._t("inno_3_p1"))
        self._add_para(self._t("inno_3_p2"))

        self._add_heading(self._t("8.1.4"), level=3)
        self._add_para(self._t("inno_4_p1"))
        self._add_para(self._t("inno_4_p2"))

        self._add_heading(self._t("8.1.5"), level=3)
        self._add_para(self._t("inno_6_p1"))

        # ── 8.2 效率量化 ──
        self._add_heading(self._t("8.2"), level=2)
        eff_headers = [
            self._t("eff_col_stage"), self._t("eff_col_task"),
            self._t("eff_col_manual"), self._t("eff_col_auto"),
            self._t("eff_col_saving"),
        ]
        eff_rows = [
            [self._t("eff_search"), self._t("eff_search_task"),
             self._t("eff_search_manual"), self._t("eff_search_auto"),
             self._t("eff_search_saving")],
            [self._t("eff_screening"), self._t("eff_screening_task"),
             self._t("eff_screening_manual"), self._t("eff_screening_auto"),
             self._t("eff_screening_saving")],
            [self._t("eff_extraction"), self._t("eff_extraction_task"),
             self._t("eff_extraction_manual"), self._t("eff_extraction_auto"),
             self._t("eff_extraction_saving")],
            [self._t("eff_verification"), self._t("eff_verification_task"),
             self._t("eff_verification_manual"), self._t("eff_verification_auto"),
             self._t("eff_verification_saving")],
            [self._t("eff_total"), self._t("eff_total_task"),
             self._t("eff_total_manual"), self._t("eff_total_auto"),
             self._t("eff_total_saving")],
        ]
        self._add_para(self._t("eff_table_title"))
        self._add_table(eff_headers, eff_rows, col_widths=[2, 2.5, 3.5, 3.5, 1.5])
        self._add_para(self._t("eff_note"))

        # ── 8.3 数据库架构总览 ──
        self._add_heading(self._t("8.3"), level=2)
        arch_headers = [
            self._t("arch_col_table"), self._t("arch_col_fields"),
            self._t("arch_col_rows"), self._t("arch_col_purpose"),
            self._t("arch_col_relations"),
        ]
        arch_rows = [
            ["cases", "31", "686",
             self._t("aux_cases"),
             self._t("arch_rel_cases")],
            ["raw_cases", "26+",
             self._t("arch_raw_cases_rows"),
             self._t("aux_raw_cases"),
             self._t("arch_rel_raw")],
            ["case_sources", "5",
             self._t("arch_cs_rows"),
             self._t("arch_cs_purpose"),
             "FK → raw_cases.case_id, FK → cases.id"],
            ["field_sources", "6",
             self._t("arch_fs_rows"),
             self._t("arch_fs_purpose"),
             "FK → cases.id"],
            ["verification_log", "7",
             self._t("arch_vl_rows"),
             self._t("arch_vl_purpose"),
             "FK → raw_cases.case_id"],
            ["operation_log", "6",
             self._t("arch_ol_rows"),
             self._t("arch_ol_purpose"),
             "—"],
            ["literature", "14",
             self._t("arch_lit_rows"),
             self._t("arch_lit_purpose"),
             "—"],
            ["lit_match_log", "6",
             self._t("arch_lml_rows"),
             self._t("arch_lml_purpose"),
             "FK → literature.id, FK → cases.id"],
            ["filtered_results", "7",
             self._t("arch_fr_rows"),
             self._t("aux_fr"),
             "FK → raw_search_results.id"],
            ["raw_search_results", "7",
             self._t("arch_rs_rows"),
             self._t("aux_rs"),
             "—"],
            ["scholar_search_results", "9",
             self._t("arch_ssr_rows"),
             self._t("aux_ssr"),
             "—"],
            ["building_types", "4",
             self._t("arch_bt_rows"),
             self._t("arch_bt_purpose"),
             "—"],
            ["case_images / image_trash", "6",
             self._t("arch_img_rows"),
             self._t("arch_img_purpose"),
             "FK → cases.id"],
            ["verified_cases", "6",
             self._t("arch_vc_rows"),
             self._t("aux_vc"),
             "FK → cases.case_id"],
            ["dedup_removed", "5",
             self._t("arch_dr_rows"),
             self._t("aux_dr"),
             "FK → raw_cases / cases"],
        ]
        self._add_para(self._t("arch_schema_title"))
        self._add_table(arch_headers, arch_rows, col_widths=[2.5, 1.2, 1.2, 5.0, 5.5])

        # ── 8.4 数据处理全流程 ──
        self._add_heading(self._t("8.4"), level=2)
        self._add_para(self._t("pipeline_note"))

        self.doc.add_page_break()

    def build_appendix(self):
        self._add_heading(self._t("appendix_a"), level=1)
        self._add_heading(self._t("a1"), level=2)
        self._add_para(self._t("app_a1_desc"))
        self._add_heading(self._t("a2"), level=2)
        self._add_para(self._t("app_a2_desc"))
        self._add_heading(self._t("a3"), level=2)
        self._add_para(self._t("app_a3_desc"))
        self.doc.add_page_break()

        self._add_heading(self._t("appendix_b"), level=1)
        api_headers = [self._t("api_col_name"), self._t("api_col_endpoint"), self._t("api_col_model"), self._t("api_col_purpose")]
        api_rows = [
            ["DeepSeek", "https://api.deepseek.com/v1", "deepseek-chat", self._t("api_deepseek_purpose")],
            ["Gemini (via Yunwu)", "https://yunwu.ai/v1", "gemini-2.5-flash", self._t("api_gemini_purpose")],
            ["Serper.dev", "https://google.serper.dev/search", "-", self._t("api_serper_purpose")],
            ["Baidu Qianfan", "https://appbuilder.baidu.com/v2/...", "-", self._t("api_baidu_purpose")],
        ]
        self._add_table(api_headers, api_rows, col_widths=[3, 5, 3, 4.5])
        self._add_para(self._t("api_keys_note"))
        self.doc.add_page_break()

        self._add_heading(self._t("appendix_c"), level=1)
        cmd_headers = [self._t("cmd_col_shortcut"), self._t("cmd_col_function")]
        cmd_rows = [
            ["Ctrl+B / Cmd+B", self._t("cmd_ctrl_b")],
            [self._t("cmd_search_keyword"), self._t("cmd_search")],
            [self._t("cmd_verify_cmd"), self._t("cmd_verify")],
            [self._t("cmd_ds_cmd"), self._t("cmd_ds_verify")],
            [self._t("cmd_g_cmd"), self._t("cmd_g_verify")],
            [self._t("cmd_oneclick_cmd"), self._t("cmd_oneclick_search")],
            [self._t("cmd_switch_cmd"), self._t("cmd_switch_tab")],
            [self._t("cmd_refresh_cmd"), self._t("cmd_refresh")],
            [self._t("cmd_import_cmd"), self._t("cmd_import_export")],
            [self._t("cmd_autocomplete_cmd"), self._t("cmd_autocomplete")],
            [self._t("cmd_delete_cmd"), self._t("cmd_delete")],
            [self._t("cmd_add_cmd"), self._t("cmd_add_case")],
            [self._t("cmd_detail_cmd"), self._t("cmd_view_detail")],
            [self._t("cmd_nav_cmd"), self._t("cmd_nav")],
            [self._t("cmd_submit_cmd"), self._t("cmd_submit_verify")],
            [self._t("cmd_approve_cmd"), self._t("cmd_approve")],
            [self._t("cmd_reject_cmd"), self._t("cmd_reject")],
            [self._t("cmd_select_all_cmd"), self._t("cmd_select_all")],
            [self._t("cmd_cancel_cmd"), self._t("cmd_cancel")],
            [self._t("cmd_help_cmd"), self._t("cmd_help")],
        ]
        self._add_table(cmd_headers, cmd_rows, col_widths=[5, 10.5])
        self.doc.add_page_break()

        self._add_heading(self._t("appendix_d"), level=1)
        self._add_para(self._t("migration_v1"))
        self._add_para(self._t("migration_v2"))
        self._add_para(self._t("migration_v3_v12"))
        self._add_para(self._t("migration_note"))

    def build_all(self):
        self.build_cover()
        self.build_toc()
        self.build_ch1_overview()
        self.build_ch8_innovation()
        self.build_ch2_architecture()
        self.build_ch3_database()
        self.build_ch4_modules()
        self.build_ch5_tech()
        self.build_ch6_deployment()
        self.build_ch7_summary()
        self.build_appendix()

    def save(self):
        self.doc.save(OUTPUT_PATH)
        return OUTPUT_PATH


if __name__ == "__main__":
    gen = ReportGenerator()
    gen.build_all()
    path = gen.save()
    print(gen._t("done_msg", path=path))
