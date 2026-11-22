import json
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openai

import config

YUNWU_BASE_URL, DEFAULT_MODEL, _KW_API_KEY = config.get_llm("keyword_translate")

try:
    from utils.keyword_translations import KEYWORD_TRANSLATIONS
except ImportError:
    KEYWORD_TRANSLATIONS = {}


COUNTRY_GROUPS = [
    ("亚洲", [
        ("中国", "China", "Chinese"),
        ("日本", "Japan", "Japanese"),
        ("韩国", "South Korea", "Korean"),
        ("朝鲜", "North Korea", "Korean"),
        ("蒙古", "Mongolia", "Mongolian"),
        ("越南", "Vietnam", "Vietnamese"),
        ("老挝", "Laos", "Lao"),
        ("柬埔寨", "Cambodia", "Khmer"),
        ("泰国", "Thailand", "Thai"),
        ("缅甸", "Myanmar", "Burmese"),
        ("马来西亚", "Malaysia", "Malay"),
        ("新加坡", "Singapore", "English, Malay, Chinese, Tamil"),
        ("印度尼西亚", "Indonesia", "Indonesian"),
        ("菲律宾", "Philippines", "Filipino, English"),
        ("东帝汶", "Timor-Leste", "Tetum, Portuguese"),
        ("文莱", "Brunei", "Malay"),
        ("印度", "India", "Hindi, English"),
        ("巴基斯坦", "Pakistan", "Urdu, English"),
        ("孟加拉国", "Bangladesh", "Bengali"),
        ("斯里兰卡", "Sri Lanka", "Sinhala, Tamil"),
        ("尼泊尔", "Nepal", "Nepali"),
        ("不丹", "Bhutan", "Dzongkha"),
        ("马尔代夫", "Maldives", "Dhivehi"),
        ("阿富汗", "Afghanistan", "Dari, Pashto"),
        ("哈萨克斯坦", "Kazakhstan", "Kazakh, Russian"),
        ("吉尔吉斯斯坦", "Kyrgyzstan", "Kyrgyz, Russian"),
        ("塔吉克斯坦", "Tajikistan", "Tajik"),
        ("乌兹别克斯坦", "Uzbekistan", "Uzbek"),
        ("土库曼斯坦", "Turkmenistan", "Turkmen"),
        ("伊朗", "Iran", "Persian"),
        ("伊拉克", "Iraq", "Arabic, Kurdish"),
        ("叙利亚", "Syria", "Arabic"),
        ("约旦", "Jordan", "Arabic"),
        ("黎巴嫩", "Lebanon", "Arabic"),
        ("以色列", "Israel", "Hebrew, Arabic"),
        ("巴勒斯坦", "Palestine", "Arabic"),
        ("沙特阿拉伯", "Saudi Arabia", "Arabic"),
        ("也门", "Yemen", "Arabic"),
        ("阿曼", "Oman", "Arabic"),
        ("阿联酋", "UAE", "Arabic"),
        ("卡塔尔", "Qatar", "Arabic"),
        ("巴林", "Bahrain", "Arabic"),
        ("科威特", "Kuwait", "Arabic"),
        ("土耳其", "Turkey", "Turkish"),
        ("塞浦路斯", "Cyprus", "Greek, Turkish"),
        ("格鲁吉亚", "Georgia", "Georgian"),
        ("亚美尼亚", "Armenia", "Armenian"),
        ("阿塞拜疆", "Azerbaijan", "Azerbaijani"),
    ]),
    ("欧洲", [
        ("英国", "United Kingdom", "English"),
        ("爱尔兰", "Ireland", "Irish, English"),
        ("法国", "France", "French"),
        ("德国", "Germany", "German"),
        ("意大利", "Italy", "Italian"),
        ("西班牙", "Spain", "Spanish"),
        ("葡萄牙", "Portugal", "Portuguese"),
        ("荷兰", "Netherlands", "Dutch"),
        ("比利时", "Belgium", "Dutch, French, German"),
        ("卢森堡", "Luxembourg", "Luxembourgish, French, German"),
        ("瑞士", "Switzerland", "German, French, Italian"),
        ("奥地利", "Austria", "German"),
        ("波兰", "Poland", "Polish"),
        ("捷克", "Czech Republic", "Czech"),
        ("斯洛伐克", "Slovakia", "Slovak"),
        ("匈牙利", "Hungary", "Hungarian"),
        ("罗马尼亚", "Romania", "Romanian"),
        ("保加利亚", "Bulgaria", "Bulgarian"),
        ("塞尔维亚", "Serbia", "Serbian"),
        ("克罗地亚", "Croatia", "Croatian"),
        ("斯洛文尼亚", "Slovenia", "Slovenian"),
        ("波黑", "Bosnia and Herzegovina", "Bosnian, Croatian, Serbian"),
        ("黑山", "Montenegro", "Montenegrin"),
        ("北马其顿", "North Macedonia", "Macedonian"),
        ("阿尔巴尼亚", "Albania", "Albanian"),
        ("希腊", "Greece", "Greek"),
        ("丹麦", "Denmark", "Danish"),
        ("瑞典", "Sweden", "Swedish"),
        ("挪威", "Norway", "Norwegian"),
        ("芬兰", "Finland", "Finnish, Swedish"),
        ("冰岛", "Iceland", "Icelandic"),
        ("俄罗斯", "Russia", "Russian"),
        ("白俄罗斯", "Belarus", "Belarusian, Russian"),
        ("乌克兰", "Ukraine", "Ukrainian"),
        ("摩尔多瓦", "Moldova", "Romanian"),
        ("立陶宛", "Lithuania", "Lithuanian"),
        ("拉脱维亚", "Latvia", "Latvian"),
        ("爱沙尼亚", "Estonia", "Estonian"),
        ("安道尔", "Andorra", "Catalan"),
        ("摩纳哥", "Monaco", "French"),
        ("列支敦士登", "Liechtenstein", "German"),
        ("马耳他", "Malta", "Maltese, English"),
        ("圣马力诺", "San Marino", "Italian"),
        ("梵蒂冈", "Vatican City", "Italian, Latin"),
    ]),
    ("美洲", [
        ("美国", "United States", "English"),
        ("加拿大", "Canada", "English, French"),
        ("墨西哥", "Mexico", "Spanish"),
        ("危地马拉", "Guatemala", "Spanish"),
        ("伯利兹", "Belize", "English"),
        ("萨尔瓦多", "El Salvador", "Spanish"),
        ("洪都拉斯", "Honduras", "Spanish"),
        ("尼加拉瓜", "Nicaragua", "Spanish"),
        ("哥斯达黎加", "Costa Rica", "Spanish"),
        ("巴拿马", "Panama", "Spanish"),
        ("古巴", "Cuba", "Spanish"),
        ("牙买加", "Jamaica", "English"),
        ("海地", "Haiti", "French, Haitian Creole"),
        ("多米尼加", "Dominican Republic", "Spanish"),
        ("巴哈马", "Bahamas", "English"),
        ("特立尼达和多巴哥", "Trinidad and Tobago", "English"),
        ("巴巴多斯", "Barbados", "English"),
        ("圣基茨和尼维斯", "Saint Kitts and Nevis", "English"),
        ("圣卢西亚", "Saint Lucia", "English"),
        ("圣文森特和格林纳丁斯", "Saint Vincent and the Grenadines", "English"),
        ("格林纳达", "Grenada", "English"),
        ("多米尼克", "Dominica", "English"),
        ("安提瓜和巴布达", "Antigua and Barbuda", "English"),
        ("哥伦比亚", "Colombia", "Spanish"),
        ("委内瑞拉", "Venezuela", "Spanish"),
        ("圭亚那", "Guyana", "English"),
        ("苏里南", "Suriname", "Dutch"),
        ("厄瓜多尔", "Ecuador", "Spanish"),
        ("秘鲁", "Peru", "Spanish"),
        ("玻利维亚", "Bolivia", "Spanish, Quechua, Aymara"),
        ("巴西", "Brazil", "Portuguese"),
        ("巴拉圭", "Paraguay", "Spanish, Guarani"),
        ("乌拉圭", "Uruguay", "Spanish"),
        ("阿根廷", "Argentina", "Spanish"),
        ("智利", "Chile", "Spanish"),
    ]),
    ("非洲", [
        ("埃及", "Egypt", "Arabic"),
        ("利比亚", "Libya", "Arabic"),
        ("突尼斯", "Tunisia", "Arabic"),
        ("阿尔及利亚", "Algeria", "Arabic, French"),
        ("摩洛哥", "Morocco", "Arabic, French"),
        ("毛里塔尼亚", "Mauritania", "Arabic"),
        ("苏丹", "Sudan", "Arabic, English"),
        ("南苏丹", "South Sudan", "English"),
        ("埃塞俄比亚", "Ethiopia", "Amharic"),
        ("厄立特里亚", "Eritrea", "Tigrinya, Arabic, English"),
        ("吉布提", "Djibouti", "French, Arabic"),
        ("索马里", "Somalia", "Somali, Arabic"),
        ("肯尼亚", "Kenya", "Swahili, English"),
        ("乌干达", "Uganda", "English, Swahili"),
        ("坦桑尼亚", "Tanzania", "Swahili, English"),
        ("卢旺达", "Rwanda", "Kinyarwanda, French, English"),
        ("布隆迪", "Burundi", "Kirundi, French"),
        ("刚果民主共和国", "DR Congo", "French"),
        ("刚果共和国", "Congo Republic", "French"),
        ("加蓬", "Gabon", "French"),
        ("赤道几内亚", "Equatorial Guinea", "Spanish, French"),
        ("喀麦隆", "Cameroon", "French, English"),
        ("中非共和国", "Central African Republic", "French"),
        ("乍得", "Chad", "French, Arabic"),
        ("尼日尔", "Niger", "French"),
        ("尼日利亚", "Nigeria", "English"),
        ("贝宁", "Benin", "French"),
        ("多哥", "Togo", "French"),
        ("加纳", "Ghana", "English"),
        ("科特迪瓦", "Côte d'Ivoire", "French"),
        ("利比里亚", "Liberia", "English"),
        ("塞拉利昂", "Sierra Leone", "English"),
        ("几内亚", "Guinea", "French"),
        ("几内亚比绍", "Guinea-Bissau", "Portuguese"),
        ("冈比亚", "Gambia", "English"),
        ("塞内加尔", "Senegal", "French"),
        ("马里", "Mali", "French"),
        ("布基纳法索", "Burkina Faso", "French"),
        ("佛得角", "Cape Verde", "Portuguese"),
        ("圣多美和普林西比", "São Tomé and Príncipe", "Portuguese"),
        ("安哥拉", "Angola", "Portuguese"),
        ("纳米比亚", "Namibia", "English"),
        ("博茨瓦纳", "Botswana", "English"),
        ("南非", "South Africa", "Afrikaans, English, Zulu, etc."),
        ("津巴布韦", "Zimbabwe", "English, Shona, Ndebele"),
        ("赞比亚", "Zambia", "English"),
        ("马拉维", "Malawi", "English, Chichewa"),
        ("莫桑比克", "Mozambique", "Portuguese"),
        ("马达加斯加", "Madagascar", "Malagasy, French"),
        ("科摩罗", "Comoros", "Comorian, Arabic, French"),
        ("毛里求斯", "Mauritius", "English, French"),
        ("塞舌尔", "Seychelles", "English, French, Seychellois Creole"),
        ("斯威士兰", "Eswatini", "Swazi, English"),
        ("莱索托", "Lesotho", "Sesotho, English"),
    ]),
    ("大洋洲", [
        ("澳大利亚", "Australia", "English"),
        ("新西兰", "New Zealand", "English"),
        ("巴布亚新几内亚", "Papua New Guinea", "English, Tok Pisin, Hiri Motu"),
        ("斐济", "Fiji", "English, Fijian, Hindi"),
        ("所罗门群岛", "Solomon Islands", "English"),
        ("瓦努阿图", "Vanuatu", "Bislama, English, French"),
        ("萨摩亚", "Samoa", "Samoan, English"),
        ("汤加", "Tonga", "Tongan, English"),
        ("基里巴斯", "Kiribati", "English, Gilbertese"),
        ("密克罗尼西亚", "Micronesia", "English"),
        ("马绍尔群岛", "Marshall Islands", "Marshallese, English"),
        ("帕劳", "Palau", "Palauan, English"),
        ("瑙鲁", "Nauru", "Nauruan, English"),
        ("图瓦卢", "Tuvalu", "Tuvaluan, English"),
    ]),
]

LANGUAGE_OPTIONS = ["中文", "English", "official"]

CORE_KEYWORDS = {
    "zh": ["中式建筑", "中国人建筑", "中国风格建筑", "中国式建筑", "模仿中国建筑", "学习中国建筑",
           "中国传统建筑", "中华建筑", "中式园林", "唐人街建筑", "中国寺庙", "中式民居", "东方建筑",
           "建筑"],
    "en": ["Chinese architecture", "Chinese people architecture", "Chinese-style architecture",
           "Chinese-style building", "imitating Chinese architecture", "learning from Chinese architecture",
           "traditional Chinese architecture", "Sino-architecture", "Chinese garden",
           "Chinatown architecture", "Chinese temple", "Chinese vernacular dwelling",
           "Oriental architecture", "architecture"],
}


def _get_country_langs(country_record):
    return [l.strip() for l in country_record[2].split(",")]


def _all_countries():
    result = []
    for _region, countries in COUNTRY_GROUPS:
        result.extend(countries)
    return result


class KeywordGenDialog:
    def __init__(self, parent, callback=None):
        self.callback = callback
        self.country_vars = {}
        self.country_widgets = {}
        self._translation_cache = {}
        self._load_prebuilt_translations()
        self.win = tk.Toplevel(parent)
        self.win.title("关键词生成器")
        self.win.geometry("850x880")
        self.win.transient(parent)

        # Scrollable area
        canvas = tk.Canvas(self.win, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(8, 0))
        scrollbar.pack(side="right", fill="y", padx=(0, 8), pady=(8, 0))

        self._build_country_section()
        self._build_language_section()
        self._build_keyword_section()
        self._build_academic_section()
        self._build_generic_section()
        self._build_result_area()
        self._build_bottom()

        self.win.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * int(e.delta / 120), "units"))

    def _build_country_section(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 1. 选择国家 ", padding=8)
        frame.pack(fill="x", padx=8, pady=(8, 0))

        # Search bar + global action buttons
        search_bar = ttk.Frame(frame)
        search_bar.pack(fill="x")
        ttk.Label(search_bar, text="搜索：").pack(side="left")
        self.country_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_bar, textvariable=self.country_search_var, width=25)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda e: self._filter_countries())
        ttk.Button(search_bar, text="全选", command=self._select_all_countries).pack(side="left", padx=2)
        ttk.Button(search_bar, text="清空", command=self._clear_all_countries).pack(side="left", padx=2)
        self._exclude_country_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(search_bar, text="不放入词组中", variable=self._exclude_country_var).pack(side="left", padx=5)
        self.country_count_label = ttk.Label(search_bar, text="")
        self.country_count_label.pack(side="right", padx=5)

        # Regional groups — each in its own LabelFrame with per-region buttons
        self._region_frames = {}

        for region, countries in COUNTRY_GROUPS:
            region_frame = ttk.LabelFrame(frame, text=f" {region} ", padding=5)
            region_frame.pack(fill="x", pady=(4, 0))
            self._region_frames[region] = region_frame

            btn_bar = ttk.Frame(region_frame)
            btn_bar.pack(fill="x", pady=(0, 3))
            ttk.Button(btn_bar, text="全选",
                       command=lambda r=region: self._select_region(r, True)).pack(side="left", padx=2)
            ttk.Button(btn_bar, text="清空",
                       command=lambda r=region: self._select_region(r, False)).pack(side="left", padx=2)

            grid_frame = ttk.Frame(region_frame)
            grid_frame.pack(fill="x")
            n_cols = 4
            for i, c in enumerate(countries):
                var = tk.BooleanVar(value=False)
                self.country_vars[c[0]] = var
                row = i // n_cols
                col = i % n_cols
                cb = tk.Checkbutton(grid_frame, text=c[0], variable=var,
                                   command=self._update_country_count)
                cb.grid(row=row, column=col, sticky="w", padx=3, pady=1)
                self.country_widgets[c[0]] = cb

        self._update_country_count()

    def _select_region(self, region, select):
        for name, countries in COUNTRY_GROUPS:
            if name == region:
                for c in countries:
                    self.country_vars[c[0]].set(select)
                break
        self._update_country_count()

    def _filter_countries(self):
        query = self.country_search_var.get().lower()
        for region, countries in COUNTRY_GROUPS:
            visible = 0
            for c in countries:
                cb = self.country_widgets[c[0]]
                if query and query not in c[0].lower():
                    cb.grid_remove()
                else:
                    cb.grid()
                    visible += 1
            if query and visible == 0:
                self._region_frames[region].pack_forget()
            else:
                # Re-pack if hidden
                try:
                    self._region_frames[region].pack_info()
                except Exception:
                    self._region_frames[region].pack(fill="x", pady=(4, 0))

    def _select_all_countries(self):
        for var in self.country_vars.values():
            var.set(True)
        self._update_country_count()

    def _clear_all_countries(self):
        for var in self.country_vars.values():
            var.set(False)
        self._update_country_count()

    def _update_country_count(self):
        selected = sum(1 for v in self.country_vars.values() if v.get())
        total = len(self.country_vars)
        self.country_count_label.config(text=f"已选 {selected} / {total} 个国家")

    def _build_language_section(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 2. 选择检索语言 ", padding=8)
        frame.pack(fill="x", padx=8, pady=(5, 0))

        # API Key for translation
        key_frame = ttk.Frame(frame)
        key_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(key_frame, text="API 密钥：").pack(side="left")
        self.api_key_var = tk.StringVar(value=_KW_API_KEY)
        api_entry = ttk.Entry(key_frame, textvariable=self.api_key_var, width=50, show="*")
        api_entry.pack(side="left", padx=5)
        ttk.Button(key_frame, text="显示", command=lambda: api_entry.config(show="" if api_entry.cget("show") else "*")).pack(side="left")

        self.lang_vars = {}
        for lang in LANGUAGE_OPTIONS:
            var = tk.BooleanVar(value=(lang == "中文"))
            self.lang_vars[lang] = var
            label = {"official": "官方语言（根据所选国家自动确定，使用AI翻译）", "English": "英文", "中文": "中文"}
            tk.Checkbutton(frame, text=label[lang], variable=var).pack(anchor="w")

    def _build_keyword_section(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 3. 核心关键词 ", padding=8)
        frame.pack(fill="x", padx=8, pady=(5, 0))
        self._kw_checkbutton_parent = frame

        kw_btn_bar = ttk.Frame(frame)
        kw_btn_bar.pack(fill="x", pady=(0, 3))
        ttk.Button(kw_btn_bar, text="全选", command=self._select_all_kw).pack(side="left", padx=2)
        ttk.Button(kw_btn_bar, text="清除", command=self._clear_all_kw).pack(side="left", padx=2)

        self.kw_vars = {}
        for kw in CORE_KEYWORDS["zh"]:
            var = tk.BooleanVar(value=True)
            self.kw_vars[kw] = var
            tk.Checkbutton(frame, text=kw, variable=var).pack(anchor="w")

        custom_frame = ttk.Frame(frame)
        custom_frame.pack(fill="x", pady=(5, 0))
        ttk.Label(custom_frame, text="自定义：").pack(side="left")
        self.custom_kw_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=self.custom_kw_var, width=30).pack(side="left", padx=5)
        ttk.Button(custom_frame, text="添加", command=self._add_custom_keyword).pack(side="left")

    def _build_academic_section(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 4. 学术词语 ", padding=8)
        frame.pack(fill="x", padx=8, pady=(5, 0))

        ac_btn_bar = ttk.Frame(frame)
        ac_btn_bar.pack(fill="x", pady=(0, 3))
        ttk.Button(ac_btn_bar, text="全选", command=self._select_all_academic).pack(side="left", padx=2)
        ttk.Button(ac_btn_bar, text="清除", command=self._clear_all_academic).pack(side="left", padx=2)

        self._academic_terms = [
            "Chinoiserie", "Sharawadgi", "Chinesische", "Tang style",
            "Great Buddha style", "Zen style", "Islamic Chinoiserie(khitayi)",
            "Peranakan House", "Shophouse", "Han",
        ]
        self.academic_vars = {}
        for term in self._academic_terms:
            var = tk.BooleanVar(value=False)
            self.academic_vars[term] = var
            tk.Checkbutton(frame, text=term, variable=var).pack(anchor="w")

    def _select_all_academic(self):
        for var in self.academic_vars.values():
            var.set(True)

    def _clear_all_academic(self):
        for var in self.academic_vars.values():
            var.set(False)

    def _build_generic_section(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 5. 通用关键词 ", padding=8)
        frame.pack(fill="x", padx=8, pady=(5, 0))

        g_btn_bar = ttk.Frame(frame)
        g_btn_bar.pack(fill="x", pady=(0, 3))
        ttk.Button(g_btn_bar, text="全选", command=self._select_all_generic).pack(side="left", padx=2)
        ttk.Button(g_btn_bar, text="清除", command=self._clear_all_generic).pack(side="left", padx=2)

        self._generic_terms = ["建筑"]
        self.generic_vars = {}
        for term in self._generic_terms:
            var = tk.BooleanVar(value=False)
            self.generic_vars[term] = var
            tk.Checkbutton(frame, text=term, variable=var).pack(anchor="w")

    def _select_all_generic(self):
        for var in self.generic_vars.values():
            var.set(True)

    def _clear_all_generic(self):
        for var in self.generic_vars.values():
            var.set(False)

    def _build_result_area(self):
        frame = ttk.LabelFrame(self.scroll_frame, text=" 生成结果（双击单元格可编辑 | Ctrl+F 查找替换） ", padding=5)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        # Find/Replace bar (hidden by default)
        self._fr_bar = ttk.Frame(frame)
        ttk.Label(self._fr_bar, text="查找：").pack(side="left")
        self._fr_find_var = tk.StringVar()
        self._fr_find_entry = ttk.Entry(self._fr_bar, textvariable=self._fr_find_var, width=18)
        self._fr_find_entry.pack(side="left", padx=3)
        self._fr_find_entry.bind("<KeyRelease>", self._on_find_key)
        self._fr_find_entry.bind("<Escape>", lambda e: self._hide_find_bar())
        ttk.Label(self._fr_bar, text="替换为：").pack(side="left", padx=(8, 0))
        self._fr_replace_var = tk.StringVar()
        self._fr_replace_entry = ttk.Entry(self._fr_bar, textvariable=self._fr_replace_var, width=18)
        self._fr_replace_entry.pack(side="left", padx=3)
        self._fr_replace_entry.bind("<Escape>", lambda e: self._hide_find_bar())
        ttk.Button(self._fr_bar, text="替换全部", command=self._replace_all).pack(side="left", padx=5)
        ttk.Button(self._fr_bar, text="✕", width=2, command=self._hide_find_bar).pack(side="left")
        self._fr_status_var = tk.StringVar(value="")
        ttk.Label(self._fr_bar, textvariable=self._fr_status_var, foreground="gray").pack(side="left", padx=5)

        columns = ("index", "country", "language", "keyword")
        self.result_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="extended")
        self.result_tree.heading("index", text="#")
        self.result_tree.heading("country", text="国家")
        self.result_tree.heading("language", text="语言")
        self.result_tree.heading("keyword", text="检索关键词")
        self.result_tree.column("index", width=35)
        self.result_tree.column("country", width=100)
        self.result_tree.column("language", width=80)
        self.result_tree.column("keyword", width=500)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        self._fr_bar.pack(fill="x", pady=(0, 3))
        self.result_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._fr_bar.pack_forget()  # hidden

        self.result_tree.bind("<Double-1>", self._on_double_click)
        self.result_tree.bind("<Control-f>", self._show_find_bar)
        self.result_tree.bind("<Command-f>", self._show_find_bar)

    def _show_find_bar(self, event=None):
        self._fr_bar.pack(before=self.result_tree, fill="x", pady=(0, 3))
        self._fr_find_entry.focus_set()
        self._highlight_matches()

    def _hide_find_bar(self):
        self._fr_bar.pack_forget()
        self._clear_highlights()

    def _on_find_key(self, event=None):
        self._highlight_matches()

    def _highlight_matches(self):
        query = self._fr_find_var.get().strip()
        self._clear_highlights()
        if not query:
            self._fr_status_var.set("")
            return
        count = 0
        for item in self.result_tree.get_children():
            vals = self.result_tree.item(item, "values")
            for i, v in enumerate(vals[1:], 1):  # skip index column
                if query.lower() in str(v).lower():
                    col = f"#{i+1}"
                    try:
                        self.result_tree.tag_has_match  # ensure tag exists
                    except Exception:
                        pass
                    count += 1
        self.result_tree.tag_configure("match", background="#FFEB3B")
        for item in self.result_tree.get_children():
            vals = self.result_tree.item(item, "values")
            hit = any(query.lower() in str(v).lower() for v in vals[1:])
            if hit:
                self.result_tree.item(item, tags=("match",))
        self._fr_status_var.set(f"{count} 处匹配" if count else "无匹配")

    def _clear_highlights(self):
        self.result_tree.tag_configure("match", background="")
        for item in self.result_tree.get_children():
            self.result_tree.item(item, tags=())

    def _replace_all(self):
        find_text = self._fr_find_var.get()
        replace_text = self._fr_replace_var.get()
        if not find_text:
            return
        count = 0
        for item in self.result_tree.get_children():
            vals = list(self.result_tree.item(item, "values"))
            changed = False
            for i in range(1, len(vals)):  # skip index column
                if find_text in str(vals[i]):
                    vals[i] = str(vals[i]).replace(find_text, replace_text)
                    changed = True
                    count += 1
            if changed:
                self.result_tree.item(item, values=vals)
        self._fr_status_var.set(f"已替换 {count} 处")
        self._highlight_matches()

    def _on_double_click(self, event):
        """Inline-edit the clicked cell."""
        region = self.result_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.result_tree.identify_column(event.x)
        item = self.result_tree.identify_row(event.y)
        if not item or column == "#1":
            return

        col_idx = int(column[1:]) - 1
        bbox = self.result_tree.bbox(item, column)
        if not bbox:
            return

        x, y, w, h = bbox
        old_value = self.result_tree.set(item, column)

        edit = tk.Entry(self.result_tree)
        edit.insert(0, old_value)
        edit.select_range(0, "end")
        edit.focus_set()
        edit.place(x=x, y=y, width=w, height=h)

        def _save_edit(event=None):
            new_value = edit.get().strip()
            if new_value and new_value != old_value:
                self.result_tree.set(item, column, new_value)
                # Update the values tuple so export/use picks up changes
                vals = list(self.result_tree.item(item, "values"))
                vals[col_idx] = new_value
                self.result_tree.item(item, values=vals)
            edit.destroy()

        def _cancel_edit(event=None):
            edit.destroy()

        edit.bind("<Return>", _save_edit)
        edit.bind("<FocusOut>", _save_edit)
        edit.bind("<Escape>", _cancel_edit)

    def _build_bottom(self):
        bar = ttk.Frame(self.scroll_frame)
        bar.pack(fill="x", padx=8, pady=8)

        self.count_label = ttk.Label(bar, text="")
        self.count_label.pack(side="left")

        ttk.Button(bar, text="生成关键词", command=self._generate).pack(side="right", padx=3)
        ttk.Button(bar, text="清空", command=self._clear_results).pack(side="right", padx=3)
        ttk.Button(bar, text="删除选中", command=self._delete_selected).pack(side="right", padx=3)
        ttk.Button(bar, text="导出 Excel", command=self._export_excel).pack(side="right", padx=3)
        ttk.Button(bar, text="保存结果", command=self._save_results).pack(side="right", padx=3)
        ttk.Button(bar, text="加载结果", command=self._load_results).pack(side="right", padx=3)
        ttk.Button(bar, text="使用选中词搜索", command=self._use_selected).pack(side="right", padx=3)

    def _save_original_auto(self):
        """Auto-save original generated results (before any manual edits)."""
        if not hasattr(self, "_original_rows") or not self._original_rows:
            return
        save_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "keywords")
        os.makedirs(save_dir, exist_ok=True)
        import datetime
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(save_dir, f"keywords_original_{ts}.csv")
        data = []
        for i, row in enumerate(self._original_rows, 1):
            data.append({"序号": i, "国家": row[0], "语言": row[1], "检索关键词": row[2]})
        df = pd.DataFrame(data)
        df.to_csv(path, index=False, encoding="utf-8-sig")

    def _save_results(self):
        """Save current results (including manual edits) to CSV."""
        items = self.result_tree.get_children()
        if not items:
            messagebox.showinfo("提示", "没有可保存的结果。")
            return

        import datetime
        save_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "keywords")
        os.makedirs(save_dir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"keywords_{ts}.csv"

        filepath = filedialog.asksaveasfilename(
            title="保存当前结果",
            initialdir=save_dir,
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv"), ("Excel 文件", "*.xlsx")]
        )
        if not filepath:
            return

        data = []
        for item in items:
            vals = self.result_tree.item(item, "values")
            data.append({"序号": vals[0], "国家": vals[1], "语言": vals[2], "检索关键词": vals[3]})
        df = pd.DataFrame(data)
        try:
            if filepath.endswith(".xlsx"):
                df.to_excel(filepath, index=False, engine="openpyxl")
            else:
                df.to_csv(filepath, index=False, encoding="utf-8-sig")
            self.count_label.config(text=f"已保存 {len(data)} 条 → {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _load_results(self):
        """Merge keywords from a saved CSV file into the result tree.
        Only updates entries whose keyword has changed; keeps unmatched entries."""
        save_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "keywords")
        os.makedirs(save_dir, exist_ok=True)

        filepath = filedialog.askopenfilename(
            title="加载关键词结果（合并到当前列表）",
            initialdir=save_dir,
            filetypes=[("CSV 文件", "*.csv"), ("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        try:
            if filepath.endswith(".xlsx"):
                df = pd.read_excel(filepath, engine="openpyxl")
            else:
                df = pd.read_csv(filepath, encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("加载失败", f"读取文件出错: {e}")
            return

        required_cols = ["国家", "语言", "检索关键词"]
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("格式错误", f"文件缺少「{col}」列。")
                return

        if not self.result_tree.get_children():
            # Tree is empty: just load all
            for i, row in df.iterrows():
                idx = i + 1
                country = str(row.get("国家", ""))
                lang = str(row.get("语言", ""))
                keyword = str(row.get("检索关键词", ""))
                self.result_tree.insert("", "end", values=(idx, country, lang, keyword))
            self.count_label.config(text=f"已加载 {len(df)} 条 → {os.path.basename(filepath)}")
            return

        # Tree has entries: merge by (country, language) position
        loaded_groups = {}
        for _, row in df.iterrows():
            key = (str(row["国家"]), str(row["语言"]))
            loaded_groups.setdefault(key, []).append(str(row["检索关键词"]))

        tree_groups = {}
        tree_items = list(self.result_tree.get_children())
        for item in tree_items:
            vals = self.result_tree.item(item, "values")
            key = (vals[1], vals[2])
            tree_groups.setdefault(key, []).append(item)

        updated = 0
        added = 0

        for key, loaded_kws in loaded_groups.items():
            if key in tree_groups:
                existing_items = tree_groups[key]
                for pos, loaded_kw in enumerate(loaded_kws):
                    if pos < len(existing_items):
                        item = existing_items[pos]
                        old_vals = list(self.result_tree.item(item, "values"))
                        if old_vals[3] != loaded_kw:
                            old_vals[3] = loaded_kw
                            self.result_tree.item(item, values=old_vals)
                            updated += 1
                    else:
                        idx = len(tree_items) + added + 1
                        self.result_tree.insert("", "end",
                            values=(idx, key[0], key[1], loaded_kw))
                        added += 1
            else:
                for loaded_kw in loaded_kws:
                    idx = len(tree_items) + added + 1
                    self.result_tree.insert("", "end",
                        values=(idx, key[0], key[1], loaded_kw))
                    added += 1

        # Re-number
        for i, item in enumerate(self.result_tree.get_children(), 1):
            vals = list(self.result_tree.item(item, "values"))
            vals[0] = i
            self.result_tree.item(item, values=vals)

        self.count_label.config(
            text=f"合并完成: 更新 {updated} 条 + 新增 {added} 条 ← {os.path.basename(filepath)}")

    def _add_custom_keyword(self):
        kw = self.custom_kw_var.get().strip()
        if kw and kw not in self.kw_vars:
            self.kw_vars[kw] = tk.BooleanVar(value=True)
            tk.Checkbutton(self._kw_checkbutton_parent, text=kw, variable=self.kw_vars[kw]).pack(anchor="w")
            self.custom_kw_var.set("")

    def _select_all_kw(self):
        for var in self.kw_vars.values():
            var.set(True)

    def _clear_all_kw(self):
        for var in self.kw_vars.values():
            var.set(False)

    def _load_prebuilt_translations(self):
        """Pre-populate translation cache from the pre-built translations file."""
        if not KEYWORD_TRANSLATIONS:
            return
        count = 0
        for lang, mapping in KEYWORD_TRANSLATIONS.items():
            for chinese_text, translated in mapping.items():
                if chinese_text and translated:
                    self._translation_cache[(chinese_text, lang)] = translated
                    count += 1
        if count:
            pass  # silently loaded

    def _resolve_languages(self):
        request = []
        use_chinese = self.lang_vars["中文"].get()
        use_english = self.lang_vars["English"].get()
        use_official = self.lang_vars["official"].get()

        selected = [c for c in _all_countries() if self.country_vars[c[0]].get()]
        for country in selected:
            langs = []
            if use_chinese:
                langs.append("中文")
            if use_english and "English" not in langs:
                langs.append("English")
            if use_official:
                for l in _get_country_langs(country):
                    if l not in langs:
                        langs.append(l)
            if not langs and use_english:
                langs = ["English"]
            request.append((country, langs))
        return request

    def _translate_keyword(self, keyword, lang):
        if lang == "中文":
            return keyword
        if lang == "English":
            for i, kw in enumerate(CORE_KEYWORDS["zh"]):
                if kw == keyword:
                    if i < len(CORE_KEYWORDS.get("en", [])):
                        return CORE_KEYWORDS["en"][i]
                    break
            return keyword
        # Other languages: use AI translation cache
        cached = self._translation_cache.get((keyword, lang))
        if cached:
            return cached
        # Fallback: return English translation if available, otherwise original
        for i, kw in enumerate(CORE_KEYWORDS["zh"]):
            if kw == keyword:
                if i < len(CORE_KEYWORDS.get("en", [])):
                    return CORE_KEYWORDS["en"][i]
                break
        return keyword

    def _translate_country(self, country_record, lang):
        if lang == "中文":
            return country_record[0]
        if lang == "English":
            return country_record[1]
        # Other languages: translate Chinese country name to target language
        chinese_name = country_record[0]
        cached = self._translation_cache.get((chinese_name, lang))
        if cached:
            return cached
        # Fallback to English name
        return country_record[1]

    def _batch_translate(self, texts_by_lang):
        """Batch translate texts grouped by target language via AI API.
        texts_by_lang: {lang: set of texts}
        Fills self._translation_cache.
        """
        api_key = self.api_key_var.get().strip()
        if not api_key:
            return

        client = openai.OpenAI(
            base_url=YUNWU_BASE_URL,
            api_key=api_key,
        )

        for lang, texts in texts_by_lang.items():
            texts = [t for t in texts if t and (t, lang) not in self._translation_cache]
            if not texts:
                continue

            texts_list = list(texts)
            prompt = f"""Translate the following Chinese phrases to {lang}.
Return a JSON object mapping each original Chinese phrase to its translation in {lang}.

Phrases to translate:
{json.dumps(texts_list, ensure_ascii=False)}

Return ONLY the JSON object, nothing else. Example format:
{{"中式建筑": "translated text", "日本": "translated text"}}"""

            try:
                response = client.chat.completions.create(
                    model=DEFAULT_MODEL,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=4000,
                )
                content = response.choices[0].message.content.strip()
                json_match = re.search(r'\{[\s\S]*\}', content)
                if json_match:
                    result = json.loads(json_match.group(0))
                    for original, translated in result.items():
                        if isinstance(translated, str) and translated.strip():
                            self._translation_cache[(original, lang)] = translated.strip()
            except Exception:
                pass

    def _generate(self):
        if not any(v.get() for v in self.country_vars.values()):
            messagebox.showinfo("提示", "请先选择国家。")
            return

        selected_kw = [kw for kw, var in self.kw_vars.items() if var.get()]
        selected_academic = [t for t, var in self.academic_vars.items() if var.get()]
        selected_generic = [t for t, var in self.generic_vars.items() if var.get()]
        if not selected_kw and not selected_academic and not selected_generic:
            messagebox.showinfo("提示", "请至少选择一个核心关键词、学术词语或通用关键词。")
            return

        lang_request = self._resolve_languages()
        if not any(langs for _, langs in lang_request):
            messagebox.showinfo("提示", "请至少选择一种检索语言。")
            return

        # Collect texts that need AI translation, grouped by target language
        texts_by_lang = {}
        for country_record, langs in lang_request:
            for lang in langs:
                if lang in ("中文", "English"):
                    continue
                if lang not in texts_by_lang:
                    texts_by_lang[lang] = set()
                # Add Chinese country name for translation
                texts_by_lang[lang].add(country_record[0])
                # Add Chinese keywords for translation
                for kw in selected_kw:
                    texts_by_lang[lang].add(kw)
                for term in selected_generic:
                    texts_by_lang[lang].add(term)

        if texts_by_lang:
            self.count_label.config(text="正在 AI 翻译中，请稍候...")
            self.win.update()
            self._batch_translate(texts_by_lang)

        all_rows = []
        for country_record, langs in lang_request:
            for lang in langs:
                exclude_country = self._exclude_country_var.get()
                country_name = self._translate_country(country_record, lang)
                # Build generic suffix (appended to every keyword)
                generic_suffix = ""
                for term in selected_generic:
                    term_translated = self._translate_keyword(term, lang)
                    generic_suffix += f" {term_translated}"
                for kw in selected_kw:
                    kw_translated = self._translate_keyword(kw, lang)
                    if exclude_country:
                        keyword = kw_translated + generic_suffix
                    else:
                        keyword = f"{country_name} {kw_translated}{generic_suffix}"
                    all_rows.append((country_record[0], lang, keyword))
                for term in selected_academic:
                    if exclude_country:
                        keyword = term + generic_suffix
                    else:
                        keyword = f"{country_name} {term}{generic_suffix}"
                    all_rows.append((country_record[0], lang, keyword))
                # If no core or academic selected, output generic alone
                if not selected_kw and not selected_academic:
                    for term in selected_generic:
                        term_translated = self._translate_keyword(term, lang)
                        if exclude_country:
                            keyword = term_translated
                        else:
                            keyword = f"{country_name} {term_translated}"
                        all_rows.append((country_record[0], lang, keyword))

        start_idx = len(self.result_tree.get_children()) + 1

        # Extend snapshot
        if not hasattr(self, '_original_rows') or self._original_rows is None:
            self._original_rows = []
        self._original_rows.extend([list(r) for r in all_rows])
        self._save_original_auto()

        for i, row in enumerate(all_rows, start_idx):
            self.result_tree.insert("", "end", values=(i, row[0], row[1], row[2]))

        count = len(self.result_tree.get_children())
        self.count_label.config(text=f"共生成 {count} 条关键词")

    def _clear_results(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self._original_rows = []
        self.count_label.config(text="")

    def _delete_selected(self):
        sel = self.result_tree.selection()
        if not sel:
            return
        for item in sel:
            self.result_tree.delete(item)
        # Re-index
        for i, item in enumerate(self.result_tree.get_children(), 1):
            vals = list(self.result_tree.item(item, "values"))
            vals[0] = i
            self.result_tree.item(item, values=vals)
        self.count_label.config(text=f"共生成 {len(self.result_tree.get_children())} 条关键词")

    def _export_excel(self):
        items = self.result_tree.get_children()
        if not items:
            messagebox.showinfo("提示", "请先生成关键词。")
            return
        filepath = filedialog.asksaveasfilename(
            title="导出关键词", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")]
        )
        if not filepath:
            return
        data = []
        for item in items:
            vals = self.result_tree.item(item, "values")
            data.append({"序号": vals[0], "国家": vals[1], "语言": vals[2], "检索关键词": vals[3]})
        df = pd.DataFrame(data)
        try:
            if filepath.endswith(".csv"):
                df.to_csv(filepath, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(filepath, index=False, engine="openpyxl")
            messagebox.showinfo("导出完成", f"已导出 {len(data)} 条。")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _use_selected(self):
        sel = self.result_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择要使用的关键词。")
            return
        keywords = [self.result_tree.item(s, "values")[3] for s in sel]
        if self.callback:
            self.callback(keywords)
        self.win.destroy()
