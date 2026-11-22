# -*- coding: utf-8 -*-
"""国家中英对照 + M49 五大洲归类（单一数据源，供软件地图与补充数据表共用）。

cases 表 country 字段为英文 M49 短名；本模块提供：
  - M49_REGION : 英文 M49 名 -> (五大区域, 亚洲亚区)   [亚区用 M49 标准拼写 South-eastern Asia]
  - EN_TO_ZH   : 英文 M49 名 -> 中文名（与 city_classification.csv 的 Country 列一致）
  - ZH_TO_EN   : 中文名 -> 英文 M49 名

论文口径（正文脚注）：伊朗按地缘文化归中亚（M49 为南亚）；俄罗斯归东欧（M49 归欧洲）。
"""

# 英文 M49 -> (Region, Subregion)。非亚洲国家 Subregion 为空。
M49_REGION = {
    # ---- Asia / Eastern Asia ----
    "Japan": ("Asia", "Eastern Asia"),
    "Republic of Korea": ("Asia", "Eastern Asia"),
    "Democratic People's Republic of Korea": ("Asia", "Eastern Asia"),
    "Mongolia": ("Asia", "Eastern Asia"),
    # ---- Asia / South-eastern Asia ----
    "Thailand": ("Asia", "South-eastern Asia"),
    "Malaysia": ("Asia", "South-eastern Asia"),
    "Viet Nam": ("Asia", "South-eastern Asia"),
    "Singapore": ("Asia", "South-eastern Asia"),
    "Indonesia": ("Asia", "South-eastern Asia"),
    "Philippines": ("Asia", "South-eastern Asia"),
    "Cambodia": ("Asia", "South-eastern Asia"),
    "Myanmar": ("Asia", "South-eastern Asia"),
    "Lao People's Democratic Republic": ("Asia", "South-eastern Asia"),
    "Timor-Leste": ("Asia", "South-eastern Asia"),
    "Brunei Darussalam": ("Asia", "South-eastern Asia"),
    # ---- Asia / Southern Asia ----
    "India": ("Asia", "Southern Asia"),
    "Bhutan": ("Asia", "Southern Asia"),
    "Nepal": ("Asia", "Southern Asia"),
    # ---- Asia / Central Asia (伊朗按论文口径归中亚) ----
    "Kyrgyzstan": ("Asia", "Central Asia"),
    "Kazakhstan": ("Asia", "Central Asia"),
    "Iran (Islamic Republic of)": ("Asia", "Central Asia"),
    # ---- Asia / Western Asia ----
    "Türkiye": ("Asia", "Western Asia"),
    # ---- Europe ----
    "France": ("Europe", ""),
    "Russian Federation": ("Europe", ""),
    "Germany": ("Europe", ""),
    "United Kingdom of Great Britain and Northern Ireland": ("Europe", ""),
    "Sweden": ("Europe", ""),
    "Netherlands": ("Europe", ""),
    "Italy": ("Europe", ""),
    "Denmark": ("Europe", ""),
    "Ukraine": ("Europe", ""),
    "Poland": ("Europe", ""),
    "Spain": ("Europe", ""),
    "Belgium": ("Europe", ""),
    "Belarus": ("Europe", ""),
    "Austria": ("Europe", ""),
    "Switzerland": ("Europe", ""),
    "Hungary": ("Europe", ""),
    "Czechia": ("Europe", ""),
    "Republic of Moldova": ("Europe", ""),
    "Portugal": ("Europe", ""),
    "Latvia": ("Europe", ""),
    "Ireland": ("Europe", ""),
    # ---- Americas ----
    "United States of America": ("Americas", ""),
    "Canada": ("Americas", ""),
    "Cuba": ("Americas", ""),
    "Mexico": ("Americas", ""),
    "Brazil": ("Americas", ""),
    "Peru": ("Americas", ""),
    "Panama": ("Americas", ""),
    "Guyana": ("Americas", ""),
    "Ecuador": ("Americas", ""),
    "Chile": ("Americas", ""),
    "Argentina": ("Americas", ""),
    # ---- Oceania ----
    "Australia": ("Oceania", ""),
    "New Zealand": ("Oceania", ""),
    "Micronesia (Federated States of)": ("Oceania", ""),
    "Kiribati": ("Oceania", ""),
    # ---- Africa ----
    "Mauritius": ("Africa", ""),
    "Zimbabwe": ("Africa", ""),
    "Zambia": ("Africa", ""),
    "South Africa": ("Africa", ""),
    "Nigeria": ("Africa", ""),
    "Seychelles": ("Africa", ""),
    "Namibia": ("Africa", ""),
    "Malawi": ("Africa", ""),
    "Madagascar": ("Africa", ""),
    "Botswana": ("Africa", ""),
    "Benin": ("Africa", ""),
}

# 英文 M49 -> 中文（与 city_classification.csv / 所有国家中英文.xlsx 的「中文」列一致）
EN_TO_ZH = {
    "Japan": "日本", "Thailand": "泰国", "Malaysia": "马来西亚",
    "Republic of Korea": "韩国", "Viet Nam": "越南",
    "United States of America": "美国", "Singapore": "新加坡",
    "Indonesia": "印度尼西亚", "France": "法国",
    "Russian Federation": "俄罗斯", "Australia": "澳大利亚", "Germany": "德国",
    "Canada": "加拿大",
    "United Kingdom of Great Britain and Northern Ireland": "英国",
    "Democratic People's Republic of Korea": "朝鲜", "Sweden": "瑞典",
    "Philippines": "菲律宾", "Mongolia": "蒙古", "Netherlands": "荷兰",
    "Italy": "意大利", "New Zealand": "新西兰", "Denmark": "丹麦",
    "Cambodia": "柬埔寨", "Myanmar": "缅甸", "Ukraine": "乌克兰",
    "Poland": "波兰", "India": "印度",
    "Lao People's Democratic Republic": "老挝", "Spain": "西班牙",
    "Mauritius": "毛里求斯", "Kyrgyzstan": "吉尔吉斯斯坦", "Cuba": "古巴",
    "Bhutan": "不丹", "Belgium": "比利时", "Belarus": "白俄罗斯",
    "Austria": "奥地利", "Zimbabwe": "津巴布韦", "Zambia": "赞比亚",
    "Switzerland": "瑞士", "South Africa": "南非", "Nigeria": "尼日利亚",
    "Mexico": "墨西哥", "Kazakhstan": "哈萨克斯坦",
    "Iran (Islamic Republic of)": "伊朗", "Hungary": "匈牙利",
    "Czechia": "捷克", "Brazil": "巴西", "Türkiye": "土耳其",
    "Timor-Leste": "东帝汶", "Seychelles": "塞舌尔",
    "Republic of Moldova": "摩尔多瓦共和国", "Portugal": "葡萄牙",
    "Peru": "秘鲁", "Panama": "巴拿马", "Nepal": "尼泊尔",
    "Namibia": "纳米比亚",
    "Micronesia (Federated States of)": "密克罗尼西亚联邦",
    "Malawi": "马拉维", "Madagascar": "马达加斯加", "Latvia": "拉脱维亚",
    "Kiribati": "基里巴斯", "Ireland": "爱尔兰", "Guyana": "圭亚那",
    "Ecuador": "厄瓜多尔", "Chile": "智利", "Brunei Darussalam": "文莱",
    "Botswana": "博茨瓦纳", "Benin": "贝宁", "Argentina": "阿根廷",
}

ZH_TO_EN = {zh: en for en, zh in EN_TO_ZH.items()}

REGION_ORDER = ["Asia", "Europe", "Americas", "Oceania", "Africa"]
ASIA_SUB_ORDER = ["South-eastern Asia", "Eastern Asia", "Southern Asia",
                  "Central Asia", "Western Asia"]
