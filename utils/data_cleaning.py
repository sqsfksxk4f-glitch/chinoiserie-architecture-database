import pandas as pd


def clean_whitespace(df):
    """去除所有文本列的首尾空白"""
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "None": "", "NaT": ""})
    return df


def normalize_province(series):
    """统一省份名称（补全'省'/'市'后缀）"""
    province_map = {
        "北京": "北京市", "上海": "上海市", "天津": "天津市", "重庆": "重庆市",
        "广东": "广东省", "江苏": "江苏省", "浙江": "浙江省", "山东": "山东省",
        "河南": "河南省", "河北": "河北省", "四川": "四川省", "湖北": "湖北省",
        "湖南": "湖南省", "福建": "福建省", "安徽": "安徽省", "江西": "江西省",
        "辽宁": "辽宁省", "陕西": "陕西省", "云南": "云南省", "贵州": "贵州省",
        "山西": "山西省", "甘肃": "甘肃省", "吉林": "吉林省", "黑龙江": "黑龙江省",
        "海南": "海南省", "青海": "青海省", "台湾": "台湾省",
        "广西": "广西壮族自治区", "内蒙古": "内蒙古自治区",
        "宁夏": "宁夏回族自治区", "新疆": "新疆维吾尔自治区", "西藏": "西藏自治区",
        "香港": "香港特别行政区", "澳门": "澳门特别行政区",
    }
    return series.replace(province_map)


def find_duplicates_in_df(df, subset=None):
    """在 DataFrame 中查找重复行"""
    if subset is None:
        subset = ["建筑名称", "省份"]
    available = [c for c in subset if c in df.columns]
    if not available:
        return pd.DataFrame()
    dup_mask = df.duplicated(subset=available, keep=False)
    return df[dup_mask].sort_values(by=available)


def get_missing_summary(df):
    """统计各字段的缺失情况"""
    missing = {}
    for col in df.columns:
        count = df[col].isna().sum() + (df[col].astype(str).str.strip() == "").sum()
        if count > 0:
            missing[col] = count
    return missing
