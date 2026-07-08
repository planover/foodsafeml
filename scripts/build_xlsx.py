# -*- coding: utf-8 -*-
"""
FoodSafeML (foodsafeml) v1.8.0 — 工作簿生成器（全量矩阵版）
==========================================================
给定某食品目录下的 basic.json / translations.json / risks.json，生成
<食品名>_食品安全风险识别表.xlsx（3 个工作表）。

v1.8.0 核心升级（全量矩阵）：
  ① Sheet3 食品安全风险识别表：全量指标清单（5大类 72 项）× 有数据地区，
     无数据的指标也列出框架行（值标 [待填写]），确保指标维度完整。
  ② Sheet4 指标对比查询：全量指标(72) × 全量地区(36) 完整矩阵，
     每个单元格有数据则 HYPERLINK 跳转源表，无数据标 [待填写]。

防幻觉规则：缺失的多语字段 / 具体物质一律以 [待填写] 标注，绝不编造。

用法：
    python build_xlsx.py <食品目录> [<食品中文名>]
例：
    python build_xlsx.py ./food-safety-芒果干 芒果干
"""
import os
import re
import json
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SKILL_VER = "1.8.0"

REGION_INFO = {
    "CN": ("中国", "China", "中国"), "HK": ("中国香港", "Hong Kong", "香港"),
    "MO": ("中国澳门", "Macao", "澳門"), "TW": ("中国台湾", "Taiwan", "臺灣"),
    "JP": ("日本", "Japan", "日本"), "KR": ("韩国", "Korea", "한국"),
    "US": ("美国", "United States", "United States"), "CA": ("加拿大", "Canada", "Canada"),
    "BR": ("巴西", "Brazil", "Brasil"), "MX": ("墨西哥", "Mexico", "México"),
    "EU": ("欧盟", "European Union", "Europäische Union"), "UK": ("英国", "United Kingdom", "United Kingdom"),
    "FR": ("法国", "France", "France"), "DE": ("德国", "Germany", "Deutschland"),
    "ES": ("西班牙", "Spain", "España"), "PT": ("葡萄牙", "Portugal", "Portugal"),
    "AU/NZ": ("澳大利亚/新西兰", "Australia/New Zealand", "Australia/New Zealand"),
    "CODEX": ("国际食品法典", "Codex Alimentarius", "Codex Alimentarius"),
    "WHO/JECFA": ("联合国粮农/世卫 食品添加剂专家委员会", "WHO/FAO JECFA", "JECFA"),
    "ISO": ("国际标准化组织", "ISO", "ISO"),
    "TH": ("泰国", "Thailand", "ประเทศไทย"), "VN": ("越南", "Vietnam", "Việt Nam"),
    "KH": ("柬埔寨", "Cambodia", "កម្ពុជា"), "ID": ("印度尼西亚", "Indonesia", "Indonesia"),
    "MY": ("马来西亚", "Malaysia", "Malaysia"), "PH": ("菲律宾", "Philippines", "Pilipinas"),
    "SG": ("新加坡", "Singapore", "Singapore"), "IN": ("印度", "India", "India"),
    "RU/EAEU": ("俄罗斯/欧亚经济联盟", "Russia/EAEU", "Россия/EАЭС"),
    "ASEAN": ("东盟", "ASEAN", "ASEAN"), "MERCOSUR": ("南方共同市场", "MERCOSUR", "MERCOSUR"),
    "CARICOM": ("加勒比共同体", "CARICOM", "CARICOM"), "SPC/PIF": ("太平洋共同体/论坛", "Pacific Community/Forum", "Pacific"),
    "WTO/SPS": ("世贸组织/SPS", "WTO/SPS", "WTO/SPS"), "APEC": ("亚太经合", "APEC", "APEC"),
    "综合": ("综合/多辖区协调框架", "Multi-jurisdiction framework", "Framework"),
}
ALL_REGIONS = list(REGION_INFO.keys())

# 地区代码 -> 官方全称（用于 Sheet4 列头）
REGION_FULLNAME = {
    "CN": "中华人民共和国 国家卫生健康委员会 / 国家食品安全风险评估中心（GB 标准）",
    "HK": "中国香港特别行政区政府 食物环境卫生署食物安全中心",
    "MO": "中国澳门特别行政区政府 市政署",
    "TW": "中国台湾地区 卫生福利部食品药物管理署",
    "JP": "日本国 厚生劳动省",
    "KR": "大韩民国 食品药品安全部（MFDS）",
    "US": "美国 食品药品监督管理局（FDA）/ 环境保护署（EPA）",
    "CA": "加拿大 卫生部（Health Canada）",
    "BR": "巴西联邦共和国 国家卫生监督局（ANVISA）",
    "MX": "墨西哥合众国 卫生部（COFEPRIS）",
    "EU": "欧洲联盟委员会（欧盟法规 / European Commission）",
    "UK": "大不列颠及北爱尔兰联合王国 食品标准局（FSA）",
    "FR": "法兰西共和国 竞争、消费与反欺诈总局（DGCCRF）",
    "DE": "德意志联邦共和国 联邦食品与农业部（BMEL）",
    "ES": "西班牙王国 食品安全与营养局（AESAN）",
    "PT": "葡萄牙共和国 食品安全局（ASAE）",
    "AU/NZ": "澳大利亚/新西兰食品标准局（FSANZ）",
    "CODEX": "国际食品法典委员会（Codex Alimentarius Commission）",
    "WHO/JECFA": "联合国粮农组织/世界卫生组织 食品添加剂联合专家委员会（JECFA）",
    "ISO": "国际标准化组织（International Organization for Standardization）",
    "TH": "泰王国 食品药品监督管理局（FDA Thailand）",
    "VN": "越南社会主义共和国 卫生部食品安全局",
    "KH": "柬埔寨王国 工业与手工业部标准局",
    "ID": "印度尼西亚共和国 食品药品监督局（BPOM）",
    "MY": "马来西亚 卫生部",
    "PH": "菲律宾共和国 食品药品监督管理局（FDA Philippines）",
    "SG": "新加坡共和国 食品局（SFA）",
    "IN": "印度共和国 食品安全标准局（FSSAI）",
    "RU/EAEU": "俄罗斯联邦 消费者权益及公民平安保护监督局（Rospotrebnadzor）/ 欧亚经济联盟",
    "ASEAN": "东南亚国家联盟（ASEAN）",
    "MERCOSUR": "南方共同市场（MERCOSUR）",
    "CARICOM": "加勒比共同体（CARICOM）",
    "SPC/PIF": "太平洋共同体（SPC）/ 太平洋岛国论坛",
    "WTO/SPS": "世界贸易组织 卫生与植物卫生措施委员会（WTO/SPS）",
    "APEC": "亚太经济合作组织（Asia-Pacific Economic Cooperation, APEC）",
    "综合": "综合/多辖区协调框架",
}

# 简<->繁 转换表
PAIRS = [
    ("铅", "鉛"), ("镉", "鎘"), ("亚", "亞"), ("盐", "鹽"), ("农", "農"), ("药", "藥"),
    ("残", "殘"), ("黄", "黃"), ("霉", "黴"), ("标", "標"), ("准", "準"), ("规", "規"),
    ("剂", "劑"), ("检", "檢"), ("验", "驗"), ("签", "籤"), ("过", "過"), ("敏", "敏"),
    ("总", "總"), ("数", "數"), ("肠", "腸"), ("杆", "桿"), ("门", "門"), ("干", "乾"),
    ("类", "類"), ("测", "測"), ("氯", "氯"), ("氰", "氰"), ("菊", "菊"), ("酯", "酯"),
    ("醚", "醚"), ("唑", "唑"), ("啉", "啉"), ("胺", "胺"), ("脒", "脒"), ("虫", "蟲"),
    ("噻", "噻"), ("嗪", "嗪"), ("威", "威"), ("鲜", "鮮"), ("锰", "錳"), ("吡", "吡"),
    ("嘧", "嘧"), ("灵", "靈"), ("苯", "苯"), ("环", "環"), ("啶", "啶"), ("磺", "磺"),
    ("钠", "鈉"), ("钾", "鉀"), ("钙", "鈣"), ("镁", "鎂"), ("铁", "鐵"), ("铜", "銅"), ("锌", "鋅"),
    ("镍", "鎳"), ("铬", "鉻"), ("锡", "錫"), ("汞", "汞"), ("砷", "砷"), ("硝", "硝"),
    ("芘", "芘"), ("胺", "胺"), ("联", "聯"), ("丙", "丙"), ("醇", "醇"), ("菌", "菌"),
    ("霉", "黴"), ("酵", "酵"), ("母", "母"), ("希", "希"), ("氏", "氏"),
]
SIMP2TRAD = {s: t for s, t in PAIRS}
TRAD2SIMP = {t: s for s, t in SIMP2TRAD.items()}

# ====== 全量指标清单（5 大类 72 项，按类别有序）======
# (类别, 简中名, 英文名, 关联代码, 标准框架)
ALL_INDICATORS = [
    # A. 食品添加剂（GB 2760-2024）
    ("食品添加剂", "二氧化硫/SO₂残留", "Sulfur dioxide / Sulphite residues (SO₂)", "INS E220", "GB 2760-2024"),
    ("食品添加剂", "焦亚硫酸钠", "Sodium metabisulfite", "INS 223", "GB 2760-2024"),
    ("食品添加剂", "苯甲酸及其钠盐", "Benzoic acid / Sodium benzoate", "INS 210/211", "GB 2760-2024"),
    ("食品添加剂", "山梨酸及其钾盐", "Sorbic acid / Potassium sorbate", "INS 200/202", "GB 2760-2024"),
    ("食品添加剂", "丙酸及其钠盐", "Propionic acid / Sodium propionate", "INS 280/281", "GB 2760-2024"),
    ("食品添加剂", "糖精钠", "Sodium saccharin", "INS 954", "GB 2760-2024"),
    ("食品添加剂", "甜蜜素(环己基氨基磺酸钠)", "Sodium cyclamate", "INS 952", "GB 2760-2024"),
    ("食品添加剂", "安赛蜜(乙酰磺胺酸钾)", "Acesulfame potassium", "INS 950", "GB 2760-2024"),
    ("食品添加剂", "三氯蔗糖(蔗糖素)", "Sucralose", "INS 955", "GB 2760-2024"),
    ("食品添加剂", "阿斯巴甜", "Aspartame", "INS 951", "GB 2760-2024"),
    ("食品添加剂", "胭脂红", "Ponceau 4R", "INS 124", "GB 2760-2024"),
    ("食品添加剂", "柠檬黄", "Tartrazine", "INS 102", "GB 2760-2024"),
    ("食品添加剂", "日落黄", "Sunset yellow FCF", "INS 110", "GB 2760-2024"),
    ("食品添加剂", "诱惑红", "Allura red AC", "INS 129", "GB 2760-2024"),
    ("食品添加剂", "亮蓝", "Brilliant blue FCF", "INS 133", "GB 2760-2024"),
    ("食品添加剂", "赤藓红", "Erythrosine", "INS 127", "GB 2760-2024"),
    ("食品添加剂", "柠檬酸", "Citric acid", "INS 330", "GB 2760-2024"),
    ("食品添加剂", "乳酸", "Lactic acid", "INS 270", "GB 2760-2024"),
    # B. 农药残留（GB 2763-2026）
    ("农药残留", "毒死蜱", "Chlorpyrifos", "CAS 2921-88-2", "GB 2763-2026"),
    ("农药残留", "吡虫啉", "Imidacloprid", "CAS 138261-41-3", "GB 2763-2026"),
    ("农药残留", "氯氰菊酯", "Cypermethrin", "CAS 52315-07-8", "GB 2763-2026"),
    ("农药残留", "高效氯氰菊酯", "Beta-cypermethrin", "CAS 65731-84-2", "GB 2763-2026"),
    ("农药残留", "联苯菊酯", "Bifenthrin", "CAS 82657-04-3", "GB 2763-2026"),
    ("农药残留", "咪鲜胺", "Prochloraz", "CAS 67747-09-5", "GB 2763-2026"),
    ("农药残留", "啶虫脒", "Acetamiprid", "CAS 135410-20-7", "GB 2763-2026"),
    ("农药残留", "吡唑醚菌酯", "Pyraclostrobin", "CAS 175013-18-0", "GB 2763-2026"),
    ("农药残留", "噻虫嗪", "Thiamethoxam", "CAS 153719-23-4", "GB 2763-2026"),
    ("农药残留", "克百威", "Carbofuran", "CAS 1563-66-2", "GB 2763-2026"),
    ("农药残留", "多菌灵", "Carbendazim", "CAS 10605-21-7", "GB 2763-2026"),
    ("农药残留", "苯醚甲环唑", "Difenoconazole", "CAS 119446-68-3", "GB 2763-2026"),
    ("农药残留", "甲氨基阿维菌素苯甲酸盐", "Emamectin benzoate", "CAS 155569-91-8", "GB 2763-2026"),
    ("农药残留", "氟氯氰菊酯", "Cyfluthrin", "CAS 68359-37-5", "GB 2763-2026"),
    ("农药残留", "阿维菌素", "Abamectin", "CAS 71751-41-2", "GB 2763-2026"),
    ("农药残留", "氯菊酯", "Permethrin", "CAS 52645-53-1", "GB 2763-2026"),
    ("农药残留", "溴氰菊酯", "Deltamethrin", "CAS 52918-63-5", "GB 2763-2026"),
    ("农药残留", "氰戊菊酯", "Fenvalerate", "CAS 51630-58-1", "GB 2763-2026"),
    ("农药残留", "敌敌畏", "Dichlorvos", "CAS 62-73-7", "GB 2763-2026"),
    ("农药残留", "乐果", "Dimethoate", "CAS 60-51-5", "GB 2763-2026"),
    ("农药残留", "百菌清", "Chlorothalonil", "CAS 1897-45-6", "GB 2763-2026"),
    ("农药残留", "噻螨酮", "Hexythiazox", "CAS 78587-05-0", "GB 2763-2026"),
    # C. 污染物（GB 2762-2025）
    ("污染物", "铅(Pb)", "Lead (Pb)", "CAS 7439-92-1", "GB 2762-2025"),
    ("污染物", "镉(Cd)", "Cadmium (Cd)", "CAS 7440-43-9", "GB 2762-2025"),
    ("污染物", "汞(Hg)", "Mercury (Hg)", "CAS 7439-97-6", "GB 2762-2025"),
    ("污染物", "砷(As)", "Arsenic (As)", "CAS 7440-38-2", "GB 2762-2025"),
    ("污染物", "铬(Cr)", "Chromium (Cr)", "CAS 7440-47-3", "GB 2762-2025"),
    ("污染物", "锡(Sn)", "Tin (Sn)", "CAS 7440-31-5", "GB 2762-2025"),
    ("污染物", "镍(Ni)", "Nickel (Ni)", "CAS 7440-02-0", "GB 2762-2025"),
    ("污染物", "亚硝酸盐", "Nitrite", "NO₂⁻", "GB 2762-2025"),
    ("污染物", "硝酸盐", "Nitrate", "NO₃⁻", "GB 2762-2025"),
    ("污染物", "苯并[a]芘", "Benzo[a]pyrene", "CAS 50-32-8", "GB 2762-2025"),
    ("污染物", "N-二甲基亚硝胺", "N-Nitrosodimethylamine (NDMA)", "CAS 62-75-9", "GB 2762-2025"),
    ("污染物", "多氯联苯", "Polychlorinated biphenyls (PCBs)", "", "GB 2762-2025"),
    ("污染物", "3-氯-1,2-丙二醇(3-MCPD)", "3-Monochloropropane-1,2-diol", "CAS 96-24-2", "GB 2762-2025"),
    # D. 真菌毒素（GB 2761）
    ("真菌毒素", "黄曲霉毒素B1", "Aflatoxin B1", "CAS 1162-65-8", "GB 2761"),
    ("真菌毒素", "黄曲霉毒素M1", "Aflatoxin M1", "CAS 6795-23-9", "GB 2761"),
    ("真菌毒素", "脱氧雪腐镰刀菌烯醇(DON)", "Deoxynivalenol (DON)", "CAS 51481-10-8", "GB 2761"),
    ("真菌毒素", "展青霉素", "Patulin", "CAS 149-29-1", "GB 2761"),
    ("真菌毒素", "赭曲霉毒素A", "Ochratoxin A", "CAS 303-47-9", "GB 2761"),
    ("真菌毒素", "玉米赤霉烯酮", "Zearalenone", "CAS 17924-92-4", "GB 2761"),
    # E. 微生物/致病菌（GB 29921-2021 + GB 14884-2016）
    ("微生物/致病菌", "菌落总数", "Aerobic plate count (APC)", "", "GB 14884-2016"),
    ("微生物/致病菌", "大肠菌群", "Coliforms", "", "GB 14884-2016"),
    ("微生物/致病菌", "大肠杆菌", "Escherichia coli (E. coli)", "", "GB 29921-2021"),
    ("微生物/致病菌", "沙门氏菌", "Salmonella spp.", "", "GB 29921-2021"),
    ("微生物/致病菌", "金黄色葡萄球菌", "Staphylococcus aureus", "", "GB 29921-2021"),
    ("微生物/致病菌", "蜡样芽胞杆菌", "Bacillus cereus", "", "GB 14884-2016"),
    ("微生物/致病菌", "霉菌", "Moulds", "", "GB 14884-2016"),
    ("微生物/致病菌", "酵母", "Yeasts", "", "GB 14884-2016"),
    ("微生物/致病菌", "单核细胞增生李斯特氏菌", "Listeria monocytogenes", "", "GB 29921-2021"),
    ("微生物/致病菌", "致泻大肠埃希氏菌", "Diarrheagenic Escherichia coli", "", "GB 29921-2021"),
    ("微生物/致病菌", "副溶血性弧菌", "Vibrio parahaemolyticus", "", "GB 29921-2021"),
    ("微生物/致病菌", "克罗诺杆菌属", "Cronobacter spp.", "", "GB 29921-2021"),
    ("微生物/致病菌", "志贺氏菌", "Shigella spp.", "", "GB 29921-2021"),
]
# 全量指标名 -> (类别, 英文名, 代码, 标准框架)
ALL_INDICATOR_MAP = {e[1]: e for e in ALL_INDICATORS}

PEST_EN = {
    "咪鲜胺": "Prochloraz", "啶虫脒": "Acetamiprid", "吡唑醚菌酯": "Pyraclostrobin",
    "噻虫嗪": "Thiamethoxam", "克百威": "Carbofuran", "多菌灵": "Carbendazim",
    "苯醚甲环唑": "Difenoconazole", "吡虫啉": "Imidacloprid", "毒死蜱": "Chlorpyrifos",
    "氯氰菊酯": "Cypermethrin", "高效氯氰菊酯": "Beta-cypermethrin", "联苯菊酯": "Bifenthrin",
    "甲氨基阿维菌素": "Emamectin benzoate", "氟氯氰菊酯": "Cyfluthrin",
}

STD_EN = {
    "GB 2760-2024": "National Food Safety Standard – Standards for Uses of Food Additives (GB 2760-2024)",
    "GB 2762-2025": "National Food Safety Standard – Maximum Levels of Contaminants in Foods (GB 2762-2025)",
    "GB 2762-2022": "National Food Safety Standard – Maximum Levels of Contaminants in Foods (GB 2762-2022)",
    "GB 2763-2026": "National Food Safety Standard – Maximum Residue Limits for Pesticides in Food (GB 2763-2026)",
    "GB 2763-2021": "National Food Safety Standard – Maximum Residue Limits for Pesticides in Food (GB 2763-2021)",
    "GB 2761": "National Food Safety Standard – Maximum Levels of Mycotoxins in Foods (GB 2761)",
    "GB 29921-2021": "National Food Safety Standard – Pathogen Limit in Pre-packaged Foods (GB 29921-2021)",
    "GB 14884-2016": "National Food Safety Standard for Candied Fruit (GB 14884-2016)",
    "GB 5009.34": "National Food Safety Standard – Determination of Sulfur Dioxide in Foods (GB 5009.34)",
    "GB 7718-2011": "National Food Safety Standard – General Rules for the Labelling of Prepackaged Foods (GB 7718-2011)",
    "GB 2763": "National Food Safety Standard – Maximum Residue Limits for Pesticides in Food (GB 2763)",
    "(EC) 1881/2006": "Commission Regulation (EC) No 1881/2006 setting maximum levels for certain contaminants in foodstuffs",
    "(EC) 1169/2011": "Regulation (EU) No 1169/2011 on the provision of food information to consumers",
    "(EC) 1333/2008": "Regulation (EC) No 1333/2008 on food additives",
    "(EC) 396/2005": "Regulation (EC) No 396/2005 on maximum residue levels of pesticides in or on food and feed of plant and animal origin",
    "21 CFR 130.9": "U.S. Code of Federal Regulations, 21 CFR 130.9 (sulfiting agents – declaration)",
    "C.R.C., c. 870": "Canada – Food and Drug Regulations (C.R.C., c. 870)",
    "ISO 5522:1981": "ISO 5522:1981 – Dried fruits – Determination of sulfur dioxide content",
}
STD_NATIVE = {
    "(EC) 1881/2006": {"EU": "Verordnung (EG) Nr. 1881/2006"},
    "(EC) 396/2005": {"EU": "Verordnung (EG) Nr. 396/2005"},
}
QUOTE_CN = {
    "(EC) 1881/2006": "干果及其制品（直接食用或作配料）中黄曲霉毒素 B1≤2.0 µg/kg，B1+B2+G1+G2 总和≤4.0 µg/kg。",
    "(EU) 1169/2011": "当二氧化硫及亚硫酸盐总量超过 10 mg/kg 或 10 mg/L（以总 SO₂ 计）时，必须在配料表中标示。",
    "(EC) 1333/2008": "E220–E228（二氧化硫及亚硫酸盐）授权用于干制水果和蔬菜，以总 SO₂ 计的最大允许量。",
    "(EC) 396/2005": "设定植物源食品农药最大残留限量；未列明具体值的农药适用默认 MRL 0.01 mg/kg。",
    "21 CFR 130.9": "可检出的亚硫酸盐指终产品中亚硫酸盐≥10 mg/kg（ppm）。",
    "40 CFR 180.507": "建立杀菌剂嘧菌酯在芒果中的残留 tolerance 为 4 ppm。",
    "C.R.C., c. 870 B.01.010.2": "若亚硫酸盐总含量≥10 mg/kg，必须在标签上标示亚硫酸盐。",
    "JECFA（apps.who.int Chemical 985）": "ADI：0–0.7 mg/kg 体重（以 SO₂ 计），为亚硫酸盐类 Group ADI。",
    "ISO 5522:1981": "方法：酸化并加热试样，以氮气流带出释放气体，吸收并氧化后，用标准氢氧化钠溶液滴定测定硫酸。",
    "WTO SPS": "SPS 协定指定粮农组织/世卫组织 Codex 食品法典为相关标准制定组织。",
}

KNOWN_PESTICIDES = set(PEST_EN.keys()) | set(PEST_EN.values())

def G(v):
    return "" if v is None else str(v)

def is_cjk(s):
    return any('\u4e00' <= c <= '\u9fff' for c in (s or ""))

def to_trad(s):
    return "".join(SIMP2TRAD.get(c, c) for c in (s or ""))

def to_simp(s):
    return "".join(TRAD2SIMP.get(c, c) for c in (s or ""))

def classify_indicator(raw):
    """返回物质级简中指标名；无法识别（非具体物质）返回 None。"""
    s = to_simp(G(raw))
    low = s.lower()
    if "二氧化硫" in s or "亞硫酸" in s or "亚硫酸" in s or "sulfit" in low:
        return "二氧化硫/SO₂残留"
    if "焦亚硫酸" in s:
        return "焦亚硫酸钠"
    if "苯甲酸" in s:
        return "苯甲酸及其钠盐"
    if "山梨酸" in s:
        return "山梨酸及其钾盐"
    if "丙酸" in s and "亚硝" not in s:
        return "丙酸及其钠盐"
    if "糖精" in s:
        return "糖精钠"
    if "甜蜜素" in s or "环己基" in s:
        return "甜蜜素(环己基氨基磺酸钠)"
    if "安赛蜜" in s or "乙酰磺胺" in s:
        return "安赛蜜(乙酰磺胺酸钾)"
    if "三氯蔗糖" in s or "蔗糖素" in s:
        return "三氯蔗糖(蔗糖素)"
    if "阿斯巴甜" in s:
        return "阿斯巴甜"
    if "胭脂红" in s:
        return "胭脂红"
    if "柠檬黄" in s:
        return "柠檬黄"
    if "日落黄" in s:
        return "日落黄"
    if "诱惑红" in s:
        return "诱惑红"
    if "亮蓝" in s:
        return "亮蓝"
    if "赤藓红" in s:
        return "赤藓红"
    if "柠檬酸" in s:
        return "柠檬酸"
    if "乳酸" in s and "亚硝" not in s:
        return "乳酸"
    if "鉛" in s or "铅" in s:
        return "铅(Pb)"
    if "鎘" in s or "镉" in s:
        return "镉(Cd)"
    if "汞" in s:
        return "汞(Hg)"
    if "砷" in s:
        return "砷(As)"
    if "鉻" in s or "铬" in s:
        return "铬(Cr)"
    if "錫" in s or "锡" in s and "亚硝" not in s:
        return "锡(Sn)"
    if "鎳" in s or "镍" in s:
        return "镍(Ni)"
    if "亚硝酸盐" in s or "亞硝酸" in s:
        return "亚硝酸盐"
    if "硝酸盐" in s or "硝酸" in s and "亚" not in s:
        return "硝酸盐"
    if "苯并" in s and "芘" in s:
        return "苯并[a]芘"
    if "二甲基亚硝胺" in s or "NDMA" in s:
        return "N-二甲基亚硝胺"
    if "多氯联苯" in s or "PCB" in s:
        return "多氯联苯"
    if "3-氯" in s or "3-MCPD" in s or "氯丙二醇" in s:
        return "3-氯-1,2-丙二醇(3-MCPD)"
    if "黃曲霉" in s or "黄曲霉" in s:
        if "M1" in s or "M-1" in s:
            return "黄曲霉毒素M1"
        return "黄曲霉毒素B1"
    if "脱氧雪腐" in s or "DON" in s:
        return "脱氧雪腐镰刀菌烯醇(DON)"
    if "展青霉素" in s or "patulin" in low:
        return "展青霉素"
    if "赭曲霉" in s or "ochratoxin" in low:
        return "赭曲霉毒素A"
    if "玉米赤霉" in s or "zearalenone" in low:
        return "玉米赤霉烯酮"
    if "農藥" in s or "农药" in s or "除害" in s or "pesticide" in low or "mrl" in low:
        return "农药残留(MRL)"
    if s in KNOWN_PESTICIDES:
        return s
    if "菌落" in s:
        return "菌落总数"
    if "大肠菌群" in s:
        return "大肠菌群"
    if "大肠杆菌" in s or "大肠埃希" in s:
        if "致泻" in s:
            return "致泻大肠埃希氏菌"
        return "大肠杆菌"
    if "沙门" in s:
        return "沙门氏菌"
    if "葡萄" in s and "球" in s:
        return "金黄色葡萄球菌"
    if "蜡样" in s or "芽胞" in s:
        return "蜡样芽胞杆菌"
    if "黴菌" in s or "霉菌" in s:
        return "霉菌"
    if "酵母" in s:
        return "酵母"
    if "李斯特" in s or "listeria" in low:
        return "单核细胞增生李斯特氏菌"
    if "副溶血" in s:
        return "副溶血性弧菌"
    if "克罗诺" in s or "cronobacter" in low:
        return "克罗诺杆菌属"
    if "志贺" in s or "shigella" in low:
        return "志贺氏菌"
    return None

def sheet3_indicator(r):
    si = classify_indicator(r.get("indicator") or "")
    if si:
        return si
    si = classify_indicator((r.get("risk_type") or "") + " " + (r.get("quote") or ""))
    if si:
        return si
    return G(r.get("risk_type")) or "[待填写]"

def ind_multiling(ind_simp):
    """指标名四语：简中/繁中/英/官方语言。优先从全量清单取英文名。"""
    entry = ALL_INDICATOR_MAP.get(ind_simp)
    if entry:
        en = entry[2]
    else:
        en = PEST_EN.get(ind_simp, "[待填写]")
    return {"simp": ind_simp, "trad": to_trad(ind_simp), "en": en, "native": "[待填写]"}

def std_multiling(std_no, std_name, region):
    en = STD_EN.get(G(std_no).strip(), "[待填写]")
    native = STD_NATIVE.get(G(std_no).strip(), {}).get(region, "[待填写]")
    return {"zh": G(std_name) or "[待填写]", "en": en, "native": native}

def claimed_std(r):
    rt = G(r.get("risk_type"))
    limit_like = any(k in rt for k in ["限量", "MAX", "MRL", "农残", "污染物", "食品添加剂", "残留"])
    if limit_like and (G(r.get("std_no")) or G(r.get("std_name"))):
        name = G(r.get("std_name")) or ""
        no = G(r.get("std_no")) or ""
        return f"{name}（{no}）" if no else name
    return "[待填写]"

def expand_rows(risks):
    """将农药 MRL 行展开为具体物质行（仅当数据含 '中文名(值)' 序列）。"""
    out = []
    for r in risks:
        code = G(r.get("code")); rt = G(r.get("risk_type")); std = G(r.get("std_no")); quote = G(r.get("quote"))
        is_pesticide = ("农残" in code) or ("2763" in std) or ("农残" in rt) or ("农药" in rt) or ("除害" in rt)
        if is_pesticide:
            seg = quote
            m = re.search(r'[：:]\s*([^。.\n]+)', quote)
            if m:
                seg = m.group(1)
            pairs = []
            for chunk in re.split(r"[、，,；;]", seg):
                chunk = chunk.strip()
                mm = re.search(r'([一-鿿·]+?)\s*[（(]?(?:及[^）)]*)?[）)]?\s*([\d.]+)\s*(mg/kg|ppm)?', chunk)
                if not mm:
                    continue
                name = mm.group(1).strip()
                if name not in KNOWN_PESTICIDES and not (2 <= len(name) <= 8 and "农药" not in name):
                    continue
                val = mm.group(2) + ((" " + mm.group(3)) if mm.group(3) else "")
                pairs.append((name, val.strip()))
            valid = [(n, v) for n, v in pairs if n in KNOWN_PESTICIDES or (2 <= len(n) <= 8)]
            if valid:
                for name, val in valid:
                    nr = dict(r)
                    nr["indicator"] = name
                    nr["limit"] = val
                    nr["method"] = nr.get("method") or "GB 23200.113-2018（植物源性食品农药多残留测定 LC-MS/MS）"
                    nr["note"] = (G(nr.get("note")) + f"；由农药 MRL 行展开：{name}").strip("；")
                    out.append(nr)
                continue
        r = dict(r)
        r["indicator"] = sheet3_indicator(r)
        out.append(r)
    return out

def build_workbook(out_dir, food_name, basic, translations, risks, merge_mode=False):
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(os.path.join(out_dir, "standards"), exist_ok=True)
    out_path = os.path.join(out_dir, f"{food_name}_食品安全风险识别表.xlsx")
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="BBBBBB"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="0F766E"); head_font = Font(bold=True, color="FFFFFF")
    sub_fill = PatternFill("solid", fgColor="CCFBF1"); sub_font = Font(bold=True, color="134E4A")
    cat_fill = PatternFill("solid", fgColor="F0F9FF"); cat_font = Font(bold=True, color="1E3A5F")
    gap_fill = PatternFill("solid", fgColor="FEF3C7")
    wrap = Alignment(vertical="top", wrap_text=True)

    # ===== Sheet1 基础信息（列表 + 配料多语对照） =====
    ws = wb.active; ws.title = "基础信息"
    ws.append(["字段", "内容"])
    ws.cell(1, 1).fill = head_fill; ws.cell(1, 1).font = head_font
    ws.cell(1, 2).fill = head_fill; ws.cell(1, 2).font = head_font
    base_rows = [
        ("食品中文名", basic.get("name_cn")), ("国际/英文名", basic.get("name_en") or basic.get("aliases")),
        ("别名/同义词", basic.get("aliases")), ("类别", basic.get("category")),
        ("配料", basic.get("ingredients_cn")),
        ("食品添加剂(含代码)", basic.get("additives")), ("添加剂说明", basic.get("additive_note")),
        ("过敏原", basic.get("allergens")), ("翻译置信度提示", basic.get("trans_conf")),
        ("skill 版本", SKILL_VER), ("生成日期", basic.get("gen_date")),
    ]
    for k, v in base_rows:
        ws.append([k, G(v)])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=2):
        row[0].font = Font(bold=True); row[0].border = border; row[0].alignment = wrap
        row[1].border = border; row[1].alignment = wrap
    ws.append([])
    ws.append(["配料多语对照（下列字段均按 14 种语言列出，来源 translations.json）"])
    hr = ws.max_row
    ws.cell(hr, 1).fill = sub_fill; ws.cell(hr, 1).font = sub_font
    langs = ["简中", "繁中", "英", "泰", "高棉", "越", "阿", "法", "德", "俄", "西", "葡", "日", "韩"]
    ws.append(["字段"] + langs)
    for c in range(1, len(langs) + 2):
        ws.cell(ws.max_row, c).fill = head_fill; ws.cell(ws.max_row, c).font = head_font
    tmap = {t.get("field"): t for t in (translations or [])}
    ing_fields = [f for f in tmap if f.startswith("配料")]
    for field in ["名称", "类别"] + ing_fields:
        t = tmap.get(field, {})
        ws.append([field] + [G(t.get(l)) for l in langs])
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row, max_col=len(langs) + 1):
        for cell in row:
            cell.border = border; cell.alignment = wrap
    ws.column_dimensions["A"].width = 26
    for i in range(2, len(langs) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.column_dimensions["B"].width = 40

    # ===== Sheet3 食品安全风险识别表（全量指标框架，31 列） =====
    ws3 = wb.create_sheet("食品安全风险识别表")
    cols = [
        "序号", "国家/地区(组织)代码", "中文名称", "英文名称", "本国语言名称",
        "法规名称(简中全称)", "法规名称(英全称)", "法规名称(官方语言全称)", "标准编号",
        "风险类型",
        "指标名称(简中)", "指标名称(繁中)", "指标名称(英)", "指标名称(官方语言)",
        "关联代码(INS/E/CI/CAS)", "适用对象/范围", "原文引用(本国语言)", "原文中文翻译",
        "限量值", "检测方法(标准编号/名称)", "限制条件", "声称执行标准",
        "来源文件名称", "来源网址", "网页存档路径", "命中别名/检索轨迹",
        "置信度(高/中/低)", "发布日期", "有效性", "访问日期", "备注",
    ]
    LIMIT_COL = cols.index("限量值") + 1
    ws3.append(cols)
    for c in range(1, len(cols) + 1):
        ws3.cell(1, c).fill = head_fill; ws3.cell(1, c).font = head_font

    # 去重 + 展开
    seen = set(); deduped = []
    for r in risks:
        key = (G(r.get("region")), G(r.get("std_no")), G(r.get("risk_type")), G(r.get("indicator")))
        if key in seen:
            continue
        seen.add(key); deduped.append(r)
    risks_expanded = expand_rows(deduped)

    # 按全量指标归类：指标简中名 -> [risk行列表]
    risk_by_indicator = {}
    for r in risks_expanded:
        ind = sheet3_indicator(r)
        risk_by_indicator.setdefault(ind, []).append(r)

    # 遍历全量指标清单，生成 Sheet3 行
    source_map = {}  # (指标简中名, 地区代码) -> Sheet3 行号
    r3 = 1
    seq = 0
    current_cat = None
    for cat, ind_simp, ind_en, ind_code, ind_std in ALL_INDICATORS:
        # 类别分隔行
        if cat != current_cat:
            current_cat = cat
            r3 += 1
            ws3.cell(r3, 1, f"━━ {cat} ━━")
            for c in range(1, len(cols) + 1):
                ws3.cell(r3, c).fill = cat_fill; ws3.cell(r3, c).font = cat_font

        # 找该指标的风险数据行
        matched = risk_by_indicator.get(ind_simp, [])
        if matched:
            for r in matched:
                seq += 1
                r3 += 1
                reg = G(r.get("region"))
                zh, en, nat = REGION_INFO.get(reg, (reg, reg, reg))
                quote = G(r.get("quote"))
                qcn = G(r.get("quote_cn")) or QUOTE_CN.get(G(r.get("std_no")).strip(),
                                                           "（原文为中文）" if is_cjk(quote) else "待补充")
                im = ind_multiling(ind_simp)
                sm = std_multiling(r.get("std_no"), r.get("std_name"), reg)
                claimed = claimed_std(r)
                ws3.append([
                    seq, reg, zh, en, nat,
                    sm["zh"], sm["en"], sm["native"], G(r.get("std_no")),
                    G(r.get("risk_type")),
                    im["simp"], im["trad"], im["en"], im["native"],
                    G(r.get("code")) or ind_code, G(r.get("scope")), quote, qcn,
                    G(r.get("limit")), G(r.get("method")), G(r.get("restriction")), claimed,
                    G(r.get("file")), G(r.get("url")), G(r.get("page_path")), G(r.get("hit_alias")),
                    G(r.get("confidence")), G(r.get("published")), G(r.get("validity")),
                    G(r.get("access")), G(r.get("note")),
                ])
                if G(r.get("limit")) not in ("", "[待填写]", "待补充"):
                    source_map[(ind_simp, reg)] = r3
                else:
                    # 有数据行但限量值为空，仍链接到该行（显示空值，区别于"无数据"）
                    source_map.setdefault((ind_simp, reg), r3)
        else:
            # 无数据框架行：以 CN 为框架，限量值标 [待填写]
            seq += 1
            r3 += 1
            reg = "CN"
            zh, en, nat = REGION_INFO.get(reg, (reg, reg, reg))
            im = ind_multiling(ind_simp)
            sm = std_multiling(ind_std, "", reg)
            ws3.append([
                seq, reg, zh, en, nat,
                "[待填写]", sm["en"], "[待填写]", ind_std,
                cat,
                im["simp"], im["trad"], im["en"], im["native"],
                ind_code, food_name, "", "待补充",
                "[待填写]", "[待填写]", "适用于该食品类别，本次检索未命中具体限量值", "[待填写]",
                "", "", "", "",
                "低", "", "", "", "全量指标清单框架行，待检索补充",
            ])
            # 框架行也加入 source_map（标为框架，Sheet4 链接到此行显示 [待填写]）
            if (ind_simp, reg) not in source_map:
                source_map[(ind_simp, reg)] = r3

    widths = [6, 16, 12, 14, 16, 30, 30, 26, 16, 16, 16, 16, 20, 18, 18, 22, 46, 40, 16, 30, 22, 26, 18, 38, 28, 20, 10, 14, 12, 14, 24]
    for i, w in enumerate(widths, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row):
        for cell in row:
            cell.border = border; cell.alignment = wrap
    ws3.freeze_panes = "C2"; ws3.auto_filter.ref = ws3.dimensions

    # ===== Sheet4 指标对比查询（全量指标 × 全量地区 完整矩阵） =====
    ws4 = wb.create_sheet("指标对比查询")
    # 全量地区列头
    regions_full = [REGION_FULLNAME.get(reg, reg) for reg in ALL_REGIONS]

    ws4.append(["指标名称 \\ 国家/地区(组织)"] + regions_full)
    for c in range(1, len(regions_full) + 2):
        ws4.cell(1, c).fill = head_fill; ws4.cell(1, c).font = head_font
    ws4.cell(1, 1).alignment = Alignment(vertical="center", wrap_text=True)

    current_cat4 = None
    for cat, ind_simp, ind_en, ind_code, ind_std in ALL_INDICATORS:
        # 类别分隔行
        if cat != current_cat4:
            current_cat4 = cat
            row = [f"━━ {cat} ━━"] + [""] * len(ALL_REGIONS)
            ws4.append(row)
            for c in range(1, len(ALL_REGIONS) + 2):
                ws4.cell(ws4.max_row, c).fill = cat_fill; ws4.cell(ws4.max_row, c).font = cat_font

        row = [ind_simp]
        for reg in ALL_REGIONS:
            src_row = source_map.get((ind_simp, reg))
            if src_row:
                cell_ref = f"{get_column_letter(LIMIT_COL)}{src_row}"
                sheet_ref = f"'食品安全风险识别表'!{cell_ref}"
                row.append(f'=HYPERLINK("#{sheet_ref}", {sheet_ref})')
            else:
                row.append("[待填写]")
        ws4.append(row)

    for row in ws4.iter_rows(min_row=1, max_row=ws4.max_row):
        for cell in row:
            cell.border = border; cell.alignment = wrap
    ws4.column_dimensions["A"].width = 26
    for i in range(2, len(regions_full) + 2):
        ws4.column_dimensions[get_column_letter(i)].width = 22
    ws4.freeze_panes = "B2"; ws4.auto_filter.ref = ws4.dimensions

    # 信息缺口声明
    gap_note_row = ws4.max_row + 2
    ws4.cell(gap_note_row, 1, "信息缺口声明")
    ws4.cell(gap_note_row, 1).font = sub_font; ws4.cell(gap_note_row, 1).fill = sub_fill
    notes = [
        f"1. 本表为全量矩阵：指标 {len(ALL_INDICATORS)} 项 × 地区 {len(ALL_REGIONS)} 个 = {len(ALL_INDICATORS)*len(ALL_REGIONS)} 个单元格。有数据的单元格可点击跳转源表，无数据的标 [待填写]。",
        "2. 微生物/致病菌、真菌毒素指标为本表强制保留类别；部分辖区检索未覆盖，对应单元格置 [待填写]，需后续补充。",
        "3. 指标名称(官方语言)、法规名称(官方语言全称) 因检索未返回各辖区官方语言原文，暂置 [待填写]，需检索补充。",
        "4. 部分辖区仅给出框架性农药/污染物规定（无逐项具体限量值），其透视单元格亦为 [待填写]。",
        "5. 污染物中锡(Sn)主要针对罐装/金属容器食品；多氯联苯主要针对油脂类食品；对本食品可能不适用，但仍全量列出以供参考。",
    ]
    for j, n in enumerate(notes, 1):
        ws4.cell(gap_note_row + j, 1, n)
        ws4.cell(gap_note_row + j, 1).alignment = wrap

    wb.save(out_path)
    return out_path


def main(argv):
    if len(argv) < 2:
        print("用法: python build_xlsx.py <食品目录> [<食品中文名>]")
        return 2
    food_dir = argv[1]
    if not os.path.isdir(food_dir):
        print(f"错误：目录不存在 {food_dir}")
        return 2
    basic_path = os.path.join(food_dir, "basic.json")
    trans_path = os.path.join(food_dir, "translations.json")
    risks_path = os.path.join(food_dir, "risks.json")
    if not os.path.exists(basic_path):
        print(f"错误：缺少 {basic_path}")
        return 2
    basic = json.load(open(basic_path, encoding="utf-8"))
    translations = json.load(open(trans_path, encoding="utf-8")) if os.path.exists(trans_path) else []
    risks = json.load(open(risks_path, encoding="utf-8")) if os.path.exists(risks_path) else []
    food_name = argv[2] if len(argv) > 2 else G(basic.get("name_cn")) or "食品"
    out = build_workbook(food_dir, food_name, basic, translations, risks)
    print(f"SAVED: {out}")
    print(f"风险行(去重+展开后): {len(risks)}")
    print(f"全量指标: {len(ALL_INDICATORS)} 项 × 全量地区: {len(ALL_REGIONS)} 个")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
