---
name: foodsafeml
description: >-
  技能中文名「食安多语风险图鉴」，英文名 FoodSafeML。给定食品名，引导录类别+配料，翻译14种语言，
  扩展别名检索，全球全量检索各国/地区+ISO+WHO(经JECFA)+FAO(Codex)+各联盟+区域论坛的食品安全要求
  （执行标准/限量/检测方法/限制条件），识别添加剂代码(INS/E/CI/CAS)并关联，原文引用、按地区分类、
  附网址/网页存档/发布日期/有效性，输出多语 Excel《食品安全风险识别表》。
  工作簿结构(v1.7.0)：①基础信息(含配料多语对照) ②食品安全风险识别表(31列，含指标名称四语/法规名称三语/「声称执行标准」)
  ③指标对比查询(公式透视，地区列用官方全称，点击单元格可跳转源表)。农残/污染物逐物质展开，强制含微生物/致病菌指标。
  触发词：食品多语言安全风险识别、食品安全风险识别表、录入食品、食品合规、food safety risk、全球食品标准。
version: 1.7.0
author: user-request
---

# 食品多语言安全风险识别与信息整合

## 一、用途与触发

用户给出**任意食品名称**，希望产出可用于合规评估的 **Excel《食品安全风险识别表》** 并**留存标准原文与网页证据**时启用。完整闭环（Phase 0–5 见第三节）。

- **默认增强档检索，无需指定具体目标市场**；引入「检索深度档位」控制规模（见 Phase 0）。全量档需显式选择。
- 凡有适用标准的辖区或国际组织均应收录，按 ISO 3166 国家代码或组织代码分组。

## 二、目标语言清单（14 种，可裁剪）

| 代码 | 语言 | 对应主要市场 | 翻译置信度提示 |
|------|------|--------------|----------------|
| zh-CN | 简体中文 | 中国内地（基准） | 高 |
| zh-TW | 繁体中文 | 港澳台 | 高 |
| en | English | 国际通用 | 高 |
| th | 泰语 | 泰国 | 中（建议核验） |
| km | 柬埔寨语(高棉语) | 柬埔寨 | **低（需核验）** |
| vi | 越南语 | 越南 | 中 |
| ar | 阿拉伯语 | 中东/GCC | **低（需核验）** |
| fr | 法语 | 法国/欧盟 | 高 |
| de | 德语 | 德国/欧盟 | 高 |
| ru | 俄语 | 俄罗斯/EAEU | 中 |
| es | 西班牙语 | 西班牙/拉美 | 高 |
| pt | 葡萄牙语 | 葡萄牙/巴西 | 高 |
| ja | 日语 | 日本 | 高 |
| ko | 韩语 | 韩国 | 高 |

> 标注「低」的语言，翻译结果须在该食品行/备注中标「翻译置信度低，待专业核验」，不得当作已确认标签。

## 三、交互流程（Phase）

### Phase 0 · 输入判别与范围设定（新增，必做）
- **输入判别**：若用户输入命中「附录 A 类别」却是**品类而非具体产品**（如「饮料」「零食」），须用 AskUserQuestion 确认：①提供具体产品名；或 ②按该品类做扫描（明确标注「品类级扫描，非单品」）。
- **检索深度档位**（AskUserQuestion 选择，默认「增强」）：
  - **标准档**：核心辖区（CN/US/EU/Codex/ISO + 用户曾提市场）+ 该食品直接相关的 4 类；每辖区 ≤2 次查询。
  - **增强档（默认）**：主要辖区 + 主要联盟/区域论坛；每辖区 ≤4 次查询。
  - **全量档**：全部辖区 + 全部组织；每辖区 ≤6 次查询，建议启用并行（见 Phase 4 并行模式）。
- 初始化 **「已覆盖辖区」集合** 与 **查询计数器**，用于去重与上限控制；并初始化本档位的**目标辖区清单总数**（作为进度播报的 Y 值：标准档≈12、增强档≈30、全量档≈全部 ISO 3166 辖区+组织数）。
- **输出目录**：默认 `<workspace>/food-safety-<食品名>/`（如 `food-safety-螺蛳粉/`）；可用 AskUserQuestion 让用户指定其他路径。该目录即生成器 `out_dir`，其下自动生成 `<食品名>_食品安全风险识别表.xlsx` 与 `standards/`。

### Phase 1 · 基础信息录入引导
- 接收一个或多个食品名称（支持批量）。
- **类别归属**：用 AskUserQuestion 在「附录 A」确认/选择；用户已给则回填。
  - **批量跳过模式**：若一次多个食品，先问一次「逐条确认 / 全部自动分类(跳过提问，按名称推断类别)」，选后者则不逐条弹问。
- **配料/成分补充**：提示填写完整配料表（可粘贴标签原文）。结构化字段：主料/辅料/食品添加剂（INS 编号或功能类别）/已知过敏原。
- 用户不知道时 WebSearch 检索「典型配料」作建议，但必须标「建议值，需用户确认」，**不得**当事实写入。
- 未提供且无法确认字段统一写 `[待补充]`，**禁止编造**。

### Phase 2 · 自动多语言翻译
- 对「食品名称/类别/配料(逐条)」翻译到目标语言清单。
- 原则：食品名称用常见商品名/通用名；配料用标签标准译法（添加剂优先 INS 编号+标准名）；各语言与中文一一对应，缺失标 `[待翻译]`。
- 模型直接产出；对关键市场（日/韩/泰/阿等）建议联网核验官方标签术语并标核验来源。
- **低资源语言**（km/ar 等，见第二节提示）翻译结果须标「翻译置信度低，待核验」。

### Phase 3 · 名称/别名/同义词扩展（提升召回）
- 基于用户名称，生成**名称集合**：常用中文别名、方言/regional 叫法、英文通用名/商品名、拉丁学名（如适用）、近似/相关产品名。
- **按语言映射**：将别名尽量映射到对应语言（英文别名用于 EN 检索、泰语别名用于 TH 检索…），提升跨语言命中。
- 例：「螺蛳粉」→ 柳州螺蛳粉 / snail rice noodle / Liuzhou river snail rice noodle / river snail noodles。
- 该集合在 Phase 4 全部检索中复用。

### Phase 4 · 全球权威信息检索（全量，核心）
- **范围=全球全量**（见附录 B），按档位控制规模（Phase 0）；按 ISO 3166 或组织代码分组。
- 调用 **WebSearch + WebFetch**，对 Phase 3 名称集合检索；严格遵守**每辖区查询上限**。
- **去重键（关键）**：行级去重键 = `(辖区, 标准编号, 风险类型)`，**不得仅按辖区去重**——同一辖区（如 CN）内的 GB 2760/GB 2762/GB 2763 属不同标准，必须分别成行。「已覆盖辖区」集合仅用于进度播报（X/Y 辖区），不参与行级去重。
- **检索工具全失败降级**：若 WebSearch 与 WebFetch 均不可用/失败，不得生成空表，应产出「人工核验清单」（列出本应检索的辖区+风险类型+建议官方入口），并在 Excel 备注与对话中显式提示「检索工具不可用，待人工核验」。
- **进度落盘**：全量档/大批量时，除每约 5 辖区在对话回传摘要外，须将进度与已命中条目持续写入 `<输出目录>/<食品名>/progress.md`，避免上下文超长导致中间结果丢失。
- **分批回报**：每完成若干辖区（如每 5 个或达上限比例），向用户回传一次「已覆盖 X/Y 辖区，命中 N 条」进度摘要，避免长任务黑盒。
- 每检索目标提取四类风险项：①执行标准 ②相关限量 ③检测方法 ④限制条件（同 v1.2）。
- **添加剂代码关联**（重点，含同号判定）：
  - 识别 INS（如 INS 102 柠檬黄）、E 编号（欧盟 E102）、CI 号（色素，如 CI 19140）、CAS 号；
  - 对每个代码，检索其在 **Codex GSFA / JECFA / EU 附录 / GB 2760** 的使用范围与限量；
  - **同号判定**：标注「该地区是否采用同一编号 / 采用不同编号(写明) / 无对应编号(如香料香精常无 INS) / CI 仅适用于色素」；
  - 作为「添加剂代码关联」风险行，「关联代码」列回填该代码。
- **原文引用与置信度分级**：每条摘录官方原文关键句，并打**置信度**：
  - **高**=官网全文可引（如 eCFR/EUR-Lex/标准公开库）；
  - **中**=官方摘要页或二手官方页；
  - **低**=仅见引用、未获原文，需核验。
  - 附：来源文件名称、官方网址、发布/修订日期、有效性状态（现行有效/已废止/即将实施/修订中/**未知**）、本次访问日期。
- **文件与网页存档（含降级，重点）**：
  - 公开可获取 PDF → 下载到 `<输出目录>/<食品名>/standards/`，命名 `<代码>_<标准号>_<风险类型缩写>.pdf`（风险类型缩写：STD=执行标准 / MAX=相关限量 / MET=检测方法 / RES=限制条件 / ADD=添加剂代码关联，避免同标准跨类型文件名碰撞覆盖）；
  - **网页存档**：WebFetch 返回的是**模型摘要**，另存为 `<代码>_<标准号>_<风险类型缩写>_page.md`；若沙箱允许用 `curl -L` 抓取**原始 HTML**，存 `<代码>_<标准号>_<风险类型缩写>_page.html`（**注意：JS 渲染页面抓不全，亦可能被墙/拦截**）；路径记入 Excel「网页存档路径」；
  - **下载/抓取失败降级**：仅保留 URL + WebFetch 摘要（`.md`），「备注」标「下载受限/抓取受限」，**不得伪造文件**。
- **并行模式（可选，呼应大规模检索）**：当档位=全量或辖区数>20，将区域分簇（如 东亚/欧盟/美洲/中东非/大洋洲），用多个 sub-agent（Agent 工具）并行检索各簇。
  - **sub-agent 上下文内联（必须）**：sub-agent 是独立上下文，不会继承本 skill；派发时须在提示词中**完整内联**附录 B 源清单、Sheet3 列结构（31 列）、置信度定义、去重键规则与文件命名规则，否则检索会脱靶。
  - 各 sub-agent 回传其簇的风险行（含来源/置信度/去重键）；主 agent 做**跨簇去重（按去重键）+ 冲突仲裁 + 合并**，再进入 Phase 5。

### Phase 5 · 生成 Excel + 文件归档
- 使用 **Python(openpyxl)** 生成 `.xlsx`（受管 venv `pip install openpyxl`；受限可改用 `minimax-xlsx` skill）。生成器见附录八，已含空值保护与 `standards/` 目录自动创建。
- **输出目录结构**（建议）：`<输出目录>/<食品名>/{<食品名>_食品安全风险识别表.xlsx, standards/}`；批量时每食品一个子目录，或与用户确认合并为单工作簿（合并时 Sheet3 含「食品名称」列区分）。
- 对话中以 Markdown 表格预览关键结果，并列出已存档文件清单（含下载受限项）。

## 四、质量与合规红线
- **不编造**标准编号、限量值、检测方法或代码含义；检索无果标「未检索到官方依据 / 待补充」。
- **原文引用**优先于转述；转述注「译自原文」。
- **置信度如实标注**（高/中/低）；低置信项须在结论中显式提示需核验。
- **存档诚实**：WebFetch 摘要 ≠ 原网页；JS 页/被墙页抓不全；下载失败不伪造文件。
- 翻译与合规结论均标 **「仅供参考，正式标签/合规以目标国主管机构及国际组织最新发布为准」**。
- 柬埔寨等本地源稀缺时回退 Codex/ASEAN 并标低置信。

## 五、附录 A · 食品类别分类法（供 Phase 1 勾选）
饮料类 / 零食类(膨化·坚果炒货) / 饼干糕点类 / 糖果巧克力类 / 方便食品(方便面·速食) /
调味品(酱料·油·盐·味精) / 乳制品 / 粮油谷物 / 肉蛋水产制品 / 冷冻冷藏食品 / 果蔬制品(蜜饯·罐头) /
保健食品 / 婴幼儿食品 / 其他（用户自定义）。

## 六、附录 B · 权威官方源清单（应扩至全球全量）
**国家/地区（节选主要，凡有适用标准者均收录，按 ISO 3166 代码）**
| 代码 | 主管机构 / 入口 | 关键法规或标准 |
|------|----------------|----------------|
| CN | CFSA `https://www.cfsa.net.cn`；卫健委 `https://www.nhc.gov.cn`；标准公开 `https://std.samr.gov.cn` | GB 2760/2761/2762/2763、GB 5009 |
| HK | 食物安全中心 CFS `https://www.cfs.gov.hk` | 食物搀假/掺杂规例、除害剂（农药）残留规例、金属杂质（铅等）限量、亚硫酸盐于干果的指引/限量、营养标签 |
| MO | 市政署 IAM `https://www.iam.gov.mo` | 食品添加剂/农药残留/重金属污染物/二氧化硫规定（澳门多对齐 Codex/欧盟） |
| US | FDA `https://www.fda.gov/food`；eCFR `https://www.ecfr.gov`；FSIS `https://www.fsis.usda.gov`；EPA `https://www.epa.gov` | 21 CFR 100-199、EPA tolerance |
| EU | EUR-Lex `https://eur-lex.europa.eu`；EFSA `https://www.efsa.europa.eu` | Reg 1881/2006、1333/2008、396/2005、EU 1169/2011 |
| UK | FSA `https://www.food.gov.uk` | 脱欧后自主法规 |
| CA | Health Canada `https://www.canada.ca`；CFIA `https://inspection.canada.ca` | FDR |
| AU/NZ | FSANZ `https://www.foodstandards.gov.au` | Food Standards Code |
| JP | MHLW `https://www.mhlw.go.jp` | 食品卫生法、肯定列表、添加物公定書 |
| KR | MFDS `https://www.mfds.go.kr` | Food Code |
| TH | Thai FDA `https://www.fda.moph.go.th` | Food Act |
| VN | MOH `https://moh.gov.vn` | TCVN |
| KH | 卫生部/商务部公告（源少，回退 Codex/ASEAN） | — |
| ASEAN | `https://www.asean.org` | ASEAN Harmonization |
| GCC | SFDA `https://sfda.gov.sa`；GSO `https://www.gso.org.sa` | GCC 标准 |
| RU/EAEU | Rospotrebnadzor `https://www.rospotrebnadzor.ru` | TR TS 021/2011 |
| FR/DE/ES/PT | 各国主管 + 欧盟法规 | 同 EU |
| BR | ANVISA `https://www.gov.br/anvisa` | RDC |
| MX | COFEPRIS `https://www.gob.mx/cofepris` | 卫生法规 |
| IN | FSSAI `https://www.fssai.gov.in` | FSS Regs |
| ZA | NHC/DALRRD | R.965 |
| … | 其余所有有食品标准的辖区依 ISO 3166 自行补充 | — |

**国际组织 / 联盟 / 区域论坛**
| 组织 | 入口 | 性质与作用（准确性修正） |
|------|------|--------------------------|
| **ISO** | `https://www.iso.org` | 检测方法/管理体系国际标准（直接发布） |
| **FAO / Codex** | `https://www.fao.org/fao-who-codexalimentarius`；GSFA `https://www.fao.org/gsfaonline` | 直接发布 CXS 通用标准、GSFA 添加剂用量 |
| **WHO（经 JECFA）** | `https://www.who.int` | **不直接发布强制标准**；通过 JECFA 发布添加剂/污染物风险评估与 ADI、健康指南 |
| **JECFA** | WHO/FAO 联合 | 添加剂/污染物评估、ADI（评估机构，非标准强制方） |
| **MERCOSUR** | 南方共同市场 | 南美共同标准（直接） |
| **CARICOM** | 加勒比共同体 | 区域标准（直接） |
| **SPC / PIF** | 太平洋共同体 | 太平洋岛国论坛标准（直接） |
| **AU / ARSO** | 非盟/非洲标准组织 | 非洲区域标准（直接） |
| **WTO / SPS** | `https://www.wto.org` | **引用框架**：SPS 协定指向 Codex/OIE/IPPC，非直接标准发布方 |
| **APEC** | `https://www.apec.org` | **协调论坛**：促进标准一致，非直接标准发布方 |

> 检索组合：`site:官方域名 + (名称集合任一) + standard|maximum level|limit|detection method|additive`。名单未覆盖但有适用标准的辖区/组织，按代码自行补充并标注来源。

## 七、附录 C · Excel 结构（3 个工作表）

> **v1.7.0 结构变更**：① 取消独立的「多语翻译对照」Sheet，多语字段直接嵌入业务表（配料多语进 Sheet1、指标/法规多语进 Sheet3）；② 指标对比查询升级为**公式透视**（地区列用官方全称、点击单元格可跳转源表）；③ 农残/污染物逐物质展开；④ 强制保留微生物/致病菌指标；⑤ 指标名四语、法规名三语；⑥ 新增「声称执行标准」。

**Sheet1「基础信息」**（列表呈现 + 配料多语对照子表）
| 字段 | 说明 |
|------|------|
| 食品中文名 | |
| 国际/英文名 | 含别名/拉丁名 |
| 别名/同义词 | Phase3 扩展集合 |
| 类别 | Phase1 确认值 |
| 食品添加剂(含代码) | INS/E/CI/CAS |
| 添加剂说明 | 工艺/残留背景 |
| 过敏原 | |
| 翻译置信度提示 | 低资源语言标注 |
| skill 版本 | 1.7.0（可追溯） |
| 生成日期 | |

**配料多语对照（子表，14 语列）**：字段行 = 名称 / 类别 / 配料(逐条)；列 = 简中/繁中/英/泰/高棉/越/阿/法/德/俄/西/葡/日/韩。所有配料同步给出多语名称（需求 #5）。

**Sheet3「食品安全风险识别表」**（核心，单品 31 列 / 合并 32 列）
> 单品模式无「食品名称」列；合并多食品时在该列位置插回。

| 列 | 说明 |
|----|------|
| 序号 | |
| 国家/地区(组织)代码 | CN/HK/MO/TW/JP/KR/US/…/EU/CODEX/JECFA/ISO/APEC… |
| 中文名称 | 辖区中文名 |
| 英文名称 | 辖区英文名 |
| 本国语言名称 | 辖区本国语言名 |
| **法规名称(简中全称)** | 如 食品安全国家标准 食品添加剂使用标准 |
| **法规名称(英全称)** | 如 National Food Safety Standard – … |
| **法规名称(官方语言全称)** | 辖区官方语言全称；无原文标 `[待填写]` |
| 标准编号 | |
| 风险类型 | 执行标准/相关限量/检测方法/限制条件/添加剂代码关联/引用框架/协调论坛 |
| **指标名称(简中)** | 具体物质：二氧化硫/SO₂残留 / 铅(Pb) / 黄曲霉毒素 / 咪鲜胺 … |
| **指标名称(繁中)** | 简中→繁体转换 |
| **指标名称(英)** | 词典映射（铅→Lead(Pb) 等） |
| **指标名称(官方语言)** | 辖区官方语言；无原文标 `[待填写]` |
| 关联代码(INS/E/CI/CAS) | |
| 适用对象/范围 | |
| 原文引用(本国语言) | 官方原文；无标「待补充」 |
| 原文中文翻译 | 原文已中文标「（原文为中文）」 |
| 限量值 | 含单位 |
| 检测方法(标准编号/名称) | 具体编号 |
| 限制条件 | |
| **声称执行标准** | 该产品声称符合的海外强制/自愿性标准（替代国内企业执行标准逻辑）；无标 `[待填写]` |
| 来源文件名称 | |
| 来源网址 | |
| 网页存档路径 | |
| 命中别名/检索轨迹 | |
| 置信度(高/中/低) | |
| 发布日期 | |
| 有效性 | |
| 访问日期 | |
| 备注 | |

Sheet3 启用筛选、冻结首行、表头着色、列宽自适应。

**农残/污染物/微生物展开规则（需求 #3）**
- **农残**：检索到具体物质（如 CN GB 2763 列明咪鲜胺 2、啶虫脒 2…）须**逐物质拆成独立行**（指标名称=该农药具体名）；仅获框架性规定（如「农药 MRL 肯定列表」）且未拿到具体值时，保留一行、指标名称写「农药残留(MRL)」并标「待补充」，**不得编造**。
- **污染物**：取消笼统「污染物(铅、镉等)」，透视/对比中**逐一展开**为 铅(Pb)/镉(Cd)/汞(Hg)/砷(As) 等独立项目；有具体值则成行，无值则透视单元格标 `[待填写]`。
- **微生物/致病菌**：**强制保留** 菌落总数/大肠菌群/大肠杆菌/沙门氏菌/金黄色葡萄球菌/蜡样芽胞杆菌/霉菌/酵母 等类别，不得遗漏；本样本无检索数据则透视单元格标 `[待填写]` 并在 Sheet4 末附「信息缺口声明」。

**Sheet4「指标对比查询」**（公式透视表，v1.7.0 升级）
- 行 = **指标名称**（二氧化硫/SO₂残留、铅(Pb)、镉(Cd)、汞(Hg)、砷(As)、黄曲霉毒素、各具体农药、微生物/致病菌…）；
- 列 = **国家/地区(组织)**，列头用**官方全称**（如 AU/NZ→澳大利亚/新西兰食品标准局（FSANZ）；APEC→亚太经济合作组织（APEC）），仅纳入有数据的辖区；
- 交叉单元格 = `=HYPERLINK("#'食品安全风险识别表'!S{行}", '食品安全风险识别表'!S{行})`：**所有数据通过公式从 Sheet3 源表调取，点击单元格可直接跳转至源表对应原单元格**；无数据标 `[待填写]`；
- 末附「信息缺口声明」（微生物/官方语言原文缺失等）。启用筛选、冻结首行首列、表头着色。

## 八、Python 生成器参考实现（agent 按数据填充后调用）

> v1.7.0 参考实现：单品模式不输出「食品名称」列（合并多食品时 `merge_mode=True` 在序号后插入）。多语字段嵌入业务表——`REGION_INFO`/`REGION_FULLNAME` 提供辖区中/英/本国语言名与官方全称；`SIMP2TRAD`/`TRAD2SIMP` 做简繁转换；`IND_EN`/`PEST_EN`/`STD_EN` 提供指标与法规多语名称；`classify_indicator()`/`expand_rows()` 做物质级展开；`Sheet4` 为 `HYPERLINK` 公式透视（地区列用官方全称、点击跳转源表）。缺失多语/具体物质一律 `[待填写]`，不编造。

```python
import os, re, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

# 地区代码 -> 官方全称（用于 Sheet4 列头，需求 #2）
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

# 简<->繁 转换表（覆盖食品安全领域常用字）
PAIRS = [
    ("铅","鉛"),("镉","鎘"),("亚","亞"),("盐","鹽"),("农","農"),("药","藥"),
    ("残","殘"),("黄","黃"),("霉","黴"),("标","標"),("准","準"),("规","規"),
    ("剂","劑"),("检","檢"),("验","驗"),("签","籤"),("过","過"),("敏","敏"),
    ("总","總"),("数","數"),("肠","腸"),("杆","桿"),("门","門"),("干","乾"),
    ("类","類"),("测","測"),("氯","氯"),("氰","氰"),("菊","菊"),("酯","酯"),
    ("醚","醚"),("唑","唑"),("啉","啉"),("胺","胺"),("脒","脒"),("虫","蟲"),
    ("噻","噻"),("嗪","嗪"),("威","威"),("鲜","鮮"),("锰","錳"),("吡","吡"),
    ("嘧","嘧"),("灵","靈"),("苯","苯"),("环","環"),("啶","啶"),("磺","磺"),
    ("钠","鈉"),("钾","鉀"),("钙","鈣"),("镁","鎂"),("铁","鐵"),("铜","銅"),("锌","鋅"),
]
SIMP2TRAD = {s: t for s, t in PAIRS}
TRAD2SIMP = {t: s for s, t in SIMP2TRAD.items()}

def to_trad(s):
    return "".join(SIMP2TRAD.get(c, c) for c in (s or ""))

def to_simp(s):
    return "".join(TRAD2SIMP.get(c, c) for c in (s or ""))

IND_EN = {
    "二氧化硫/SO₂残留": "Sulfur dioxide / Sulphite residues (SO₂)",
    "铅(Pb)": "Lead (Pb)", "镉(Cd)": "Cadmium (Cd)", "汞(Hg)": "Mercury (Hg)",
    "砷(As)": "Arsenic (As)", "黄曲霉毒素": "Aflatoxins", "农药残留(MRL)": "Pesticide residues (MRL)",
    "菌落总数": "Aerobic plate count (APC)", "大肠菌群": "Coliforms", "大肠杆菌": "Escherichia coli (E. coli)",
    "沙门氏菌": "Salmonella", "金黄色葡萄球菌": "Staphylococcus aureus", "蜡样芽胞杆菌": "Bacillus cereus",
    "霉菌": "Moulds", "酵母": "Yeasts",
}
IND_NATIVE = {
    "二氧化硫/SO₂残留": {"EU": "Dioxyde de soufre / sulfites", "US": "Sulfur dioxide / sulfites", "JP": "二酸化硫黄／亜硫酸塩", "KR": "아황산가스/아황산염"},
    "铅(Pb)": {"EU": "Plomb (Pb)", "US": "Lead (Pb)", "JP": "鉛（Pb）", "KR": "납(Pb)"},
    "黄曲霉毒素": {"EU": "Aflatoxines", "US": "Aflatoxins", "JP": "アフラトキシン", "KR": "아플라톡신"},
    "农药残留(MRL)": {"EU": "Résidus de pesticides (LMR)", "US": "Pesticide residues (tolerance)", "JP": "農薬残留（MRL）", "KR": "농약 잔류（MRL）"},
}
PEST_EN = {
    "咪鲜胺": "Prochloraz", "啶虫脒": "Acetamiprid", "吡唑醚菌酯": "Pyraclostrobin",
    "噻虫嗪": "Thiamethoxam", "克百威": "Carbofuran", "多菌灵": "Carbendazim",
    "苯醚甲环唑": "Difenoconazole", "吡虫啉": "Imidacloprid", "毒死蜱": "Chlorpyrifos",
    "氯氰菊酯": "Cypermethrin", "高效氯氰菊酯": "Beta-cypermethrin", "联苯菊酯": "Bifenthrin",
    "甲氨基阿维菌素": "Emamectin benzoate", "氟氯氰菊酯": "Cyfluthrin",
}
STD_EN = {
    "GB 2760-2024": "National Food Safety Standard – Standards for Uses of Food Additives (GB 2760-2024)",
    "GB 2762-2022": "National Food Safety Standard – Maximum Levels of Contaminants in Foods (GB 2762-2022)",
    "GB 2763-2021": "National Food Safety Standard – Maximum Residue Limits for Pesticides in Food (GB 2763-2021)",
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

def classify_indicator(raw):
    """返回物质级简中指标名；无法识别（非具体物质）返回 None。"""
    s = to_simp(G(raw))
    low = s.lower()
    if "二氧化硫" in s or "亞硫酸" in s or "亚硫酸" in s or "sulfit" in low:
        return "二氧化硫/SO₂残留"
    if "鉛" in s or "铅" in s:
        return "铅(Pb)"
    if "鎘" in s or "镉" in s:
        return "镉(Cd)"
    if "汞" in s:
        return "汞(Hg)"
    if "砷" in s:
        return "砷(As)"
    if "黃曲霉" in s or "黄曲霉" in s or "aflatoxin" in low:
        return "黄曲霉毒素"
    if "農藥" in s or "农药" in s or "除害" in s or "pesticide" in low or "mrl" in low:
        return "农药残留(MRL)"
    if s in KNOWN_PESTICIDES:
        return s
    if any(k in s for k in ["菌落", "大肠", "沙门", "葡萄", "蜡样", "霉菌", "酵母"]):
        return s
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
    en = IND_EN.get(ind_simp)
    if en is None:
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
    wrap = Alignment(vertical="top", wrap_text=True)

    # ===== Sheet1 基础信息（列表 + 配料多语对照） =====
    ws = wb.active; ws.title = "基础信息"
    ws.append(["字段", "内容"])
    ws.cell(1, 1).fill = head_fill; ws.cell(1, 1).font = head_font
    ws.cell(1, 2).fill = head_fill; ws.cell(1, 2).font = head_font
    base_rows = [
        ("食品中文名", basic.get("name_cn")), ("国际/英文名", basic.get("aliases")),
        ("别名/同义词", basic.get("aliases")), ("类别", basic.get("category")),
        ("食品添加剂(含代码)", basic.get("additives")), ("添加剂说明", basic.get("additive_note")),
        ("过敏原", basic.get("allergens")), ("翻译置信度提示", basic.get("trans_conf")),
        ("skill 版本", "1.7.0"), ("生成日期", basic.get("gen_date")),
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
    for field in ["名称", "类别", "配料(芒果)"]:
        t = tmap.get(field, {})
        ws.append([field] + [G(t.get(l)) for l in langs])
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row, max_col=len(langs) + 1):
        for cell in row:
            cell.border = border; cell.alignment = wrap
    ws.column_dimensions["A"].width = 26
    for i in range(2, len(langs) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.column_dimensions["B"].width = 40

    # ===== Sheet3 食品安全风险识别表（31 列，源表） =====
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

    seen = set(); deduped = []
    for r in risks:
        key = (G(r.get("region")), G(r.get("std_no")), G(r.get("risk_type")), G(r.get("indicator")))
        if key in seen:
            continue
        seen.add(key); deduped.append(r)
    risks = expand_rows(deduped)

    source_map = {}
    r3 = 1
    for i, r in enumerate(risks, 1):
        r3 += 1
        reg = G(r.get("region"))
        zh, en, nat = REGION_INFO.get(reg, (reg, reg, reg))
        quote = G(r.get("quote"))
        qcn = G(r.get("quote_cn")) or QUOTE_CN.get(G(r.get("std_no")).strip(),
                                                   "（原文为中文）" if is_cjk(quote) else "待补充")
        ind_simp = sheet3_indicator(r)
        im = ind_multiling(ind_simp)
        sm = std_multiling(r.get("std_no"), r.get("std_name"), reg)
        claimed = claimed_std(r)
        ws3.append([
            i, reg, zh, en, nat,
            sm["zh"], sm["en"], sm["native"], G(r.get("std_no")),
            G(r.get("risk_type")),
            im["simp"], im["trad"], im["en"], im["native"],
            G(r.get("code")), G(r.get("scope")), quote, qcn,
            G(r.get("limit")), G(r.get("method")), G(r.get("restriction")), claimed,
            G(r.get("file")), G(r.get("url")), G(r.get("page_path")), G(r.get("hit_alias")),
            G(r.get("confidence")), G(r.get("published")), G(r.get("validity")),
            G(r.get("access")), G(r.get("note")),
        ])
        sub = classify_indicator(r.get("indicator") or (r.get("risk_type") + " " + r.get("quote")))
        if sub and G(r.get("limit")) not in ("", "[待填写]", "待补充"):
            source_map.setdefault((sub, reg), r3)
    widths = [6, 16, 12, 14, 16, 30, 30, 26, 16, 16, 16, 16, 20, 18, 18, 22, 46, 40, 16, 30, 22, 26, 18, 38, 28, 20, 10, 14, 12, 14, 24]
    for i, w in enumerate(widths, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row):
        for cell in row:
            cell.border = border; cell.alignment = wrap
    ws3.freeze_panes = "C2"; ws3.auto_filter.ref = ws3.dimensions

    # ===== Sheet4 指标对比查询（公式透视 + 官方全称列头） =====
    ws4 = wb.create_sheet("指标对比查询")
    reg_set = sorted({G(r.get("region")) for r in risks})
    regions_full = [REGION_FULLNAME.get(reg, reg) for reg in reg_set]
    BASE_CATS = ["二氧化硫/SO₂残留", "铅(Pb)", "镉(Cd)", "汞(Hg)", "砷(As)", "黄曲霉毒素", "农药残留(MRL)"]
    MICROBIAL = ["菌落总数", "大肠菌群", "大肠杆菌", "沙门氏菌", "金黄色葡萄球菌", "蜡样芽胞杆菌", "霉菌", "酵母"]
    specific = []
    for r in risks:
        ind = G(r.get("indicator"))
        if ind and ind not in BASE_CATS and ind not in MICROBIAL and ind in KNOWN_PESTICIDES:
            if ind not in specific:
                specific.append(ind)
    categories = BASE_CATS + specific + MICROBIAL

    ws4.append(["指标名称 \\ 国家/地区(组织)"] + regions_full)
    for c in range(1, len(regions_full) + 2):
        ws4.cell(1, c).fill = head_fill; ws4.cell(1, c).font = head_font
    ws4.cell(1, 1).alignment = Alignment(vertical="center", wrap_text=True)
    for cat in categories:
        row = [cat]
        for reg in reg_set:
            src_row = source_map.get((cat, reg))
            if src_row:
                cell_ref = f"{get_column_letter(LIMIT_COL)}{src_row}"
                sheet_ref = f"'食品安全风险识别表'!{cell_ref}"
                # 所有数据通过公式从 Sheet3 源表调取，点击可跳转源单元格
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
        "1. 微生物/致病菌指标为本表强制保留类别；本次检索数据未覆盖微生物限量条目，对应单元格置 [待填写]，需后续补充。",
        "2. 指标名称(官方语言)、法规名称(官方语言全称) 因检索未返回各辖区官方语言原文，暂置 [待填写]，需检索补充。",
        "3. 部分辖区仅给出框架性农药/污染物规定（无逐项具体限量值），其透视单元格亦为 [待填写]。",
    ]
    for j, n in enumerate(notes, 1):
        ws4.cell(gap_note_row + j, 1, n)
        ws4.cell(gap_note_row + j, 1).alignment = wrap

    wb.save(out_path)
    return out_path
```

## 九、示例交互
用户：「录入『螺蛳粉』，出食品安全风险识别表，覆盖全球。」（默认增强档）
→ Phase0 确认单品 + 选「增强」档，初始化覆盖集合；
→ Phase1 类别「方便食品」+ 配料（米粉、汤料包、酸笋、腐竹、花生、添加剂…）；
→ Phase2 翻 14 语（高棉/阿标低置信）；
→ Phase3 扩展：柳州螺蛳粉 / snail rice noodle / Liuzhou river snail rice noodle；
→ Phase4 检索 CN(GB 2760/2762)/US(21 CFR)/EU(1881/2006,1333/2008)/JP(肯定列表)/KR(Food Code)/TH/Codex GSFA/ISO/WHO-JECFA/ASEAN…；识别色素 INS 102 并关联（同号判定）；下载 PDF + 存 `_page.md`（失败则降级）；分批回报进度；
→ Phase5 输出 `<输出目录>/螺蛳粉/螺蛳粉_食品安全风险识别表.xlsx` + `standards/`。

## 十、实跑验证记录与运维须知（v1.5.0 新增）

> 以下经验来自「芒果干」全量档 6 簇并行实跑（2026-07-07），81 行原始 → 79 行去重，31 个存档文件。记录踩坑与对策，防止下次重蹈。

### 10.1 团队生命周期陷阱（⚠️ 最高频踩坑）
**现象**：TeamCreate 建立的团队会在**跨轮次会话间被自动回收**，连带把仍在跑的 sub-agent 一起清掉，表现为"无声消失"——agent 永远不回传，文件也不落盘。
**对策**：
- **同消息建队+派工**：TeamCreate 和 Agent 派发必须在**同一条消息**里完成（先 TeamCreate，紧接 Agent 调用），避免团队还没建好就派工导致"No active team found"；
- **重派策略**：若某簇超时未回传，先 `ls` 检查文件是否已落盘；若无，重建团队+同消息重派，缩减 `max_turns`（如 45→40）和每辖区查询上限（如 6→3）加速；
- **不要依赖跨轮次的团队上下文**：每次恢复执行时，先重建团队再继续派工。

### 10.2 sub-agent 完成但未落盘结果文件
**现象**：agent 完成了检索（甚至下载了 PDF/TXT 原文），但返回时只发消息、没写 `.md` 结果文件，导致整合脚本解析不到该簇数据。
**对策**：
- 派发 sub-agent 时，在提示词中**强制要求**：「完成后必须用 Write 工具把完整结果写入 `<输出目录>/<食品名>/<食品名>_<簇名>_检索结果.md`，并在回复中附小结」；
- 主 agent 回收后**立即 `ls` 验证文件存在**；若缺失，从已下载的 standards/ 文件（PDF/TXT）中提取原文补建结果文件；
- 整合脚本 `integrate.py` 应对"部分簇有文件、部分无"做容错（跳过缺失文件而非崩溃）。

### 10.3 Markdown 表头格式一致性
**现象**：不同 sub-agent 写的表头列名可能带或不带括号后缀（如 `有效性(现行有效/已废止/…)` vs `有效性`），导致解析器按精确匹配时整表 0 行。
**对策**：
- 派发提示词中**固定表头列名**（不带括号后缀），列序 18 列严格一致；
- 整合脚本用**列序索引**而非列名匹配解析（按 `|` 分割后的列位置取值），对表头微差异容错。

### 10.4 后端网络 502 / ENOTFOUND 抖动
**现象**：`copilot.tencent.com` 偶发 502 ENOTFOUND，导致 sub-agent 整体失败（非逻辑错误）。
**对策**：按兜底规则重试 1 次；重试时精简查询量、缩短 `max_turns`；若仍失败，**先把已完成的簇出最终 Excel**，在「待完善事项」标注缺失簇，用户可随时让我重跑那一段（脚本已就绪）。

### 10.5 整合与生成脚本（工具链）
实跑验证了以下脚本链，已落盘在输出目录，可复用：
- `integrate.py`：解析各簇 `.md` 结果文件 → 统一 `risks.json`（跨簇去重，键=辖区+标准号+风险类型）；
- `build_xlsx.py`：读 `basic.json` + `translations.json` + `risks.json` → 生成 3-sheet Excel；
- 两个脚本均用受管 venv 的 openpyxl，重跑只需两条命令：
  ```bash
  <venv>/Scripts/python integrate.py
  <venv>/Scripts/python build_xlsx.py
  ```

### 10.6 实跑关键发现
- 全量档 6 簇并行是**重任务**（总耗时约 1.5 小时含重试），建议优先用**增强档**验证 skill 健康度，再放开全量；
- 芒果干（纯芒果、无添加剂）场景下，SO₂ 残留限量仍是各国核心关注点（因干燥工艺可能带入微量），日本 0.030 g/kg 最严、中国 0.1 g/kg、马来西亚/印度 1000-2000 mg/kg；
- 农残/污染物具体 MRL 多数辖区需查官方数据库（非网页可直取），普遍标「待补充」是诚实做法，**不应编造**。

### 10.7 v1.7.0 结构变更记录（2026-07-07）
- 取消独立的「多语翻译对照」Sheet：配料多语进 Sheet1 子表，指标/法规多语进 Sheet3（指标名四语、法规名三语）；
- Sheet3 由 25 列扩至 31 列（新增指标名称·简中/繁中/英/官方语言、法规名称·简中/英/官方语言全称、声称执行标准）；
- Sheet4 改为**公式透视**：地区列头用官方全称（AU/NZ→澳大利亚/新西兰食品标准局（FSANZ）、APEC→亚太经济合作组织（APEC）…），交叉单元格 `=HYPERLINK("#'食品安全风险识别表'!S{行}", …)` 从源表调取并可点击跳转；
- 农残逐物质展开（咪鲜胺/啶虫脒/吡虫啉…）、污染物展开为 铅(Pb)/镉(Cd)/汞(Hg)/砷(As)；**强制保留微生物/致病菌指标**（菌落总数/大肠菌群/大肠杆菌/沙门氏菌/金黄色葡萄球菌/蜡样芽胞杆菌/霉菌/酵母），不得遗漏；
- 新增「声称执行标准」字段（对应海外强制法规/自愿性认证，替代国内企业执行标准逻辑）。未检索到官方语言原文或微生物限量时，相关单元格标 `[待填写]` 并在 Sheet4 末附「信息缺口声明」。

---
> 免责声明：本 skill 由 AI 驱动，翻译与合规检索结果仅供参考，重要出口/上市决策须由具备资质的专业人员依据目标国主管机构及国际组织最新发布文件核验；本表不构成法律意见。
