Option Explicit

Dim objWord, objDoc, objTable, objSelection
Dim todayStr, outputPath

Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

todayStr = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2)
outputPath = "业绩数据整合脚本执行流程说明_" & todayStr & ".docx"

objSelection.Style = "标题"
objSelection.TypeText "业绩数据整合脚本执行流程说明"
objSelection.Font.Name = "微软雅黑"
objSelection.Font.Size = 16
objSelection.Font.Bold = True
objSelection.ParagraphFormat.Alignment = 1
objSelection.TypeParagraph()

objSelection.Style = "正文"
objSelection.Font.Name = "宋体"
objSelection.Font.Size = 11
objSelection.Font.Bold = False
objSelection.ParagraphFormat.Alignment = 0

objSelection.TypeText "生成日期: " & todayStr
objSelection.TypeParagraph()
objSelection.TypeText "脚本文件: performance_data_integrated_generator.py"
objSelection.TypeParagraph()
objSelection.TypeText "脚本版本: v1.0"
objSelection.TypeParagraph()
objSelection.TypeText "合并来源: query_to_v01.py (v2.5) + ngp_to_v0ngp.py (v1.8)"
objSelection.TypeParagraph()
objSelection.TypeParagraph()

AddHeading objSelection, "一、脚本概述", 1
AddHeading objSelection, "1.1 功能定位", 2

objSelection.TypeText "本脚本是业绩数据整合统一生成脚本，主要功能包括："
objSelection.TypeParagraph()

AddNumberedList objSelection, Array( _
    "一次执行即可生成包含 V0 和 V0-NGP 两个 sheet 的业绩数据-整合文件", _
    "V0 sheet: 综合查询结果数据 → 80列V0格式（与参考文件V0数据格式一致）", _
    "V0-NGP sheet: NGP数据 → 39列V0-NGP格式（与参考文件-NGP数据格式一致）", _
    "同时生成 V0-IYB业绩 sheet（含端口/分层映射）", _
    "自动生成转换规则说明 Word 文档", _
    "自动生成结果验证报告（对比参考文件）" _
)

AddHeading objSelection, "1.2 主入口函数", 2

objSelection.TypeText "核心入口函数：generate_integrated_file()"
objSelection.TypeParagraph()

CreateTable objSelection, Array("参数", "说明"), Array( _
    Array("query_file_path", "综合查询结果 xlsx 文件路径"), _
    Array("ngp_file_path", "NGP xlsx 文件路径"), _
    Array("mapping_table_path", "业务部门字段映射表路径（可选，自动查找）"), _
    Array("reference_file_path", "参考文件路径（必填，用于保留其他sheet公式结构）"), _
    Array("output_dir", "输出目录（可选，默认与综合查询文件同目录）") _
)

objSelection.TypeText "返回值：output_file_path（主输出文件路径）、doc_path（转换规则文档路径）、report_path（验证报告路径）"
objSelection.TypeParagraph()

AddHeading objSelection, "二、执行步骤详解", 1

AddHeading objSelection, "2.1 Part 1: V0 数据处理", 2
objSelection.TypeText "处理函数：process_v0_data(query_file_path, mapping_table_path)"
objSelection.TypeParagraph()
objSelection.TypeText "处理目标：综合查询结果 → V0 格式（80列）"
objSelection.TypeParagraph()

CreateTable objSelection, Array("步骤", "操作", "执行内容", "关键说明"), Array( _
    Array("Step 1", "读取综合查询数据", "读取 xlsx 文件，获取原始数据", "原始数据: N行 × M列"), _
    Array("Step 2", "删除永领致远", "删除签单供应商=""永领致远顾问有限公司""的行", "删除后剩余: N-D行"), _
    Array("Step 2b", "保单状态回填", "保单状态优先，为空时用订单状态回填", "保单状态为空 → 使用订单状态值"), _
    Array("Step 3", "繁简转换", "繁体→简体（使用 OpenCC t2s）", "自动跳过数值列"), _
    Array("Step 4", "产品名称映射", "按映射表清洗产品名称", "旧产品名称 → 新产品名称"), _
    Array("Step 5", "列名映射", "综合查询列名 → V0 格式（80列）", "丢弃不需要的列，缺失列填NaN"), _
    Array("Step 6", "保单状态映射", "原始值 → V0 标准值", "如""已生效""→""生效""、""PENDING""→""pending"""), _
    Array("Step 7", "数据清洗", "多项清洗操作", "详见下表") _
)

AddHeading objSelection, "Step 7 数据清洗详细内容", 3

CreateTable objSelection, Array("清洗项", "操作内容", "示例"), Array( _
    Array("产品品类映射", "按映射表转换产品品类", "危疾计划→重疾、医疗计划→医疗"), _
    Array("币种映射", "币种标准化转换", "美元→美金、港元→港币"), _
    Array("供款方式映射", "供款方式标准化", "整付保费→整付"), _
    Array("电话修复", "去掉.0后缀（float→str残留）", "12345678.0→12345678"), _
    Array("保单号码格式化", "纯数字→int，非纯数字→str", "00123→123、A123→A123"), _
    Array("客户分群匹配", "根据PI&NONPI文件匹配", "是→PI、其他→NONPI"), _
    Array("年期清洗", "去掉""年""字，整付保费特殊处理", "10年→10、整付保费→1"), _
    Array("数值列处理", "数值列中""0""→NaN", "保监征费、合计等列"), _
    Array("佣金模式", "NaN→0", "-") _
)

AddHeading objSelection, "2.2 Part 2: V0-NGP 数据处理", 2
objSelection.TypeText "处理函数：process_v0ngp_data(ngp_file_path, mapping_table_path)"
objSelection.TypeParagraph()
objSelection.TypeText "处理目标：NGP 数据 → V0-NGP 格式（39列）"
objSelection.TypeParagraph()

CreateTable objSelection, Array("步骤", "操作", "执行内容"), Array( _
    Array("Step 1", "读取 NGP 数据", "读取 xlsx 文件 Sheet1"), _
    Array("Step 2", "列顺序对齐", "确保列顺序与 V0-NGP 一致（39列），缺失补None，多余删除"), _
    Array("Step 3", "繁简转换", "繁体→简体转换（文本列）"), _
    Array("Step 4", "产品名称映射", "按映射表清洗产品名称"), _
    Array("Step 5", "数据清洗", "日期/金额类型转换、币种/保单状态/产品品类/供款方式映射、年期清洗、保单号码格式化、电话修复") _
)

AddHeading objSelection, "2.3 Part 2c: V0-IYB业绩 数据生成", 2
objSelection.TypeText "处理函数：process_iyb_v0_data(df_v0)"
objSelection.TypeParagraph()
objSelection.TypeText "处理目标：从 V0 数据中筛选 IYB 供应商数据，生成 V0-IYB业绩 格式（20列）"
objSelection.TypeParagraph()

CreateTable objSelection, Array("步骤", "执行内容"), Array( _
    Array("供应商筛选", "仅保留 IYB_SUPPLIERS 白名单中的供应商（9家）"), _
    Array("日期过滤", "排除 2024 年及以前的保单（未签单的活跃保单保留）"), _
    Array("端口映射", "业务细分 → 端口（如""BK业务""→""三方机构""）"), _
    Array("分层映射", "市场分层 → 分层（如""银行网点""→""银行""）"), _
    Array("列结构", "输出 20 列，顺序与 V0-IYB业绩 模板一致") _
)

AddHeading objSelection, "2.4 Part 3: 创建输出文件并写入数据", 2
objSelection.TypeText "核心策略：openpyxl 复制参考文件 + Excel COM 写入数据并重算公式"
objSelection.TypeParagraph()

AddHeading objSelection, "Excel COM 操作步骤", 3

CreateTable objSelection, Array("阶段", "操作", "执行内容"), Array( _
    Array("1", "清理残留进程", "强制关闭 Excel 进程，避免文件锁定"), _
    Array("2", "复制参考文件", "shutil.copy2 字节级复制，保留公式 sheet XML"), _
    Array("3", "写入 V0 数据", "分块写入（chunk=1000），修复非日期列误设日期格式"), _
    Array("4", "写入 V0-NGP 数据", "分块写入，修复非日期列误设日期格式"), _
    Array("5", "写入 V0-IYB业绩 数据", "分块写入，修复 G2/H2 XLOOKUP 公式短范围问题"), _
    Array("6", "保存数据写入结果", "防止后续重算崩溃丢失数据"), _
    Array("7", "修复 spill sheet 日期格式错配", "按列名判断而非索引，避免误改"), _
    Array("8", "重算所有公式", "CalculateFullRebuild() 基于新数据重算"), _
    Array("9", "V0-IYB业绩 端口/分层兜底", "验证 G/H 列，对空值直接写入映射值"), _
    Array("10", "统计公式重算结果", "统计各 sheet 公式数量和错误值"), _
    Array("11", "清理 spilled 范围错误值", "清除错误常量"), _
    Array("12", "修复 #DIV/0! 除法公式", "包裹 IFERROR，消除零除错误"), _
    Array("13", "修复 V0-合并表 E列 #N/A", "CHOOSE/MATCH 包裹 IFERROR"), _
    Array("14", "设置 sheet 缩放比例", "所有 sheet 设置为 100%"), _
    Array("15", "另存 V0-SunLife", "保存为""永明业绩数据-YYYYMMDD.xlsx""（公式转静态值）"), _
    Array("16", "另存 V0-IYB业绩", "保存为""IYB业绩追踪周报-YYYYMMDD.xlsx""（V0+V2+匹配表）"), _
    Array("17", "关闭 Excel", "退出 COM 进程") _
)

AddHeading objSelection, "2.5 Part 4: 生成转换规则 Word 文档", 2
objSelection.TypeText "处理函数：generate_rules_document(doc_path, date_str)"
objSelection.TypeParagraph()
objSelection.TypeText "生成内容："
objSelection.TypeParagraph()

AddBulletedList objSelection, Array( _
    "目录（TOC 字段）", _
    "V0 Sheet 转换规则：数据源、数据逻辑、列名映射、值域映射、数据格式化、格式设置", _
    "V0-NGP Sheet 转换规则：同上", _
    "共享映射规则汇总：保单状态、产品品类、币种、供款方式、产品名称映射" _
)

AddHeading objSelection, "2.6 Part 5: 生成结果验证报告", 2
objSelection.TypeText "处理函数：run_comparison() + generate_comparison_report()"
objSelection.TypeParagraph()
objSelection.TypeText "对比内容："
objSelection.TypeParagraph()

AddBulletedList objSelection, Array( _
    "数据量验证：各 sheet 行数、列数对比", _
    "字段一致性验证：逐列值一致率、不一致样本、映射规则验证", _
    "映射规则验证：列出所有映射规则及其执行效果", _
    "格式一致性验证：表头格式、数据行格式、列宽对比" _
)

AddHeading objSelection, "三、核心映射规则汇总", 1

AddHeading objSelection, "3.1 保单状态映射", 2
objSelection.TypeText "共24条映射规则，分为""保单状态""和""订单状态回填""两类"
objSelection.TypeParagraph()

CreateTable objSelection, Array("原始值", "V0标准值", "类型"), Array( _
    Array("已生效", "生效", "保单状态"), _
    Array("已生效-未回执", "生效", "保单状态"), _
    Array("已生效-待核验", "生效", "保单状态"), _
    Array("保单失效", "失效", "保单状态"), _
    Array("已交单至保险公司", "已签单", "保单状态"), _
    Array("PENDING", "pending", "保单状态"), _
    Array("PENDING-待补充资料", "pending", "保单状态"), _
    Array("PENDING-资料已补充待审核", "pending", "保单状态"), _
    Array("PENDING-内部处理中", "pending", "保单状态"), _
    Array("待核保", "待批核", "保单状态"), _
    Array("待生效", "待批核", "保单状态"), _
    Array("取消投保中", "取消投保", "保单状态"), _
    Array("冷静期内退保", "退保", "保单状态"), _
    Array("申请退保中", "退保", "保单状态"), _
    Array("已撤销", "取消预约", "订单状态回填"), _
    Array("预约成功", "排期", "订单状态回填"), _
    Array("投保文件待复核", "pending", "订单状态回填"), _
    Array("签单完成", "已签单", "订单状态回填"), _
    Array("投保文件复核驳回", "拒保", "订单状态回填"), _
    Array("已提交预约待审核", "pending", "订单状态回填"), _
    Array("待交单至保险公司", "已签单", "订单状态回填"), _
    Array("预约中", "排期", "订单状态回填"), _
    Array("待确认转介信息", "pending", "订单状态回填"), _
    Array("预约资料待修改", "pending", "订单状态回填") _
)

AddHeading objSelection, "3.2 产品品类映射", 2

CreateTable objSelection, Array("细分品类", "标准品类"), Array( _
    Array("危疾计划", "重疾"), _
    Array("人寿保障", "人寿"), _
    Array("医疗计划", "医疗"), _
    Array("年金计划", "年金"), _
    Array("高端医疗", "医疗"), _
    Array("医疗计划(自愿医保)", "医疗"), _
    Array("人寿险", "人寿"), _
    Array("投资相连人寿保险计划", "投连险"), _
    Array("其它类型", "医疗"), _
    Array("意外及伤残", "医疗") _
)

AddHeading objSelection, "3.3 币种映射", 2

CreateTable objSelection, Array("原始值", "标准值"), Array( _
    Array("美元", "美金"), _
    Array("港元", "港币") _
)

AddHeading objSelection, "3.4 供款方式映射", 2

CreateTable objSelection, Array("原始值", "标准值"), Array( _
    Array("整付保费", "整付") _
)

AddHeading objSelection, "3.5 IYB 端口映射", 2

CreateTable objSelection, Array("业务细分", "端口"), Array( _
    Array("BK业务", "三方机构"), _
    Array("成事家办", "A端"), _
    Array("天领业务", "A端"), _
    Array("合伙转介业务", "B端"), _
    Array("永明经代", "B端"), _
    Array("同行经代", "B端"), _
    Array("ICLUB业务", "B端"), _
    Array("IFA业务", "B端") _
)

AddHeading objSelection, "3.6 IYB 分层映射", 2

CreateTable objSelection, Array("市场分层", "分层"), Array( _
    Array("银行网点", "银行"), _
    Array("同行机构转介", "非持牌转介人"), _
    Array("持牌转介", "非持牌转介人"), _
    Array("异业机构转介", "非持牌转介人"), _
    Array("同行个人转介", "非持牌转介人"), _
    Array("持牌代理人", "非持牌转介人"), _
    Array("持牌资管", "持牌转介人"), _
    Array("持牌经纪", "持牌转介人"), _
    Array("IYB分公司", "成事家办"), _
    Array("天领机构", "天领"), _
    Array("天领贴牌", "天领"), _
    Array("IFA业务", "非持牌转介人") _
)

AddHeading objSelection, "四、输出文件清单", 1

CreateTable objSelection, Array("文件名", "说明", "生成时机"), Array( _
    Array("业绩数据-整合-YYYYMMDD.xlsx", "主输出文件（含 V0、V0-NGP、V0-IYB业绩 及其他公式 sheet）", "Part 3"), _
    Array("永明业绩数据-YYYYMMDD.xlsx", "V0-SunLife sheet 单独导出（公式转静态值）", "Part 3"), _
    Array("IYB业绩追踪周报-YYYYMMDD.xlsx", "V0-IYB业绩 单独导出（含 V0+V2+匹配表）", "Part 3"), _
    Array("业绩数据-整合-转换规则说明-YYYYMMDD.docx", "转换规则说明文档", "Part 4"), _
    Array("业绩数据-整合-结果验证报告-YYYYMMDD.docx", "结果验证报告（对比参考文件）", "Part 5") _
)

AddHeading objSelection, "五、关键技术点", 1

AddTechPoint objSelection, "5.1 公式保留策略", "通过 shutil.copy2 字节级复制参考文件，再用 Excel COM 写入数据并重算，确保公式引用结构完整保留"
AddTechPoint objSelection, "5.2 XLOOKUP 公式修复", "自动将参考模板中过小的查找范围（如 $J$1:$J$10）替换为完整列引用（J:J）"
AddTechPoint objSelection, "5.3 日期格式修复", "针对 spill sheet 中非日期列误设日期格式的问题，按列名判断并修复"
AddTechPoint objSelection, "5.4 端口/分层兜底", "即使 XLOOKUP 公式修复，仍对空值行直接写入映射值作为兜底"
AddTechPoint objSelection, "5.5 错误值处理", "清理 #N/A、#VALUE!、#DIV/0! 等错误值，修复除法公式零除问题"
AddTechPoint objSelection, "5.6 数据分块写入", "COM 写入采用 chunk=1000 分块策略，避免内存溢出"
AddTechPoint objSelection, "5.7 进程清理", "写入前强制关闭残留 Excel 进程，避免文件锁定"

AddHeading objSelection, "六、执行流程图", 1

objSelection.TypeText "开始"
objSelection.TypeParagraph()
objSelection.TypeText "  │"
objSelection.TypeParagraph()
objSelection.TypeText "  ▼"
objSelection.TypeParagraph()
objSelection.TypeText "┌─────────────────────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "│ 1. 读取输入参数                              │"
objSelection.TypeParagraph()
objSelection.TypeText "│    - 综合查询文件                            │"
objSelection.TypeParagraph()
objSelection.TypeText "│    - NGP文件                                 │"
objSelection.TypeParagraph()
objSelection.TypeText "│    - 映射表（自动查找）                       │"
objSelection.TypeParagraph()
objSelection.TypeText "│    - 参考文件（自动查找）                     │"
objSelection.TypeParagraph()
objSelection.TypeText "└─────────────────────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │"
objSelection.TypeParagraph()
objSelection.TypeText "  ├──────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 1: V0 数据处理           │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ process_v0_data()             │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ 7个步骤 → 80列 V0 格式        │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 2: V0-NGP 数据处理       │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ process_v0ngp_data()          │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ 5个步骤 → 39列 V0-NGP 格式    │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 2c: V0-IYB业绩 数据生成  │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ process_iyb_v0_data()         │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ 筛选IYB供应商 → 20列格式      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 3: 创建输出文件           │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ - openpyxl 复制参考文件        │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ - Excel COM 写入数据并重算     │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ - 修复公式和格式              │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ - 另存永明/IYB业绩文件        │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 4: 生成转换规则文档       │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ generate_rules_document()     │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Word文档：映射规则汇总         │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "  │  ┌───────────────────────────────┐"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ Part 5: 生成结果验证报告       │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ run_comparison()              │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  │ 对比参考文件，生成验证报告     │"
objSelection.TypeParagraph()
objSelection.TypeText "  │  └───────────────────────────────┘"
objSelection.TypeParagraph()
objSelection.TypeText "  │                      │"
objSelection.TypeParagraph()
objSelection.TypeText "  ▼                      ▼"
objSelection.TypeParagraph()
objSelection.TypeText "完成"
objSelection.TypeParagraph()
objSelection.TypeText "  │"
objSelection.TypeParagraph()
objSelection.TypeText "  └── 输出文件：业绩数据-整合-YYYYMMDD.xlsx"
objSelection.TypeParagraph()
objSelection.TypeText "      转换规则说明.docx"
objSelection.TypeParagraph()
objSelection.TypeText "      结果验证报告.docx"
objSelection.TypeParagraph()
objSelection.TypeText "      永明业绩数据-YYYYMMDD.xlsx"
objSelection.TypeParagraph()
objSelection.TypeText "      IYB业绩追踪周报-YYYYMMDD.xlsx"
objSelection.TypeParagraph()

objDoc.SaveAs(outputPath)
objDoc.Close()
objWord.Quit()

MsgBox "Word文档已生成: " & outputPath

Function AddHeading(objSel, text, level)
    objSel.TypeParagraph()
    objSel.Style = "标题" & level
    objSel.Font.Name = "黑体"
    objSel.Font.Size = 14 - level * 1
    objSel.Font.Bold = True
    objSel.TypeText text
    objSel.TypeParagraph()
    objSel.Style = "正文"
    objSel.Font.Name = "宋体"
    objSel.Font.Size = 11
    objSel.Font.Bold = False
End Function

Function AddNumberedList(objSel, items)
    Dim i
    For i = LBound(items) To UBound(items)
        objSel.TypeText (i + 1) & ". " & items(i)
        objSel.TypeParagraph()
    Next
End Function

Function AddBulletedList(objSel, items)
    Dim i
    For i = LBound(items) To UBound(items)
        objSel.TypeText "· " & items(i)
        objSel.TypeParagraph()
    Next
End Function

Function AddTechPoint(objSel, title, desc)
    objSel.Font.Bold = True
    objSel.Font.Color = RGB(26, 95, 180)
    objSel.TypeText title
    objSel.TypeParagraph()
    objSel.Font.Bold = False
    objSel.Font.Color = RGB(0, 0, 0)
    objSel.TypeText "  " & desc
    objSel.TypeParagraph()
End Function

Function CreateTable(objSel, headers, data)
    Dim objTable, numCols, numRows, i, j
    numCols = UBound(headers) + 1
    numRows = UBound(data) + 1
    
    Set objTable = objSel.Tables.Add(objSel.Range, numRows + 1, numCols)
    
    For j = 0 To numCols - 1
        objTable.Cell(1, j + 1).Range.Text = headers(j)
        objTable.Cell(1, j + 1).Range.Font.Bold = True
        objTable.Cell(1, j + 1).Range.Font.Name = "宋体"
        objTable