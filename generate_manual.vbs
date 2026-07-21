Option Explicit

Dim objWord, objDoc, objRange, objTable, objRow, objCell, objPara
Dim strOutputPath, strToday

strToday = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
strOutputPath = "业绩数据整合脚本执行流程说明-" & strToday & ".docx"

Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.DisplayAlerts = False

Set objDoc = objWord.Documents.Add()

' 设置页面
With objDoc.PageSetup
    .TopMargin = objWord.InchesToPoints(0.7)
    .BottomMargin = objWord.InchesToPoints(0.7)
    .LeftMargin = objWord.InchesToPoints(0.7)
    .RightMargin = objWord.InchesToPoints(0.7)
End With

' 设置默认样式
With objDoc.Styles("Normal").Font
    .Name = "宋体"
    .Size = 11
End With

' ==================== 标题 ====================
Set objPara = objDoc.Content.Paragraphs.Add()
objPara.Range.Text = "业绩数据整合脚本执行流程说明"
objPara.Range.Font.Name = "微软雅黑"
objPara.Range.Font.Size = 16
objPara.Range.Font.Bold = True
objPara.Alignment = 1 ' wdAlignParagraphCenter
objPara.Range.InsertParagraphAfter

' 元信息
Set objPara = objDoc.Content.Paragraphs.Add()
objPara.Range.Text = "生成日期: " & Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)
objPara.Range.Font.Name = "宋体"
objPara.Range.Font.Size = 11
objPara.Range.InsertParagraphAfter

Set objPara = objDoc.Content.Paragraphs.Add()
objPara.Range.Text = "脚本文件: performance_data_integrated_generator.py"
objPara.Range.Font.Name = "宋体"
objPara.Range.Font.Size = 11
objPara.Range.InsertParagraphAfter

Set objPara = objDoc.Content.Paragraphs.Add()
objPara.Range.Text = "脚本版本: v1.0"
objPara.Range.Font.Name = "宋体"
objPara.Range.Font.Size = 11
objPara.Range.InsertParagraphAfter

Set objPara = objDoc.Content.Paragraphs.Add()
objPara.Range.Text = "合并来源: query_to_v01.py (v2.5) + ngp_to_v0ngp.py (v1.8)"
objPara.Range.Font.Name = "宋体"
objPara.Range.Font.Size = 11
objPara.Range.InsertParagraphAfter
objPara.Range.InsertParagraphAfter

' ==================== 目录 ====================
AddHeading objDoc, "目录", 1
objPara.Range.InsertParagraphAfter

' ==================== 一、脚本概述 ====================
AddHeading objDoc, "一、脚本概述", 1

AddHeading objDoc, "1.1 功能定位", 2
AddParagraph objDoc, "本脚本是业绩数据整合统一生成脚本，主要功能包括："

AddBullet objDoc, "一次执行即可生成包含 V0 和 V0-NGP 两个 sheet 的业绩数据-整合文件"
AddBullet objDoc, "V0 sheet: 综合查询结果数据 → 80列V0格式（与参考文件V0数据格式一致）"
AddBullet objDoc, "V0-NGP sheet: NGP数据 → 39列V0-NGP格式（与参考文件-NGP数据格式一致）"
AddBullet objDoc, "同时生成 V0-IYB业绩 sheet（含端口/分层映射）"
AddBullet objDoc, "自动生成转换规则说明 Word 文档"
AddBullet objDoc, "自动生成结果验证报告（对比参考文件）"

AddHeading objDoc, "1.2 主入口函数", 2
AddParagraph objDoc, "核心入口函数：generate_integrated_file()"

' 参数表格
Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter

Set objCell = objTable.Cell(1, 1)
objCell.Range.Text = "参数"
SetCellFont objCell, True, 10

Set objCell = objTable.Cell(1, 2)
objCell.Range.Text = "说明"
SetCellFont objCell, True, 10

AddTableRow objTable, Array("query_file_path", "综合查询结果 xlsx 文件路径")
AddTableRow objTable, Array("ngp_file_path", "NGP xlsx 文件路径")
AddTableRow objTable, Array("mapping_table_path", "业务部门字段映射表路径（可选，自动查找）")
AddTableRow objTable, Array("reference_file_path", "参考文件路径（必填，用于保留其他sheet公式结构）")
AddTableRow objTable, Array("output_dir", "输出目录（可选，默认与综合查询文件同目录）")

objTable.Range.InsertParagraphAfter
AddParagraph objDoc, "返回值：output_file_path（主输出文件路径）、doc_path（转换规则文档路径）、report_path（验证报告路径）、manual_path（执行流程说明文档路径）"

' ==================== 二、执行步骤详解 ====================
AddHeading objDoc, "二、执行步骤详解", 1

' 2.1 Part 1: V0 数据处理
AddHeading objDoc, "2.1 Part 1: V0 数据处理", 2
AddParagraph objDoc, "处理函数：process_v0_data(query_file_path, mapping_table_path)"
AddParagraph objDoc, "处理目标：综合查询结果 → V0 格式（80列）"

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 4)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "步骤"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "操作"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "执行内容"
SetCellFont objTable.Cell(1, 4), True, 10
objTable.Cell(1, 4).Range.Text = "关键说明"

AddTableRow objTable, Array("Step 1", "读取综合查询数据", "读取 xlsx 文件，获取原始数据", "原始数据: N行 × M列")
AddTableRow objTable, Array("Step 2", "删除永领致远", "删除签单供应商=""永领致远顾问有限公司""的行", "删除后剩余: N-D行")
AddTableRow objTable, Array("Step 2b", "保单状态回填", "保单状态优先，为空时用订单状态回填", "保单状态为空 → 使用订单状态值")
AddTableRow objTable, Array("Step 3", "繁简转换", "繁体→简体（使用 OpenCC t2s）", "自动跳过数值列")
AddTableRow objTable, Array("Step 4", "产品名称映射", "按映射表清洗产品名称", "旧产品名称 → 新产品名称")
AddTableRow objTable, Array("Step 5", "列名映射", "综合查询列名 → V0 格式（80列）", "丢弃不需要的列，缺失列填NaN")
AddTableRow objTable, Array("Step 6", "保单状态映射", "原始值 → V0 标准值", "如""已生效""→""生效""、""PENDING""→""pending""")
AddTableRow objTable, Array("Step 7", "数据清洗", "多项清洗操作", "详见下表")

objTable.Range.InsertParagraphAfter

AddHeading objDoc, "Step 7 数据清洗详细内容", 3

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 3)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "清洗项"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "操作内容"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "示例"

AddTableRow objTable, Array("产品品类映射", "按映射表转换产品品类", "危疾计划→重疾、医疗计划→医疗")
AddTableRow objTable, Array("币种映射", "币种标准化转换", "美元→美金、港元→港币")
AddTableRow objTable, Array("供款方式映射", "供款方式标准化", "整付保费→整付")
AddTableRow objTable, Array("电话修复", "去掉.0后缀（float→str残留）", "12345678.0→12345678")
AddTableRow objTable, Array("保单号码格式化", "纯数字→int，非纯数字→str", "00123→123、A123→A123")
AddTableRow objTable, Array("客户分群匹配", "根据PI&NONPI文件匹配", "是→PI、其他→NONPI")
AddTableRow objTable, Array("年期清洗", "去掉""年""字，整付保费特殊处理", "10年→10、整付保费→1")
AddTableRow objTable, Array("数值列处理", "数值列中""0""→NaN", "保监征费、合计等列")
AddTableRow objTable, Array("佣金模式", "NaN→0", "-")

' 2.2 Part 2: V0-NGP 数据处理
AddHeading objDoc, "2.2 Part 2: V0-NGP 数据处理", 2
AddParagraph objDoc, "处理函数：process_v0ngp_data(ngp_file_path, mapping_table_path)"
AddParagraph objDoc, "处理目标：NGP 数据 → V0-NGP 格式（39列）"

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 3)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "步骤"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "操作"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "执行内容"

AddTableRow objTable, Array("Step 1", "读取 NGP 数据", "读取 xlsx 文件 Sheet1")
AddTableRow objTable, Array("Step 2", "列顺序对齐", "确保列顺序与 V0-NGP 一致（39列），缺失补None，多余删除")
AddTableRow objTable, Array("Step 3", "繁简转换", "繁体→简体转换（文本列）")
AddTableRow objTable, Array("Step 4", "产品名称映射", "按映射表清洗产品名称")
AddTableRow objTable, Array("Step 5", "数据清洗", "日期/金额类型转换、币种/保单状态/产品品类/供款方式映射、年期清洗、保单号码格式化、电话修复")

' 2.3 Part 2c: V0-IYB业绩 数据生成
AddHeading objDoc, "2.3 Part 2c: V0-IYB业绩 数据生成", 2
AddParagraph objDoc, "处理函数：process_iyb_v0_data(df_v0)"
AddParagraph objDoc, "处理目标：从 V0 数据中筛选 IYB 供应商数据，生成 V0-IYB业绩 格式（20列）"

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "步骤"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "执行内容"

AddTableRow objTable, Array("供应商筛选", "仅保留 IYB_SUPPLIERS 白名单中的供应商（9家）")
AddTableRow objTable, Array("日期过滤", "排除 2024 年及以前的保单（未签单的活跃保单保留）")
AddTableRow objTable, Array("端口映射", "业务细分 → 端口（如""BK业务""→""三方机构""）")
AddTableRow objTable, Array("分层映射", "市场分层 → 分层（如""银行网点""→""银行""）")
AddTableRow objTable, Array("列结构", "输出 20 列，顺序与 V0-IYB业绩 模板一致")

' 2.4 Part 3: 创建输出文件并写入数据
AddHeading objDoc, "2.4 Part 3: 创建输出文件并写入数据", 2
AddParagraph objDoc, "核心策略：openpyxl 复制参考文件 + Excel COM 写入数据并重算公式"

AddHeading objDoc, "Excel COM 操作步骤", 3

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 3)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "阶段"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "操作"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "执行内容"

AddTableRow objTable, Array("1", "清理残留进程", "强制关闭 Excel 进程，避免文件锁定")
AddTableRow objTable, Array("2", "复制参考文件", "shutil.copy2 字节级复制，保留公式 sheet XML")
AddTableRow objTable, Array("3", "写入 V0 数据", "分块写入（chunk=1000），修复非日期列误设日期格式")
AddTableRow objTable, Array("4", "写入 V0-NGP 数据", "分块写入，修复非日期列误设日期格式")
AddTableRow objTable, Array("5", "写入 V0-IYB业绩 数据", "分块写入，修复 G2/H2 XLOOKUP 公式短范围问题")
AddTableRow objTable, Array("6", "保存数据写入结果", "防止后续重算崩溃丢失数据")
AddTableRow objTable, Array("7", "修复 spill sheet 日期格式错配", "按列名判断而非索引，避免误改")
AddTableRow objTable, Array("8", "重算所有公式", "CalculateFullRebuild() 基于新数据重算")
AddTableRow objTable, Array("9", "V0-IYB业绩 端口/分层兜底", "验证 G/H 列，对空值直接写入映射值")
AddTableRow objTable, Array("10", "统计公式重算结果", "统计各 sheet 公式数量和错误值")
AddTableRow objTable, Array("11", "清理 spilled 范围错误值", "清除错误常量")
AddTableRow objTable, Array("12", "修复 #DIV/0! 除法公式", "包裹 IFERROR，消除零除错误")
AddTableRow objTable, Array("13", "修复 V0-合并表 E列 #N/A", "CHOOSE/MATCH 包裹 IFERROR")
AddTableRow objTable, Array("14", "设置 sheet 缩放比例", "所有 sheet 设置为 100%")
AddTableRow objTable, Array("15", "另存 V0-SunLife", "保存为""永明业绩数据-YYYYMMDD.xlsx""（公式转静态值）")
AddTableRow objTable, Array("16", "另存 V0-IYB业绩", "保存为""IYB业绩追踪周报-YYYYMMDD.xlsx""（V0+V2+匹配表）")
AddTableRow objTable, Array("17", "关闭 Excel", "退出 COM 进程")

' 2.5 Part 4: 生成转换规则 Word 文档
AddHeading objDoc, "2.5 Part 4: 生成转换规则 Word 文档", 2
AddParagraph objDoc, "处理函数：generate_rules_document(doc_path, date_str)"
AddParagraph objDoc, "生成内容："

AddBullet objDoc, "目录（TOC 字段）"
AddBullet objDoc, "V0 Sheet 转换规则：数据源、数据逻辑、列名映射、值域映射、数据格式化、格式设置"
AddBullet objDoc, "V0-NGP Sheet 转换规则：同上"
AddBullet objDoc, "共享映射规则汇总：保单状态、产品品类、币种、供款方式、产品名称映射"

' 2.6 Part 5: 生成结果验证报告
AddHeading objDoc, "2.6 Part 5: 生成结果验证报告", 2
AddParagraph objDoc, "处理函数：run_comparison() + generate_comparison_report()"
AddParagraph objDoc, "对比内容："

AddBullet objDoc, "数据量验证：各 sheet 行数、列数对比"
AddBullet objDoc, "字段一致性验证：逐列值一致率、不一致样本、映射规则验证"
AddBullet objDoc, "映射规则验证：列出所有映射规则及其执行效果"
AddBullet objDoc, "格式一致性验证：表头格式、数据行格式、列宽对比"

' 2.7 Part 6: 生成执行流程说明文档
AddHeading objDoc, "2.7 Part 6: 生成执行流程说明文档", 2
AddParagraph objDoc, "处理函数：generate_execution_manual_document(doc_path, date_str)"
AddParagraph objDoc, "生成内容：脚本执行流程说明文档，包含完整的执行步骤、映射规则、输出文件清单和执行流程图"

' ==================== 三、核心映射规则汇总 ====================
AddHeading objDoc, "三、核心映射规则汇总", 1

' 3.1 保单状态映射
AddHeading objDoc, "3.1 保单状态映射", 2
AddParagraph objDoc, "共24条映射规则，分为""保单状态""和""订单状态回填""两类"

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 3)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "原始值"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "V0标准值"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "类型"

AddTableRow objTable, Array("已生效", "生效", "保单状态")
AddTableRow objTable, Array("已生效-未回执", "生效", "保单状态")
AddTableRow objTable, Array("已生效-待核验", "生效", "保单状态")
AddTableRow objTable, Array("保单失效", "失效", "保单状态")
AddTableRow objTable, Array("已交单至保险公司", "已签单", "保单状态")
AddTableRow objTable, Array("PENDING", "pending", "保单状态")
AddTableRow objTable, Array("PENDING-待补充资料", "pending", "保单状态")
AddTableRow objTable, Array("PENDING-资料已补充待审核", "pending", "保单状态")
AddTableRow objTable, Array("PENDING-内部处理中", "pending", "保单状态")
AddTableRow objTable, Array("待核保", "待批核", "保单状态")
AddTableRow objTable, Array("待生效", "待批核", "保单状态")
AddTableRow objTable, Array("取消投保中", "取消投保", "保单状态")
AddTableRow objTable, Array("冷静期内退保", "退保", "保单状态")
AddTableRow objTable, Array("申请退保中", "退保", "保单状态")
AddTableRow objTable, Array("已撤销", "取消预约", "订单状态回填")
AddTableRow objTable, Array("预约成功", "排期", "订单状态回填")
AddTableRow objTable, Array("投保文件待复核", "pending", "订单状态回填")
AddTableRow objTable, Array("签单完成", "已签单", "订单状态回填")
AddTableRow objTable, Array("投保文件复核驳回", "拒保", "订单状态回填")
AddTableRow objTable, Array("已提交预约待审核", "pending", "订单状态回填")
AddTableRow objTable, Array("待交单至保险公司", "已签单", "订单状态回填")
AddTableRow objTable, Array("预约中", "排期", "订单状态回填")
AddTableRow objTable, Array("待确认转介信息", "pending", "订单状态回填")
AddTableRow objTable, Array("预约资料待修改", "pending", "订单状态回填")

' 3.2 产品品类映射
AddHeading objDoc, "3.2 产品品类映射", 2

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "细分品类"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "标准品类"

AddTableRow objTable, Array("危疾计划", "重疾")
AddTableRow objTable, Array("人寿保障", "人寿")
AddTableRow objTable, Array("医疗计划", "医疗")
AddTableRow objTable, Array("年金计划", "年金")
AddTableRow objTable, Array("高端医疗", "医疗")
AddTableRow objTable, Array("医疗计划(自愿医保)", "医疗")
AddTableRow objTable, Array("人寿险", "人寿")
AddTableRow objTable, Array("投资相连人寿保险计划", "投连险")
AddTableRow objTable, Array("其它类型", "医疗")
AddTableRow objTable, Array("意外及伤残", "医疗")

' 3.3 币种映射
AddHeading objDoc, "3.3 币种映射", 2

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "原始值"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "标准值"

AddTableRow objTable, Array("美元", "美金")
AddTableRow objTable, Array("港元", "港币")

' 3.4 供款方式映射
AddHeading objDoc, "3.4 供款方式映射", 2

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "原始值"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "标准值"

AddTableRow objTable, Array("整付保费", "整付")

' 3.5 IYB 端口映射
AddHeading objDoc, "3.5 IYB 端口映射", 2

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "业务细分"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "端口"

AddTableRow objTable, Array("BK业务", "三方机构")
AddTableRow objTable, Array("成事家办", "A端")
AddTableRow objTable, Array("天领业务", "A端")
AddTableRow objTable, Array("合伙转介业务", "B端")
AddTableRow objTable, Array("永明经代", "B端")
AddTableRow objTable, Array("同行经代", "B端")
AddTableRow objTable, Array("ICLUB业务", "B端")
AddTableRow objTable, Array("IFA业务", "B端")

' 3.6 IYB 分层映射
AddHeading objDoc, "3.6 IYB 分层映射", 2

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 2)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "市场分层"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "分层"

AddTableRow objTable, Array("银行网点", "银行")
AddTableRow objTable, Array("同行机构转介", "非持牌转介人")
AddTableRow objTable, Array("持牌转介", "非持牌转介人")
AddTableRow objTable, Array("异业机构转介", "非持牌转介人")
AddTableRow objTable, Array("同行个人转介", "非持牌转介人")
AddTableRow objTable, Array("持牌代理人", "非持牌转介人")
AddTableRow objTable, Array("持牌资管", "持牌转介人")
AddTableRow objTable, Array("持牌经纪", "持牌转介人")
AddTableRow objTable, Array("IYB分公司", "成事家办")
AddTableRow objTable, Array("天领机构", "天领")
AddTableRow objTable, Array("天领贴牌", "天领")
AddTableRow objTable, Array("IFA业务", "非持牌转介人")

' ==================== 四、输出文件清单 ====================
AddHeading objDoc, "四、输出文件清单", 1

Set objTable = objDoc.Tables.Add(objDoc.Content, 1, 3)
objTable.Style = "Table Grid"
objTable.Range.ParagraphFormat.Alignment = 1

SetCellFont objTable.Cell(1, 1), True, 10
objTable.Cell(1, 1).Range.Text = "文件名"
SetCellFont objTable.Cell(1, 2), True, 10
objTable.Cell(1, 2).Range.Text = "说明"
SetCellFont objTable.Cell(1, 3), True, 10
objTable.Cell(1, 3).Range.Text = "生成时机"

AddTableRow objTable, Array("业绩数据-整合-YYYYMMDD.xlsx", "主输出文件（含 V0、V0-NGP、V0-IYB业绩 及其他公式 sheet）", "Part 3")
AddTableRow objTable, Array("永明业绩数据-YYYYMMDD.xlsx", "V0-SunLife sheet 单独导出（公式转静态值）", "Part 3")
AddTableRow objTable, Array("IYB业绩追踪周报-YYYYMMDD.xlsx", "V0-IYB业绩 单独导出（含 V0+V2+匹配表）", "Part 3")
AddTableRow objTable, Array("业绩数据-整合-转换规则说明-YYYYMMDD.docx", "转换规则说明文档", "Part 4")
AddTableRow objTable, Array("业绩数据-整合-结果验证报告-YYYYMMDD.docx", "结果验证报告（对比参考文件）", "Part 5")
AddTableRow objTable, Array("业绩数据整合脚本执行流程说明-YYYYMMDD.docx", "脚本执行流程说明文档", "Part 6")

' ==================== 五、关键技术点 ====================
AddHeading objDoc, "五、关键技术点", 1

AddHeading objDoc, "5.1 公式保留策略", 2
AddParagraph objDoc, "通过 shutil.copy2 字节级复制参考文件，再用 Excel COM 写入数据并重算，确保公式引用结构完整保留"

AddHeading objDoc, "5.2 XLOOKUP 公式修复", 2
AddParagraph objDoc, "自动将参考模板中过小的查找范围（如 $J$1:$J$10）替换为完整列引用（J:J）"

AddHeading objDoc, "5.3 日期格式修复", 2
AddParagraph objDoc, "针对 spill sheet 中非日期列误设日期格式的问题，按列名判断并修复"

AddHeading objDoc, "5.4 端口/分层兜底", 2
AddParagraph objDoc, "即使 XLOOKUP 公式修复，仍对空值行直接写入映射值作为兜底"

AddHeading objDoc, "5.5 错误值处理", 2
AddParagraph objDoc, "清理 #N/A、#VALUE!、#DIV/0! 等错误值，修复除法公式零除问题"

AddHeading objDoc, "5.6 数据分块写入", 2
AddParagraph objDoc, "COM 写入采用 chunk=1000 分块策略，避免内存溢出"

AddHeading objDoc, "5.7 进程清理", 2
AddParagraph objDoc, "写入前强制关闭残留 Excel 进程，避免文件锁定"

' ==================== 六、执行流程图 ====================
AddHeading objDoc, "六、执行流程图", 1

AddParagraph objDoc, "开始" & vbCrLf & _
    "  │" & vbCrLf & _
    "  ▼" & vbCrLf & _
    "┌─────────────────────────────────────────────┐" & vbCrLf & _
    "│ 1. 读取输入参数                              │" & vbCrLf & _
    "│    - 综合查询文件                            │" & vbCrLf & _
    "│    - NGP文件                                 │" & vbCrLf & _
    "│    - 映射表（自动查找）                       │" & vbCrLf & _
    "│    - 参考文件（自动查找）                     │" & vbCrLf & _
    "└─────────────────────────────────────────────┘" & vbCrLf & _
    "  │" & vbCrLf & _
    "  ├──────────────────────┐" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 1: V0 数据处理           │" & vbCrLf & _
    "  │  │ process_v0_data()             │" & vbCrLf & _
    "  │  │ 7个步骤 → 80列 V0 格式        │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 2: V0-NGP 数据处理       │" & vbCrLf & _
    "  │  │ process_v0ngp_data()          │" & vbCrLf & _
    "  │  │ 5个步骤 → 39列 V0-NGP 格式    │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 2c: V0-IYB业绩 数据生成  │" & vbCrLf & _
    "  │  │ process_iyb_v0_data()         │" & vbCrLf & _
    "  │  │ 筛选IYB供应商 → 20列格式      │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 3: 创建输出文件           │" & vbCrLf & _
    "  │  │ - openpyxl 复制参考文件        │" & vbCrLf & _
    "  │  │ - Excel COM 写入数据并重算     │" & vbCrLf & _
    "  │  │ - 修复公式和格式              │" & vbCrLf & _
    "  │  │ - 另存永明/IYB业绩文件        │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 4: 生成转换规则文档       │" & vbCrLf & _
    "  │  │ generate_rules_document()     │" & vbCrLf & _
    "  │  │ Word文档：映射规则汇总         │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 5: 生成结果验证报告       │" & vbCrLf & _
    "  │  │ run_comparison()              │" & vbCrLf & _
    "  │  │ 对比参考文件，生成验证报告     │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  │                      ▼" & vbCrLf & _
    "  │  ┌───────────────────────────────┐" & vbCrLf & _
    "  │  │ Part 6: 生成执行流程说明文档   │" & vbCrLf & _
    "  │  │ generate_execution_manual_document() │" & vbCrLf & _
    "  │  │ Word文档：脚本执行流程说明     │" & vbCrLf & _
    "  │  └───────────────────────────────┘" & vbCrLf & _
    "  │                      │" & vbCrLf & _
    "  ▼                      ▼" & vbCrLf & _
    "完成" & vbCrLf & _
    "  │" & vbCrLf & _
    "  └── 输出文件：业绩数据-整合-YYYYMMDD.xlsx" & vbCrLf & _
    "      转换规则说明.docx" & vbCrLf & _
    "      结果验证报告.docx" & vbCrLf & _
    "      永明业绩数据-YYYYMMDD.xlsx" & vbCrLf & _
    "      IYB业绩追踪周报-YYYYMMDD.xlsx" & vbCrLf & _
    "      业绩数据整合脚本执行流程说明-YYYYMMDD.docx"

objDoc.Content.Font.Name = "宋体"
objDoc.Content.Font.Size = 9

' 保存文档
objDoc.SaveAs(strOutputPath)
objDoc.Close()
objWord.Quit()

Set objCell = Nothing
Set objRow = Nothing
Set objTable = Nothing
Set objRange = Nothing
Set objDoc = Nothing
Set objWord = Nothing

WScript.Echo "✅ Word 文档已生成: " & strOutputPath


' ==================== 辅助函数 ====================

Sub AddHeading(objDoc, strText, intLevel)
    Dim objPara
    Set objPara = objDoc.Content.Paragraphs.Add()
    objPara.Range.Text = strText
    objPara.Style = objDoc.Styles("Heading " & intLevel)
    objPara.Range.Font.Name = "宋体"
    objPara.Range.Font.Size = 11 + (1 - intLevel) * 2
    objPara.Range.Font.Bold = True
    objPara.Range.Font.Color = RGB(26, 95, 180)
    objPara.Range.InsertParagraphAfter
End Sub

Sub AddParagraph(objDoc, strText)
    Dim objPara
    Set objPara = objDoc.Content.Paragraphs.Add()
    objPara.Range.Text = strText
    objPara.Range.Font.Name = "宋体"
    objPara.Range.Font.Size = 11
    objPara.Range.InsertParagraphAfter
End Sub

Sub AddBullet(objDoc, strText)
    Dim objPara
    Set objPara = objDoc.Content.Paragraphs.Add()
    objPara.Range.Text = strText
    objPara.Range.Font.Name = "宋体"
    objPara.Range.Font.Size = 11
    objPara.Range.ListFormat.ApplyBulletDefault
    objPara.Range.InsertParagraphAfter
End Sub

Sub SetCellFont(objCell, blnBold, intSize)
    objCell.Range.Font.Name = "宋体"
    objCell.Range.Font.Size = intSize
    objCell.Range.Font.Bold = blnBold
    objCell.VerticalAlignment = 1 ' wdCellAlignVerticalCenter
    objCell.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
End Sub

Sub AddTableRow(objTable, arrValues)
    Dim intI, objRow, objCell
    Set objRow = objTable.Rows.Add()
    For intI = 0 To UBound(arrValues)
        Set objCell = objRow.Cells(intI + 1)
        objCell.Range.Text = arrValues(intI)
        objCell.Range.Font.Name = "宋体"
        objCell.Range.Font.Size = 9
        objCell.VerticalAlignment = 1
        objCell.Range.ParagraphFormat.Alignment = 1
    Next
End Sub
