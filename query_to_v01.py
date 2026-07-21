#!/usr/bin/env python3
"""
综合查询结果 + V0-NGP-1 → V0-1 Sheet 自动转换脚本

功能：
  - 读取综合查询结果 xlsx 文件，删除签单供应商=永领致远顾问有限公司的数据
  - 繁简转换 + 产品名称映射（可选）
  - 保单状态映射（综合查询原始值 → V0标准简写值）
  - 保单号码格式与V0一致（纯数字→int无前缀零，非纯数字→str）
  - 将综合查询列名映射为 V0 格式的列名（80列）
  - 从业绩数据-整合文件读取 V0-NGP-1 数据，追加在末尾
  - 直接在整合文件中新增 V0-1 sheet（不另存新文件）

版本：2.5
更新日志：
  - 2026-07-16: v2.5 供款方式映射（缴费方式"整付保费"→"整付"），与V0值域一致；供款方式"0"→NaN
  - 2026-07-16: v2.4 电话列.0后缀修复（float→str残留，如62982319.0→62982319）
  - 2026-07-16: v2.3 数据行格式与V0完全一致（per-column精细设置：39列styled+41列plain+2列特殊）
  - 2026-07-16: v2.2 产品品类映射（11种细分品类→V0标准8种），与V0值域一致
  - 2026-07-16: v2.0 表头样式与V0完全一致（bold/wrap_text/border逐列设置）
  - 2026-07-16: v1.9 保单号码增强：str类型.0后缀(float→str残留)去掉.0转为int
  - 2026-07-16: v1.8 年期=整付保费→年期=1，且对应行供款方式=整付
  - 2026-07-16: v1.7 币种映射（美元→美金，港元→港币），与V0命名一致
  - 2026-07-16: v1.6 年期列去"年"字（如"25年"→"25"，与V0格式一致）
  - 2026-07-15: v1.5 V0-1全部繁体→简体（列名+数据值），含V0-NGP-1列名繁简转换
  - 2026-07-15: v1.4 预约签单时间→签单日期（匹配V0列结构），不再映射到签单时间
  - 2026-07-15: v1.3 保单号码格式与V0一致（纯数字→int/General，非纯数字→str/@）
  - 2026-07-15: v1.2 添加保单状态映射规则（综合查询原始值→V0标准简写值）
  - 2026-07-15: v1.1 改为直接在整合文件中新增 V0-1 sheet，不再另存新文件
  - 2026-07-15: v1.0 初版
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import numpy as np
import sys
import os
from pathlib import Path
from datetime import datetime

# ============================================================
# 常量定义
# ============================================================

# V0 列顺序（80列）
V0_COLUMNS = [
    '简称', '订单编号', '提交日期', '转介公司', '转介人', '转介日期', '转介时间',
    '保单状态', '签单日期', '签单时间', '签单地点', '签单供应商', '注册编号IA',
    'TR', 'TR助理', '业务行政', '业务细分', '市场分层', 'KEY ACCOUNT',
    '机构公司名字', '合作伙伴', '赴港联系人', '电话', '预约备注', '所签地区',
    '保险公司', '产品品类', '产品名称', '年期', '供款方式', '币种',
    '计划书年龄', '保费', '保费（港币）', 'APE', '保监征费', '合计', '保额',
    '投保人  (中文)', '投保人  (拼音)', '投保人国籍', '投保人证件号',
    '投保人出生日期', '投保人电话', '投保人职业', '投保人邮箱',
    '投保人邮寄通讯地址', '受保人(中文）', '受保人(拼音)', '受保人证件号',
    '申请表递交状态', '递交日期', 'pending原因', '查单最新更新日期',
    '首期付款方式', '是否第三者付款', '是否融资单', '客户分群',
    '首年特殊折扣', '保单号码', '批核日（年/月/日）',
    '生效日期（年/月/日）', '保费到期日（年/月/日）',
    '首期保费日（年/月/日）', '预计冷静期截止日',
    '快递单号',     'DDA状态', '回执状况', '保费逾期未缴',
    '续保状态', '续保金额', '保费储备金户口', '是否转入单',
    '备注', '订单编号（旧）', '佣金模式',
    '计划书编号(仅平安需登记）', '港分经理（公式读取）',
    'IS老师（公式读取）', '预约类型（公式读取）'
]

# 综合查询列名 → V0列名 映射
QUERY_TO_V0_MAP = {
    '订单编号': '订单编号',
    '预约提交日期': '提交日期',
    '预约签单时间': '签单日期',
    '保单状态': '保单状态',
    '申请表递交状态': '申请表递交状态',
    '签单供应商': '签单供应商',
    '保单号码': '保单号码',
    '批核日（年/月/日）': '批核日（年/月/日）',
    '业务细分': '业务细分',
    '市场分层': '市场分层',
    'KEY ACCOUNT': 'KEY ACCOUNT',
    '公司名称': '机构公司名字',
    '合作伙伴': '合作伙伴',
    '保险公司': '保险公司',
    '产品名称': '产品名称',
    '供款年限': '年期',
    '缴费方式': '供款方式',
    '保单币种': '币种',
    '首年保费': '保费',
    '保费（港币）': '保费（港币）',
    'APE（港币）': 'APE',
    '投保人姓名（中文）': '投保人  (中文)',
    '受保人名字（中文）': '受保人(中文）',
    '计划书年龄': '计划书年龄',
    '投保人国籍': '投保人国籍',
    '是否融资单': '是否融资单',
    '折扣比例': '首年特殊折扣',
    '转介公司': '转介公司',
    '所签地区': '所签地区',
    '产品类型': '产品品类',
    'TR': 'TR',
    '订单编号（旧）': '订单编号（旧）',
    '递交保险公司日期': '递交日期',
}

# 综合查询中需要丢弃的列（不在V0中）
QUERY_DROP_COLS = [
    '订单状态', '是否预缴', '是否专业投资者',
    '结算模式', '订单子状态',
]

# 要删除的签单供应商
DELETE_SUPPLIER = '永领致远顾问有限公司'

# 币种映射：综合查询/V0-NGP-1 原始值 → V0 标准值
CURRENCY_MAP = {
    '美元': '美金',
    '港元': '港币',
}

# 产品品类映射：综合查询/V0-NGP-1 原始值 → V0 标准值
PRODUCT_CATEGORY_MAP = {
    # 综合查询产品类型/V0-NGP-1产品品类 细分值 → V0标准值
    '危疾计划': '重疾',
    '人寿保障': '人寿',
    '医疗计划': '医疗',
    '年金计划': '年金',
    '高端医疗': '医疗',
    '医疗计划(自愿医保)': '医疗',
    '人寿险': '人寿',
    '投资相连人寿保险计划': '投连险',
    '其它类型': '医疗',
    '意外及伤残': '医疗',
}

# 供款方式映射：综合查询缴费方式/V0-NGP-1原始值 → V0 标准值
PAYMENT_METHOD_MAP = {
    '整付保费': '整付',
}

# 保单状态映射：综合查询/V0-NGP-1 原始值 → V0 标准值
POLICY_STATUS_MAP = {
    # 综合查询保单状态原始值 → V0简写值
    '已生效': '生效',
    '已生效-未回执': '生效',
    '已生效-待核验': '生效',
    '保单失效': '失效',
    '已交单至保险公司': '已签单',
    'PENDING': 'pending',
    'PENDING-待补充资料': 'pending',
    'PENDING-资料已补充待审核': 'pending',
    'PENDING-内部处理中': 'pending',
    '待核保': '待批核',
    '待生效': '待批核',
    '取消投保中': '取消投保',
    '冷静期内退保': '退保',
    '申请退保中': '退保',
    # 订单状态回填值 → V0简写值（保单状态为空时使用订单状态）
    '已撤销': '取消预约',
    '预约成功': '排期',
    '投保文件待复核': 'pending',
    '签单完成': '已签单',
    '投保文件复核驳回': '拒保',
    '已提交预约待审核': 'pending',
    '待交单至保险公司': '已签单',
    '预约中': '排期',
    '待确认转介信息': 'pending',
    '预约资料待修改': 'pending',
}


# 日期列
DATE_COLUMNS = [
    '提交日期', '签单日期', '签单时间', '转介日期', '转介时间',
    '批核日（年/月/日）', '生效日期（年/月/日）', '保费到期日（年/月/日）',
    '首期保费日（年/月/日）', '预计冷静期截止日', '递交日期',
    '查单最新更新日期', '投保人出生日期',
]

# 金额列（number_format = #,##0.00_ )
AMOUNT_COLUMNS = ['保费', '保费（港币）', 'APE', '合计']

# 文本列（number_format = @）
TEXT_FORMAT_COLUMNS = [
    '订单编号', '签单地点', '签单供应商', '注册编号IA',
    'TR', 'TR助理', '业务行政', '赴港联系人', '预约备注',
    '产品品类', '产品名称', '供款方式', '币种',
    '投保人  (中文)', '投保人  (拼音)', '投保人国籍', '投保人证件号',
    '投保人电话', '投保人职业', '投保人邮箱', '投保人邮寄通讯地址',
    '受保人(中文）', '受保人(拼音)', '受保人证件号',
    '首期付款方式', '是否第三者付款', '是否融资单',
    '首年特殊折扣', '保额',
]

# 日期格式列（number_format = yyyy-mm-dd）
DATE_FORMAT_COLUMNS = [
    '提交日期', '签单日期', '签单时间', '转介日期', '转介时间',
    '批核日（年/月/日）', '生效日期（年/月/日）', '保费到期日（年/月/日）',
    '首期保费日（年/月/日）', '预计冷静期截止日', '递交日期',
    '查单最新更新日期', '投保人出生日期',
]

# V0 表头颜色分组
HEADER_COLOR_GROUPS = {
    'FFC55A11': [  # 橙色 - 主要字段
        '简称', '订单编号', '提交日期', '签单日期', '签单时间', '签单地点',
        '签单供应商', '注册编号IA', 'TR', 'TR助理', '业务行政', '业务细分',
        '市场分层', 'KEY ACCOUNT', '机构公司名字', '合作伙伴', '赴港联系人',
        '电话', '预约备注', '所签地区', '保险公司', '产品品类', '产品名称',
        '年期', '供款方式', '币种', '计划书年龄', '保费', '保监征费', '保额',
        '投保人  (中文)', '投保人  (拼音)', '投保人国籍', '投保人证件号',
        '投保人出生日期', '投保人电话', '投保人职业', '投保人邮箱',
        '投保人邮寄通讯地址', '受保人(中文）', '受保人(拼音)', '受保人证件号',
        '申请表递交状态', '递交日期', '首期付款方式', '是否第三者付款',
        '是否融资单', '客户分群', '首年特殊折扣', '保单号码',
        '计划书编号(仅平安需登记）', '港分经理（公式读取）',
        'IS老师（公式读取）', '预约类型（公式读取）',
    ],
    'FF548235': [  # 绿色 - 转介 + 金额
        '转介公司', '转介人', '转介日期', '转介时间',
        '保费（港币）', 'APE', '合计', '订单编号（旧）', '佣金模式',
    ],
    'FF7030A0': [  # 紫色 - 状态/备注
        '保单状态', 'pending原因', '查单最新更新日期', '备注',
    ],
    'FF2E75B6': [  # 蓝色 - 保单周期
        '批核日（年/月/日）', '生效日期（年/月/日）', '保费到期日（年/月/日）',
        '首期保费日（年/月/日）', '预计冷静期截止日', '快递单号',
        'DDA状态', '回执状况', '保费逾期未缴', '续保状态', '续保金额',
        '保费储备金户口', '是否转入单',
    ],
}

# V0 列宽
V0_COLUMN_WIDTHS = {
    '简称': 11.0, '订单编号': 22.7, '提交日期': 14.14,
    '转介公司': 18.59, '转介人': 11.38, '转介日期': 14.38, '转介时间': 17.14,
    '保单状态': 15.87, '签单日期': 12.5, '签单时间': 14.25,
    '签单地点': 16.38, '签单供应商': 29.38, '注册编号IA': 15.0,
    'TR': 14.14, 'TR助理': 15.63, '业务行政': 14.38,
    '业务细分': 12.87, '市场分层': 15.0, 'KEY ACCOUNT': 22.75,
    '机构公司名字': 25.5, '合作伙伴': 32.74, '赴港联系人': 10.62,
    '电话': 13.0, '预约备注': 10.62, '所签地区': 8.1,
    '保险公司': 30.12, '产品品类': 13.75, '产品名称': 25.0,
    '年期': 10.14, '供款方式': 9.5, '币种': 12.38,
    '计划书年龄': 12.75, '保费': 14.19, '保费（港币）': 19.0,
    'APE': 16.75, '保监征费': 12.0, '合计': 17.87, '保额': 19.0,
    '投保人  (中文)': 24.25, '投保人  (拼音)': 33.25,
    '投保人国籍': 22.14, '投保人证件号': 25.16,
    '投保人出生日期': 22.25, '投保人电话': 20.63,
    '投保人职业': 28.0, '投保人邮箱': 34.38,
    '投保人邮寄通讯地址': 64.5, '受保人(中文）': 26.63,
    '受保人(拼音)': 23.25, '受保人证件号': 18.63,
    '申请表递交状态': 17.5, '递交日期': 18.61,
    'pending原因': 37.25, '查单最新更新日期': 30.87,
    '首期付款方式': 15.0, '是否第三者付款': 16.87,
    '是否融资单': 21.0, '客户分群': 21.14,
    '首年特殊折扣': 26.75, '保单号码': 25.63,
    '批核日（年/月/日）': 24.63, '生效日期（年/月/日）': 22.14,
    '保费到期日（年/月/日）': 24.14, '首期保费日（年/月/日）': 26.38,
    '预计冷静期截止日': 33.25, '快递单号': 50.46,
    'DDA状态': 21.0, '回执状况': 16.87, '保费逾期未缴': 15.87,
    '续保状态': 36.14, '续保金额': 23.38, '保费储备金户口': 17.81,
    '是否转入单': 20.16, '备注': 39.25,
    '订单编号（旧）': 27.35, '佣金模式': 10.78,
    '计划书编号(仅平安需登记）': 9.0,
    '港分经理（公式读取）': 13.0, 'IS老师（公式读取）': 13.0,
    '预约类型（公式读取）': 13.0,
}

# ============================================================
# 繁简转换 & 产品名称映射
# ============================================================

def load_product_name_mapping(mapping_table_path):
    """从映射表加载产品名称映射规则。"""
    if not Path(mapping_table_path).exists():
        print(f"   ⚠️ 映射表不存在: {mapping_table_path}，跳过产品名称映射")
        return {}

    df_map = pd.read_excel(mapping_table_path, sheet_name='产品名称匹配')
    name_map = {}
    for _, row in df_map.iterrows():
        raw_name = str(row['旧产品名称']).strip()
        std_name = str(row['新产品名称']).strip()
        if raw_name and std_name and raw_name != std_name:
            name_map[raw_name] = std_name
    return name_map


def traditional_to_simplified(df, skip_cols=None):
    """对DataFrame中的文本列进行繁体→简体转换。"""
    try:
        from opencc import OpenCC
        cc = OpenCC('t2s')
    except ImportError:
        print("   ⚠️ opencc 未安装，跳过繁简转换")
        return df

    if skip_cols is None:
        skip_cols = []

    converted_count = {}
    for c in df.columns:
        if c in skip_cols:
            continue
        # 检查是否是文本列
        if df[c].dtype == 'object' or pd.api.types.is_string_dtype(df[c]):
            # 进一步跳过数值型伪装成文本的列
            if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_string_dtype(df[c]):
                continue
            old_vals = df[c].fillna('').astype(str)
            new_vals = old_vals.apply(lambda x: cc.convert(x))
            diff_count = (old_vals != new_vals).sum()
            if diff_count > 0:
                converted_count[c] = diff_count
                # 显示前3个转换示例
                examples = []
                for idx in old_vals.index[old_vals != new_vals][:3]:
                    examples.append(f"     {old_vals[idx]} → {new_vals[idx]}")
                print(f"   '{c}' 有 {diff_count} 个值发生了繁简转换")
                for ex in examples:
                    print(ex)
            df[c] = new_vals

    if converted_count:
        print(f"   繁简转换总计: {sum(converted_count.values())} 个值，涉及 {len(converted_count)} 列")
    else:
        print("   无繁简转换需要处理")
    return df


def apply_product_name_mapping(df, name_map):
    """对产品名称列应用映射。"""
    if not name_map:
        return df
    if '产品名称' not in df.columns:
        return df

    old_names = df['产品名称'].fillna('').astype(str)
    new_names = old_names.apply(lambda x: name_map.get(x.strip(), x.strip()))
    matched = (old_names != new_names).sum()
    if matched > 0:
        print(f"   产品名称映射: 匹配到标准名 {matched} 个")
        for old, new in name_map.items():
            count = (old_names == old).sum()
            if count > 0:
                print(f"     {old} → {new} ({count}条)")
    else:
        print(f"   产品名称映射: 无匹配")
    df['产品名称'] = new_names
    return df


# ============================================================
# 核心转换函数
# ============================================================

def convert_query_to_v01(query_file_path, integrate_file_path,
                         mapping_table_path=None, auto_find_mapping=True):
    """将综合查询数据 + V0-NGP-1 数据合并为 V0-1 sheet，直接写入整合文件。

    Args:
        query_file_path: 综合查询结果 xlsx 文件路径
        integrate_file_path: 业绩数据-整合文件路径（含 V0-NGP-1 sheet），直接在此文件新增 V0-1
        mapping_table_path: 产品名称映射表路径（可选，默认自动查找同目录下的业务部门字段映射表.xlsx）
        auto_find_mapping: 是否自动查找映射表（默认True，当mapping_table_path为None时生效）

    Returns:
        integrate_file_path: 修改后的文件路径
    """
    # 自动查找映射表：优先查找整合文件目录，其次查找综合查询文件目录
    if mapping_table_path is None and auto_find_mapping:
        # 优先从整合文件所在目录查找
        integrate_dir = Path(integrate_file_path).parent
        candidate = integrate_dir / '业务部门字段映射表.xlsx'
        if candidate.exists():
            mapping_table_path = str(candidate)
            print(f"📂 自动找到映射表: {mapping_table_path}")
        else:
            # 再从综合查询文件目录查找
            query_dir = Path(query_file_path).parent
            candidate = query_dir / '业务部门字段映射表.xlsx'
            if candidate.exists():
                mapping_table_path = str(candidate)
                print(f"📂 自动找到映射表: {mapping_table_path}")

    print()
    print("=" * 60)
    print("综合查询 + V0-NGP-1 → V0-1 转换脚本")
    print("=" * 60)
    print()

    print(f"📂 综合查询文件: {query_file_path}")
    print(f"📂 整合文件: {integrate_file_path}")
    if mapping_table_path:
        print(f"📂 映射表: {mapping_table_path}")
    else:
        print(f"📂 映射表: 未提供（繁简转换和产品名称映射将跳过）")
    print()

    # ---- Step 1: 读取综合查询数据 ----
    print("📖 Step 1: 读取综合查询数据...")
    df_query = pd.read_excel(query_file_path)
    print(f"   原始数据: {df_query.shape[0]} 行 x {df_query.shape[1]} 列")

    # ---- Step 2: 删除签单供应商=永领致远的数据 ----
    print()
    print("🗑️ Step 2: 删除签单供应商=永领致远顾问有限公司的数据...")
    delete_count = df_query[df_query['签单供应商'] == DELETE_SUPPLIER].shape[0]
    df_query = df_query[df_query['签单供应商'] != DELETE_SUPPLIER].copy()
    print(f"   删除 {delete_count} 行，剩余 {df_query.shape[0]} 行")

    # ---- Step 2b: 保单状态回填（保单状态优先，为空时用订单状态） ----
    print()
    print("🔄 Step 2b: 保单状态回填（保单状态优先，为空时回填订单状态）...")
    if '保单状态' in df_query.columns and '订单状态' in df_query.columns:
        # 判断保单状态为空的条件：NaN / 空字符串 / "0"
        bao_empty_mask = df_query['保单状态'].isna() | \
                         (df_query['保单状态'].astype(str).str.strip() == '') | \
                         (df_query['保单状态'].astype(str).str.strip() == '0')
        # 订单状态有值且非空
        ding_valid_mask = df_query['订单状态'].notna() & \
                         (df_query['订单状态'].astype(str).str.strip() != '') & \
                         (df_query['订单状态'].astype(str).str.strip() != '0')
        # 需要回填的行：保单状态为空 且 订单状态有值
        fallback_mask = bao_empty_mask & ding_valid_mask
        fallback_count = fallback_mask.sum()
        if fallback_count > 0:
            from collections import Counter
            fallback_details = Counter(df_query.loc[fallback_mask, '订单状态'].astype(str).str.strip())
            print(f"   回填 {fallback_count} 行：保单状态为空 → 使用订单状态值")
            for val, cnt in fallback_details.most_common():
                print(f"     订单状态=\"{val}\" → 保单状态 ({cnt}条)")
            # 执行回填
            df_query.loc[fallback_mask, '保单状态'] = df_query.loc[fallback_mask, '订单状态'].astype(str).str.strip()
        else:
            print("   无需回填的行（保单状态均有值，或订单状态也为空）")
        # 统计保单状态仍然为空的行
        still_empty = bao_empty_mask.sum() - fallback_count
        if still_empty > 0:
            print(f"   保单状态仍为空: {still_empty} 行（订单状态也为空）")
        print(f"   回填后保单状态值域: {sorted([str(x) for x in df_query['保单状态'].dropna().unique()])}")
    else:
        print("   ⚠️ 无保单状态或订单状态列，跳过回填")

    # ---- Step 3: 繁简转换（始终执行） ----
    print()
    print("🔤 Step 3: 繁体→简体转换（综合查询数据）...")
    # 确定需要跳过的列（日期列、数值列）
    skip_cols_query = set()
    for c in df_query.columns:
        try:
            if pd.api.types.is_numeric_dtype(df_query[c]) and not pd.api.types.is_string_dtype(df_query[c]):
                skip_cols_query.add(c)
        except Exception:
            pass
    df_query = traditional_to_simplified(df_query, skip_cols=skip_cols_query)

    # ---- Step 4: 产品名称映射（可选） ----
    print()
    print("🏷️ Step 4: 产品名称按映射表清洗（综合查询数据）...")
    if mapping_table_path:
        name_map = load_product_name_mapping(mapping_table_path)
        df_query = apply_product_name_mapping(df_query, name_map)
    else:
        print("   未提供映射表，跳过产品名称映射")

    # ---- Step 5: 综合查询列名映射为 V0 格式 ----
    print()
    print("🔄 Step 5: 综合查询列名映射为 V0 格式 (38列 → 80列)...")

    # 先丢弃不需要的列
    for col in QUERY_DROP_COLS:
        if col in df_query.columns:
            df_query = df_query.drop(columns=[col])

    # 重命名列
    rename_dict = {}
    for old_name, new_name in QUERY_TO_V0_MAP.items():
        if old_name in df_query.columns:
            rename_dict[old_name] = new_name
    df_query = df_query.rename(columns=rename_dict)

    # 添加 V0 中缺失的列（留空）
    for col in V0_COLUMNS:
        if col not in df_query.columns:
            df_query[col] = np.nan

    # 按 V0 列顺序重排
    df_query = df_query[V0_COLUMNS]
    print(f"   映射完成: {df_query.shape[1]} 列（按V0顺序）")
    print(f"   综合查询有效数据行: {df_query.shape[0]} 行")

    # ---- Step 6: 读取 V0-NGP-1 数据 ----
    print()
    print("📖 Step 6: 读取 V0-NGP-1 数据...")
    df_ngp1 = pd.read_excel(integrate_file_path, sheet_name='V0-NGP-1')
    print(f"   V0-NGP-1 数据: {df_ngp1.shape[0]} 行 x {df_ngp1.shape[1]} 列")

    # ---- Step 6b: V0-NGP-1 列名繁简转换 ----
    print()
    print("🔤 Step 6b: V0-NGP-1 列名繁体→简体转换...")
    try:
        from opencc import OpenCC
        cc_col = OpenCC('t2s')
        rename_col_dict = {}
        for col in df_ngp1.columns:
            simp_col = cc_col.convert(str(col))
            if str(col) != simp_col:
                rename_col_dict[col] = simp_col
        if rename_col_dict:
            print(f"   列名繁简转换: {len(rename_col_dict)} 个列名")
            for old, new in rename_col_dict.items():
                print(f"     {old} -> {new}")
            df_ngp1 = df_ngp1.rename(columns=rename_col_dict)
        else:
            print("   无需转换的列名")
    except ImportError:
        print("   ⚠️ opencc 未安装，跳过列名繁简转换")

    # ---- Step 7: V0-NGP-1 繁简转换（始终执行） ----
    print()
    print("🔤 Step 7: 繁体→简体转换（V0-NGP-1 数据）...")
    skip_cols_ngp = set(DATE_COLUMNS)
    for c in df_ngp1.columns:
        try:
            if pd.api.types.is_numeric_dtype(df_ngp1[c]) and not pd.api.types.is_string_dtype(df_ngp1[c]):
                skip_cols_ngp.add(c)
        except Exception:
            pass
    df_ngp1 = traditional_to_simplified(df_ngp1, skip_cols=skip_cols_ngp)

    print()
    print("🏷️ Step 7b: 产品名称按映射表清洗（V0-NGP-1 数据）...")
    if mapping_table_path:
        df_ngp1 = apply_product_name_mapping(df_ngp1, name_map)
    else:
        print("   未提供映射表，跳过产品名称映射")

    # ---- Step 8: V0-NGP-1 列映射为 V0 格式 ----
    print()
    print("🔄 Step 8: V0-NGP-1 列映射为 V0 格式 (39列 → 80列)...")

    # V0-NGP-1 的列名已与 V0 一致，只需添加缺失列
    for col in V0_COLUMNS:
        if col not in df_ngp1.columns:
            df_ngp1[col] = np.nan

    # 按 V0 列顺序重排
    df_ngp1 = df_ngp1[V0_COLUMNS]
    print(f"   映射完成: {df_ngp1.shape[1]} 列（按V0顺序）")
    print(f"   V0-NGP-1 有效数据行: {df_ngp1.shape[0]} 行")

    # ---- Step 9: 合并数据 ----
    print()
    print("🔗 Step 9: 合并数据（综合查询在前，V0-NGP-1 在末尾）...")
    df_combined = pd.concat([df_query, df_ngp1], ignore_index=True)
    print(f"   合计: {df_combined.shape[0]} 行 x {df_combined.shape[1]} 列")
    print(f"   其中综合查询数据: {df_query.shape[0]} 行")
    print(f"   其中V0-NGP-1数据: {df_ngp1.shape[0]} 行")

    # ---- Step 10: 保单状态映射 ----
    print()
    print("🔄 Step 10: 保单状态映射（原始值 → V0标准值）...")
    if '保单状态' in df_combined.columns:
        old_statuses = df_combined['保单状态'].fillna('').astype(str)
        new_statuses = old_statuses.apply(
            lambda x: POLICY_STATUS_MAP.get(x.strip(), x.strip()) if x.strip() and x.strip() != '0' else x
        )
        mapped_count = (old_statuses != new_statuses).sum()
        if mapped_count > 0:
            print(f"   映射了 {mapped_count} 个保单状态值")
            # 统计映射详情
            from collections import Counter
            mapping_details = Counter()
            for old, new in zip(old_statuses, new_statuses):
                if old != new:
                    mapping_details[(old, new)] += 1
            for (old, new), cnt in mapping_details.most_common():
                print(f"     \"{old}\" → \"{new}\" ({cnt}条)")
            df_combined['保单状态'] = new_statuses
        else:
            print("   无需映射的保单状态值")
        # 保单状态值 '0' → NaN（来自V0-NGP-1的空值）
        mask_zero = df_combined['保单状态'].astype(str).str.strip() == '0'
        if mask_zero.sum() > 0:
            print(f"   保单状态值\"0\" → 空值 ({mask_zero.sum()}条)")
            df_combined.loc[mask_zero, '保单状态'] = np.nan
        # 打印映射后的值域
        mapped_statuses = df_combined['保单状态'].dropna().unique()
        print(f"   映射后保单状态值域: {sorted([str(x) for x in mapped_statuses])}")
    else:
        print("   ⚠️ 无保单状态列，跳过映射")

    # ---- Step 10b: 数据清洗 ----
    print()
    print("🧹 Step 10b: 数据清洗...")
    # 产品品类映射（细分品类→V0标准值）
    if '产品品类' in df_combined.columns:
        old_category = df_combined['产品品类'].fillna('').astype(str)
        new_category = old_category.apply(
            lambda x: PRODUCT_CATEGORY_MAP.get(x.strip(), x) if x.strip() and x.strip() != '0' else x
        )
        category_mapped = (old_category != new_category).sum()
        if category_mapped > 0:
            print(f"   产品品类映射: {category_mapped}个值")
            from collections import Counter
            category_details = Counter()
            for old, new in zip(old_category, new_category):
                if old != new and old.strip() != '0':
                    category_details[(old, new)] += 1
            for (old, new), cnt in category_details.most_common():
                print(f"     \"{old}\" → \"{new}\" ({cnt}条)")
            df_combined['产品品类'] = new_category
        else:
            print("   无需映射的产品品类值")
        # 产品品类值 '0' → NaN（来自V0-NGP-1的空值）
        mask_cat_zero = df_combined['产品品类'].astype(str).str.strip() == '0'
        if mask_cat_zero.sum() > 0:
            print(f"   产品品类值\"0\" → 空值 ({mask_cat_zero.sum()}条)")
            df_combined.loc[mask_cat_zero, '产品品类'] = np.nan
        # 打印映射后值域
        category_vals = df_combined['产品品类'].dropna().unique()
        print(f"   映射后产品品类值域: {sorted([str(x) for x in category_vals])}")

    # 币种映射（美元→美金，港元→港币）
    if '币种' in df_combined.columns:
        old_currency = df_combined['币种'].fillna('').astype(str)
        new_currency = old_currency.apply(lambda x: CURRENCY_MAP.get(x.strip(), x) if x.strip() and x.strip() != '0' else x)
        currency_mapped = (old_currency != new_currency).sum()
        if currency_mapped > 0:
            print(f"   币种映射: {currency_mapped}个值")
            from collections import Counter
            currency_details = Counter()
            for old, new in zip(old_currency, new_currency):
                if old != new and old.strip() != '0':
                    currency_details[(old, new)] += 1
            for (old, new), cnt in currency_details.most_common():
                print(f"     \"{old}\" → \"{new}\" ({cnt}条)")
            df_combined['币种'] = new_currency
        # 币种值 '0' → NaN
        mask_currency_zero = df_combined['币种'].astype(str).str.strip() == '0'
        if mask_currency_zero.sum() > 0:
            print(f"   币种值\"0\" → 空值 ({mask_currency_zero.sum()}条)")
            df_combined.loc[mask_currency_zero, '币种'] = np.nan
        # 打印映射后值域
        currency_vals = df_combined['币种'].dropna().unique()
        print(f"   映射后币种值域: {sorted([str(x) for x in currency_vals])}")

    # 供款方式映射（缴费方式"整付保费"→"整付"，与V0值域一致）
    if '供款方式' in df_combined.columns:
        old_pay = df_combined['供款方式'].fillna('').astype(str)
        new_pay = old_pay.apply(
            lambda x: PAYMENT_METHOD_MAP.get(x.strip(), x) if x.strip() and x.strip() != '0' else x
        )
        pay_mapped = (old_pay != new_pay).sum()
        if pay_mapped > 0:
            print(f"   供款方式映射: {pay_mapped}个值")
            from collections import Counter
            pay_details = Counter()
            for old, new in zip(old_pay, new_pay):
                if old != new and old.strip() != '0':
                    pay_details[(old, new)] += 1
            for (old, new), cnt in pay_details.most_common():
                print(f"     \"{old}\" → \"{new}\" ({cnt}条)")
            df_combined['供款方式'] = new_pay
        else:
            print("   无需映射的供款方式值")
        # 供款方式值 '0' → NaN（来自V0-NGP-1的空值）
        mask_pay_zero = df_combined['供款方式'].astype(str).str.strip() == '0'
        if mask_pay_zero.sum() > 0:
            print(f"   供款方式值\"0\" → 空值 ({mask_pay_zero.sum()}条)")
            df_combined.loc[mask_pay_zero, '供款方式'] = np.nan
        # 打印映射后值域
        pay_vals = df_combined['供款方式'].dropna().unique()
        print(f"   映射后供款方式值域: {sorted([str(x) for x in pay_vals])}")

    # 电话列 .0 后缀修复（float→str残留，如 "62982319.0" → "62982319"）
    tel_cols = [c for c in df_combined.columns if '电话' in str(c)]
    for col in tel_cols:
        if col in df_combined.columns:
            tel_str = df_combined[col].astype(str)
            dot0_mask = tel_str.str.endswith('.0')
            dot0_count = dot0_mask.sum()
            if dot0_count > 0:
                # 去掉 .0 后缀，检查是否为纯数字 → 转为 int
                tel_fixed = 0
                for idx in df_combined.index:
                    raw = str(df_combined[col].iloc[df_combined.index.get_loc(idx)]) if df_combined[col].iloc[df_combined.index.get_loc(idx)] is not np.nan else ''
                    # Check .0 suffix
                    s = str(df_combined.loc[idx, col])
                    if s == 'nan' or s == '0':
                        if s == '0':
                            df_combined.loc[idx, col] = np.nan
                        continue
                    if s.endswith('.0'):
                        s_no_dot = s[:-2]  # 去掉 ".0"
                        if s_no_dot.isdigit():
                            df_combined.loc[idx, col] = int(s_no_dot)
                            tel_fixed += 1
                            continue
                    # 常规纯数字 → int
                    if s.isdigit():
                        df_combined.loc[idx, col] = int(s)
                if tel_fixed > 0:
                    print(f"   电话列({col}).0后缀修复: {tel_fixed}个值去掉.0转为int")
                # 检查残留 .0
                remaining_dot0 = df_combined[col].astype(str).str.endswith('.0').sum()
                if remaining_dot0 > 0:
                    print(f"   ⚠️ 电话列({col})仍有 {remaining_dot0} 个值带.0后缀（非纯数字）")

    # 数值列中字符串 '0' → NaN
    numeric_string_cols = ['保监征费', '合计', '保额', '续保金额', '保费储备金户口']
    for col in numeric_string_cols:
        if col in df_combined.columns:
            mask = df_combined[col].astype(str).str.strip() == '0'
            df_combined.loc[mask, col] = np.nan
    # 佣金模式 NaN → 0
    if '佣金模式' in df_combined.columns:
        df_combined['佣金模式'] = df_combined['佣金模式'].fillna(0)
    # 计划书年龄 int → float（避免精度问题）
    if '计划书年龄' in df_combined.columns:
        df_combined['计划书年龄'] = pd.to_numeric(df_combined['计划书年龄'], errors='coerce')
    # 年期列去"年"字（如"25年"→"25"，"15年"→"15"，与V0格式一致）
    if '年期' in df_combined.columns:
        nianqi_old = df_combined['年期'].fillna('').astype(str)
        # 整付保费 → 年期=1，且对应行供款方式改为"整付"
        zhengfu_mask = nianqi_old.str.strip() == '整付保费'
        if zhengfu_mask.sum() > 0:
            print(f"   整付保费特殊处理: {zhengfu_mask.sum()}行 → 年期=1, 供款方式=整付")
            if '供款方式' in df_combined.columns:
                df_combined.loc[zhengfu_mask, '供款方式'] = '整付'
        nianqi_new = nianqi_old.str.rstrip('年').str.strip()
        # 纯数字值转为int/float，非纯数字保留str；"整付保费"→1
        SPECIAL_NIANQI_MAP = {'整付保费': 1}
        def _clean_nianqi(x):
            if x == '' or x == '0':
                return np.nan
            if x in SPECIAL_NIANQI_MAP:
                return SPECIAL_NIANQI_MAP[x]
            try:
                return int(x)
            except ValueError:
                try:
                    return float(x)
                except ValueError:
                    return x  # 如"至100岁"等保留原值
        nianqi_cleaned = nianqi_new.apply(_clean_nianqi)
        changed_count = (nianqi_old != nianqi_cleaned.astype(str)).sum()
        if changed_count > 0:
            print(f"   年期去\"年\"字: {changed_count}个值发生变化")
            # 显示转换示例
            from collections import Counter
            nianqi_changes = Counter()
            for old, new in zip(nianqi_old, nianqi_cleaned.astype(str)):
                if old != new and old != '' and old != '0':
                    nianqi_changes[(old, new)] += 1
            for (old, new), cnt in nianqi_changes.most_common(10):
                print(f"     \"{old}\" → \"{new}\" ({cnt}条)")
        df_combined['年期'] = nianqi_cleaned

    # ---- Step 10c: 保单号码格式化（与V0一致） ----
    # V0规则：纯数字保单号码 → int类型（去掉前缀零），非纯数字 → str类型
    print()
    print("🔢 Step 10c: 保单号码格式化（纯数字→int无前缀零，非纯数字→str）...")
    if '保单号码' in df_combined.columns:
        pn_col = df_combined['保单号码'].copy()
        # 保单号码值为 '0' → NaN
        mask_pn_zero = pn_col.astype(str).str.strip() == '0'
        zero_count = mask_pn_zero.sum()
        if zero_count > 0:
            print(f"   保单号码值\"0\" → 空值 ({zero_count}条)")
            pn_col[mask_pn_zero] = np.nan

        # 去掉前缀零，判断是否是纯数字
        pn_str = pn_col.dropna().astype(str)
        stripped = pn_str.str.lstrip('0')
        # 去掉前缀零后是纯数字 → 转为int
        pure_numeric_mask = stripped.str.isdigit()
        numeric_count = pure_numeric_mask.sum()
        # 去掉前缀零后非纯数字（含字母、括号等） → 保留str
        non_numeric_count = (~pure_numeric_mask).sum()

        # 应用转换
        # 先处理 .0 后缀（float→str残留，如 "602068558.0" → "602068558" → int）
        dot0_fixed = 0
        for idx in pn_col.dropna().index:
            raw = str(pn_col[idx])
            s = raw.lstrip('0')
            # 处理 .0 后缀：去掉后是纯数字则转为int
            if s.endswith('.0'):
                s_no_dot = s[:-2]  # 去掉 ".0"
                if s_no_dot.isdigit():
                    pn_col[idx] = int(s_no_dot)
                    dot0_fixed += 1
                    continue
            # 常规判断：纯数字→int，非纯数字→str（去掉前缀零）
            if s.isdigit():
                pn_col[idx] = int(s)
            else:
                # 含字母/括号等，保留原值
                pn_col[idx] = raw.lstrip('0')

        df_combined['保单号码'] = pn_col
        print(f"   .0后缀修复（float→str残留→去掉.0→int）: {dot0_fixed}个")
        print(f"   纯数字保单号码（去掉前缀零→int）: {numeric_count - dot0_fixed}个")
        print(f"   非纯数字保单号码（保留str）: {non_numeric_count}个")
        if zero_count > 0:
            print(f"   空值保单号码: {zero_count + pn_col.isna().sum() - zero_count}个")

    print("   数据清洗完成")

    # ---- Step 11: 在整合文件中新增 V0-1 Sheet ----
    print()
    print("✨ Step 11b: 在整合文件中新增 V0-1 Sheet...")

    # 直接打开整合文件
    wb = openpyxl.load_workbook(integrate_file_path)
    original_sheets = wb.sheetnames
    print(f"   文件原有 sheet: {original_sheets}")

    # 如果已存在 V0-1 sheet，先删除
    if 'V0-1' in wb.sheetnames:
        del wb['V0-1']
        print("   已删除旧 V0-1 sheet")

    # 创建 V0-1 sheet（放在 V0-NGP-1 左边）
    ws = wb.create_sheet('V0-1')

    # 写入表头
    for col_idx, col_name in enumerate(V0_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)

    # 写入数据
    for row_idx, (_, row) in enumerate(df_combined.iterrows(), 2):
        for col_idx, col_name in enumerate(V0_COLUMNS, 1):
            val = row[col_name]
            # 处理 NaN
            if pd.isna(val):
                cell = ws.cell(row=row_idx, column=col_idx)
                # 不写入值，保持空单元格
            else:
                cell = ws.cell(row=row_idx, column=col_idx, value=val)

    # ---- Step 12: 设置 V0 格式 ----
    print()
    print("🎨 Step 12b: 设置 Excel 格式（复刻 V0 样式）...")

    # 构建 列名→列序号 映射
    col_name_to_idx = {}
    for i, name in enumerate(V0_COLUMNS, 1):
        col_name_to_idx[name] = i

    # 表头样式（与V0完全一致）
    # V0规则: bold=True for Col1-76, bold=False for Col77-80
    # wrap_text: 只有Col51(申请表递交状态)=True, 其余79列=None
    # border_bottom=thin: Col4-7(转介区域), Col57-59(融资/分群/折扣), Col77-80(公式读取); 其余=None
    # border_left=None: Col77(计划书编号); 其余=thin
    # border_right=None: Col68(回执状况); 其余=thin
    # border_top=thin: 全部80列
    
    # 定义哪些列的表头有特殊样式
    HEADER_BOLD_FALSE_COLS = [
        '计划书编号(仅平安需登记）', '港分经理（公式读取）',
        'IS老师（公式读取）', '预约类型（公式读取）'
    ]  # Col77-80: bold=False
    HEADER_WRAP_TRUE_COLS = ['申请表递交状态']  # Col51: wrap_text=True
    HEADER_BOTTOM_THIN_COLS = [
        '转介公司', '转介人', '转介日期', '转介时间',  # Col4-7
        '是否融资单', '客户分群', '首年特殊折扣',  # Col57-59
        '计划书编号(仅平安需登记）', '港分经理（公式读取）',
        'IS老师（公式读取）', '预约类型（公式读取）'  # Col77-80
    ]  # 11列 bottom=thin, 其余=None
    HEADER_LEFT_NONE_COLS = ['计划书编号(仅平安需登记）']  # Col77: left=None
    HEADER_RIGHT_NONE_COLS = ['回执状况']  # Col68: right=None

    # V0 数据行格式分组（与V0完全一致）
    # V0规则: 39列有center/center/wrap_text=True + 4边thin border (styled组)
    #         41列有None/None/None alignment + None border + General格式 (plain组)
    #         Col8(保单状态)有黄色填充, Col37(合计)有微软雅黑字体+右对齐+无边框

    DATA_STYLED_COLS = [
    # Col2-14: 核心业务字段
    '订单编号', '提交日期', '保单状态', '签单日期', '签单时间',
    '签单地点', '签单供应商', '注册编号IA', 'TR',
    # Col17-26: 业务细分字段
    '业务细分', '市场分层', 'KEY ACCOUNT', '机构公司名字', '合作伙伴',
    '赴港联系人', '电话', '预约备注', '所签地区', '保险公司',
    # Col27-35: 产品/金额字段
    '产品品类', '产品名称', '年期', '供款方式', '币种',
    '计划书年龄', '保费', '保费（港币）', 'APE',
    # Col39-41: 投保人姓名/国籍
    '投保人  (中文)', '投保人  (拼音)', '投保人国籍',
    # Col48: 受保人中文
    '受保人(中文）',
    # Col51-52: 递交状态/日期
    '申请表递交状态', '递交日期',
    # Col57-61: 融单/分群/折扣/保单号码/批核日
    '是否融资单', '客户分群', '首年特殊折扣', '保单号码', '批核日（年/月/日）',
    ]  # 39列: center/center/wrap_text=True, 4边thin border

    DATA_YELLOW_FILL_COLS = ['保单状态']  # Col8: 黄色填充 FFFFFF00

    DATA_SPECIAL_FONT_COLS = {'合计': '微软雅黑'}  # Col37: 微软雅黑(其余列宋体)

    DATA_RIGHT_ALIGN_COLS = ['合计']  # Col37: horizontal=right (其余styled列=center)

    # V0 数据行逐列 number_format（仅列出不等于 General 的列）
    # plain列(41列)全部=General, styled列按V0逐列指定
    DATA_NUMBER_FORMAT_MAP = {
    # @ (文本格式)
    '订单编号': '@',
    '签单地点': '@', '签单供应商': '@', '注册编号IA': '@', 'TR': '@',
    '业务细分': '@', '市场分层': '@', 'KEY ACCOUNT': '@',
    '机构公司名字': '@', '合作伙伴': '@', '赴港联系人': '@',
    '电话': '@', '预约备注': '@', '所签地区': '@', '保险公司': '@',
    '产品名称': '@', '年期': '@', '供款方式': '@', '币种': '@',
    '投保人  (中文)': '@', '投保人  (拼音)': '@', '投保人国籍': '@',
    '受保人(中文）': '@', '申请表递交状态': '@',
    '是否融资单': '@', '客户分群': '@', '首年特殊折扣': '@',
    # yyyy-mm-dd (日期格式)
    '提交日期': 'yyyy\\-mm\\-dd', '签单日期': 'yyyy\\-mm\\-dd', '签单时间': 'yyyy\\-mm\\-dd',
    '递交日期': 'yyyy\\-mm\\-dd', '批核日（年/月/日）': 'yyyy\\-mm\\-dd',
    # #,##0.00_ (金额格式)
    '保费': '#,##0.00_ ', '保费（港币）': '#,##0.00_ ', 'APE': '#,##0.00_ ',
    '合计': '#,##0.00_ ',
    # General (数值格式, 在styled列中)
    '保单状态': 'General', '产品品类': 'General', '计划书年龄': 'General', '保单号码': 'General',
    }

    header_font_bold = Font(name='微软雅黑', size=11, bold=True, color='FFFFFFFF')
    header_font_not_bold = Font(name='微软雅黑', size=11, bold=False, color='FFFFFFFF')

    # 构建 列名→颜色 映射
    col_to_color = {}
    for color_hex, cols in HEADER_COLOR_GROUPS.items():
        for col in cols:
            col_to_color[col] = color_hex

    # 设置表头格式（逐列，与V0完全一致）
    for col_name in V0_COLUMNS:
        idx = col_name_to_idx[col_name]
        cell = ws.cell(1, idx)
        color_hex = col_to_color.get(col_name, 'FFC55A11')
        
        # Font: bold=True (Col1-76) or bold=False (Col77-80)
        if col_name in HEADER_BOLD_FALSE_COLS:
            cell.font = header_font_not_bold
        else:
            cell.font = header_font_bold
        
        # Fill
        cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
        
        # Alignment: 只有Col51 wrap_text=True，其余None
        if col_name in HEADER_WRAP_TRUE_COLS:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=None)
        
        # Border: 逐边设置（与V0一致）
        left_style = 'thin' if col_name not in HEADER_LEFT_NONE_COLS else None
        right_style = 'thin' if col_name not in HEADER_RIGHT_NONE_COLS else None
        top_style = 'thin'
        bottom_style = 'thin' if col_name in HEADER_BOTTOM_THIN_COLS else None
        cell.border = Border(
            left=Side(style=left_style, color='FF000000') if left_style else Side(style=None),
            right=Side(style=right_style, color='FF000000') if right_style else Side(style=None),
            top=Side(style=top_style, color='FF000000'),
            bottom=Side(style=bottom_style, color='FF000000') if bottom_style else Side(style=None),
        )

    # 数据行样式
    data_font = Font(name='宋体', size=11)
    data_align_center = Alignment(horizontal='center', vertical='center')
    data_align_default = Alignment(horizontal=None, vertical=None)
    data_border = Border(
        left=Side(style='thin', color='FF000000'),
        right=Side(style='thin', color='FF000000'),
        top=Side(style='thin', color='FF000000'),
        bottom=Side(style='thin', color='FF000000'),
    )

    # 设置数据行格式
    total_rows = df_combined.shape[0] + 1  # 含表头

    for row_idx in range(2, total_rows + 1):
        for col_name in V0_COLUMNS:
            idx = col_name_to_idx[col_name]
            cell = ws.cell(row_idx, idx)
            cell.font = data_font
            cell.border = data_border

            # 对齐方式
            if col_name in TEXT_FORMAT_COLUMNS or col_name in DATE_FORMAT_COLUMNS or col_name in AMOUNT_COLUMNS or col_name in ['保单状态', '计划书年龄', '保监征费', '保单号码']:
                cell.alignment = data_align_center
            elif col_name == '合计':
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = data_align_default

            # 数字格式
            if col_name in DATE_FORMAT_COLUMNS:
                cell.number_format = 'yyyy\\-mm\\-dd'
            elif col_name in AMOUNT_COLUMNS:
                cell.number_format = '#,##0.00_ '
            elif col_name in TEXT_FORMAT_COLUMNS:
                cell.number_format = '@'
            elif col_name == '保监征费':
                cell.number_format = 'General'
            elif col_name == '保额':
                cell.number_format = 'General'
            elif col_name == '保单号码':
                # 保单号码：纯数字 → General（int类型），非纯数字 → @（str类型）
                cell_val = cell.value
                if cell_val is None or (isinstance(cell_val, str) and cell_val.strip() == ''):
                    cell.number_format = '@'
                elif isinstance(cell_val, int):
                    cell.number_format = 'General'
                elif isinstance(cell_val, str):
                    cell.number_format = '@'
                else:
                    cell.number_format = 'General'

    # 设置列宽
    for col_name in V0_COLUMNS:
        idx = col_name_to_idx[col_name]
        col_letter = get_column_letter(idx)
        width = V0_COLUMN_WIDTHS.get(col_name, 15)
        ws.column_dimensions[col_letter].width = width

    # ---- Step 13: 将 V0-1 sheet 移到 V0-NGP-1 左边 ----
    print()
    print("📍 Step 13b: 将 V0-1 sheet 移到 V0-NGP-1 左边...")

    # 找到 V0-NGP-1 的位置
    if 'V0-NGP-1' in wb.sheetnames:
        v0ngp1_index = wb.sheetnames.index('V0-NGP-1')
        v01_index = wb.sheetnames.index('V0-1')
        # 移动 V0-1 到 V0-NGP-1 之前
        wb.move_sheet('V0-1', offset=v0ngp1_index - v01_index - 1)
        print(f"   V0-1 已移至 V0-NGP-1 左边")
        print(f"   新的 sheet 顺序: {wb.sheetnames}")
    else:
        print("   ⚠️ 未找到 V0-NGP-1 sheet，V0-1 放在末尾")

    # 保存（直接修改整合文件）
    wb.save(integrate_file_path)
    wb.close()

    print()
    print("=" * 60)
    print("✅ 转换完成！")
    print(f"   整合文件: {integrate_file_path}")
    print(f"   V0-1 Sheet 总行数: {df_combined.shape[0]}")
    print(f"     综合查询数据行: {df_query.shape[0]}")
    print(f"     V0-NGP-1数据行: {df_ngp1.shape[0]}")
    print(f"   V0-1 Sheet 列数: {len(V0_COLUMNS)}")
    print("=" * 60)

    return integrate_file_path


# ============================================================
# 命令行入口
# ============================================================

def _resolve_path(filepath):
    """如果路径是相对路径（不含盘符和绝对路径标记），则解析为脚本所在目录下的路径。"""
    p = Path(filepath)
    if not p.is_absolute():
        # 相对路径 → 解析为脚本自身所在目录
        script_dir = Path(__file__).resolve().parent
        p = script_dir / p
    return str(p.resolve())


def main():
    if len(sys.argv) < 3:
        print("用法: python query_to_v01.py <综合查询文件> <整合文件> [映射表路径]")
        print()
        print("参数说明:")
        print("  综合查询文件    - 综合查询结果 xlsx 文件名或路径（如 综合查询结果20260710 (源数据).xlsx）")
        print("  整合文件        - 业绩数据-整合文件名或路径（含 V0-NGP-1 sheet）")
        print("                    脚本直接在此文件中新增 V0-1 sheet")
        print("  映射表路径(可选) - 产品名称映射表路径（默认自动查找同目录下的业务部门字段映射表.xlsx）")
        print()
        print("示例（使用文件名，自动在同目录下查找）:")
        print("  python query_to_v01.py \"综合查询结果20260710 (源数据).xlsx\" \"业绩数据-整合.xlsx\"")
        print("  python query_to_v01.py \"综合查询结果20260710 (源数据).xlsx\" \"业绩数据-整合.xlsx\" \"业务部门字段映射表.xlsx\"")
        print()
        print("说明:")
        print("  - 传入文件名（如 综合查询结果20260710 (源数据).xlsx）时，自动在脚本所在目录查找")
        print("  - 传入绝对路径时，直接使用该路径")
        print("  - 删除签单供应商=永领致远顾问有限公司的数据")
        print("  - 综合查询列名映射为V0格式（80列）")
        print("  - V0-NGP-1数据追加在末尾")
        print("  - 直接在整合文件中新增 V0-1 sheet，放在 V0-NGP-1 左边")
        print("  - ⚠️ 执行前确保整合文件没有被 Excel 打开！")
        print("  - 映射表默认自动查找：优先整合文件目录，其次综合查询文件目录")
        print("  - 繁简转换 + 产品名称映射（自动查找或手动指定映射表）")
        sys.exit(1)

    query_file_path = sys.argv[1]
    integrate_file_path = sys.argv[2]
    mapping_table_path = sys.argv[3] if len(sys.argv) >= 4 else None

    # 相对路径 → 解析为脚本所在目录
    query_file_path = _resolve_path(query_file_path)
    integrate_file_path = _resolve_path(integrate_file_path)
    if mapping_table_path:
        mapping_table_path = _resolve_path(mapping_table_path)

    if not Path(query_file_path).exists():
        print(f"❌ 综合查询文件不存在: {query_file_path}")
        sys.exit(1)

    if not Path(integrate_file_path).exists():
        print(f"❌ 整合文件不存在: {integrate_file_path}")
        sys.exit(1)

    result = convert_query_to_v01(query_file_path, integrate_file_path,
                                   mapping_table_path=mapping_table_path)
    print(f"\n🎉 文件已更新: {result}")


if __name__ == '__main__':
    main()
