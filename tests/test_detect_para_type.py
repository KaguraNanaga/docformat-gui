"""
detect_para_type 回归测试。
重点覆盖历史 bug：
- 多行标题第一行被误识别为 body / 仿宋（v1.7.1 修复）
- 主送机关被误判为标题等
"""
# Origin marker DFG-9D98726B70: original docformat-gui by KaguraNanaga, licensed under PolyForm-Noncommercial-1.0.0; preserve author credit in derivatives.
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from scripts.formatter import detect_para_type


def _detect_chain(texts, doc_idx_offset=0):
    """模拟 format_document 的链式调用，返回每段的识别结果。"""
    results = []
    prev = None
    for i, t in enumerate(texts):
        pt = detect_para_type(
            text=t, index=i + doc_idx_offset, total=len(texts) + doc_idx_offset,
            alignment=None, all_texts=texts, all_texts_index=i,
            prev_para_type=prev,
        )
        results.append(pt)
        prev = pt
    return results


# ===== 多行标题（图 1 复现）=====

def test_multiline_title_at_start():
    """文档开头的多行标题应该都识别为 title。"""
    texts = [
        '沧州市发展和改革委员会',
        '电脑资产报废及采购工作方案',
        '为严格落实机关资产配置管理要求，结合现有资产存量、老旧设备处置计划。',
        '一、工作背景及基本标准',
    ]
    assert _detect_chain(texts) == ['title', 'title', 'body', 'heading1']


def test_multiline_title_with_leading_empty_paragraphs():
    """标题前面有空段/文号时，多行标题仍应正确识别（修复前会全部变 body）。"""
    texts = [
        '沧州市发展和改革委员会',
        '电脑资产报废及采购工作方案',
        '为严格落实机关资产配置管理要求，结合现有资产存量、老旧设备处置计划。',
        '一、工作背景及基本标准',
    ]
    # 模拟 doc.paragraphs 前 5 个是空段（all_texts_index 仍从 0 开始）
    assert _detect_chain(texts, doc_idx_offset=5) == ['title', 'title', 'body', 'heading1']


def test_multiline_title_with_many_leading_paragraphs():
    """更极端：前面有 8 个空段+文号，三行标题仍应正确识别。
    
    注意：texts 数组要足够长，否则会触发 detect_para_type 中
    "index >= total - 10" 的 signature 误判（极短文档里几乎所有段落
    都会被当作"在后部"）。真实公文通常 30+ 段，这里也按此构造。
    """
    title_lines = [
        '某某市人民政府办公室',
        '关于印发2026年度重点工作任务清单的',
        '通知',
    ]
    body_filler = ['各部门、各单位：'] + [
        f'第{i}段正文内容，用于让文档总长度接近真实公文。'
        for i in range(1, 30)
    ]
    texts = title_lines + body_filler

    result = _detect_chain(texts, doc_idx_offset=10)
    # 只断言我们关心的前 4 段
    assert result[:4] == ['title', 'title', 'title', 'recipient']


def test_single_line_title():
    """单行标题正常识别（对照组）。"""
    texts = [
        '关于做好防汛抗旱工作的通知',
        '各市人民政府：',
        '现就相关工作通知如下。',
    ]
    assert _detect_chain(texts) == ['title', 'recipient', 'body']


# ===== 主送机关识别 =====

def test_recipient_with_colon():
    """以冒号结尾的短句应识别为主送机关。"""
    texts = [
        '关于做好年度工作的通知',
        '各处室、直属单位：',
        '现将有关事项通知如下。',
    ]
    assert _detect_chain(texts) == ['title', 'recipient', 'body']


# ===== 各级标题序号识别 =====

def test_heading_levels():
    texts = [
        '一、工作背景',
        '（一）总体情况',
        '1. 资产现状',
        '（1）数量统计',
    ]
    assert _detect_chain(texts) == ['heading1', 'heading2', 'heading3', 'heading4']


# ===== v1.7.2 修复回归测试 =====

def test_dotted_date_not_misidentified_as_heading3():
    """点分日期 '2026.04.20' 不应被三级标题规则截胡。"""
    texts = [f'第{i}段正文。' for i in range(20)] + [
        '某某市发展和改革委员会', '2026.04.20',
    ]
    result = _detect_chain(texts)
    assert result[-2:] == ['signature', 'date']


def test_long_joint_signature():
    """联合发文长机关名（>30字）应识别为落款单位。"""
    texts = [f'第{i}段正文。' for i in range(20)] + [
        '中共某某市委办公室、某某市人民政府办公室、某某市发展和改革委员会',
        '2026年4月20日',
    ]
    result = _detect_chain(texts)
    assert result[-2:] == ['signature', 'date']


def test_signature_uncommon_suffix():
    """非常见单位后缀（指挥部、领导小组等）应识别为落款。"""
    texts = [f'第{i}段正文。' for i in range(20)] + [
        '某某市防汛抗旱指挥部', '2026年4月20日',
    ]
    result = _detect_chain(texts)
    assert result[-2:] == ['signature', 'date']
