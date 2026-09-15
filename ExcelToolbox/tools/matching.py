# -*- coding: utf-8 -*-
"""
编号匹配引擎（纯逻辑，无界面依赖）

从 ExcelQingXi.py 原样移植，未改动任何匹配逻辑。
如需调整匹配行为，只改这里即可，所有工具都会生效。
"""
import re
import unicodedata

DASH_MAP = {
    '\u2010': '-',  # HYPHEN
    '\u2011': '-',  # NON-BREAKING HYPHEN
    '\u2012': '-',  # FIGURE DASH
    '\u2013': '-',  # EN DASH
    '\u2014': '-',  # EM DASH
    '\u2015': '-',  # HORIZONTAL BAR
    '\u2043': '-',  # HYPHEN BULLET
    '\u2212': '-',  # MINUS SIGN
    '\ufe58': '-',  # SMALL EM DASH
    '\ufe63': '-',  # SMALL HYPHEN-MINUS
    '\uff0d': '-',  # FULLWIDTH HYPHEN-MINUS
    '\u02d7': '-',  # MODIFIER LETTER MINUS SIGN
}

# 各种引号/英寸/双撇号 → ASCII '"'（随后由"英寸→in"派生统一）
QUOTE_MAP = {
    '\u201c': '"',  # 左双引号 “
    '\u201d': '"',  # 右双引号 ”   ← 你的 7”Plate 就是这个
    '\u201e': '"',  # 双低引号 „
    '\u2033': '"',  # 双撇号 ″（英寸）
    '\u2036': '"',  # 反转双撇号 ‶
    '\u3003': '"',  # 同上
    '\u301e': '"',  # 同上
    '\u2018': "'",  # 左单引号 ‘
    '\u2019': "'",  # 右单引号 ’
    '\u2032': "'",  # 单撇号 ′（英尺）
}

# 不可见字符（必须删除）
INVISIBLE = '\u200b\u200c\u200d\u200e\u200f\u2060\ufeff\u00ad\u180e'


def to_text(v):
    """单元格值 → 文本。关键修复：把 123.0 这类浮点还原成 123"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))          # 123.0 -> "123"（原代码会变成 "123.0" 导致失配）
    return str(v)


def normalize(s, strip_bracket=False):
    """标准化：统一码点 + 去空白/不可见字符（用于比较）

    处理顺序很关键：必须先做【引号/横杠映射】，再做 NFKC。
    否则 NFKC 会把 ″(U+2033) 拆成两个 ′(U+2032)，英寸号就识别不出来了。
    """
    t = to_text(s)
    # 1) 先统一引号/英寸符号（在 NFKC 之前！否则 U+2033 会被拆成两个 U+2032）
    for k, v in QUOTE_MAP.items():
        t = t.replace(k, v)
    # 2) 先统一各种横杠
    for k, v in DASH_MAP.items():
        t = t.replace(k, v)
    # 3) NFKC：全角字母数字、全角横杠等 → 半角
    t = unicodedata.normalize('NFKC', t)
    # 4) NFKC 可能又产生引号/横杠变体，再扫一遍兜底
    for k, v in QUOTE_MAP.items():
        t = t.replace(k, v)
    for k, v in DASH_MAP.items():
        t = t.replace(k, v)
    # 5) 删除不可见字符
    t = ''.join(ch for ch in t if ch not in INVISIBLE)
    # 6) 删除所有空白（半角/全角/NBSP/制表/换行）
    t = ''.join(ch for ch in t if not ch.isspace())
    # 7) 括号统一成半角
    t = t.replace('（', '(').replace('）', ')').replace('【', '[').replace('】', ']')
    # 8) 可选：删除括号及内容
    if strip_bracket:
        t = re.sub(r'[\(\[][^\)\]]*[\)\]]', '', t)
        t = t.replace('(', '').replace(')', '').replace('[', '').replace(']', '')
    return t.strip()


# ============================================================================
# 多级"派生"匹配
#
# 思路：不再只做一种匹配，而是对同一个值生成多种等价写法，逐个尝试。
#   进入规则之前，键会先过 clean_key()：normalize 之后抹掉尾部句点
#       （123. = 123，这样后面前导零/去序号等规则也能继续组合生效）
#   然后是 10 条派生规则，顺序即优先级：
#   精确 → 英寸号当 in → 英寸号去掉 → 去尾部 -序号 → 去全部尾部 -序号
#        → 去尾部 -序号再去了横杠 → 括号内外互换 → 去括号内容
#        → 去所有横杠 → 数字去前导零
# 并且把【对照表的正确编号】也一并建索引，
# 这样 "123-1 找 123" 和 "123 找 123-1" 两个方向都能命中。
# ============================================================================

def _inch_in(s):
    """英寸号视为 in：7"Plate -> 7inPlate"""
    return s.replace('"', 'in')


def _inch_drop(s):
    """直接去掉英寸号：7"Plate -> 7Plate"""
    return s.replace('"', '')


# 尾部句点（半角 . 和全角 。）—— Excel/系统导出的编号常常多带一个句点
_TAIL_DOT = '.。'


def _strip_tail_dot(s):
    """去掉尾部句点：123. -> 123

    只删【结尾】的句点，中间的不动，所以 1.5、A.B.C 完全不受影响。
    123.. 这类连续多个也会一次去干净（rstrip 是重复剥离）。
    """
    return s.rstrip(_TAIL_DOT)


def clean_key(s, strip_bracket=False):
    """【匹配专用键】= normalize 之后再抹掉尾部句点

    为什么放在这里、而不是当成一条派生规则：
      "0650090." 这种是【尾部句点 + 前导零】的组合，
      单条规则去不掉前导零（_num_lz 要求整串都是数字）。
      在键上先统一抹掉句点，下面 10 条规则就都能正常组合生效了，
      并且建索引和查询两侧都走这个函数，天然对称、不会漏方向。

    注意：normalize() 本身不动，所以输出到 Excel 的"清洗后编号"列
    仍由 clean_match 决定（那边也调 clean_key，见该文件）。
    """
    return _strip_tail_dot(normalize(s, strip_bracket))


def _strip_suffix(s):
    """去掉一个尾部 -数字：123-1 -> 123"""
    return re.sub(r'-\d+$', '', s)


def _strip_suffix_all(s):
    """去掉所有尾部 -数字：123-1-2 -> 123"""
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r'-\d+$', '', s)
    return s


def _nodash(s):
    """去掉所有横杠：123-1 -> 1231"""
    return s.replace('-', '')


def _num_lz(s):
    """纯数字去前导零：0123 -> 123"""
    if s.isdigit():
        t = s.lstrip('0')
        return t if t else '0'
    return s


def _strip_bracket_content(s):
    """删除括号及括号内内容：DRBSS(FCS) -> DRBSS
    用于兜底匹配：源=DRBSS、对照表=DRBSS(FCS) 时也能命中。
    """
    s = re.sub(r'\([^)]*\)', '', s)
    s = re.sub(r'\[[^\]]*\]', '', s)
    return s.replace('(', '').replace(')', '').replace('[', '').replace(']', '')


# 交换不变键的分隔符（用不可见控制符，避免与真实数据冲突）
_SWAP_SEP = '\x1f'


def _swap_paren(s):
    """【括号内外互换不变】123(FC2002) 与 FC2002(123) 得到同一个键

    做法：把"括号外部分"与"每个括号内部分"视为独立片段，
          排序后用 _SWAP_SEP 连接 → 得到与顺序无关的规范形式。
    仅当含括号时生效；无括号则原样返回（不引入新的等价关系）。
    用分隔符连接可避免 12(3) 与 1(23) 这类误撞（12|3 ≠ 1|23）。
    """
    if '(' not in s and '[' not in s:
        return s
    outer = re.sub(r'\([^)]*\)', '', s)
    outer = re.sub(r'\[[^\]]*\]', '', outer)
    inners = re.findall(r'\(([^)]*)\)', s) + re.findall(r'\[([^\]]*)\]', s)
    toks = []
    for piece in [outer] + inners:
        piece = piece.strip('-')
        if piece:
            toks.append(piece)
    if len(toks) < 2:
        return s                      # 只有一段，谈不上互换
    return _SWAP_SEP.join(sorted(toks))


# 顺序即优先级（越靠前越可信）
DERIVS = [
    ('精确',              lambda s: s),
    ('英寸→in',           _inch_in),
    ('英寸去除',           _inch_drop),
    ('去尾部序号',         _strip_suffix),
    ('去全部尾部序号',      _strip_suffix_all),
    ('去尾部序号+去横杠',   lambda s: _nodash(_strip_suffix(s))),
    # ↓ 括号内外互换：保留全部组成信息，比"去括号内容"更可信，故排在它前面
    ('括号内外互换',        _swap_paren),
    ('去括号内容',         _strip_bracket_content),
    ('去横杠',            _nodash),
    ('数字去前导零',        _num_lz),
]


def build_index(pairs):
    """
    由对照表 (原始编号, 正确编号) 构建多级索引。
    每一级都同时收录【原始编号】和【正确编号】两种来源，
    因此 123-1→123 与 123→123-1 两个方向都能匹配。
    """
    idx = {name: {} for name, _ in DERIVS}
    for raw_key, raw_val in pairs:
        nk = clean_key(raw_key).lower()
        nv = clean_key(raw_val).lower()
        for name, fn in DERIVS:
            dk = fn(nk)
            if dk:
                idx[name].setdefault(dk, []).append((raw_val, 0, raw_key))
            dv = fn(nv)
            if dv:
                idx[name].setdefault(dv, []).append((raw_val, 1, raw_key))
    return idx


def lookup(raw, idx, strip_bracket=False):
    """
    返回 (正确编号, 命中规则, 是否歧义)
    """
    plain = normalize(raw, strip_bracket)
    n = _strip_tail_dot(plain).lower()
    if not n:
        return "", "", False
    # 原值带尾部句点时，在规则名前加个前缀，方便在日志里看出"是这句点救了它"
    prefix = '去尾部句点+' if n != plain.lower() else ''
    for name, fn in DERIVS:
        k = fn(n)
        if not k:
            continue
        cands = idx[name].get(k)
        if cands:
            # 择优：先"来源=原始编号"(prio 0)，再取较短的，最后按字符串稳定排序
            best = sorted(cands, key=lambda x: (x[1], len(str(x[0])), str(x[0])))[0]
            distinct = {str(c[0]) for c in cands}
            rule = ('去尾部句点' if name == '精确' else prefix + name) if prefix else name
            return best[0], rule, len(distinct) > 1
    return "", "", False
