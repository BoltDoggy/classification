#!/usr/bin/env python3
"""从 GB/T 4754—2017《国民经济行业分类》docx 中提取分类数据,生成 data.js。

用法: python3 tools/extract_docx.py [input.docx] [output.js]
默认输入 .input/P020210608556900301740.docx,输出 ./data.js
"""
import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

SRC = sys.argv[1] if len(sys.argv) > 1 else '.input/P020210608556900301740.docx'
OUT = sys.argv[2] if len(sys.argv) > 2 else 'data.js'


def load_tables(path):
    doc = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    body = doc.find(W + 'body')

    def cell_text(tc):
        return ''.join(t.text or '' for t in tc.iter(W + 't')).strip()

    tables = []
    for tbl in body.findall(W + 'tbl'):
        raw = [[cell_text(tc) for tc in tr.findall(W + 'tc')]
               for tr in tbl.findall(W + 'tr')]
        width = max(len(r) for r in raw)
        tables.append([(r + [''] * width)[:width] for r in raw])
    return tables


def parse_main(rows):
    """主表: 门类/大类/中类/小类/名称/说明 -> 扁平条目列表。

    处理三类特殊行:
    1. 中类与其唯一小类同行(如 701/7010 房地产开发经营),共 200 行;
       其中 727、905 两行连中类/小类名称也挤在同一单元格,需显式拆分;
    2. 大类 27 下 277/278 两个中类挤在一行(2770/2780);
    3. 小类 8052/8053(足浴服务/养生保健服务)名称连排挤在一行。
    """
    entries = []
    for men, da, zhong, xiao, name, desc in rows[2:]:
        if not name:
            continue
        if zhong == '727':
            entries.append({'code': '727', 'name': '安全保护服务',
                            'desc': desc, 'level': 3})
            entries.append({'code': '7271', 'name': '安全服务', 'desc': '', 'level': 4})
            continue
        if zhong == '905':
            entries.append({'code': '905', 'name': '文化体育娱乐活动与经纪代理服务',
                            'desc': '', 'level': 3})
            entries.append({'code': '9051', 'name': '文化活动服务',
                            'desc': desc, 'level': 4})
            continue
        if zhong == '277278':
            names = name.split('  ')
            descs = [desc[:desc.index('指药品用')], desc[desc.index('指药品用'):]]
            pairs = [('277', '2770'), ('278', '2780')]
            for (zc, xc), n, d in zip(pairs, names, descs):
                entries.append({'code': zc, 'name': n, 'desc': '', 'level': 3})
                entries.append({'code': xc, 'name': n, 'desc': d, 'level': 4})
            continue
        if xiao == '80528053':
            entries.append({'code': '8052', 'name': '足浴服务', 'desc': '', 'level': 4})
            entries.append({'code': '8053', 'name': '养生保健服务',
                            'desc': desc, 'level': 4})
            continue
        if zhong and xiao:
            entries.append({'code': zhong, 'name': name, 'desc': '', 'level': 3})
            entries.append({'code': xiao, 'name': name, 'desc': desc, 'level': 4})
        elif men:
            entries.append({'code': men, 'name': name, 'desc': desc, 'level': 1})
        elif da:
            entries.append({'code': da, 'name': name, 'desc': desc, 'level': 2})
        elif zhong:
            entries.append({'code': zhong, 'name': name, 'desc': desc, 'level': 3})
        else:
            entries.append({'code': xiao, 'name': name, 'desc': desc, 'level': 4})
    return entries


def build_tree(entries):
    root = {'code': '', 'name': '国民经济行业分类', 'level': 0, 'desc': '', 'children': []}
    stack = [root]
    for e in entries:
        node = {'code': e['code'], 'name': e['name'], 'level': e['level'],
                'desc': e['desc'], 'children': []}
        while stack and stack[-1]['level'] >= e['level']:
            stack.pop()
        stack[-1]['children'].append(node)
        stack.append(node)
    return root


def parse_stats(rows):
    """附录 A 统计表: 每门类 2017/2011 大中小类数量。"""
    stats = {}
    for label, d17, z17, x17, _l11, d11, z11, x11, *_ in rows[2:]:
        m = re.match(r'([A-T])\s*(.+)', label)
        if not m:
            continue
        stats[m.group(1)] = {
            'name': m.group(2), 'd': int(d17), 'z': int(z17), 'x': int(x17),
            'd11': int(d11), 'z11': int(z11), 'x11': int(x11)}
    return stats


def split_pairs(codes, names):
    """对照侧拼接代码(如 '13201363')与名称('饲料加工  水产饲料制造')按位置拆分。

    名称无法按数量对齐时返回 None,由调用方降级为合并展示。
    """
    chunks = [codes[i:i + 4] for i in range(0, len(codes), 4)]
    parts = [p.strip() for p in re.split(r'\s{2,}', names.strip()) if p.strip()]
    if len(parts) == len(chunks):
        return list(zip(chunks, parts))
    return None


# 附录 B 中代码挤在一行的行(中类+小类或两个小类),按标准原文显式拆分
T2_SPECIALS = {
    '2454': ([('2450', '玩具制造')], '新增，将原2450分解'),
    '2455': ([('2450', '玩具制造')], '新增，将原2450分解'),
    '4831': ([('4830', '海洋工程建筑')], '新增，将原4830分解'),
    '6511': ([('6510', '软件开发')], '新增，将原6510分解'),
    '7271': ([('7281', '安全服务')], ''),
    '7272': ([('7282', '安全系统监控服务')], ''),
}

# 附录 C 中同样挤在一行的行
T3_SPECIALS = {
    '2454': [('3240', '游艺用品及玩具的制造')],
    '2455': [('3240', '游艺用品及玩具的制造')],
    '7271': [('8010', '私人安全活动')],
    '7272': [('8020', '安全系统服务活动')],
}

# 附录 C 原文档的行代码错误: 该行名称对应主表 2456,代码却排成 2455
T3_CODE_FIX = {('2455', '儿童乘骑玩耍的童车类产品制造'): '2456'}

# 附录 C 原文档 751 区块排版错位(7512 重复、后续错一行),
# 该区块所有小类均对应 ISIC 7490,按此对齐补全 7517
T3_751X = {'c': '7490', 'n': '未另分类的其他专业、科学和技术活动'}


def packed_codes(c17):
    """展开挤在一格的 2017 代码: 7 位=中类+小类,8 位=两个小类。"""
    if len(c17) == 7:
        return [c17[3:]]
    if len(c17) == 8:
        return [c17[:4], c17[4:]]
    return [c17]


def parse_2011_map(rows):
    """附录 B 对照表: 2017 小类代码 -> {t: [{c,n,note}...], note(公共说明)}。"""
    m = {}
    for row in rows[2:]:
        for c17 in packed_codes(row[0]):
            _n17, c11, n11, note = row[1], row[2], row[3], row[4]
            if len(c17) != 4:
                continue
            if c17 in T2_SPECIALS:
                targets, note = T2_SPECIALS[c17]
                m[c17] = {'t': [{'c': c, 'n': n, 'note': note}
                                for c, n in targets]}
                continue
            if not c11:
                m[c17] = {'t': [{'c': '', 'n': '', 'note': note or '新增'}]}
                continue
            if len(c11) == 4:
                targets = [(c11, n11)]
            else:
                targets = split_pairs(c11, n11) or \
                    [('/'.join(c11[i:i + 4] for i in range(0, len(c11), 4)), n11)]
            # 说明与目标数量一致时按目标拆分(双空格分隔)
            parts = [p.strip() for p in re.split(r'\s{2,}', note)] if note else []
            if len(parts) == len(targets):
                m[c17] = {'t': [{'c': c, 'n': n, 'note': np}
                                for (c, n), np in zip(targets, parts)]}
            else:
                m[c17] = {'t': [{'c': c, 'n': n, 'note': note}
                                for c, n in targets]}
    return m


def parse_isic(rows):
    """附录 C 对照表: 2017 代码 -> ISIC Rev.4 对应列表(可能多条)。"""
    m = {}
    for row in rows[2:]:
        _n17, isic, isic_name = row[1], row[2], row[3]
        c17 = T3_CODE_FIX.get((row[0], _n17), row[0])
        for c17 in packed_codes(c17):
            if not c17 or not isic:
                continue
            if c17 in T3_SPECIALS:
                pairs = T3_SPECIALS[c17]
            elif len(isic) == 4:
                pairs = [(isic, isic_name)]
            else:
                pairs = split_pairs(isic, isic_name) or \
                    [('/'.join(isic[i:i + 4] for i in range(0, len(isic), 4)),
                      isic_name)]
            bucket = m.setdefault(c17, [])
            for c, n in pairs:
                if not any(e['c'] == c and e['n'] == n for e in bucket):
                    bucket.append({'c': c, 'n': n})
    for code in ['7511', '7512', '7513', '7514', '7515', '7516', '7517', '7519']:
        m[code] = [dict(T3_751X)]
    return m


def count_level(node, lv):
    n = 1 if node['level'] == lv else 0
    return n + sum(count_level(c, lv) for c in node['children'])


def compact(node):
    """转为紧凑结构 {c,n,d,ch},剔除空字段。"""
    out = {'c': node['code'], 'n': node['name']}
    if node['desc']:
        out['d'] = node['desc']
    if node['children']:
        out['ch'] = [compact(c) for c in node['children']]
    return out


def main():
    tables = load_tables(SRC)
    entries = parse_main(tables[0])
    root = build_tree(entries)
    stats = parse_stats(tables[1])
    m2011 = parse_2011_map(tables[2])
    isic = parse_isic(tables[3])

    got = tuple(count_level(root, lv) for lv in (1, 2, 3, 4))
    expect = (20, 97, 473, 1382)
    assert got == expect, f'层级数量不符: {got} != {expect}'

    codes = [e['code'] for e in entries]
    assert len(codes) == len(set(codes)), '代码存在重复'
    import re as _re
    weird = [c for c in codes if not _re.fullmatch(r'[A-T]|\d{2}|\d{3}|\d{4}', c)]
    assert not weird, f'存在异常代码(疑为合并单元格未拆分): {weird[:5]}'

    xiao_codes = {e['code'] for e in entries if e['level'] == 4}
    missing = xiao_codes - set(m2011)
    assert not missing, f'2011 对照缺失小类: {sorted(missing)[:8]}'

    def walk(node, parent_code, parent_level):
        assert node['level'] == parent_level + 1, \
            f"{node['code']} 层级跳级(父级 {parent_code})"
        if parent_level >= 2:
            assert node['code'].startswith(parent_code), \
                f"{node['code']} 不是 {parent_code} 的合法子级"
        for c in node['children']:
            walk(c, node['code'], node['level'])
    for child in root['children']:
        walk(child, '', 0)

    data = {
        'meta': {
            'title': '国民经济行业分类',
            'standard': 'GB/T 4754—2017',
            'replaces': 'GB/T 4754—2011',
            'issued': '2017-06-30',
            'effective': '2017-10-01',
            'counts': {'men': got[0], 'da': got[1], 'zhong': got[2], 'xiao': got[3]},
        },
        'tree': [compact(c) for c in root['children']],
        'stats': stats,
        'v2011': m2011,
        'isic': isic,
    }
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('const DATA = ')
        f.write(payload)
        f.write(';\n')

    total = sum(got)
    print(f'OK 门类 {got[0]} / 大类 {got[1]} / 中类 {got[2]} / 小类 {got[3]}(共 {total} 条)')
    print(f'2011 对照 {len(m2011)} 条, ISIC 对照 {len(isic)} 条, 门类统计 {len(stats)} 条')
    print(f'输出: {OUT} ({len(payload) // 1024} KB)')


if __name__ == '__main__':
    main()
