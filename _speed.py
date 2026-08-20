# -*- coding: utf-8 -*-
import os, time, kaoping_core as kc
sld = r'E:\2.武钢兴达工作\2.月度固定工作\2026\2026.07\扫雷'
xlsx = sorted([os.path.join(sld, f) for f in os.listdir(sld) if f.endswith('.xlsx')])[0]
img = r'E:\2.武钢兴达工作\2.月度固定工作\2026\2026.07\（兴达炼铁）夏季交通安全学习\供矿作业区\1.jpg'
tpl = kc.find_template()
base = os.path.dirname(os.path.abspath(tpl))
def run(name, materials):
    items = {i: {'desc': '', 'score': '20', 'material_text': '', 'materials': [], 'eval_desc': '', 'super_score': ''} for i in range(1, 13)}
    items[2]['materials'] = materials
    out = os.path.join(base, '_speed_test.doc')
    t0 = time.time()
    kc.generate_doc(tpl, out, '测试', 8, items, year=2026)
    dt = time.time() - t0
    sz = os.path.getsize(out)
    os.remove(out)
    print(f'{name}: {dt:.1f}s, {sz/1024:.0f}KB', flush=True)
run('empty', [])
run('image x1', [img])
run('xlsx-embed x1', [xlsx])
print('ALLDONE', flush=True)
