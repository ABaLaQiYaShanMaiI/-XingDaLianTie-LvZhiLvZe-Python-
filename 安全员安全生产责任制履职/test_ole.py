# -*- coding: utf-8 -*-
"""测试：图片 + OLE 文件混合插入"""
import os, traceback
import win32com.client
import kaoping_core as kc

IMG = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\高温天气预警\8.3.jpg"
DOCX = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.08\新员工入职和电瓶车置换\（供矿）蒋千林\电动车审批蒋千林.docx"
XLSX = r"E:\2.武钢兴达工作\2.月度固定工作\2026\2026.05\履职履责\2026年履职履责表5月（兴达炼铁保产事业部）.xlsx"

for p in (IMG, DOCX, XLSX):
    print(os.path.exists(p), p)

tpl = kc.find_template()
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = None
try:
    doc = word.Documents.Open(os.path.abspath(tpl))
    tb = doc.Tables(1)
    # 项2(行5): 图片
    c = tb.Cell(5, 8)
    kc._clear_cell_content(c)
    kc._insert_materials(c, [IMG])
    # 项5(行8): docx OLE
    c = tb.Cell(8, 8)
    kc._clear_cell_content(c)
    kc._insert_materials(c, [DOCX])
    # 项9(行12): xlsx OLE + 图片混合
    c = tb.Cell(12, 8)
    kc._clear_cell_content(c)
    kc._insert_materials(c, [XLSX, IMG])
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_ole_test.doc")
    doc.SaveAs2(os.path.abspath(out), 0)
    print("保存:", out, os.path.getsize(out))
    print("文档 InlineShapes:", doc.InlineShapes.Count)
    # 检查各单元格
    for r in (5, 8, 12):
        cell = tb.Cell(r, 8)
        print(f"R{r}C8 shapes:", len(cell.Range.InlineShapes),
              "text:", repr(cell.Range.Text[:30]))
finally:
    if doc is not None:
        try:
            doc.Close(False)
        except Exception:
            pass
    word.Quit()
print("DONE")
