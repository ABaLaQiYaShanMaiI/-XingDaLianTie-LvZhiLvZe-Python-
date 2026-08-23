# -*- coding: utf-8 -*-
"""手工回归脚本：图片 + OLE 文件混合插入（需本机安装 Word）。

与旧版不同：不再硬编码 E:\ 绝对路径，素材在临时目录现场生成
（Pillow 造图、openpyxl 造 xlsx），换机器可直接运行。

用法：
    python test_ole.py
"""
import os
import sys
import tempfile

import win32com.client

import kaoping_core as kc


def main():
    tmp = tempfile.mkdtemp(prefix="ole_test_")
    try:
        from PIL import Image
        img = os.path.join(tmp, "测试图片.jpg")
        Image.new("RGB", (1200, 900), (200, 120, 60)).save(img)

        from openpyxl import Workbook
        xlsx = os.path.join(tmp, "测试材料.xlsx")
        wb = Workbook()
        wb.active["A1"] = "测试"
        wb.save(xlsx)

        txt = os.path.join(tmp, "测试说明.txt")
        with open(txt, "w", encoding="utf-8") as f:
            f.write("OLE 材料测试")

        for p in (img, xlsx, txt):
            print(os.path.exists(p), p)

        tpl = kc.find_template()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = None
        try:
            doc = word.Documents.Open(os.path.abspath(tpl))
            tb = doc.Tables(1)
            # 用岗位实际行映射与材料列，避免硬编码行列（各岗位模板行列不同）
            rows = []
            for idx in (1, 2):
                v = kc.ITEM_ROWS[idx]
                rows.append(v[0] if isinstance(v, (list, tuple)) else v)
            col = kc.COL_MATERIAL
            for r in rows:
                c = tb.Cell(r, col)
                kc._clear_cell_content(c)
                kc._insert_materials(c, [img, xlsx, txt])
            out = os.path.join(tmp, "_ole_test.doc")
            doc.SaveAs2(os.path.abspath(out), 0)
            print("已保存:", out, os.path.getsize(out), "bytes")
            for r in rows:
                cell = tb.Cell(r, col)
                print("R%dC%d shapes:" % (r, col), len(cell.Range.InlineShapes),
                      "text:", repr(cell.Range.Text[:30]))
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            word.Quit()
        print("DONE")
        return 0
    finally:
        try:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
