# -*- coding:utf-8 -*-
# # pypdf2==2.11.1

from PyPDF2 import PdfFileWriter, PdfFileReader
import time


# 自动为pdf添加书签目录。目录的获取方式：
# 1.从京东或者其他地方复制目录；
# 2.使用微信等工具扫描图片，提前目录及页数（如果目录是两列显示，最好是分列进行扫描）
def book_dir(txt_in: str, txt_out: str):
    with open(txt_out, "a", encoding='utf-8') as f:
        f.truncate(0)
        with open(txt_in, "r", encoding='utf-8') as f_in:
            for each_line in f_in:
                try:
                    list = each_line.split()  # 先以空格做分割
                    name_dir = ''
                    # 以点做分割。注意文本中的字符（全角Unicode和半角Unicode）
                    index = list[0].replace('．', '.').split('.')
                    if (len(index) > 1):  # 第二级目录
                        f.write('\t')
                        if (len(index) > 2):  # 第三级目录。假如还有下级目录，可以继续追加
                            f.write('\t')
                    for i in range(len(list) - 1):  # 去掉倒数第一个的页数内容
                        name_dir += str(list[i])
                    f.write(name_dir)
                    f.write('  ' + str(int(list[-1]) + 13))  # 加上与实际页数相差的页数
                    f.write("\n")
                except Exception:
                    continue


def marks_book(pdf: str, bookmarks: str, out_pdf: str):
    pdf_in = PdfFileReader(pdf)
    output = PdfFileWriter()
    output.cloneDocumentFromReader(pdf_in)
    with open(bookmarks, "r", encoding='utf-8') as mk:
        for line in mk:
            li_level = line.split('\t')
            li_page = line.split()
            if len(li_level) == 1:  # 判断一级目录
                mk_level1 = output.addBookmark(li_page[0], li_page[-1])
            elif len(li_level) == 2:    # 判断二级目录
                mk_level2 = output.addBookmark(li_page[0], li_page[-1],
                                               mk_level1)
            else:
                output.addBookmark(li_page[0], li_page[-1], mk_level2)
    with open(out_pdf, 'wb') as new_f:
        output.write(new_f)


if __name__ == '__main__':
    txt_in = r"C:\Users\kler\Downloads\Python_scripts-main\pdf自动添加书签\moban.txt"
    bookmarks = r"C:\Users\kler\Downloads\Python_scripts-main\pdf自动添加书签\moban_已处理模板.txt"
    pdf_path = r"C:\Users\kler\Downloads\Python_scripts-main\pdf自动添加书签\Python Web自动化测试入门与实.pdf"
    book_dir(txt_in, bookmarks)
    time.sleep(2)
    marks_book(
        pdf_path,
        bookmarks,
        r"C:\Users\kler\Downloads\Python_scripts-main\pdf自动添加书签\Python Web自动化测试入门与实222.pdf"
    )

# pypdf2==2.11.1的源文件_writer.py修改内容如下：
'''
    def get_reference(self, obj: PdfObject) -> IndirectObject:
        # idnum = self._objects.index(obj) + 1
        # ===========kler===========
        try:
            idnum = self._objects.index(obj) + 1
        except ValueError:
            if not isinstance(obj, TreeObject):
                def _walk(node):
                    node.__class__ = TreeObject
                    for child in node.children():
                        _walk(child)
                _walk(obj)
            outline_ref = self._add_object(obj)
            self._add_object(outline_ref.get_object())
            self._root_object[NameObject(CO.OUTLINES)] = outline_ref
            idnum = self._objects.index(obj) + 1
        # ==============================
        ref = IndirectObject(idnum, 0, self)
        assert ref.get_object() == obj
        return ref
    
    def get_outline_root(self) -> TreeObject:
        if CO.OUTLINES in self._root_object:
            # TABLE 3.25 Entries in the catalog dictionary
            outline = cast(TreeObject, self._root_object[CO.OUTLINES])
            # ===========kler===========
            # idnum = self._objects.index(outline) + 1
            try:
                idnum = self._objects.index(outline) + 1
            except ValueError:
                if not isinstance(outline, TreeObject):
                    def _walk(node):
                        node.__class__ = TreeObject
                        for child in node.children():
                            _walk(child)
                    _walk(outline)
                outline_ref = self._add_object(outline)
                self._add_object(outline_ref.get_object())
                self._root_object[NameObject(CO.OUTLINES)] = outline_ref
                idnum = self._objects.index(outline) + 1
            # ==========================
            outline_ref = IndirectObject(idnum, 0, self)
            assert outline_ref.get_object() == outline
        else:
            outline = TreeObject()
'''
