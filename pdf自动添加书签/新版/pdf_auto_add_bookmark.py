# -*- coding:utf-8 -*-
# pypdf>=3.0
from pypdf import PdfWriter
from pypdf import PdfReader
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
                    f.write('  ' + str(int(list[-1]) + 13))  # 加上与实际页数相差的12页
                    f.write("\n")
                except Exception:
                    continue
    # with open(txt_out, 'r', encoding='utf-8') as file:
    #     lines = file.readlines()
    #     for line in lines:
    #         if line.strip() == '':
    #             continue
    #         title = line.


def marks_book(in_pdf: str, out_pdf: str, bookmarks: str):
    pdf_read = PdfReader(in_pdf)
    pdf_write = PdfWriter()
    # 会删除已有书签
    for page in range(len(pdf_read.pages)):
        pdf_write.add_page(pdf_read.pages[page])
    # 不会删除已有书签
    # pdf_write.clone_document_from_reader(pdf_read)
    with open(bookmarks, "r", encoding='UTF-8') as f:
        for line in f:
            li_level = line.split('\t')
            li_page = line.split()
            book_title = li_page[0]
            book_page = li_page[-1]
            if len(li_level) == 1:
                bookmark1 = pdf_write.add_outline_item(book_title, book_page)
            elif len(li_level) == 2:
                bookmark2 = pdf_write.add_outline_item(book_title, book_page, parent=bookmark1)
            else:
                pdf_write.add_outline_item(book_title, book_page, parent=bookmark2)
            # bookmark = (len(book_level), book_title, book_page)
            # print(bookmark)
    with open(out_pdf, 'wb') as new_f:
        pdf_write.write(new_f)


if __name__ == '__main__':
    txt_in = r"C:\Users\kler\Downloads\Python_scripts-m\pdf自动添加书签\moban.txt"
    bookmarks = r"C:\Users\kler\Downloads\Python_scripts-m\pdf自动添加书签\moban2222.txt"
    pdf_path = r"C:\Users\kler\Downloads\Python_scripts-m\pdf自动添加书签\Python Web自动化测试入门与实.pdf"
    book_dir(txt_in, bookmarks)
    time.sleep(2)
    marks_book(
        pdf_path,
        r"C:\Users\kler\Downloads\Python_scripts-m\pdf自动添加书签\Python Web自动化测试入门与实2222.pdf",
        bookmarks
    )


'''
报错：
  File "D:\program\python3.11\Lib\site-packages\pypdf\_writer.py", line 1930, in add_outline_item
    page_ref,
    ^^^^^^^^
UnboundLocalError: cannot access local variable 'page_ref' where it is not associated with a value
'''

# pypdf2==2.11.1的源文件_writer.py修改内容如下：
'''
    def add_outline_item(
        self,
        title: str,
        page_number: Union[None, PageObject, IndirectObject, int, str],
        ...
            ...
            elif isinstance(page_number, int):
                try:
                    page_ref = self.pages[page_number].indirect_reference
                except IndexError:
                    page_ref = NumberObject(page_number)
            ======================kler=======================
            elif isinstance(page_number, str):
                page_ref = NumberObject(page_number)
            ======================kler=====================
            if page_ref is None:
                logger_warning(
                    f"can not find reference of page {page_number}",
                    __name__,
                )
                page_ref = NullObject()
'''
