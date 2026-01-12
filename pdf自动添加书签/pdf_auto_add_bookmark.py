# -*- coding:utf-8 -*-
# pypdf>=3.0

import time

from pypdf import PdfReader, PdfWriter

# 自动为pdf添加书签目录。目录的获取方式：
# 1.从京东或者其他地方复制目录；
# 2.使用微信\chatgpt等工具扫描图片，提前目录及页数（如果目录是两列显示，最好是分列进行扫描）


class pdf_add_bookmark:
    def __init__(self):
        self.current_time = time.strftime("%Y%m%d%H%M%S", time.localtime())

    def book_dir(self, pdf_catalog_txt: str, page_offset: int = 17) -> None:
        pdf_catalog_txt_out = pdf_catalog_txt.split(
            '.')[0] + f'_{self.current_time}.' + pdf_catalog_txt.split('.')[1]
# 优化结构
        with open(pdf_catalog_txt_out, "a", encoding='utf-8') as f:
            f.truncate(0)
            with open(pdf_catalog_txt, "r", encoding='utf-8') as f_in:
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
                        # 加上与实际页数相差的页数
                        f.write('  ' + str(int(list[-1]) + page_offset))
                        f.write("\n")
                    except Exception:
                        continue
        # with open(txt_out, 'r', encoding='utf-8') as file:
        #     lines = file.readlines()
        #     for line in lines:
        #         if line.strip() == '':
        #             continue
        #         title = line.
        print('目录处理完成，目录文件的保存路径为：', pdf_catalog_txt_out)
        return pdf_catalog_txt_out

    def add_bookmarks(self, pdf_file_path: str, pdf_catalog_txt: str, page_offset: int = 17) -> None:
        pdf_read = PdfReader(pdf_file_path)
        pdf_write = PdfWriter()

        pdf_file_path_out = pdf_file_path.split(
            '.')[0] + f'_{self.current_time}.' + pdf_file_path.split('.')[1]

        pdf_catalog_txt_out = self.book_dir(pdf_catalog_txt, page_offset)
        # 会删除已有书签
        for page in range(len(pdf_read.pages)):
            pdf_write.add_page(pdf_read.pages[page])
        # 不会删除已有书签
        # pdf_write.clone_document_from_reader(pdf_read)
        with open(pdf_catalog_txt_out, "r", encoding='UTF-8') as f:
            for line in f:
                li_level = line.split('\t')
                li_page = line.split()
                book_title = li_page[0]
                book_page = int(li_page[-1])
                if len(li_level) == 1:
                    bookmark1 = pdf_write.add_outline_item(
                        book_title, book_page)
                elif len(li_level) == 2:
                    bookmark2 = pdf_write.add_outline_item(
                        book_title, book_page, parent=bookmark1)
                else:
                    pdf_write.add_outline_item(
                        book_title, book_page, parent=bookmark2)
                # bookmark = (len(book_level), book_title, book_page)
                # print(bookmark)

        with open(pdf_file_path_out, 'wb') as f:
            pdf_write.write(f)
        print('书签添加完成，文件保存路径为：', pdf_file_path_out)
        return pdf_file_path_out


if __name__ == '__main__':
    pdf_catalog_txt = r"C:\Users\wii\Desktop\目录.txt"
    pdf_file_path = r"C:\Users\wii\Desktop\awesome-books-master\JavaScript 忍者秘籍.pdf"
    a = pdf_add_bookmark()
    a.add_bookmarks(
        pdf_catalog_txt=pdf_catalog_txt,
        pdf_file_path=pdf_file_path,
        page_offset=20)


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
