# -*- coding:utf-8 -*-
# __author  :FZ
# Date  :2021-04-06

import time
import pandas as pd
import xlwt
from xmindparser import xmind_to_dict


def styles():
    """设置单元格的样式的基础方法"""
    style = xlwt.XFStyle()
    return style


def borders(status=1):
    """设置单元格的边框
    细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，
    细点虚线:7大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    """
    border = xlwt.Borders()
    border.left = status
    border.right = status
    border.top = status
    border.bottom = status
    return border


def heights(worksheet, line, size=4):
    """设置单元格的高度"""
    worksheet.row(line).height_mismatch = True
    worksheet.row(line).height = size * 256


def widths(worksheet, line, size=11):
    """设置单元格的宽度"""
    worksheet.col(line).width = size * 256


def alignments(**kwargs):
    """设置单元格的对齐方式
    status有两种：horz（水平），vert（垂直）
    horz中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,GENERAL,CENTER_ACROSS_SEL（分散）,RIGHT（右边）,LEFT（左边）
    vert中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,BOTTOM(下方),TOP（上方）
    """
    alignment = xlwt.Alignment()

    if "horz" in kwargs.keys():
        alignment.horz = eval("xlwt.Alignment.HORZ_{}".format(
            kwargs['horz'].upper()))
    if "vert" in kwargs.keys():
        alignment.vert = eval("xlwt.Alignment.VERT_{}".format(
            kwargs['vert'].upper()))
    alignment.wrap = 1  # 设置自动换行
    return alignment


def fonts(name='宋体',
          bold=False,
          underline=False,
          italic=False,
          colour='black',
          height=11):
    """设置单元格中字体的样式
    默认字体为宋体，不加粗，没有下划线，不是斜体，黑色字体
    """
    font = xlwt.Font()
    # 字体
    font.name = name
    # 加粗
    font.bold = bold
    # 下划线
    font.underline = underline
    # 斜体
    font.italic = italic
    # 颜色
    font.colour_index = xlwt.Style.colour_map[colour]
    # 大小
    font.height = 20 * height
    return font


def patterns(colors=1):
    """设置单元格的背景颜色，该数字表示的颜色在xlwt库的其他方法中也适用，默认颜色为白色
    0 = Black, 1 = White,2 = Red, 3 = Green, 4 = Blue,5 = Yellow, 6 = Magenta, 7 = Cyan,
    16 = Maroon, 17 = Dark Green,18 = Dark Blue, 19 = Dark Yellow ,almost brown), 20 = Dark Magenta,
    21 = Teal, 22 = Light Gray,23 = Dark Gray
    """
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = colors
    return pattern


def main(xmindName, reportName, module):
    xm = xmind_to_dict(xmindName)[0]['topic']
    # print(json.dumps(xm, indent=2, ensure_ascii=False))  # indent为显示json格式，ensure_ascii为显示为中文，不显示ASCII码
    workbook = xlwt.Workbook(encoding='utf-8')  # 创建workbook对象
    worksheet = workbook.add_sheet(xm["title"], cell_overwrite_ok=True)  # 创建工作表
    row0 = ['用例编号', '需求', '概要', '步骤ID', '步骤', '期望结果', '报告人', '模块', '优先级', '标签', '测试用例集']
    sizes = [10, 15, 40, 11, 50, 60, 15, 11, 11, 11, 15]
    dicts = {"horz": "CENTER", "vert": "CENTER"}
    style2 = styles()
    style2.alignment = alignments(**dicts)
    style2.font = fonts()
    style2.borders = borders()
    style2.pattern = patterns(7)
    heights(worksheet, 0)
    for i in range(len(row0)):
        worksheet.write(0, i, row0[i], style2)
        widths(worksheet, i, size=sizes[i])
    style = styles()
    style.borders = borders()
    x = 0  # 写入数据的当前行数
    z = 0  # 用例的编号
    for i in range(len(xm["topics"])):
        test_module = xm["topics"][i]
        # print(xm["topics"])
        for j in range(len(test_module["topics"])):
            test_suit = test_module["topics"][j]
            # print(test_suit)
            for k in range(len(test_suit["topics"])):
                test_case = test_suit["topics"][k]
                # print(test_case)        # 打印所有用例
                # print(test_case.get('makers', 1))   # 打印用例的类型(makers字段的值),如果没有'makers'字段则返回1。
                z += 1
                c1 = len(test_case["topics"])  # 每条用例的执行步骤有几条
                for n in range(len(test_case["topics"])):
                    x += 1
                    test_step = test_case["topics"][n]
                    m = len(test_step["topics"])
                    txt = ""
                    for m1 in range(m):
                        test_except = test_step["topics"][m1]
                        txt += str(test_except["title"])
                        txt += '\n'
                    worksheet.write(x, 5, txt, style)  # 预期结果
                    worksheet.write(x, 4, test_step["title"], style)  # 执行步骤
                    worksheet.write(x, 3, "{}".format(n + 1), style)  # 步骤ID
                    worksheet.write(x, 6, reportName, style)  # 报告人
                    worksheet.write(x, 7, module, style)  # 模块
                    if test_case.get('makers', 1) == ['flag-red']:
                        worksheet.write(x, 8, "高", style)  # 优先级
                        worksheet.write(x, 9, "冒烟测试", style)  # 标签
                    elif test_case.get('makers', 1) == 1:
                        worksheet.write(x, 8, "中", style)
                        worksheet.write(x, 9, "功能测试", style)
                    else:
                        raise BaseException('提示：用例类型的标记不正确！')
                    worksheet.write(x, 10, xm["topics"][i]["note"], style)  # 测试用例集
                worksheet.write_merge(x - c1 + 1, x, 0, 0, z, style)  # 用例编号
                worksheet.write_merge(x - c1 + 1, x, 1, 1, test_module["title"], style)  # 测试需求名称---需求
                worksheet.write_merge(x - c1 + 1, x, 2, 2, test_case["title"], style)  # 测试用例名称---概要
    time_stamp = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime(time.time()))
    filename_xls = xm["title"] + time_stamp + ".xls"    # xls名称 = xmind的主题名称 + 当前日期和时间
    filename_csv = xm["title"] + time_stamp + '.csv'
    workbook.save(filename_xls)
    df = pd.read_excel(filename_xls, sheet_name=xm["title"], header=None)
    df.to_csv(filename_csv, header=None, sep=',', index_label=False, index=None)


# 待完善：测试用例集为空时，使用需求名称代替；用例集中的反斜杠\转为短横，避免导入时jira错误的创建多个层级的用例集
if __name__ == "__main__":
    main(r"D:\Code\Git\测试用例示例_FZ.xmind", "w_fangzheng23", "工作台")
