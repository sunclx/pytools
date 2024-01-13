"""
用MS Office 转换文件
"""
from win32com import client
from pathlib import Path


def word_convert(
    path, before_suffix=".doc", after_suffix=".docx", after_suffix_num=None
):
    """
    # 使用Word.SaveAs2函数转换文件
    https://learn.microsoft.com/zh-cn/office/vba/api/Word.SaveAs2
    # suffix number 表格
    https://learn.microsoft.com/zh-cn/office/vba/api/word.wdsaveformat

    名称	值	Description
    wdFormatDocument	0	Microsoft Office Word 97 - 2003 binary file format.
    wdFormatDOSText	4	Microsoft DOS text format.
    wdFormatDOSTextLineBreaks	5	Microsoft DOS text with line breaks preserved.
    wdFormatEncodedText	7	Encoded text format.
    wdFormatFilteredHTML	10	Filtered HTML format.
    wdFormatFlatXML	19	Open XML file format saved as a single XML file.
    wdFormatFlatXMLMacroEnabled	20	Open XML file format with macros enabled saved as a single XML file.
    wdFormatFlatXMLTemplate	21	Open XML template format saved as a XML single file.
    wdFormatFlatXMLTemplateMacroEnabled	22	Open XML template format with macros enabled saved as a single XML file.
    wdFormatOpenDocumentText	23	OpenDocument Text format.
    wdFormatHTML	8	Standard HTML format.
    wdFormatRTF	6	Rich text format (RTF).
    wdFormatStrictOpenXMLDocument	24	Strict Open XML document format.
    wdFormatTemplate	1	Word template format.
    wdFormatText	2	Microsoft Windows text format.
    wdFormatTextLineBreaks	3	Windows text format with line breaks preserved.
    wdFormatUnicodeText	7	Unicode text format.
    wdFormatWebArchive	9	Web archive format.
    wdFormatXML	11	Extensible Markup Language (XML) format.
    wdFormatDocument97	0	Microsoft Word 97 document format.
    wdFormatDocumentDefault	16	Word default document file format. For Word, this is the DOCX format.
    wdFormat PDF	17	PDF format.
    wdFormatTemplate97	1	Word 97 template format.
    wdFormatXMLDocument	12	XML document format.
    wdFormatXMLDocumentMacroEnabled	13	XML document format with macros enabled.
    wdFormatXMLTemplate	14	XML template format.
    wdFormatXMLTemplateMacroEnabled	15	XML template format with macros enabled.
    wdFormatXPS	18	XPS format.
    """
    suffix_dict = {
        ".doc": 0,
        ".docx": 12,
        ".docm": 13,
        ".dot": 1,
        ".dotx": 14,
        ".dotm": 15,
        ".mht": 9,
        ".html": 8,
        ".odt": 24,
        ".rtf": 6,
        ".xml": 11,
        ".pdf": 17,
        ".txt": 7,
        ".xps": 18,
    }
    suffix_num_dict = {
        # wdFormatDocument	0	Microsoft Office Word 97 - 2003 binary file format.
        0: "doc",
        # wdFormatDOSText	4	Microsoft DOS text format.
        4: "txt",
        # wdFormatDOSTextLineBreaks	5	Microsoft DOS text with line breaks preserved.
        5: "txt",
        # wdFormatEncodedText	7	Encoded text format.
        7: "txt",
        # wdFormatFilteredHTML	10	Filtered HTML format.
        10: "html",
        # wdFormatFlatXML	19	Open XML file format saved as a single XML file.
        19: "xml",
        # wdFormatFlatXMLMacroEnabled	20	Open XML file format with macros enabled saved as a single XML file.
        20: "xml",
        # wdFormatFlatXMLTemplate	21	Open XML template format saved as a XML single file.
        21: "xml",
        # wdFormatFlatXMLTemplateMacroEnabled	22	Open XML template format with macros enabled saved as a single XML file.
        22: "xml",
        # wdFormatOpenDocumentText	23	OpenDocument Text format.
        23: "odt",
        # wdFormatHTML	8	Standard HTML format.
        8: "html",
        # wdFormatRTF	6	Rich text format (RTF).
        6: "rtf",
        # wdFormatStrictOpenXMLDocument	24	Strict Open XML document format.
        24: "xml",
        # wdFormatTemplate	1	Word template format.
        1: "dot",
        # wdFormatText	2	Microsoft Windows text format.
        2: "txt",
        # wdFormatTextLineBreaks	3	Windows text format with line breaks preserved.
        3: "txt",
        # wdFormatUnicodeText	7	Unicode text format.
        # 7:"txt",
        # wdFormatWebArchive	9	Web archive format.
        9: "mht",
        # wdFormatXML	11	Extensible Markup Language (XML) format.
        11: "xml",
        # wdFormatDocument97	0	Microsoft Word 97 document format.
        # 0:"doc",
        # wdFormatDocumentDefault	16	Word default document file format. For Word, this is the DOCX format.
        16: "docx",
        # wdFormat PDF	17	PDF format.
        17: "pdf",
        # wdFormatTemplate97	1	Word 97 template format.
        # 1:"dot",
        # wdFormatXMLDocument	12	XML document format.
        12: "docx",
        # wdFormatXMLDocumentMacroEnabled	13	XML document format with macros enabled.
        13: "docm",
        # wdFormatXMLTemplate	14	XML template format.
        14: "dotx",
        # wdFormatXMLTemplateMacroEnabled	15	XML template format with macros enabled.
        15: "dotm",
        # wdFormatXPS	18	XPS format.
        18: "xps",
    }
    path = Path(path)
    if not path.exists():
        print(f"{path} 不存在")
        return
    if path.is_file():
        paths = [path]
        paths = [p for p in paths if p.suffix.lower() == before_suffix]
    else:
        paths = list(path.glob(f"*{before_suffix}"))

    # 打开word应用程序
    word = client.Dispatch("Word.Application")
    for path in paths:
        path = path.resolve()
        fn = str(path)
        print(f"正在转换文件: {fn}...")
        doc = word.Documents.Open(fn)  # 打开word文件
        if after_suffix_num is None:
            after_suffix_num = suffix_dict[after_suffix]
        else:
            after_suffix = suffix_num_dict[after_suffix_num]
        newname = f"{path.with_suffix(after_suffix)}"
        doc.SaveAs(newname, after_suffix_num)
        print(f"文件转换完成: {newname}.")
        doc.Close()  # 关闭原来word文件
    word.Quit()

    return


# 转换文档为docx
def convert2docx(path, suffix=".doc"):
    """转换文档为docx"""
    word_convert(path, before_suffix=suffix, after_suffix=".docx", after_suffix_num=12)
    # path = Path(path)
    # if not path.exists():
    #     print(f"{path} 不存在")
    #     return
    # if path.is_file():
    #     paths = [path]
    #     paths = [p for p in paths if p.suffix.lower() == suffix]
    # else:
    #     paths = list(path.glob(f"*{suffix}"))

    # # 打开word应用程序
    # word = client.Dispatch("Word.Application")
    # for path in paths:
    #     path = path.resolve()
    #     fn= str(path)
    #     print(f"正在转换文件: {fn}...")
    #     doc = word.Documents.Open(fn)  # 打开word文件
    #     newname = f"{path.with_suffix(".docx")}"
    #      # 另存为后缀为".docx"的文件，其中参数12或16指docx文件
    #      # https://learn.microsoft.com/zh-cn/office/vba/api/word.wdsaveformat
    #     doc.SaveAs(newname, 12)
    #     print(f"文件转换完成: {newname}.")
    #     doc.Close()  # 关闭原来word文件
    # word.Quit()
    return


# 转换文档为mht
def convert2mht(path, suffix=".doc"):
    """转换文档为mht"""
    word_convert(path, before_suffix=suffix, after_suffix=".mht", after_suffix_num=9)
    # path = Path(path)
    # if not path.exists():
    #     print(f"{path} 不存在")
    #     return
    # if path.is_file():
    #     paths = [path]
    #     paths = [p for p in paths if p.suffix.lower() == suffix]
    # else:
    #     paths = list(path.glob(f"*{suffix}"))

    # # 打开word应用程序
    # word = client.Dispatch("Word.Application")
    # for path in paths:
    #     path = path.resolve()
    #     fn= str(path)
    #     print(f"正在转换文件: {fn}...")
    #     doc = word.Documents.Open(fn)  # 打开word文件
    #     newname = f"{path.with_suffix(".mht")}"
    #      # 另存为后缀为".mht"的文件，其中参数9或指mht文件
    #      # https://learn.microsoft.com/zh-cn/office/vba/api/word.wdsaveformat
    #     doc.SaveAs(newname, 9)
    #     print(f"文件转换完成: {newname}.")
    #     doc.Close()  # 关闭原来word文件
    # word.Quit()
    return


def excel_convert(
    path, before_suffix=".xls", after_suffix=".xlsx", after_suffix_num=None
):
    """
    # 使用Word.SaveAs2函数转换文件
    https://learn.microsoft.com/zh-cn/office/vba/api/excel.workbook.saveas
    # suffix number 表格
    https://learn.microsoft.com/zh-cn/office/vba/api/excel.xlfileformat
    名称	值	说明	扩展名
    xlAddIn	18	Microsoft Excel 97-2003 外接程序	*.xla
    xlAddIn8	18	Microsoft Excel 97-2003 外接程序	*.xla
    xlCSV	6	CSV	*.csv
    xlCSVMac	22	Macintosh CSV	*.csv
    xlCSVMSDOS	24	MSDOS CSV	*.csv
    xlCSVUTF8	62	UTF8 CSV	*.csv
    xlCSVWindows	23	Windows CSV	*.csv
    xlCurrentPlatformText	-4158	当前平台文本	*.txt
    xlDBF2	7	Dbase 2 格式	*.dbf
    xlDBF3	8	Dbase 3 格式	*.dbf
    xlDBF4	11	Dbase 4 格式	*.dbf
    xlDIF	9	数据交换格式	*.dif
    xlExcel12	50	Excel 二进制工作簿	*.xlsb
    xlExcel2	16	Excel 版本 2.0 (1987)	*.xls
    xlExcel2FarEast	27	Excel 版本 2.0 Asia (1987)	*.xls
    xlExcel3	29	Excel 版本 3.0 (1990)	*.xls
    xlExcel4	33	Excel 版本 4.0 (1992)	*.xls
    xlExcel4Workbook	35	Excel 版本 4.0 工作簿格式 (1992)	*.xlw
    xlExcel5	39	Excel 版本 5.0 (1994)	*.xls
    xlExcel7	39	Excel 95（版本 7.0）	*.xls
    xlExcel8	56	Excel 97-2003 工作簿	*.xls
    xlExcel9795	43	Excel 版本 95 和 97	*.xls
    xlHtml	44	HTML 格式	*.htm；*.html
    xlIntlAddIn	26	国际外接程序	无文件扩展名
    xlIntlMacro	25	国际宏	无文件扩展名
    xlOpenDocumentSpreadsheet	60	OpenDocument 电子表格	*.ods
    xlOpenXMLAddIn	55	Open XML 外接程序	*.xlam
    xlOpenXMLStrictWorkbook	61 (&H3D)	Strict Open XML 文件	*.xlsx
    xlOpenXMLTemplate	54	Open XML 模板	*.xltx
    xlOpenXMLTemplateMacroEnabled	53	启用 Open XML 模板宏	*.xltm
    xlOpenXMLWorkbook	51	Open XML 工作簿	*.xlsx
    xlOpenXMLWorkbookMacroEnabled	52	启用 Open XML 工作簿宏	*.xlsm
    xlSYLK	2	符号链接格式	*.slk
    xlTemplate	17	Excel 模板格式	*.xlt
    xlTemplate8	17	模板 8	*.xlt
    xlTextMac	19	Macintosh 文本	*.txt
    xlTextMSDOS	21	MSDOS 文本	*.txt
    xlTextPrinter	36	打印机文本	*.prn
    xlTextWindows	20	Windows 文本	*.txt
    xlUnicodeText	42	Unicode 文本	无文件扩展名；*.txt
    xlWebArchive	45	Web 档案	*.mht；*.mhtml
    xlWJ2WD1	14	日语 1-2-3	*.wj2
    xlWJ3	40	日语 1-2-3	*.wj3
    xlWJ3FJ3	41	日语 1-2-3 格式	*.wj3
    xlWK1	5	Lotus 1-2-3 格式	*.wk1
    xlWK1ALL	31	Lotus 1-2-3 格式	*.wk1
    xlWK1FMT	30	Lotus 1-2-3 格式	*.wk1
    xlWK3	15	Lotus 1-2-3 格式	*.wk3
    xlWK3FM3	32	Lotus 1-2-3 格式	*.wk3
    xlWK4	38	Lotus 1-2-3 格式	*.wk4
    xlWKS	4	Lotus 1-2-3 格式	*.wks
    xlWorkbookDefault	51	默认工作簿	*.xlsx
    xlWorkbookNormal	-4143	常规工作簿	*.xls
    xlWorks2FarEast	28	Microsoft Works 2.0 亚洲格式	*.wks
    xlWQ1	34	Quattro Pro 格式	*.wq1
    xlXMLSpreadsheet	46	XML 电子表格	*.xml
    """
    suffix_dict = {
        # xlAddIn	18	Microsoft Excel 97-2003 外接程序	*.xla
        # ".xla":18,
        # # xlAddIn8	18	Microsoft Excel 97-2003 外接程序	*.xla
        # ".xla":18,
        # xlCSV	6	CSV	*.csv
        # ".csv":6,
        # xlCSVMac	22	Macintosh CSV	*.csv
        # ".csv":22,
        # xlCSVMSDOS	24	MSDOS CSV	*.csv
        # ".csv":24,
        # xlCSVUTF8	62	UTF8 CSV	*.csv
        ".csv": 62,
        # xlCSVWindows	23	Windows CSV	*.csv
        # ".csv":23,
        # xlCurrentPlatformText	-4158	当前平台文本	*.txt
        # ".txt":-4158,
        # xlDBF2	7	Dbase 2 格式	*.dbf
        # ".dbf":7,
        # xlDBF3	8	Dbase 3 格式	*.dbf
        # ".dbf":8,
        # xlDBF4	11	Dbase 4 格式	*.dbf
        ".dbf": 11,
        # xlDIF	9	数据交换格式	*.dif
        ".dif": 9,
        # xlExcel12	50	Excel 二进制工作簿	*.xlsb
        ".xlsb": 50,
        # xlExcel2	16	Excel 版本 2.0 (1987)	*.xls
        # ".xls":16,
        # xlExcel2FarEast	27	Excel 版本 2.0 Asia (1987)	*.xls
        # ".xls":27,
        # xlExcel3	29	Excel 版本 3.0 (1990)	*.xls
        # ".xls":29,
        # xlExcel4	33	Excel 版本 4.0 (1992)	*.xls
        # ".xls":33,
        # xlExcel4Workbook	35	Excel 版本 4.0 工作簿格式 (1992)	*.xlw
        ".xlw": 35,
        # xlExcel5	39	Excel 版本 5.0 (1994)	*.xls
        # ".xls":39,
        # xlExcel7	39	Excel 95（版本 7.0）	*.xls
        # ".xls":39,
        # xlExcel8	56	Excel 97-2003 工作簿	*.xls
        # ".xls":56,
        # xlExcel9795	43	Excel 版本 95 和 97	*.xls
        # ".xls":43,
        # xlHtml	44	HTML 格式	*.htm；*.html
        "htm": 44,
        "html": 44,
        # xlIntlAddIn	26	国际外接程序	无文件扩展名
        ".xla": 26,
        # xlIntlMacro	25	国际宏	无文件扩展名
        # ".xla":25,
        # xlOpenDocumentSpreadsheet	60	OpenDocument 电子表格	*.ods
        ".ods": 60,
        # xlOpenXMLAddIn	55	Open XML 外接程序	*.xlam
        ".xlam": 55,
        # xlOpenXMLStrictWorkbook	61 (&H3D)	Strict Open XML 文件	*.xlsx
        # ".xlsx":61,
        # xlOpenXMLTemplate	54	Open XML 模板	*.xltx
        ".xltx": 54,
        # xlOpenXMLTemplateMacroEnabled	53	启用 Open XML 模板宏	*.xltm
        ".xltm": 53,
        # xlOpenXMLWorkbook	51	Open XML 工作簿	*.xlsx
        ".xlsx": 51,
        # xlOpenXMLWorkbookMacroEnabled	52	启用 Open XML 工作簿宏	*.xlsm
        ".xlsm": 52,
        # xlSYLK	2	符号链接格式	*.slk
        ".slk": 2,
        # xlTemplate	17	Excel 模板格式	*.xlt
        ".xlt": 17,
        # xlTemplate8	17	模板 8	*.xlt
        # ".xlt":17,
        # xlTextMac	19	Macintosh 文本	*.txt
        # ".txt":19,
        # xlTextMSDOS	21	MSDOS 文本	*.txt
        # ".txt":21,
        # xlTextPrinter	36	打印机文本	*.prn
        # ".prn":36,
        # xlTextWindows	20	Windows 文本	*.txt
        # ".txt":20,
        # xlUnicodeText	42	Unicode 文本	无文件扩展名；*.txt
        ".txt": 42,
        # xlWebArchive	45	Web 档案	*.mht；*.mhtml
        ".mht": 45,
        # xlWJ2WD1	14	日语 1-2-3	*.wj2
        ".wj2": 14,
        # xlWJ3	40	日语 1-2-3	*.wj3
        # ".wj3":40,
        # xlWJ3FJ3	41	日语 1-2-3 格式	*.wj3
        ".wj3": 41,
        # xlWK1	5	Lotus 1-2-3 格式	*.wk1
        # ".wk1":5,
        # xlWK1ALL	31	Lotus 1-2-3 格式	*.wk1
        ".wk1": 31,
        # xlWK1FMT	30	Lotus 1-2-3 格式	*.wk1
        # ".wk1":30,
        # xlWK3	15	Lotus 1-2-3 格式	*.wk3
        # ".wk3":15,
        # xlWK3FM3	32	Lotus 1-2-3 格式	*.wk3
        ".wk3": 32,
        # xlWK4	38	Lotus 1-2-3 格式	*.wk4
        ".wk4": 38,
        # xlWKS	4	Lotus 1-2-3 格式	*.wks
        # ".wks":4,
        # xlWorkbookDefault	51	默认工作簿	*.xlsx
        # ".xlsx":51,
        # xlWorkbookNormal	-4143	常规工作簿	*.xls
        ".xls": -4143,
        # xlWorks2FarEast	28	Microsoft Works 2.0 亚洲格式	*.wks
        ".wks": 28,
        # xlWQ1	34	Quattro Pro 格式	*.wq1
        ".wq1": 34,
        # xlXMLSpreadsheet	46	XML 电子表格	*.xml
        ".xml": 46,
    }
    suffix_num_dict = {
        # xlAddIn	18	Microsoft Excel 97-2003 外接程序	*.xla
        18: ".xla",
        # # xlAddIn8	18	Microsoft Excel 97-2003 外接程序	*.xla
        # 18:".xla",
        # xlCSV	6	CSV	*.csv
        6: ".csv",
        # xlCSVMac	22	Macintosh CSV	*.csv
        22: ".csv",
        # xlCSVMSDOS	24	MSDOS CSV	*.csv
        24: ".csv",
        # xlCSVUTF8	62	UTF8 CSV	*.csv
        62: ".csv",
        # xlCSVWindows	23	Windows CSV	*.csv
        23: ".csv",
        # xlCurrentPlatformText	-4158	当前平台文本	*.txt
        -4158: ".txt",
        # xlDBF2	7	Dbase 2 格式	*.dbf
        7: ".dbf",
        # xlDBF3	8	Dbase 3 格式	*.dbf
        8: ".dbf",
        # xlDBF4	11	Dbase 4 格式	*.dbf
        11: ".dbf",
        # xlDIF	9	数据交换格式	*.dif
        9: ".dif",
        # xlExcel12	50	Excel 二进制工作簿	*.xlsb
        50: ".xlsb",
        # xlExcel2	16	Excel 版本 2.0 (1987)	*.xls
        16: ".xls",
        # xlExcel2FarEast	27	Excel 版本 2.0 Asia (1987)	*.xls
        27: ".xls",
        # xlExcel3	29	Excel 版本 3.0 (1990)	*.xls
        29: ".xls",
        # xlExcel4	33	Excel 版本 4.0 (1992)	*.xls
        33: ".xls",
        # xlExcel4Workbook	35	Excel 版本 4.0 工作簿格式 (1992)	*.xlw
        35: ".xlw",
        # xlExcel5	39	Excel 版本 5.0 (1994)	*.xls
        # 39:".xls",
        # xlExcel7	39	Excel 95（版本 7.0）	*.xls
        39: ".xls",
        # xlExcel8	56	Excel 97-2003 工作簿	*.xls
        56: ".xls",
        # xlExcel9795	43	Excel 版本 95 和 97	*.xls
        43: ".xls",
        # xlHtml	44	HTML 格式	*.htm；*.html
        44: ".htm;*.html",
        # xlIntlAddIn	26	国际外接程序	无文件扩展名
        26: ".xla",
        # xlIntlMacro	25	国际宏	无文件扩展名
        25: ".xla",
        # xlOpenDocumentSpreadsheet	60	OpenDocument 电子表格	*.ods
        60: ".ods",
        # xlOpenXMLAddIn	55	Open XML 外接程序	*.xlam
        55: ".xlam",
        # xlOpenXMLStrictWorkbook	61 (&H3D)	Strict Open XML 文件	*.xlsx
        61: ".xlsx",
        # xlOpenXMLTemplate	54	Open XML 模板	*.xltx
        54: ".xltx",
        # xlOpenXMLTemplateMacroEnabled	53	启用 Open XML 模板宏	*.xltm
        53: ".xltm",
        # xlOpenXMLWorkbook	51	Open XML 工作簿	*.xlsx
        51: ".xlsx",
        # xlOpenXMLWorkbookMacroEnabled	52	启用 Open XML 工作簿宏	*.xlsm
        52: ".xlsm",
        # xlSYLK	2	符号链接格式	*.slk
        2: ".slk",
        # xlTemplate	17	Excel 模板格式	*.xlt
        17: ".xlt",
        # xlTemplate8	17	模板 8	*.xlt
        # 17:".xlt",
        # xlTextMac	19	Macintosh 文本	*.txt
        19: ".txt",
        # xlTextMSDOS	21	MSDOS 文本	*.txt
        21: ".txt",
        # xlTextPrinter	36	打印机文本	*.prn
        36: ".prn",
        # xlTextWindows	20	Windows 文本	*.txt
        20: ".txt",
        # xlUnicodeText	42	Unicode 文本	无文件扩展名；*.txt
        42: ".txt",
        # xlWebArchive	45	Web 档案	*.mht；*.mhtml
        45: ".mht",
        # xlWJ2WD1	14	日语 1-2-3	*.wj2
        14: ".wj2",
        # xlWJ3	40	日语 1-2-3	*.wj3
        40: ".wj3",
        # xlWJ3FJ3	41	日语 1-2-3 格式	*.wj3
        41: ".wj3",
        # xlWK1	5	Lotus 1-2-3 格式	*.wk1
        5: ".wk1",
        # xlWK1ALL	31	Lotus 1-2-3 格式	*.wk1
        31: ".wk1",
        # xlWK1FMT	30	Lotus 1-2-3 格式	*.wk1
        30: ".wk1",
        # xlWK3	15	Lotus 1-2-3 格式	*.wk3
        15: ".wk3",
        # xlWK3FM3	32	Lotus 1-2-3 格式	*.wk3
        32: ".wk3",
        # xlWK4	38	Lotus 1-2-3 格式	*.wk4
        38: ".wk4",
        # xlWKS	4	Lotus 1-2-3 格式	*.wks
        4: ".wks",
        # xlWorkbookDefault	51	默认工作簿	*.xlsx
        # 51:".xlsx",
        # xlWorkbookNormal	-4143	常规工作簿	*.xls
        -4143: ".xls",
        # xlWorks2FarEast	28	Microsoft Works 2.0 亚洲格式	*.wks
        28: ".wks",
        # xlWQ1	34	Quattro Pro 格式	*.wq1
        34: ".wq1",
        # xlXMLSpreadsheet	46	XML 电子表格	*.xml
        46: ".xml",
    }
    path = Path(path)
    if not path.exists():
        print(f"{path} 不存在")
        return
    if path.is_file():
        paths = [path]
        paths = [p for p in paths if p.suffix.lower() == before_suffix]
    else:
        paths = list(path.glob(f"*{before_suffix}"))

    # 打开Excel应用程序
    excel = client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    for path in paths:
        path = path.resolve()
        fn = str(path)
        print(f"正在转换文件: {fn}...")
        # 打开excel文件
        xls = excel.Workbooks.Open(fn)
        if after_suffix_num is None:
            after_suffix_num = suffix_dict[after_suffix]
        else:
            after_suffix = suffix_num_dict[after_suffix_num]
        newname = f"{path.with_suffix(after_suffix)}"
        xls.SaveAs(newname, after_suffix_num)
        print(f"文件转换完成: {newname}.")
        # 关闭原来excel文件
        xls.Close()
    excel.Quit()
    return


# 转换文档为xlsx
def convert2xlsx(path, suffix=".xls"):
    """# 转换文档为xlsx"""
    excel_convert(path, before_suffix=suffix, after_suffix=".xlsx", after_suffix_num=51)
    # path = Path(path)
    # if not path.exists():
    #     print(f"{path} 不存在")
    #     return
    # if path.is_file():
    #     paths = [path]
    #     paths = [p for p in paths if p.suffix.lower() == suffix]
    # else:
    #     paths = list(path.glob(f"*{suffix}"))

    # # 打开Excel应用程序
    # excel = client.Dispatch("Excel.Application")
    # for path in paths:
    #     path = path.resolve()
    #     fn= str(path)
    #     print(f"正在转换文件: {fn}...")
    #     xls = excel.Workbooks.Open(fn)  # 打开excel文件
    #     newname = f"{path.with_suffix(".xlsx")}"
    #      # 另存为后缀为".xlsx"的文件，其中参数51指xlsx文件
    #      # https://learn.microsoft.com/zh-cn/office/vba/api/excel.xlfileformat
    #     xls.SaveAs(newname, 51)
    #     print(f"文件转换完成: {newname}.")
    #     xls.Close()  # 关闭原来word文件
    # excel.Quit()
    return


if __name__ == "__main__":
    convert2docx(path=".")
