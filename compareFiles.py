import pandas as pd
import numpy as np
import difflib
from pdf2docx import parse
import docx2txt
import codecs
import sys
def process_excel(file1, file2,userId):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    compare_values = df1.values == df2.values
    rows, cols = np.where(compare_values == False)

    for item in zip(rows, cols):
        df1.iloc[item[0], item[1]] = ' {} --> {} '.format(df1.iloc[item[0], item[1]], df2.iloc[item[0], item[1]])
    filename = f"Tempchunks/{userId}_output.xlsx"
    df1.to_excel(filename, index=False, header=True)
    print(filename)
def process_txt(file1, file2,userId):

    first_file_lines=open(file1).readlines()
    second_file_lines= open(file2).readlines()
    difference =difflib.HtmlDiff().make_file(first_file_lines, second_file_lines, first_file, second_file)
    filename = f"Tempchunks/{userId}_difference_report.html"
    difference_report= open(filename, 'w')
    difference_report.write(difference)
    difference_report.close()
    print(filename)

def process_docx(file1, file2,userId):

    first_file =  docx2txt.process(file1)
    second_file =docx2txt.process(file2)
    with open("temp1.txt", "w") as text_file:
        print(first_file, file=text_file)
    with open("temp2.txt", "w") as text_file:
        print(second_file, file=text_file)

    first_file_lines=open("temp1.txt").readlines()
    second_file_lines= open("temp2.txt").readlines()
    difference = difflib.HtmlDiff().make_file(first_file_lines, second_file_lines, file1, file2, charset='utf-8')
    filename = f"Tempchunks/{userId}_difference_report.html"
    with codecs.open(filename, 'w', "utf-8") as difference_report:
        difference_report.write(difference)
    print("difference_report.html")
def process_pdf(file1, file2,userId):
    parse(file1, 'temp1.docx', start=0, end=None)
    parse(file2, 'temp2.docx', start=0, end=None)

    first_txt_file =  docx2txt.process("temp1.docx")
    second_txt_file =docx2txt.process("temp2.docx")
   

    with open("temp1.txt", "w") as text_file:
        print(first_txt_file, file=text_file)
    with open("temp2.txt", "w") as text_file:
        print(second_txt_file, file=text_file)

    first_file_lines=open("temp1.txt").readlines()
    second_file_lines= open("temp2.txt").readlines()
    difference = difflib.HtmlDiff().make_file(first_file_lines, second_file_lines, file1, file2, charset='utf-8')
    filename = f"Tempchunks/{userId}_difference_report.html"
    with codecs.open(filename, 'w', "utf-8") as difference_report:
        difference_report.write(difference)
    print("difference_report.html")
# .NET tarafından tip parametresi ve dosya yollarını alacağım
 # Örneğin, 1: Excel, 2: TXT, 3: DOCX, 4: PDF şeklinde kurguluyorum
if len(sys.argv) == 5:
    # python_exe_dosyasi.exe dosya1.xlsx dosya2.xlsx 1
    dosya_tipi = int(sys.argv[3])
    ilk_dosya_yolu = sys.argv[1]
    ikinci_dosya_yolu = sys.argv[2]
    userId=sys.argv[4]
    if dosya_tipi == 1:
        process_excel(ilk_dosya_yolu, ikinci_dosya_yolu,userId)
    elif dosya_tipi == 2:
        process_txt(ilk_dosya_yolu, ikinci_dosya_yolu,userId)
    elif dosya_tipi == 3:
        process_docx(ilk_dosya_yolu, ikinci_dosya_yolu,userId)
    elif dosya_tipi == 4:
        process_pdf(ilk_dosya_yolu, ikinci_dosya_yolu,userId)
    else:
        print("Geçersiz dosya tipi.")
else:
    print("Lütfen dosya yollarını ve dosya tipini belirtin.")