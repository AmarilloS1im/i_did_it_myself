import os
import re
import openpyxl
from pdfminer.high_level import extract_pages, extract_text


def main():
    wrightFile(getInfoFromInv(extractTextFromPdf(os_dir_path())))
    remove_tmp_files()

def os_dir_path():
    path = os.getcwd()
    files = os.listdir(path=path)
    files = ' '.join(files)
    match = re.findall(r"\w{2}\s*\d+\s+\w+\s+\w+\.pdf",files)
    file_name = match[1]
    pdf_path = path +"\\" + file_name
    return pdf_path


def extractTextFromPdf(pdf_path):
    text = extract_text(pdf_path)
    with open("text_PL.txt", 'w+', encoding='utf-8') as file:
        file.write(text)
        file.close()
    text = open('text_PL.txt').read()
    return text


def getInfoFromInv(text):
    art_list = []
    place_list = []
    gross_list = []
    net_list = []
    total_list = []

    match = re.findall(r'\d{8}\s*TU|\d{8}\s*RP|\d{8}\s*8|\d{8}\s+CO', text)
    for x in match:
        x = x.replace('\n', '')
        x = x.replace('TU', '')
        x = x.replace('RP', '')
        if len(x) == 8:
            art_list.append(x)
        else:
            art_list.append((x[:-1]))

    match = re.findall(r"\s+\d{1,3}\s+9\d{7}|\s+\d{1,3}\s+[A-z]{6,}", text)
    tmp_list = []
    for x in match[1:]:
        tmp_list.append(x)
    match = " ".join(tmp_list)
    match = re.findall(r"\s+\d{1,4}\s+",match)
    for x in match:
        x = x.replace('\x0c','')
        place_list.append(x.replace('\n',''))

    match = re.findall(r"\d*[.]\d{3}\s{1}KG|\d*\s{1}KG", text)
    match = match[2:]
    for x in range(len(match)):
        if x % 2 != 0:
            match[x] = match[x].replace(' ', '')
            match[x] = match[x].replace('\n', '')
            match[x] = match[x].replace('KG', '')
            net_list.append(match[x])
        else:
            match[x] = match[x].replace(' ', '')
            match[x] = match[x].replace('\n', '')
            match[x] = match[x].replace('KG', '')
            gross_list.append((match[x]))

    total_list.append(art_list[1:])
    total_list.append(list(filter(None, place_list)))
    total_list.append(net_list)
    total_list.append(gross_list)
    return total_list


def wrightFile(total_list):
    book = openpyxl.Workbook()
    sheet = book.active
    for i in range(len(total_list)):
        for info in range(len(total_list[i])):
            sheet.cell(row=i + 1, column=info + 1)
            sheet[info + 2][i].value = total_list[i][info]
    headers_name_list = ['art', 'place', 'net_weight', 'gross_weight']
    for x in range(len(headers_name_list)):
        sheet[1][x].value = headers_name_list[x]
    book.save("PL.xlsx")
    book.close()

def remove_tmp_files():
    path = os.getcwd()
    del_file = path + "\\" + 'text_PL.txt'
    os.remove(del_file)
    print("tmp_file text_PL.txt removed successfull")


if __name__ == '__main__':
    main()
