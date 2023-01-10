import os
import re
import openpyxl
from pdfminer.high_level import extract_pages, extract_text

def os_dir_path():
    path = os.getcwd()
    files = os.listdir(path=path)
    files = ' '.join(files)
    match = re.findall(r"\d+\s+\w+\s+\w+\.pdf",files)
    file_name = match[0]
    pdf_path = path +"\\" + file_name
    return pdf_path

def main():
    wrightFile(getInfoFromInv(extractTextFromPdf(os_dir_path())))
    remove_tmp_files()


def extractTextFromPdf(pdf_path):
    text = extract_text(pdf_path)
    with open("text.txt", 'w+', encoding='utf-8') as file:
        file.write(text)
        file.close()
    text = open('text.txt').read()
    return text


def getInfoFromInv(text):
    art_list = []
    country_list = []
    qty_list=[]
    summ_list = []
    total_list=[]

    match =re.findall(r'\d{6,}\s\d{2}',text)
    for x in match:
        art_list.append(x.replace('\n',''))

    match = re.findall(r'\d{9}0\s{2}\w{,2}', text)
    for x in match:
            country_list.append(x[-2::])


    match = re.findall(r'\d+[,]\d+\sPC|\d+[,]\d+\sSET|\d+PC|\d+\sSET|\d+\sPC', text)
    for x in match:
        x = x.replace('PC','')
        x = x.replace('SET','')
        qty_list.append(x.replace(',',''))

    match = re.findall(r'\d*[,]\d*[.]\d{2}\s*EUR|\d*[.]\d{2}\s*EUR', text)
    for x in match:
        summ_list.append(x[:-5].replace(',',''))


    total_list.append(art_list)
    total_list.append(qty_list)
    total_list.append(country_list)
    total_list.append(summ_list[:-2])
    return total_list

def wrightFile(total_list):
   book = openpyxl.Workbook()
   sheet = book.active
   for i in range(len(total_list)):
       for info in range(len(total_list[i])):
           sheet.cell(row=i+1, column=info+1)
           sheet[info+2][i].value = total_list[i][info]
   headers_name_list = ['art', 'qty', 'country', 'total']
   for x in range(len(headers_name_list)):
       sheet[1][x].value = headers_name_list[x]
   book.save("Invoice.xlsx")
   book.close()

def remove_tmp_files():
    path = os.getcwd()
    del_file = path + "\\" + 'text.txt'
    os.remove(del_file)
    print("tmp_file text.txt removed successfull")



if __name__ == '__main__':
    main()
