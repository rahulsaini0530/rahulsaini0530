from PIL import Image
from pytesseract import pytesseract
import re
import pandas as pd
from urllib.request import Request, urlopen
import urllib.request
from bs4 import BeautifulSoup as soup
import requests
import os
import fitz
import io   
from urllib.parse import urlparse
from urllib.parse import parse_qs
import csv
import ntpath


file_location = "US 01032019 31052019.xls"
sheet = pd.read_excel(file_location)
# print(sheet['Application Id'])

parent_dir = "D:\Python\scrapping"

# for id in sheet['Application Id']:
#     print("app", id)

id = 'WO2020190855'
url = "https://patentscope.wipo.int/search/en/detail.jsf?docId="+id+"&tab=PCTDOCUMENTS"
req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
webpage = urlopen(req).read()
page_soup = soup(webpage, "html.parser")
title = page_soup.find("title")
pdfs = page_soup.select("a[href$='.pdf']")

download_dir = "download-pdf"
records_dir = "records"

if not os.path.exists(download_dir):
    os.mkdir(download_dir, 777)

if not os.path.exists(records_dir):
    os.mkdir(records_dir, 777)

current_dir = os.getcwd()
download_dir_path = os.path.join(current_dir, download_dir)
records = os.path.join(current_dir, records_dir)


# recordsfolder = os.path.join(current_dir, records_dir)
# csvFolder = os.path.join(recordsfolder, recordsfolder)
print("Processing - ", id)

for container in pdfs:
    current_link = ""
    current_link = container.get('href')

    # r = requests.get("https://patentscope.wipo.int" + current_link)
    # res = requests.get("https://patentscope.wipo.int" + current_link)
    # area = open(r"https://patentscope.wipo.int/search/docs2/iasr/WO2020231547/pdf/6wvtB6_QSR3IfLgH1MsgxkgvOIX3FaA_lxRwoFuzBRgqS-u1DQ5-OPPzhH28WoEZWlUWQnwjjo0ZXJAAA4QTjZqM6ZRc6XTs5iF0IAQKaE4Pq0rq_OOJqlyqAG97tAi9;jsessionid=D2048851AD72EC68656BEB9B77FD61E0.wapp2nC?filename=US2020027068-IASR.pdf", "r")
    # durl = "https://patentscope.wipo.int/search/docs2/pct/WO2020197844/pdf/c72oxS9vc2K_wsKwQngdlEKlrzyEWVR8-A8UpgFK3eb7BOGoKxXEH5ED9-fGH2mK8BY2V2OE_28Sau5F0Zhfef6F6Z14ILOhB60XjlTVfrMvPBrLm33C6U5J6VkDjmvB?docId=id00000061102316&amp;filename=WO2020197844-IB308-20210713-2316.pdf"

    durl = "https://patentscope.wipo.int" + current_link
    # print(durl)

    parsed_url = urlparse(durl)
    filename = parse_qs(parsed_url.query)['filename'][0]

    filepath = os.path.join(download_dir_path, filename)
    fileNameWithExt = os.path.splitext(filepath)[0]
    fileNameWithoutExt = ntpath.basename(fileNameWithExt)
    # print(ntpath.basename(fileNameWithoutExt))
    
    # csvFolderName = os.path.join(download_dir_path, fileNameWithoutExt)
    
    response = urllib.request.urlopen(durl)
    with open(filename, 'wb') as file:
        file.write(response.read())
        file.close()

    pdffile = fitz.open(filename)
    for pageNumber, page in enumerate(pdffile.pages(), start=1):
        text = page.getText()
        # print(text)

    for pageNumber, page in enumerate(pdffile.pages(), start=1):
        for imgNumber, img in enumerate(page.getImageList(), start=1):
            xref = img[0]
            pix = fitz.Pixmap(pdffile, xref)

            if pix.n > 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)

            pix.writePNG(f'image_Page{pageNumber}_{imgNumber}.png')
            path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            image_path = r'image_Page' + str(pageNumber) + '_' + str(imgNumber)+'.png'

            img = Image.open(image_path)

            pytesseract.tesseract_cmd = path_to_tesseract

            text = pytesseract.image_to_string(img)
            emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
            if emails:
                # print('emails')
                # print(emails)
                csvFolderCurrentID = os.path.join(records, id)
                access = 0o755
                if not os.path.exists(csvFolderCurrentID):
                    os.makedirs(csvFolderCurrentID, access)
                
                recordsFilePath = os.path.join(csvFolderCurrentID, fileNameWithoutExt + '.csv')
                
                with open(recordsFilePath, 'w') as recordfile:
                    writer = csv.writer(recordfile)
                    rows = zip(emails)
                    for email in rows:
                        writer.writerow(email)
                        print(email)
                    recordfile.close()
                    print('Done')
            # else:
            #     print('No email found')
            
            os.remove(image_path)        
            # pix.writePNG(fileNameWithoutExt + '/image_Page{pageNumber}_{imgNumber}.png')
    pdffile.close()
    os.remove(filename)
    
# f = io.BytesIO(r.content)
# print(f)
# file = fitz.open(stream=area, filetype="pdf")
# for pageNumber, page in enumerate(f.pages(), start=1):
#     text = page.getText()
#     path = os.path.join(parent_dir, id)
#     if text:
#         if not os.path.exists(path):
#             os.mkdir(path)
#             f = open(r"parent_dir/+id+.txt", "w")
#             f.write(text)
#             # print(text)
#     else:
#         for pageNumber, page in enumerate(file.pages(), start=1):
#             for imgNumber, img in enumerate(page.getImageList(), start=1):
#                 xref = img[0]
#                 pix = fitz.Pixmap(file, xref)

#                 if pix.n > 4:
#                     pix = fitz.Pixmap(fitz.csRGB, pix)
#                 pix.writePNG(f'parent_dir/+image_Page{pageNumber}_{imgNumber}.png')


# my_url = 'https://patentscope.wipo.int/search/en/detail.jsf?docId=WO2020231547&tab=PCTDOCUMENTS'
# html = urllib3.urlopen(my_url).read()
# sopa = BeautifulSoup(html)
# current_link = ''
# for link in sopa.find_all('a'):
#     current_link = link.get('href')
#     if current_link.endswith('pdf'):
#             print('Tengo un pdf: ' + current_link)
