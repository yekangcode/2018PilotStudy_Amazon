import requests
from xlrd import open_workbook
from xlutils.copy import copy
from bs4 import BeautifulSoup
import time


def spider(max_pages):
    page = 1

    filelocation = "C:/Users/Youngeun Kang/Desktop/iPhone5.xls"
    rb = open_workbook(filelocation)
    wb = copy(rb)
    s = wb.get_sheet(0)
    while page <= max_pages:
        url = "https://www.amazon.com/Apple-iPhone-16GB-Certified-Refurbished/product-reviews/B00WZR5ULU/ref=cm_cr_arp_d_paging_btm_3?ie=UTF8&reviewerType=all_reviews&pageNumber=" + str(page)
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'}
        source_code = requests.get(url, headers=headers)

        time.sleep(1)

        rating = 0
        title = 0
        text = 0

        if source_code.status_code == 200:
            plain_text = source_code.text
            soup = BeautifulSoup(plain_text, "html.parser")
            for link in soup.findAll("i", {'data-hook': 'review-star-rating'}):
                label = link.get_text().split(" ")[0]
                s.write(10*(page - 1) + rating, 1, label)
                rating += 1
            for link in soup.findAll("a", {'data-hook': 'review-title'}):
                table = link.get_text()
                s.write(10*(page - 1) + title, 2, table)
                title += 1
            for link in soup.findAll("span", {'data-hook': 'review-body'}):
                table = link.get_text()
                s.write(10*(page - 1) + text, 3, table)
                text += 1
        else:
            break
        page += 1
        print(page)
    wb.save(filelocation)  # 엑셀 파일 저장


spider(55)
print("debug")