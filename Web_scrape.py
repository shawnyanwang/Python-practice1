__author__ = 'shawn.wang'

import os
import html2text
from urllib import urlopen
import xlsxwriter
import string


def find_suffix_file(path = "C:\Users\shawn.wang\Desktop\pages"):
    file_list = os.listdir(path)
    i = 0
    while i < len(file_list):
        item = file_list[i]
        if item.find('.htm') == -1:
            print 'delete: ', item
            file_list.remove(item)
        else:
            file_list[i] ='file:///' + path + '/' + item
            i += 1
    return file_list


def get_page_content(url):
    html = urlopen(url).read()
    return html2text.html2text(html).__str__()

# if __name__ == "__main__":
#     temp_page = get_page_content(Links[0])
#
#
#     print temp_page


def get_test_case_id_test_statuses(page):
    # print page
    test_id_start_point, i = 0, 0
    dist = {}
    req = ''

    while 1:
        test_id_start_point = page.find("testID=", test_id_start_point)+7
        if test_id_start_point == 6:
            break
        test_id_end_point = page.find(")", test_id_start_point)
        test_id_item = page[test_id_start_point:test_id_end_point]
        if len(test_id_item) > 9 and test_id_start_point != 6:
            test_id_start_point = page.find("|", test_id_start_point)+2
            test_id_end_point = page.find("|", test_id_start_point)-1
            test_id_item = page[test_id_start_point:test_id_end_point]+'-'
            test_id_start_point = page.find("|", test_id_start_point)+2
            test_id_end_point = page.find("\n", test_id_start_point)-2
            test_id_item += page[test_id_start_point:test_id_end_point]
            req = test_id_item
            continue
        test_status_start_point = page.find("|", test_id_end_point)+2
        test_status_end_point = page.find("|", test_status_start_point)-2
        test_status_item = page[test_status_start_point:test_status_end_point]
        dist[test_id_item] = test_status_item
    return dist, req
# if __name__ == "__main__":
#      dist = get_test_case_id_test_statuses(temp_page)
#      print dist.keys()
#      print dist.values()


def generate_column_name(n):
    res = ''
    while n >= 0:
        res += chr(n % 26+ord('A'))
        n /= 26
        n -= 1
    return res[::-1]


def write_xlsx_data(dist_1, req, m, worksheet):
    # print len(dist_1)
    column_num = generate_column_name(0) + str(m+1)
    worksheet.write(column_num, req)
    for i in range(0, len(dist_1)):
        column_num = generate_column_name(1) + str(i+1+m)
        # print column_num
        worksheet.write(column_num, string.atoi(dist_1.keys()[i]))
        column_num = generate_column_name(2) + str(i+1+m)
        worksheet.write(column_num, dist_1.values()[i])
    return len(dist_1)



def get_id_statuses_xlsx(path, file_name):
    path = path + "/"
    workbook = xlsxwriter.Workbook(path+file_name)
    worksheet = workbook.add_worksheet()
    Links = find_suffix_file(path)
    # print Links
    i, m = 0, 0
    for item in Links:
        page = get_page_content(item)
        content, req = get_test_case_id_test_statuses(page)
        m += write_xlsx_data(content, req, m, worksheet)
        i += 3




if __name__ == "__main__":
    get_id_statuses_xlsx("C:\Users\shawn.wang\Desktop\pages\BCM_PROXY_POPUP", 'demo.xlsx')


