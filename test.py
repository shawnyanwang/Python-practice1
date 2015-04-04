__author__ = 'shawn.wang'
import html2text
from urllib import urlopen

url = "http://qc.carefusion.com/qcreporting/td/"
html = urlopen(url).read()


url = "http://qc.carefusion.com/qcreporting/td/TheOne_TestCaseApprovalReport.asp?folder=&designer=&testfilter=5072&Submit=Submit"
html = urlopen(url).read()
print html2text.html2text(html).__str__()

# def generate_column_name(n):
#     i, m = 0, n
#     res = ''
#     while n >= 0:
#         res += chr(n % 26+ord('A'))
#         n /= 26
#         n -= 1
#     return res[::-1]
#
# for i in range(0, 26*26+26*26):
#     print generate_column_name(i)
#
#
# # print generate_column_name(1+26*28)
# print str(26)
# # print chr(ord('d')+2)