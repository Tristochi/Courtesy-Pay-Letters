from __future__ import print_function
from datetime import date
from mailmerge import MailMerge
import os   

text_file = open('J:\cpletters.txt', 'r')
courtesy_pay_list = text_file.readlines()
text_file.close()
delimeted_text = [line.split('\t') for line in courtesy_pay_list]
header = delimeted_text[0]
fifteen_days_late = []
fifteen_days_late.append(header)
twentyfive_days_late = []
twentyfive_days_late.append(header)
thirtyfive_days_late = []
thirtyfive_days_late.append(header)
for row in delimeted_text[1:]:
    days_delinquent = int(row[5])
    if days_delinquent == 15:
        fifteen_days_late.append(row)
    elif days_delinquent == 25:
        twentyfive_days_late.append(row)
    elif days_delinquent == 35:
        thirtyfive_days_late.append(row)

out15 = open('J:\cpletters15.txt', 'w+')
out25 = open('J:\cpletters25.txt', 'w+')
out35 = open('J:\cpletters35.txt', 'w+')


for line in fifteen_days_late:
    out15.write('\t'.join(line))
    out15.write('\n')
out15.close()

for line in twentyfive_days_late:
    out25.write('\t'.join(line))
    out25.write('\n')
   
out25.close()

for line in thirtyfive_days_late:
    out35.write('\t'.join(line))
    out35.write('\n')
out35.close()

template = 'J:\courtesy_pay_15_test.docx'
document = MailMerge(template)

cpayltr15 = {
    'MEMBER_NBR': 
    }




