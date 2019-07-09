from mailmerge import MailMerge
import time
import os   

# This method uses that delimited list and writes it to a new text file.
def writeNewTxt(aList, output):
    out = open(output, 'w+')
    for line in aList:
        out.write('\t'.join(line))
    out.close()


# This method takes a file path as input and a file path to write to as output.
def sort_txt_file(input):
    text_file = open(input, 'r')
    newList = text_file.readlines()
    text_file.close()
    delimeted = [line.split('\t') for line in newList]
    header = delimeted[0]
    list1 = []
    list1.append(header)
    list2 = []
    list2.append(header)
    list3 = []
    list3.append(header)

    for row in delimeted[1:]:
        days_delinquent = int(row[5])
        if days_delinquent == 15:
           list1.append(row)
        elif days_delinquent == 25:
            list2.append(row)
        elif days_delinquent == 35:
            list3.append(row)

    writeNewTxt(list1, 'G:\exports\cpletters\cpletters15.txt')
    writeNewTxt(list2, 'G:\exports\cpletters\cpletters25.txt')
    writeNewTxt(list3, 'G:\exports\cpletters\cpletters35.txt')

    count1 = len(list1) - 1
    count2 = len(list2) - 1
    count3 = len(list3) - 1
    if count1 > 1:
        print("There are ", count1, "records on courtestpay15.txt")
    elif count1 == 1:
        print("There is ", count1, " record on courtesypay15.txt")
    if count2 > 1:
        print("There are ", count2, "records on courtestpay25.txt")
    elif count2 == 1:
        print("There is ", count2, "record on courtesypay25.txt")
    if count3 > 1:
        print("There are ", count3, "records on courtestpay35.txt")
    elif count3 == 1:
        print("There is ", count3, "record on courtesypay35.txt")
    return None

#This requires a docx for templatePath and outputPath. Both need to have merge fields already inplace.
#The txtFilePath takes the path of one previously sorted txt file.
def writeToDocx(txtFilePath, templatePath, outputPath):
    text_file = open(txtFilePath, 'r')
    newList = text_file.readlines()
    text_file.close()
    delimeted = [line.split('\t') for line in newList]
    header = delimeted[0]
    values = delimeted[1:]
    template = templatePath
    document = MailMerge(template)

    big_Dict = ([{head:val for head, val in zip(header, val)} for val in values])
    document.merge_templates(big_Dict, 'nextPage_section')
    document.write(outputPath)

def getRecordsCount(txtFilePath):
    text_file = open(txtFilePath, 'r')
    newList = text_file.readlines()
    text_file.close()
    delimeted = [line.split('\t') for line in newList]
    header = delimeted[0]
    values = delimeted[1:]

    return int(len(values))

def main():
    sort_txt_file('G:\exports\cpletters\cpletters.txt')

    writeToDocx('G:\exports\cpletters\cpletters15.txt', 'G:\exports\cpletters\courtesy_pay_15_temp.docx', 
                'G:\exports\cpletters\courtesy_pay_15.docx')

    writeToDocx('G:\exports\cpletters\cpletters25.txt', 'G:\exports\cpletters\courtesy_pay_25_temp.docx', 
                'G:\exports\cpletters\courtesy_pay_25.docx')

    writeToDocx('G:\exports\cpletters\cpletters35.txt', 'G:\exports\cpletters\courtesy_pay_35_temp.docx', 
                'G:\exports\cpletters\courtesy_pay_35.docx')

    count1 = getRecordsCount('G:\exports\cpletters\cpletters15.txt') 
    count2 = getRecordsCount('G:\exports\cpletters\cpletters25.txt') 
    count3 = getRecordsCount('G:\exports\cpletters\cpletters35.txt') 

    if count1 > 0:
        os.startfile('G:\\exports\cpletters\courtesy_pay_15.docx', 'print')
        time.sleep(2)
        print(count1, " 15 day records printed.")
    else:
        print("No 15 day records to print.")

    if count2 > 0:
        os.startfile('G:\\exports\cpletters\courtesy_pay_25.docx', 'print')
        time.sleep(2)
        print(count2, " 25 day records printed.")
    else:
        print("No 25 day records to print.")

    if count3 > 0:
        os.startfile('G:\\exports\cpletters\courtesy_pay_35.docx', 'print')
        time.sleep(2)
        print(count3, " 35 day records printed.")
    else: 
        print("No 35 day records to print.")
        
if __name__ == "__main__":
    main()
    
