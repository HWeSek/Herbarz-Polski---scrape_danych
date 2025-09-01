import requests
import re
import time
import xlsxwriter

address="http://gajl.wielcy.pl/"
letters = list(map(chr, range(97, 123)))

familynames = []
herby_list = []

timer_start = time.time()

#Get list of armories
for letter in letters:
    html = requests.get(address+"herby_alfa.php?phase=2&lang=pl&letter="+letter)
    for name in re.findall(r'List\.Add\("(.*)"\)', html.text):
        herby_list.append(name)

#get familyname for each armory
for herb in herby_list:
    html = requests.get(address + "herby_nazwiska.php?lang=pl&herb=" + herb).text
    html = re.sub(r'<a.*?</a>', '', html)
    html = re.sub(r'<b>', '', html)
    html = re.sub(r'</b>', '', html)
    html = re.sub(r'\.', '', html)

    fn_string = ""
    for fn in re.findall(r'<p class="indent">(.*?)<\/p>', html):
        fn_string += fn + ","
    fn_string = fn_string.replace(" ", "")
    familynames.append(fn_string.split(","))
    print(herb,fn_string.split(","))




#Create excel file
workbook = xlsxwriter.Workbook('nazwiska_herby.xlsx')
worksheet = workbook.add_worksheet()

dark_color = workbook.add_format({"bg_color": "#00AA00"})
light_color = workbook.add_format({"bg_color": "#00FF00"})

worksheet.set_column(0, 0, 40)
worksheet.set_column(0, 1000, 25)
for row in range(0, len(herby_list)):
    worksheet.write(row*2, 0, herby_list[row], dark_color)
    worksheet.write(row*2+1, 0, address + "/images/" + herby_list[row] + ".gif")
    for col in range(0, len(familynames[row])):
        worksheet.write(row*2, col+1, familynames[row][col], light_color)

workbook.close()
timer_end = time.time()

print("Execution time:", timer_end - timer_start)


