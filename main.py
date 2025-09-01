import requests
import re
import time
import xlsxwriter

address="http://gajl.wielcy.pl/"
letters = list(map(chr, range(97, 123)))

familynames = []
herby = []

timer_start = time.time()

#Get list of familynames
for letter in letters:
    html = requests.get(address+"herby_alfa_nazwiska.php?phase=2&lang=pl&letter="+letter)
    for name in re.findall(r'List\.Add\("(.*)"\)', html.text):
        familynames.append(name.split("--h?")[0])

#Get armories for each familyname
for nazwisko in familynames:
    print(nazwisko)
    html = requests.post(address+"herby_nazwisko_herby.php?", {'nazwisko': nazwisko})
    herb = re.findall(r'<img src=".(.*?)"', html.text, )
    herb.pop(0)
    if  len(herb) > 0 and "h?.gif" not in herb[0]:
        herby.append(herb)
    else:
        herby.append([])


#Create excel file
workbook = xlsxwriter.Workbook('herby_nazwiska.xlsx')
worksheet = workbook.add_worksheet()


merge_format = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "yellow",
    }
)


for row in range(0, len(familynames)):
    worksheet.merge_range(row*2 + 1, 0, row*2 + 2, 0, familynames[row].capitalize(), merge_format)
    for col in range(0, len(herby[row])):
        worksheet.write(row*2 + 1, col + 1, herby[row][col].split("/")[-1].split(".")[0].capitalize())
        worksheet.write(row*2+2, col + 1, address + herby[row][col])
        worksheet.set_column(row, col+1, 20)


workbook.close()
timer_end = time.time()

print("Execution time:", timer_end - timer_start)


