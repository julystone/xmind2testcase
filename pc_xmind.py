import xmind

workbook = xmind.load('./tempFiles/testcases.xmind')
sheet = workbook.getPrimarySheet()
# print(workbook.getData())
# print(workbook.to_prettify_json())

print(sheet.getData())
