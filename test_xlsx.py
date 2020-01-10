import xlsx

DATABASE = "pwddatabase.xlsx"
PWD_COL = "B"
xlsx = xlsx.Xlsx(DATABASE)

while True:
    name = raw_input("Give me the name whose password you wanna know: ")
    cell = xlsx.get_cell(name)
    if None != cell:
        break
    else:
        print "Couldn't find name, are you sure it is correct? try again."

row = ""
for char in cell:
	if True == char.isdigit():
		row += char

print xlsx.get_content(PWD_COL + row)
print "Bye!"
