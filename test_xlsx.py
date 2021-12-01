import xlsx
import pandas
import shutil

DATABASE_BACKUP = "pwddatabase.xlsx"
DATABASE = "pwddatabase.tmp.xlsx"
# Copy the file to something else that I do not keep track of and I feel
# free to edit
shutil.copy(DATABASE_BACKUP, DATABASE)
PWD_COL = "B"
xlsx = xlsx.Xlsx(DATABASE)

# I change the content of the file only temporarly
xlsx.write_cell("A6", "riccardo")

while True:
    name = input("Give me the name whose password you wanna know: ")
    cell = xlsx.get_cell(name)
    if None != cell:
        break
    else:
        print("Couldn't find name, are you sure it is correct? try again.")

row = ""
for char in cell:
	if True == char.isdigit():
		row += char

print(xlsx.get_content(PWD_COL + row))

cell = "D12"

print("")
print("Content of " + cell + ":")
print(xlsx.get_content(cell))

print("Bye!")
