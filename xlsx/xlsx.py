import zipfile
import xml.dom.minidom as minidom
import tempfile
import os

shared_strings_path = "xl/sharedStrings.xml"
sheet1_path = "xl/worksheets/sheet1.xml"

class Xlsx:
    def __init__(self, filename):
        self.filename = filename
        self.zip_file = zipfile.ZipFile(filename)

        self.shared_strings_file = self.zip_file.open(shared_strings_path)
        self.sheet_file = self.zip_file.open(sheet1_path)

        self.shared_strings = minidom.parse(self.shared_strings_file)
        self.sheet = minidom.parse(self.sheet_file)

    # return string if text cell, None otherwise
    def get_content(self, cell):
        cell = cell.strip()
        c_list = self.sheet.getElementsByTagName("c")
        for c in c_list:
            if c.getAttribute("r") == cell:
                try:
                    content = c.firstChild.firstChild.data
                except AttributeError:
                    # void cell
                    return None
                if c.hasAttribute("t") and c.getAttribute("t") == "s":
                    return self.__get_string_from_index(int(content))
                else:
                    try:
                        return int(content)
                    except ValueError:
                        pass
                    try:
                        return float(content)
                    except ValueError:
                        pass
                    return None
        return None

    def get_cell(self, content):
        index = self.__get_index_of_string(content)
        v_list = self.sheet.getElementsByTagName("v")
        for v in v_list:
            if str(index) == v.firstChild.data:
                c =  v.parentNode
                if c.hasAttribute("t") and c.getAttribute("t") == "s":
                    return c.getAttribute("r")
        return None

    # return index that xlsx uses for internal mapping
    def __get_index_of_string(self, string):
        t_list = self.shared_strings.getElementsByTagName("si")
        for i in range(0, len(t_list)):
            if string in t_list[i].firstChild.firstChild.data:
                return i
        return None

    def __get_string_from_index(self, index):
        t_list = self.shared_strings.getElementsByTagName("si")
        return t_list[index].firstChild.firstChild.data

    def __del__(self):
        self.shared_strings_file.close()
        self.sheet_file.close()
        self.zip_file.close()

    # can only write string cells
    def write_cell(self, cell, string):
        # Extract xlsx in a temp folder
        with tempfile.TemporaryDirectory() as tmpdirname:
            # Created temporary directory tmpdirname
            pass
        zip_file = zipfile.ZipFile(self.filename)
        zip_file.extractall(tmpdirname)
        zip_file.close()

        # Change the string file
        cell = cell.strip()
        c_list = self.sheet.getElementsByTagName("c")
        for c in c_list:
            if c.getAttribute("r") == cell:
                try:
                    content = c.firstChild.firstChild.data
                except AttributeError:
                    # void cell
                    return False
                if c.hasAttribute("t") and c.getAttribute("t") == "s":
                    self.__write_string_from_index(int(content), string)

        # Open string file in write
        shared_strings_long_path = os.path.join(tmpdirname, shared_strings_path)
        with open(shared_strings_long_path, "w") as shared_strings_file:
            shared_strings_file.write(self.shared_strings.toxml())

        # Zip content of the folder overwritting old file
        zip_file = zipfile.ZipFile(self.filename, "w")

        for root, dirs, files in os.walk(tmpdirname):
            for file in files:
                # remove path up to tmpdirname
                complete_path = os.path.join(root, file)
                rel_path = os.path.relpath(complete_path, tmpdirname)
                zip_file.write(complete_path, arcname=rel_path)


    def __write_string_from_index(self, index, string):
        t_list = self.shared_strings.getElementsByTagName("si")
        t_list[index].firstChild.firstChild.data = string
