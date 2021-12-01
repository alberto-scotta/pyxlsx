import zipfile
import xml.dom.minidom as minidom

class Xlsx:
    def __init__(self, filename):
        zip_file = zipfile.ZipFile(filename)

        shared_strings = zip_file.open("xl/sharedStrings.xml")
        sheet = zip_file.open("xl/worksheets/sheet1.xml")

        self.shared_strings = minidom.parse(shared_strings)
        self.sheet = minidom.parse(sheet)

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

