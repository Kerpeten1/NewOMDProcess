import json


def fill_in_company_into_txt(company):
    with open("OMDLetter\cxt Dateien\company.txt", "w") as file:
        file.write(json.dumps(company))
        file.close()
    with open("OMDLetter\cxt Dateien\company.txt", "r") as file:
        first_line = file.readline()
        file.close()
        return first_line


def fill_in_item_into_txt(item):
    with open("OMDLetter\cxt Dateien\items.txt", "w") as file:
        file.write(json.dumps(item))
        file.close()
    with open("OMDLetter\cxt Dateien\items.txt", "r") as file:
        first_line = file.readline()
        file.close()
        return first_line