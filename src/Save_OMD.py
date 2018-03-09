from datetime import date
from docx import Document
import os
import comtypes.client

# path
def save_omd_as_docx(company_name):
    document = Document(r"C:\Users\M261651\Desktop\Dokumente\Files für Pycharm Projekte\OMD Prozess\OMDLetter\OMD Vorlage\TCUST_standard_disclosure_letter_company_and_items_filled.docx")
    todays_date = str(date.today()).replace("-", "")
    country = " US"
    omd = " OMD "
    docx = ".docx"
    pdf = ".pdf"
    company_name_docx = company_name + country + omd + todays_date + docx
    company_name_docx = generate_docx_name(company_name, country, omd, todays_date, docx)
    company_name_pdf = company_name + country + omd + todays_date + pdf

    path = "C:\\Users\M261651\Desktop\Dokumente\Files für Pycharm Projekte\OMD Prozess\OMDLetter\Fertige OMDs\\"
    file_name_docx = path + company_name_docx

    file_name_pdf = path + company_name_pdf
    print(" ")
    document.save(file_name_docx)
    save_omd_as_pdf(file_name_docx, file_name_pdf)


def save_omd_as_pdf(file_name_docx, file_name_pdf):
    wdFormatPDF = 17
    in_file = os.path.abspath(file_name_docx)
    out_file = os.path.abspath(file_name_pdf)

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


# path
def generate_docx_name(company_name, country, omd, todays_date, docx):
    file_docx_name = company_name + country + omd + todays_date
    file_docx_name2 = company_name + country + omd + todays_date
    all_file_names = next(os.walk(r"C:\Users\M261651\Desktop\Dokumente\Files für Pycharm Projekte\OMD Prozess\OMDLetter\Fertige OMDs\\"))[2]
    count_of_files = len(all_file_names)
    count_equal_files = []
    count_equal_files2 = []
    # print("file_docx_name2: ", file_docx_name2)
    # print("all_file_names: ", all_file_names)
    if company_name in str(all_file_names):
        for i in range(0,count_of_files):
            print("i: ", i)
            if file_docx_name2 in all_file_names[i]:
                # print("Index: ", i, " ", all_file_names[i])
                count_equal_files.append("true")
                count_equal_files2.append(count_equal_files[-1])
                number = len(count_equal_files2)
                number_iterate = number + 1
                # print("number: ",number)
                file_docx_name = company_name + country + omd + todays_date + " " + str(number_iterate)
    else:
        print("not found")
    return file_docx_name + docx



# na du
# asdsad