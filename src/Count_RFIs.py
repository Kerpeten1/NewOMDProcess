import os
import time
from src.Read_Docx import read_rfi_data


def read_name_and_count__rfis_in_folder():
    file_names = next(os.walk("RFIs"))[2]
    count_of_files = len(file_names)
    start_rfi_process(count_of_files, file_names)


def start_rfi_process(count_of_files, name):
    for i in range(0, count_of_files):
        read_rfi_data(name[i])


read_name_and_count__rfis_in_folder()