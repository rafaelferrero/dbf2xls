# -*- coding: UTF-8 -*-
import pyexcel as pe
from os import listdir, path, makedirs
from dbfread import DBF


def dbf2xls(input_file, output_file):
    dbf = DBF(input_file, load=True, encoding='windows-1252')
    data = []
    data.append(dbf.field_names)
    for record in dbf:
        data.append(list(record.values()))
    pe.save_as(array=data, dest_file_name=output_file)


def main():

    files_i = listdir('./dbf-directory/SICA/')
    files_o = listdir('./xls-directory/SICA/')

    try:
        opts, _ = getopt.getopt(sys.argv[1:], 'i:o:', ['file_dir_input=', 'file_dir_output='])
    except getopt.GetoptError:
        sys.exit(2)

    for opt, arg in opts:
        if opt in ('-i', '--file_dir_input'):
            files_i = arg
        if opt in ('-o', '--file_dir_output'):
            files_o = arg

    for file_i in files_i:
        output_name = path.splitext(file)[0]
        dbf2xls(files_i + file_i, files_o + output_name + '.xls')


if __name__ == "__main__":
    main()
