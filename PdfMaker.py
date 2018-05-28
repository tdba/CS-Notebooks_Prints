# -*- coding: utf-8 -*-

import os
import sys
import pickle
import xlrd

import cairosvg
from string import Template
from barcode import generate
from BarCodeGenerator import render

"""
Script made for the Cardboard Scents company
This extractor convert information from doctors to produce prescription books and postal labels
Should be adapted if used in another occasion
Structure used:
    - Excel file containing the doctors' data
    - Folder templates containing the SVG templates
Output:
    - Pickle file containing the doctors' data
    - Directory mail-labels 
        - Subdirectory memos 
        - Subdirectory notebooks
    - Directory notebooks
"""

g_relevant_c = [3]
h_relevant_c = [5, 6, 7, 8]
l_relevant_c = [9, 10, 11, 12, 13, 14, 15, 16]
s_relevant_c = [17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]

doctors_c = [5, 6, 7, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]
mail_c = [0, 1, 9, 10, 11, 12, 13, 14, 15]


def relevant_columns_names(num):
    """
    Switch implementation matching the column number and the column name
    :param num: The column number
    :return: The column name (used in the storage dictionaries)
    """
    return {
        0: 'g_order_number', 1: 'g_target_name', 2: 'g_language', 3: 'g_inami_number', 5: 'h_first_name',
        6: 'h_last_name', 7: 'h_inami_number', 8: 'h_bar_code', 9: 'l_first_name', 10: 'l_last_name',
        11: 'l_title', 12: 'l_street', 13: 'l_zip', 14: 'l_city', 15: 'l_institute', 16: 'l_bar_code',
        17: 's_title', 18: 's_first_name', 19: 's_last_name', 20: 's_speciality', 21: 's_institute',
        22: 's_street', 23: 's_zip', 24: 's_city', 25: 's_phone', 26: 's_inami_number', 27: 's_fax', 28: 's_email'
    }[num]


def extractor(file):
    """
    Extract data from provided file to store it in dictionary with doctors as key
    :param file: Excel file containing the doctors' data
    :return: A dictionary with the doctors' data (also exported with pickle)
    """
    doctors_hm = {}
    mail_labels_hm = {}

    workbook = xlrd.open_workbook(file)
    worksheet = workbook.sheet_by_index(0)

    for row in range(1, worksheet.nrows):
        row_values = [int(val) if type(val) == float else val for val in worksheet.row_values(row)]

        g_values = {}
        for i in g_relevant_c:
            g_values[relevant_columns_names(i)] = row_values[i]
        h_values = {}
        for i in h_relevant_c:
            h_values[relevant_columns_names(i)] = row_values[i]
        l_values = {}
        for i in l_relevant_c:
            l_values[relevant_columns_names(i)] = row_values[i]
        s_values = {}
        for i in s_relevant_c:
            s_values[relevant_columns_names(i)] = row_values[i]

        inami_key = int(h_values['h_inami_number'].replace('.', ''))
        if inami_key not in doctors_hm:
            doctors_hm[inami_key] = {'g': g_values, 'h': h_values, 'l': l_values, 's': s_values}
        if inami_key in mail_labels_hm:
            mail_labels_hm[inami_key][row_values[0]] = (row_values[1], row_values[2], row_values[16])
        else:
            mail_labels_hm[inami_key] = {row_values[0]: (row_values[1], row_values[2], row_values[16])}

    with open("doctors_hm", mode='wb+') as f:
        pickle.dump(doctors_hm, f)
    with open("mail_labels_hm", mode='wb+') as f:
        pickle.dump(mail_labels_hm, f)

    return doctors_hm, mail_labels_hm


def mail_label_maker(doctor, num, command_type, labels):
    """
    Create a pdf mail label for the corresponding command number
    :param doctor: Dictionary containing the doctor's data
    :param num: Command number
    :param command_type: Command type (memo or prescription)
    :param labels: Mail labels dictionary
    :return: -
    """
    doc = dict(list(doctor['l'].items()) + list(doctor['g'].items()))
    doc['g_order_number'] = num
    doc['g_target_name'] = command_type
    inami_number = int(doctor['h']['h_inami_number'].replace('.', ''))
    code = labels[inami_number][num][2]
    doc['l_bar_code'] = code

    if len(code) == 24:
        generate('code128', code, output='barcode')
        with open("barcode.svg", mode='r') as f:
            svg_barcode = ''.join(f.readlines()[5:])
            top = '<svg x="157px" y="117px" height="10.000mm" width="40.000mm" xmlns="http://www.w3.org/2000/svg">'
            doc['image_bar_code'] = top + svg_barcode
        os.remove('barcode.svg')

        with open("templates/mail_labels.svg", mode='r') as f:
            template = Template(f.read())
            filled_template = template.substitute(doc)
        with open(str(num) + ".svg", mode='w+') as f:
            f.write(filled_template)

        if command_type == 'algemene':
            path = 'mail_labels/notebooks/'
        else:
            path = 'mail_labels/memos/'
        cairosvg.svg2pdf(url=str(num) + '.svg', write_to=path + str(num) + '.pdf')
        labels[inami_number][num] += (path + str(num) + '.pdf',)

        os.remove(str(num) + '.svg')
    else:
        print("An error occurred, we can't provide this mail label because the given code doesn't fit the requirements"
              "\nDenied order number:" + num + '\n With code' + code + '\n\n')


def lang_prescription(doctor, file, num, lang):
    """
    Create a pdf prescription notebooks for the corresponding doctor by language
    :param doctor: Dictionary containing the doctor's data
    :param file: Template corresponding to the language
    :param num: Number of notebooks needed
    :param lang: Language of the notebooks created
    :return: -
    """
    doctor['image_bar_code'] = render('0' + str(doctor['h_bar_code'])[:-1])

    with open(file, mode='r') as f:
        template = Template(f.read())
        filled_template = template.substitute(doctor)
    with open(str(num) + ".svg", mode='w+') as f:
        f.write(filled_template)

    inami_number = int(doctor['h_inami_number'].replace('.', ''))
    for i in range(num):
        cairosvg.svg2pdf(url=str(num) + '.svg', write_to="notebooks/" + str(lang) + '.' + str(inami_number) + '.'
                                                         + str(i) + '.pdf')

    os.remove(str(num) + '.svg')


def prescription_maker(doctor, num_by_lang):
    """
    Create a pdf prescription notebooks for the corresponding doctor
    :param doctor: Dictionary containing the doctor's data
    :param num_by_lang: Number of notebooks needed by language
    :return: -
    """
    doc = dict(list(doctor['g'].items()) + list(doctor['h'].items()) + list(doctor['s'].items()))
    doc['h_last_name'] = doc['h_last_name'].upper()
    doc['h_first_name'] = doc['h_first_name'].upper()

    if num_by_lang[0] > 0:
        lang_prescription(doc, "templates/NL_prescriptions.svg", num_by_lang[0], 'NL')
    if num_by_lang[1] > 0:
        lang_prescription(doc, "templates/FR_prescriptions.svg", num_by_lang[1], 'FR')


def pdf_maker(doctors_hm, mails_hm):
    """
    Launch the creations of the needed pdf files (mail labels and prescription notebooks)
    :param doctors_hm: Dictionary containing doctors'data under the form of dictionaries by doctor
    :param mails_hm: Dictionary containing the commands number as a list by doctor
    :return: -
    """
    for inami in mails_hm.keys():
        quantity = (sum([1 for command in mails_hm[inami].values() if command[0] == 'algemene' and command[1] == 'N']),
                    sum([1 for command in mails_hm[inami].values() if command[0] == 'algemene' and command[1] != 'N']))
        doctor = doctors_hm[inami]
        for k, e in mails_hm[inami].items():
            mail_label_maker(doctor, k, e[0], mails_hm)
        prescription_maker(doctor, quantity)


if __name__ == '__main__':
    print("Initiating the extractor")
    doctors, mail_labels = extractor(sys.argv[1])
    print("Extraction achieved")
    print("Pdf creation")
    pdf_maker(doctors, mail_labels)
    with open("mail_labels_hm", mode='wb+') as hm:
        pickle.dump(mail_labels, hm)
    print("Creations achieved")
