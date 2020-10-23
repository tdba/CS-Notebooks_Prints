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

relevants = {'g': g_relevant_c, 'h': h_relevant_c, 'l': l_relevant_c, 's': s_relevant_c}

doctors_c = [5, 6, 7, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]
mail_c = [0, 1, 9, 10, 11, 12, 13, 14, 15]

notebooks = ["algemene", "ALGEMENE", "Algemene", "CSP100", "CSP50"]
signet = [['s_title', 's_first_name', 's_last_name'], 's_speciality', 's_institute',
          's_street', ['s_zip', 's_city'], 's_phone', 's_inami_number', 's_fax', 's_email']


def relevant_columns_names(num):
    """
    Switch implementation matching the column number and the column name
    :param num: Int
    :return: Str
    """
    return {
        0: 'g_order_number', 1: 'g_target_name', 2: 'g_language', 3: 'g_inami_number', 5: 'h_first_name',
        6: 'h_last_name', 7: 'h_inami_number', 8: 'h_bar_code', 9: 'l_first_name', 10: 'l_last_name',
        11: 'l_title', 12: 'l_street', 13: 'l_zip', 14: 'l_city', 15: 'l_institute', 16: 'l_bar_code',
        17: 's_title', 18: 's_first_name', 19: 's_last_name', 20: 's_speciality', 21: 's_institute',
        22: 's_street', 23: 's_zip', 24: 's_city', 25: 's_phone', 26: 's_inami_number', 27: 's_fax', 28: 's_email'
    }[num]


def pdf_maker(file, option):
    """
    Extract data from provided file to store it in dictionary with doctors as key
    :param file: Str (filename)
    :param option: str
    :return: -
    """

    workbook = xlrd.open_workbook(file)
    worksheet = workbook.sheet_by_index(0)

    def format_svg(values):
        """
        Transform the values to make them fit to svg standards
        :param values: List
        :return: -
        """
        for idx in range(len(values)):
            val = values[idx]
            if type(val) == str and '&' in val:
                values[idx] = val.replace('&', '&amp;')

    for row in range(1, worksheet.nrows):
        row_values = [int(val) if type(val) == float else val for val in worksheet.row_values(row)]
        format_svg(row_values)

        doc = {'l': {}, 'g': {}, 'h': {}, 's': {}}
        for letter, col_list in relevants.items():
            for i in col_list:
                doc[letter][relevant_columns_names(i)] = row_values[i]

        if 'l' in option:
            mail_label_maker(doc, row_values[16], row_values[1], str(row_values[0]))
        if 'n' in option and row_values[1] in notebooks:
            prescription_maker(doc, row_values[2], row_values[0])


def mail_label_maker(doctor, barcode_num, command_type, order_num):
    """
    Create a pdf mail label for the corresponding command number
    :param doctor: Dictionary
    :param barcode_num: Int
    :param command_type: Str
    :param order_num: Str
    :return: -
    """
    doc = dict(list(doctor['l'].items()) + list(doctor['g'].items()) + list(doctor['h'].items()))
    doc['g_order_number'] = order_num
    doc['g_target_name'] = command_type
    doc['l_bar_code'] = barcode_num

    if len(barcode_num) == 24:
        generate('code128', barcode_num, output='barcode')
        with open("barcode.svg", mode='r') as f:
            svg_barcode = ''.join(f.readlines()[5:])
            top = '<svg x="121px" y="81px" height="10.000mm" width="40.000mm" xmlns="http://www.w3.org/2000/svg">'
            doc['image_bar_code'] = top + svg_barcode
        os.remove('barcode.svg')

        with open("templates/mail_labels.svg", mode='r') as f:
            template = Template(f.read())
            filled_template = template.substitute(doc)
        with open("tmp_label.svg", mode='w+') as f:
            f.write(filled_template)

        if command_type in notebooks:
            path = 'mail_labels/notebooks/'
        else:
            path = 'mail_labels/memos/'

        file_path = path + order_num
        tmp = file_path
        i = 0
        while os.path.exists(tmp + '.pdf'):
            tmp = file_path + '_' + str(i)
            i += 1

        cairosvg.svg2pdf(url="tmp_label.svg", write_to=tmp + '.pdf')
        os.remove("tmp_label.svg")
    else:
        print("An error occurred, we can't provide this mail label because the given code doesn't fit the requirements"
              "\nDenied barcode:" + barcode_num + '\n With order number' + order_num + '\n\n')


def signet_maker(doctor):
    """
    Create from the doctor's data the svg of the signet adapted to blanks and oversized cells
    :param doctor: Dictionary
    :return: Str (svg code of the signet)
    """
    line = '<text transform ="matrix(1 0 0 1 10.3467 {})" class ="st16 {}"> {} </text>\n'
    step = 8.4004
    y = 454.2532
    result = ''

    def text_adapt(field, ordinate, bold=''):
        """
        Adapted line of the signet regarding the size of the input or its emptiness
        :param field: Str (content for the fields)
        :param ordinate: Int
        :param bold: Str
        :return: Str
        """
        if 0 < len(field) <= 40:
            return ordinate, line.format(ordinate, 'st2 ' + bold, field)
        elif len(field) > 40:
            index = field[:41].rfind(' ')
            if index == -1:
                return ordinate, line.format(ordinate, 'st4 ' + bold, field)
            else:
                first_part = line.format(ordinate, 'st2 ' + bold, field[:index])
                ordinate += step
                ordinate, second_part = text_adapt(field[index+1:], ordinate, bold)
                return ordinate, first_part + second_part
        else:
            return ordinate, field

    for e in signet:
        if type(e) == list:
            if 's_title' in e:
                y, text = text_adapt(' '.join([str(doctor[k]) for k in e]), y, 'st22')
                result += text
            else:
                y, text = text_adapt(' '.join([str(doctor[k]) for k in e]), y)
                result += text
        else:
            y, text = text_adapt(str(doctor[e]), y)
            result += text
        y += step
    return result


def lang_prescription(doctor, file, lang, order_num):
    """
    Create a pdf prescription notebooks for the corresponding doctor by language
    :param doctor: Dictionary
    :param file: Str (filename)
    :param lang: Str
    :param order_num: str
    :return: -
    """
    doctor['image_bar_code'] = render(str(doctor['h_bar_code'])[:-1])
    doctor['signet'] = signet_maker(doctor)

    with open(file, mode='r') as f:
        template = Template(f.read())
        filled_template = template.substitute(doctor)
    with open("tmp_nb.svg", mode='w+') as f:
        f.write(filled_template)

    file_path = "notebooks/" + str(lang) + '.' + str(order_num)
    tmp = file_path
    i = 0
    while os.path.exists(tmp + '.pdf'):
        tmp = file_path + '_' + str(i)
        i += 1

    cairosvg.svg2pdf(url='tmp_nb.svg', write_to=tmp + '.pdf', dpi=72)

    os.remove('tmp_nb.svg')


def prescription_maker(doctor, lang, order_num):
    """
    Create a pdf prescription notebooks for the corresponding doctor
    :param doctor: Dictionary
    :param lang: str
    :param order_num: str
    :return: -
    """
    doc = dict(list(doctor['g'].items()) + list(doctor['h'].items()) + list(doctor['s'].items()))
    doc['h_last_name'] = doc['h_last_name'].upper()
    doc['h_first_name'] = doc['h_first_name'].upper()

    lang_prescription(doc, f"templates/{lang}_prescriptions.svg", lang, order_num)


if __name__ == '__main__':
    if not os.path.exists('mail_labels/notebooks/'):
        os.makedirs('mail_labels/notebooks/')
    if not os.path.exists('mail_labels/memos/'):
        os.makedirs('mail_labels/memos/')
    if not os.path.exists('notebooks/'):
        os.makedirs('notebooks/')

    print("Initiating the extractor")
    print("---Pdf creation---")
    pdf_maker(sys.argv[1], sys.argv[2])
    print("Extraction achieved")
