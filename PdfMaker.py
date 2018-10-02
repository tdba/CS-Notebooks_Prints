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

notebooks = {"algemene", "ALGEMENE", "Algemene", "CSP100", "CSP50"}
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


def extractor(file):
    """
    Extract data from provided file to store it in dictionary with doctors as key
    :param file: Str (filename)
    :return: Dictionary (doctors)
    """
    #try:
    #    with open('docs_hm', 'rb') as f_docs, open('inami_match_hm', 'rb') as f_match:
    #        doctors_hm = pickle.load(f_docs)
    #        mail_labels_name_hm = pickle.load(f_match)
    #except FileNotFoundError:
    doctors_hm = {}
    mail_labels_name_hm = {}
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

        try:
            inami_key = int(h_values['h_inami_number'].replace('.', ''))
        except ValueError:
            inami_key = int(''.join([str(ord(e)) for e in h_values['h_last_name'] + h_values['h_first_name']]))
        name = h_values['h_last_name'].lower() + ' ' + h_values['h_first_name'].lower()

        if inami_key not in doctors_hm:
            doctors_hm[inami_key] = {'g': g_values, 'h': h_values, 'l': l_values, 's': s_values}
        if inami_key in mail_labels_hm:
            mail_labels_hm[inami_key][row_values[16]] = (row_values[1], row_values[2], row_values[16], row_values[0])
        else:
            mail_labels_hm[inami_key] = {row_values[16]: (row_values[1], row_values[2], row_values[16], row_values[0])}
            if row_values[1] not in notebooks:
                mail_labels_name_hm[name] = inami_key

    with open("docs_hm", 'wb+') as f_doc, open("inami_match_hm", 'wb+') as i_match:
        pickle.dump(doctors_hm, f_doc)
        pickle.dump(mail_labels_name_hm, i_match)

    return doctors_hm, mail_labels_hm


def mail_label_maker(doctor, num, command_type, labels):
    """
    Create a pdf mail label for the corresponding command number
    :param doctor: Dictionary
    :param num: Int
    :param command_type: Str
    :param labels: Dictionary
    :return: -
    """
    doc = dict(list(doctor['l'].items()) + list(doctor['g'].items()) + list(doctor['h'].items()))
    if '.' in doctor['h']['h_inami_number']:
        inami_number = int(doctor['h']['h_inami_number'].replace('.', ''))
    else:
        inami_number = int(''.join([str(ord(e)) for e in doc['h_last_name'] + doc['h_first_name']]))
    doc['g_order_number'] = labels[inami_number][num][3]
    doc['g_target_name'] = command_type
    doc['l_bar_code'] = num
    for k in doc:
        if '&' in doc[k]:
            doc[k] = doc[k].replace('&', '&amp;')

    if len(num) == 24:
        generate('code128', num, output='barcode')
        with open("barcode.svg", mode='r') as f:
            svg_barcode = ''.join(f.readlines()[5:])
            top = '<svg x="121px" y="81px" height="10.000mm" width="40.000mm" xmlns="http://www.w3.org/2000/svg">'
            doc['image_bar_code'] = top + svg_barcode
        os.remove('barcode.svg')

        with open("templates/mail_labels.svg", mode='r') as f:
            template = Template(f.read())
            filled_template = template.substitute(doc)
        with open(str(num) + ".svg", mode='w+') as f:
            f.write(filled_template)

        if command_type in notebooks:
            path = 'mail_labels/notebooks/'
        else:
            path = 'mail_labels/memos/'
        cairosvg.svg2pdf(url=str(num) + '.svg', write_to=path + str(num) + '.pdf')
        labels[inami_number][num] += (path + str(num) + '.pdf',)

        os.remove(str(num) + '.svg')
    else:
        print("An error occurred, we can't provide this mail label because the given code doesn't fit the requirements"
              "\nDenied order number:" + num + '\n With code' + code + '\n\n')


def signet_maker(doctor):
    """
    Create from the doctor's data the svg of the signet adapted to blanks and oversized cells
    :param doctor: Dictionary
    :return: Str (svg code of the signet)
    """
    line = '<text transform ="matrix(1 0 0 1 10.3467 {})" class ="st16 {}"> {} </text>\n'
    step = 8.4004
    y = 473.2532
    result = ''

    def text_adapt(field, ordinate, bold=''):
        """
        Adapted line of the signet regarding the size of the input or its emptiness
        :param field: Str (content for the fields)
        :param ordinate: Int
        :param bold: Str
        :return: Str
        """
        field = str(field)
        if '&' in field:
            field = field.replace('&', '&amp;')
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
            y, text = text_adapt(doctor[e], y)
            result += text
        y += step
    return result


def lang_prescription(doctor, file, num, lang):
    """
    Create a pdf prescription notebooks for the corresponding doctor by language
    :param doctor: Dictionary
    :param file: Str (filename)
    :param num: Int
    :param lang: Str
    :return: -
    """
    doctor['image_bar_code'] = render(str(doctor['h_bar_code'])[:-1])
    doctor['signet'] = signet_maker(doctor)

    with open(file, mode='r') as f:
        template = Template(f.read())
        filled_template = template.substitute(doctor)
    with open(str(num) + ".svg", mode='w+') as f:
        f.write(filled_template)

    inami_number = int(doctor['h_inami_number'].replace('.', ''))
    for i in range(num):
        cairosvg.svg2pdf(url=str(num) + '.svg', write_to="notebooks/" + str(lang) + '.' + str(inami_number) + '.'
                                                         + str(i) + '.pdf', dpi=72)

    os.remove(str(num) + '.svg')


def prescription_maker(doctor, num_by_lang):
    """
    Create a pdf prescription notebooks for the corresponding doctor
    :param doctor: Dictionary
    :param num_by_lang: Int
    :return: -
    """
    doc = dict(list(doctor['g'].items()) + list(doctor['h'].items()) + list(doctor['s'].items()))
    doc['h_last_name'] = doc['h_last_name'].upper()
    doc['h_first_name'] = doc['h_first_name'].upper()

    if num_by_lang[0] > 0:
        lang_prescription(doc, "templates/NL_prescriptions.svg", num_by_lang[0], 'NL')
    if num_by_lang[1] > 0:
        lang_prescription(doc, "templates/FR_prescriptions.svg", num_by_lang[1], 'FR')


def pdf_maker(doctors_hm, mails_hm, prod_type):
    """
    Launch the creations of the needed pdf files (mail labels and prescription notebooks)
    :param doctors_hm: Dictionary
    :param mails_hm: Dictionary
    :param prod_type: String ('n' or 'l' or concatenation of both)
    :return: -
    """
    for inami in mails_hm.keys():
        doctor = doctors_hm[inami]
        if 'l' in prod_type:
            for k, e in mails_hm[inami].items():
                mail_label_maker(doctor, k, e[0], mails_hm)
        if 'n' in prod_type:
            q = (sum([1 for command in mails_hm[inami].values() if command[0] in notebooks and command[1] == 'N']),
                 sum([1 for command in mails_hm[inami].values() if command[0] in notebooks and command[1] != 'N']))

            prescription_maker(doctor, q)


if __name__ == '__main__':
    if not os.path.exists('mail_labels/notebooks/'):
        os.makedirs('mail_labels/notebooks/')
    if not os.path.exists('mail_labels/memos/'):
        os.makedirs('mail_labels/memos/')
    if not os.path.exists('notebooks/'):
        os.makedirs('notebooks/')

    print("Initiating the extractor")
    doctors, mail_labels = extractor(sys.argv[1])
    print("Extraction achieved")

    print("---Pdf creation---")
    pdf_maker(doctors, mail_labels, sys.argv[2])
    if 'l' in sys.argv[2]:
        #try:
        #    with open("mail_i_hm", 'rb') as f_mail:
        #        mail_labels_permanent = pickle.load(f_mail)
        #except FileNotFoundError:
        #    mail_labels_permanent = {}
        with open("mail_i_hm", 'wb+') as f_mail:
        #    pickle.dump({**mail_labels_permanent, **mail_labels}, f_mail)
            pickle.dump(mail_labels, f_mail)
    print("Creations achieved")
