# -*- coding: utf-8 -*-

import os
import pickle
# import popen2

"""
Script made for the Cardboard Scents company
This matcher allows to retrieve mail_labels pdf files to be printed matching the inami number of the doctor
Should be adapted if used in another occasion
Structure used:
    - Pickle file containing the mail labels' data
    - Directory mail-labels 
        - Subdirectory memos 
        - Subdirectory notebooks
    - Directory notebooks
"""


def loader():
    """
    Load the pickle file containing the dictionaries useful to the match
    :return: Mail labels dictionaries
    """
    with open('mail_i_hm', 'rb') as f_labels, open('inami_match_hm', 'rb') as f_i_match:
        i_labels = pickle.load(f_labels)
        i_match = pickle.load(f_i_match)
    return i_labels, i_match


def save(labels_ht):
    """
    Save the modified dictionary into the pickle file
    :param labels_ht: Updated mail labels dictionary with less entries
    :return: -
    """
    with open("mail_i_hm", mode='wb+') as f:
        pickle.dump(labels_ht, f)


def print_label(pos_labels, labels, number):
    """
    Print a file, remove it, and remove it from the labels yet to print
    :param pos_labels: Mail labels from the dictionary fitting the requested document (memo or notebook)
    :param labels: Mail labels dictionary
    :param number: Inami number of the doctor
    :return: -
    """
    if len(pos_labels) > 0:
        file = pos_labels[0][3]
        order_number = pos_labels[0][4]
        while True:
            # TODO print
            # popen2.popen4("lpr -P [printer] " + file)
            os.system(file)

            print("---Printing---")
            print("Did it print well? (y/n)")
            print_input = input().lower()
            if print_input == 'y':
                os.remove(file)
                del labels[number][order_number]
                save(labels)

    else:
        print("There is no more file to print for this inami number with this kind of document")


def look_up_name(i_labels, i_match):
    """
    Perform a look in the memo bloc labels to print one with a matching name
    :param i_labels: Dictionary containing the commands numbers as a list by doctor with inami number as key
    :param i_match: Dictionary containing the inami number by doctor with names as key
    :return: -
    """
    while True:
        print("Please enter the last name of the doctor you're looking for (or q to quit):")
        name = input().lower()
        if name == 'q':
            break
        print("Please enter the first name of the doctor you're looking for:")
        name += ' ' + input().lower()

        if name not in i_match.keys():
            print("The entered name does not match any registered doctor memo order")
        else:
            number = i_match[name]
            pos_labels = [e + (k,) for k, e in i_labels[number].items() if e[0] != 'algemene']
            print_label(pos_labels, i_labels, number)
            break


def look_up_inami(labels, memo):
    """
    Perform a look in the labels to print one with a matching inami number for a bloc memo or a notebook
    :param labels: Mail labels dictionary
    :param memo: Indicate if the mail label looked for is for a memo bloc or a notebook
    :return: -
    """
    while True:
        print("Please enter the inami number you wish to search or q to quit this search:")
        cmd_input = input()

        if cmd_input == 'q':
            break

        try:
            number = int(cmd_input)
            if number not in labels.keys():
                print("The entered inami number does not match any registered doctor")
            else:
                if memo:
                    pos_labels = [e + (k,) for k, e in labels[number].items() if e[0] != 'algemene']
                else:
                    pos_labels = [e + (k,) for k, e in labels[number].items() if e[0] == 'algemene']
                print_label(pos_labels, labels, number)
                break

        except ValueError:
            print("Please enter an input respecting the expected format (an inami number or q) \n")


def match(i_labels, i_match):
    """
    Interrogate the user on the inami number to match
    to open the corresponding mail label pdf file
    Unless the user asks to quit, it will propose to redo the operation (while it's possible)
    :param i_labels: Dictionary containing the commands numbers as a list by doctor with inami number as key
    :param i_match: Dictionary containing the commands numbers as a list by doctor with names as key
    :return: -
    """
    while True:
        print("Please enter the requested document (n (notebook) or m (memo)) or q to quit the program:")
        cmd_input = input().lower()

        if cmd_input == 'q':
            break

        elif cmd_input == 'm':
            while True:
                print("Do you wish to search by inami number (i) or bt first name and last name (n) or quit this memo "
                      "search (q)?")
                type_input = input().lower()
                if type_input == 'q':
                    break
                elif type_input == 'n':
                    look_up_name(i_labels, i_match)
                    break
                elif type_input == 'i':
                    look_up_inami(i_labels, True)
                    break
                else:
                    print("The input didn't match any of the expected option, let's start again")

        elif cmd_input == 'n':
            look_up_inami(i_labels, False)

        else:
            print("Please enter an input respecting the expected format (m, n or q) \n")


if __name__ == '__main__':
    mail_labels_by_i, inami_match = loader()
    print("Initiating the matcher")
    match(mail_labels_by_i, inami_match)
    print("Everything ended as planned")
