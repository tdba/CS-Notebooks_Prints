# -*- coding: utf-8 -*-

import os
import re
import sys
import pickle

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
    :return: Mail labels dictionary
    """
    with open('mail_labels_hm', mode='rb') as f_labels:
        labels = pickle.load(f_labels)
    return labels


def save(labels_ht):
    """
    Save the modified dictionary into the pickle file
    :param labels_ht: Updated mail labels dictionary with less entries
    :return: -
    """
    with open("mail_labels_hm", mode='wb+') as f:
        pickle.dump(labels_ht, f)


def match(labels):
    """
    Interrogate the user on the inami number to match
    to open the corresponding mail label pdf file
    Unless the user asks to quit, it will propose to redo the operation (while it's possible)
    :param labels: Dictionary containing the commands number as a list by doctor
    :return: -
    """
    while True:
        print("Please enter the requested inami number or q to quit the program:")
        cmd_input = input()
        if cmd_input.lower() == 'q':
            break
        try:
            number = int(cmd_input)
            if number not in labels.keys():
                print("The entered inami number does not match any registered doctor")
            else:
                # TODO
                pass
        except ValueError:
            print("Please enter an input respecting the expected format (an inami number or q) \n")

    pass


if __name__ == '__main__':
    mail_labels = loader()
    print("Initiating the matcher")
    match(mail_labels)
    print("Everything ended as planned")
