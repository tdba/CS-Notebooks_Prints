# -*- coding: utf-8 -*-

import os
import sys
import pickle

"""
Script made for the Cardboard Scents company
This matcher allows to retrieve mail_labels pdf files to be printed matching the inami number of the doctor
Should be adapted if used in another occasion
Structure used:
    - Pickle file containing the doctors' data
    - Directory mail-labels 
        - Subdirectory memos 
        - Subdirectory notebooks
    - Directory notebooks
"""


def loader():
    """
    Load the pickle file containing the dictionaries useful to the match
    :return: Doctor and mail labels dictionaries
    """
    with open('doctors_hm', mode='r') as f_doc, open('mail_labels_hm', mode='r') as f_labels:
        docs = pickle.load(f_doc)
        labels = pickle.load(f_labels)
    return docs, labels


def save(labels_ht):
    """
    Save the modified dictionary into the pickle file
    :param labels_ht: Updated mail labels dictionary with less entries
    :return: -
    """
    with open("mail_labels_hm", mode='wb+') as f:
        pickle.dump(labels_ht, f)


def match(docs, labels):
    """
    Interrogate the user on the inami number to match
    to open the corresponding mail label pdf file
    Unless the user asks to quit, it will propose to redo the operation (while it's possible)
    :param docs: Dictionary containing doctors'data under the form of dictionaries by doctor
    :param labels: Dictionary containing the commands number as a list by doctor
    :return:
    """
    pass


if __name__ == '__main__':
    doctors, mail_labels = loader()
    print("Initiating the matcher")
    match(doctors, mail_labels)
    print("Everything ended as planned")
