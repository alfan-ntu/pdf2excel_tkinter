# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
class_set.py
============

This module provides functions to maintain a statemachine for the script customs_pdf_2_excel.py
to operate

.. note::
    This documentation follows the reStructuredText (reST) format, which is
    compatible with Sphinx autodoc.

:author: Maoyi Fan
:email: maoyi.fan@yapro.com.tw
:date: 2025-03-09
:version: 1.1
:license: MIT
:history:
    - 1.1 (2025-03-09) - Added a state variable tax_bill_info_OK to handle the exceptional
                         case where only one record is on the last page
"""
import math


# Function to check
# Log base 2
def Log2(x):
    if x == 0:
        return False
    return (math.log10(x) / math.log10(2))


# Function to check
# if x is power of 2
def isPowerOfTwo(n):
    if n == 0:
        return False
    else:
        return (math.ceil(Log2(n)) == math.floor(Log2(n)))


class entry_setting:
    """
    Class to maintain a statemachine for text parser extraction

    :param tax_bill_entry: State indicator marking 稅單編號 is being processed
    :type tax_bill_entry: bool
    :param decl_form_entry: State indicator marking 報單號碼 is being processed
    :type decl_form_entry: bool
    :param tax_ID_entry: State indicator marking 納稅義務人統編 is being processed
    :type tax_ID_entry: bool
    :param tax_amount_entry: State indicator marking 金額 is being processed
    :type tax_amount_entry: bool
    :param tax_bill_info_OK: A flag indicating if at least one record extract on one page
    :type tax_bill_info_OK: bool

    :return: get_current_setting() method returns flags of current state of the statemachine
    :rtype: flag list

    """
    def __init__(self, tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry, tax_bill_info_OK):
        self.tax_bill_entry = tax_bill_entry          # bit 0 of entry_combination
        self.decl_form_entry = decl_form_entry        # bit 1 of entry_combination
        self.tax_ID_entry = tax_ID_entry              # bit 2 of entry_combination
        self.tax_amount_entry = tax_amount_entry      # bit 3 of entry_combination
        self.tax_bill_info_OK = tax_bill_info_OK      # if any 稅單號碼+報單號碼 extracted
        self.entry_combination = 0

    def set_current_entry(self, tax_bill_entry, decl_form_entry, tax_ID_entry, tax_amount_entry, tax_bill_info_OK):
        self.tax_bill_entry = tax_bill_entry          # bit 0 of entry_combination
        self.decl_form_entry = decl_form_entry        # bit 1 of entry_combination
        self.tax_ID_entry = tax_ID_entry              # bit 2 of entry_combination
        self.tax_amount_entry = tax_amount_entry      # bit 3 of entry_combination
        self.tax_bill_info_OK = tax_bill_info_OK      # if any 稅單號碼+報單號碼 extracted; False right after '報單號碼' tag
                                                      # is raised
        self.entry_combination = 0

        if self.tax_bill_entry is True:
            self.entry_combination = self.entry_combination | 1
#            print("tax_bill_entry true & ec = ", self.entry_combination)
        else:
            self.entry_combination = self.entry_combination & 14    # 0b1110
#            print("tax_bill_entry false & ec = ", self.entry_combination)

        if self.decl_form_entry is True:
            self.entry_combination = self.entry_combination | 2
#            print("decl_form_entry true & ec = ", self.entry_combination)
        else:
            self.entry_combination = self.entry_combination & 13    # 0b1101
#            print("decl_form_entry false & ec = ", self.entry_combination)

        if self.tax_ID_entry is True:
            self.entry_combination = self.entry_combination | 4
#            print("tax_ID_entry true & ec = ", self.entry_combination)
        else:
            self.entry_combination &= 11    # 0b1011
#            print("tax_ID_entry false & ec = ", self.entry_combination)

        if self.tax_amount_entry is True:
            self.entry_combination = self.entry_combination | 8
#            print("tax_amount_entry true & ec = ", self.entry_combination)
        else:
            self.entry_combination &= 7     # 0b0111
#            print("tax_amount_entery false & ec = ", self.entry_combination)

        if isPowerOfTwo(self.entry_combination) is True:
            pass
#            print("<<<class_set.py>>>[true]self.entry_combination=",
#                    self.entry_combination)
            return True
        else:
            pass
#            print("<<<class_set.py>>>[false]self.entry_combination=",
#                    self.entry_combination)
            return False

    def get_current_setting(self):
        """
        Retrieve the statemachine flags for text parser to categorize the current text
        :return: statemachine flags
        """
        if isPowerOfTwo(self.entry_combination) is True:
            return True, self.tax_bill_entry, self.decl_form_entry, \
                self.tax_ID_entry, self.tax_amount_entry, self.tax_bill_info_OK
        else:
            return False, self.tax_bill_entry, self.decl_form_entry, \
                self.tax_ID_entry, self.tax_amount_entry, self.tax_bill_info_OK

    def clear_current_setting(self):
        """
        Clear the statemachine flags
        :return: statemachine flags after reset
        """
        self.tax_bill_entry = False
        self.decl_form_entry = False
        self.tax_ID_entry = False
        self.tax_amount_entry = False
        # self.tax_bill_info_OK = False
        self.entry_combination = 0
        return self.tax_bill_entry, self.decl_form_entry, self.tax_ID_entry, \
            self.tax_amount_entry
