##############################################################################
#
# A hello world spreadsheet using the xlsxwriter_lite Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2024, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter_lite as xlsxwriter

workbook = xlsxwriter.Workbook('hello_world.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write_number(0, 0, 123.0)

workbook.push_worksheet(worksheet)

workbook.close()
