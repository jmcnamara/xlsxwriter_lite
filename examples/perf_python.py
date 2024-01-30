##############################################################################
#
# A simple performance test for Python backed XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2024, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('perf_python.xlsx')
worksheet = workbook.add_worksheet()

for row in range(10000):
    for col in range(50):
        worksheet.write_number(row, col, 123.0)

workbook.close()
