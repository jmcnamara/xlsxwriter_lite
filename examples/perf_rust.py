##############################################################################
#
# A simple performance test for Rust backed XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2024, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter_lite as xlsxwriter

workbook = xlsxwriter.Workbook('perf_rust.xlsx')
worksheet = workbook.add_worksheet()

for row in range(10000):
    for col in range(50):
        worksheet.write_number(row, col, 123.0)

workbook.push_worksheet(worksheet)

workbook.close()
