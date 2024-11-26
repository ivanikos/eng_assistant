import os
import sqlite3
import xlsxwriter

# Specify the path to your .pspx file
file_path = r"C:\Users\ivana\Documents\Plant 3D files\_PLANT 3D 2025 IMPERIAL TEMPLATE\Spec Sheets"

summary_specs = []

for spec_file in os.listdir(file_path):
    if ".pspc" in spec_file:
        print(spec_file)
        try:
            # Connect to the SQLite database
            connection = sqlite3.connect(f"{file_path}\\{spec_file}")
            cursor = connection.cursor()

            cursor.execute(f"SELECT * FROM 'EngineeringItems';")  # Adjust LIMIT as needed
            rows = cursor.fetchall()

            column_names = [description[0] for description in cursor.description]
            # print(len(column_names))

            if len(column_names) == 42:
                summary_specs.append(
                    [column_names[0], column_names[3], column_names[5], column_names[6], column_names[7],
                     column_names[9], column_names[10], column_names[11], column_names[16], column_names[18],
                     column_names[19], column_names[20], column_names[21], column_names[22], column_names[23],
                     column_names[24], column_names[25], column_names[26], column_names[27], column_names[28],
                     column_names[29], column_names[30], column_names[31], column_names[32], column_names[33],
                     column_names[34], column_names[38], column_names[39], column_names[40],
                     "None", "None", "Spec_Sheet"])

                for row in rows:
                    summary_specs.append(
                        [row[0], row[3], row[5], row[6], row[7],
                         row[9], row[10], row[11], row[16], row[18],
                         row[19], row[20], row[21], row[22], row[23],
                         row[24], row[25], row[26], row[27], row[28],
                         row[29], row[30], row[31], row[32], row[33],
                         row[34], row[38], row[39], row[40],
                         "None", "None", spec_file])

            else:
                summary_specs.append([column_names[0], column_names[3], column_names[5], column_names[6], column_names[7],
                                      column_names[9], column_names[10], column_names[11], column_names[16], column_names[18],
                                      column_names[19], column_names[20], column_names[21], column_names[22], column_names[23],
                                      column_names[24], column_names[25], column_names[26], column_names[27], column_names[28],
                                      column_names[29], column_names[30], column_names[31], column_names[32], column_names[33],
                                      column_names[34], column_names[38], column_names[39], column_names[40],
                                      column_names[42], column_names[43], "Spec_Sheet"])

                for row in rows:
                    summary_specs.append(
                        [row[0], row[3], row[5], row[6], row[7],
                         row[9], row[10], row[11], row[16], row[18],
                         row[19], row[20], row[21], row[22], row[23],
                         row[24], row[25], row[26], row[27], row[28],
                         row[29], row[30], row[31], row[32], row[33],
                         row[34], row[38], row[39], row[40],
                         row[42], row[43], spec_file])

            connection.close()

        except sqlite3.Error as e:
            print(f"SQLite error: {e}")



workbook_summary = xlsxwriter.Workbook(r"C:\Users\ivana\Documents\Plant 3D files\spec_summary.xlsx")
# -------------------------------------Краткая сводка
ws0 = workbook_summary.add_worksheet('spec_summary')
ws0.set_column(0, 2, 25)
ws0.set_row(0, 45)
# ws0.set_column(1, 1, 40)
# ws0.set_column(4, 15, 12)
# ws0.set_column(2, 2, 12)
# ws0.set_column(3, 3, 12)

cell_format_green = workbook_summary.add_format()
cell_format_green.set_bg_color('#6c8784')
cell_format_blue = workbook_summary.add_format()
cell_format_blue.set_bg_color('#b0dacc')
cell_format_hat = workbook_summary.add_format()
cell_format_hat.set_bg_color('#65878f')
cell_format_filename = workbook_summary.add_format()
cell_format_filename.set_bg_color('#f8e8b5')


for i, (a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21,
        a22, a23, a24, a25, a26, a27, a28, a29, a30, a31, a32) in enumerate(summary_specs, start=1):
    color = cell_format_blue
    # if "Working" in one:
    #     color = cell_format_filename
    #     color.set_text_wrap(text_wrap=1)
    #     ws0.merge_range('A1:C1', one, color)
    #     continue
    # if "***" in one:
    #     color = cell_format_green
    #
    # if "Checking SD-TAGs complete" in one or \
    #     "DETAIL CHECKING" in one or "Load tags does" in one or "SD tag" in one:
    #     color.set_text_wrap(text_wrap=1)
    #     color = cell_format_hat
    # try:
    #     color.set_border(style=1)
    #     color.set_text_wrap(text_wrap=1)
    # except:
    #     pass

    ws0.write(f'A{i}', a1, color)
    ws0.write(f'B{i}', a2, color)
    ws0.write(f'C{i}', a3, color)
    ws0.write(f'D{i}', a4, color)
    ws0.write(f'E{i}', a5, color)
    ws0.write(f'F{i}', a6, color)
    ws0.write(f'G{i}', a7, color)
    ws0.write(f'H{i}', a8, color)
    ws0.write(f'I{i}', a9, color)
    ws0.write(f'J{i}', a10, color)
    ws0.write(f'K{i}', a11, color)
    ws0.write(f'L{i}', a12, color)
    ws0.write(f'M{i}', a13, color)
    ws0.write(f'N{i}', a14, color)
    ws0.write(f'O{i}', a15, color)
    ws0.write(f'P{i}', a16, color)
    ws0.write(f'Q{i}', a17, color)
    ws0.write(f'R{i}', a18, color)
    ws0.write(f'S{i}', a19, color)
    ws0.write(f'T{i}', a20, color)
    ws0.write(f'U{i}', a21, color)
    ws0.write(f'V{i}', a22, color)
    ws0.write(f'W{i}', a23, color)
    ws0.write(f'X{i}', a24, color)
    ws0.write(f'Y{i}', a25, color)
    ws0.write(f'Z{i}', a26, color)
    ws0.write(f'AA{i}', a27, color)
    ws0.write(f'AB{i}', a28, color)
    ws0.write(f'AC{i}', a29, color)
    ws0.write(f'AD{i}', a30, color)
    ws0.write(f'AE{i}', a31, color)
    ws0.write(f'AF{i}', a32, color)

workbook_summary.close()


print(len(summary_specs))
# for i in summary_specs:
#     print(i)