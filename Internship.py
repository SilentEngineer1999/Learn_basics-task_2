# ------------------------------------ My approach to this problem ------------------------------------#
""" I created the header for the Output separately then I row values for all student details separately
seeing as how they repeat for different tests from the same student making it easier to iterate. Then I
wrote those values into a new Excel sheet and got the Output."""
# ------------------------------------------------Start----------------------------------------------- #
import pandas

# ----------------------------------------- Input File Path --------------------------------------------#
file_path = input("enter file the File path for to optimize Student Test :-")
# file_path = input_1.xlsx
# file_path = Input_2.xlsx
file = pandas.read_excel("input_1.xlsx")
# ------------------------------------------Side Note---------------------------------------------------#
"""I divided the header into student details and test parameters and made a separate list for test names"""
# --------------------------------------- get values of head -------------------------------------------#
header_user = []
header_parameters = []
test_names = []

for cols in file.columns:
    if "-" in cols:
        test = cols.split("- ")
        if test[1] not in header_parameters:
            header_parameters.append(test[1])
        if test[0].rstrip() not in test_names:
            test_names.append(test[0].rstrip())
    else:
        header_user.append(cols)

header_user.insert(3, "Test_Name")
test_names = sorted(test_names)

# --------------------------------Separate list of Names,id,chapter tag,marks---------------------------#
name = file[header_user[0]].to_list()
username = file[header_user[1]].to_list()
chapter_tag = file[header_user[2]].to_list()


per_stud_test = []
for j in range(0, len(test_names)):
    test_details = []
    for i in range(0, len(header_parameters)):
        if i > len(header_parameters) - 4:
            test_details.append(file[f"{test_names[j]}- {header_parameters[i]}"].to_list())
        else:
            test_details.append(file[f"{test_names[j]} - {header_parameters[i]}"].to_list())

    per_stud_test.append(test_details)

# ------------------------------------- Side Note --------------------------------------------------------#
"""Using the header i created a list of column values in one row so one list within a list had student details
test details together thus pairing it with header {row:column}"""
# ---------------------------------- Combined list of all cell values ---------------------------------- #

row = []
for i in range(0, len(name)):
    for j in range(0, len(test_names)):
        for k in range(0, len(test_names)):
            if per_stud_test[j][0][k] != "-":
                row.append([name[i], username[i], chapter_tag[i], test_names[j], per_stud_test[k][2][j],
                            per_stud_test[k][3][j], per_stud_test[k][0][j], per_stud_test[j][5][k],
                            per_stud_test[k][1][j], per_stud_test[k][4][j]])

header = header_user + sorted(header_parameters)
# ---------------------------------------dont edit from here ------------------------------------------- #
"""This part is kind of set in stone so that writing part is standard no matter what the input be if any edits
needs to be made for any new non recurring student parameters for eg. email id of student then only edits need to
be made to the header and row"""
# ------------------------------------ Writing part -----------------------------------------------------#
# ----------------------------------------- Creating a dataframe --------------------------------------- #
df = pandas.DataFrame(row, columns=header)
column_list = df.columns

# --------------------- Create a Pandas Excel writer using XlsxWriter engine --------------------------- #
output_path = input("Enter the output path ")
writer = pandas.ExcelWriter("output.xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)

# ------------------------------ Get workbook and worksheet objects ------------------------------------ #
workbook = writer.book
worksheet = writer.sheets['Sheet1']

for idx, val in enumerate(column_list):
    worksheet.write(0, idx, val)

writer.save()
