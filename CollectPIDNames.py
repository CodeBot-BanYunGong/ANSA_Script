# -*- coding: utf-8 -*-
"""
Created on Wed May 15 11:21:01 2024

@author: XWEN14
"""

# This script should be run from within ANSA and will search all PIDs within the ANSA model and output a unique list of PIDs in excel format.
# Use only alphanumeric characters and '.' or '_' in the PID names
# If the PID is part of an interface, it must follow the nomenclature of PART_1_NAME.I.PART_2_NAME where PART_1_NAME and PART_2_NAME contain only alphanumeric characters and '.' and '_'

# The user should change the OutputFilePath and Filename as needed.


import os
import ansa
from datetime import datetime


from ansa import base
from ansa import constants
from ansa import utils

print("Select save directory for output Excel file.")
output_filepath = utils.SelectSaveDir(os.getcwd())
# Export to current working directory
filename = "Unique_PIDs.xlsx"  # User defined

pids = base.CollectEntities(constants.NASTRAN, None, "PSHELL")

# Clean all PID names in ANSA file be removing all whitespaces in the names of the PIDs
for pid in pids:
    base.SetEntityCardValues(constants.NASTRAN, pid, {"Name": pid._name.strip()})

pid_list = []

for pid in pids:
    pid_list = pid_list + [pid._name.strip()]

unique_pid_list = []
names_to_check = []

for i in range(len(pid_list)):
    if ".I." in pid_list[i]:
        names_to_check = pid_list[i].split(".I.")
        for name in names_to_check:
            if name not in unique_pid_list:
                unique_pid_list.append(name)
    else:
        if pid_list[i] not in unique_pid_list:
            unique_pid_list.append(pid_list[i])
    names_to_check.clear()

xl_object = utils.XlsxCreate()
# write traceability data into excel sheet
# script version
trace_script_version = "28092023"
# username of script user
trace_user = os.getlogin()
# date/time script was run
now = datetime.now()
trace_runtime = now.strftime("%d%m%Y, %H:%M:%S")
# ansa model name
trace_ansa_db = base.DataBaseName()
utils.XlsxInsertSheet(xl_object)
trace_data_header = ["Script Version", "User", "Run Time", "ANSA DB Name"]
trace_data = [trace_script_version, trace_user, trace_runtime, trace_ansa_db]
xl_row = 0
for i in range(len(trace_data_header)):
    utils.XlsxSetCellValue(
        xl_object,
        "Sheet2",
        xl_row,
        0,
        trace_data_header[i],
    )
    utils.XlsxSetCellValue(
        xl_object,
        "Sheet2",
        xl_row,
        1,
        trace_data[i],
    )
    xl_row += 1

for i in range(len(unique_pid_list)):
    utils.XlsxSetCellValue(xl_object, "Sheet1", i, 0, unique_pid_list[i])

utils.XlsxSave(xl_object, os.path.join(output_filepath, filename))
utils.XlsxClose(xl_object)

print("Done.")
