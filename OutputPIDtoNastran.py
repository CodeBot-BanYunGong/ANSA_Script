# PYTHON script created by Matt Obrigkeit 2/26/2020
# This script should be run from within ANSA and will ouput one Nastran file for each row in the InputXLFile.
# Use only alphanumeric characters and '.' or '_' in the PID names
# If the PID is part of an interface, it must follow the nomenclature of PART_1_NAME.I.PART_2_NAME where PART_1_NAME and PART_2_NAME contain only alphanumeric characters and '.' and '_'

# The script will create a nastran_files directory in the OutputFilePath diretory and save each Nastran File with the same names as are listed in the InputXLFile.

# The user should change the OutputFilePath and InputXLFile as needed.

# V2 Updates (4/20/2020):
# Added functionality to determine how many volumes can be detected for each part that is to be exported.
# An excel file is written that summarizes how many volumes it found and whether it deleted the volumes that were created in the detections process.

# V4 Updates (4/9/2022):
# Allows user to select directory for output. Program requires that this directory will include the input Excel file (Unique_PIDs.xlsx).  The Excel #summary file will be saved in the user selected directory. The program will create a sub-directory 'nastran_files' to save all of the nastran #files to.

# PID Naming convention (26-Sep-2023):
# 	PID_NAME base + Region Identifier + Boundary Identifier
#   regional identifiers:
# 	_SI: solid internals
# 	_C: conductor
# 	_SE: solid externals
# 	aux current/voltage sources: ${part_name}_AUX_CVTMS_CS. Must end in _CS
# 	Boundary Identifiers:
# 	_VS, _CS, MI\d+\w+
# 	PID name will always end in _VS, _CS, _MI#A or MI#B (if applicable)
# 	_INLET and _OUTLET
#   _Q\d+\w+
# updated GCARLS14 16-July-2024

import os
import sys
import re
import math
import time
import datetime as dt
from datetime import datetime
import getpass

from ansa import base
from ansa import constants
from ansa import utils
from ansa import mesh


def output_pid_to_nastran():
    debug_mode = False

    # User must input an output file path for location to store output file summary xls. This is also the root directory where a folder called 			"nastran_files" will be created to store all the nastran files. This directory must also include the input Excel file ('Unique_PIDs.xlsx') which should contain a unique list of parts that the user wants to export to NASTRAN files. Part names should start in A1.
    print("Select save directory for output NASTRAN files.")
    output_dir = utils.SelectSaveDir(os.getcwd())
    input_xl_file = os.path.join(output_dir, "Unique_PIDs.xlsx")
    print("Input Excel file: " + input_xl_file)
    # Output summary xls file showing how many volumes were identified for each part in the InputXlFile
    output_xl_file = os.path.join(
        output_dir, os.path.basename(input_xl_file).split(".")[0] + "_Summary.xlsx"
    )
    print("Output Excel summary file: " + output_xl_file)
    m = utils.Messenger()
    m.clear()
    m.echo(False)

    # Check to see if InputXLFile exists and is a file
    if os.path.isfile(input_xl_file):
        # Read parts list from Excel worksheet
        unique_pid_list = []

        xl_wb = utils.XlsxOpen(input_xl_file)
        m.print("Read Input File:" + input_xl_file)
    else:
        m.print("Following file could not be found:" + input_xl_file)
        sys.exit()

    output_filepath = os.path.join(output_dir, "nastran_files")
    print("NASTRAN files saved to: " + output_filepath)

    if os.path.isdir(output_filepath) == False:
        os.makedirs(output_filepath)

    # remember start directory to return at end of script
    start_dir = os.getcwd()
    os.chdir(output_filepath)

    # read excel sheet to make a list of PIDs
    i = 0
    while (utils.XlsxGetCellValue(xl_wb, "Sheet1", i, 0)) != None:
        part_name = utils.XlsxGetCellValue(xl_wb, "Sheet1", i, 0)
        print("Splitting " + part_name + "...")
        split_part_name = []
        split_part_name = part_name.split(".")

        # this block of code is needed for xls files that were modified by the user by deleting certain parts. For some reason, a " was being found in these modified files.
        num_true = 0
        for n in range(len(split_part_name)):
            if len(split_part_name[n]) != 0:
                num_true = num_true + 1
        if num_true == len(split_part_name):
            unique_pid_list.append(part_name)
        i = i + 1
    utils.XlsxClose(xl_wb)

    # write header row of excel sheet
    xl_summary = utils.XlsxCreate()  # Excel Summary Object
    xl_headers = [
        "Part Name",
        "Number of Volumes Identified",
        "Number of Volumes Deleted",
        "Free Edges Detected",
        "Number of Matching Entities",
        "Matching Entities",
    ]
    xl_col = 0
    for i in range(len(xl_headers)):
        utils.XlsxSetCellValue(
            xl_summary,
            "Sheet1",
            0,
            xl_col,
            xl_headers[i],
        )
        xl_col += 1

    # PIDs to ignore based on naming convention
    ignore_prefix_list = ["AIR_EXT"]
    ignore_pid_list = ["INLET", "OUTLET", "AIR_EXT"]
    ignore_suffix_list = [
        "_CS",
        "_VS",
        "_MI",
        "_INLET",
        "_OUTLET",
        "_Q",
        "_CR",
        "_AIR_EXT_",
    ]
    # fraction of PID list to go through (for faster testing)
    list_fraction = 1

    # write traceability data into excel sheet
    # script version
    trace_script_version = "07162024"
    # username of script user
    trace_user = getpass.getuser()
    # date/time script was run
    now = datetime.now()
    trace_runtime = now.strftime("%d%m%Y, %H:%M:%S")
    # input/output excel names
    trace_input_file = input_xl_file
    # ansa model name
    trace_ansa_db = base.DataBaseName()
    utils.XlsxInsertSheet(xl_summary)
    trace_data_header = [
        "Script Version",
        "User",
        "Run Time",
        "Input File",
        "ANSA DB Name",
    ]
    trace_data = [
        trace_script_version,
        trace_user,
        trace_runtime,
        trace_input_file,
        trace_ansa_db,
    ]
    xl_row = 0
    for i in range(len(trace_data_header)):
        utils.XlsxSetCellValue(
            xl_summary,
            "Sheet2",
            xl_row,
            0,
            trace_data_header[i],
        )
        utils.XlsxSetCellValue(
            xl_summary,
            "Sheet2",
            xl_row,
            1,
            trace_data[i],
        )
        xl_row += 1

    # Cycle thru all PSHELL entities in the model to find surfaces that belong to each part to export
    all_pids = base.CollectEntities(constants.NASTRAN, None, "PSHELL")
    excel_row = 1  # row number to start printing data in xls summary file

    start = time.time()
    for part_num in range(math.floor(len(unique_pid_list) * list_fraction)):
        # script timing
        cur_iter = part_num + 1
        max_iter = math.floor(len(unique_pid_list) * list_fraction)
        prstime = calcProcessTime(start, cur_iter, max_iter)
        if cur_iter > 1:  # skip the first PID since time estimate will be absurd
            print(
                "time elapsed: %s(s), time left: %s(s), estimated finish time: %s"
                % prstime
            )

        # filter out exported part if PID matches ignore list
        part_to_export = unique_pid_list[part_num]
        print(part_to_export)
        if any(
            re.search("\w+" + suffix + "+(\d*\w*)*$", part_to_export)
            for suffix in ignore_suffix_list
        ):
            print(part_to_export + " contains an ignored suffix.")
            continue
        elif any(
            re.search("^" + prefix + "\w+", part_to_export)
            for prefix in ignore_prefix_list
        ):
            print(part_to_export + " begins with an ignored prefix.")
            continue
        elif part_to_export in ignore_pid_list:
            print(part_to_export + " is an ignored PID.")
            continue

        # find any PIDs related to PartToExport that match PartToExport or PartToExport + an identifier (including ignore list)
        # this list will provide the entities that belong to each part to export
        matching_entities = []
        matching_entity_names = []  # list of names for the matching entities
        print("Finding matches for " + part_to_export + "...")
        for pid in all_pids:
            split_pid = []
            split_pid = pid._name.split(".I.")
            if (part_to_export in pid._name) and debug_mode:
                print(pid._name + " may be associated with " + part_to_export)
            if len(split_pid) > 2:
                m.print(
                    pid._name
                    + "Name is invalid because it has more than one '.I.' separator"
                )
            elif len(split_pid) == 1:
                if split_pid[0] == part_to_export or any(
                    re.search(
                        part_to_export + "(_AUX\d*)*" + suffix + "+(\d*\w*)*$",
                        split_pid[0],
                    )
                    for suffix in ignore_suffix_list
                ):
                    print("Matching " + pid._name + " with " + part_to_export)
                    matching_entities.append(pid)
                    matching_entity_names.append(pid._name)
            elif len(split_pid) == 2:
                # if either element of PIDsplit contains the PartToExport or PartToExport + IgnoreSuffix, add to MatchingEntities
                if split_pid[0] == part_to_export or any(
                    re.search(
                        part_to_export + "(_AUX\d*)*" + suffix + "+(\d*\w*)*$",
                        split_pid[0],
                    )
                    for suffix in ignore_suffix_list
                ):
                    print("Matching " + pid._name + " with " + part_to_export)
                    matching_entities.append(pid)
                    matching_entity_names.append(pid._name)
                if split_pid[1] == part_to_export or any(
                    re.search(
                        part_to_export + "(_AUX\d*)*" + suffix + "+(\d*\w*)*$",
                        split_pid[1],
                    )
                    for suffix in ignore_suffix_list
                ):
                    print("Matching " + pid._name + " with " + part_to_export)
                    matching_entities.append(pid)
                    matching_entity_names.append(pid._name)

        # Determine the id for each PID in the list of entities that belong to each part to export
        id_to_export = []
        for i in range(len(matching_entities)):
            id = matching_entities[i]._id
            id_to_export.append(id)
        # writes name to column 1
        utils.XlsxSetCellValue(xl_summary, "Sheet1", excel_row, 0, part_to_export)

        # export part and related PIDs to nastran file
        print(
            "Exporting "
            + part_to_export
            + "... ("
            + str(part_num + 1)
            + "/"
            + str(len(unique_pid_list))
            + ")"
        )
        if matching_entities == None:
            m.print("There are no matching surfaces")
        else:
            if debug_mode:
                m.print("Running in debug mode for :" + str(part_to_export))
                # Writes 0 for number of volumes identified in column 2
                utils.XlsxSetCellValue(xl_summary, "Sheet1", excel_row, 1, str(0))
                # Writes number of volumes deleted in column 3
                utils.XlsxSetCellValue(xl_summary, "Sheet1", excel_row, 2, str(0))
                utils.XlsxSetCellValue(
                    xl_summary, "Sheet1", excel_row, 3, "N/A"
                )  # Writes number of free edges in column 4
                utils.XlsxSetCellValue(
                    xl_summary, "Sheet1", excel_row, 4, str(len(matching_entity_names))
                )  # Writes number of matching entities in column 5
                utils.XlsxSetCellValue(
                    xl_summary, "Sheet1", excel_row, 5, str(matching_entity_names)
                )  # Writes matching entities in column 6
            else:
                base.All()
                base.Or(matching_entities, constants.NASTRAN, "PSHELL")

                # Orient Surface Normals
                base.Orient

                # Now export parts and/or surfaces to Nastran File (Exporting of unclosed volumes is necessary for coolant inlet/outlets)
                # PartFileName = OutputFilePath + '\\' + PartToExport + '.nas'

                part_filename = os.path.join(output_filepath, part_to_export + ".nas")
                base.OutputNastran(
                    part_filename,
                    mode="visible",
                    write_comments="above_key",
                    format="short",
                    continuation_lines="on",
                    enddata="on",
                    disregard_includes="on",
                    second_as_first="on",
                    beginbulk="on",
                    version="msc nastran",
                )

                # Check to see if the selected entities form a closed volume
                num_volumes = 0  # of volumes identified
                volumes = mesh.VolumesDetect(
                    1, return_volumes=True, include_facets=False, whole_db=False
                )

                # Check number of free edges found and report
                base.F11ShellsOptionsSet("growth ratio", True, "OPEN-FOAM", 1.2)
                # Setting up quality parameter
                obj = base.checks.mesh.SingleBounds()
                check_reports = obj.execute(
                    exec_mode=base.Check.EXEC_ON_VIS, report=base.Check.REPORT_NONE
                )
                error_entities = None
                for check_report in check_reports:
                    for issue in check_report.issues:
                        error_entities = issue.entities
                if error_entities:
                    print("Free edges detected for part " + str(part_to_export))
                    free_edges_detected = True
                else:
                    free_edges_detected = False

                # output summary to excel file, row by row
                if volumes == None:
                    m.print("No closed volumes identified for :" + str(part_to_export))
                    # Writes 0 for number of volumes identified in column 2
                    utils.XlsxSetCellValue(xl_summary, "Sheet1", excel_row, 1, str(0))
                    # Writes number of volumes deleted in column 3
                    utils.XlsxSetCellValue(xl_summary, "Sheet1", excel_row, 2, str(0))
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 3, str(free_edges_detected)
                    )  # Writes number of free edges in column 4
                    utils.XlsxSetCellValue(
                        xl_summary,
                        "Sheet1",
                        excel_row,
                        4,
                        str(len(matching_entity_names)),
                    )  # Writes number of matching entities in column 5
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 5, str(matching_entity_names)
                    )  # Writes matching entities in column 6
                else:
                    num_volumes = len(volumes)
                    m.print(
                        str(num_volumes)
                        + " closed volumes identified for:"
                        + str(part_to_export)
                    )
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 1, str(num_volumes)
                    )  # Writes number of volumes identified in column 2
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 3, str(free_edges_detected)
                    )  # Writes number of free edges in column 4
                    utils.XlsxSetCellValue(
                        xl_summary,
                        "Sheet1",
                        excel_row,
                        4,
                        str(len(matching_entity_names)),
                    )  # Writes number of matching entities in column 5
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 5, str(matching_entity_names)
                    )  # Writes matching entities in column 6

                    # Now Delete Volumes
                    num_volumes_deleted = 0
                    for vol in volumes:
                        # returns 1 if volume is deleted, otherwise returns 0
                        volsDeleted = mesh.VolumesDelete(vol)
                        num_volumes_deleted = num_volumes_deleted + volsDeleted
                        volsDeleted = 0
                    m.print(str(num_volumes_deleted) + " volumes have been deleted")
                    utils.XlsxSetCellValue(
                        xl_summary, "Sheet1", excel_row, 2, str(num_volumes_deleted)
                    )  # Writes number of volumes deleted in column 3

                    # Now delete add PIDs
                    deleteVolumePIDs()
        excel_row = excel_row + 1

    utils.XlsxSave(xl_summary, output_xl_file)
    utils.XlsxClose(xl_summary)
    # print('Nastran files output directory: ' + OutputFilePath)
    # print('Excel summary file location: ' + OutputXlFile)

    # change back to starting directory
    os.chdir(start_dir)
    print("Done.")
    return


def deleteVolumePIDs():
    # When Auto Detect Volumes is run and volumes are found, PIDs are created named "Auto Detected Volume"
    # This function searches for all PIDs with this name and deletes them
    all_volume_pids = base.CollectEntities(constants.NASTRAN, None, "PSOLID")
    for pid in all_volume_pids:
        if pid._name == "Auto Detected Volume":
            base.DeleteEntity(pid)
    return


def calcProcessTime(starttime, cur_iter, max_iter):
    time_elapsed = time.time() - starttime
    time_estimated = (time_elapsed / cur_iter) * (max_iter)
    time_finish = starttime + time_estimated
    time_finish = dt.datetime.fromtimestamp(time_finish).strftime("%H:%M:%S")  # in time
    time_remaining = time_estimated - time_elapsed  # in seconds
    return (int(time_elapsed), int(time_remaining), time_finish)


def main():
    output_pid_to_nastran()


if __name__ == "__main__":
    main()
