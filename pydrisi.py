'''
A few simple functions for making some things in Idrisi simpler.

TODO:
- All of these should raise exceptions rather than just exiting.s
'''

import sys
import os
import win32com.client

idrisi = win32com.client.Dispatch('idrisi32.IdrisiAPIServer')

def set_workspace(workspace):
    idrisi.SetWorkingDir(workspace)

def list_rasters(workspace):
    raster_files = []
    files = os.listdir(workspace)
    for file_name in files:
        if file_name[-4:] == ".rst":
            raster_files.append(file_name)
    return raster_files

def reclass(in_raster, out_raster, rcl_file):
    idrisi.RunModule('reclass', r"I*%s*%s*3*%s*1"%(in_raster, out_raster, rcl_file),1,'','','','',1)

def write_rcl(values, output_rcl):
    rcl_file = open(output_rcl, 'w')
    for value_list in values:
        reclass_to_value = value_list[0]
        reclass_start = value_list[1]
        reclass_till = value_list[2]
        rcl_file.write("%s %s %s\n"%(reclass_to_value, reclass_start, reclass_till))
    rcl_file.close()

def extract(extract_values_from_an_image, by_overlaying_this_image, output_type, summary_type, output_name = "temp"):
    output_types = {"VALUES_FILE":"1",
               "TABLE_ON_SCREEN":"2",
               "IMAGE_FILE":"3"}
    summary_types = {"MIN":"1",
                     "MAX":"2",
                     "SUM":"3",
                     "AVERAGE":"4",
                     "MODE":"5",
                     "RANGE":"6",
                     "POP_SD":"7",
                     "SAMPLE_SD":"8",
                     "ALL":"9"}
    
    try:
        out = output_types[output_type.upper()]
    except KeyError:
        print("ERROR : EXTRACT parameter output_type must be a member of:\nVALUES_FILE\nTABLE_ON_SCREEN\nIMAGE_FILE")
        sys.exit()
        
    try:
        summary = summary_types[summary_type.upper()]
    except KeyError:
        print("ERROR : EXTRACT parameter summary_type must be a member of:\nMIN\nMAX\nSUM\nAVERAGE\nMODE\nRANGE\nPOP_SD\nSAMPLE_SD\nALL")
        sys.exit()

    if summary == "9" and out != "3":
        print("ERROR : summary_type ALL is only allowed if output_type is IMAGE_FILE")
        sys.exit()
    else:
        if out == "2":
            idrisi.RunModule('extract',r"%s*%s*%s*%s"%(by_overlaying_this_image, extract_values_from_an_image, out, summary),1,'','','','',1)
        else:
            idrisi.RunModule('extract',r"%s*%s*%s*%s*%s"%(by_overlaying_this_image, extract_values_from_an_image, out, summary, output_name),1,'','','','',1)

def initial_from_copy(image_to_create, image_to_copy, output_data_type, output_file_type, initial_value, units):
    data_types = {"INTEGER":"1",
                  "REAL":"2",
                  "BYTE":"3"}
    file_types = {"BINARY":"1",
                  "ASCII":"2"}
    try:
        o_d_type = data_types[output_data_type.upper()]
    except KeyError:
        print("ERROR : INITIAL parameter output_data_type must be a member of : ")
        for key in data_types:
            print(" - %s"%(key))
        sys.exit()
    try:
        o_f_type = file_types[output_file_type.upper()]
    except KeyError:
        print("ERROR : INITIAL parameter output_file_type must be a member of : ")
        for key in file_types:
            print(" - %s"%(key))
        sys.exit()
    idrisi.RunModule('initial', r"%s*%s*%s*%s*1*%s*%s"%(image_to_create, o_d_type, o_f_type, initial_value, image_to_copy, units), 1,'','','','',1)


def overlay(image_1, image_2, operation, output_image):
    operations = {"ADD":"1",
                  "SUBTRACT":"2",
                  "MULTIPLY":"3",
                  "RATIO_1":"41",
                  "RATIO_2":"42",
                  "RATIO_3":"43",
                  "NORMALIZED_RATIO":"5",
                  "EXPONENT":"6",
                  "COVER":"7",
                  "MIN":"8",
                  "MAX":"9"}
    try:
        operation_id = operations[operation.upper()]
    except KeyError:
        print("ERROR : OVERLAY parameter operation must be a member of :")
        for key in operations:
            print(" - %s"%(key))
    idrisi.RunModule('overlay', r"%s*%s*%s*%s"%(operation_id, image_1, image_2, output_image), 1,'','','','',1)

def crosstab_hard(raster_1, raster_2, mask, output_type, raster_3 = "none", output_image = "none", output_table = "none"):
    o_types = {"IMAGE":"1",
               "TABLE":"2",
               "IMAGE_AND_TABLE":"3",
               "ASSOCIATION_DATA":"4"}
    try:
        o_type = o_types[output_type.upper()]
    except KeyError:
        print("ERROR : CROSSTAB_HARD parameter output_type must be a member of : ")
        for key in o_types:
            print(" - %s"%(key))
        sys.exit()
    if o_type in ["2", "3"]:
        idrisi.RunModule('crosstab', r"1*%s*%s*%s*%s*%s*%s*Y"%(raster_1, raster_2, raster_3, mask, o_type, output_image), 1,'','','','',1)
    else:
        idrisi.RunModule('crosstab', r"1*%s*%s*%s*%s*%s*%s"%(raster_1, raster_2, raster_3, mask, o_type, output_image), 1,'','','','',1)
    if o_type in ["2", "3"] and output_table != "none":
        workspace =idrisi.GetWorkingDir()
        idr_tmp_ctab = open(workspace+"\\ctab1.id$")
        saved_table = open(output_table, 'w')
        for line in idr_tmp_ctab:
            saved_table.write(line)
        idr_tmp_ctab.close()
        saved_table.close()
        os.remove(workspace+"\\ctab1.id$")
    elif o_type in ["2", "3"] and output_table == "none":
        print("ERROR : CROSSTAB_HARD requires output_table value when output_type is set to IMAGE or IMAGE_AND_TABLE.")
        sys.exit()
    elif o_type in ["1", "4"] and output_table != "none":
        print("ERROR : CROSSTAB_HARD output_table value specified, but output_type is not set to IMAGE or IMAGE_AND_TABLE.")
        sys.exit()
    else:
        pass
