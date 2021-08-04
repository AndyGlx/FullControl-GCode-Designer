'''
pre-commit-VBA.py

Last updated 30/07/2020

Script extracts VBA modules, forms and class modules into a new 'FILENAME.VBA' subdirectory within the repository root folder (by default)

With the standard .gitignore entries, only Excel files located within the root directory of the repository will be processed & added to a commit
'''

import os
import shutil
import win32com.client
import time
import traceback


def init_Xl(file_path):
    '''function to open Excel and open Workbook at file_path

    :param file_path: full path of Workbook to open/process
    :type file_path: str
    :return: Xl_app ('Microsoft Excel'), Excel Workbook (<COMObject Open>)
    :rtype: str
    '''
    # open Excel as hidden instance in the background
    Xl_app = win32com.client.DispatchEx('Excel.Application')
    # required to disable all macros in all files opened programmatically without showing any security alerts. Equivalent of Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Xl_app.AutomationSecurity = 3
    # required to suppress the overwriting of existing file
    Xl_app.DisplayAlerts = False
    # required to suppress any application events from running on opening excel files
    Xl_app.EnableEvents = False
    # open Workbook
    Xl_wb = Xl_app.Workbooks.Open(Filename=file_path, Editable=True)
    print('Starting Excel instance...')
    return Xl_app, Xl_wb


def close_Xl(Xl_app, Xl_wb):  # , file_path, Xl_file_format):
    '''function to save & close Workbook instance

    :param Xl_app: 'Microsoft Excel'
    :type Xl_app: str
    :param Xl_wb: Microsoft Excel Workbook COMObject
    :type Xl_wb: str
    '''
    print('Closing Excel instance...')
    Xl_app.Workbooks(Xl_wb.Name).Close(SaveChanges=False)


def module_file_ext(Xl_module_type):
    '''function to return the standard extensions based on module type numeric reference

    :param Xl_module_type: Module type codes
        * Type = 1   = Standard Module    -- (*.bas)
        * Type = 2   = Class Module       -- (*.cls)
        * Type = 3   = Userform Module    -- (*.frm)
        * Type = 100 = Sheet Class Module -- (*.cls)
    :type Xl_module_type: int
    :return: standard file extension based on module type
    :rtype: str
    '''
    module_type_dict = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}
    return module_type_dict[Xl_module_type]


# function to extract the VBA modules from an Excel archive
def extract_VBA_files(workbook_name):

    # Find index of standard '- Rev X.xxx' tag in workbook filename if it exists
    index_rev = workbook_name.find(Rev_tag)

    # If revision tag existed in the filename, remove it based on index position, else remove only the file extension
    if index_rev != -1:
        filename = workbook_name[0:index_rev]
    else:
        filename = os.path.splitext(workbook_name)[0]

    # Setup directory name based on filename
    vba_path = directory + VBA_suffix

    # Make new directory (to replace existing where it previously existed and was removed)
    os.mkdir(vba_path)

    # print to console for information/debugging purposes
    max_len = max(len('Directory name -- ' + vba_path), len('Workbook name  -- ' + workbook_name))
    print()
    print('-' * max_len)
    print('Workbook name  -- ' + workbook_name)
    print('Directory name -- ' + vba_path)
    print('-' * max_len)

    # Open Excel instance and open workbook
    Xl_app, Xl_wb = init_Xl(directory + workbook_name)

    # Iterate through each VBComponent (Sheet Class Modules, Class Modules, Standard Modules, Userforms)
    try:
        for Xl_module in Xl_wb.VBProject.VBComponents:
            # get name of module
            Xl_module_name = Xl_module.Name
            # get module type
            Xl_module_type = Xl_module.Type

            module_name = f'{Xl_module_name}{module_file_ext(Xl_module_type)}'  # filename e.g. 'module.bas'
            module_path = f'{vba_path}\\{module_name}'

            # for standard modules, class modules and userforms
            if Xl_module_type in [1, 2, 3]:
                # export module
                Xl_wb.VBProject.VBComponents(Xl_module_name).Export(module_path)
                print(f'Exporting -- {module_name}')

            # for sheet class modules
            elif Xl_module_type in [100]:
                # export module
                Xl_wb.VBProject.VBComponents(Xl_module_name).Export(module_path)
                print(f'Exporting -- {module_name}')

    except Exception as ex:
        print(ex)

    finally:
        # close Excel instance
        close_Xl(Xl_app, Xl_wb)


try:
    # list excel extensions that will be processed and have VBA modules extracted
    excel_file_extensions = ('xlsb', 'xls', 'xlsx', 'xlsm', 'xla', 'xlt', 'xlam', 'xltm')

    # process Excel files in root directory (not recursive to subdirectories)
    directory = os.getcwd() + '//'

    # String to append to end of subdirectory
    VBA_suffix = '.vba'

    # String to remove from end of filename
    Rev_tag = ' - Rev '

    # remove all previous '.VBA' directories including contents
    for directories in os.listdir(directory):
        if directories.endswith(VBA_suffix):
            for x in range(0, 5):  # try 5 times
                try:
                    shutil.rmtree(directories)
                    err = False
                except Exception:
                    err = True
                    with open("precommit_exceptions_log.log", "a") as logfile:
                        traceback.print_exc(file=logfile)
                    pass

                if err:
                    time.sleep(2)  # wait for 2 seconds before trying to remove directory again
                else:
                    break

    # loop through files in given directory and process those that are Excel files (excluding temporary Excel files ~$*)
    for filenames in os.listdir(directory):
        if filenames.endswith(excel_file_extensions):
            # skip temporary excel files ~$*.*, otherwise process Excel file and extract VBA modules
            if not filenames.startswith('~$'):
                # extract VBA modules
                extract_VBA_files(filenames)

    # print trailing line to separate output from multiple files
    print()

    # loop through directories & remove '.VBA' directory if nothing was extracted and directory is empty
    for directories in os.listdir(directory):
        if directories.endswith(VBA_suffix):
            if not os.listdir(directories):
                for x in range(0, 5):  # try 5 times
                    try:
                        os.rmdir(directories)
                        err = False
                    except Exception:
                        err = True
                        with open("precommit_exceptions_log.log", "a") as logfile:
                            traceback.print_exc(file=logfile)
                        pass

                    if err:
                        time.sleep(2)  # wait for 2 seconds before trying to remove directory again
                    else:
                        break

except Exception:
    with open("precommit_exceptions_log.log", "a") as logfile:
        traceback.print_exc(file=logfile)
    raise
    