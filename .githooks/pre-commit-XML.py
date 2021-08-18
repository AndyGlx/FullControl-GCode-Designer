'''
pre-commit-XML.py

Last updated 23/02/2019

Script extracts customUI.xml/customUI14.xml files into a new 'FILENAME.XML' subdirectory within the repository root folder (by default)

With the standard .gitignore entries, only Excel files located within the root directory of the repository will be processed & added to a commit
'''

import os
import shutil

from zipfile import ZipFile

# list excel extensions that will be processed and have VBA modules extracted
excel_file_extensions = ('xlsb', 'xlsx', 'xlsm', 'xlam', 'xltm')

# process Excel files in this directory (not recursive to subdirectories)
directory = '.'

# String to append to end of subdirectory based on filename
XML_suffix = '.xml'

# String to remove everything after in Excel filename
Rev_tag = ' - Rev '


# function to extract the XML files from an Excel archive
def extract_XML_files(archive_name, full_item_name, extract_folder):

    with ZipFile(archive_name) as zf:
        file_data = zf.read(full_item_name)
    with open(os.path.join(extract_folder, os.path.basename(full_item_name)), "wb") as file_out:
        file_out.write(file_data)


# remove all previous '.XML' directories including contents
for directories in os.listdir(directory):
    if directories.endswith(XML_suffix):
        shutil.rmtree(directories)

# loop through files in given directory and process those that are Excel files (excluding temporary Excel files ~$*)
for filename in os.listdir(directory):
    if filename.endswith(excel_file_extensions):

        # skip temporary excel files ~$*.*
        if not filename.startswith('~$'):

            # Setup variable for wookbook name
            workbook_name = filename

            # Find index of standard '- Rev X.xxx' tag in workbook filename if it exists
            index_rev = filename.find(Rev_tag)

            # If revision tag existed in the filename, remove it based on index position, else remove only the file extension
            if index_rev != -1:
                filename = filename[0:index_rev]
            else:
                filename = os.path.splitext(filename)[0]

                # Setup directory name based on filename
            xml_path = XML_suffix

            # Make new directory (to replace existing where it previously existed)
            os.mkdir(xml_path)

            # print to console for information/debugging purposes
            max_len = max(len('Directory name -- ' + xml_path), len('Workbook name  -- ' + workbook_name))
            print()
            print('-' * max_len)
            print('Workbook name  -- ' + workbook_name)
            print('Directory name -- ' + xml_path)
            print('-' * max_len)

            # Extract customUI.xml & customUI14.xml files from temporary zip file if they exist
            try:
                extract_XML_files(workbook_name, 'customUI/customUI.xml', xml_path)
                print('Extracted -- customUI.xml from ' + workbook_name)
            except KeyError:
                print('customUI.xml does not exists in ' + workbook_name)
            try:
                extract_XML_files(workbook_name, 'customUI/customUI14.xml', xml_path)
                print('Extracted -- customUI14.xml from ' + workbook_name)
            except KeyError:
                print('customUI14.xml does not exists in ' + workbook_name)

# print trailing line to separate output
print()

# loop through directories & remove '.XML' directory if nothing was extracted and it is therefore empty
for directories in os.listdir(directory):
    if directories.endswith(XML_suffix):
        if not os.listdir(directories):
            os.rmdir(directories)
