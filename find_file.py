import os
import shutil
import re
import time
import datetime

import tkinter as tk
from tkinter import filedialog, Tk, messagebox

"""
Automated Latest Version Extraction Tool for Design Drawing PDFs

This Program is used for Automated Extraction of PDF files with specific namaing formats, capable of: 
1.Searching for PDF files with specific naming pattern in designated directories 
2.Comparing file versions and keeping the latest versions 
3.Selecting the best quality files based on file suffix priority 
4.Copying filtered latest version files to a newly created target folder

Supported file naming formats: 
- Basic format: XX#-#-XXX##-X####-[version].pdf 
- Singed version: XX#-#-XXX##-X####-[version]-Signed.pdf 
- Searchable version: XX#-#-XXX##-X####-[version]-Searchable.pdf 

Author: [Michael; JIUN-AN, TSAI; 蔡濬安]
Version: 1.1
Last Updated: 2025/09/01

"""

'''coding Update History: 

20240516:
    Update the search format in "file_name_formats" to allow for 1 or 2 unmbers or English alphabets, 
    as the previous vertion only allowed searching for 1 unmber or English Alphabet.
    - (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-(\d|[A-Z])\.pdf) 
        -> (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])\.pdf)
    - (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-(\d|[A-Z])-(Signed)\.pdf) 
        -> (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Signed)\.pdf)
    - (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-(\d|[A-Z])-(Searchable)\.pdf) 
        -> (([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Searchable)\.pdf)
'''

class Inf_file:
    """
    File Information Class

    Used to store and manage information related to PDF files, including file name, version, suffix, etc.
    """
    def __init__(self, name, version, suffix, pathfile, filename):
        """
        Initialize file information object

        Args:
            name (str): Base name of the file (without version number and suffix)
            version (str): File version number
            suffix (str): File suffix (such as 'Signed', 'Searchable' or empty string)
            pathfile (str): Complete path of file
            filename (str): Complete filename
        """
        self.name = name
        self.version = version
        self.suffix = suffix
        self.pathfile = pathfile
        self.filename = filename
    def update_inf(self, new_version, new_suffix, new_pathfile, new_filename):
        """
        Update file information

        Args:
            new_version (str): New version number
            new_suffix (str): New file suffix
            new_pathfile (str): New file path
            new_filename (str): New filename
        """
        self.version = new_version
        self.suffix = new_suffix
        self.pathfile = new_pathfile
        self.filename = new_filename



file_name_formats = [
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])\.pdf)',
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Signed)\.pdf)',
    r'(([A-Z]{2}\d{1}-\d{1}-[A-Z]{3}\d{2}-[A-Z]{1}\d{4})-([0-9]{1,2}|[A-Z])-(Searchable)\.pdf)'
]

new_file = []

def find_files(search_path, target_path, file_format, file_name_formats):
    for root, dirs, files in os.walk(search_path):
        for filename in files:
            print('目前檔名', filename)
            # 只選擇檔案是pdf
            if filename.endswith(file_format):
                # 如果檔案與我們要的格式相符，就暫存
                for regex in file_name_formats:
                    match = re.match(regex, filename)
                    if match != None:

                        now_name = match.group(2)
                        now_version = match.group(3)

                        now_groups = match.groups()
                        if len(now_groups) == 4 and now_groups[3]:
                            now_suffix = now_groups[3]
                        else:
                            now_suffix = ''

                        now_fillname = match.group(1)
                        now_file_path = os.path.join(root, filename)

                        print('目前檔案：', match)                    
                        print('目前檔名：', now_name)                        
                        print('目前版本：', now_version)                   
                        print('目前尾數：', now_suffix)                
                        print('目前檔案路徑：', now_file_path)
                        print('目前檔案全名：', now_fillname)

                        # determine whether the new_file list has data
                        if len(new_file) == 0:
                            _ = Inf_file(now_name, now_version, now_suffix, now_file_path, now_fillname)
                            new_file.append(_)
                            
                        else:
                            # determine whether the file name exists in the new_file list
                            for _ in new_file:
                                if _.name == now_name:
                                    exist_name = 'Ture'
                                    exist_obj = _
                                    break
                                else:
                                    exist_name = 'Flase'
                                    exist_obj = _
                            print('目前檔名是否存在於list內：', exist_name)
                            print('目前class名稱：', exist_obj)
                                
                            
                            if exist_name == 'Ture':

                                # obtain relevant information about the current original file name 
                                origin_name = exist_obj.name
                                origin_version = exist_obj.version
                                origin_suffix = exist_obj.suffix
                                origin_filename = exist_obj.filename

                                if origin_version.isdigit():
                                    temorigin_versionformat = int(origin_version)
                                else:
                                    temorigin_versionformat = origin_version

                                if now_version.isdigit():
                                    temnow_versionformat = int(now_version)
                                else:
                                    temnow_versionformat = now_version

                                print('現在檔名：', now_name)
                                print('現在版本', now_version)
                                print('目前版本是否為數字：', now_version.isdigit())

                                print('原始檔名：', exist_obj.name)                                
                                print('原始版本：', exist_obj.version)
                                print('原始版本是否為數字：', origin_version.isdigit())
                                                                                                              
                                                                
                                if isinstance(temorigin_versionformat, int) and isinstance(temnow_versionformat, int) and temorigin_versionformat <= temnow_versionformat : 
                                    print('原始版本比較小')
                                    version_update(temnow_versionformat, now_suffix, now_file_path, now_fillname, temorigin_versionformat, exist_obj)
                                elif isinstance(temorigin_versionformat, str) and isinstance(temnow_versionformat, int):
                                    print('原始版本是字串，現在版本是數字')
                                    version_update(now_version, now_suffix, now_file_path, now_fillname, origin_version, exist_obj)
                                elif isinstance(temorigin_versionformat, str) and isinstance(temnow_versionformat, str) and temorigin_versionformat <= temnow_versionformat:
                                    print('原始版本/現在版本都是字串，但現在本版字串大於等於原版本')
                                    version_update(now_version, now_suffix, now_file_path, now_fillname, origin_version, exist_obj)
                                else:
                                    continue

                            # new_file list dosen't have the same file name, crate.
                            elif exist_name == 'Flase':
                                name = Inf_file(now_name, now_version, now_suffix, now_file_path, now_fillname)
                                print('印出檔名：', now_name)
                                new_file.append(name)
                                    
                            for _ in new_file:
                                print('list目前class名稱：', _)
                                print('list目前檔名', _.name)
                                print('list目前版本', _.version)
                                print('list目前尾碼', _.suffix)
                                print('list目前路徑', _.pathfile)

            else:
                continue

    
    for _ in new_file:
        name_pathfile = _.pathfile
        filename = _.filename
        copy_file(name_pathfile, target_path, filename)


def version_update(version, suffix, file_path, filename, origin_version, classname):
    if version > origin_version:
        print('現在版本大於原版本，更新覆蓋....')           
        classname.version = str(version)
        classname.suffix = suffix
        classname.pathfile = file_path
        classname.filename = filename
    elif not str(origin_version).isdigit() and str(version).isdigit():
        print('原版本是英文，目前版本為數字，更新覆蓋.......')
        classname.version = version
        classname.suffix = suffix
        classname.pathfile = file_path
        classname.filename = filename
    elif version == origin_version:
        print('原版本與現有本版相同，但解析度不同.....')
        if suffix == 'Searchable' and classname.suffix != 'Searchable':
            print('現在解析度為最高，原解析度非最高，更新中............')
            classname.suffix = suffix
            classname.pathfile = file_path
            classname.filename = filename

        elif suffix == '' and classname.suffix == 'Signed':
            print('原版本原Scan檔，現在版本解析度較大，更新中...........')
            classname.suffix = ''
            classname.pathfile = file_path
            classname.filename = filename

def copy_file(name_pathfile, target_path, filename):
    new_source_path = os.path.join(name_pathfile)
    new_target_path = os.path.join(target_path, filename)
    print('檔案來源：', new_source_path)
    print('要移動的路徑及檔名：', new_target_path)
    shutil.copyfile(new_source_path, new_target_path)


def open_dir():
    window_root = tk.Tk()
    window_root.withdraw()
    choose_folder_path = filedialog.askdirectory()

    today_date = datetime.date.today()
    new_folder_path = os.path.join(choose_folder_path, '★★★最新版本({}更新)★★★'.format(today_date))

    folders = os.listdir(choose_folder_path)

    delete_folders = []
    for _ in folders:
        if '最新版本' in _:
            delete_folders.append(_)
        else:
            continue

    if len(delete_folders) == 0:
        messagebox.showinfo('Information', 'There is no ★★★最新版本★★★ of relrvant Folder')
    else:
        for _ in delete_folders:
            folder_to_delete = os.path.join(choose_folder_path, _)
            shutil.rmtree(folder_to_delete)
            messagebox.showinfo('Success', '{} Folder have been deleted.'.format(_))
    os.makedirs(new_folder_path)
    messagebox.showinfo('Success', 'Create new ★★★最新版本({}更新)★★★ Folder.'.format(today_date))        

    return (choose_folder_path, new_folder_path)

def main():
    try:
        start_time = time.time()    
        result_path = open_dir()
        search_path = result_path[0]
        target_path = result_path[1]
        file_format = 'pdf'

        find_files(search_path, target_path, file_format, file_name_formats)
        end_time = time.time()
        print('本次執行程式共花：', round(end_time - start_time, 2), 'sec')
        messagebox.showinfo('Information', '本次執行程式共花：{} sec'.format(round(end_time - start_time, 2)))
    except Exception as _:
        messagebox.showerror('Error', 'An error occurred:{}'.format(_))


if __name__ == '__main__':
    main()