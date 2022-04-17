from genericpath import exists
import glob
import os
import string
import zipfile
import subprocess
import win32api
import win32con
import win32file
import shutil

from pathlib import Path

def latest_file(path):
    list_of_files = glob.glob(f'{path}/*')
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

def downloads_path():
    return str(Path.home() / "Downloads")

def unzip_file(path, file):
    with zipfile.ZipFile(file, 'r') as zip_ref:
        zip_ref.extractall(path)

def create_mobi_file_path(file):
    file_list = file.split(".")
    file_list[-1] = "mobi"
    path = ".".join(file_list)
    return path

def convert_to_mobi(source, destination):
    program_files = os.environ["ProgramFiles"]
    program_files_86 = os.environ["ProgramFiles(x86)"]
    calibre = f'{program_files}\\Calibre2\\ebook-convert.exe'
    calibre_86 = f'{program_files_86}\\Calibre2\\ebook-convert.exe'
    
    if os.path.exists(calibre):
        execution = calibre
    elif os.path.exists(calibre_86):
        execution = calibre_86
    else:
        raise Exception("Calibre doesn't exist")

    subprocess.call(f'"{execution}" "{source}" "{destination}"')
    return destination

def find_kindle_path():
    drives = [i for i in win32api.GetLogicalDriveStrings().split('\x00') if i]

    for i in drives:
        drivename = win32api.GetVolumeInformation(i)[0]+'('+i+')'
        if drivename != None and "kindle" in drivename.lower():
            return f'{i}documents'
    
    raise Exception("No Kindle drive was found")

def copy_to_kindle(file, folder):
    dst = os.path.join(folder, os.path.basename(file))
    shutil.copyfile(file, dst)

def get_file_extension(file):
    return file.split(".")[-1]

def check_supported_file(extension):
    if extension.lower() not in ['zip', 'fb2', 'epub']:
        raise Exception('Not supported file type')

def main():
    latest_download = latest_file(downloads_path())
    extension = get_file_extension(latest_download)
    check_supported_file(extension)

    
    if extension.lower() == "zip":
        unzip_file(downloads_path(), latest_download)
        latest_download = latest_file(downloads_path())
    
    if len(latest_download) == 0:
        return

    mobi_path = convert_to_mobi(latest_download, create_mobi_file_path(latest_download))

    copy_to_kindle(mobi_path, find_kindle_path())
    
if __name__ == "__main__":
    main()


