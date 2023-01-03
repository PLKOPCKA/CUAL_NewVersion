import zipfile
import os
from pandas import read_csv, DataFrame
from Parameters import win_user


def open_zippedCSV_as_DataFrame(csv_name: str, zip_name=f'C:/Users/{win_user}/Downloads/AllCSVExports.zip',
                                delete_zip=True) -> DataFrame:
    """
    Open Zip Archive and read CSV file from it, then optionally delete Zip Archive.

    :param zip_name: Absolute path to Zip Archive name to open. With file extension, e.g. '.zip'
    :param csv_name: CSV file name in specified Zip Archive. With file extension: '.csv'
    :param delete_zip: default True: Specified if you want to delete Zip Archive after operation.
    :return:  DataFrame -> CSV data
    """

    archive = zipfile.ZipFile(zip_name, 'r')
    df = read_csv(archive.open(csv_name))
    archive.close()

    if delete_zip:
        os.remove(zip_name)

    return df


def check_if_file_exist(zip_name=f'C:/Users/{win_user}/Downloads/AllCSVExports.zip', delete_zip=False) -> int:
    """
    Check if Zip Archive file exist and optionally delete it.

    :param zip_name: Absolute path to Zip Archive name to open. With file extension, e.g. '.zip'
    :param delete_zip: default True: Specified if you want to delete Zip Archive after operation.
    :return: Integer -> 1 if file exist, 0 if file don not exist, -1 if file existed and was succesfully deleted
    """

    if os.path.exists(zip_name):
        if delete_zip:
            os.remove(zip_name)
            return -1
        else:
            return 1
    else:
        return 0


def move_file_to_new_directory(new_directory: str, zip_name=f'C:/Users/{win_user}/Downloads/AllCSVExports.zip'):
    os.replace(zip_name, new_directory)


if __name__ == '__main__':
    # print(check_if_file_exist(r"C:\Users\plgrzfil\Downloads\OneDrive_2022-10-20.zip", delete_zip=True))
    # print(check_if_file_exist(r"C:\Users\plgrzfil\Downloads\Hotel list in ZB_2021 (1).docx", delete_zip=False))
    # move_file_to_new_directory(r'C:\Users\plgrzfil\Desktop\Nowy folder\Hotel list in ZB_2021 (1).docx',
    #                            r"C:\Users\plgrzfil\Downloads\Hotel list in ZB_2021 (1).docx")
    # print(check_if_file_exist(r"C:\Users\plgrzfil\Downloads\Hotel list in ZB_2021 (1).docx", delete_zip=False))
    if check_if_file_exist(zip_name=r"C:\Users\plgrzfil\Desktop\Nowy folder\abc.zip"):
        move_file_to_new_directory(rf"C:\Users\plgrzfil\Desktop\Nowy folder\aaa\abc_{0.4}.zip",
                                   zip_name=r"C:\Users\plgrzfil\Desktop\Nowy folder\abc.zip")