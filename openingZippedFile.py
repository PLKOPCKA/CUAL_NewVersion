import zipfile
import os
import pandas as pd
from Parameters import win_user


def open_zippedCSV_as_DataFrame(csv_name: str, zip_name=f'C:/Users/{win_user}/Downloads/AllCSVExports.zip', delete_zip= True):
    '''
    Open Zip Archive and read CSV file from it, then optionally delete Zip Archive.

    :param zip_name: Absolute path to Zip Archive name to open. With file extension, e.g. '.zip'
    :param csv_name: CSV file name in specified Zip Archive. With file extension: '.csv'
    :param delete_zip: default True: Specified if you want to delete Zip Archive after operation.
    :return:  DataFrame -> CSV data
    '''

    archive = zipfile.ZipFile(zip_name, 'r')
    df = pd.read_csv(archive.open(csv_name))
    archive.close()

    if delete_zip:
        os.remove(zip_name)

    return df