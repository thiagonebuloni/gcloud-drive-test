import io
import os
import pandas as pd
import sheets.sheets as sh
from pprint import pprint
from Google import create_service
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload

CLIENT_SECRET_FILE = 'client_secret_566868905417-g3difppfkga03a99hdvofmnguokj4hgn.apps.googleusercontent.com.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# UPLOADING FILES TO DRIVE FOLDER


def upload(folder_id: str, file_names: list[str],
           mime_types=list[str]):
    """Uploads files to a folder on Drive"""

    for file_name, mime_type in zip(file_names, mime_types):
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }

        media = MediaFileUpload(
            './files/{0}'.format(file_name), mimetype=mime_type)

        service.files().create(body=file_metadata, media_body=media, fields='id').execute()


# upload file
# folder_id = '1HlnMQzQfTmvFj-_KK4q_-O4CuM-eejo8'
# file_names = ['file.xlsx', 'file.jpg']
# mime_types = [
#     'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'image/jpeg']

# upload(folder_id, file_names, mime_types)


# CREATING FOLDER
def create_folder(folder_names: list[str],
                  parent_folder_name: list[str]):
    """Creating folders in drive"""

    for folder_name in folder_names:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_name]
        }

        service.files().create(body=file_metadata).execute()


# parent_folder_name = '1Rd5MVPuQUQ97G8JFfH8DvPCh96AWsy8W'
# folder_names = ['Untitled folder']
# create_folder(folder_names, parent_folder_name)


# DOWNLOADING FILES
def download_files(file_ids: list[str],
                   file_names: list[str],
                   download_folder: str):
    """Download files from drive"""
    for file_id, file_name in zip(file_ids, file_names):
        request = service.files().get_media(fileId=file_id)

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fd=fh, request=request)
        done = False

        while not done:
            status, done = downloader.next_chunk()
            print('Download progress {0}'.format(status.progress() * 100))

        fh.seek(0)

        with open(os.path.join(download_folder, file_name), 'wb') as f:
            f.write(fh.read())
            f.close()


# file_ids = ['11pWt-0Q9wbhaH12VB1iUZWHrFQ3cOcX6',
#             '1BMTdi6kknW7CSI6Lf2alNzyHeQ5v8mP7']
# file_names = ['spreadsheet.xlsx', 'image.jpg']
# download_folder = './files'

# download_files(file_ids, file_names, download_folder)


# COPYING FILES INSIDE DRIVE
def copy_files(source_file_id: str,
               folder_ids: list[str],
               file_name: str = None,
               file_description: str = None):
    """Copy fildes in given folders"""

    for folder_id in folder_ids:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id],
            'starred': True,
            'description': file_description
        }
        service.files().copy(fileId=source_file_id,
                             body=file_metadata
                             ).execute()


# source_file_id = '11pWt-0Q9wbhaH12VB1iUZWHrFQ3cOcX6'
# folder_ids = ['1-0LpQG-XBiS3UMJmOiDyM2Ar7ev0Wn66',
#               '1Z5HJOMveCTkz3CGsOyv_LIHya_lZsMHk',
#               '1Vhmec3W2_tFbQJ1l7zDFKanF7xPIFd7h']

# copy_files(source_file_id, folder_ids)


# GETTING INFORMATION ABOUT ACCOUNT
def about_drive_account():
    """Get information about account, usage and capacity"""
    response = service.about().get(fields='*').execute()
    # pprint(response)

    for key, value in response.get('storageQuota').items():
        print('{0}: {1:.2f} MB'.format(key, int(value) / 1024**2))
        print('{0}: {1:.2f} GB'.format(key, int(value) / 1024**3))


# about_drive_account()


# EXPORTING CSV FILES AND CONVERTING TO GOOGLE SHEETS FILE
def export_csv_file(file_path: str, parents: list[str] = None):
    if not os.path.exists(file_path):
        print(f'{file_path} not found.')
        return
    try:
        file_metadata = {
            'name': os.path.basename(file_path).replace('.csv', ''),
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': parents
        }

        media = MediaFileUpload(filename=file_path, mimetype='text/csv')

        response = service.files().create(
            media_body=media,
            body=file_metadata
        ).execute()

        print(response)
        return response
    except Exception as e:
        print(e)
        return


# import csv files as google sheets
# csv_files = os.listdir('./files/csv')

# for csv_file in csv_files:
#     # export_csv_file(os.path.join('./files/csv', csv_file))
#     export_csv_file(os.path.join('./files/csv', csv_file), parents=[
#                     '1Z5HJOMveCTkz3CGsOyv_LIHya_lZsMHk'])


# UPLOADING EXCEL FILES AND CONVERTING TO GOOGLE SHEETS
def convert_excel_file(file_path: str, folder_ids: list[str] = None):
    if not os.path.exists(file_path):
        print(f'{file_path} not found')
        return

    try:
        file_metadata = {
            'name': os.path.splitext(os.path.basename(file_path))[0],
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': folder_ids
        }

        media = MediaFileUpload(
            filename=file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        response = service.files().create(
            media_body=media,
            body=file_metadata
        ).execute()

        print(response)
        return response

    except Exception as e:
        print(e)
        return


# excel_files = os.listdir('./files/excel')

# for excel_file in excel_files:
#     convert_excel_file(os.path.join('./files/excel', excel_file),
#                        ['1Z5HJOMveCTkz3CGsOyv_LIHya_lZsMHk'])


# LISTING FILES IN A DRIVE FOLDER
def list_files(folder_id: str, show: bool = True):
    query = f"parents = '{folder_id}'"

    response = service.files().list(q=query).execute()
    files = response.get('files')
    next_page_token = response.get('nextPageToken')

    while next_page_token:
        response = service.files().list(q=query, pageToken=next_page_token).execute()
        files.extend(response.get('files'))
        next_page_token = response.get('nextPageToken')

    pd.set_option('display.max_columns', 100)
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.min_rows', 500)
    pd.set_option('display.max_colwidth', 150)
    pd.set_option('display.width', 200)
    pd.set_option('expand_frame_repr', True)
    df = pd.DataFrame(files)

    if not show:
        return df

    print(df)


# folder_id = '1Z5HJOMveCTkz3CGsOyv_LIHya_lZsMHk'
# list_files(folder_id, show=True)


# MOVING FILES IN DRIVE FOLDERS
# change to move only selected files
def move_file(source_folder_id: str, target_folder_id: str):
    query = f"parents = '{source_folder_id}'"

    response = service.files().list(q=query).execute()
    files = response.get('files')
    nextPageToken = response.get('nextPageToken')

    while nextPageToken:
        response = service.files().list(q=query, pageToken=nextPageToken).execute()
        files.extend(response.get('files'))
        nextPageToken = response.get('nextPageToken')

    for f in files:
        if f['mimeType'] != 'application/vnd.google-apps.folder':
            service.files().update(
                fileId=f.get('id'),
                addParents=target_folder_id,  # not needed if you want root
                removeParents=source_folder_id
            ).execute()


# source_folder_id = '1Rd5MVPuQUQ97G8JFfH8DvPCh96AWsy8W'
# target_folder_id = '1RyruFYbBlIn3VNu2CEH2X_By_yeX-Nx3'
# move_file(source_folder_id, target_folder_id)
