import pandas as pd
import requests
from tqdm import tqdm
import os
import time

excel_file_location = './urls.xlsx'  # your excel file, change to whatever file name you want.
download_dir = './downloads/'  # your default download location.

grab_time_start = time.time()

# This reads your file and sheet name is "Sheet1"
df = pd.read_excel(excel_file_location, sheet_name='Sheet1')  # Your root data location
total_rows = len(df.index)  # total rows in your root file.


# This converts seconds into hours, minutes
def sec_to_hours(seconds):
    a = str(seconds // 3600)
    b = str((seconds % 3600) // 60)
    c = str((seconds % 3600) % 60)
    d = "{} hours {} mins {} seconds".format(a, b, c)
    return d


# This function downloading the file and showing the download process
def download(name: str, url: str):
    resp = requests.get(url, stream=True)
    total = int(resp.headers.get('content-length', 0))
    # Can also replace 'file' with a io.BytesIO object
    with open(name, 'wb') as file, tqdm(
            desc=name,
            total=total,
            unit='iB',
            unit_scale=True,
            unit_divisor=1024,
    ) as bar:
        for data in resp.iter_content(chunk_size=1024):
            size = file.write(data)
            bar.update(size)


# This looping all urls until finished. Broken URLs will be skipped.
def download_loop(header_name):
    i = 0
    while i < total_rows:
        try:
            url = df.at[i, header_name]  # getting url for every row by header name.
            basename = os.path.basename(url)  # extracts file name from your URL
            save_to = str(download_dir) + str(basename)  # saving location + final name
            download(save_to, url)  # starting the downloads

        except:
            found_error_text = colored('Found some error or None value, skipped row: ', 'red')
            print(found_error_text, i)
            pass  # remove this and use raise if you want to debug
            # raise  # use this for debug otherwise errors will be skipped

        i += 1  # adding 1 to i, so it will continue row by row until finished.
        
        if i == total_rows:
            grab_time_end = time.time()
            total_time = grab_time_end - grab_time_start
            print('Download completed in', sec_to_hours(int(total_time)))
            break


download_loop('urls')  # start downloads
