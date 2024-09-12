# This script upload all the files for specific local folder into the CZO ShaPoint folder.
# Since the local folder has almost the same structure that the folder in the SharePoint, the script
# will remove the first part of the local folder path and use the rest of the path to create the folder structure.
# In this case will remove the 'E:/Data/' and use the rest of the path to create the folder structure.

from pathlib import Path
from datetime import datetime, timedelta

import office365_api
import Log
import ElapsedTime

folder_path = Path(r'C:/temp/data2/Bahada/CR3000/L0/Flux/')
root_folder = r'C:/temp/data2/'  # r'E:/Data/'

if __name__ == '__main__':
    # Create the log file
    log = Log.Log('upload_folder_script.txt')
    # Create the elapsed time object
    et = ElapsedTime.ElapsedTime()
    # set connection to SharePoint
    sp = office365_api.SharePoint(log=log)
    # Get the current time and the time from two days ago
    specific_time = datetime.now() - timedelta(days=2)
    # Specific time, if needed (uncomment and set if you want a specific cutoff time)
    # specific_time = datetime(2024, 9, 5, 10, 30)  # Replace with your specific date and time
    # Get the list of files in the local folder
    files = [f for f in folder_path.rglob('*') if
             f.is_file() and datetime.fromtimestamp(f.stat().st_mtime) >= specific_time]
    idx = 1
    for item in files:
        log.info(f'File: {item.name}, ({idx}/{len(files)})')
        # print(item)
        upload_file = item.relative_to(root_folder)
        # print(upload_file)
        # Upload the files to the SharePoint folder
        sp.upload_large_file(local_file_path=item, target_file_url=upload_file)
        idx += 1
    # Log the elapsed time
    log.info(f'Elapsed time: {et.elapsed()}')
