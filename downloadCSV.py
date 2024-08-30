import win32com.client as win32
import os
import time
import shutil

def click_hyperlink(file_name, sheet_name, cell_address):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    full_path = os.path.abspath(file_name)
    
    try:
        workbook = excel.Workbooks.Open(full_path)
        sheet = workbook.Sheets(sheet_name)
        cell = sheet.Range(cell_address)

        if cell.Hyperlinks.Count > 0:
            hyperlink = cell.Hyperlinks(1)
            hyperlink.Follow()
            print("CSVのダウンロードを開始します。( ´ー｀)y-~~")
        else:
            print(f"No hyperlink found in cell {cell_address}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Workbook is closed only after download is complete
        workbook.Close(SaveChanges=False)
        excel.Application.Quit()

def wait_for_download(download_folder, file_extension='.csv', timeout=600):
    start_time = time.time()
    while True:
        # List all files in the download folder with the specified extension
        files = [f for f in os.listdir(download_folder) if f.endswith(file_extension)]
        if files:
            return max(files, key=lambda f: os.path.getmtime(os.path.join(download_folder, f)))
        
        if time.time() - start_time > timeout:
            raise TimeoutError("Timeout while waiting for download to complete.")
        
        print("CSVのダウンロード完了を待っています。(´･ω･)っ旦")
        time.sleep(5)  # Wait before checking again

def move_and_rename_latest_csv():
    download_folder = os.path.expanduser('~/Downloads')  # ダウンロードフォルダのパス
    destination_folder = os.path.dirname(os.path.abspath(__file__))  # スクリプトのあるフォルダ
    new_csv_name = 'issues.csv'  # 新しいCSVファイル名
    
    # ダウンロード完了まで待機し、最新のCSVファイルを取得
    latest_file = wait_for_download(download_folder)
    downloaded_file_path = os.path.join(download_folder, latest_file)
    
    # ファイルをスクリプトと同じ階層に移動してリネーム
    destination_file_path = os.path.join(destination_folder, new_csv_name)
    
    # スクリプトと同じフォルダに既にissues.csvが存在する場合は削除
    if os.path.exists(destination_file_path):
        os.remove(destination_file_path)
        print(f"Existing file deleted: {destination_file_path}")
    
    shutil.move(downloaded_file_path, destination_file_path)
    print(f"File moved and renamed to: {destination_file_path}")

# 使用例
click_hyperlink('CSVダウンロード.xlsx', 'CSVダウンロード', 'A1')
time.sleep(10)  # 追加で待機することでダウンロードが確実に完了するようにする（調整可能）
move_and_rename_latest_csv()
