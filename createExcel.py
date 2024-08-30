import csv
import datetime
import os
import shutil
import pandas as pd
import chardet
import xlwings as xw

today = datetime.date.today()


# 起票チームごとに分割したCSVを格納する辞書
csv_dict = {}

# ヘッダーの退避
headers = []

# CSVファイルの読み込みと仕分け
with open('issues.csv') as f:
    reader = csv.reader(f)
    for row in reader:
        # ヘッダー行を退避
        if row[32] == '起票チーム':
            headers = row
            continue

        # 起票チームが空欄ならNone
        team = '入力なし'
        if row[32] != '':
            # 起票チームがあれば取得
            team = row[32]

        if team in csv_dict:
            # dictに存在するチームなら行を追加
            csv_dict[team].append(row)
        else:
            # 初めてのチームならdictに空配列を追加してから行を入れる
            csv_dict[team] = []
            csv_dict[team].append(row)

# dictをループして出力
os.makedirs('output', exist_ok=True)
team_files = []

for team, items in csv_dict.items():
    team_file = f'output/{team}.csvを出力しました'
    team_files.append(team_file)  # チームごとのファイル名を記録
    # チーム名のCSVファイルを作成。Windows環境ではnewline=''を入れないと改行コード周りで出力結果がおかしくなる
    with open(team_file, 'w', newline='') as f:
        writer = csv.writer(f)
        # すべてのCSVで1行目はヘッダー
        writer.writerow(headers)
        # チーム単位に仕分けした行を順次出力
        for row in items:
            writer.writerow(row)

    # チームごとにさらにプロジェクトごとに分割
    project_dict = {}
    with open(team_file, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            if row == headers:
                continue  # ヘッダー行はスキップ
            project = row[1]  # 2列目のプロジェクト名を取得
            if project in project_dict:
                project_dict[project].append(row)
            else:
                project_dict[project] = []
                project_dict[project].append(row)

    # プロジェクトごとに出力
    for project, project_items in project_dict.items():
        project_file = f'output/{team}_{project}.csv'
        with open(project_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(project_items)

# デバッグ情報: team_filesの内容を確認

print("CSVをExcelに転記します。")
for team_file in team_files:
    print(team_file)
# チームごとのCSVファイルを削除
for team_file in team_files:
    if os.path.exists(team_file):
        os.remove(team_file)
    else:
        print(f"File not found: {team_file}")
# ここからExcelへの書き込み処理
def copy_csv_to_excel(csv_file_path, sheet):
    # CSVファイルのエンコーディングを検出
    encoding = detect_encoding(csv_file_path)
    
    # CSVファイルの読み込み
    df = pd.read_csv(csv_file_path, encoding='cp932')
    
    # DataFrameの内容をExcelシートに書き込む
    # ヘッダーを含めて転記
    for i, col in enumerate(df.columns):
        sheet.range(6, i + 1).value = col  # ヘッダーを書き込み
    for i, row in enumerate(df.values):
        sheet.range(7 + i, 1).value = row  # データを書き込み
    
    # シート全体の折り返して表示を解除
    used_range = sheet.used_range
    used_range.api.WrapText = False

def detect_encoding(file_path):
    # ファイルのエンコーディングを自動的に検出する
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def main():
    # フォルダを指定（outputフォルダに固定）
    folder_path = os.path.join(os.path.dirname(__file__), 'output')
    
    # テンプレートファイルのパス
    template_path = 'テンプレート.xlsx'
    if not os.path.exists(template_path):
        print(f"テンプレートファイルが見つかりません: {template_path}")
        return
    
    # フォルダ内のCSVファイルを取得
    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    if not csv_files:
        print("CSVファイルが見つかりませんでした。")
        return
    
    # CSVファイルごとにテンプレートファイルをコピー
    for csv_file in csv_files:
        csv_name = os.path.splitext(csv_file)[0]  # 拡張子を除いたファイル名
        new_file_name = f"全体_EY障害管理簿({csv_name}).xlsx"
        new_file_path = os.path.join(folder_path, new_file_name)
        
        # テンプレートファイルをコピー
        shutil.copy(template_path, new_file_path)
        
        # コピーしたExcelファイルを開く
        app = xw.App(visible=False)
        wb = app.books.open(new_file_path)
        sheet = wb.sheets['【別紙15】障害管理簿']
        
        # D4セルにCSVファイルの名前を設定
        # sheet.range('D4').value = csv_name
        
        # CSVの内容を転記
        csv_file_path = os.path.join(folder_path, csv_file)
        copy_csv_to_excel(csv_file_path, wb.sheets['Redmine出力ファイル'])
        
        # Excelファイルを保存して閉じる
        wb.save()
        wb.close()
        app.quit()
        
        # print(f"コピー完了: {new_file_path} - D4に設定: {csv_name}")

if __name__ == "__main__":
    main()
