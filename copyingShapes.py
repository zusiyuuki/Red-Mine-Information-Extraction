import win32com.client
import os
import re

def fullwidth_to_halfwidth(text):
    # 全角文字を半角に変換する変換辞書
    fullwidth_chars = '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
    halfwidth_chars = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"#$%&\'()*+,-./:;<=>?@[\]^_`abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
    trans = str.maketrans(fullwidth_chars, halfwidth_chars)
    return text.translate(trans)

def copy_shapes(template_sheet, target_sheet):
    for shape in template_sheet.Shapes:
        try:
            # 図形を新しく作成
            new_shape = target_sheet.Shapes.AddShape(
                shape.AutoShapeType, shape.Left, shape.Top, shape.Width, shape.Height
            )
            # 図形の書式をコピー
            new_shape.Fill.ForeColor.RGB = shape.Fill.ForeColor.RGB
            new_shape.Line.ForeColor.RGB = shape.Line.ForeColor.RGB
            
            # 図形内の文字をコピー
            text_range = new_shape.TextFrame2.TextRange
            text_range.Text = shape.TextFrame2.TextRange.Text
            
            # テキストを半角に変換
            text_range.Text = fullwidth_to_halfwidth(text_range.Text)
            
            # 文字のフォント、サイズ、色を上書き
            text_range.Font.Name = "Meiryo UI"  # フォントをMeiryo UIに設定
            text_range.Font.Size = 14  # フォントサイズを14に設定
            text_range.Font.Fill.ForeColor.RGB = 0  # 黒色に設定
            
            # テキストの位置を中央揃え（水平および垂直）
            new_shape.TextFrame2.HorizontalAnchor = 2  # 水平中央揃え
            new_shape.TextFrame2.VerticalAnchor = 3    # 垂直中央揃え
            text_range.ParagraphFormat.Alignment = 2   # テキストの中央揃え

        except Exception as e:
            print(f"図形 {shape.Name} のコピー中にエラーが発生しました: {e}")

def main():
    template_path = 'テンプレート.xlsx'
    output_folder = os.path.join(os.path.dirname(__file__), 'output')

    if not os.path.exists(template_path):
        print(f"テンプレートファイルが見つかりません: {template_path}")
        return
    
    excel_files = [f for f in os.listdir(output_folder) if f.endswith('.xlsx')]
    if not excel_files:
        print("出力フォルダにExcelファイルが見つかりませんでした。")
        return
    
    xl_app = win32com.client.Dispatch('Excel.Application')
    xl_app.Visible = False
    
    wb_template = xl_app.Workbooks.Open(os.path.abspath(template_path))
    template_sheet = wb_template.Sheets('【別紙15】障害管理簿')
    
    for excel_file in excel_files:
        new_file_path = os.path.join(output_folder, excel_file)
        
        wb_target = xl_app.Workbooks.Open(os.path.abspath(new_file_path))
        target_sheet = wb_target.Sheets('【別紙15】障害管理簿')
        
        copy_shapes(template_sheet, target_sheet)
        
        wb_target.Save()
        wb_target.Close()
    
    wb_template.Close()
    xl_app.Quit()
    
    print("図形の転記が完了しました。")

if __name__ == "__main__":
    main()
