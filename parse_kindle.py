import os
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import zipfile

def parse_kindle_metadata(xml_file):
    """
    XMLファイルを解析して書籍データを抽出
    """
    tree = ET.parse(xml_file)
    root = tree.getroot()

    books = []
    for meta_data in root.findall('.//meta_data'):
        asin = meta_data.findtext('ASIN', default='N/A')
        title = meta_data.findtext('title', default='N/A')
        authors = ', '.join([author.text for author in meta_data.findall('.//authors/author')])
        publisher = ', '.join([publisher.text for publisher in meta_data.findall('.//publishers/publisher')])
        publication_date = meta_data.findtext('publication_date', default='N/A')
        purchase_date = meta_data.findtext('purchase_date', default='N/A')
        content_type = meta_data.findtext('cde_contenttype', default='N/A')
        origin_type = ', '.join([origin.findtext('type') for origin in meta_data.findall('.//origins/origin')])

        books.append({
            "ASIN": asin,
            "Title": title,
            "Authors": authors,
            "Publisher": publisher,
            "Publication Date": publication_date,
            "Purchase Date": purchase_date,
            "Content Type": content_type,
            "Origin Type": origin_type
        })

    return books

def create_diff_from_csvs(xml_file, csv_folder):
    """
    指定フォルダ内の全CSVファイルとXMLデータを比較し、差分データを出力し、ZIPに圧縮
    """
    # XMLからデータを取得
    new_books = parse_kindle_metadata(xml_file)
    new_books_df = pd.DataFrame(new_books)

    # フォルダ内のすべてのCSVファイルを読み込む
    existing_books_df = pd.DataFrame()
    for filename in os.listdir(csv_folder):
        if filename.endswith(".csv"):
            file_path = os.path.join(csv_folder, filename)
            print(f"CSVファイルを読み込んでいます: {file_path}")
            temp_df = pd.read_csv(file_path)
            existing_books_df = pd.concat([existing_books_df, temp_df])

    # 差分を検出（新しいタイトルのみ）
    if not existing_books_df.empty:
        diff_df = new_books_df[~new_books_df['Title'].isin(existing_books_df['Title'])]
    else:
        diff_df = new_books_df

    if not diff_df.empty:
        # カラム順を変更してTitleを左側に配置
        columns = ["Title"] + [col for col in diff_df.columns if col != "Title"]
        diff_df = diff_df[columns]

        # 日付と時間を含む新しいCSVファイル名を作成
        now = datetime.now()
        timestamp = now.strftime('%Y%m%d%H%M%S')
        diff_csv_file = f"kindle_metadata_diff_{timestamp}.csv"

        # 差分データを保存
        diff_df.to_csv(diff_csv_file, index=False, encoding='utf-8-sig')
        print(f"差分データが {diff_csv_file} に保存されました。")

        # ZIPファイル名を作成
        zip_file = f"kindle_metadata_diff_{timestamp}.zip"

        # CSVをZIPに圧縮
        with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(diff_csv_file, os.path.basename(diff_csv_file))
        print(f"CSVがZIPファイルとして保存されました: {zip_file}")

        # 圧縮後に元のCSVファイルを削除（必要に応じて）
        os.remove(diff_csv_file)
    else:
        print("差分データはありません。")

# 実行例
xml_file = "KindleSyncMetadataCache.xml"  # XMLファイル
csv_folder = "./"  # CSVファイルが格納されているフォルダ
create_diff_from_csvs(xml_file, csv_folder)
