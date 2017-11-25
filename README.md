# OutlookFolderReaderWithPython
Outlook folder reader:  Read specified folder and save items in the folder to csv

You have to make Outlook com object interface.
```
cd path_to_python\Lib\site-packages\win32com\client
python makepy.py
```
Select outlook com object.

## OutlookFolderReaderWithPython.ipynb
Outlookのフォルダ構造のパースと、特定フォルダのメールをCSV出力します。
出力は、送信者,送り先,CC(BCCがわかればここに追加),サブジェクト,送信日,本文,メールヘッダーです。
win32comを利用していますので、Windowsでのみ利用。

環境：windows only, Python3以上

## AppendEmailData.ipynb
上記CSVを読み込み、本文から返信元メッセージを削除し、ヘッダから各種情報を取得します。
返信元メッセージの削除は特定パターン(--original message--以降)のみ削除。
ヘッダ情報は、メッセージID,in-reply-to,references,x-mailer(あれば)を出力します。

環境：Pyhton3以上
