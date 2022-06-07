---

### Outlook VBA

受信フォルダ内にある添付ファイル付きのメールから添付ファイルを自動ダウンロードする

---

#### ThisOutlookSession  ~ マクロ自動実行処理 ~

* Outlookの予定アイテム件名を指定し、条件を満たすとアラームをトリガーとしマクロを呼びだし実行する

#### 受信メールの添付ファイルを自動ダウンロードする

* MacroName -> AutoSave2.bas

  * 添付ファイルのあるメールフォルダを指定し、フォルダ内にある添付ファイルがあるメールのファイルを自動ダウンロードする

  * 同名ファイルが指定フォルダ内に複数存在しており、尚且つファイル内容を更新している場合は直近に受信したファイルをダウンロードする

  * Downloadファイルの保存先は、ドキュメントフォルダを指定

---
<img src="https://github.com/TA1851/Outlook_file_download/blob/main/img/File_DL.png">
#### Mail_list.bas -> 改訂版Macro(MacroName:File_Auto_Saving)

* ダウンロードしたファイルをOneDriveに移動させる
* ダウンロードしたファイルとメールをLinkさせる（Excelファイルへの書き出し）

---
#### Requirement

* Library list

  * Visual Basic For Applicasions
  * Microsoft Outlook 16.0 Object Library
  * OLE Automation
  * Microsoft Office 16.0 Object Library
  * Microsoft Script Host Object Model
  * Microsoft Scriping runtime
  * Microsoft Excel 16.0 Object Library

---

【参考記事】
 
* Outlook の予定表設定を行い、Alarm をトリガーとしmacro自動化を行う

  [マクロの自動化方法](https://extan.jp/?p=866&cpage=1&unapproved=1125&moderation-hash=02ff48a4830507554d307dde3b90caf0#:~:text=%E3%81%93%E3%81%AE%E3%83%9E%E3%82%AF%E3%83%AD%E3%82%92ThisOutlookSession%E3%81%AB%E8%BF%BD%E5%8A%A0)

* Library の追加（FileSystemObjectを利用するため）

  [FileSystemObjectの追加](https://www.tipsfound.com/vba/18001)
