Attribute VB_Name = "Mail_list"
Option Explicit
Sub File_Auto_Saving()
    
    Call Filedel1  'ファイルの保存状態を常に最新に保つ
    Call Filedel2
    
    Dim objInbox As Object
    Dim objFolder As Object
    Dim strPath As String
    Dim i As Long
    Dim objItem As Object
    Dim OneDrive_Path As String
    
    'Excel用定義
    Dim myExcel As Excel.Application
    Dim objBook As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim n As Long
    
    'Excelオブジェクト生成、ブックの追加
    Set myExcel = CreateObject("Excel.Application")
    Set objBook = myExcel.Workbooks.Add()
    Set objSheet = objBook.Sheets(1)

    '項目目を追加
    objSheet.Cells(1, 1) = "subject"
    objSheet.Cells(1, 2) = "To"
    objSheet.Cells(1, 3) = "Date"
    objSheet.Cells(1, 4) = "Files"
    objSheet.Cells(1, 5) = "File_Path"
    
    '添付ファイルリストを書き込む行の位置
    n = 2
    
    Set objInbox = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    '添付ファイルがあるメールのフォルダを指定します。2階層以上ある場合は「.Folders.Item(＜フォルダ名＞)」を追加してください。
    Set objFolder = objInbox.Folders.Item("title").Folders.Item("name").Folders.Item("item name")
    
    '添付ファイルの保存先をパスで指定します。
    strPath = "C:\Users\tosaka\Documents\Mail_Files\"
    OneDrive_Path = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\mail_files\"
     
    For Each objItem In objFolder.Items
        For i = 1 To objItem.Attachments.Count
            '添付ファイルに拡張子がある場合のみ処理します。
            If InStr(objItem.Attachments.Item(i), ".xlsx") <> 0 Then
                objItem.Attachments.Item(i).SaveAsFile strPath & objItem.Attachments.Item(i)
                
                'Excelへ添付ファイル情報を追加
                objSheet.Cells(n, 1) = objItem.ConversationTopic '件名
                objSheet.Cells(n, 2) = objItem.SenderName '送信者
                objSheet.Cells(n, 3) = objItem.ReceivedTime '受信日時
                objSheet.Cells(n, 4) = objItem.Attachments.Item(i) '添付ファイル
                objSheet.Cells(n, 5) = OneDrive_Path & objItem.Attachments.Item(i) '添付ファイルのパス"
                n = n + 1
            End If
        Next i
    
    Next objItem
    
    
    objBook.SaveAs strPath & "file_list.xlsx"  '添付ファイルを指定保存場所へ保存
    
    'ファイルを閉じて、ExcelのProcessを終了する
    objBook.Close
    myExcel.Application.Quit
    
    Set objItem = Nothing
    Set objInbox = Nothing
    Set objFolder = Nothing
    Set objSheet = Nothing
    Set objBook = Nothing
    
    Call move_files1
    Call move_files2

End Sub
Sub Filedel1()
    On Error Resume Next
    
    Const cnsSOUR1 As String = "C:\Users\tosaka\OneDrive - XXXXXXXXXX, Inc\mail_files\*.xlsx"
    
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    Dim destinationFile1 As String
    
    destinationFile1 = "C:\Users\tosaka\OneDrive - XXXXXX, Inc\mail_files\"
    
    ' FSOによるファイル削除
    If (destinationFile1) <> "" Then
        objFso.DeleteFile cnsSOUR1
        Else
    End If
    
    Set objFso = Nothing
    
End Sub
Sub Filedel2()
    On Error Resume Next
    
    Const cnsSOUR1 As String = "C:\Users\tosaka\OneDrive - XXXXXXX, Inc\mail_files\*.pdf"
    
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    Dim destinationFile1 As String
    Dim destinationFile2 As String
    
    destinationFile1 = "C:\Users\tosaka\OneDrive - XXXXXXXXX, Inc\mail_files"
    
    ' FSOによるファイル削除
    If (destinationFile1) <> "" Then
        objFso.DeleteFile cnsSOUR1
        Else
    End If
    
    Set objFso = Nothing
    
End Sub
Sub move_files1()
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFile As String
    Dim destinationFile As String
    Dim currentdir As String
    
    
    sourceFile = "C:\Users\tosaka\Documents\Mail_Files\*.xlsx"
    destinationFile = "C:\Users\tosaka\OneDrive - XXXXXXXXXX, Inc\mail_files\"
    currentdir = "C:\Users\tosaka\Documents\Mail_Files\"
    
    If (currentdir) <> "" Then
        fso.MoveFile sourceFile, destinationFile
    End If
    
    Set fso = Nothing
End Sub
Sub move_files2()

    Dim fso As New Scripting.FileSystemObject
    Dim sourceFile As String
    Dim destinationFile As String
    Dim currentdir As String
    
    
    sourceFile = "C:\Users\tosaka\Documents\Mail_Files\*.pdf"
    destinationFile = "C:\Users\tosaka\OneDrive - XXXXXXXXXX, Inc\mail_files"
    currentdir = "C:\Users\tosaka\Documents\Mail_Files"
    
    If (currentdir) <> "" Then
        fso.MoveFile sourceFile, destinationFile
    End If
    
    Set fso = Nothing
    
End Sub
