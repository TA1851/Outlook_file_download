Attribute VB_Name = "Mail_list"
Option Explicit
Sub File_Auto_Saving()
    
    Call Filedel1  '�t�@�C���̕ۑ���Ԃ���ɍŐV�ɕۂ�
    Call Filedel2
    
    Dim objInbox As Object
    Dim objFolder As Object
    Dim strPath As String
    Dim i As Long
    Dim objItem As Object
    
    'Excel�p��`
    Dim myExcel As Excel.Application
    Dim objBook As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim n As Long
    
    'Excel�I�u�W�F�N�g�����A�u�b�N�̒ǉ�
    Set myExcel = CreateObject("Excel.Application")
    Set objBook = myExcel.Workbooks.Add()
    Set objSheet = objBook.Sheets(1)

    '���ږڂ�ǉ�
    objSheet.Cells(1, 1) = "subject"
    objSheet.Cells(1, 2) = "To"
    objSheet.Cells(1, 3) = "Date"
    objSheet.Cells(1, 4) = "Files"
    objSheet.Cells(1, 5) = "File_Path"
    
    '�Y�t�t�@�C�����X�g���������ލs�̈ʒu
    n = 2
    
     
    Set objInbox = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    '�Y�t�t�@�C�������郁�[���̃t�H���_���w�肵�܂��B2�K�w�ȏ゠��ꍇ�́u.Folders.Item(���t�H���_����)�v��ǉ����Ă��������B
    Set objFolder = objInbox.Folders.Item("�]���Č�").Folders.Item("�n��").Folders.Item("TEL")
    
    '�Y�t�t�@�C���̕ۑ�����p�X�Ŏw�肵�܂��B
    strPath = "C:\Users\tosaka\Documents\Mail_Files\test\"
     
    For Each objItem In objFolder.Items
        For i = 1 To objItem.Attachments.Count
            '�Y�t�t�@�C���Ɋg���q������ꍇ�̂ݏ������܂��B
            If InStr(objItem.Attachments.Item(i), ".xlsx") <> 0 Then
                objItem.Attachments.Item(i).SaveAsFile strPath & objItem.Attachments.Item(i)
                
                'Excel�֓Y�t�t�@�C������ǉ�
                objSheet.Cells(n, 1) = objItem.ConversationTopic '����
                objSheet.Cells(n, 2) = objItem.SenderName '���M��
                objSheet.Cells(n, 3) = objItem.ReceivedTime '��M����
                objSheet.Cells(n, 4) = objItem.Attachments.Item(i) '�Y�t�t�@�C��
                objSheet.Cells(n, 5) = strPath & objItem.Attachments.Item(i) '�Y�t�t�@�C���̃p�X"
                n = n + 1
            End If
        Next i
    Next objItem
 
    '�Y�t�t�@�C���ۑ��ꏊ��Excel��ۑ��@���t�@�C�����͓K���Ȗ��O�ɕς��Ă��������B
    objBook.SaveAs strPath & "file_list.xlsx"
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
    
    Const cnsSOUR1 As String = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\Excel\*.xlsx"
    
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    Dim destinationFile1 As String
    Dim destinationFile2 As String
    
    destinationFile1 = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\Excel"
    
    ' FSO�ɂ��t�@�C���폜
    If (destinationFile1) <> "" Then
        objFso.DeleteFile cnsSOUR1
        Else
    End If
    
    Set objFso = Nothing
End Sub
Sub Filedel2()
    On Error Resume Next
    
    Const cnsSOUR1 As String = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\PDF\*.pdf"
    
    Dim objFso As FileSystemObject
    Set objFso = New FileSystemObject
    Dim destinationFile1 As String
    Dim destinationFile2 As String
    
    destinationFile1 = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\PDF"
    
    ' FSO�ɂ��t�@�C���폜
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
    
    
    sourceFile = "C:\Users\tosaka\Documents\Mail_Files\test\*.xlsx"
    destinationFile = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\Excel"
    currentdir = "C:\Users\tosaka\Documents\Mail_Files\test"
    
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
    
    
    sourceFile = "C:\Users\tosaka\Documents\Mail_Files\test\*.pdf"
    destinationFile = "C:\Users\tosaka\OneDrive - Micron Technology, Inc\test\PDF"
    currentdir = "C:\Users\tosaka\Documents\Mail_Files\test"
    
    If (currentdir) <> "" Then
        fso.MoveFile sourceFile, destinationFile
    End If
    
    Set fso = Nothing
    
End Sub
