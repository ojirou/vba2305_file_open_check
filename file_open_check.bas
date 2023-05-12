Attribute VB_Name = "Module1"
'#############################################################################
' 【EXL】 ファイルが開いているかをチェック
'
'　filee_open_check
'#############################################################################
Option Explicit
Sub FileOpenCheck()
    Dim wb As Workbook, bookcheck As Workbook
    Dim openfile, openfile_path As String
    Dim bookflag As Boolean
    openfile = "list_sample.xlsx"
    openfile_path = "C:\Users\user\git\github\vba2305_file_open_check\data_table\list_sample.xlsx"
    bookflag = False
    
    For Each bookcheck In Workbooks
        If bookcheck.Name = openfile Then
            bookflag = True
        End If
    Next bookcheck
    
    If bookflag Then
        MsgBox "File is open."
        Workbooks(openfile).Activate
    Else
        MsgBox "File is not open."
        Workbooks.Open (openfile_path)
    End If
End Sub
