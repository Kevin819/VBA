Attribute VB_Name = "ģ��1"

Sub Fenshu()

    Dim dataExcel, Workbook, mySheet
    mypath = ThisWorkbook.Path + "\"
    fileNameList = Dir(mypath, vbDirectory)
    Set dataExcel = CreateObject("Excel.Application")
    a = 2
    Do While fileNameList <> ""
        If fileNameList <> "��������.xlsx" And fileNameList <> "." And fileNameList <> ".." Then
            tmpName = mypath + fileNameList
            Set Workbook = dataExcel.Workbooks.Open(tmpName)
            Set mySheet = Workbook.Worksheets(1)
            Sheets("sheet1").Cells(a, 2) = mySheet.Cells(4, 2)
            Sheets("sheet1").Cells(a, 3) = mySheet.Cells(18, 2)
            a = a + 1
            Workbook.Close
        End If
        fileNameList = Dir
    Loop
    MsgBox "��ȡ�ɹ���", vbSystemModal '��ȡ��󵯿�����

End Sub
