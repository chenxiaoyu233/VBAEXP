Attribute VB_Name = "main"
' Global Var
Global excel, word As Object
Global lookup, template As Object
Global wordDir, excelPath, templatePath As String
Global excelWorkSheet As String
Global houseIDCol As String
Global houseIDRowL, houseIDRowR As Integer
Global inlineShapeID As Integer
Global templateContentL, templateContentR As Integer
Private Sub setupConstant()
    wordDir = "/Users/chenxiaoyu/Desktop/docx-excel/docx/"
    excelPath = "/Users/chenxiaoyu/Desktop/docx-excel/excel/lookup.xls"
    templatePath = "/Users/chenxiaoyu/Desktop/docx-excel/template/template.xlsx"
End Sub
' read the template.xlsx and set some vars
Private Sub readTemplateHead()
    Set template = excel.Application.WorkBooks.Open(filename:=templatePath, ReadOnly:=True).WorkSheets("Sheet1")
    excelWorkSheet = template.Range("A1").Text
    houseIDCol = template.Range("A2").Text
    houseIDRowL = CInt(template.Range("B2").Value)
    houseIDRowR = CInt(template.Range("C2").Value)
    inlineShapeID = CInt(template.Range("A3").Value)
    templateContentL = CInt(template.Range("A4").Value)
    templateContentR = CInt(template.Range("B4").Value)
    Set lookup = excel.Application.WorkBooks.Open(filename:=excelPath, ReadOnly:=True).WorkSheets(excelWorkSheet)
End Sub
' Program Entry
Sub main()
    setupConstant
    setupHandle
    readTemplateHead
    Dim Line As Integer
    For Line = houseIDRowL To houseIDRowR
        Dim c As Object
        Set c = lookup.Range(houseIDCol & Line)
        If c.Value <> "" Then
            Dim file As String
            file = queryWordFile(c.Value)
            If file <> "" Then
                word.Application.Documents.Open (file)
                Call setContent(Line, file)
                'I do not know how this code works.
                'But if I don't write it like this, the changes could not save in the .docx files.
                'Maybe just another MS Office bug.
                'This annoying thing wastes me a lot of time.
                'I do not want to do anything related with this.
                word.Application.Activate
                word.Application.Documents(file).Activate
                word.Application.Documents(file).Save
                word.Application.Documents(file).Close
            End If
        End If
    Next
    'MsgBox ("Done all")
    quitHandle
End Sub
' set the Content by the lookup excel and template
Private Sub setContent(ByVal Line As Integer, ByVal file As String)
    Dim name As String
    word.Application.Documents(file).InlineShapes(inlineShapeID).OLEFormat.Edit
    name = excel.Application.ActiveWorkbook.name
    Dim curExcel As Object
    Set curExcel = excel.Application.WorkBooks(name).WorkSheets("Sheet1")
    For pt = templateContentL To templateContentR
        Dim target, kind, info, content As String
        target = template.Range("A" & pt).Value
        kind = template.Range("B" & pt).Value
        info = template.Range("C" & pt).Value
        If kind = "content" Then
            content = info
        ElseIf kind = "copy" Then
            content = lookup.Range(info & Line).Value
        End If
       curExcel.Range(target).Value = content
    Next
    ' clean up
    excel.Application.WorkBooks(curExcel.Parent.name).Save
    excel.Application.WorkBooks(curExcel.Parent.name).Close
End Sub

' test the open and close of a document
Private Sub testOpenDocument()
    setupConstant
    setupExcel
    Dim file As String
    file = "/Users/chenxiaoyu/Desktop/docx-excel/docx/97-A40’‘¥∫¡’.doc"
    Documents.Open (file)
    Dim tmp As String
    Documents(file).InlineShapes(2).OLEFormat.Activate
    tmp = excel.Application.ActiveWorkbook.name
    MsgBox (excel.Application.WorkBooks(tmp).WorkSheets("Sheet1").Range("B2").Value)
    'MsgBox (tmp.Range("3B").Value)
    MsgBox ("open")
    Documents(file).Close
    quitExcel
End Sub
' Use House ID to get the current wordFile
' @param string houseID
' @return string querywordFile
'   case find: filepath
'   case not find: empty string
Private Function queryWordFile(ByVal houseID As String)
    Dim file As String
    file = Dir(wordDir)
    Do While file <> ""
        Dim pos As Integer
        pos = InStr(1, file, houseID)
        If pos <> 0 Then
            queryWordFile = wordDir & file
            Exit Function
        End If
        file = Dir
    Loop
    queryWordFile = ""
End Function
' test module for the function: queryWordFile
Private Sub testQueryWordFile()
    setupConstant
    Dim file As String
    Dim msg As Object
    file = queryWordFile(houseID:="A150")
    MsgBox (file)
End Sub
' Open An Excel Application with the main table loaded
Private Sub setupHandle()
    Set excel = CreateObject("Excel.Sheet")
    Set word = CreateObject("Word.Document")
End Sub
' Close the Excel when all the tasks are done
Private Sub quitHandle()
    excel.Application.WorkBooks(template.Parent.name).Close
    excel.Application.WorkBooks(lookup.Parent.name).Close
    word.Close
    excel.Close
End Sub

