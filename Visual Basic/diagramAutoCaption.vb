Sub PictureAutoCaption()
    Dim excelApp As Object
    Dim workbook As Object
    Dim sheet As Object
    Dim doc As Document
    Dim picIndex As Integer
    Dim pic As InlineShape
    Dim captionText As String
    Dim excelFilePath As String
    Dim startPage As Integer
    Dim picRange As Range
    Dim currentPage As Integer
    
    excelFilePath = "" ' Set your Excel file path
    startPage = 3 ' Change this to your desired page number
    
    ' Open Excel
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
    If excelApp Is Nothing Then
        MsgBox "Excel is not installed or not accessible! Please close excel application", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Set workbook = excelApp.Workbooks.Open(excelFilePath) ' Open the workbook
    Set sheet = workbook.Sheets(2) ' Assuming data is on the first sheet
    
    ' Loop through pictures in the Word document
    Set doc = ActiveDocument
    picIndex = 1
    
    For Each pic In doc.InlineShapes
        ' Determine the page number of the picture
        Set picRange = pic.Range
        currentPage = picRange.Information(wdActiveEndPageNumber)
        
        ' Skip pictures inside tables
        If picRange.Tables.Count > 0 Then
            ' Picture is inside a table; skip it
            GoTo NextPicture
        End If
        
        If currentPage >= startPage Then
            ' Read the caption from Excel
            On Error Resume Next
            captionText = sheet.Cells(picIndex, 2).Value
            On Error GoTo 0
            
            ' Add caption using Word's built-in feature
            If captionText <> "" Then
                Dim Caption As String
                Caption = "Rajah " & picIndex & ": " & captionText
                
                pic.Select
                Selection.InsertCaption Label:="Rajah", TitleAutoText:="", Title:=captionText, _
                                        Position:=wdCaptionPositionBelow, ExcludeLabel:=False
                                        
                With Selection.ParagraphFormat
                    .Alignment = wdAlignParagraphCenter
                End With
            Else
                MsgBox "No caption found for picture " & picIndex, vbExclamation
            End If
            
            picIndex = picIndex + 1
        End If
NextPicture:
    Next pic

    ' Close Excel
    workbook.Close SaveChanges:=False
    excelApp.Quit
    Set excelApp = Nothing
    Set workbook = Nothing
    Set sheet = Nothing

    MsgBox "Captions added successfully!", vbInformation
End Sub
