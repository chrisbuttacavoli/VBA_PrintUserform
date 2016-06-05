Attribute VB_Name = "Module1"
' ////////////////////////////////////////////////////////////////////////////
' // Copyright (C) 2016 by Chris Buttacavoli
' // All rights reserved. This coding project may not be reproduced, modified,
' // or used in any manner whatsoever without the express permission of the
' // author.
' //
' // Author: Chris Buttacavoli
' // Project: Printing ODOT calculations
' // Requirements provided by: Britney Buttacavoli
' // Source Code: https://github.com/chrisbuttacavoli/VBA_PrintUserform
' //
' // Description: This project was created to assist users in choosing which
' // sheets to print from an Excel workbook. Users can select and reorder
' // pages and run the print dialog from the userform.
' ////////////////////////////////////////////////////////////////////////////

Public Sub OpenPrintForm()
    frmPrint.Show
End Sub


' Spits out an array of all sheet names in this workbook
Public Function GetWorksheetNames() As Variant
    Dim allSheets() As String
    Dim wks As Worksheet
    Dim i As Integer
    
    ReDim allSheets(Worksheets.Count - 1)
    i = 0
    
    ' Save all sheet names to an array
    For Each wks In Worksheets
        allSheets(i) = wks.Name
        i = i + 1
    Next
    
    GetWorksheetNames = allSheets()
End Function


Public Sub InsertSheet()
    Dim val As Variant
    
    val = InputBox("Enter the new sheet name as you want it to appear " & _
                "on the Table of Contents", "Enter a new sheet name")
    
    ' Check to see if they entered in anything
    If Not IsNull(val) Then
        
        ' Replace any special characters the user inputted
        Dim sheetName As String
        sheetName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(val, "`", ""), "~", ""), "!", ""), "@", ""), "#", ""), "$", ""), "%", ""), "^", ""), "&", ""), "*", ""), "(", ""), ")", ""), "=", ""), "+", ""), "\", ""), "|", ""), "[", ""), "{", ""), "]", ""), "}", ""), ";", ""), ":", ""), "'", ""), """", ""), ",", ""), "<", ""), ".", ""), ">", ""), "/", ""), "?", "")
        
        If Len(sheetName) = 0 Then Exit Sub
        If Len(sheetName) > 31 Then
            MsgBox "The sheet name you provided was " & Len(sheetName) & " characters. 31 characters is the maximum length."
            Exit Sub
        End If
        
        ' Create the sheet
        Dim ws As Worksheet
        With ThisWorkbook
            Set ws = ThisWorkbook.Sheets.Add(After:=.Sheets(.Sheets.Count))
        End With
        ws.Name = sheetName
        
        ' Set formatting on new sheet
        With ws
            With .Cells.Font
                .Name = "Cambria"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMajor
            End With
            With .Range("B2")
                .Value = sheetName
                With .Font
                    .Name = "Cambria"
                    .FontStyle = "Bold"
                    .Size = 18
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .ThemeFont = xlThemeFontNone
                End With
            End With
        End With
    End If
    
    ' Format the header and footer
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ws.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "&""Cambria,Regular""Printed: &D at &T"
        .LeftFooter = "&""Cambria,Regular""&Z&F"
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ws.Range("A1").Select
End Sub


' Self explanatory
Public Function ExistsInArray(val, arr) As Boolean
    Dim i As Integer
    ExistsInArray = False
    For i = 0 To UBound(arr)
        If val = arr(i) Then
            ExistsInArray = True
            Exit For
        End If
    Next
End Function
