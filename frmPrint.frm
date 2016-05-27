VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrint 
   Caption         =   "Print Setup"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3405
   OleObjectBlob   =   "frmPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

' Set up choices
Private Sub UserForm_Initialize()
    Dim sheetNames As Variant
    sheetNames = GetWorksheetNames
    
    ' Populate options here
    cbChoice.List = sheetNames
    lstSelected.Clear
    cmdUp.Caption = ChrW(8593)
    cmdDown.Caption = ChrW(8595)
    
    ' Add previously selected items
    Dim rng As Range
    Dim cell As Range
    Set rng = Sheets("Table of Contents").Range("B6:B50")
    For Each cell In rng
        If cell.Value <> "" Then
            lstSelected.AddItem cell.Value
        End If
    Next
    
    ' Choose our control to start on
    Me.cbChoice.SetFocus
End Sub


' Magic happens here my friends...
Private Sub cmdPrint_Click()
    Dim numSheets As Integer
    numSheets = lstSelected.ListCount
    
    If numSheets > 0 Then
        ' Show the message to user
        lblMessage.Visible = True
        
        ' Begin our fun stuff
        Application.ScreenUpdating = False
        Dim printList() As String
        
        ' First clear table of contents sheet
        Sheets("Table of Contents").Range("B6:B50").ClearContents
        
        ' Dump all selected sheets into an array
        ReDim printList(numSheets - 1)
        Dim j As Integer
        For i = 0 To UBound(printList)
            printList(i) = lstSelected.List(i)
            
            ' Set items in Table of Contents sheet while we're here but do not include certain pages on it
            If Not (lstSelected.List(i) = "Title Sheet" Or lstSelected.List(i) = "Table of Contents") Then
                Sheets("Table of Contents").Range("B6").Offset(j, 0).Value = lstSelected.List(i)
                j = j + 1
            End If
        Next
        
        ' Dance-a little monkey, dance-a....
        Me.Hide
        lblMessage.Visible = False
        Application.ScreenUpdating = True
        
        ' Rock 'n roll baby!
        Application.Dialogs(xlDialogPrinterSetup).Show
        Worksheets(printList).PrintOut , preview:=True
        
        ' Close the form
        Sheets("Table of Contents").Select
    Else
        MsgBox "You did not select any sheets to print"
    End If
End Sub


' Adds item to list box after choosing something in the drop down
Private Sub cbChoice_Change()
    If cbChoice.Value <> "" Then lstSelected.AddItem cbChoice.Value
End Sub


' Moves a single list item up
Private Sub cmdUp_Click()
    Dim selectedItem As Integer
    selectedItem = CanMove("Up")
    
    If selectedItem <> -1 Then
        Dim oldItem As String
        
        ' Swap with the one above it
        With lstSelected
            oldItem = .List(selectedItem - 1)
            .List(selectedItem - 1) = .List(selectedItem)
            .List(selectedItem) = oldItem
            .selected(selectedItem) = False
            .selected(selectedItem - 1) = True
        End With
    End If
End Sub


' Moves a single list item down
Private Sub cmdDown_Click()
    Dim selectedItem As Integer
    selectedItem = CanMove("Down")
    
    If selectedItem <> -1 Then
        Dim oldItem As String
        
        ' Swap with the one below it
        With lstSelected
            oldItem = .List(selectedItem + 1)
            .List(selectedItem + 1) = .List(selectedItem)
            .List(selectedItem) = oldItem
            .selected(selectedItem) = False
            .selected(selectedItem + 1) = True
        End With
    End If
End Sub


' Clears all items in the list box
Private Sub cmdClearAll_Click()
    lstSelected.Clear
End Sub


' Clears only selected items in the list box
Private Sub cmdClearSelected_Click()
    Dim i As Integer
    
    On Error Resume Next
    For i = Me.lstSelected.ListCount - 1 To 0 Step -1
        If Me.lstSelected.selected(i) Then
            Me.lstSelected.RemoveItem i
        End If
    Next
End Sub


' Helper function to check if we can move an item up or down the list box.
' Returns -1 if invalid move, otherwise returns position of selected item.
Private Function CanMove(direction As String) As Integer
    Dim i As Integer
    Dim selectedItem As Integer
    Dim numSelected As Integer

    For i = 0 To lstSelected.ListCount - 1
        If lstSelected.selected(i) Then
            numSelected = numSelected + 1
            selectedItem = i
        End If
    Next
    
    If (numSelected > 1) Or _
        (selectedItem = 0 And direction = "Up") Or _
        (selectedItem = lstSelected.ListCount - 1 And direction <> "Up") Then
        
        CanMove = -1
    Else
        CanMove = selectedItem
    End If
End Function


