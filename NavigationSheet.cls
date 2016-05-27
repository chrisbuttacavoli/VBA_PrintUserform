VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Dim btn As Shape
    Dim sheetNames As Variant
    Dim startHeight As Double
    
    Application.ScreenUpdating = False
    sheetNames = GetWorksheetNames
    ActiveSheet.DrawingObjects.Delete
    startHeight = ActiveSheet.Range("B4").Top
    
    ' Create the sheet name buttons
    For i = 0 To UBound(sheetNames)
        Set btn = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 53.25, startHeight + i * 28, 242.25, 22.5)
        
        ' Assign text, hyperlink, and text size
        Dim sheetHyperlink As String
        If InStr(1, sheetNames(i), " ") Then
            sheetHyperlink = "'" & sheetNames(i) & "'"
        Else
            sheetHyperlink = sheetNames(i)
        End If
        ' Add hyperlink and text
        ActiveSheet.Hyperlinks.Add anchor:=btn, Address:="", SubAddress:=sheetHyperlink & "!A1"
        btn.TextFrame.Characters.Text = sheetNames(i)
            
        ' Set the font text size
        With btn.TextFrame2.TextRange.Characters().Font
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
            .Fill.ForeColor.TintAndShade = 0
            .Fill.ForeColor.Brightness = 0
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Name = "+mn-lt"
        End With
    Next
    
    ' Now we format said buttons
    With ActiveSheet.DrawingObjects.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With
    With ActiveSheet.DrawingObjects.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
    End With
    ActiveSheet.DrawingObjects.ShapeRange.TextFrame2.TextRange.Font.Name = "+mj-lt"
    With ActiveSheet.DrawingObjects.ShapeRange.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    
    MakeInsertNewSheetButton
    Application.ScreenUpdating = True
End Sub


Private Sub MakeInsertNewSheetButton()
    Dim btn As Shape
    Set btn = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 325, ActiveSheet.Range("B4").Top, 242.25, 45)
    
    With btn
        .TextFrame.Characters.Text = "Insert a new sheet"
        .OnAction = "InsertSheet"
    End With
    
    With btn.TextFrame2.TextRange.Font
        .BaselineOffset = 0
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 16
        .Name = "Cambria"
    End With
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    With btn.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.349999994
        .Transparency = 0
        .Solid
    End With
    With btn.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Transparency = 0
    End With
End Sub
