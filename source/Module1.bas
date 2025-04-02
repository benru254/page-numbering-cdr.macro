Attribute VB_Name = "Module1"
Sub ShowPageNumberForm()
    UserForm1.Show
End Sub

Sub InsertPageNumbers(frm As Object)
    Dim doc As Document
    Dim pg As Page
    Dim shp As Shape
    Dim fontSize As Double
    Dim fontName As String
    Dim margin As Double
    Dim pageIndex As Integer
    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim position As String
    Dim textContent As String
    Dim prefixSuffix As String

    ' Get user input from the form
    fontName = frm.cmbFont.Text
    fontSize = Val(frm.txtFontSize.Text)
    position = IIf(frm.optLeft.Value, "Left", IIf(frm.optRight.Value, "Right", "Center"))

    ' Use prefix/suffix if filled; otherwise, just use the number
    If Trim(frm.txtPrefixSuffix.Text) <> "" Then
        prefixSuffix = frm.txtPrefixSuffix.Text & " "
    Else
        prefixSuffix = ""
    End If

    ' Set bottom margin distance
    margin = 10

    ' Get the active document
    Set doc = ActiveDocument

    ' Loop through each page
    For Each pg In doc.Pages
        pageIndex = pg.Index
        pageWidth = pg.SizeWidth
        pageHeight = pg.SizeHeight

        ' Activate the page
        doc.Pages.Item(pageIndex).Activate

        ' Create the text shape with optional prefix/suffix
        textContent = prefixSuffix & pageIndex
        Set shp = doc.ActiveLayer.CreateArtisticText(0, margin, textContent, cdrTextAlignmentCenter)

        ' Set text properties
        shp.Text.Story.Font = fontName
        shp.Text.Story.Size = fontSize

        ' Positioning
        Select Case position
            Case "Left"
                shp.PositionX = 20
            Case "Center"
                shp.PositionX = pageWidth / 2
            Case "Right"
                shp.PositionX = pageWidth - 20
        End Select

        shp.PositionY = margin
    Next pg

    MsgBox "Page numbers inserted successfully!", vbInformation, "Done"
End Sub

