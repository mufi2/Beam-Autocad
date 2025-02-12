Sub Draw(p1x As Variant, p1y As Variant, p2x As Variant, p2y As Variant, z As Double)
    Dim j As Integer
    Dim p1 As Object
    Dim p2 As Object
    Dim line1 As Object

    For j = LBound(p1x) To UBound(p1x)
        Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(p1x(j), p1y(j) + z, 0))
        Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(p2x(j), p2y(j) + z, 0))
        Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
        line1.Layer = "Beam_LSection"
    Next j
End Sub

Sub LDimDraw(long_dim_x0 As Variant, long_dim_y0 As Variant, long_dim_x1 As Variant, long_dim_y1 As Variant, z As Double)
    Dim j As Integer
    Dim p1 As Object
    Dim p2 As Object
    Dim line1 As Object

    For j = LBound(long_dim_x0) To UBound(long_dim_x0)
        Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(long_dim_x0(j), long_dim_y0(j) + z, 0))
        Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(long_dim_x1(j), long_dim_y1(j) + z, 0))
        Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
        line1.Layer = "text"
    Next j
End Sub
Sub DrawLongitudinalRebar(p1x As Variant, p1y As Variant, p2x As Variant, p2y As Variant, z As Double)
    Dim j As Integer
    Dim p1 As Object
    Dim p2 As Object
    Dim line1 As Object

    For j = LBound(p1x) To UBound(p1x)
        Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(p1x(j), p1y(j) + z, 0))
        Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(p2x(j), p2y(j) + z, 0))
        Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
        line1.Layer = "Long_rebar"
    Next j
End Sub

Sub DrawWithoutCantilever(column_width As Double, column_cc As Double, a As Double, a_initial As Double, d As Double, beam_cc As Double, z As Double)
    Dim p3x(5) As Double
    Dim p3y(5) As Double
    Dim p4x(5) As Double
    Dim p4y(5) As Double

    ' Lower longitudinal rebar (first three points)
    p3x(0) = a_initial - column_width + column_cc
    p3x(1) = a_initial - column_width + column_cc
    p3x(2) = a - column_cc
    p3x(3) = a_initial - column_width + column_cc
    p3x(4) = a_initial - column_width + column_cc
    p3x(5) = a - column_cc

    p3y(0) = beam_cc
    p3y(1) = beam_cc
    p3y(2) = beam_cc
    p3y(3) = d - beam_cc
    p3y(4) = d - beam_cc
    p3y(5) = d - beam_cc

    p4x(0) = a_initial - column_width + column_cc
    p4x(1) = a - column_cc
    p4x(2) = a - column_cc
    p4x(3) = a_initial - column_width + column_cc
    p4x(4) = a - column_cc
    p4x(5) = a - column_cc

    p4y(0) = d * 2
    p4y(1) = beam_cc
    p4y(2) = d * 2
    p4y(3) = -d
    p4y(4) = d - beam_cc
    p4y(5) = -d

    ' Call DrawLongitudinalRebar to draw the lines
    Call DrawLongitudinalRebar(p3x, p3y, p4x, p4y, z)
End Sub
Sub Stirrup(l_next As Double, z As Double, d As Double, cc As Double, a As Double)
    Dim l As Double
    Dim spacing_left As Integer
    Dim spacing_center As Integer
    Dim spacing_right As Integer
    Dim n_left As Integer, n_center As Integer, n_right As Integer
    Dim stirrup_dia As Integer
    Dim sp As Double
    Dim p1 As Object, p2 As Object, ptext_1 As Object, ptext_2 As Object
    Dim text_dim(2) As Variant
    Dim text_length As Object
    Dim line1 As Object, line2 As Object, line3 As Object

    l = l_next
    stirrup_dia = 8 ' Fixed stirrup diameter, can be replaced by user input if needed

    ' Get user input for spacing on different sections of the beam
    spacing_left = InputBox("Enter Spacing(mm) on left section(l/3) of Beam : ")
    spacing_center = InputBox("Enter Spacing(mm) on center section of Beam : ")
    spacing_right = InputBox("Enter Spacing(mm) on right section(l/3) of Beam : ")

    sp = 0 ' Initialize spacing

    ' Process for stirrup lines (left section)
    n_left = Round(((l / 3 - 50 * 2) / spacing_left)) + 1
    ' Annotate stirrup information
    Set ptext_1 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 6, d / 2 + z, 0))
    Set ptext_2 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 6 + 46.6, -86.4 + z, 0))
    Set line1 = ThisDrawing.ModelSpace.AddLine(ptext_1, ptext_2)
    line1.Layer = "text"
    
    ' Flat line for annotation
    Set ptext_3 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 6 + 46.6, -86.4 + z, 0))
    Set ptext_4 = ThisDrawing.ModelSpace.AddPoint(Array(a + 46.6 + 77 + l / 6, -86.4 + z, 0))
    Set line2 = ThisDrawing.ModelSpace.AddLine(ptext_3, ptext_4)
    line2.Layer = "text"

    ' Text to be shown for stirrup dimensions
    text_dim(0) = n_left
    text_dim(1) = stirrup_dia
    text_dim(2) = spacing_left
    Set text_length = ThisDrawing.ModelSpace.AddText(" " & text_dim(0) & "-" & text_dim(1) & "mm dia@ " & text_dim(2) & "mm c/c", ptext_4, 25)
    text_length.Layer = "text"

    ' Stirrup lines at the middle
    Set ptext_5 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 6 - 250, d / 2 + z, 0))
    Set ptext_6 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 6 + 250, d / 2 + z, 0))
    Set line3 = ThisDrawing.ModelSpace.AddLine(ptext_5, ptext_6)
    line3.Linetype = "strrp"

    ' Create stirrup lines on the left section
    For j = 1 To n_left
        If j = 1 Then
            Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(a + 50, cc + z, 0))
            Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(a + 50, z + d - cc, 0))
            Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
            line1.Layer = "Stirrups_sides"
            sp = sp + 50
        Else
            sp = sp + spacing_left
            Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + cc, 0))
            Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + d - cc, 0))
            Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
            line1.Layer = "Stirrups_sides"
        End If
    Next j

    ' Process for stirrup lines (center section)
    n_center = Round(((l / 3 - 50 * 2) / spacing_center)) + 1
    ' Annotate stirrup information for the center section
    Set ptext_1 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 2, d / 2 + z, 0))
    Set ptext_2 = ThisDrawing.ModelSpace.AddPoint(Array(a + 46.6 + l / 2, d + z + 86.4, 0))
    Set line1 = ThisDrawing.ModelSpace.AddLine(ptext_1, ptext_2)
    line1.Layer = "text"
    
    ' Flat line for annotation
    Set ptext_3 = ThisDrawing.ModelSpace.AddPoint(Array(a + 46.6 + l / 2, d + z + 86.4, 0))
    Set ptext_4 = ThisDrawing.ModelSpace.AddPoint(Array(a + 46.6 + 77 + l / 2, d + z + 86.4, 0))
    Set line2 = ThisDrawing.ModelSpace.AddLine(ptext_3, ptext_4)
    line2.Layer = "text"

    ' Text to be shown for stirrup dimensions
    text_dim(0) = n_center
    text_dim(1) = stirrup_dia
    text_dim(2) = spacing_center
    Set text_length = ThisDrawing.ModelSpace.AddText(" " & text_dim(0) & "-" & text_dim(1) & "mm dia@ " & text_dim(2) & "mm c/c", ptext_4, 25)
    text_length.Layer = "text"

    ' Stirrup lines at the middle
    Set ptext_5 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 2 - 250, d / 2 + z, 0))
    Set ptext_6 = ThisDrawing.ModelSpace.AddPoint(Array(a + l / 2 + 250, d / 2 + z, 0))
    Set line3 = ThisDrawing.ModelSpace.AddLine(ptext_5, ptext_6)
    line3.Linetype = "strrp"

    ' Create stirrup lines on the center section
    For j = 1 To n_center
        sp = sp + spacing_center
        Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + cc, 0))
        Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + d - cc, 0))
        Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
        line1.Layer = "Stirrups_center"
    Next j

    ' Process for stirrup lines (right section)
    n_right = Round(((l / 3 - 50 * 2) / spacing_right)) + 1
    ' Annotate stirrup information for the right section
    Set ptext_1 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6, d / 2 + z, 0))
    Set ptext_2 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 - 46.6, -86.4 + z, 0))
    Set line1 = ThisDrawing.ModelSpace.AddLine(ptext_1, ptext_2)
    line1.Layer = "text"
    
    ' Flat line for annotation
    Set ptext_3 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 - 46.6, -86.4 + z, 0))
    Set ptext_4 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 - 46.6 - 77, -86.4 + z, 0))
    Set ptext_5 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 - 46.6 - 77 - 468, -86.4 + z, 0))
    Set line2 = ThisDrawing.ModelSpace.AddLine(ptext_3, ptext_4)
    line2.Layer = "text"

    ' Text to be shown for stirrup dimensions
    text_dim(0) = n_right
    text_dim(1) = stirrup_dia
    text_dim(2) = spacing_right
    Set text_length = ThisDrawing.ModelSpace.AddText(" " & text_dim(0) & "-" & text_dim(1) & "mm dia@ " & text_dim(2) & "mm c/c", ptext_5, 25)
    text_length.Layer = "text"

    ' Stirrup lines at the middle
    Set ptext_5 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 - 250, d / 2 + z, 0))
    Set ptext_6 = ThisDrawing.ModelSpace.AddPoint(Array(a + 5 * l / 6 + 250, d / 2 + z, 0))
    Set line3 = ThisDrawing.ModelSpace.AddLine(ptext_5, ptext_6)
    line3.Linetype = "strrp"

    ' Create stirrup lines on the right section
    For j = 1 To n_right
        sp = sp + spacing_right
        If sp > l - 50 Then
            sp = sp - 50
            Exit For
        End If
        Set p1 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + cc, 0))
        Set p2 = ThisDrawing.ModelSpace.AddPoint(Array(a + sp, z + d - cc, 0))
        Set line1 = ThisDrawing.ModelSpace.AddLine(p1, p2)
        line1.Layer = "Stirrups_sides"
    Next j
End Sub
