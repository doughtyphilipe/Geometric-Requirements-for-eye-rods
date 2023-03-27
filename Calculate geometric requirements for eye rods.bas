Option Explicit

Private Sub OptionButton1_Click()

End Sub

Private Sub CommandButton1_Click()
    
End Sub

Private Sub cmdAC_Click()

End Sub
Private Sub cmdBerechnen_Click()
Dim Gamma As Single
'MA Dim
Dim t1 As Single, Fed1 As Single, fy1 As Single, d01 As Single, a As Single, c As Single
'MB Dim
Dim Fed2 As Single, fy2 As Single, t2 As Single, d02 As Single
Dim d0_03 As Single, d0_075 As Single, d0_13 As Single, d0_16 As Single, d0_25 As Single
'Array Dim
'Script by Philipe Doughty - TKE Intern from 01.09.21 until 28.02.22 -  doughtyphilipe@gmail.com


If optbutMA.Value = True Then
    'input MA Values
    Gamma = boxGamma.Text
    t1 = boxT1.Text
    Fed1 = boxFed1.Text
    fy1 = boxFy1.Text
    d01 = boxD0.Text
    'Calculate a and c
    a = (Fed1 * Gamma) / (2 * t1 * fy1) + 2 * d01 / 3
    c = (Fed1 * Gamma) / (2 * t1 * fy1) + d01 / 3
    'Output MA
    lblA1.Caption = FormatNumber(a, 1)
    lblC1.Caption = FormatNumber(c, 1)
    
    'print values to Spreadsheet
    If exportBox.Value = True Then
        'Create Array
        Dim list1(1 To 6, 1 To 2) As Variant
        Dim i As Single, j As Single
        list1(1, 1) = "t [mm]"
        list1(1, 2) = t1
        list1(2, 1) = "Fed1 [N]"
        list1(2, 2) = Fed1
        list1(3, 1) = "fy1 [N/mm2]"
        list1(3, 2) = fy1
        list1(4, 1) = "d_0 [mm]"
        list1(4, 2) = d01
        list1(5, 1) = "a [mm]"
        list1(5, 2) = a
        list1(6, 1) = "c [mm]"
        list1(6, 2) = c
        'Populate Spreadsheet
        For i = 1 To UBound(list1, 1)
            For j = 1 To UBound(list1, 2)
            Cells(i, j).Value = list1(i, j)
            Next j
        Next i
        
    End If
    
    
    
ElseIf optbutMB.Value = True Then
    'input MB Values
    Fed2 = boxFed2.Text
    fy2 = boxFy2.Text
    Gamma = boxGamma.Text
    
    'Calculate t and d0
    t2 = 0.7 * ((Fed2 * Gamma) / fy2) ^ (0.5)
    d02 = 2.5 * t2
    d0_03 = 0.3 * d02
    d0_075 = 0.75 * d02
    d0_13 = 1.3 * d02
    d0_16 = 1.6 * d02
    d0_25 = 2.5 * d02
    
    'Output MB
    lblT2.Caption = FormatNumber(t2, 1)
    lblD02.Caption = FormatNumber(d02, 1)
    lblD0_03.Caption = FormatNumber(d0_03, 1)
    lblD0_075.Caption = FormatNumber(d0_075, 1)
    lblD0_13.Caption = FormatNumber(d0_13, 1)
    lblD0_16.Caption = FormatNumber(d0_16, 1)
    lblD0_25.Caption = FormatNumber(d0_25, 1)
    
    'print values to Spreadsheet
    If exportBox.Value = True Then
        'Create Array
        Dim list2(1 To 9, 1 To 2) As Variant
        Dim m As Single, n As Single
        list2(1, 1) = "Fed2 [N]"
        list2(1, 2) = Fed2
        list2(2, 1) = "fy2 [N/mm2]"
        list2(2, 2) = fy2
        list2(3, 1) = "t2 [mm]"
        list2(3, 2) = t2
        list2(4, 1) = "d0 [mm]"
        list2(4, 2) = d02
        list2(5, 1) = "0,3*d0 [mm]"
        list2(5, 2) = d0_03
        list2(6, 1) = "0,75*d0 [mm]"
        list2(6, 2) = d0_075
        list2(7, 1) = "1,3*d0 [mm]"
        list2(7, 2) = d0_13
        list2(8, 1) = "1,6*d0 [mm]"
        list2(8, 2) = d0_16
        list2(9, 1) = "2,5*d0 [mm]"
        list2(9, 2) = d0_25
        'Populate Spreadsheet
        For m = 1 To UBound(list2, 1)
            For n = 1 To UBound(list2, 2)
            Cells(m, n).Value = list2(m, n)
            Next n
        Next m
    End If
    
    
End If

End Sub

Private Sub cmdClear_Click()
    Call clearAll
    Cells.Clear
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub fraInputMB_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub optbutMA_Click()
    fraInputMA.Visible = True
    fraInputMB.Visible = False

    
    imgMB.Visible = False
    imgMA.Visible = True
    
    Call clearAll

End Sub

Private Sub optbutMB_Click()
    fraInputMB.Visible = True
    fraInputMA.Visible = False

    
    imgMA.Visible = False
    imgMB.Visible = True
    
    Call clearAll
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub clearAll()
    boxT1.Text = ""
    boxFed1.Text = ""
    boxFy1.Text = ""
    boxD0.Text = ""
    lblA1.Caption = ""
    lblC1.Caption = ""
    boxFed2.Text = ""
    boxFy2.Text = ""
    lblT2.Caption = ""
    lblD02.Caption = ""
    lblD0_03.Caption = ""
    lblD0_075.Caption = ""
    lblD0_13.Caption = ""
    lblD0_16.Caption = ""
    lblD0_25.Caption = ""
End Sub
