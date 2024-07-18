Private Sub FillButton1_Click()
    Const StudentsListSheet = "Members"
    Const TargetSheet = "Layout1"
    Dim R As Integer
    Const Rmax As Integer = 11
    Const Rmin As Integer = 4
    R = Rmin
    Dim C As Integer
    Const Cmax As Integer = 6
    Const Cmin As Integer = 1
    C = Cmax
    With Worksheets(StudentsListSheet)
        For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
            Worksheets(TargetSheet).Cells(R, C).Value = .Cells(i, 1).Value & " " & .Cells(i, 2).Value
            R = R + 1
            If R > Rmax Then
                R = Rmin
                C = C - 1
            End If
        Next
    End With
    MsgBox "Layout1に出席番号順に格納完了しました"
End Sub
