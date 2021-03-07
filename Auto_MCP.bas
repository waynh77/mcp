Attribute VB_Name = "Auto_MCP"

Public Function Peg_auto()
Dim urutan As String * 10
Dim hitung As Single
With Peg_frm.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "NIP" & "0000001"
    Else
        .MoveLast
        If Val(Left(.Fields("No_peg"), 7)) <> "0000000" Then
            urutan = "0000000" & "0000001"
        Else
            hitung = Val(Right(.Fields("no_peg"), 7)) + 1
            urutan = "NIP" & Right("0000000" & hitung, 7)
        End If
    End If
    Peg_frm.Text1(0).Text = urutan
End With
End Function

Public Function inv_auto()
End Function

