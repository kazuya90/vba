'combobox‚Ö1,100‚ğ’Ç‰Á‚·‚éŠÖ”
Sub addItemsToComboBox()
    Dim i As Integer
    For i = 1 To 100
        UserForm.rbox.AddItem i
        UserForm.box.AddItem i
    Next i
End Sub

'ŠÛ”š‚ğ•Ô‚·
Function ŠÛ”š(ByVal n As Long) As String
    Select Case n
    Case 1 To 20
        ŠÛ”š = Chr(Asc("‡@") + n - 1)
        
    Case 21 To 35
        ŠÛ”š = ChrW(12881 + n - 21)

    Case 36 To 50
        ŠÛ”š = ChrW(12977 + n - 21)
        
    Case 0
        ŠÛ”š = ChrW(9450)
        
    Case Else
        ŠÛ”š = "(" & n & ")"
    End Select
End Function