'combobox��1,100��ǉ�����֐�
Sub addItemsToComboBox()
    Dim i As Integer
    For i = 1 To 100
        UserForm.rbox.AddItem i
        UserForm.box.AddItem i
    Next i
End Sub

'�ې�����Ԃ�
Function �ې���(ByVal n As Long) As String
    Select Case n
    Case 1 To 20
        �ې��� = Chr(Asc("�@") + n - 1)
        
    Case 21 To 35
        �ې��� = ChrW(12881 + n - 21)

    Case 36 To 50
        �ې��� = ChrW(12977 + n - 21)
        
    Case 0
        �ې��� = ChrW(9450)
        
    Case Else
        �ې��� = "(" & n & ")"
    End Select
End Function