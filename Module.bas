'comboboxへ1,100を追加する関数
Sub addItemsToComboBox()
    Dim i As Integer
    For i = 1 To 100
        UserForm.box.AddItem i
    Next i
    
End Sub
'comboboxへA2,3,4のセルの値を追加する関数
Sub addItemsTocopytargetbox()
    Dim i As Integer
    For i = 2 To 4
        UserForm.copytargetbox.AddItem Range("A" & i).value
    Next i
End Sub
'丸数字を返す
Function 丸数字(ByVal n As Long) As String
    Select Case n
    Case 1 To 20
        丸数字 = Chr(Asc("�@") + n - 1)
        
    Case 21 To 35
        丸数字 = ChrW(12881 + n - 21)

    Case 36 To 50
        丸数字 = ChrW(12977 + n - 21)
        
    Case 0
        丸数字 = ChrW(9450)
        
    Case Else
        丸数字 = "(" & n & ")"
    End Select
End Function
Sub show()
    UserForm.show 0
End Sub
