Sub show()
    UserForm.show 1
End Sub

'main関数
Sub main(returnNumber As Integer, colNumber As Integer)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    Dim rownumbers As Variant
    rownumbers = FindRowsWithValue(sheetName, colNumber, "●")
    Dim values As Variant
    values = GetValuesByRowAndColumn(sheetName, rownumbers, returnNumber)
    Dim numbersArray As Variant
    numbersArray = GetNumbersArray(values)
    CopyArrayToClipboard (numbersArray)
End Sub

'アクティブなセルの範囲を配列として返す関数
'一次元
Function GetActiveRange() As Variant
    Dim activeRange As Range
    Dim arr() As Variant
    Set activeRange = Application.Selection
    ReDim arr(activeRange.count - 1)
    Dim i As Integer
    For i = 0 To activeRange.count - 1
        arr(i) = activeRange(i + 1)
    Next i
    GetActiveRange = arr
End Function

'引数に可変長の配列をとり、numbers+引数の配列として返す関数、
Function GetNumbersArray(arr As Variant) As Variant
    Dim i As Integer
    Dim numbersArray() As Variant
    ReDim numbersArray(UBound(arr))

    
    For i = 1 To UBound(arr)
    if(isContainsNewLine(arr(i)))then
        numbersArray(i) = 丸数字(i) & vbCrLf & arr(i)
    else
        numbersArray(i) = 丸数字(i) & arr(i)
    Next i
    GetNumbersArray = numbersArray
End Function

'引数に配列をとり、その各配列を結合後、クリップボードへ格納する関数
Sub CopyArrayToClipboard(arr As Variant)
    Dim i As Integer
    Dim str As String
    For i = 0 To UBound(arr)
        str = str & arr(i) & vbCrLf
    Next i
    With New DataObject
        .SetText str
        .PutInClipboard
    End With
End Sub

'丸数字を返す
Function 丸数字(ByVal n As Long) As String
    Select Case n
    Case 1 To 20
        丸数字 = Chr(Asc("①") + n - 1)
        
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

'特定の値のセルの行番号を配列として返す関数
Function FindRowsWithValue(sheetName As String, columnNumber As Integer, searchValue As String) As Variant
  Dim lastRow As Long
  Dim i As Long
  Dim result() As Long
  Dim count As Long
  
  lastRow = Sheets(sheetName).Cells(Rows.count, columnNumber).End(xlUp).Row
  
  ReDim result(1 To lastRow)
  
  For i = 1 To lastRow
    If Sheets(sheetName).Cells(i, columnNumber).Value = searchValue Then
      count = count + 1
      result(count) = i
    End If
  Next i
  
  If count = 0 Then
    FindRowsWithValue = Empty
  Else
    ReDim Preserve result(1 To count)
    FindRowsWithValue = result
  End If
End Function

'数値の配列と数値を引数にとり、その配列は行番号として、引数の数値を列番号として持つセルの値を配列として返す関数
Function GetValuesByRowAndColumn(sheetName As String, rownumbers As Variant, columnNumber As Integer) As Variant
  Dim i As Long
  Dim result() As Variant
  Dim count As Long
  
  If IsEmpty(rownumbers) Then
    MsgBox "コピー対象がありません"
    End
  Else
  
  ReDim result(1 To UBound(rownumbers))
  
  For i = 1 To UBound(rownumbers)
    result(i) = Sheets(sheetName).Cells(rownumbers(i), columnNumber).Value
  Next i
  
  GetValuesByRowAndColumn = result
  End If
End Function

'comboboxへ1,100を追加する関数
Sub addItemsToComboBox()
    Dim i As Integer
    For i = 1 To 100
        UserForm.ComboBox1.AddItem i
    Next i
End Sub

'改行を含む文字列か判断する関数
Function isContainsNewLine(str As String) As Boolean
    If InStr(str, vbCrLf) > 0 Then
        isContainsNewLine = True
    Else
        isContainsNewLine = False
    End If
End Function