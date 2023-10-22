'main関数
Sub main(returnrowNumber As Integer, rowNumber As Integer)
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    Dim columnNumbers As Variant
    columnNumbers = GetColumnNumberByValue(sheetName, rowNumber, "●")
    Dim values As Variant
    values = GetValuesByColumnAndRowNumber(columnNumbers, returnrowNumber)
    Dim numbersArray As Variant
    numbersArray = GetNumbersArray(values)
    CopyArrayToClipboard numbersArray
    MsgBox "コピーしました"
End Sub


'特定の値のセルの列番号を配列として返す関数
'引数はシート名、行番号、検索する値
Function GetColumnNumberByValue(sheetName As String, rowNumber As Integer, value As String) As Variant
    Dim i As Integer
    Dim result() As Variant
    Dim count As Integer
    count = 0
    For i = 1 To Sheets(sheetName).Cells(rowNumber, Columns.count).End(xlToLeft).Column
        If Sheets(sheetName).Cells(rowNumber, i).value = value Then
            count = count + 1
            ReDim Preserve result(1 To count)
            result(count) = i
        End If
    Next i
    GetColumnNumberByValue = result
End Function

'数値の配列と数値を引数にとり、その配列は列番号として、引数の数値を行番号として持つセルの値を配列として返す関数
Function GetValuesByColumnAndRowNumber(arr As Variant, rowNumber As Integer) As Variant
    Dim i As Integer
    Dim result() As Variant
    Dim count As Integer
    count = 0
    On Error GoTo Label100
    For i = 1 To UBound(arr)
        count = count + 1
        ReDim Preserve result(1 To count)
        result(count) = ActiveSheet.Cells(rowNumber, arr(i)).value
    Next i
    On Error GoTo 0
    GetValuesByColumnAndRowNumber = result
    Exit Function

Label100:
    MsgBox "対象がありません"
    End
    
End Function

'引数に可変長の配列をとり、numbers+引数の配列として返す関数、
Function GetNumbersArray(arr As Variant) As Variant
    Dim i As Integer
    Dim numbersArray() As Variant
    ReDim numbersArray(UBound(arr))
    For i = 1 To UBound(arr)
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











