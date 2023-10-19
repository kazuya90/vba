Function GetActiveRange() As Variant
  Dim activeRange As Range
  Set activeRange = Selection
  
  GetActiveRange = activeRange.Value
End Function

'GetActiveRangeで取得した値を結合して、クリップボードにコピーする
Sub CopyActiveRange()
  Dim activeRange As Variant
  activeRange = GetActiveRange()
  
  Dim i As Integer
  Dim j As Integer
  Dim str As String
  For i = 1 To UBound(activeRange, 1)
    For j = 1 To UBound(activeRange, 2)
      str = str & activeRange(i, j)
    Next j
  Next i
  
  Dim dataObj As New MSForms.DataObject
  dataObj.SetText str
  dataObj.PutInClipboard
End Sub

'GetActiveRangeで取得した値の先頭へ①②③④などの番号をつけて、クリップボードにコピーする
Sub CopyActiveRangeWithNumber()
  Dim activeRange As Variant
  activeRange = GetActiveRange()
  
  Dim i As Integer
  Dim j As Integer
  Dim str As String
  For i = 1 To UBound(activeRange, 1)
    For j = 1 To UBound(activeRange, 2)
      str = str & i & "." & j & " " & activeRange(i, j) & vbCrLf
    Next j
  Next i
  
  Dim dataObj As New MSForms.DataObject
  dataObj.SetText str
  dataObj.PutInClipboard
End Sub
'①-⑳の番号を配列として宣言する
Dim numbers As Variant

'①-⑳の番号を配列に格納する
numbers = Array("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳")

'GetActiveRangeで取得した値の先頭へ①②③④の番号をつけて、クリップボードにコピーする
Sub CopyActiveRangeWithNumber2()
  Dim activeRange As Variant
  activeRange = GetActiveRange()
  
  Dim i As Integer
  Dim j As Integer
  Dim str As String
  For i = 1 To UBound(activeRange, 1)
    For j = 1 To UBound(activeRange, 2)
      str = str & numbers(i - 1) & " " & activeRange(i, j) & vbCrLf
    Next j
  Next i
  
  Dim dataObj As New MSForms.DataObject
  dataObj.SetText str
  dataObj.PutInClipboard
End Sub
```

