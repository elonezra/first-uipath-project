'split the text to rows
Dim rows As String() = copiedTable.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

' create matrix
Dim data(rows.Length-1)() As String

' פיצול כל שורה לעמודות
For i As Integer = 0 To rows.Length-1
    data(i) = rows(i).Split(vbTab)
Next