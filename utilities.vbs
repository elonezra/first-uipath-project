sub WriteToCell(sheetName As String, cell As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Range(cell).Value = FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 4)
End Sub


Function readRow(sheetName As String, index As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim fname As String, lname As String, email As String, phone As String, google As String

    fname = ws.Range("A"+index).Value
    lname = ws.Range("B"+index).Value
    email = ws.Range("C"+index).Value
    phone = ws.Range("D"+index).Value
    google = ws.Range("E"+index).Value
    Dim arr As Variant

    If IsEmpty(google) Then google = "Non"

    arr = Array(fname, lname,  email, phone, google)
    readRow = Join(arr, ",")
End Function

' copiedTable.Split({ vbCrLf }, StringSplitOptions.RemoveEmptyEntries). 
' 		Select(Function(line) line.Split({"	"c}, StringSplitOptions.None).
' 			Select(Function(word) If(String.IsNullOrWhiteSpace(word), "Non", word)).ToArray()).ToArray()