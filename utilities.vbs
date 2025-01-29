sub WriteToCell(sheetName As String, cell As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Range(cell).Value = FormatDateTime(Now, 2) & " " & FormatDateTime(Now, 4)
End Sub


Function readRow(sheetName As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    Dim lastRow As Integer
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' מוצאים את השורה האחרונה עם נתונים
    
    Dim i As Integer
    Dim fname As String, lname As String, email As String, phone As String, google As String
    Dim rowValues As String
    Dim result As String
    
    result = "" ' אתחול מחרוזת התוצאה

    For i = 1 To lastRow ' לולאה על כל השורות הקיימות
        fname = ws.Range("A" & i).Value
        lname = ws.Range("B" & i).Value
        email = ws.Range("C" & i).Value
        phone = ws.Range("D" & i).Value
        google = ws.Range("E" & i).Value

        ' בדיקה אם התא ריק
        If IsEmpty(google) Then google = "Non"

        ' חיבור כל הערכים של השורה למחרוזת אחת
        rowValues = Join(Array(fname, lname, email, phone, google), ",")

        ' הוספת השורה לתוצאה עם הפרדה בטאב בין שורות
        If result = "" Then
            result = rowValues
        Else
            result = result & vbTab & rowValues
        End If
    Next i
    
    readRow = result 
End Function

' copiedTable.Split({ vbCrLf }, StringSplitOptions.RemoveEmptyEntries). 
' 		Select(Function(line) line.Split({"	"c}, StringSplitOptions.None).
' 			Select(Function(word) If(String.IsNullOrWhiteSpace(word), "Non", word)).ToArray()).ToArray()