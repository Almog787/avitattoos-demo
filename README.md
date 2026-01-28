Sub ReplaceZWithXY()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' הגדרת הגיליון הפעיל
    Set ws = ActiveSheet
    
    ' מציאת השורה האחרונה בעמודה AS (כדי לדעת מתי לעצור)
    lastRow = ws.Cells(ws.Rows.Count, "AS").End(xlUp).Row
    
    ' כיבוי עדכון מסך כדי שהריצה תהיה מהירה יותר
    Application.ScreenUpdating = False
    
    ' לולאה הרצה משורה 2 ועד השורה האחרונה
    For i = 2 To lastRow
        
        ' ---------------------------------------------------------
        ' בדיקה עבור עמודה AS מול עמודה X
        ' ---------------------------------------------------------
        ' הפונקציה InStr בודקת אם האות z קיימת בתוך התא (לא משנה אם גדולה או קטנה)
        If InStr(1, ws.Cells(i, "AS").Value, "z", vbTextCompare) > 0 Then
            ' אם נמצא z, העתק את הערך מעמודה X (עמודה 24) לעמודה AS
            ws.Cells(i, "AS").Value = ws.Cells(i, "X").Value
        End If
        
        ' ---------------------------------------------------------
        ' בדיקה עבור עמודה AT מול עמודה Y
        ' ---------------------------------------------------------
        If InStr(1, ws.Cells(i, "AT").Value, "z", vbTextCompare) > 0 Then
            ' אם נמצא z, העתק את הערך מעמודה Y (עמודה 25) לעמודה AT
            ws.Cells(i, "AT").Value = ws.Cells(i, "Y").Value
        End If
        
    Next i
    
    ' החזרת עדכון מסך והודעת סיום
    Application.ScreenUpdating = True
    MsgBox "תהליך ההחלפה הסתיים בהצלחה!", vbInformation
End Sub
