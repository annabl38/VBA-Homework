Attribute VB_Name = "Module1"
Sub stockcheck()
For Each ws In Worksheets

Dim row As Long
Dim totalvol As Double
Dim lastrow As Long
Dim summaryrow As Long
Dim perchange As Double
Dim yearstart As Double
Dim yearend As Double
Dim yearchange As Double
Dim greatperchange As Double
Dim lowperchange As Double
Dim greatvol As Double
Dim greatperticker As String, lowperticker As String, greatvolticker As String


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
totalvol = 0
summaryrow = 2
greatperchange = 0
lowperchange = 0
greatvol = 0
yearstart = 0
yearend = 0
yearchange = 0

For row = 2 To lastrow
    
    If ws.Cells(row, 1) <> ws.Cells(row - 1, 1) Then
    yearstart = ws.Cells(row, 3)
    totalvol = ws.Cells(row, 7) + totalvol
        
    ElseIf ws.Cells(row, 1) = ws.Cells(row + 1, 1) Then
     totalvol = ws.Cells(row, 7) + totalvol
     If yearstart = 0 Then
     yearstart = ws.Cells(row, 3)
    End If
    
           
    ElseIf ws.Cells(row, 1) <> ws.Cells(row + 1, 1) Then
        totalvol = ws.Cells(row, 7) + totalvol
        yearend = ws.Cells(row, 6)
       ws.Cells(summaryrow, 9).Value = ws.Cells(row, 1)
        ws.Cells(summaryrow, 10).Value = totalvol
        yearchange = yearend - yearstart
        ws.Cells(summaryrow, 11).Value = yearchange
        
        If yearstart = 0 And yearend = 0 Then
        perchange = 0
        Else
        perchange = yearchange / yearstart
       End If
        
        ws.Cells(summaryrow, 12).Value = perchange
        ws.Cells(summaryrow, 12).NumberFormat = "0.00%"
            If yearchange >= 0 Then
           ws.Cells(summaryrow, 11).Interior.Color = RGB(0, 255, 0)
            Else
            ws.Cells(summaryrow, 11).Interior.Color = RGB(255, 0, 0)
            End If
            
        If perchange > greatperchange Then
            greatperchange = perchange
            greatperticker = ws.Cells(summaryrow, 9).Value
        ElseIf perchange <= lowperchange Then
            lowperchange = perchange
            lowperticker = ws.Cells(summaryrow, 9).Value
        End If
        
        If totalvol > greatvol Then
            greatvol = totalvol
            greatvolticker = ws.Cells(summaryrow, 9).Value
        End If
        
        summaryrow = summaryrow + 1
        totalvol = 0
    End If
Next row

ws.Range("o2").Value = greatperticker
ws.Range("P2").Value = greatperchange
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("O3").Value = lowperticker
ws.Range("P3").Value = lowperchange
ws.Range("p3").NumberFormat = "0.00%"
ws.Range("O4").Value = greatvolticker
ws.Range("P4").Value = greatvol

Next ws




End Sub

