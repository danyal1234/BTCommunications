 Option Explicit
 Sub Update()
    ActiveWorkbook.Sheets("SharePointLink").Activate
    ActiveSheet.ListObjects("Communications_Booking_Requests").Refresh
    
    ActiveWorkbook.Sheets("ORDERS").Activate
 
    If ActiveWorkbook.Worksheets("ORDERS").Cells(8, 3) <> "" Then
        Worksheets("ORDERS").Rows(8 & ":" & Worksheets("ORDERS").Rows.Count).Delete
    End If
 
    Dim i As Long
    Dim sh As Worksheet
    Dim rn As Range
    Dim k As Long
    Dim counter As Long
    
    Set sh = ThisWorkbook.Sheets("SharePointLink")
    Set rn = sh.UsedRange
    k = rn.Rows.Count + rn.Row - 1
    counter = 2
    
    For i = 2 To k
        If ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 7).Value = 0 Then
        ElseIf ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value <> "Approved As Is" And ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value <> "Approved with Changes" Then
        Else
            'Event Name
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 1).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 4).Value
            'Order Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 2).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 17).Value
            'Asset
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 3).Value = "Canada Flag"
            'Qty
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 4).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 7).Value
            'Rent Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 5).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 5).Value
            'Return Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter + 6, 6).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 6).Value
            counter = counter + 1
        End If
    Next i
    
    For i = 2 To k
        If ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 8).Value <> "" And ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 8).Value <> 0 And (ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved As Is" Or ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved with Changes") Then
            Set sh = ThisWorkbook.Sheets("ORDERS")
            Set rn = sh.UsedRange
            counter = rn.Rows.Count + rn.Row
            'Event Name
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 1).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 4).Value
            'Order Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 2).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 17).Value
            'Asset
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 3).Value = "Ontario Flag"
            'Qty
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 4).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 8).Value
            'Rent Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 5).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 5).Value
            'Return Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 6).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 6).Value
        End If
    Next i
    
    For i = 2 To k
        If ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 9).Value <> "" And ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 9).Value <> 0 And (ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved As Is" Or ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved with Changes") Then
            Set sh = ThisWorkbook.Sheets("ORDERS")
            Set rn = sh.UsedRange
            counter = rn.Rows.Count + rn.Row
            'Event Name
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 1).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 4).Value
            'Order Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 2).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 17).Value
            'Asset
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 3).Value = "Blue Podium"
            'Qty
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 4).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 9).Value
            'Rent Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 5).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 5).Value
            'Return Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 6).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 6).Value
        End If
    Next i
    
    For i = 2 To k
        If ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 10).Value <> "" And ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 10).Value <> 0 And (ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved As Is" Or ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 11).Value = "Approved with Changes") Then
            Set sh = ThisWorkbook.Sheets("ORDERS")
            Set rn = sh.UsedRange
            counter = rn.Rows.Count + rn.Row
            'Event Name
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 1).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 4).Value
            'Order Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 2).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 17).Value
            'Asset
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 3).Value = "Grey Podium"
            'Qty
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 4).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 10).Value
            'Rent Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 5).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 5).Value
            'Return Date
            ActiveWorkbook.Worksheets("ORDERS").Cells(counter, 6).Value = ActiveWorkbook.Worksheets("SharePointLink").Cells(i, 6).Value
        End If
    Next i
    
    Set sh = ThisWorkbook.Sheets("ORDERS")
    Set rn = sh.UsedRange
    k = rn.Rows.Count + rn.Row - 1
    
    For i = 8 To k
        If Cells(i, 6).Value < Date Then Cells(i, 6).Font.Color = vbRed
    Next i

End Sub

