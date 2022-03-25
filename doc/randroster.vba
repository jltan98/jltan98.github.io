Sub GenerateRoster()
    Dim fulfil As Integer

    fulfil = 1 * 30 'Static for now

    If Worksheets("Roster").Range("J3") <> fulfil Then
        MsgBox "Please check the availability of staffs in Leave Planner!", , "Insufficient PSAs on Duty!"
        Exit Sub
    End If

 Call ScheduleRoster
    
End Sub
Sub ScheduleRoster()
        
    Dim ArrayVal As ArrayList 'Types of Job
    Dim i As Integer 'Random Num
    Dim c As Integer 'Last row
    Dim rownum1 As Integer
    Dim colnum1 As Integer
    Dim cnt As Integer
    Dim items As Integer
    Dim wks As Worksheet
    
    ' Initialize variable.
    Set ArrayVal = New ArrayList
    ArrayVal.Add "CS"
    ArrayVal.Add "Reg"
    ArrayVal.Add "P"
    ArrayVal.Add "Ref"
    ArrayVal.Add "A"
        
    Set wks = Sheets("Roster")
    
    c = 26 + wks.Range("F4") 'Last row
    cnt = 0
    colnum1 = 6 'Column F - Day 1 of month
    i = 0
    items = 5
    
    Worksheets("Roster").Range("F27:AI41").ClearContents

    Range("F27").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC2,StaffAvailability!R14C3:R28C3,StaffAvailability!R14C:R28C) = 1,1,0)"
    Selection.AutoFill Destination:=Range("F27:F41")
    Range("F27:F41").AutoFill Destination:=Range("F27:AI41"), Type:=xlFillDefault
    
    ' Loop Min Required Staff Roles
    Do Until colnum1 = 30 + 6
        
        If wks.Cells(25, colnum1).Text = "Sat" Then
            ' Saturday
            rownum1 = 15
            items = 0
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("F8")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
                
            rownum1 = rownum1 + 1
            items = items + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("F9")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            rownum1 = rownum1 + 1
            items = items + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("F10")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            rownum1 = rownum1 + 1
            items = items + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("F11")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            rownum1 = rownum1 + 1
            items = items + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("F12")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            Do While wks.Cells(14, colnum1) > 0
                
                Do While wks.Cells(22, colnum1) < wks.Range("J5")
                    i = 27
                    cnt = Int(Rnd() * (4 - 0 + 1) + 0) 'Random generate 0 to 4
                    wks.Cells(26 + Int(Application.WorksheetFunction.Match(1, wks.Range(Cells(i, colnum1), Cells(c, colnum1)), 0)), colnum1) = ArrayVal(cnt)
                    i = i + 1
                Loop
                
                i = 27
                wks.Cells(26 + Int(WorksheetFunction.Match(1, wks.Range(Cells(i, colnum1), Cells(c, colnum1)), 0)), colnum1) = "Standby"
                i = i + 1
            Loop
            
            colnum1 = colnum1 + 1
            
        ElseIf wks.Cells(25, colnum1).Text <> "Sun" Then
            ' Not Sunday
            rownum1 = 15
            items = 0
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("E8")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            items = items + 1
            rownum1 = rownum1 + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("E9")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            items = items + 1
            rownum1 = rownum1 + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("E10")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            items = items + 1
            rownum1 = rownum1 + 1
            
           Do Until wks.Cells(rownum1, colnum1) = wks.Range("E11")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            items = items + 1
            rownum1 = rownum1 + 1
            
            Do Until wks.Cells(rownum1, colnum1) = wks.Range("E12")
                i = Int(Rnd() * (c - 27 + 1) + 27) 'Random generate row number
                If wks.Cells(i, colnum1) = 1 Then
                    wks.Cells(i, colnum1) = ArrayVal(items)
                End If
            Loop
            
            Do While wks.Cells(14, colnum1) > 0
                i = 27
                cnt = Int(Rnd() * (4 - 0 + 1) + 0) 'Random generate 0 to 4
                wks.Cells(26 + Int(Application.WorksheetFunction.Match(1, wks.Range(Cells(i, colnum1), Cells(c, colnum1)), 0)), colnum1) = ArrayVal(cnt)
                i = i + 1
            Loop
            
            colnum1 = colnum1 + 1
        Else
            colnum1 = colnum1 + 1
        End If
    Loop
    
End Sub


