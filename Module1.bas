' change sample

Attribute VB_Name = "Module1"
Private Sub Distribution_BM()


Dim rngSU As Range
Dim rngS As Range


Application.ScreenUpdating = False



If Date > DateSerial(2019, 8, 1) Then
    
        MsgBox "사용 가능 날짜가 지났습니다." & vbCr & "사용 가능기한 : 2019.10.01"
        
    ThisWorkbook.ActiveSheet.Protect "whwhdrms1!"
        
    Exit Sub
        
End If



For i = 1 To 6

Sheets(i).Unprotect ("whwhdrms1!")


    For Each rngS In Range(Sheets(i).Cells(5, 1), Sheets(i).Cells(Sheets(i).Columns(1).Find("LL", , , 1).Row, 1))
    
        If rngS.Interior.ColorIndex = -4142 Then
        
            If rngSU Is Nothing Then
            
                Set rngSU = rngS
                
            Else
        
                Set rngSU = Union(rngSU, rngS)
                                
            End If
            
        End If
         
    
            
    Next rngS
      
    If Not rngSU Is Nothing Then
    
        rngSU.EntireRow.Delete (xlUp)
        
    End If

    Set rngSU = Nothing
    
    Range(Sheets(i).Cells(5, 12), Sheets(i).Cells(Sheets(i).Columns(1).Find("LL", , , 1).Row, 12)).Value = ""
    Range(Sheets(i).Cells(5, 12), Sheets(i).Cells(Sheets(i).Columns(1).Find("LL", , , 1).Row, 12)).FormatConditions.Delete
    Range(Sheets(i).Cells(5, 12), Sheets(i).Cells(Sheets(i).Columns(1).Find("LL", , , 1).Row, 12)).Interior.ColorIndex = -4142
    

Next i
    


ThisWorkbook.ActiveSheet.Unprotect "whwhdrms1!"

Range(Cells(5, 57), Cells(300, 57)).Interior.Color = xlNone

    For Each rngA In Range(Cells(5, 57), Cells(300, 57))
    
        If rngA <> 0 Then
        
            If Cells(rngA.Row, 1).Value = "B" Then
                
                    If rngA > 0 Then
                    
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet5.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If
                                
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                        
                            If rngB <> 0 Then
                                                    
                                Set FIN = Sheet5.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                               
                                FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                
                                FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                Sheet5.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                                                        
                                FIN.Offset(1, 10).Value = rngB.Value
                                
                                
                            End If
                            
                        Next rngB
                        
                    ElseIf rngA < 0 Then
                        
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet6.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If

                        
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                            
                            If rngB <> 0 Then
                                
                                Set FIN = Sheet6.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                               
                                FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                    
                                FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                Sheet6.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                    
                                FIN.Offset(1, 10).Value = rngB.Value


                            End If
                            
                        Next rngB
                        
                    End If
                
                
                
                
                
            ElseIf Cells(rngA.Row, 1).Value = "S" Then
                
                    If rngA > 0 Then
                    
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet1.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If
                                
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                        
                            If rngB <> 0 Then
                                                    
                                Set FIN = Sheet1.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                                If FIN.Offset(, 7) <> 0 Then
                                
                                
                                    If FIN.Offset(1).Value <> 0 Then
                                    
                                        If FIN.Offset(, 7).Value = Cells(4, rngB.Column).Value Then
                                        
                                            FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                        
                                            FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                            
                                            Sheet1.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                        
                                            FIN.Offset(1, 10).Value = rngB.Value
                                        
                                        Else
                                        
                                            rngB.Interior.Color = RGB(255, 0, 0)
                                        
                                        End If
                                        
                                    Else
                                        
                                        Set UNIT_RNG = Range(FIN.Offset(, 7), FIN.End(4).Offset(-1, 7))
                                            
                                        Set UNIT_RNG_FIN = UNIT_RNG.Find(Cells(4, rngB.Column), , , 1)
                                        
                                        If UNIT_RNG_FIN Is Nothing Then
                                        
                                           Set UNIT_LS = UNIT_RNG.Find("LS Negative", , , 1)
                                            
                                           If UNIT_LS Is Nothing Then
                                           
                                                GoTo NEXT_rngSB
                                                
                                            Else
                                            
                                                UNIT_LS.Offset(1).EntireRow.Insert Shift:=xlDown
                                                
                                                UNIT_LS.Offset(1, -1).Value = Cells(3, rngB.Column).Value
                                                
                                                Sheet1.Rows(UNIT_LS.Row + 1).Interior.Pattern = xlNone
                                                
                                                UNIT_LS.Offset(1, 3).Value = rngB.Value
                                            
                                            End If
                                        
                                        Else
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1).EntireRow.Insert Shift:=xlDown
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1, -1).Value = Cells(3, rngB.Column).Value
                                            
                                            Sheet1.Rows(UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Row + 1).Interior.Pattern = xlNone
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1, 3).Value = rngB.Value
                                            
                                        End If
                                    
                                    End If
                                
                                
                                ElseIf FIN.Offset(1, 4).Value <> 0 Then
                                
                                    FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet1.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 10).Value = rngB.Value
                                    
                                Else
        
                                    FIN.Offset(1, 4).End(4).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 2).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet1.Rows(FIN.Offset(1, 4).End(4).Row - 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 6).Value = rngB.Value
                                    
                                End If
                                
                            End If
                            
NEXT_rngSB:
                         
On Error GoTo 0
                         
                        Next rngB
                        
                    ElseIf rngA < 0 Then
                        
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet2.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If

                        
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                            
                            If rngB <> 0 Then
                                
                                Set FIN = Sheet2.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                                If FIN.Offset(1, 4).Value <> 0 Then
                                
                                    FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet2.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 10).Value = rngB.Value
                                    
                                Else
        
                                    FIN.Offset(1, 4).End(4).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 2).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet2.Rows(FIN.Offset(1, 4).End(4).Row - 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 6).Value = rngB.Value
                                    
                                End If
                            
                            End If
                            
                        Next rngB
                        
                    End If
                                
                
                
                
                
                
            ElseIf Cells(rngA.Row, 1).Value = "P" Then
                
                
                    If rngA > 0 Then
                    
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet3.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If
                                
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                        
                            If rngB <> 0 Then
                                                    
                                Set FIN = Sheet3.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                                If FIN.Offset(, 7) <> 0 Then
                                
                                
                                    If FIN.Offset(1).Value <> 0 Then
                                    
                                        If FIN.Offset(, 7).Value = Cells(4, rngB.Column).Value Then
                                        
                                            FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                        
                                            FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                            
                                            Sheet3.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                        
                                            FIN.Offset(1, 10).Value = rngB.Value
                                        
                                        Else
                                        
                                            rngB.Interior.Color = RGB(255, 0, 0)
                                        
                                        End If
                                        
                                    Else
                                        
                                        Set UNIT_RNG = Range(FIN.Offset(, 7), FIN.End(4).Offset(-1, 7))
                                            
                                        Set UNIT_RNG_FIN = UNIT_RNG.Find(Cells(4, rngB.Column), , , 1)
                                        
                                        If UNIT_RNG_FIN Is Nothing Then
                                        
                                           Set UNIT_LS = UNIT_RNG.Find("LS Negative", , , 1)
                                            
                                           If UNIT_LS Is Nothing Then
                                           
                                                GoTo NEXT_rngPB
                                                
                                            Else
                                            
                                                UNIT_LS.Offset(1).EntireRow.Insert Shift:=xlDown
                                                
                                                UNIT_LS.Offset(1, -1).Value = Cells(3, rngB.Column).Value
                                                
                                                Sheet3.Rows(UNIT_LS.Row + 1).Interior.Pattern = xlNone
                                                
                                                UNIT_LS.Offset(1, 3).Value = rngB.Value
                                            
                                            End If
                                        
                                        Else
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1).EntireRow.Insert Shift:=xlDown
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1, -1).Value = Cells(3, rngB.Column).Value
                                            
                                            Sheet3.Rows(UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Row + 1).Interior.Pattern = xlNone
                                            
                                            UNIT_RNG.Find(Cells(4, rngB.Column), , , 1).Offset(1, 3).Value = rngB.Value
                                            
                                        End If
                                    
                                    End If
                                
                                
                                ElseIf FIN.Offset(1, 4).Value <> 0 Then
                                
                                    FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet3.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 10).Value = rngB.Value
                                    
                                Else
        
                                    FIN.Offset(1, 4).End(4).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 2).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet3.Rows(FIN.Offset(1, 4).End(4).Row - 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 6).Value = rngB.Value
                                    
                                End If
                                
                            End If
                            
NEXT_rngPB:
                         
On Error GoTo 0
                         
                        Next rngB
                        
                    ElseIf rngA < 0 Then
                        
                        On Error Resume Next
                        
                            Cells(rngA.Row, 58).Value = Sheet4.Columns(1).Find(Cells(rngA.Row, 2).Value, , , 1).Row
                            
                        If Err > 0 Then
                        
                            Cells(rngA.Row, 58).Interior.Color = RGB(255, 0, 0)
                            GoTo NOCODEA
                            
                        End If

                        
                        For Each rngB In Union(Range(Cells(rngA.Row, 6), Cells(rngA.Row, 29)), Range(Cells(rngA.Row, 31), Cells(rngA.Row, 55)))
                            
                            If rngB <> 0 Then
                                
                                Set FIN = Sheet4.Columns(1).Find(Cells(rngA.Row, 2), , , 1)
                                
                                If FIN.Offset(1, 4).Value <> 0 Then
                                
                                    FIN.Offset(1).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 6).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet4.Rows(FIN.Row + 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 10).Value = rngB.Value
                                    
                                Else
        
                                    FIN.Offset(1, 4).End(4).EntireRow.Insert Shift:=xlDown
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 2).Value = Cells(3, rngB.Column).Value
                                    
                                    Sheet4.Rows(FIN.Offset(1, 4).End(4).Row - 1).Interior.Pattern = xlNone
                                    
                                    FIN.Offset(1, 4).End(4).Offset(-1, 6).Value = rngB.Value
                                    
                                End If
                            
                            End If
                            
                        Next rngB
                        
                    End If
                
                
                
                
                
                
            End If

        End If

NOCODEA:

On Error GoTo 0

rngA.Interior.Color = RGB(128, 128, 128)

    Next rngA






For i = 1 To 6

    Set RD = Sheets(i).Cells(5, 7).End(4)
    
    Set RDST = RD
    
        Do Until RDST.Row = Sheets(i).Columns(1).Find("LL", , , 1).Row
        
        
            If Not RDST.Row = Sheets(i).Columns(1).Find("LL", , , 1).Row Then
    
        
                Set RDST = RD
                
                If RDST.Offset(1) = "" Or RDST.Offset(1) = 0 Then
                
                    Set RD = RDST.End(4)
                    
                    Sheets(i).Cells(RDST.Row, 12) = RDST.Offset(, 4).Value
                
                Else
                
                    Set RDEN = RD.End(4)
                            
                    Set RD = RDEN.End(4)
                    
                    Sheets(i).Cells(RDST.Row, 12) = Application.Sum(Range(RDST.Offset(, 4), RDEN.Offset(, 4)))
                    
                End If
                
                   
                Sheets(i).Cells(RDST.Row, 12).Font.Underline = xlUnderlineStyleSingle
                    
                Sheets(i).Cells(RDST.Row, 12).Offset(-1).Value = Sheets(i).Cells(RDST.Row, 12).Offset(-1, -1).Value - Sheets(i).Cells(RDST.Row, 12)
                
                
                If Sheets(i).Cells(RDST.Row, 12).Offset(-1).Value = 0 Then
                
                    Sheets(i).Cells(RDST.Row, 12).Offset(-1).Interior.Color = RGB(0, 255, 0)
                
                Else
                
                    Sheets(i).Cells(RDST.Row, 12).Offset(-1).Interior.Color = RGB(255, 0, 0)
                    
                End If
                
                
                With Sheets(i).Cells(RDST.Row, 12).Offset(-1)
                    .FormatConditions.AddIconSetCondition
                    .FormatConditions(.FormatConditions.Count).SetFirstPriority
                    With .FormatConditions(1)
                            .ReverseOrder = False
                            .ShowIconOnly = False
                            .IconSet = ActiveWorkbook.IconSets(xl3Triangles)
                    End With
                    With .FormatConditions(1).IconCriteria(2)
                            .Type = xlConditionValueNumber
                            .Value = 0
                            .Operator = 7
                    End With
                    With .FormatConditions(1).IconCriteria(3)
                            .Type = xlConditionValueNumber
                            .Value = 0
                            .Operator = 5
                    End With
                        
                End With
                
            End If
        
        
        Loop
        
        
    For Each rngN In Range(Sheets(i).Cells(5, 9), Sheets(i).Cells(Sheets(i).Columns(1).Find("LL", , , 1).Row, 9))
    
        If rngN.Interior.Color = RGB(252, 228, 214) And rngN <> 0 And rngN.Offset(, 3) = "" Then
            
                rngN.Offset(, 3).Interior.Color = RGB(255, 128, 128)
                
                rngN.Offset(, 3).Value = "?"
            
        End If
        
    Next rngN
    
    
    
Sheets(i).Protect ("whwhdrms1!")

Next i

ThisWorkbook.ActiveSheet.Protect "whwhdrms1!"

Application.ScreenUpdating = True



End Sub









