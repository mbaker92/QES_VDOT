Attribute VB_Name = "Code_For_Reformat_Button1"
Option Explicit
Private Sub Auto_Open()
    Dim buttons_objects As Shape
    Dim test_sheet As Worksheet
    Dim ReformatButton As Button
    
        'Remove existing buttons (prevent button accumulation)
        For Each test_sheet In Sheets
            For Each buttons_objects In test_sheet.Shapes
                buttons_objects.Delete
            Next
        Next
        'Make the reformat button and assign it to the appropriate macro
        Set ReformatButton = Sheets("ShoulderComparison").Buttons.Add(Sheets("ShoulderComparison").Range("Z1").Left, Sheets("ShoulderComparison").Range("Z1").Top, Sheets("ShoulderComparison").Range("Z1:AA1").Width, Sheets("ShoulderComparison").Range("Z1:Z2").Height)
        ReformatButton.OnAction = "Reformat_Update_Shoulder"
        ReformatButton.Characters.Text = "Reformat Table"

    
End Sub
Sub Reformat_Update_Shoulder()
    
    SecondPassUpdate
    MsgBox "Reformat Finished", , "Complete"
    
End Sub
Function SheetNameExists(SheetName As String) As Boolean
    Dim test_sheet As Variant
    For Each test_sheet In Sheets
        If test_sheet.Name = SheetName Then
            SheetNameExists = True
            Exit Function
        End If
    Next
    SheetNameExists = False
    
End Function
Function SecondPassUpdate()

    Sheets("ShoulderComparison").Activate
    Dim iter As Long
    Dim lastrow As Long
    lastrow = Sheets("ShoulderComparison").Range("A1048576").End(xlUp).Row
    
    For iter = 2 To lastrow Step 1
        Sheets("ShoulderComparison").Range(CStr("Z" & iter)).Formula = "=Concatenate($A$" & iter & ",$C$" & iter & ")"
    Next

    Sheets("ShoulderComparison").Range("A2:Z" & lastrow).Sort Key1:=Range("B2:B" & lastrow), Order1:=xlAscending, Key2:=Range("C2:C" & lastrow), Order2:=xlDescending, Key3:=Range("G2:G" & lastrow), Order3:=xlAscending

    
    
    Dim Cols(8, 2) As String
    Dim Sets1(26, 1) As String
    
    Cells.FormatConditions.Delete
    
    
    For iter = 2 To lastrow
        Sheets("ShoulderComparison").Range(CStr("A" & iter)) = Sheets("ShoulderComparison").Range(CStr("D" & iter)) & Sheets("ShoulderComparison").Range(CStr("G" & iter)) & Sheets("ShoulderComparison").Range(CStr("H" & iter))
        If Sheets("ShoulderComparison").Range(CStr("Q" & iter)).Value = "Matching FWG data not found" Then
            Sheets("ShoulderComparison").Range(CStr("Q" & iter)).Value = ""
        End If
    Next
    Sheets("ShoulderComparison").Range(CStr("R2:Y" & lastrow)).ClearContents
    
    For iter = 1 To 8
        Cols(iter, 0) = CStr(ColNumToLet(ColLetToNum("Q") + iter))
        Cols(iter, 1) = CStr(ColNumToLet(ColLetToNum("H") + iter))
    Next
    'Pass/Fail threshold
        Cols(1, 2) = CStr(0.84)
        Cols(2, 2) = CStr(0.29)
        Cols(3, 2) = CStr(0.7)
        Cols(4, 2) = CStr(0.48)
        Cols(5, 2) = CStr(0.83)
        Cols(6, 2) = CStr(0.48)
        Cols(7, 2) = CStr(0.84)
        Cols(8, 2) = CStr(0.26)
    
    Dim iter2 As Long
    Dim iter3 As Long
    Dim iter4 As Long
    Dim RateCheck As String
    Dim CheckCheck As Long
    Dim ComTarQES As String
    Dim ComTarFWG As String
    Dim Match1(500, 1) As String
    Dim Match2 As Long
    Dim InputFormula1 As String
    Dim InputFormula2 As String
    Dim ShoulderCompare As Sheets
    Dim ID_range As Range
    
    Set ID_range = Sheets("ShoulderComparison").Range(CStr("Z:Z"))
    
    Dim SRow2 As Long
    
    CheckCheck = 0
    Match2 = 0
    SRow2 = 2
    
    'Iterate through all rows comparing QES rating to FWG rating
    For iter = 2 To lastrow
        If (iter Mod 10) = 0 Then
            Application.StatusBar = "Step 1:  " & Format(iter / lastrow, "0%") & "   Overall:  " & Format(iter / lastrow * 0.4, "0%")
            DoEvents
        End If
        RateCheck = Sheets("ShoulderComparison").Range(CStr("C" & CStr(iter)))
        If RateCheck = "QES" Then
            CheckCheck = 0
            For iter2 = 1 To 8
                If Sheets("ShoulderComparison").Range(CStr("Q" & iter)) = "" Then
                    If ID_range.Find(CStr(Sheets("ShoulderComparison").Range(CStr("A" & iter)) & "FWG"), , xlValues) Is Nothing Then
                        Sheets("ShoulderComparison").Range("Q" & iter).Value = "Matching FWG data not found"
                    Else
                    
                        ComTarFWG = Application.Match(CStr(Sheets("ShoulderComparison").Range(CStr("A" & iter)) & "FWG"), Sheets("ShoulderComparison").Range(CStr("Z1:Z" & lastrow)), 0)
                        ComTarQES = CStr(Sheets("ShoulderComparison").Range(CStr(Cols(iter2, 0) & iter)).Address)
                        If iter2 = 1 Or iter2 = 2 Or iter2 = 5 Or iter2 = 6 Then
                            Sheets("ShoulderComparison").Range(ComTarQES).Formula = CStr("=IF($" & Cols(iter2, 1) & "$" & iter & "=$" & Cols(iter2, 1) & "$" & ComTarFWG & ",1,0)")
                        Else
                            Sheets("ShoulderComparison").Range(ComTarQES).Formula = CStr("=IF(OR(AND(VALUE($" & Cols(iter2, 1) & "$" & iter & ")>=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")-2,VALUE($" & Cols(iter2, 1) & "$" & iter & ")<=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")+2),AND(VALUE($" & Cols(iter2, 1) & "$" & iter & ")=4,VALUE($" & Cols(iter2, 1) & "$" & iter & ")>=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")-4,VALUE($" & Cols(iter2, 1) & "$" & iter & ")<=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")+2),AND(VALUE($" & Cols(iter2, 1) & "$" & iter & ")=8,VALUE($" & Cols(iter2, 1) & "$" & iter & ")>=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")-4,VALUE($" & Cols(iter2, 1) & "$" & iter & ")<=VALUE($" & Cols(iter2, 1) & "$" & ComTarFWG & ")+4)),1,0)")
                        End If
                    End If
                    
                End If
            Next
        End If
        Sheets("ShoulderComparison").Range(CStr("A" & iter)).Select
        'ActiveWindow.SmallScroll down:=0.1
    Next




    Sheets("ShoulderComparison").Activate
    'Iterate through all rows and create percentage agreement and calculate pass/fail conditions
    For iter = 2 To lastrow
        If (iter Mod 10) = 0 Then
            Application.StatusBar = "Step 2:  " & Format(iter / lastrow, "0%") & "   Overall:  " & Format(iter / lastrow * 0.6 + 0.4, "0%")
            DoEvents
        End If
        Sheets("ShoulderComparison").Range(CStr("B" & iter)).Select
        RateCheck = Sheets("ShoulderComparison").Range(CStr("C" & CStr(iter)))
        If RateCheck = "FWG" Then
            CheckCheck = CheckCheck + 1
            If CheckCheck = 2 Then
                For iter2 = 1 To 8
                    For iter3 = 2 To lastrow
                        If Sheets("ShoulderComparison").Range(CStr("C" & iter3)) = "QES" Then
                            If Sheets("ShoulderComparison").Range(CStr("B" & iter)) = Sheets("ShoulderComparison").Range(CStr("B" & iter3)) Then
                                Match1(Match2, 0) = Match2
                                Match1(Match2, 1) = CStr(Sheets("ShoulderComparison").Range(CStr(Cols(iter2, 0) & iter3)).Address)
                                Match2 = Match2 + 1
                            End If
                        End If
                    Next
                    InputFormula1 = Match1(0, 1)
                    InputFormula2 = Match1(0, 1)
                    For iter4 = 1 To Match2 - 1
                        If (Match1(iter4, 1) <> "") Then
                            InputFormula1 = CStr(InputFormula1 & "+" & Match1(iter4, 1))
                            InputFormula2 = CStr(InputFormula2 & "," & Match1(iter4, 1))
                        End If
                    Next
                    Sheets("ShoulderComparison").Range(CStr(Cols(iter2, 0) & iter)).Formula = CStr("=IFERROR(ROUND(SUM(" & InputFormula2 & ")/Count(" & InputFormula2 & "),2),-1)")
                    Erase Match1
                    Match2 = 0
                    
                    'Update ShoulderCompare tab as agreement data is generated
                    Sheets("ShoulderCompare").Range(CStr(ColNumToLet(2 + iter2) & SRow2)).Formula = CStr("=ShoulderComparison!$" & Cols(iter2, 0) & "$" & iter)
                    Worksheets("ShoulderComparison").Activate
                Next
                
                'Update ShoulderCompare tab with sample numbers and agreement line headers
                InputFormula1 = Sheets("ShoulderComparison").Range(CStr("B" & iter))
                Sheets("ShoulderCompare").Range(CStr("A" & SRow2)) = InputFormula1
                Sheets("ShoulderCompare").Range(CStr("B" & SRow2)) = "Agreement"
                SRow2 = SRow2 + 1
            ElseIf CheckCheck = 3 Then
                For iter2 = 1 To 8
                    Sheets("ShoulderComparison").Range(CStr(Cols(iter2, 0) & iter)).Formula = CStr("=IF($" & Cols(iter2, 0) & "$" & (iter - 1) & "<" & Cols(iter2, 2) & ",""FAIL"",""PASS"")")
                    Sheets("ShoulderCompare").Range(CStr(ColNumToLet(2 + iter2) & SRow2)).Formula = CStr("=ShoulderComparison!$" & Cols(iter2, 0) & "$" & iter)
                    Sheets("ShoulderCompare").Range(CStr(ColNumToLet(2 + iter2) & SRow2 + 1)) = ""
                    
                Next
                InputFormula1 = Sheets("ShoulderComparison").Range(CStr("B" & iter))
                Sheets("ShoulderCompare").Range(CStr("A" & SRow2)) = InputFormula1
                Sheets("ShoulderCompare").Range(CStr("B" & SRow2)) = "PASS/FAIL"
                Sheets("ShoulderCompare").Range(CStr("A" & SRow2 + 1)) = ""
                Sheets("ShoulderCompare").Range(CStr("B" & SRow2 + 1)) = ""
                
                SRow2 = SRow2 + 2
                
            End If
        Else
            CheckCheck = 0
        End If
    Next
    Application.StatusBar = False
    With Sheets("ShoulderComparison").Range("$R:$Y")
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="FAIL"
        .FormatConditions(1).Interior.Color = RGB(255, 255, 0)
    End With
    Sheets("ShoulderComparison").Range("A1").Select
    
    EHSFormat2
    
End Function
Function EHSFormat2()
        
    Dim iter As Long
    Dim lastrow As Long
    lastrow = Sheets("ShoulderCompare").Range("A1048576").End(xlUp).Row

    For iter = 2 To lastrow Step 3
        Sheets("ShoulderCompare").Range(CStr("K" & iter & ":K" & iter + 1)).Merge
        Sheets("ShoulderCompare").Range(CStr("A" & iter & ":K" & iter + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Sheets("ShoulderCompare").Range(CStr("A" & iter & ":K" & iter + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Sheets("ShoulderCompare").Range(CStr("A" & iter & ":K" & iter + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Sheets("ShoulderCompare").Range(CStr("A" & iter & ":K" & iter + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
    Next
    
End Function

Public Function ColLetToNum(ColInput As String) As Long
' Convert Column letter to a number
    Dim Leng1 As Long
    Dim Output1 As Long
    
    Leng1 = Len(ColInput)
    
    If Leng1 = 1 Then
        Output1 = Asc(ColInput) - 64
    End If
    If Leng1 = 2 Then
        Output1 = (Asc(Left(ColInput, 1)) - 64) * 26 + Asc(Right(ColInput, 1)) - 64
    End If
    If Leng1 = 3 Then
        Output1 = (Asc(Left(ColInput, 1)) - 64) * 26 * 26 + (Asc(Mid(ColInput, 2, 1)) - 64) * 26 + Asc(Right(ColInput, 1)) - 64
    End If
    
    ColLetToNum = Output1
End Function
Public Function ColNumToLet(ColInput) As String
    ' Convert Column Number to Letter
    Dim Output1 As String
    Output1 = Split(Cells(, ColInput).Address, "$")(1)
    
    ColNumToLet = Output1
End Function

