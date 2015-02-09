
Private Sub ExpansionSetBox_Change()
    Dim i As Long
    Dim value, selection As String
    selection = ExpansionSetBox.value
    i = 0
    Do
        value = cells(7 + i, 8).value
        cells(7 + i, 4).Select
        i = i + 1
    Loop Until value = selection
End Sub

Private Sub ResetButton_Click()
    If VerifyPassword = False Then
        MsgBox ("Invalid password")
        Exit Sub
    End If
    Call ResetWorkSheet
End Sub

Private Sub RunReportButton_Click()
    
    If VerifyPassword = False Then
        MsgBox ("Invalid password")
        Exit Sub
    End If

    Dim _
        price, payOutCredit, payOutCash As Double, _
        q1, q2, quantity, unitOfMeasure, factor, i, progressCounter As Long, _
        value As String
        
    i = 0                   'i is our loop counter
    progressCounter = 0     'displays number of records processed in status bar
    
    With ExpansionSetBox
        .Clear
        .value = ""
    End With
    
    Do
        'reset variables
        price = 0
        factor = 0
        'check for empty row
        value = cells(7 + i, 4).value
    
        If value <> "" Then 'Execute the following until rows are empty
            'Check expansion set, add to DDL if necessary
            Dim x As Long, bool As Boolean
            bool = False
            For x = 0 To ExpansionSetBox.ListCount - 1
                If ExpansionSetBox.ListCount > 0 Then
                    If ExpansionSetBox.List(x) = cells(7 + i, 8).value Then bool = True
                End If
            Next
            If Not bool Then ExpansionSetBox.AddItem (cells(7 + i, 8))
            
            'get quantity
            
            'If two rows for a card do the following:
            If cells(7 + i, 4).value = cells(7 + i + 1, 4).value Then
                If cells(7 + i, 14).value < 0 Then
                    q1 = 0
                Else
                    q1 = cells(7 + i, 14).value
                End If
                If cells(7 + i + 1, 14).value < 0 Then
                    q2 = 0
                Else
                    q2 = cells(7 + i + 1, 14).value
                End If
                quantity = q1 + q2
            Else
            'If only one row:
                quantity = cells(7 + i, 14).value
            End If
            
            'get unitOfMeasure
            unitOfMeasure = cells(7 + i, 2).value
            
            'get price
            'If two rows for a card do the following:
            If cells(7 + i, 4).value = cells(7 + i + 1, 4).value Then
                If cells(7 + i, 6).value = "NM" Then
                    price = cells(7 + i, 16).value
                    Rows(7 + i + 1).Delete
                Else
                    price = cells(7 + i + 1, 16).value
                    Rows(7 + i).Delete
                End If
            Else
            'If only one row:
                price = cells(7 + i, 16).value
            End If
            
            'get factor based on unitOfMeasure
            '**************************************************************
            Select Case unitOfMeasure
            '***UoM # 1
                Case 1
                    If quantity <= 8 Then
                        factor = 0.7
                    ElseIf quantity <= 16 Then
                        factor = 0.5
                    ElseIf quantity >= 17 Then
                        factor = 0.25
                    End If
                
                '***UoM # 2
                Case 2
                    If quantity <= 4 Then
                        factor = 0.7
                    ElseIf quantity <= 8 Then
                        factor = 0.5
                    ElseIf quantity <= 16 Then
                        factor = 0.33
                    ElseIf quantity <= 20 Then
                        factor = 0.25
                    ElseIf quantity >= 21 Then
                        factor = 0
                    End If
                
                '***UoM # 3
                Case 3
                    If quantity <= 4 Then
                        factor = 0.5
                    ElseIf quantity <= 8 Then
                        factor = 0.33
                    ElseIf quantity <= 12 Then
                        factor = 0.25
                    ElseIf quantity >= 13 Then
                        factor = 0
                    End If
                
                '***UoM # 4
                Case 4
                    If quantity <= 4 Then
                        factor = 0.5
                    ElseIf quantity <= 8 Then
                        factor = 0.25
                    ElseIf quantity >= 9 Then
                        factor = 0
                    End If
            End Select
            '**************************************************************
            'Calculate and display payOut
            payOutCredit = price * factor
            payOutCash = payOutCredit * 0.6
            If payOutCredit = 0 Then
                Rows(7 + i).Delete
            Else
                cells(7 + i, 18).value = payOutCredit
                cells(7 + i, 20).value = payOutCash
                
                'increment counter
                i = i + 1
            End If
        End If
        progressCounter = progressCounter + 1
        Application.StatusBar = "Progress: Processed " & progressCounter & " records."
    Loop Until value = ""
    
    'Save ExpansionSetBox list values for later
    Dim y, z As Long
    z = 1
    If ExpansionSetBox.ListCount > 0 Then
        For y = 0 To ExpansionSetBox.ListCount - 1
             Sheet2.cells(1 + z, 1) = ExpansionSetBox.List(y)
             z = z + 1
        Next
    End If

    
    'Reset Worksheet formatting
    Call ReformatWorksheet
    'ActiveWindow.Zoom = 200
    
    'Hide quantity and price columns ***added blank column A, UoM and attribute columns--AFH 6.11.13
    'Columns("N:P").EntireColumn.Hidden = True
    'Columns("A:B").EntireColumn.Hidden = True
    'Columns("F").EntireColumn.Hidden = True
    
    MsgBox ("Report Complete")
    Application.StatusBar = False

End Sub

Private Sub Worksheet_Activate()

'Call ResetWorkSheet

End Sub

Private Sub ResetWorkSheet()

    Call ReformatWorksheet
    
    With cells(7, 1)
        .Select
        .value = "(Paste Data Here)"
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
        .Font.Bold = True
        .Font.Name = "Arial"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    
    'Unhide quantity and price columns ***added blank column A, UoM and attribute columns--AFH 6.11.13
    Columns("N:P").EntireColumn.Hidden = False
    Columns("A:B").EntireColumn.Hidden = False
    Columns("F").EntireColumn.Hidden = False
    
    With Columns("B:T")
        .Rows("7:" & .Rows.Count).ClearContents
    End With
    
    ActiveWindow.Zoom = 100
    
End Sub

Public Sub ReformatWorksheet()
    Columns(1).Interior.Color = RGB(220, 220, 220)
    Columns(3).Interior.Color = RGB(220, 220, 220)
    Columns(5).Interior.Color = RGB(220, 220, 220)
    Columns(7).Interior.Color = RGB(220, 220, 220)
    Columns(9).Interior.Color = RGB(220, 220, 220)
    Columns(11).Interior.Color = RGB(220, 220, 220)
    Columns(13).Interior.Color = RGB(220, 220, 220)
    Columns(15).Interior.Color = RGB(220, 220, 220)
    Columns(17).Interior.Color = RGB(220, 220, 220)
    Columns(19).Interior.Color = RGB(220, 220, 220)
    Columns(21).Interior.Color = RGB(220, 220, 220)
    Columns("U:AB").Interior.Color = RGB(220, 220, 220)
    Rows("1:5").Interior.Color = RGB(220, 220, 220)
    Rows(5).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B6:T6").Interior.Color = RGB(256, 256, 256)
    Range("B6:T6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    cells(6, 2).Borders(xlEdgeLeft).LineStyle = xlContinuous
    cells(6, 20).Borders(xlEdgeRight).LineStyle = xlContinuous
End Sub

Private Function VerifyPassword() As Boolean
    PasswordUC.Show
    Dim pwd As String
    Do
        'Wait
    Loop Until PasswordUC.Visible = False
    pwd = PasswordUC.PasswordTextBox.Text
    If pwd = "Welcome1!" Then
        VerifyPassword = True
    Else
        VerifyPassword = False
    End If
End Function

