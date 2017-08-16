Option Explicit
Public PriorVal As String

Private Sub Workbook_Open()
Dim NR As Long
    With Sheets("AuditLog")
        NR = .Range("C" & .Rows.Count).End(xlUp).Row + 1
        Application.EnableEvents = False
        .Range("A" & NR).Value = Environ("UserName")
        .Range("B" & NR).Value = Environ("ComputerName")
        Application.EnableEvents = True
    End With
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    If Selection(1).Value = "" Then
        PriorVal = "Blank"
    Else
        PriorVal = Selection(1).Value
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
Dim NR As Long
If sh.Name = "AuditLog" Then Exit Sub     'allows you to edit the log sheet

Application.EnableEvents = False
    With Sheets("AuditLog")
        NR = .Range("C" & .Rows.Count).End(xlUp).Row + 1
        
        .Range("A1").Value = "Username"
        .Range("B1").Value = "Computer Name"
        .Range("C1").Value = "Date and Time"
        .Range("D1").Value = "Sheet Name"
        .Range("E1").Value = "Cell or Cell Range"
        .Range("F1").Value = "Prior Value"
        .Range("G1").Value = "Current Value"
        .Range("H1").Value = "Audit Trail Reason"
        
        .Range("C" & NR).Value = Now
        .Range("D" & NR).Value = sh.Name
        .Range("E" & NR).Value = Target.Address
        .Range("F" & NR).Value = PriorVal
        .Range("G" & NR).Value = Target(1).Value
        NR = NR + 1
    End With
    Application.EnableEvents = True
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If ActiveSheet.CodeName = "AuditLog" Then Exit Sub 'Allows you to edit log sheet without getting stuck in Reason loop
    Dim NR As Long
    With Sheets("AuditLog")
        NR = .Range("C" & .Rows.Count).End(xlUp).Row 'Get Last Row with content present
    Do While .Range("H" & NR).Value = "" 'While Column H and last row containing content has nothing in column H
        If .Range("H" & NR).Value = "" Then 'If column H and last row containing content has nothing in column H Then
            .Range("H" & NR).Value = InputBox("Please enter an audit reason.", "Audit Trail Reason") 'Present input box to garner an audit trail reason
        End If
        If .Range("H" & NR).Value = "" Then 'If after all that, column H and last row containing content still has nothing in it
            MsgBox "You MUST provide an audit trail reason.", vbCritical, "No Audit Reason Given" 'Present a warning to the user
        End If
    Loop 'Loop until they give an audit trail reason
    End With
End Sub
