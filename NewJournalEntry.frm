VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewJournalEntry 
   Caption         =   "New Journal Entry"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   OleObjectBlob   =   "NewJournalEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

JEEntryDateBox.Value = Format(Date, "mm/dd/yyyy")

End Sub

Private Sub JECheckedCheckBox_Click()
    JECheckedDateBox.Enabled = JECheckedCheckBox
    If JECheckedCheckBox.Value = True Then
        JECheckedDateBox.BackColor = &H80000005
    Else: JECheckedDateBox.BackColor = &H80000011
    End If
    
End Sub

Private Sub JEReturnCheckBox_Click()
    JEReturnDateBox.Enabled = JEReturnCheckBox
    If JEReturnCheckBox.Value = True Then
        JEReturnDateBox.BackColor = &H80000005
    Else: JEReturnDateBox.BackColor = &H80000011
    End If

End Sub
Private Sub JECompletedCheckBox_Click()
    JECompletedDateBox.Enabled = JECompletedCheckBox
    If JECompletedCheckBox.Value = True Then
        JECompletedDateBox.BackColor = &H80000005
    Else: JECompletedDateBox.BackColor = &H80000011
    End If

End Sub


Private Sub JEScanCheckBox_Click()
    JEScanDateBox.Enabled = JEScanCheckBox
    If JEScanCheckBox.Value = True Then
        JEScanDateBox.BackColor = &H80000005
    Else: JEScanDateBox.BackColor = &H80000011
    End If
    
End Sub


Private Sub NewJournalEntryCancel_Click()

Unload Me

End Sub

Private Sub NewJournalEntrySubmit_Click()

'Validates the form
'!!!Consider adding logic which prevents entries that are not Checked AND Returned from being marked Completed
If JEEntryNoBox.Text = vbNullString Then
    MsgBox "You need to enter an entry number for this JE.", vbExclamation, "Missing JE Entry #!"
    JEEntryNoBox.SetFocus
    Exit Sub
    
ElseIf JEJEDatebox = vbNullString Then
    MsgBox "You need to enter a JE date for this JE.", vbExclamation, "Missing JE Date!"
    JEJEDatebox.SetFocus
    Exit Sub
ElseIf IsDate(JEJEDatebox.Value) = False Then
    MsgBox JEJEDatebox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    JEJEDatebox.SetFocus
    Exit Sub
    
ElseIf JEEntryDateBox = vbNullString Then
    MsgBox "You need to enter an entry date for this JE.", vbExclamation, "Missing JE Entry Date!"
    JEEntryDateBox.SetFocus
    Exit Sub
ElseIf IsDate(JEEntryDateBox.Value) = False Then
    MsgBox JEJEDatebox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    JEEntryDateBox.SetFocus
    Exit Sub

'Consider moving/changing this.
ElseIf JECheckedCheckBox = True And JECheckedDateBox = vbNullString Then
    MsgBox "You need to enter a checked date for this JE.", vbExclamation, "Missing Checked Date!"
    JECheckedDateBox.SetFocus
    Exit Sub
ElseIf JECheckedCheckBox = True And IsDate(JECheckedDateBox.Value) = False Then
    MsgBox JEJECheckedBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    JECheckedDateBox.SetFocus
    Exit Sub
    
'Supra
ElseIf JEReturnCheckBox = True And JEReturnDateBox = vbNullString Then
    MsgBox "You need to enter a return date for this JE.", vbExclamation, "Missing Return Date!"
    JEReturnDateBox.SetFocus
    Exit Sub
ElseIf JEReturnCheckBox = True And IsDate(JEReturnDateBox.Value) = False Then
    MsgBox JEReturnDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    JEReturnDateBox.SetFocus
    Exit Sub
    
ElseIf JECompletedCheckBox = True And JECompletedDateBox = vbNullString Then
    MsgBox "You need to enter a completion date for this JE.", vbExclamation, "Missing JE Completion Date!"
    JECompletedDateBox.SetFocus
    Exit Sub
ElseIf JECompletedCheckBox = True And IsDate(JECompletedDateBox.Value) = False Then
    MsgBox JECompletedDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid JE Entry Date!"
    JECompletedDateBox.SetFocus
    Exit Sub
    
ElseIf JEScanCheckBox = True And JEScanDateBox = vbNullString Then
    MsgBox "You need to enter a scan date for this JE.", vbExclamation, "Missing JE Scan Date!"
    JEScanDateBox.SetFocus
    Exit Sub
ElseIf JEScanCheckBox = True And IsDate(JEScanDateBox.Value) = False Then
    MsgBox JEScanDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid JE Scan Date!"
    JEScanDateBox.SetFocus
    Exit Sub
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Creates a new JournalEntry Object
Dim JE As JournalEntry
Set JE = New JournalEntry

'Checks to see if the form data is valid
'   If it is,
'      Loads the JournalEntry Object with form data.
'   Else
'       Raises exceptions.

JE.Entry = JEEntryNoBox.Value
JE.JEDate = JEJEDatebox.Value
JE.KeyDate = JEEntryDateBox.Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Really questionable way of catching JEs missing dates.

If JECheckedDateBox.Value = vbNullString Then
    JE.CheckDate = "12/25/9999"
Else: JE.CheckDate = JECheckedDateBox.Value
End If

If JEReturnDateBox.Value = vbNullString Then
    JE.ReturnDate = "12/25/9999"
Else: JE.ReturnDate = JEReturnDateBox.Value
End If

If JECompletedDateBox.Value = vbNullString Then
    JE.CompletedDate = "12/25/9999"
Else: JE.CompletedDate = JECompletedDateBox.Value
End If

If JEScanDateBox.Value = vbNullString Then
    JE.ScanDate = "12/25/9999"
Else: JE.ScanDate = JEScanDateBox.Value
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Selects the proper KEYLOG sheet
If Month(JE.KeyDate) = 1 Then
    'JE.Prefix = JE.Prefix & "A"
    Worksheets("JANUARY KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 2 Then
    'JE.Prefix = JE.Prefix & "B"
    Worksheets("FEBRUARY KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 3 Then
    'JE.Prefix = JE.Prefix & "H"
   Worksheets("MARCH KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 4 Then
    'JE.Prefix = JE.Prefix & "R"
    Worksheets("APRIL KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 5 Then
    'JE.Prefix = JE.Prefix & "Y"
    Worksheets("MAY KEYLOG").Activate

ElseIf Month(JE.KeyDate) = 6 Then
    'JE.Prefix = JE.Prefix & "E"
    Worksheets("JUNE KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 7 Then
    'JE.Prefix = JE.Prefix & "L"
    Worksheets("JULY KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 8 Then
    'JE.Prefix = JE.Prefix & "G"
    Worksheets("AUGUST KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 9 Then
    'JE.Prefix = JE.Prefix & "P"
    Worksheets("SEPTEMBER KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 10 Then
    'JE.Prefix = JE.Prefix & "T"
    Worksheets("OCTOBER KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 11 Then
    'JE.Prefix = JE.Prefix & "V"
    Worksheets("NOVEMBER KEYLOG").Activate
    
ElseIf Month(JE.KeyDate) = 12 Then
    'JE.Prefix = JE.Prefix & "D"
    Worksheets("DECEMBER KEYLOG").Activate
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Catches unchecked/unreturned/incomplete JEs
If JE.CheckDate = "12/25/9999" Then
    JE.CheckStatus = "NOT CHECKED"
    'Worksheets("PENDING").Activate
    'Range("D1").Value = JE.CheckStatus
Else: JE.CheckStatus = "CHECKED"
    'Range("D1").Value = JE.CheckDate
End If

If JE.ReturnDate = "12/25/9999" Then
    JE.ReturnStatus = "NOT RETURNED"
    'Worksheets("PENDING").Activate
    'Range("E1").Value = JE.ReturnStatus
Else: JE.ReturnStatus = "RETURNED"
    'Range("E1").Value = JE.ReturnDate
End If

If JE.CompletedDate = "12/25/9999" Then
    JE.CompletedStatus = "NOT COMPLETE"
    'Worksheets("PENDING").Activate
    'Range("F1").Value = JE.CompletedStatus
Else: JE.CompletedStatus = "COMPLETE"
    'Range("F1").Value = JE.CompletedDate
End If

If JE.ScanDate = "12/25/9999" Then
    JE.ScanStatus = "NOT SCANNED"
    'Worksheets("PENDING").Activate
    'Range("F1").Value = JE.CompletedStatus
Else: JE.ScanStatus = "SCANNED"
    'Range("F1").Value = JE.CompletedDate
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'!!!If JE is Complete,
'       Load the JE data into the keylog,
'       ???Store the JE in the Keyed_JE_Dictionary
'       ???Store the JE in the Completed_JE_Dictionary.
'
'   Else,
'       Load the JE data into the keylog,
'       ???Load the JE data into the Pending screen,
'       ???Store the JE in the Keyed_JE_Dictionary,
'       ???Store the JE in the Pending_JE_Dictionary.

'Uncomment next two lines to route to Testing Output sheet
'Worksheets("TestingOut").Activate
'crpage = "TestingOut"

If JE.StatusArray(0) = "CHECKED" And JE.StatusArray(1) = "RETURNED" And JE.StatusArray(2) = "COMPLETE" Then
    Dim LastRow As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row + 1
    Range("A" & LastRow).Value = JE.KeyDate
    'Range("B" & LastRow).Value = ???
    Range("C" & LastRow).Value = JE.Entry
    Range("D" & LastRow).Value = JE.CheckDate
    Range("E" & LastRow).Value = JE.ReturnDate
    Range("F" & LastRow).Value = JE.CompletedDate
    Range("N" & LastRow).Value = JE.ScanDate
    
'???Else log the entry in Pending.

Else:
    LastRow = Range("A" & Rows.Count).End(xlUp).Row + 1
    Range("A" & LastRow).Value = JE.KeyDate
    Range("B" & LastRow).Value = JE.Entry
    'Range("C" & LastRow).Value = ???
    If JE.CheckStatus = "NOT CHECKED" Then
        Range("D" & LastRow).Value = JE.CheckStatus
    Else:
        Range("D" & LastRow).Value = JE.CheckDate
    End If
    
    If JE.ReturnStatus = "NOT RETURNED" Then
        Range("E" & LastRow).Value = JE.ReturnStatus
    Else:
        Range("E" & LastRow).Value = JE.ReturnDate
    End If
        
    If JE.CompletedStatus = "NOT COMPLETE" Then
        Range("F" & LastRow).Value = JE.CompletedStatus
    Else:
        Range("F" & LastRow).Value = JE.CompletedDate
    End If
    
    If JE.ScanStatus = "NOT SCANNED" Then
        Range("N" & LastRow).Value = JE.ScanStatus
    Else:
        Range("N" & LastRow).Value = JE.ScanDate
    End If

End If

Unload Me

End Sub

