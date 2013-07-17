VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewCashReceipt 
   Caption         =   "New Cash Receipt"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
   OleObjectBlob   =   "NewCashReceipt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewCashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

'Loads CRPrefixComboBox with the values from the "Config" sheet
'To add/remove values:
'0. Select the "Config" worksheet.
'1. Edit the column in the "CR_Config" section labeled "Prefix List" as needed.
'2. Click the "Formulas" tab.
'3. Click the "Name Manager" button.
'4. Select "PrefixList."
'5. Click the "Delete" button.
'6. Click the "New" button.
'7. Type "PrefixList" (without quotes) in the "Name:" box.
'8. Click in the "Refers to:" box.
'9. Click and drag in the "Config" worksheet, selecting all of the cells in the "Prefix List" column except for the words "Prefix List."
'10. Click OK.
'11. Click OK.
'12. Save the workbook.  CRPrefixComboBox will now contain the new prefixes.

Dim cPre As Range
Dim ws As Worksheet
Set ws = Worksheets("Config")

For Each cPre In ws.Range("PrefixList")
    With Me.CRPrefixComboBox
        .AddItem cPre.Value
    End With
Next cPre

CREntryDateBox.Value = Format(Date, "mm/dd/yyyy")

End Sub


Private Sub CRCheckedCheckBox_Click()
    CRCheckedDateBox.Enabled = CRCheckedCheckBox
    If CRCheckedCheckBox.Value = True Then
        CRCheckedDateBox.BackColor = &H80000005
    Else: CRCheckedDateBox.BackColor = &H80000011
    End If
    
End Sub

Private Sub CRReturnCheckBox_Click()
    CRReturnDateBox.Enabled = CRReturnCheckBox
    If CRReturnCheckBox.Value = True Then
        CRReturnDateBox.BackColor = &H80000005
    Else: CRReturnDateBox.BackColor = &H80000011
    End If

End Sub
Private Sub CRCompletedCheckBox_Click()
    CRCompletedDateBox.Enabled = CRCompletedCheckBox
    If CRCompletedCheckBox.Value = True Then
        CRCompletedDateBox.BackColor = &H80000005
    Else: CRCompletedDateBox.BackColor = &H80000011
    End If

End Sub


Private Sub CRScanCheckBox_Click()
    CRScanDateBox.Enabled = CRScanCheckBox
    If CRScanCheckBox.Value = True Then
        CRScanDateBox.BackColor = &H80000005
    Else: CRScanDateBox.BackColor = &H80000011
    End If
    
End Sub

Private Sub NewCashReceiptCancel_Click()

Unload Me

End Sub

Private Sub NewCashReceiptSubmit_Click()

'Validates the form
'!!!Consider adding logic which prevents entries that are not Checked AND Returned from being marked Completed
If CREntryNoBox.Text = vbNullString Then
    MsgBox "You need to enter an entry number for this CR.", vbExclamation, "Missing CR Entry #!"
    CREntryNoBox.SetFocus
    Exit Sub
    
ElseIf CRAABox.Text = vbNullString Then
    MsgBox "You need to enter an actual amount for this CR.", vbExclamation, "Missing Actual Amount!"
    CRAABox.SetFocus
    Exit Sub
ElseIf IsNumeric(CRAABox.Value) = False Then
    MsgBox CRAABox.Value & " is not a vald amount.  This field only accepts number values.", vbExclamation, "Invalid Date!"
    CRAABox.SetFocus
    Exit Sub
    
ElseIf CRCRDateBox = vbNullString Then
    MsgBox "You need to enter a CR date for this CR.", vbExclamation, "Missing CR Date!"
    CRCRDateBox.SetFocus
    Exit Sub
ElseIf IsDate(CRCRDateBox.Value) = False Then
    MsgBox CRCRDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    CRCRDateBox.SetFocus
    Exit Sub
    
ElseIf CREntryDateBox = vbNullString Then
    MsgBox "You need to enter an entry date for this CR.", vbExclamation, "Missing CR Entry Date!"
    CREntryDateBox.SetFocus
    Exit Sub
ElseIf IsDate(CREntryDateBox.Value) = False Then
    MsgBox CRCRDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    CREntryDateBox.SetFocus
    Exit Sub

'Consider moving/changing this.
ElseIf CRCheckedCheckBox = True And CRCheckedDateBox = vbNullString Then
    MsgBox "You need to enter a checked date for this CR.", vbExclamation, "Missing Checked Date!"
    CRCheckedDateBox.SetFocus
    Exit Sub
ElseIf CRCheckedCheckBox = True And IsDate(CRCheckedDateBox.Value) = False Then
    MsgBox CRCRCheckedBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    CRCheckedDateBox.SetFocus
    Exit Sub
    
'Supra
ElseIf CRReturnCheckBox = True And CRReturnDateBox = vbNullString Then
    MsgBox "You need to enter a return date for this CR.", vbExclamation, "Missing Return Date!"
    CRReturnDateBox.SetFocus
    Exit Sub
ElseIf CRReturnCheckBox = True And IsDate(CRReturnDateBox.Value) = False Then
    MsgBox CRReturnDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid Date!"
    CRReturnDateBox.SetFocus
    Exit Sub
    
ElseIf CRCompletedCheckBox = True And CRCompletedDateBox = vbNullString Then
    MsgBox "You need to enter a completion date for this CR.", vbExclamation, "Missing CR Completion Date!"
    CRCompletedDateBox.SetFocus
    Exit Sub
ElseIf CRCompletedCheckBox = True And IsDate(CRCompletedDateBox.Value) = False Then
    MsgBox CRCompletedDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid CR Entry Date!"
    CRCompletedDateBox.SetFocus
    Exit Sub
    
ElseIf CRScanCheckBox = True And CRScanDateBox = vbNullString Then
    MsgBox "You need to enter a scan date for this CR.", vbExclamation, "Missing CR Scan Date!"
    CRScanDateBox.SetFocus
    Exit Sub
ElseIf CRScanCheckBox = True And IsDate(CRScanDateBox.Value) = False Then
    MsgBox CRScanDateBox.Value & " is not a vald date.  Try mm/dd/yy formatting.", vbExclamation, "Invalid CR Scan Date!"
    CRScanDateBox.SetFocus
    Exit Sub
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Creates a new CashReceipt Object
Dim CR As CashReceipt
Set CR = New CashReceipt

CR.Prefix = CRPrefixComboBox.Value
CR.Entry = CREntryNoBox.Value
CR.AA = CRAABox.Value
CR.CRDate = CRCRDateBox.Value
CR.KeyDate = CREntryDateBox.Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Really questionable way of catching CRs missing dates.  Will have to catch
'Later with something like
'If CR.Date = "12/25/9999" Then
'   Date = "MISSING"
'   Puts(Date)
'End If

If CRCheckedDateBox.Value = vbNullString Then
    CR.CheckDate = "12/25/9999"
Else: CR.CheckDate = CRCheckedDateBox.Value
End If

If CRReturnDateBox.Value = vbNullString Then
    CR.ReturnDate = "12/25/9999"
Else: CR.ReturnDate = CRReturnDateBox.Value
End If

If CRCompletedDateBox.Value = vbNullString Then
    CR.CompletedDate = "12/25/9999"
Else: CR.CompletedDate = CRCompletedDateBox.Value
End If

If CRScanDateBox.Value = vbNullString Then
    CR.ScanDate = "12/25/9999"
Else: CR.ScanDate = CRScanDateBox.Value
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Assigns the proper month affix to the Prefix,
'!!!points the CR to the proper Keylog page,
'!!!and preps the CR for loading into the proper CRs page
'
'Example: CR AA12301 is dated 01/01/01
'CR.Prefix is loaded with "A"
'
'The logic below checks the month value of CR.CRDate
'which is loaded with the serial value of '01/01/01.'
'
'The logic returns True for If Month(CR.CRDate) = 1
'The logic then affixes the January affix, "A," to CR.Prefix
'
'CR.Prefix is now loaded with "AA"

Dim crpage As String
Dim PrefixPrefix As String

If Month(CR.CRDate) = 1 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "A"
    'Worksheets("JANUARY KEYLOG").Activate
    crpage = "JANUARY CRs"
    
ElseIf Month(CR.CRDate) = 2 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "B"
    'Worksheets("FEBRUARY KEYLOG").Activate
    crpage = "FEBRUARY CRs"
    
ElseIf Month(CR.CRDate) = 3 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "H"
    'Worksheets("MARCH KEYLOG").Activate
    crpage = "MARCH CRs"
    
ElseIf Month(CR.CRDate) = 4 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "R"
    'Worksheets("APRIL KEYLOG").Activate
    crpage = "APRIL CRs"
    
ElseIf Month(CR.CRDate) = 5 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "Y"
    'Worksheets("MAY KEYLOG").Activate
    crpage = "MAY CRs"
    
ElseIf Month(CR.CRDate) = 6 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "E"
    'Worksheets("JUNE KEYLOG").Activate
    crpage = "JUNE CRs"
    
ElseIf Month(CR.CRDate) = 7 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "L"
    'Worksheets("JULY KEYLOG").Activate
    crpage = "JULY CRs"
    
ElseIf Month(CR.CRDate) = 8 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "G"
    'Worksheets("AUGUST KEYLOG").Activate
    crpage = "AUGUST CRs"
    
ElseIf Month(CR.CRDate) = 9 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "P"
    'Worksheets("SEPTEMBER KEYLOG").Activate
    crpage = "SEPTEMBER CRs"
    
ElseIf Month(CR.CRDate) = 10 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "T"
    'Worksheets("OCTOBER KEYLOG").Activate
    crpage = "OCTOBER CRs"
    
ElseIf Month(CR.CRDate) = 11 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "V"
    'Worksheets("NOVEMBER KEYLOG").Activate
    crpage = "NOVEMBER CRs"
    
ElseIf Month(CR.CRDate) = 12 Then
    PrefixPrefix = CR.Prefix
    CR.Prefix = CR.Prefix & "D"
    'Worksheets("DECEMBER KEYLOG").Activate
    crpage = "DECEMBER CRs"
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Selects the proper KEYLOG sheet
If Month(CR.KeyDate) = 1 Then
    'CR.Prefix = CR.Prefix & "A"
    Worksheets("JANUARY KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 2 Then
    'CR.Prefix = CR.Prefix & "B"
    Worksheets("FEBRUARY KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 3 Then
    'CR.Prefix = CR.Prefix & "H"
   Worksheets("MARCH KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 4 Then
    'CR.Prefix = CR.Prefix & "R"
    Worksheets("APRIL KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 5 Then
    'CR.Prefix = CR.Prefix & "Y"
    Worksheets("MAY KEYLOG").Activate

ElseIf Month(CR.KeyDate) = 6 Then
    'CR.Prefix = CR.Prefix & "E"
    Worksheets("JUNE KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 7 Then
    'CR.Prefix = CR.Prefix & "L"
    Worksheets("JULY KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 8 Then
    'CR.Prefix = CR.Prefix & "G"
    Worksheets("AUGUST KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 9 Then
    'CR.Prefix = CR.Prefix & "P"
    Worksheets("SEPTEMBER KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 10 Then
    'CR.Prefix = CR.Prefix & "T"
    Worksheets("OCTOBER KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 11 Then
    'CR.Prefix = CR.Prefix & "V"
    Worksheets("NOVEMBER KEYLOG").Activate
    
ElseIf Month(CR.KeyDate) = 12 Then
    'CR.Prefix = CR.Prefix & "D"
    Worksheets("DECEMBER KEYLOG").Activate
    
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Catches unchecked/unreturned/incomplete CRs
If CR.CheckDate = "12/25/9999" Then
    CR.CheckStatus = "NOT CHECKED"
    'Worksheets("PENDING").Activate
    'Range("D1").Value = CR.CheckStatus
Else: CR.CheckStatus = "CHECKED"
    'Range("D1").Value = CR.CheckDate
End If

If CR.ReturnDate = "12/25/9999" Then
    CR.ReturnStatus = "NOT RETURNED"
    'Worksheets("PENDING").Activate
    'Range("E1").Value = CR.ReturnStatus
Else: CR.ReturnStatus = "RETURNED"
    'Range("E1").Value = CR.ReturnDate
End If

If CR.CompletedDate = "12/25/9999" Then
    CR.CompletedStatus = "NOT COMPLETE"
    'Worksheets("PENDING").Activate
    'Range("F1").Value = CR.CompletedStatus
Else: CR.CompletedStatus = "COMPLETE"
    'Range("F1").Value = CR.CompletedDate
End If

If CR.ScanDate = "12/25/9999" Then
    CR.ScanStatus = "NOT SCANNED"
    'Worksheets("PENDING").Activate
    'Range("F1").Value = CR.CompletedStatus
Else: CR.ScanStatus = "SCANNED"
    'Range("F1").Value = CR.CompletedDate
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'!!!If CR is Complete,
'       Load the CR data into the keylog,
'       Load the CR number into the proper month,
'       ???Store the CR in the Keyed_Cr_Dictionary
'       ???Store the CR in the Completed_CR_Dictionary.
'
'   Else,
'       Load the CR number into the proper month,
'       ???Load the CR data into the Pending screen,
'       Load the CR number into the proper month,
'       ???Store the CR in the Keyed_CR_Dictionary,
'       ???Store the CR in the Pending_CR_Dictionary.

'Uncomment next two lines to route to Testing Output sheet
'Worksheets("TestingOut").Activate
'crpage = "TestingOut"

If CR.StatusArray(0) = "CHECKED" And CR.StatusArray(1) = "RETURNED" And CR.StatusArray(2) = "COMPLETE" Then
    Dim LastRow As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row + 1
    Range("A" & LastRow).Value = CR.KeyDate
    'Range("B" & LastRow).Value = ???
    Range("C" & LastRow).Value = CR.Prefix & CR.Entry
    Range("D" & LastRow).Value = CR.AA
    Range("E" & LastRow).Value = CR.CheckDate
    Range("F" & LastRow).Value = CR.ReturnDate
    Range("G" & LastRow).Value = CR.CompletedDate
    Range("O" & LastRow).Value = CR.ScanDate
    
    'Changes to _CR page
    Worksheets(crpage).Activate
    
    'Something using:
    '
    'Dim myArray As Variant
    'myArray = Range("A1:D4").Value
    'Range("A1:D4").Value = myArray
    '?

'???Else log the entry in Pending.
Else:
    LastRow = Range("A" & Rows.Count).End(xlUp).Row + 1
    Range("A" & LastRow).Value = CR.KeyDate
    'Range("B" & LastRow).Value = ???
    Range("C" & LastRow).Value = CR.Prefix & CR.Entry
    Range("D" & LastRow).Value = CR.AA
    
    If CR.CheckStatus = "NOT CHECKED" Then
        Range("E" & LastRow).Value = CR.CheckStatus
    Else:
        Range("E" & LastRow).Value = CR.CheckDate
    End If
    
    If CR.ReturnStatus = "NOT RETURNED" Then
        Range("F" & LastRow).Value = CR.ReturnStatus
    Else:
        Range("F" & LastRow).Value = CR.ReturnDate
    End If
        
    If CR.CompletedStatus = "NOT COMPLETE" Then
        Range("G" & LastRow).Value = CR.CompletedStatus
    Else:
        Range("G" & LastRow).Value = CR.CompletedDate
    End If
    
    If CR.ScanStatus = "NOT SCANNED" Then
        Range("O" & LastRow).Value = CR.ScanStatus
    Else:
        Range("O" & LastRow).Value = CR.ScanDate
    End If
    
    Worksheets(crpage).Activate

End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Selects the right prefix header, and
'Enters the CR number under that header.
'
'If a prefix is added to the Config sheet,
'This section should be updated to reflect it.
'!!!Refactor to make this gate dynamic.
Dim Ipart As String
Dim Ip As Integer
Ipart = Right(CR.Entry, 3)
Ip = CInt(Ipart)

If PrefixPrefix = "A" Then
    Range("A2").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A3").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A4").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A5").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A6").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A7").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A8").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A9").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "B" Then
    Range("A12").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A13").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A14").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A15").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A16").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A17").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A18").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A19").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "C" Then
    Range("A22").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A23").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A24").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A25").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A26").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A27").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A28").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A29").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "D" Then
    Range("A32").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A33").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A34").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A35").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A36").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A37").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A38").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A39").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "E" Then
    Range("A42").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A43").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A44").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A45").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A46").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A47").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A48").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A49").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "F" Then
    Range("A52").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A53").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A54").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A55").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A56").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A57").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A58").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A59").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "G" Then
    Range("A62").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A63").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A64").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A65").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A66").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A67").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A68").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A69").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "H" Then
    Range("A72").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A73").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A74").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A75").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A76").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A77").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A78").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A79").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "K" Then
    Range("A82").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A83").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A84").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A85").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A86").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A87").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A88").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A89").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "L" Then
    Range("A92").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A93").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A94").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A95").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A96").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A97").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A98").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A99").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "M" Then
    Range("A102").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A103").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A104").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A105").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A106").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A107").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A108").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A109").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "N" Then
    Range("A112").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A113").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A114").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A115").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A116").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A117").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A118").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A119").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "0" Then
    Range("A122").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A123").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A124").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A125").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A126").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A127").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A128").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A129").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "T" Then
    Range("A132").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A133").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A134").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A135").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A136").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A137").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A138").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A139").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "W" Then
    Range("A142").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A143").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A144").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A145").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A146").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A147").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A148").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A149").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
    
ElseIf PrefixPrefix = "Z" Then
    Range("A152").Select
    If Ip <= 25 Then
        ActiveCell.Offset(0, Ip).Value = Ip
    ElseIf Ip <= 50 Then
        Range("A153").Select
        ActiveCell.Offset(0, Abs(25 - Ip)).Value = Ip
    ElseIf Ip <= 75 Then
        Range("A154").Select
        ActiveCell.Offset(0, Abs(50 - Ip)).Value = Ip
    ElseIf Ip <= 100 Then
        Range("A155").Select
        ActiveCell.Offset(0, Abs(75 - Ip)).Value = Ip
    ElseIf Ip <= 125 Then
        Range("A156").Select
        ActiveCell.Offset(0, Abs(100 - Ip)).Value = Ip
    ElseIf Ip <= 150 Then
        Range("A157").Select
        ActiveCell.Offset(0, Abs(125 - Ip)).Value = Ip
    ElseIf Ip <= 175 Then
        Range("A158").Select
        ActiveCell.Offset(0, Abs(150 - Ip)).Value = Ip
    ElseIf Ip <= 200 Then
        Range("A159").Select
        ActiveCell.Offset(0, Abs(175 - Ip)).Value = Ip
    End If
End If

Unload Me

End Sub
