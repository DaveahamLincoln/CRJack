VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryTypeMenu 
   Caption         =   "Select Entry Type"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   OleObjectBlob   =   "EntryTypeMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EntryTypeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()

Unload Me

End Sub

Private Sub CashReceiptButton_Click()

NewCashReceipt.Show
Unload Me

End Sub

Private Sub JournalEntryButton_Click()

NewJournalEntry.Show
Unload Me

End Sub

Private Sub PayrollButton_Click()
NewPayrollEntry.Show
Unload Me

End Sub

Private Sub UserForm_Click()

End Sub
