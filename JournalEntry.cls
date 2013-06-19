VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private jEEntry As String
Private jEJEDate As String
Private jEKeyDate As String

Private jECheckDate As String
Private jECheckStatus As String

Private jEReturnDate As String
Private jEReturnStatus As String

Private jECompletedDate As String
Private jECompletedStatus As String

Private jEScanDate As String
Private jEScanStatus As String

Private jEStatusArray(0 To 3) As String
  
Public Property Get Entry() As String
 
    Entry = jEEntry
     
End Property

Property Let Entry(aEntry As String)

    jEEntry = aEntry

End Property

Property Get JEDate() As String
 
    JEDate = jEJEDate
     
End Property

Public Property Let JEDate(aDate As String)

    jEJEDate = DateValue(aDate)

End Property

Property Get KeyDate() As String
 
    KeyDate = jEKeyDate
     
End Property

Public Property Let KeyDate(aDate As String)

    jEKeyDate = DateValue(aDate)

End Property

Public Property Get CheckDate() As String
 
    CheckDate = jECheckDate
     
End Property

Public Property Let CheckDate(aDate As String)

    jECheckDate = DateValue(aDate)

End Property

Public Property Get CheckStatus() As String
 
    CheckStatus = jECheckStatus
     
End Property

Public Property Let CheckStatus(aStatus As String)

    jECheckStatus = aStatus

End Property

Public Property Get ReturnDate() As String
 
    ReturnDate = jEReturnDate
     
End Property

Public Property Let ReturnDate(aDate As String)

    jEReturnDate = DateValue(aDate)

End Property


Public Property Get ReturnStatus() As String
 
    ReturnStatus = jEReturnStatus
     
End Property

Public Property Let ReturnStatus(aStatus As String)

    jEReturnStatus = aStatus

End Property

Public Property Get CompletedDate() As String
 
    CompletedDate = jECompletedDate
     
End Property

Public Property Let CompletedDate(aDate As String)

    jECompletedDate = DateValue(aDate)

End Property

Public Property Get CompletedStatus() As String
 
    CompletedStatus = jECompletedStatus
     
End Property

Public Property Let CompletedStatus(aStatus As String)

    jECompletedStatus = aStatus

End Property

Public Property Get ScanDate() As String
 
    ScanDate = jEScanDate
     
End Property

Public Property Let ScanDate(aDate As String)

    jEScanDate = DateValue(aDate)
    
End Property

Public Property Get ScanStatus() As String
 
    ScanStatus = jEScanStatus
     
End Property

Public Property Let ScanStatus(aStatus As String)

    jEScanStatus = aStatus

End Property

Public Property Get StatusArray()
    
    jEStatusArray(0) = jECheckStatus
    jEStatusArray(1) = jEReturnStatus
    jEStatusArray(2) = jECompletedStatus
    jEStatusArray(3) = jEScanStatus
    StatusArray = jEStatusArray

End Property

