VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cRPrefix As String
Private cREntry As String
Private cRAA As String
Private cRCRDate As String
Private cRKeyDate As String

Private cRCheckDate As String
Private cRCheckStatus As String

Private cRReturnDate As String
Private cRReturnStatus As String

Private cRCompletedDate As String
Private cRCompletedStatus As String

Private cRScanDate As String
Private cRScanStatus As String

Private cRStatusArray(0 To 3) As String
 
Public Property Get Prefix() As String
 
    Prefix = cRPrefix
     
End Property

Public Property Let Prefix(aPrefix As String)

    cRPrefix = aPrefix

End Property
 
Public Property Get Entry() As String
 
    Entry = cREntry
     
End Property

Property Let Entry(aEntry As String)

    cREntry = aEntry

End Property

Public Property Get AA() As String
 
    AA = cRAA
     
End Property

Property Let AA(aAA As String)

    cRAA = aAA

End Property

Property Get CRDate() As String
 
    CRDate = cRCRDate
     
End Property

Public Property Let CRDate(aDate As String)

    cRCRDate = DateValue(aDate)

End Property

Property Get KeyDate() As String
 
    KeyDate = cRKeyDate
     
End Property

Public Property Let KeyDate(aDate As String)

    cRKeyDate = DateValue(aDate)

End Property

Public Property Get CheckDate() As String
 
    CheckDate = cRCheckDate
     
End Property

Public Property Let CheckDate(aDate As String)

    cRCheckDate = DateValue(aDate)

End Property

Public Property Get CheckStatus() As String
 
    CheckStatus = cRCheckStatus
     
End Property

Public Property Let CheckStatus(aStatus As String)

    cRCheckStatus = aStatus

End Property

Public Property Get ReturnDate() As String
 
    ReturnDate = cRReturnDate
     
End Property

Public Property Let ReturnDate(aDate As String)

    cRReturnDate = DateValue(aDate)

End Property


Public Property Get ReturnStatus() As String
 
    ReturnStatus = cRReturnStatus
     
End Property

Public Property Let ReturnStatus(aStatus As String)

    cRReturnStatus = aStatus

End Property

Public Property Get CompletedDate() As String
 
    CompletedDate = cRCompletedDate
     
End Property

Public Property Let CompletedDate(aDate As String)

    cRCompletedDate = DateValue(aDate)

End Property

Public Property Get CompletedStatus() As String
 
    CompletedStatus = cRCompletedStatus
     
End Property

Public Property Let CompletedStatus(aStatus As String)

    cRCompletedStatus = aStatus

End Property

Public Property Get ScanDate() As String
 
    ScanDate = cRScanDate
     
End Property

Public Property Let ScanDate(aDate As String)

    cRScanDate = DateValue(aDate)
    
End Property

Public Property Get ScanStatus() As String
 
    ScanStatus = cRScanStatus
     
End Property

Public Property Let ScanStatus(aStatus As String)

    cRScanStatus = aStatus

End Property

Public Property Get StatusArray()
    
    cRStatusArray(0) = cRCheckStatus
    cRStatusArray(1) = cRReturnStatus
    cRStatusArray(2) = cRCompletedStatus
    cRStatusArray(3) = cRScanStatus
    StatusArray = cRStatusArray

End Property
