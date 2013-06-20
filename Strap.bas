Attribute VB_Name = "Strap"
Sub NewEntry()
Attribute NewEntry.VB_ProcData.VB_Invoke_Func = "E\n14"
    With EntryTypeMenu
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
