Attribute VB_Name = "modUserformPlatzieren"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Platzieren der Pop-ups, vor allem wenn mehrere Monitore angeschlossen sind
'------------------------------------------------------------------------------------------

Public Sub UserFormPlatzieren(frmMe As Object)
    With frmMe
        .StartUpPosition = 0
        .Top = ActiveWindow.Top + ((ActiveWindow.Height - frmMe.Height) / 2)
        .Left = ActiveWindow.Left + ((ActiveWindow.Width - frmMe.Width) / 2) + 280
    End With
End Sub
