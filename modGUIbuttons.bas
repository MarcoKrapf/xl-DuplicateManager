Attribute VB_Name = "modGUIbuttons"
Option Explicit
Option Private Module

'Modulbeschreibung:
'Die Funktionen dieses Moduls geben dynamische Beschriftungen für die GUI-Buttons
'zurück in Abhängigkeit der Sprache und des jeweiligen Button-Zustands
'------------------------------------------------------------------------------------------

Public Function buttonDuplikatfensterOnOff(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonDuplikatfensterOnOff = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(1, Sprache).Value
    Else
        buttonDuplikatfensterOnOff = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(2, Sprache).Value
    End If
    
End Function

Public Function buttonSuchModus(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonSuchModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(4, Sprache).Value
    Else
        buttonSuchModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(5, Sprache).Value
    End If
    
End Function

Public Function buttonMarkierModus(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonMarkierModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(7, Sprache).Value
    Else
        buttonMarkierModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(8, Sprache).Value
    End If
    
End Function

Public Function buttonAusgabeModus(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonAusgabeModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(10, Sprache).Value
    Else
        buttonAusgabeModus = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(11, Sprache).Value
    End If
    
End Function

Public Function buttonLoeschModusZeilen(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonLoeschModusZeilen = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(13, Sprache).Value
    Else
        buttonLoeschModusZeilen = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(14, Sprache).Value
    End If
    
End Function

Public Function buttonLoeschModusKomprimieren(ByRef Sprache As Integer, ByRef button As Boolean) As String
    
    If button = True Then
        buttonLoeschModusKomprimieren = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(16, Sprache).Value
    Else
        buttonLoeschModusKomprimieren = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(17, Sprache).Value
    End If
    
End Function

Public Function buttonDuplikatfenster(ByRef Sprache As Integer, ByRef status As Byte) As String
    
    Select Case status
        Case 0: buttonDuplikatfenster = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(19, Sprache).Value
        Case 1: buttonDuplikatfenster = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(20, Sprache).Value
        Case 2: buttonDuplikatfenster = ThisWorkbook.Worksheets("Dynamic_GUI").Cells(21, Sprache).Value
    End Select
    
End Function
