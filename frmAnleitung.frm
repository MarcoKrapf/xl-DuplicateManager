VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnleitung 
   Caption         =   "[xl DuplikatManager - Anleitung]"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   OleObjectBlob   =   "frmAnleitung.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAnleitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    'Erste Registerkarte aktivieren
    MultiPageAnleitung.Value = 0
End Sub
