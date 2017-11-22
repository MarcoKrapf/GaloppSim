VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    'Set playmode
        PlayMode = "RS"
    'Standardeinstellungen der Menüband-Elemente setzen
        betting = False
        taktik = True
        hufspur = True
        namen = True
        pferdefarbe = False
        markFav = True
        namen2 = True
        colResult = False
        spannung = True
        zoom = True
        maxi = True
        headings = False
        gitter = False

    'No race chosen
        Auswahl = ""
    
    If PlayMode = "RS" Then Call derby_RS
    
End Sub
