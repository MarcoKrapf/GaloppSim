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
    'Standardeinstellungen der Menüband-Elemente setzen
        taktik = True
        namen = False
        namen2 = False
        pferdefarbe = False
        hufspur = True
        spannung = True
        maxi = True
        zoom = True
        gitter = True
        headings = True
    'Rennen auswählen
        Auswahl = "race_DERBY17HH"
End Sub
