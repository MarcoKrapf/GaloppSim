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
    'Set standard settings for options
        betting = False
        taktik = True
        hufspur = True
        namen = True
        pferdefarbe = True
        markFav = True
        namen2 = True
        colResult = False
        spannung = False
        zoom = True
        maxi = True
        headings = False
        gitter = False
    
    'Set standard language
        Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            Case msoLanguageIDGerman, msoLanguageIDGermanAustria, _
                    msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg
                language = "DE"
'            Case msoLanguageIDRussian, msoLanguageIDUkrainian
'                language = "RU"
            Case Else
                language = "EN"
        End Select

    'No race chosen
        Auswahl = ""
    
    If PlayMode = "RS" Then Call derby_RS
    
End Sub
