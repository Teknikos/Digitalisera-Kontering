VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ###################################################
' ###### Kod f�r Startsidans "Skapa Ny Klient" ######
' ###################################################


Private Sub CommandButtonSkapaKlient_Click()
'
' Skapar en ny klient p� ny flik.
'
Dim offsetCounter As Integer
offsetCounter = 1 ' Denna beh�vs pga macro "SkapaNyaKunderFr�nLista" ska fungera mot L�nkar.L�nkaKundLista

SkapaNyKlient.SkapaKlient
PlusKnappar.Knapp_Ut�kaKundlista
L�nkar.L�nkaKundlista (offsetCounter)
Sheets(Sheets.Count).Select

End Sub

Private Sub OptionButtonEnkelKund_Click()
' Tar bort rutan med momsprocent n�r 'ingen moms'
TextBoxMomsProcent.Visible = False
End Sub

Private Sub OptionButtonMoms0_Click()
' Tar bort rutan med momsprocent n�r 'ingen moms'
TextBoxMomsProcent.Visible = False
End Sub

Private Sub OptionButtonMoms1_Click()
' Visar rutan momsprocent n�r 'moms' �r vald
TextBoxMomsProcent.Visible = True
End Sub

Private Sub OptionButtonEnkelKund_Hover()
' S�tter en info ruta �ver Enkel Kund knappen vid mouseover.
    Dim r As Range
    Set r = Selection
    
    Sheets("Start").OptionButtonEnkelKund.Select
    ActiveSheet.Hyperlinks.Add _
        Anchor:=Selection.ShapeRange.Item(1), _
        Address:="", _
        ScreenTip:="Enkla kunder konteras utifr�n ""Enkla Kontoplanen"" ist�llet f�r en satt budget."
    
    r.Select
End Sub