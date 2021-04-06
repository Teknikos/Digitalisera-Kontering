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
' ###### Kod för Startsidans "Skapa Ny Klient" ######
' ###################################################


Private Sub CommandButtonSkapaKlient_Click()
'
' Skapar en ny klient på ny flik.
'
Dim offsetCounter As Integer
offsetCounter = 1 ' Denna behövs pga macro "SkapaNyaKunderFrånLista" ska fungera mot Länkar.LänkaKundLista

SkapaNyKlient.SkapaKlient
PlusKnappar.Knapp_UtökaKundlista
Länkar.LänkaKundlista (offsetCounter)
Sheets(Sheets.Count).Select

End Sub

Private Sub OptionButtonEnkelKund_Click()
' Tar bort rutan med momsprocent när 'ingen moms'
TextBoxMomsProcent.Visible = False
End Sub

Private Sub OptionButtonMoms0_Click()
' Tar bort rutan med momsprocent när 'ingen moms'
TextBoxMomsProcent.Visible = False
End Sub

Private Sub OptionButtonMoms1_Click()
' Visar rutan momsprocent när 'moms' är vald
TextBoxMomsProcent.Visible = True
End Sub

Private Sub OptionButtonEnkelKund_Hover()
' Sätter en info ruta över Enkel Kund knappen vid mouseover.
    Dim r As Range
    Set r = Selection
    
    Sheets("Start").OptionButtonEnkelKund.Select
    ActiveSheet.Hyperlinks.Add _
        Anchor:=Selection.ShapeRange.Item(1), _
        Address:="", _
        ScreenTip:="Enkla kunder konteras utifrån ""Enkla Kontoplanen"" istället för en satt budget."
    
    r.Select
End Sub
