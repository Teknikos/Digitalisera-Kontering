Attribute VB_Name = "SkapaNyKlient"
Sub SkapaKlient()
'
' Skapar en ny klient på ny flik.
'

If Sheets("Start").OptionButtonEnkelKund.Value = True Then
' Skapa Enkel_Kund
    Sheets("Mall_Enkel_Kund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' Döper blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value 'Namnger kunden
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxFörvaltare.Value ' Namnger ekonomiskförvaltare för kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("F2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' Sätter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("F2") = Sheets("Start").OptionButtonKontorÖstersund.Caption ' Sätter kontor till Östersund
    End If

ElseIf Sheets("Start").OptionButtonMoms0.Value = True Then
' Skapa Ej_Momskund
    Sheets("Mall_Ej_Momskund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' Döper blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value 'Namnger kunden
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxFörvaltare.Value ' Namnger ekonomiskförvaltare för kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' Sätter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorÖstersund.Caption ' Sätter kontor till Östersund
    End If
    
Else
' Skapa Momskund
    Sheets("Mall_Momskund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' Döper blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value ' Namnger kunden
    Sheets(Sheets.Count).Range("B1") = Sheets("Start").TextBoxMomsProcent.Value & "%" ' Sätter momssats
    Sheets(Sheets.Count).Range("B1").NumberFormat = "0.00%"
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxFörvaltare.Value ' Namnger ekonomiskförvaltare för kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' Sätter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorÖstersund.Caption ' Sätter kontor till Östersund
    End If
End If
Sheets(Sheets.Count).Activate
MsgBox (Sheets("Start").TextBoxKlientID.Value & vbCrLf & Sheets("Start").TextBoxKundnamn.Value & " Skapad")

End Sub



