Attribute VB_Name = "SkapaNyKlient"
Sub SkapaKlient()
'
' Skapar en ny klient p� ny flik.
'

If Sheets("Start").OptionButtonEnkelKund.Value = True Then
' Skapa Enkel_Kund
    Sheets("Mall_Enkel_Kund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' D�per blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value 'Namnger kunden
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxF�rvaltare.Value ' Namnger ekonomiskf�rvaltare f�r kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("F2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' S�tter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("F2") = Sheets("Start").OptionButtonKontor�stersund.Caption ' S�tter kontor till �stersund
    End If

ElseIf Sheets("Start").OptionButtonMoms0.Value = True Then
' Skapa Ej_Momskund
    Sheets("Mall_Ej_Momskund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' D�per blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value 'Namnger kunden
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxF�rvaltare.Value ' Namnger ekonomiskf�rvaltare f�r kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' S�tter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontor�stersund.Caption ' S�tter kontor till �stersund
    End If
    
Else
' Skapa Momskund
    Sheets("Mall_Momskund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Name = Sheets("Start").TextBoxKlientID.Value ' D�per blad till KlientID
    Sheets(Sheets.Count).Range("A1") = Sheets("Start").TextBoxKundnamn.Value ' Namnger kunden
    Sheets(Sheets.Count).Range("B1") = Sheets("Start").TextBoxMomsProcent.Value & "%" ' S�tter momssats
    Sheets(Sheets.Count).Range("B1").NumberFormat = "0.00%"
    Sheets(Sheets.Count).Range("A2") = Sheets("Start").TextBoxF�rvaltare.Value ' Namnger ekonomiskf�rvaltare f�r kunden
    If Sheets("Start").OptionButtonKontorUppsala.Value = True Then
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontorUppsala.Caption ' S�tter kontor till Uppsala
    Else
        Sheets(Sheets.Count).Range("B2") = Sheets("Start").OptionButtonKontor�stersund.Caption ' S�tter kontor till �stersund
    End If
End If
Sheets(Sheets.Count).Activate
MsgBox (Sheets("Start").TextBoxKlientID.Value & vbCrLf & Sheets("Start").TextBoxKundnamn.Value & " Skapad")

End Sub



