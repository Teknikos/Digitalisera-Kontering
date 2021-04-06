Attribute VB_Name = "Länkar"
Sub LänkaKundlista(offsetCounter As Integer)
'Adderar länk till kundflik och populerar kundlista vid nyskapande av kund.
    ' Byt kund ID
    Sheets("Start").Select
    Range("StartFirstKlientID").Offset(offsetCounter, 0).Select
    Selection = Sheets(Sheets.Count).Name
    Dim tempString As String
    tempString = Selection
    ' Länka till senast skapade arket.
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=Sheets(Sheets.Count).Name & "!A1"
    ' Byt Kundnamn
    Sheets("Start").Range("StartFirstKlientID").Offset(offsetCounter, 1).Select
    Selection = Sheets(Sheets.Count).Range("KundNamn")
    ' Byt Momsnyckel om momskund
    If IsNumeric(Sheets(Sheets.Count).Range("Momsnyckel")) Then
        Sheets("Start").Range("StartFirstKlientID").Offset(offsetCounter, 2).Select
        Selection = Sheets(Sheets.Count).Range("Momsnyckel")
        Selection.NumberFormat = "0.00%"
    End If
    ' Byt Ansvarig förvaltare
    Sheets("Start").Range("StartFirstKlientID").Offset(offsetCounter, 3).Select
    Selection = Sheets(Sheets.Count).Range("Förvaltare")
End Sub


