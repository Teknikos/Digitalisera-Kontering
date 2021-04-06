Attribute VB_Name = "PlusKnappar"
Sub Knapp_UtökaInsättningar()
'
' Knapp_UtökaInsättningar Makro
' Utökar rutan "Instättningar mm" med 1 rad.
'

'
    Range("firstInsättningar").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub
Sub Knapp_UtökaKonteringsinfo()
'
' Knapp_UtökaKonteringsinfo Makro
' Utökar rutan "Konteringsinfo/anteckningar/pågående projekt" med 1 rad.
'

    Range("firstKonteringsinfo").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub
Sub Knapp_UtökaFöretagskort()
'
' Knapp_UtökaFöretagskort Makro
' Utökar rutan "BankNamn företagskort" med 1 rad.
'

    Range("firstFöretagskort").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
End Sub

Sub Knapp_UtökaMoms()
'
' Knapp_UtökaMoms Makro
' Utökar rutan "Momsdragningslista + Momsnyckel" med 1 rad.
'
    
    Range("firstMomsdragningslista").Offset(3, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Range("secondMomsdragningslista").Offset(3, 0).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinous
        .Weight = xlThin
    End With
    Range("secondMomsdragningslista").Offset(3, -2).Select
    
End Sub

Sub Knapp_UtökaKundlista()
'
' Knapp_UtökaKundlista Makro
' Utökar listan "Kundlistan" på Start-fliken med 1 rad uppifrån.
'
    Sheets("Start").Select
    Range("firstKlientlista").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
End Sub
