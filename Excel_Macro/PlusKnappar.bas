Attribute VB_Name = "PlusKnappar"
Sub Knapp_Ut�kaIns�ttningar()
'
' Knapp_Ut�kaIns�ttningar Makro
' Ut�kar rutan "Inst�ttningar mm" med 1 rad.
'

'
    Range("firstIns�ttningar").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub
Sub Knapp_Ut�kaKonteringsinfo()
'
' Knapp_Ut�kaKonteringsinfo Makro
' Ut�kar rutan "Konteringsinfo/anteckningar/p�g�ende projekt" med 1 rad.
'

    Range("firstKonteringsinfo").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub
Sub Knapp_Ut�kaF�retagskort()
'
' Knapp_Ut�kaF�retagskort Makro
' Ut�kar rutan "BankNamn f�retagskort" med 1 rad.
'

    Range("firstF�retagskort").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
End Sub

Sub Knapp_Ut�kaMoms()
'
' Knapp_Ut�kaMoms Makro
' Ut�kar rutan "Momsdragningslista + Momsnyckel" med 1 rad.
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

Sub Knapp_Ut�kaKundlista()
'
' Knapp_Ut�kaKundlista Makro
' Ut�kar listan "Kundlistan" p� Start-fliken med 1 rad uppifr�n.
'
    Sheets("Start").Select
    Range("firstKlientlista").Offset(1, 0).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
End Sub
