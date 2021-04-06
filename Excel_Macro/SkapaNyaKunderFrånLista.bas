Attribute VB_Name = "SkapaNyaKunderFr�nLista"
Sub SkapaKlienterFr�nLista()
Attribute SkapaKlienterFr�nLista.VB_ProcData.VB_Invoke_Func = " \n14"
' ###########################################################
' ###  Skapar kunder p� egna ark utifr�n listan p� start  ###
' ###########################################################
        
Dim MyCell As Range, MyRange As Range
Dim offsetCounter As Integer
offsetCounter = 0 ' Beh�vs f�r att l�nkar ska kopplas till klient id f�r hela listan

Set MyRange = Sheets("Start").Range("A6") ' S�tter sikte p� 1:a cellen i kolumnen KlientID (refereras som "MyCell")
Set MyRange = Range(MyRange, MyRange.End(xlDown)) ' Utvidgar siktet till slutet p� kolumnen

FormateraLista

For Each MyCell In MyRange
    If IsNumeric(MyCell.Offset(0, 2).Value) And MyCell.Offset(0, 2).Value > 0 Then ' Kollar om det �r momskund eller inte.
        ' Skapa en Momskund fr�n klientListan
        Sheets("Mall_Momskund").Copy After:=Sheets(Sheets.Count) ' Mall ligger som dolt ark
        Sheets(Sheets.Count).Name = MyCell.Value ' Namnger arbetsbladet till klientID
        Sheets(Sheets.Count).Range("A1") = MyCell.Offset(0, 1).Value ' Brf Namn
        Sheets(Sheets.Count).Range("B1") = MyCell.Offset(0, 2).Value & "%" ' MomsNyckel
        Sheets(Sheets.Count).Range("A2") = MyCell.Offset(0, 3).Value ' Ansvarig Ekonom
        '
    ElseIf IsEmpty(MyCell.Offset(0, 2)) Or MyCell.Offset(0, 2).Value = 0 Then
        ' Skapa en Ej_MomsKund fr�n KlientListan
        Sheets("Mall_Ej_Momskund").Copy After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = MyCell.Value
        Sheets(Sheets.Count).Range("A1") = MyCell.Offset(0, 1).Value
        Sheets(Sheets.Count).Range("A2") = MyCell.Offset(0, 3).Value
    Else
        ' Skapa en Enkel_Kund fr�n klientListan
        Sheets("Mall_Enkel_Kund").Copy After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = MyCell.Value
        Sheets(Sheets.Count).Range("A1") = MyCell.Offset(0, 1).Value
        Sheets(Sheets.Count).Range("A2") = MyCell.Offset(0, 3).Value
    End If
    Sheets(Sheets.Count).Visible = True ' S�tter bladet som synligt (Mallar ligger dolda)
    
    offsetCounter = offsetCounter + 1
    Call L�nkar.L�nkaKundlista(offsetCounter)
    
Next MyCell


End Sub

Sub FormateraLista()
' MomsnyckeListan byt ut punkt mot komma.
Columns("C:C").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "General"
End Sub
