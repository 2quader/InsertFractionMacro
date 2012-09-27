VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFractions 
   Caption         =   "Bruch einfügen"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   OleObjectBlob   =   "frmFractions.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
    Unload Me
End Sub

' Bruch an Cursorpos. in Dokument einfügen
Private Sub btnInsert_Click()
    ' Wenn die Eingabefelder nicht leer sind...
    If txtZaehler.Text <> "" And txtNenner.Text <> "" Then
        ' Feldbefehl "zusammenbauen"
        Dim code As String
        code = "EQ \F("
        code = code + txtZaehler.Text
        code = code + ";"
        code = code + txtNenner.Text
        code = code + ")"
        
        'Wurde auch ein Ergebnis eingegeben?
        If txtErgebnis.Text <> "" Then
            code = code + " = "
            code = code + txtErgebnis.Text
        End If
    
        ' Feld an Cursorpos. einfügen
        Selection.Fields.Add Selection.Range, wdFieldEmpty, code, False
    Else 'Felder sind leer
        MsgBox ("Keine Eingabe!")
    End If
End Sub

Private Sub btnZaehlerHochstellen_Click()
    Dim userInput As String
    userInput = InputBox("Hochgestelltes eingeben:")
    If userInput <> "" Then
        txtZaehler.Text = txtZaehler.Text + "\S\up4(" + userInput + ")"
    End If
    txtZaehler.SetFocus
End Sub

Private Sub btnZaehlerTiefstellen_Click()
    Dim userInput As String
    userInput = InputBox("Tiefgestelltes eingeben:")
    If userInput <> "" Then
        txtZaehler.Text = txtZaehler.Text + "\S\do4(" + userInput + ")"
    End If
    txtZaehler.SetFocus
End Sub

Private Sub btnNennerHochstellen_Click()
    Dim userInput As String
    userInput = InputBox("Hochgestelltes eingeben:")
    If userInput <> "" Then
        txtNenner.Text = txtNenner.Text + "\S\up4(" + userInput + ")"
    End If
    txtNenner.SetFocus
End Sub

Private Sub btnNennerTiefstellen_Click()
    Dim userInput As String
    userInput = InputBox("Tiefgestelltes eingeben:")
    If userInput <> "" Then
        txtNenner.Text = txtNenner.Text + "\S\do4(" + userInput + ")"
    End If
    txtNenner.SetFocus
End Sub

Private Sub btnErgebnisHochstellen_Click()
    Dim userInput As String
    userInput = InputBox("Hochgestelltes eingeben:")
    If userInput <> "" Then
        txtErgebnis.Text = txtErgebnis.Text + "\S\up4(" + userInput + ")"
    End If
    txtErgebnis.SetFocus
End Sub

Private Sub btnErgebnisTiefstellen_Click()
    Dim userInput As String
    userInput = InputBox("Tiefgestelltes eingeben:")
    If userInput <> "" Then
        txtErgebnis.Text = txtErgebnis.Text + "\S\do4(" + userInput + ")"
    End If
    txtErgebnis.SetFocus
End Sub
