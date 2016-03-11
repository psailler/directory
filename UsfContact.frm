VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsfContact 
   Caption         =   "Formulaire de contact"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   OleObjectBlob   =   "UsfContact.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UsfContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Dim L As Integer
'Confirmation d'ajout
If MsgBox("Etes-vous certain de vouloir enregistrer ce contact ?", vbYesNo, "Demande de confirmation") = vbYes Then
'Permet de se positionner sur la dernière ligne de tableau VIDE
L = Sheets(ActiveSheet.Name).Range("A1048576").End(xlUp).Row + 1

If ComboBox1.Value = "" Then
ComboBox1.BackColor = vbRed
MsgBox "Veuillez renseigner le statut du client pour enregistrer", vbOKOnly + vbCritical, "Champs Obligatoire"
Exit Sub
ElseIf TextBox8.BackColor = vbRed Then
Exit Sub
ElseIf TextBox10.BackColor = vbRed Then
Exit Sub
ElseIf TextBox11.BackColor = vbRed Then
Exit Sub
ElseIf TextBox13.BackColor = vbRed Then
Exit Sub
End If

'STATUT
Range("A" & L).Value = ComboBox1

'SECTEUR ACTIVITÉ
Range("B" & L).Value = ComboBox2


'DÉTAIL ACTIVITE
Range("C" & L).Value = TextBox1

'SOCIÉTÉ
Range("D" & L).Value = TextBox2

'TITRE
If OptionButton1.Value = True Then
Range("E" & L).Value = OptionButton1.Caption
ElseIf OptionButton2.Value = True Then
Range("E" & L).Value = OptionButton2.Caption
ElseIf OptionButton3.Value = True Then
Range("E" & L).Value = OptionButton3.Caption
ElseIf OptionButton1.Value = False Or OptionButton2.Value = False Or OptionButton3.Value = False Then
Range("E" & L).Value = ""
End If

'NOM
Range("F" & L).Value = TextBox3

'PRENOM
Range("G" & L).Value = TextBox4

'FONCTION/SERVICE
Range("H" & L).Value = TextBox5

'ADRESSE 1
Range("I" & L).Value = TextBox6

'ADRESSE 2
Range("J" & L).Value = TextBox7

'CP
Range("K" & L).Value = TextBox8


'VILLE
Range("L" & L).Value = TextBox9

'MOBILE
Range("M" & L).Value = TextBox10


'FIXE
Range("N" & L).Value = TextBox11

'MAIL
Range("O" & L).Value = TextBox12

'FAX
Range("P" & L).Value = TextBox13

'COPIE DANS L'ONGLET "BDD ADRESSES MAILS"
Range("A1048576").End(xlUp).EntireRow.copy Destination:=Worksheets("BDD Adresses Mails").Range("A1048576").End(xlUp)(2)

End If


'EFFACE LES DONNEES DU FORMULAIRE

ComboBox1 = Clear
ComboBox2 = Clear
OptionButton1 = Unchecked
OptionButton2 = Unchecked
OptionButton3 = Unchecked
TextBox1 = ""
TextBox1.BackColor = vbWhite
TextBox2 = ""
TextBox2.BackColor = vbWhite
TextBox3 = ""
TextBox3.BackColor = vbWhite
TextBox4 = ""
TextBox4.BackColor = vbWhite
TextBox5 = ""
TextBox5.BackColor = vbWhite
TextBox6 = ""
TextBox6.BackColor = vbWhite
TextBox7 = ""
TextBox7.BackColor = vbWhite
TextBox8 = ""
TextBox8.BackColor = vbWhite
TextBox9 = ""
TextBox9.BackColor = vbWhite
TextBox10 = ""
TextBox10.BackColor = vbWhite
TextBox11 = ""
TextBox11.BackColor = vbWhite
TextBox12 = ""
TextBox12.BackColor = vbWhite
TextBox13 = ""
TextBox13.BackColor = vbWhite

End Sub


'EFFACE LES DONNEES DU FORMULAIRE
Private Sub del_Click()
ComboBox1 = Clear
ComboBox1.BackColor = vbWhite
ComboBox2 = Clear
OptionButton1 = Unchecked
OptionButton2 = Unchecked
OptionButton3 = Unchecked
TextBox1 = ""
TextBox1.BackColor = vbWhite
TextBox2 = ""
TextBox2.BackColor = vbWhite
TextBox3 = ""
TextBox3.BackColor = vbWhite
TextBox4 = ""
TextBox4.BackColor = vbWhite
TextBox5 = ""
TextBox5.BackColor = vbWhite
TextBox6 = ""
TextBox6.BackColor = vbWhite
TextBox7 = ""
TextBox7.BackColor = vbWhite
TextBox8 = ""
TextBox8.BackColor = vbWhite
TextBox9 = ""
TextBox9.BackColor = vbWhite
TextBox10 = ""
TextBox10.BackColor = vbWhite
TextBox11 = ""
TextBox11.BackColor = vbWhite
TextBox12 = ""
TextBox12.BackColor = vbWhite
TextBox13 = ""
TextBox13.BackColor = vbWhite
End Sub


Private Sub ComboBox1_Change()
If ComboBox1.Value = "Partenaires" Then
UsfContact.Show 0
Sheets("PARTENAIRES").Select
ElseIf ComboBox1.Value = "Invités" Then
UsfContact.Show 0
Sheets("INVITES").Select
ElseIf ComboBox1.Value = "Exposants" Then
UsfContact.Show 0
Sheets("EXPOSANTS").Select
ElseIf ComboBox1.Value = "Cavaliers" Then
UsfContact.Show 0
Sheets("CAVALIERS").Select
ElseIf ComboBox1.Value = "Bénévoles" Then
UsfContact.Show 0
Sheets("BENEVOLES").Select
ElseIf ComboBox1.Value = "Presse" Then
UsfContact.Show 0
Sheets("PRESSE").Select
ElseIf ComboBox1.Value = "Fournisseurs" Then
UsfContact.Show 0
Sheets("FOURNISSEURS").Select
ElseIf ComboBox1.Value = "Prospects" Then
UsfContact.Show 0
Sheets("PROSPECTS").Select
End If
End Sub

Private Sub ComboBox1_AfterUpdate()
If ComboBox1.Value = "" Then
ComboBox1.BackColor = vbRed
MsgBox "Veuillez renseigner le statut du client pour enregistrer", vbOKOnly + vbCritical, "Champs Obligatoire"
ElseIf ComboBox1.Value <> "" Then
ComboBox1.BackColor = vbWhite
End If
End Sub

Private Sub quitter_Click()
Unload Me
End Sub
Private Sub TextBox2_Change()
TextBox2.Text = UCase(TextBox2.Text)
End Sub

Private Sub TextBox3_Change()
TextBox3.Text = UCase(TextBox3.Text)
End Sub

Private Sub TextBox4_Change()
TextBox4.Text = UCase(TextBox4.Text)
End Sub
Private Sub UserForm_Activate()
trois_boutons Me
End Sub

Private Sub TextBox5_Change()
TextBox5.Text = UCase(TextBox5.Text)
End Sub

Private Sub TextBox9_Change()
TextBox9.Text = UCase(TextBox9.Text)
End Sub


Private Sub TextBox10_AfterUpdate()
TextBox10.MaxLength = 14
If TextBox10.TextLength = 14 Then
TextBox10.BackColor = vbWhite
ElseIf TextBox10.TextLength = 13 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 12 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 11 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 10 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 9 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 8 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 7 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 6 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 5 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 4 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 3 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 2 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
ElseIf TextBox10.TextLength = 1 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox10.BackColor = vbRed
Exit Sub
End If
End Sub

Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
TextBox10.MaxLength = 14
Valeur = Len(TextBox10)
If Valeur = 2 Or Valeur = 5 Or Valeur = 8 Or Valeur = 11 Then TextBox10 = TextBox10 & "."
'FORCE LE CHAMPS DE SAISIE A RECEVOIR UN CHIFFRE
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
   KeyAscii = 0
    MsgBox "Il faut saisir un Numéro de Téléphone", vbOKOnly + vbCritical, "Format du Numéro de Téléphone incorrect"
    Exit Sub
End If
End Sub

Private Sub TextBox11_AfterUpdate()
TextBox11.MaxLength = 14
If TextBox11.TextLength = 14 Then
TextBox11.BackColor = vbWhite
ElseIf TextBox11.TextLength = 13 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 12 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 11 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 10 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 9 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 8 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 7 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 6 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 5 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 4 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 3 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 2 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
ElseIf TextBox11.TextLength = 1 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox11.BackColor = vbRed
Exit Sub
End If
End Sub

Private Sub TextBox11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
TextBox11.MaxLength = 14
Valeur = Len(TextBox11)
If Valeur = 2 Or Valeur = 5 Or Valeur = 8 Or Valeur = 11 Then TextBox11 = TextBox11 & "."
'FORCE LE CHAMPS DE SAISIE A RECEVOIR UN CHIFFRE
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
   KeyAscii = 0
    MsgBox "Il faut saisir un Numéro de Téléphone", vbOKOnly + vbCritical, "Format du Numéro de Téléphone incorrect"
    Exit Sub
End If
End Sub

Private Sub TextBox13_AfterUpdate()
TextBox13.MaxLength = 14
If TextBox13.TextLength = 14 Then
TextBox13.BackColor = vbWhite
ElseIf TextBox13.TextLength = 13 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 12 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 11 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 10 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 9 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 8 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 7 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 6 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 5 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 4 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 3 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 2 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
ElseIf TextBox13.TextLength = 1 Then
MsgBox "Le Numéro de Téléphone que vous avez saisi est incorrect" & vbNewLine & "Exemple : 08.25.95.49.85 ", vbOKOnly + vbCritical, "Numéro de Téléphone incorrect"
TextBox13.BackColor = vbRed
Exit Sub
End If
End Sub

Private Sub TextBox13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
TextBox13.MaxLength = 14
Valeur = Len(TextBox13)
If Valeur = 2 Or Valeur = 5 Or Valeur = 8 Or Valeur = 11 Then TextBox13 = TextBox13 & "."
'FORCE LE CHAMPS DE SAISIE A RECEVOIR UN CHIFFRE
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
   KeyAscii = 0
    MsgBox "Il faut saisir un Numéro de Téléphone", vbOKOnly + vbCritical, "Format du Numéro de Téléphone incorrect"
    Exit Sub
End If
End Sub


Private Sub TextBox8_AfterUpdate()
'MAXIMUM 5 CARACTERES
TextBox8.MaxLength = 5
'MESSAGE D'ERREUR SI MOINS DE 5 CARACTERES SAISIS
If TextBox8.TextLength = 5 Then
TextBox8.BackColor = vbWhite
ElseIf TextBox8.TextLength = 4 Then
MsgBox "Le Code Postal que vous avez saisi est incorrect" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
TextBox8.BackColor = vbRed
Exit Sub
ElseIf TextBox8.TextLength = 3 Then
MsgBox "Le Code Postal que vous avez saisi est incorrect" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
TextBox8.BackColor = vbRed
Exit Sub
'ElseIf TextBox7.TextLength = 2 Then
'MsgBox "Le format que vous avez saisi est incorrect", vbOKOnly + vbCritical, "Code Postal incorrect"
ElseIf TextBox8.TextLength = 1 Then
MsgBox "Le Code Postal que vous avez saisi est incorrect" & vbNewLine & "Exemple : 33 - 33000 ", vbOKOnly + vbCritical, "Code Postal incorrect"
TextBox8.BackColor = vbRed
Exit Sub
End If
End Sub

Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'MAXIMUM 5 CARACTERES
TextBox8.MaxLength = 5
'FORCE LE CHAMPS DE SAISIE A RECEVOIR UN CHIFFRE
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
   KeyAscii = 0
    MsgBox "Il faut saisir un Code Postal", vbOKOnly + vbCritical, "Format du Code Postal incorrect"
    Exit Sub
End If
End Sub

Private Sub UserForm_Click()

End Sub
