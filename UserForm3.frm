VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Recherche d'un matériel"
   ClientHeight    =   7575
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9915
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Ws As Worksheet
Dim Sd As Worksheet
Public Function FeuilleExiste(sNomFeuille As String) As Boolean 'fonction pour savoir si la feuille existe déja
    On Error GoTo Err_FeuilleExiste
    FeuilleExiste = False
    FeuilleExiste = Not ActiveWorkbook.Worksheets(sNomFeuille) Is Nothing
Err_FeuilleExiste:
End Function
Private Sub ComboBox2_Change()
Dim x As Integer
Dim y As String
Dim Z As String
y = ComboBox1
Z = ComboBox2
For x = 2 To Ws.Range("B" & Rows.Count).End(xlUp).Row 'ici on determine si on a besoin d'un troisième champ de recher
    If Ws.Range("B" & x) = y Then
        If Ws.Range("C" & x) = Z Then
            If y = "N/A" Then
                ComboBox3.Visible = True 'si oui on rend la combobox3 visisble
                Label11.Visible = True
                With Me.ComboBox3
                    .AddItem Ws.Range("F" & x)
                End With
            Else
                ComboBox9 = Range("A" & x).Value 'si non ont rempli les textboxs
                ComboBox8 = Range("B" & x).Value
                ComboBox7 = Range("C" & x).Value
                TextBox4 = Range("D" & x).Value
                TextBox5 = Range("E" & x).Value
                TextBox6 = Range("F" & x).Value
                ComboBox5 = Range("G" & x).Value
                ComboBox6 = Range("H" & x).Value
            End If
        End If
    End If
Next
End Sub

Private Sub ComboBox3_Change() 'Initialisation de la combobox 3
Dim w As Integer
Dim x As String
Dim y As String
Dim Z As String
x = ComboBox3
y = ComboBox1
Z = ComboBox2
For w = 2 To Ws.Range("B" & Rows.Count).End(xlUp).Row
    If Ws.Range("B" & w) = y Then
        If Ws.Range("C" & w) = Z Then
            If Ws.Range("F" & w) = x Then
                ComboBox9 = Range("A" & w).Value
                ComboBox8 = Range("B" & w).Value
                ComboBox7 = Range("C" & w).Value
                TextBox4 = Range("D" & w).Value
                TextBox5 = Range("E" & w).Value
                TextBox6 = Range("F" & w).Value
                ComboBox5 = Range("G" & w).Value
                ComboBox6 = Range("H" & w).Value
            End If
        End If
    End If
Next
End Sub

Private Sub ComboBox4_Change()
Dim w As Integer
Dim y As String
y = ComboBox4
ComboBox9 = "" 'reset des textbox
ComboBox8 = ""
ComboBox7 = ""
TextBox4 = ""
TextBox5 = ""
TextBox6 = ""
ComboBox5 = ""
ComboBox6 = ""
ComboBox1 = ""
ComboBox2 = ""
ComboBox3 = ""
For w = 2 To Ws.Range("F" & Rows.Count).End(xlUp).Row 'les champ prennent leurs nouvelles valeurs
    If Ws.Range("F" & w) = y Then
        ComboBox9 = Range("A" & w).Value
        ComboBox8 = Range("B" & w).Value
        ComboBox7 = Range("C" & w).Value
        TextBox4 = Range("D" & w).Value
        TextBox5 = Range("E" & w).Value
        TextBox6 = Range("F" & w).Value
        ComboBox5 = Range("G" & w).Value
        ComboBox6 = Range("H" & w).Value
    End If
Next
End Sub


Private Sub CommandButton4_Click()
Dim nom As String
Dim L As Integer
Dim date_du_jour As Date
date_du_jour = Format(Now, "dd/mm/yyyy")
nom = InputBox("Nom de la feuille où sauvegarder", "Confirmation", "Enregistrement du " & date_du_jour) 'nom du fichier en fonction de la date du jour
If nom = "" Then Exit Sub
If FeuilleExiste(nom) Then 'appel de la fonction qui détermine si le fichier existe
Sheets(nom).Select
Else
Sheets.Add.Name = (nom) 'création de la nouvelle fiche
        Range("A" & 1).Value = "Plateforme"
        Range("B" & 1).Value = "Numéro de position"
        Range("C" & 1).Value = "Matériel"
        Range("D" & 1).Value = "Marque"
        Range("E" & 1).Value = "Modèle"
        Range("F" & 1).Value = "N° de série"
        Range("G" & 1).Value = "Stand"
        Range("H" & 1).Value = "Etat"
End If
    L = Sheets(nom).Range("A" & Rows.Count).End(xlUp).Row + 1
        Range("A" & L).Value = ComboBox9
        Range("B" & L).Value = ComboBox8
        Range("C" & L).Value = ComboBox7
        Range("D" & L).Value = TextBox4
        Range("E" & L).Value = TextBox5
        Range("F" & L).Value = TextBox6
        Range("G" & L).Value = ComboBox5
        Range("H" & L).Value = ComboBox6


End Sub

Private Sub CommandButton5_Click()
Dim j As Integer
If MsgBox("Confirmez-vous l'enregistrement ?", vbYesNo, "Demande de confirmation") = vbYes Then
For j = 2 To Ws.Range("A" & Rows.Count).End(xlUp).Row
        If Ws.Range("B" & j).Value = ComboBox1 And Ws.Range("C" & j).Value = ComboBox2 And Ws.Range("F" & j).Value = ComboBox3 Then
            Ws.Range("A" & j).Value = ComboBox9
            Ws.Range("B" & j).Value = ComboBox8
            Ws.Range("C" & j).Value = ComboBox7
            Ws.Range("D" & j).Value = TextBox4
            Ws.Range("E" & j).Value = TextBox5
            Ws.Range("F" & j).Value = TextBox6
            Ws.Range("G" & j).Value = ComboBox5
            Ws.Range("H" & j).Value = ComboBox6
        ElseIf Ws.Range("F" & j).Value = ComboBox4 Then
            Ws.Range("A" & j).Value = ComboBox9
            Ws.Range("B" & j).Value = ComboBox8
            Ws.Range("C" & j).Value = ComboBox7
            Ws.Range("D" & j).Value = TextBox4
            Ws.Range("E" & j).Value = TextBox5
            Ws.Range("F" & j).Value = TextBox6
            Ws.Range("G" & j).Value = ComboBox5
            Ws.Range("H" & j).Value = ComboBox6
        ElseIf Ws.Range("B" & j).Value = ComboBox1 And Ws.Range("C" & j).Value = ComboBox2 Then
            Ws.Range("A" & j).Value = ComboBox9
            Ws.Range("B" & j).Value = ComboBox8
            Ws.Range("C" & j).Value = ComboBox7
            Ws.Range("D" & j).Value = TextBox4
            Ws.Range("E" & j).Value = TextBox5
            Ws.Range("F" & j).Value = TextBox6
            Ws.Range("G" & j).Value = ComboBox5
            Ws.Range("H" & j).Value = ComboBox6
        End If
Next
End If
ComboBox9 = "" 'reset des textbox
ComboBox8 = ""
ComboBox7 = ""
TextBox4 = ""
TextBox5 = ""
TextBox6 = ""
ComboBox5 = ""
ComboBox6 = ""
ComboBox1 = ""
ComboBox2 = ""
ComboBox3 = ""
End Sub

Private Sub CommandButton6_Click()
Dim y As String
y = MsgBox("- Pour rechercher un équipement, remplir toutes les collones en haut a gauche ou celles en haut à droite" + vbCrLf + "- Pour enregistrer le (ou les) matériel(s) dans une autre fiche, une fois selectionné(s) cliquer sur enregistrer et choisiser la fiche où enregistrer" + vbCrLf + "- Pour modfifier l'élement choisie, remplacer les champs voulu dans la partie basse de la fenêtre et cliquer sur enregistrer", vbOKOnly, "Aide")
End Sub

Private Sub UserForm_Initialize()
Dim j As Long
Set Ws = Sheets("Sheet1")
Set Sd = Sheets("Data")
Dim Cell As Range
Dim i As Integer
Dim x As Integer
Dim Un As New Collection
 
    On Error Resume Next
        'Recherche les doublons dans la plage A
        For Each Cell In Range("B1:B36656")
            'Utilise la propriété "Key" des collections qui
            'n'acceptent que des valeurs uniques.
            Un.Add Cell, CStr(Cell)
        Next Cell
    On Error GoTo 0
 
    For i = 2 To Un.Count
        'Afiche le résultat sans doublon dans la colonne B
        With Me.ComboBox1
            .AddItem Un.Item(i)
        End With
       
    Next i
 With Me.ComboBox4 'initialisation de la combobox4
    For j = 2 To Ws.Range("F" & Rows.Count).End(xlUp).Row
        If Ws.Range("F" & j).Value <> "N/A" And Ws.Range("F" & j).Value <> "?" Then
            .AddItem Ws.Range("F" & j)
        End If
    Next
 End With
ComboBox5.List() = Array(" ", "sur mât", " N/A ", "sur pied")
ComboBox6.List() = Array("Neuf", "Moyen", "Bon", "HS", "à réformer")

For x = 2 To Sd.Range("B" & Rows.Count).End(xlUp).Row
    With Me.ComboBox8
        .AddItem Sd.Range("B" & x)
    End With
Next
For x = 2 To Sd.Range("C" & Rows.Count).End(xlUp).Row
    With Me.ComboBox7
        .AddItem Sd.Range("C" & x)
    End With
Next
For x = 2 To Sd.Range("A" & Rows.Count).End(xlUp).Row
    With Me.ComboBox9
        .AddItem Sd.Range("A" & x)
    End With
Next
ComboBox3.Visible = False
Label11.Visible = False
End Sub


Private Sub CommandButton2_Click()
Unload Me 'bouton retour
UserForm2.Show vbModeless
End Sub

Private Sub CommandButton3_Click()
Unload Me 'bouton quitter

End Sub

Private Sub Combobox1_change()
Dim i As Integer
Dim msg As String
Sheets("Sheet1").Select
ComboBox3.Visible = False 'Nétoyage de la fenêtre si l'tilisateur lance une autre recherche
Label11.Visible = False
ComboBox2.Clear
ComboBox3.Clear
ComboBox4 = ""
ComboBox9 = ""
ComboBox8 = ""
ComboBox7 = ""
TextBox4 = ""
TextBox5 = ""
TextBox6 = ""
ComboBox5 = ""
ComboBox6 = ""
If Me.ComboBox1.ListIndex = "" Then Exit Sub
    msg = ComboBox1
    With Me.ComboBox2 'initialisation du champ de recherche "Equipement"
        For i = 2 To Ws.Range("B" & Rows.Count).End(xlUp).Row
            If Ws.Range("B" & i) = msg Then
                .AddItem Ws.Range("C" & i)
            End If
        Next i
    End With

End Sub

