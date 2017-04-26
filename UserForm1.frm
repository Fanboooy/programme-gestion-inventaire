VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Ajout d'un matériel"
   ClientHeight    =   7575
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9945
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Sd As Worksheet
Dim Ws As Worksheet

Private Sub CommandButton3_Click()
Unload Me 'bouton retour
UserForm2.Show
End Sub


Private Sub CommandButton4_Click()
Dim y As String
y = MsgBox("- Remplisser tout les champs et appuyer sur le bouton Ajouter le nouveau matériel sera ajouter a la fin du dossier", vbOKOnly, "Aide")
End Sub

'Pour initialiser les combobox 1, 2, 3

Private Sub UserForm_Initialize()
Set Ws = Sheets("Sheet1") 'Correspond au nom de votre fiche excel où ce trouve l'inventaire
Set Sd = Sheets("Data")
Dim x As Integer
    'Pour la liste déroulante stand

ComboBox1.List() = Array(" ", "sur mât", " N/A ", "sur pied")

    
    'Pour la liste déroulante état

ComboBox2.List() = Array("Neuf", "Moyen", "Bon", "HS", "à réformer")

For x = 2 To Sd.Range("B" & Rows.Count).End(xlUp).Row
    With Me.ComboBox5
        .AddItem Sd.Range("B" & x)
    End With
Next
For x = 2 To Sd.Range("C" & Rows.Count).End(xlUp).Row
    With Me.ComboBox4
        .AddItem Sd.Range("C" & x)
    End With
Next
For x = 2 To Sd.Range("A" & Rows.Count).End(xlUp).Row
    With Me.ComboBox3
        .AddItem Sd.Range("A" & x)
    End With
Next
End Sub


Private Sub CommandButton1_Click() 'bouton de confirmation et affichage des données dans les cases corespondante
Dim L As Integer
    If MsgBox("Confirmez-vous l'enregistrement ?", vbYesNo, "Demande de confirmation") = vbYes Then 'confirmation et insertion des données dans le tableau
    L = Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row + 1
        Range("B" & L).Value = ComboBox5
        Range("C" & L).Value = ComboBox4
        Range("D" & L).Value = TextBox4
        Range("E" & L).Value = TextBox5
        Range("F" & L).Value = TextBox6
        Range("G" & L).Value = ComboBox1
        Range("H" & L).Value = ComboBox2
        Range("A" & L).Value = ComboBox3
        ComboBox4 = ""    'reset des celules
        ComboBox5 = ""
        TextBox4 = ""
        TextBox5 = ""
        TextBox6 = ""
        ComboBox1 = ""
        ComboBox2 = ""
        ComboBox3 = ""
    End If
End Sub


Private Sub CommandButton2_Click() 'bouton quitter
    Unload Me
End Sub
