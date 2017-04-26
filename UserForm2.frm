VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Menu"
   ClientHeight    =   5970
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   7530
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Ws As Worksheet
Dim Sd As Worksheet
Public Function FeuilleExiste(sNomFeuille As String) As Boolean 'fonction pour savoir si la feuille existe d�ja
    On Error GoTo Err_FeuilleExiste
    FeuilleExiste = False
    FeuilleExiste = Not ActiveWorkbook.Worksheets(sNomFeuille) Is Nothing
Err_FeuilleExiste:
End Function
Private Sub CommandButton1_Click()
Unload Me 'bouton d'acc�s au formulaire pour remplir le dossier
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()

Unload Me 'bouton pour acceder au formulaire de recherche
UserForm3.Show vbModeless
End Sub

Private Sub CommandButton3_Click()
Unload Me 'bouton quitter
End Sub


Private Sub CommandButton4_Click()
Dim y As String
y = MsgBox("- Une fiche data a �t� cr�er si vous n'en aviez pas d�ja ajouter une vous pouvez la supprimer si vous souhaiter en utiliser une autre" + vbCrLf + "-Cliquer sur entrer un nouveau mat�riel pour ajouter des donn�es a votre inventaire" + vbCrLf + "-Cliquer sur rechercher pour retrouver un mat�riel, le sauvegarder dans une autre fiche ou le modifier", vbOKOnly, "Aide")
End Sub

Private Sub UserForm_Initialize()
Dim nom As String
Dim Cellt As Range
Dim Cellc As Range
Dim Cellq As Range
Dim t As Integer
Dim c As Integer
Dim q As Integer
Dim Unt As New Collection
Dim unc As New Collection
Dim unq As New Collection
nom = "Data"
If FeuilleExiste(nom) Then 'appel de la fonction qui d�termine si le fichier existe
Else
    Sheets.Add.Name = (nom) 'cr�ation de la fiche data
            Range("A" & 1).Value = "Plateforme"
            Range("B" & 1).Value = "Num�ro de position"
            Range("C" & 1).Value = "Mat�riel"
            Range("D" & 1).Value = "Marque"
            Range("E" & 1).Value = "Mod�le"
            Range("F" & 1).Value = "N� de s�rie"
            Range("G" & 1).Value = "Stand"
            Range("G" & 2).Value = "sur m�t"
            Range("G" & 3).Value = " N/A "
            Range("G" & 4).Value = "sur pied"
            Range("H" & 1).Value = "Etat"
            Range("H" & 2).Value = "Neuf"
            Range("H" & 3).Value = "Moyen"
            Range("H" & 4).Value = "Bon"
            Range("H" & 5).Value = "HS"
            Range("H" & 6).Value = "� r�former"
End If
Set Sd = Sheets("Data") 'Fiche o� se situ les donn�es
Set Ws = Sheets("Sheet1")
On Error Resume Next
     'Recherche les doublons dans la plage A
    For Each Cellt In Ws.Range("A1:A36656")
        'Utilise la propri�t� "Key" des collections qui
        'n'acceptent que des valeurs uniques.
        Unt.Add Cellt, CStr(Cellt)
    Next Cellt
On Error GoTo 0
For t = 2 To Unt.Count
    Sd.Range("A" & t).Value = Unt.Item(t)
Next t
'------------------------------
On Error Resume Next
     'Recherche les doublons dans la plage A
    For Each Cellq In Ws.Range("C1:C36656")
        'Utilise la propri�t� "Key" des collections qui
        'n'acceptent que des valeurs uniques.
        unq.Add Cellq, CStr(Cellq)
    Next Cellq
On Error GoTo 0
For q = 2 To unq.Count
    Sd.Range("C" & q).Value = unq.Item(q)
Next q

'-------------------------
On Error Resume Next
    For Each Cellc In Ws.Range("B1:B36656")
        unc.Add Cellc, CStr(Cellc)
    Next Cellc
On Error GoTo 0
For c = 2 To unc.Count
     Sd.Range("B" & c).Value = unc.Item(c)
Next c
Ws.Select
End Sub
