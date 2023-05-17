VERSION 5.00
Begin VB.Form frmEditor_Spawn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Spawn Editor"
   ClientHeight    =   4995
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   4815
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.HScrollBar scrlPokeBuff 
         Height          =   255
         Left            =   1800
         Max             =   100
         TabIndex        =   31
         Top             =   4320
         Width           =   3255
      End
      Begin VB.CheckBox chkNoExp 
         Caption         =   "Cannot Give Exp?"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   3840
         Width           =   1695
      End
      Begin VB.ComboBox cmbPokemonNum 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   3495
      End
      Begin VB.CheckBox chkCanCatch 
         Caption         =   "Cannot Catch?"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtRarity 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Text            =   "0"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtSpawnMax 
         Height          =   285
         Left            =   3360
         TabIndex        =   21
         Text            =   "0"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtSpawnMin 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Location:"
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   5055
         Begin VB.CheckBox chkRandomXY 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkRandomMap 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtY 
            Height          =   285
            Left            =   2880
            TabIndex        =   17
            Text            =   "0"
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   2880
            TabIndex        =   15
            Text            =   "0"
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   2880
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Y:"
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "X:"
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Map:"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtRespawn 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "0"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtMaxLevel 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Text            =   "0"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMinLevel 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "0"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblBuff 
         Caption         =   "Pokemon Buff: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4320
         Width           =   4815
      End
      Begin VB.Label Label11 
         Caption         =   "Common  [ 0 ~~ 100000 ] Rare"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Rarity"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Spawn Time"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "ms"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Respawn:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Level Range:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPokemon 
         Caption         =   "Pokemon: "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.ListBox lstMapPokemon 
         Height          =   4350
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmEditor_Spawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCanCatch_Click()
    Spawn(EditorIndex).CanCatch = chkCanCatch.value
End Sub

Private Sub chkNoExp_Click()
    Spawn(EditorIndex).NoExp = chkNoExp.value
End Sub

Private Sub chkRandomMap_Click()
    Spawn(EditorIndex).randomMap = chkRandomMap.value
End Sub

Private Sub chkRandomXY_Click()
    Spawn(EditorIndex).randomXY = chkRandomXY.value
End Sub

Private Sub cmbPokemonNum_Click()
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstMapPokemon.ListIndex
    lstMapPokemon.RemoveItem EditorIndex - 1
    Spawn(EditorIndex).PokeNum = cmbPokemonNum.ListIndex
    If cmbPokemonNum.ListIndex > 0 Then
        lstMapPokemon.AddItem EditorIndex & ": " & Trim$(Pokemon(cmbPokemonNum.ListIndex).Name), EditorIndex - 1
    Else
        lstMapPokemon.AddItem EditorIndex & ": ", EditorIndex - 1
    End If
    lstMapPokemon.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub lstMapPokemon_Click()
    SpawnEditorLoadIndex lstMapPokemon.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestSpawn
    End If
    CloseSpawnEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        If SpawnChange(i) Then
            SendSaveSpawn i
            SpawnChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseSpawnEditor
End Sub

Private Sub scrlPokeBuff_Change()
    lblBuff.Caption = "Pokemon Buff: " & scrlPokeBuff.value
    Spawn(EditorIndex).PokeBuff = scrlPokeBuff.value
    EditorChange = True
End Sub

Private Sub txtMap_Change()
    If IsNumeric(txtMap.Text) Then
        Spawn(EditorIndex).MapNum = Val(txtMap.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMaxLevel_Change()
    If IsNumeric(txtMaxLevel.Text) Then
        If Val(txtMaxLevel.Text) <= 0 Then txtMaxLevel.Text = 0
        If Val(txtMaxLevel.Text) >= 100 Then txtMaxLevel.Text = 100
        Spawn(EditorIndex).MaxLevel = Val(txtMaxLevel.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMinLevel_Change()
    If IsNumeric(txtMinLevel.Text) Then
        If Val(txtMinLevel.Text) <= 0 Then txtMinLevel.Text = 0
        If Val(txtMinLevel.Text) >= 100 Then txtMinLevel.Text = 100
        Spawn(EditorIndex).MinLevel = Val(txtMinLevel.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRarity_Change()
    If IsNumeric(txtRarity.Text) Then
        Spawn(EditorIndex).Rarity = Val(txtRarity.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRespawn_Change()
    If IsNumeric(txtRespawn.Text) Then
        Spawn(EditorIndex).Respawn = Val(txtRespawn.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtSpawnMax_Change()
    If IsNumeric(txtSpawnMax.Text) Then
        Spawn(EditorIndex).SpawnTimeMax = Val(txtSpawnMax.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtSpawnMin_Change()
    If IsNumeric(txtSpawnMin.Text) Then
        Spawn(EditorIndex).SpawnTimeMin = Val(txtSpawnMin.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtX_Change()
    If IsNumeric(txtX.Text) Then
        Spawn(EditorIndex).MapX = Val(txtX.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtY_Change()
    If IsNumeric(txtY.Text) Then
        Spawn(EditorIndex).MapY = Val(txtY.Text)
        EditorChange = True
    End If
End Sub
