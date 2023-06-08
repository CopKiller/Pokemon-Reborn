VERSION 5.00
Begin VB.Form frmEditor_Npc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor"
   ClientHeight    =   7095
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   6975
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtRewardExp 
         Height          =   285
         Left            =   4080
         TabIndex        =   31
         Text            =   "0"
         Top             =   6000
         Width           =   1695
      End
      Begin VB.HScrollBar scrlWinConvo 
         Height          =   255
         Left            =   3000
         Max             =   0
         TabIndex        =   30
         Top             =   6360
         Width           =   2775
      End
      Begin VB.TextBox txtReward 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Text            =   "0"
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pokemon"
         Height          =   3375
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   5535
         Begin VB.ComboBox cmbMoveset 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2880
            Width           =   1815
         End
         Begin VB.CommandButton cmdFindMove 
            Caption         =   "Find"
            Height          =   255
            Left            =   4560
            TabIndex        =   24
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtFindMoveset 
            Height          =   285
            Left            =   2880
            TabIndex        =   23
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Text            =   "0"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.ListBox lstMoveset 
            Height          =   840
            Left            =   2880
            TabIndex        =   20
            Top             =   1560
            Width           =   2415
         End
         Begin VB.ComboBox cmbPokeNum 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1680
            Width           =   2535
         End
         Begin VB.ListBox lstPokemon 
            Height          =   1035
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label6 
            Caption         =   "Move"
            Height          =   255
            Left            =   2880
            TabIndex        =   26
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Level:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Pokemon"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1440
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmbNpcType 
         Height          =   315
         ItemData        =   "frmEditor_Npc.frx":0000
         Left            =   1200
         List            =   "frmEditor_Npc.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   4575
      End
      Begin VB.HScrollBar scrlConvo 
         Height          =   255
         Left            =   3000
         Max             =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   315
         ItemData        =   "frmEditor_Npc.frx":003B
         Left            =   1200
         List            =   "frmEditor_Npc.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   7
         Top             =   720
         Width           =   3855
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   5160
         ScaleHeight     =   66.065
         ScaleMode       =   0  'User
         ScaleWidth      =   34.133
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblWinConvo 
         Caption         =   "Win Convo:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   6360
         Width           =   5535
      End
      Begin VB.Label Label7 
         Caption         =   "Reward Money / EXP:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblConvo 
         Caption         =   "Conversation: None"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "Behaviour:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   6300
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2655
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
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit(Esc)"
      End
   End
End
Attribute VB_Name = "frmEditor_Npc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PokemonIndex As Long
Private MoveIndex As Long

Private Sub cmbBehaviour_Click()
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    EditorChange = True
End Sub

Private Sub cmbMoveset_Click()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    If MoveIndex = 0 Then Exit Sub
    tmpIndex = lstMoveset.ListIndex
    lstMoveset.RemoveItem MoveIndex - 1
    Npc(EditorIndex).PokemonMoveset(PokemonIndex, MoveIndex) = cmbMoveset.ListIndex
    If cmbMoveset.ListIndex > 0 Then
        lstMoveset.AddItem MoveIndex & ": " & Trim$(PokemonMove(cmbMoveset.ListIndex).Name), MoveIndex - 1
    Else
        lstMoveset.AddItem MoveIndex & ": None", MoveIndex - 1
    End If
    lstMoveset.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbPokeNum_Click()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    tmpIndex = lstPokemon.ListIndex
    lstPokemon.RemoveItem PokemonIndex - 1
    Npc(EditorIndex).PokemonNum(PokemonIndex) = cmbPokeNum.ListIndex
    If cmbPokeNum.ListIndex > 0 Then
        lstPokemon.AddItem PokemonIndex & ": " & Trim$(Pokemon(cmbPokeNum.ListIndex).Name) & " Lv: " & Npc(EditorIndex).PokemonLevel(PokemonIndex), PokemonIndex - 1
    Else
        lstPokemon.AddItem PokemonIndex & ": None", PokemonIndex - 1
    End If
    lstPokemon.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmdFindMove_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtFindMoveset.Text)) > 0 Then
        FindChar = Trim$(txtFindMoveset.Text)
        clBound = 0
        cuBound = MAX_POKEMON_MOVE
        
        For i = clBound To cuBound
            If cmbMoveset.List(i) <> "None" Then
                ComboText = Trim$(cmbMoveset.List(i))
                indexString = i & ": "
                stringLength = Len(ComboText) - Len(indexString)
                If stringLength >= 0 Then
                    ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                    If LCase(ComboText) = LCase(FindChar) Then
                        cmbMoveset.ListIndex = i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

Private Sub cmdIndexSearch_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtIndexSearch.Text)) > 0 Then
        FindChar = Trim$(txtIndexSearch.Text)
        clBound = 1
        cuBound = MAX_NPC
        
        For i = clBound To cuBound
            ComboText = Trim$(lstIndex.List(i - 1))
            indexString = i & ": "
            stringLength = Len(ComboText) - Len(indexString)
            If stringLength >= 0 Then
                ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                If LCase(ComboText) = LCase(FindChar) Then
                    lstIndex.ListIndex = (i - 1)
                    Exit Sub
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseNpcEditor
    End If
End Sub

Private Sub Form_Load()
    scrlSprite.max = Count_Character
    txtName.MaxLength = NAME_LENGTH
    scrlConvo.max = MAX_CONVERSATION
    scrlWinConvo.max = MAX_CONVERSATION
End Sub

Private Sub lstIndex_Click()
    NpcEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub lstMoveset_Click()
    MoveIndex = lstMoveset.ListIndex + 1
    
    If PokemonIndex <= 0 Then Exit Sub
    If MoveIndex <= 0 Then Exit Sub
    
    cmbMoveset.ListIndex = Npc(EditorIndex).PokemonMoveset(PokemonIndex, MoveIndex)
End Sub

Private Sub lstPokemon_Click()
Dim X As Byte

    PokemonIndex = lstPokemon.ListIndex + 1
    
    If PokemonIndex <= 0 Then Exit Sub
    
    cmbPokeNum.ListIndex = Npc(EditorIndex).PokemonNum(PokemonIndex)
    txtLevel.Text = Npc(EditorIndex).PokemonLevel(PokemonIndex)
    lstMoveset.Clear
    For X = 1 To MAX_MOVESET
        If Npc(EditorIndex).PokemonMoveset(PokemonIndex, X) > 0 Then
            lstMoveset.AddItem X & ": " & Trim$(PokemonMove(Npc(EditorIndex).PokemonMoveset(PokemonIndex, X)).Name)
        Else
            lstMoveset.AddItem X & ": None"
        End If
    Next
    lstMoveset.ListIndex = 0
    cmbMoveset.ListIndex = Npc(EditorIndex).PokemonMoveset(PokemonIndex, 1)
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestNpc
    End If
    CloseNpcEditor
End Sub

Private Sub mnuExit_Click()
    CloseNpcEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_NPC
        If NpcChange(i) Then
            SendSaveNpc i
            NpcChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseNpcEditor
End Sub

Private Sub scrlConvo_Change()
    If scrlConvo.value > 0 Then
        lblConvo.Caption = "Conversation: " & Trim$(Conversation(scrlConvo.value).Name)
    Else
        lblConvo.Caption = "Conversation: None"
    End If
    Npc(EditorIndex).Convo = scrlConvo.value
    EditorChange = True
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Npc(EditorIndex).Sprite = scrlSprite.value
    EditorChange = True
End Sub

Private Sub scrlWinConvo_Change()
    If scrlWinConvo.value > 0 Then
        lblWinConvo.Caption = "Win Convo: " & Trim$(Conversation(scrlWinConvo.value).Name)
    Else
        lblWinConvo.Caption = "Win Convo: None"
    End If
    Npc(EditorIndex).WinEvent = scrlWinConvo.value
    EditorChange = True
End Sub

Private Sub txtLevel_Change()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    If Not IsNumeric(txtLevel.Text) Then Exit Sub
    tmpIndex = lstPokemon.ListIndex
    lstPokemon.RemoveItem PokemonIndex - 1
    Npc(EditorIndex).PokemonLevel(PokemonIndex) = Val(txtLevel.Text)
    If Npc(EditorIndex).PokemonNum(PokemonIndex) > 0 Then
        lstPokemon.AddItem PokemonIndex & ": " & Trim$(Pokemon(Npc(EditorIndex).PokemonNum(PokemonIndex)).Name) & " Lv: " & Trim$(txtLevel.Text), PokemonIndex - 1
    Else
        lstPokemon.AddItem PokemonIndex & ": None", PokemonIndex - 1
    End If
    lstPokemon.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtReward_Change()
    If IsNumeric(txtReward.Text) Then
        Npc(EditorIndex).Reward = Val(txtReward.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRewardExp_Change()
    If IsNumeric(txtRewardExp.Text) Then
        Npc(EditorIndex).RewardExp = Val(txtRewardExp.Text)
        EditorChange = True
    End If
End Sub
