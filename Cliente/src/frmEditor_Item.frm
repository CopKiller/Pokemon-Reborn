VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   4365
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraTMHM 
      Caption         =   "TM/HM"
      Height          =   1215
      Left            =   3360
      TabIndex        =   33
      Top             =   1920
      Width           =   4695
      Begin VB.ComboBox cmbMoveList 
         Height          =   315
         ItemData        =   "frmEditor_Item.frx":0000
         Left            =   1440
         List            =   "frmEditor_Item.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox chkTakeItem 
         Caption         =   "Take Item?"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Move List"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   3375
         ItemData        =   "frmEditor_Item.frx":0004
         Left            =   120
         List            =   "frmEditor_Item.frx":0006
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   4215
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Frame fraMedicine 
         Caption         =   "Medicine"
         Height          =   1455
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   4695
         Begin VB.CheckBox chkLevelUp 
            Caption         =   "Level Up"
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Text            =   "0"
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cmbMedicineType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":0008
            Left            =   1440
            List            =   "frmEditor_Item.frx":0024
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label7 
            Caption         =   "Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Medicine Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   4320
         TabIndex        =   18
         Text            =   "0"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame fraPokeball 
         Caption         =   "Pokeball Properties"
         Height          =   1695
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CheckBox chkAutoCatch 
            Caption         =   "Auto Catch?"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar scrlBallSprite 
            Height          =   255
            Left            =   1680
            Max             =   15
            TabIndex        =   16
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtCatchRate 
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Text            =   "0"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblBallSprite 
            Caption         =   "Ball Sprite: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Catch Rate"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Item.frx":0082
         Left            =   1200
         List            =   "frmEditor_Item.frx":009E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkStock 
         Caption         =   "Stock?"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5280
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin VB.Frame fraKeyItem 
         Caption         =   "Key Item Properties"
         Height          =   1215
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbKeyItemType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":00ED
            Left            =   2280
            List            =   "frmEditor_Item.frx":00F7
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   360
            Width           =   3015
         End
         Begin VB.HScrollBar scrlSpriteType 
            Height          =   255
            Left            =   2280
            Max             =   3
            TabIndex        =   29
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label8 
            Caption         =   "Key Item Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblSpriteType 
            Caption         =   "Sprite Type: None"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   5055
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Description"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Price:"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1455
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
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutoCatch_Click()
    Item(EditorIndex).Data3 = chkAutoCatch.value
    EditorChange = True
End Sub

Private Sub chkLevelUp_Click()
    Item(EditorIndex).Data3 = chkLevelUp.value
    EditorChange = True
End Sub

Private Sub chkStock_Click()
    Item(EditorIndex).Stock = chkStock.value
    EditorChange = True
End Sub

Private Sub chkTakeItem_Click()
    Item(EditorIndex).Data2 = chkTakeItem.value
    EditorChange = True
End Sub

Private Sub cmbKeyItemType_Click()
    Item(EditorIndex).Data1 = cmbKeyItemType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbMedicineType_Click()
    Item(EditorIndex).Data1 = cmbMedicineType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbMoveList_Click()
    Item(EditorIndex).Data1 = cmbMoveList.ListIndex
    EditorChange = True
End Sub

Private Sub cmbType_Click()
    Item(EditorIndex).Type = cmbType.ListIndex
    
    If Item(EditorIndex).Type = ItemTypeEnum.Pokeball Then
        fraPokeball.Visible = True
    Else
        fraPokeball.Visible = False
    End If
    
    If Item(EditorIndex).Type = ItemTypeEnum.Medicine Then
        fraMedicine.Visible = True
    Else
        fraMedicine.Visible = False
    End If
    
    If Item(EditorIndex).Type = ItemTypeEnum.keyItems Then
        fraKeyItem.Visible = True
    Else
        fraKeyItem.Visible = False
    End If
    
    If Item(EditorIndex).Type = ItemTypeEnum.TM_HM Then
        fraTMHM.Visible = True
    Else
        fraTMHM.Visible = False
    End If
    
    EditorChange = True
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
        cuBound = MAX_ITEM
        
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

Private Sub Form_Load()
    scrlSprite.max = Count_Item
    txtName.MaxLength = NAME_LENGTH
End Sub

Private Sub lstIndex_Click()
    ItemEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestItem
    End If
    CloseItemEditor
End Sub

Private Sub mnuExit_Click()
    CloseItemEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_ITEM
        If ItemChange(i) Then
            SendSaveItem i
            ItemChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseItemEditor
End Sub

Private Sub scrlBallSprite_Change()
    lblBallSprite.Caption = "Ball Sprite: " & scrlBallSprite.value
    Item(EditorIndex).Data2 = scrlBallSprite.value
    EditorChange = True
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Item(EditorIndex).Sprite = scrlSprite.value
    EditorChange = True
End Sub

Private Sub scrlSpriteType_Change()
    Select Case scrlSpriteType.value
        Case TEMP_SPRITE_GROUP_DIVE
            lblSpriteType.Caption = "Sprite Type: Dive"
        Case TEMP_SPRITE_GROUP_BIKE
            lblSpriteType.Caption = "Sprite Type: Bike"
        Case TEMP_SPRITE_GROUP_SURF
            lblSpriteType.Caption = "Sprite Type: Surf"
        Case Else
            lblSpriteType.Caption = "Sprite Type: None"
    End Select
    Item(EditorIndex).Data2 = scrlSpriteType.value
    EditorChange = True
End Sub

Private Sub txtCatchRate_Change()
    If IsNumeric(txtCatchRate.Text) Then
        Item(EditorIndex).Data1 = Val(txtCatchRate.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtDesc_Change()
    Item(EditorIndex).Desc = Trim$(txtDesc.Text)
    EditorChange = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.Text) Then
        Item(EditorIndex).Price = Val(txtPrice.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtValue_Change()
    If IsNumeric(txtValue.Text) Then
        Item(EditorIndex).Data2 = Val(txtValue.Text)
        EditorChange = True
    End If
End Sub
