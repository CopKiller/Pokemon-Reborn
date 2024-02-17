VERSION 5.00
Begin VB.Form frmEditor_Store 
   Caption         =   "Edit Store by Peixonalta"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Init"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cmbTypeStore 
         Height          =   315
         ItemData        =   "frmEditor_Store.frx":0000
         Left            =   240
         List            =   "frmEditor_Store.frx":0002
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Remove 1"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Add 1"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Store Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblMaxSlots 
         AutoSize        =   -1  'True
         Caption         =   "Max Slots: 0"
         Height          =   195
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Container"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      Begin VB.TextBox txtAvailableQuant 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   28
         Text            =   "0"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox chkIsLimited 
         Caption         =   "Is Limited?"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkIsNew 
         Caption         =   "Is New?"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CheckBox chkCustom 
         Caption         =   "Custom Description"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtItemPrice 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtSlotNum 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtItemNum 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtItemQuant 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Quant:"
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   810
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label6 
         Caption         =   "Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "SlotNum:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "ItemNum:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Quant:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEditor_Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_QUANT As Long = 999
Private Const MAX_PRICE As Long = 999

Private SlotNum As Long

Private Sub chkCustom_Click()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As Byte
    Value = chkCustom
    
    If SlotNum > VirtualShop(index).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If
    
    VirtualShop(index).Items(SlotNum).CustomDesc = Value
End Sub

Private Sub chkIsLimited_Click()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As Byte
    Value = chkIsLimited
    
    If SlotNum > VirtualShop(index).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If
    
    If Value = YES Then
        txtAvailableQuant.Enabled = True
    Else
        txtAvailableQuant.Enabled = False
    End If
    
    VirtualShop(index).Items(SlotNum).IsLimited = Value
End Sub

Private Sub chkIsNew_Click()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As Byte
    Value = chkIsNew
    
    If SlotNum > VirtualShop(index).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If
    
    VirtualShop(index).Items(SlotNum).IsNew = Value
End Sub

Private Sub cmbTypeStore_Click()
    SlotNum = 0
    RefreshControls
End Sub

Private Sub Command1_Click(index As Integer)
    Dim i As Long
    i = cmbTypeStore.ListIndex + 1
    If SlotNum > VirtualShop(i).Max_Slots Or SlotNum <= 0 Then
        Exit Sub
    End If

    If index = 0 Then
        If SlotNum <= 1 Then
            Exit Sub
        End If

        SlotNum = SlotNum - 1
    Else
        If SlotNum >= VirtualShop(i).Max_Slots Then
            Exit Sub
        End If

        SlotNum = SlotNum + 1
    End If

    RefreshControls
End Sub

Private Sub Command2_Click(index As Integer)
    Dim i As Long
    i = cmbTypeStore.ListIndex + 1
    If SlotNum > VirtualShop(i).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If

    If index = 0 Then
        If VirtualShop(i).Items(SlotNum).ItemNum <= 0 Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemNum = VirtualShop(i).Items(SlotNum).ItemNum - 1
    Else
        If VirtualShop(i).Items(SlotNum).ItemNum >= MAX_ITEM Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemNum = VirtualShop(i).Items(SlotNum).ItemNum + 1
    End If

    RefreshControls
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command3_Click(index As Integer)
    Dim i As Long
    i = cmbTypeStore.ListIndex + 1
    If SlotNum > VirtualShop(i).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If

    If index = 0 Then
        If VirtualShop(i).Items(SlotNum).ItemQuant <= 0 Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemQuant = VirtualShop(i).Items(SlotNum).ItemQuant - 1
    Else
        If VirtualShop(i).Items(SlotNum).ItemQuant >= MAX_QUANT Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemQuant = VirtualShop(i).Items(SlotNum).ItemQuant + 1
    End If

    RefreshControls
End Sub

Private Sub Command4_Click(index As Integer)
    Dim i As Long
    i = cmbTypeStore.ListIndex + 1
    If SlotNum > VirtualShop(i).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If

    If index = 0 Then
        If VirtualShop(i).Items(SlotNum).ItemPrice <= 0 Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemPrice = VirtualShop(i).Items(SlotNum).ItemPrice - 1
    Else
        If VirtualShop(i).Items(SlotNum).ItemPrice >= MAX_PRICE Then Exit Sub
        VirtualShop(i).Items(SlotNum).ItemPrice = VirtualShop(i).Items(SlotNum).ItemPrice + 1
    End If

    RefreshControls
End Sub

Private Sub Command5_Click()
    Call SaveVirtualShop
    
End Sub

Private Sub Command66_Click()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    
    If VirtualShop(index).Max_Slots <= 1 Then Exit Sub
    
    VirtualShop(index).Max_Slots = VirtualShop(index).Max_Slots - 1
    ReDim Preserve VirtualShop(index).Items(1 To VirtualShop(index).Max_Slots)
    RefreshControls
End Sub

Private Sub Command67_Click()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    
    VirtualShop(index).Max_Slots = VirtualShop(index).Max_Slots + 1
    ReDim Preserve VirtualShop(index).Items(1 To VirtualShop(index).Max_Slots)
    RefreshControls
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Select Case i
        Case VirtualShopTabsRec.Skins: frmEditor_Store.cmbTypeStore.AddItem "Skins"
        Case VirtualShopTabsRec.Mounts: frmEditor_Store.cmbTypeStore.AddItem "Mounts"
        Case VirtualShopTabsRec.Items: frmEditor_Store.cmbTypeStore.AddItem "Items"
        Case VirtualShopTabsRec.Vips: frmEditor_Store.cmbTypeStore.AddItem "Vips"
    End Select
    Next i
    'Show first Data
    SlotNum = 1
    
    cmbTypeStore.ListIndex = 0
    
End Sub

Private Sub RefreshControls()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    
    lblMaxSlots = "Max Slots: " & VirtualShop(index).Max_Slots
    
    txtSlotNum = SlotNum
    
    If UBound(VirtualShop(index).Items) < SlotNum Then
        SlotNum = SlotNum - 1
        RefreshControls
    End If
    
    txtItemNum = VirtualShop(index).Items(SlotNum).ItemNum
    txtItemQuant = VirtualShop(index).Items(SlotNum).ItemQuant
    txtItemPrice = VirtualShop(index).Items(SlotNum).ItemPrice
    chkCustom = VirtualShop(index).Items(SlotNum).CustomDesc
    chkIsNew = VirtualShop(index).Items(SlotNum).IsNew
    chkIsLimited = VirtualShop(index).Items(SlotNum).IsLimited
    txtAvailableQuant = VirtualShop(index).Items(SlotNum).AvailableQuant
    
    If txtItemNum > 0 Then
        lblItemName = "Item Name: " & Trim$(Item(txtItemNum).Name)
    Else
        lblItemName = "Item Name: None."
    End If
End Sub

Private Sub txtAvailableQuant_Change()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As Long
    
    If Not IsNumeric(txtAvailableQuant) Then
        Value = 0
        txtAvailableQuant = Value
    Else
        Value = txtAvailableQuant
    End If

    If SlotNum > VirtualShop(index).Max_Slots Or SlotNum < 1 Then
        Exit Sub
    End If

    VirtualShop(index).Items(SlotNum).AvailableQuant = Value
End Sub

Private Sub txtItemNum_Change()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As String
    Value = txtItemNum
    
    If Not IsNumeric(Value) Then
        Value = VirtualShop(index).Items(SlotNum).ItemNum
    End If
    
    If Value > MAX_ITEM Then
        Value = VirtualShop(index).Items(SlotNum).ItemNum
    ElseIf Value < 0 Then
        Value = VirtualShop(index).Items(SlotNum).ItemNum
    End If
    
    VirtualShop(index).Items(SlotNum).ItemNum = CLng(Value)
    RefreshControls
End Sub

Private Sub txtItemPrice_Change()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As String
    Value = txtItemPrice
    
    If Not IsNumeric(Value) Then
        Value = VirtualShop(index).Items(SlotNum).ItemPrice
    End If
    
    If Value > MAX_QUANT Then
        Value = VirtualShop(index).Items(SlotNum).ItemPrice
    ElseIf Value < 0 Then
        Value = VirtualShop(index).Items(SlotNum).ItemPrice
    End If
    
    VirtualShop(index).Items(SlotNum).ItemPrice = CLng(Value)
    RefreshControls
End Sub

Private Sub txtItemQuant_Change()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As String
    Value = txtItemQuant
    
    If Not IsNumeric(Value) Then
        Value = VirtualShop(index).Items(SlotNum).ItemQuant
    End If
    
    If Value > MAX_QUANT Then
        Value = VirtualShop(index).Items(SlotNum).ItemQuant
    ElseIf Value < 0 Then
        Value = VirtualShop(index).Items(SlotNum).ItemQuant
    End If
    
    VirtualShop(index).Items(SlotNum).ItemQuant = CLng(Value)
    RefreshControls
End Sub

Private Sub txtSlotNum_Change()
    Dim index As Long
    index = cmbTypeStore.ListIndex + 1
    Dim Value As String
    Value = txtSlotNum
    
    If Not IsNumeric(Value) Then
        Value = LBound(VirtualShop(index).Items)
    End If
    
    If Value > VirtualShop(index).Max_Slots Then
        Value = VirtualShop(index).Max_Slots
    ElseIf Value < LBound(VirtualShop(index).Items) Then
        Value = LBound(VirtualShop(index).Items)
    End If
    
    txtSlotNum = CLng(Value)
    SlotNum = CLng(Value)
End Sub



