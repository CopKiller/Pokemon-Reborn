VERSION 5.00
Begin VB.Form frmEditorShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   5775
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   5655
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Moeda de Troca"
         Height          =   1815
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   5535
         Begin VB.TextBox txtFind2 
            Height          =   285
            Left            =   3000
            TabIndex        =   20
            Top             =   1320
            Width           =   1935
         End
         Begin VB.OptionButton optValue 
            Caption         =   "Item"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optValue 
            Caption         =   "Money"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.HScrollBar scrlSellItemNum 
            Height          =   255
            Left            =   240
            Max             =   0
            TabIndex        =   16
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtPrice 
            Height          =   285
            Left            =   2160
            TabIndex        =   12
            Text            =   "0"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Find:"
            Height          =   195
            Left            =   3000
            TabIndex        =   19
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label lblMoney 
            AutoSize        =   -1  'True
            Caption         =   "Money"
            Height          =   195
            Left            =   1560
            TabIndex        =   13
            Top             =   600
            Width           =   480
         End
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Left            =   240
         Max             =   0
         TabIndex        =   10
         Top             =   3000
         Width           =   3255
      End
      Begin VB.ListBox lstShopItem 
         Height          =   1620
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   4575
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
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Find:"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblItemNum 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Shop Items"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   5130
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
Attribute VB_Name = "frmEditorShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseShopEditor
    End If
End Sub

Private Sub Form_Load()
    txtName.MaxLength = NAME_LENGTH
    scrlItemNum.max = MAX_ITEM
End Sub

Private Sub lstIndex_Click()
    ShopEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub lstShopItem_Click()
    scrlItemNum.value = Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).Num
    txtPrice.Text = Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).Price
    
    optValue(Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).SellValueType).value = True
    
    If Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).SellValueType = 1 Then
        scrlSellItemNum.max = MAX_ITEM
        scrlSellItemNum.value = Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).SellValueId
    End If
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestShop
    End If
    CloseShopEditor
End Sub

Private Sub mnuExit_Click()
    CloseShopEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_SHOP
        If ShopChange(i) Then
            SendSaveShop i
            ShopChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseShopEditor
End Sub

Private Sub optValue_Click(Index As Integer)
    Dim shopIndex As Long
    
    shopIndex = lstShopItem.ListIndex + 1
    If shopIndex = 0 Then Exit Sub
    
    Select Case Index
        Case 0
            lblMoney.Caption = "Money:"
            scrlSellItemNum.Enabled = False
            txtFind2.Enabled = False
        Case 1
            lblMoney.Caption = "Quant:"
            scrlSellItemNum.Enabled = True
            txtFind2.Enabled = True
    End Select
    
    Shop(EditorIndex).ShopItem(shopIndex).SellValueType = Index
End Sub

Private Sub scrlItemNum_Change()
    Dim tmpIndex As Long
    Dim shopIndex As Long

    shopIndex = lstShopItem.ListIndex + 1
    If shopIndex = 0 Then Exit Sub
    tmpIndex = lstShopItem.ListIndex
    Shop(EditorIndex).ShopItem(shopIndex).Num = scrlItemNum.value

    lstShopItem.RemoveItem shopIndex - 1
    If Shop(EditorIndex).ShopItem(shopIndex).Num > 0 Then
        If Shop(EditorIndex).ShopItem(shopIndex).SellValueType = 1 And Shop(EditorIndex).ShopItem(shopIndex).SellValueId > 0 Then
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Item>" & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).SellValueId).Name) & ">" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
        Else
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - $ " & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
        End If

        If Shop(EditorIndex).ShopItem(shopIndex).SellValueType = 0 And Shop(EditorIndex).ShopItem(shopIndex).Price = 0 Then
            Shop(EditorIndex).ShopItem(shopIndex).Price = Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Price
        End If
        
        txtPrice.Text = Shop(EditorIndex).ShopItem(shopIndex).Price
    Else
        lstShopItem.AddItem shopIndex & ": None - Price: $ 0", shopIndex - 1
        lblMoney = "Money:"
    End If


    lstShopItem.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub scrlSellItemNum_Change()
    Dim tmpIndex As Long
    Dim shopIndex As Long

    shopIndex = lstShopItem.ListIndex + 1
    If shopIndex = 0 Then Exit Sub
    If Shop(EditorIndex).ShopItem(shopIndex).SellValueType <> 1 Then Exit Sub

    tmpIndex = lstShopItem.ListIndex
    Shop(EditorIndex).ShopItem(shopIndex).SellValueId = scrlSellItemNum.value
    lstShopItem.RemoveItem shopIndex - 1
    If Shop(EditorIndex).ShopItem(shopIndex).Num > 0 Then
        If Shop(EditorIndex).ShopItem(shopIndex).SellValueId > 0 Then
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Item" & ">" & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).SellValueId).Name) & ">" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
        Else
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Item" & "> & " > " & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1"
        End If
    Else
        lstShopItem.AddItem shopIndex & ": None - Item>>", shopIndex - 1
    End If
    lstShopItem.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtFind_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte

    ' Maior Índice  \/
    MAX_INDEX = MAX_ITEM

    ' Quantidade Mínima de caracteres pra procurar
    MinChar = 2

    ' Nome deste controle
    If Not IsNumeric(txtFind) Then
        ' Nome deste controle
        Find = UCase$(Trim$(txtFind))
        If Len(Find) <= MinChar And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_INDEX
            If Not Find = "" Then
                ' Atribuição da estrutura em procura
                If InStr(1, UCase$(Trim$(Item(i).Name)), Find) > 0 Then
                    ' Nome do controle a ser alterado
                    scrlItemNum = i
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Nome deste controle
        If txtFind > MAX_INDEX Then
            ' Nome deste controle
            txtFind = MAX_INDEX
            ' Nome deste controle
        ElseIf txtFind <= 0 Then
            ' Nome deste controle
            txtFind = 1
        End If
        ' Nome do controle a ser alterado & Nome deste controle
        scrlItemNum = txtFind
    End If
End Sub

Private Sub txtFind2_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte

    ' Maior Índice  \/
    MAX_INDEX = MAX_ITEM

    ' Quantidade Mínima de caracteres pra procurar
    MinChar = 2

    ' Nome deste controle
    If Not IsNumeric(txtFind2) Then
        ' Nome deste controle
        Find = UCase$(Trim$(txtFind2))
        If Len(Find) <= MinChar And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_INDEX
            If Not Find = "" Then
                ' Atribuição da estrutura em procura
                If InStr(1, UCase$(Trim$(Item(i).Name)), Find) > 0 Then
                    ' Nome do controle a ser alterado
                    scrlSellItemNum = i
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Nome deste controle
        If txtFind2 > MAX_INDEX Then
            ' Nome deste controle
            txtFind2 = MAX_INDEX
            ' Nome deste controle
        ElseIf txtFind2 <= 0 Then
            ' Nome deste controle
            txtFind2 = 1
        End If
        ' Nome do controle a ser alterado & Nome deste controle
        scrlSellItemNum = txtFind2
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtPrice_Change()
Dim tmpIndex As Long
Dim shopIndex As Long

    shopIndex = lstShopItem.ListIndex + 1
    If shopIndex = 0 Then Exit Sub
    tmpIndex = lstShopItem.ListIndex
    If IsNumeric(txtPrice.Text) Then
        Shop(EditorIndex).ShopItem(shopIndex).Price = Val(txtPrice.Text)
    End If
    lstShopItem.RemoveItem shopIndex - 1
    If Shop(EditorIndex).ShopItem(shopIndex).Num > 0 Then

        If Shop(EditorIndex).ShopItem(shopIndex).SellValueType = 1 Then
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Item>" & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & ">" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
        Else
            lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Price: $" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
        End If
    Else
        lstShopItem.AddItem shopIndex & ": None - Price: $" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
    End If
    
    lstShopItem.ListIndex = tmpIndex
    EditorChange = True
End Sub
