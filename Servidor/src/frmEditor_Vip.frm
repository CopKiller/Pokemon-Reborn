VERSION 5.00
Begin VB.Form frmEditor_Vip 
   Caption         =   "Vip Settings"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vip Rewards"
      Height          =   2175
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtDeathPenalty 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Text            =   "0"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtShopPrice 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "0"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDropChance 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "0"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtVipCoin 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Text            =   "0"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtVipExp 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "-Death Penalty %:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "-Shop Price %:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "+Drop Chance %:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "+Coin %:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "+Vip Exp %:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vip Type"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optVipType 
         Caption         =   "Vip Gold"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optVipType 
         Caption         =   "Vip Silver"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optVipType 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Vip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    SetVipSettings
    SaveVipSettings
    SendVipAdvantageToAll
End Sub

Private Sub Form_Load()
    LoadVipSettings
End Sub

Private Sub optVipType_Click(index As Integer)
    LoadVipSettings
End Sub

Private Sub txtDeathPenalty_Change()
    If Not IsNumeric(txtDeathPenalty) Then
        txtDeathPenalty = 0
    End If
End Sub

Private Sub txtDropChance_Change()
    If Not IsNumeric(txtDropChance) Then
        txtDropChance = 0
    End If
End Sub

Private Sub txtShopPrice_Change()
    If Not IsNumeric(txtShopPrice) Then
        txtShopPrice = 0
    End If
End Sub

Private Sub txtVipCoin_Change()
    If Not IsNumeric(txtVipCoin) Then
        txtVipCoin = 0
    End If
End Sub

Private Sub txtVipExp_Change()
    If Not IsNumeric(txtVipExp) Then
        txtVipExp = 0
    End If
End Sub

Private Function GiveVipTypeSelected() As Byte
    Dim i As Long
    
    For i = 0 To EnumVipType.VipCount - 1
        If optVipType(i) = True Then GiveVipTypeSelected = i
    Next i
End Function

Private Sub LoadVipSettings()
    txtVipExp = VipSettings(GiveVipTypeSelected).VipExp
    txtVipCoin = VipSettings(GiveVipTypeSelected).VipCoin
    txtDropChance = VipSettings(GiveVipTypeSelected).VipDrop
    txtShopPrice = VipSettings(GiveVipTypeSelected).VipShopPrice
    txtDeathPenalty = VipSettings(GiveVipTypeSelected).VipDeathPenalty
End Sub

Private Sub SetVipSettings()
    VipSettings(GiveVipTypeSelected).VipExp = txtVipExp
    VipSettings(GiveVipTypeSelected).VipCoin = txtVipCoin
    VipSettings(GiveVipTypeSelected).VipDrop = txtDropChance
    VipSettings(GiveVipTypeSelected).VipShopPrice = txtShopPrice
    VipSettings(GiveVipTypeSelected).VipDeathPenalty = txtDeathPenalty
End Sub
