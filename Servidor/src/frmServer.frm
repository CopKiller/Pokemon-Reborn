VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton btnTab 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton btnTab 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer tmrEvent 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   6720
      Top             =   600
   End
   Begin VB.Frame frmUser 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton btnAcess 
         Caption         =   "Dono"
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   13
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Dev"
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   12
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Mapper"
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   11
         Top             =   1070
         Width           =   1815
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Moderador"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   10
         Top             =   660
         Width           =   1815
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Remover"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton btnDisconnect 
         Caption         =   "Desconectar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4335
      End
      Begin VB.ListBox lbPlayers 
         Height          =   1620
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame frmInfo 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton btnReload 
         Caption         =   "Recarregar os Pokémons"
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   20
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton btnReload 
         Caption         =   "Recarregar os Npc's"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   19
         Top             =   1260
         Width           =   2295
      End
      Begin VB.CommandButton btnReload 
         Caption         =   "Recarregar o Mapa"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   18
         Top             =   840
         Width           =   2295
      End
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   3960
         Max             =   5
         Min             =   1
         TabIndex        =   14
         Top             =   490
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Desligar Servidor"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkStaffOnly 
         Caption         =   "Modo Desenvolvedor"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblExp 
         Caption         =   "Exp: 1"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   3720
         X2              =   3720
         Y1              =   2280
         Y2              =   240
      End
      Begin VB.Label lblCPS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "CPS: 0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.TextBox txtCommand 
      Height          =   380
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Timer tmrTotalOnline 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   120
   End
   Begin VB.TextBox txtLog 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin MSWinsockLib.Winsock Server_Socket 
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblGameTime 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   390
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "&PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect Index"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAcess_Click(Index As Integer)
Dim i As Long
Dim IdAccess As Long

    i = lbPlayers.ListIndex + 1
    
    If (lbPlayers.ListIndex > -1) Then
                 'Player(indexReturn, TempPlayer(indexReturn).UseChar).Name
        If IsPlaying(i) Then
        
            Select Case Index
                    'IdAccess = Index
                Case 1
                    IdAccess = Index
                Case 2
                    IdAccess = Index
                Case 3
                    IdAccess = Index
                Case 4
                    IdAccess = Index
                    
            End Select
            
            Player(i, TempPlayer(i).UseChar).Access = IdAccess
            SendPlayerData i
        End If
    End If
    
End Sub

Private Sub btnDisconnect_Click()
Dim i As Long

    i = lbPlayers.ListIndex + 1
    
    If (lbPlayers.ListIndex > -1) Then
        If GetPlayerIP(i) <> vbNullString Then
            CloseSocket (lbPlayers.ListIndex + 1)
        End If
    End If
    
End Sub

Private Sub btnReload_Click(Index As Integer)
Dim i As Long
    
    Select Case Index
        Case 0
            Call LoadMaps
            TextAdd frmServer.txtLog, "Os mapas foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i), GetPlayerDir(i)
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
        Case 1
            Call LoadNpcs
            TextAdd frmServer.txtLog, "Os npc's foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        SendNpcs i
                    Else
                        Exit Sub
                    End If
                End If
            Next
        Case 2
            Call LoadSpawns
            TextAdd frmServer.txtLog, "Os pokémons foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        SendSpawns i
                    Else
                        Exit Sub
                    End If
                End If
            Next
        
    End Select
End Sub

Private Sub btnTab_Click(Index As Integer)
        
    Select Case Index
        Case 0
            frmInfo.Visible = Not frmInfo.Visible
        Case 1
            frmUser.Visible = Not frmUser.Visible
    End Select
    
End Sub

Private Sub chkStaffOnly_Click()
Dim i As Long

    If chkStaffOnly.Value = YES Then
        '//Disconnect all non staff members
        If Player_HighIndex > 0 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        If Player(i, TempPlayer(i).UseChar).Access <= 0 Then
                            Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "You have been disconnected from the server.", White, YES
                                Case LANG_EN: AddAlert i, "You have been disconnected from the server.", White, YES
                                Case LANG_ES: AddAlert i, "You have been disconnected from the server.", White, YES
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End If
End Sub

Private Sub cmdShutdown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutdown.Caption = "Desligar Servidor"
        SendGlobalMsg "Shutdown canceled.", White
    Else
        isShuttingDown = True
        cmdShutdown.Caption = "Cancelar Desligamento"
        Secs = 180
    End If
End Sub

Private Sub scrlExp_Change()
    Dim CurExp As String
    Dim CurLanguage As Byte
    Dim Index As Long
    
    CurExp = scrlExp.Value
    
    If scrlExp.Value > 1 Then
        frmServer.tmrEvent.Enabled = True
    End If
    
    For Index = 1 To MAX_PLAYER
        If IsPlaying(Index) Then
            If TempPlayer(Index).UseChar > 0 Then
            
                Dim TextPT, TextEN, TextES As String
                
                If CurExp = 1 Then
                    TextPT = "O evento de experiência foi finalizado, Experiência atual é de: "
                    TextEN = "The experience event is finished, Current experience is: "
                    TextES = "El evento de la experiencia ha sido finalizado, la experiencia actual es: "
                Else
                    TextPT = "O evento de experiência foi ativado, Experiência atual é de: "
                    TextEN = "The experience event has been activated, Current experience is: "
                    TextES = "El evento de experiencia se ha activado, la experiencia actual es: "
                End If
                    
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, TextPT + CurExp + "x", White
                    Case LANG_EN: AddAlert Index, TextEN + CurExp + "x", White
                    Case LANG_ES: AddAlert Index, TextES + CurExp + "x", White
                End Select
            End If
        End If
    Next
    
    Select Case CurExp
        Case 1
            CurExp = "Exp: 1"
        Case 2
            CurExp = "Exp: 2"
        Case 3
            CurExp = "Exp: 3"
        Case 4
            CurExp = "Exp: 4"
        Case 5
            CurExp = "Exp: 5"
    End Select
                 
    frmServer.lblExp.Caption = CurExp
End Sub

Private Sub Server_Socket_DataArrival(ByVal bytesTotal As Long)
    If IsServerConnected Then Call main_IncomingData(bytesTotal)
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim count As Byte
Dim i As Long

    ' Check connection
    count = 0
    For i = 1 To MAX_PLAYER
        If IsConnected(i) Then
            If GetPlayerIP(i) = Socket(Index).RemoteHostIP Then
                count = count + 1
                If count >= 5 Then Exit Sub
            End If
        End If
    Next
    
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
Dim count As Byte
Dim i As Long

    ' Check connection
    count = 0
    For i = 1 To MAX_PLAYER
        If IsConnected(i) Then
            If GetPlayerIP(i) = Socket(Index).RemoteHostIP Then
                count = count + 1
                If count >= 5 Then Exit Sub
            End If
        End If
    Next
    
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' *****************
' ** Form object **
' *****************
Private Sub Form_Unload(Cancel As Integer)
    DestroyServer
End Sub

Private Sub tmrEvent_Timer()
    frmServer.scrlExp.Value = 1
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
Dim Index As Long
Dim Command() As String
Dim chatMsg As String
Dim CurLanguage As Byte

    If KeyAscii = vbKeyReturn Then
        
        chatMsg = Trim$(txtCommand.Text)
        
        If Left$(chatMsg, 1) = "/" Then
            chatMsg = LCase(Trim$(txtCommand.Text))
            Command = Split(chatMsg, Space(1))
            
            Select Case Command(0)
                Case "/online"
                    TextAdd frmServer.txtLog, "Jogadores Online: " & TotalPlayerOnline
                
                Case "/clear"
                    txtLog.Text = vbNullString
                    
            End Select
            
        Else
        
            If LenB(Trim$(txtCommand.Text)) > 0 Then
                
                Select Case CurLanguage
                    Case LANG_PT: Call SendGlobalMsg("[SERVIDOR]: " + txtCommand.Text, White)
                    Case LANG_EN: Call SendGlobalMsg("[SERVER]: " + txtCommand.Text, White)
                    Case LANG_ES: Call SendGlobalMsg("[SERVIDOR]: " + txtCommand.Text, White)
                End Select

                For Index = 1 To MAX_PLAYER
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Select Case CurLanguage
                                Case LANG_PT: AddAlert Index, "Servidor:" + txtCommand.Text, White
                                Case LANG_EN: AddAlert Index, "Server:" + txtCommand.Text, White
                                Case LANG_ES: AddAlert Index, "Servidor:" + txtCommand.Text, White
                            End Select
                        End If
                    End If
                Next
                
            End If
            
        End If

        KeyAscii = 0
        txtCommand.Text = vbNullString
    End If
End Sub

Private Sub txtLog_GotFocus()
    txtCommand.SetFocus
    DoEvents
End Sub

'
Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYER
        frmServer.lbPlayers.AddItem vbNullString

        If i < 10 Then
            frmServer.lbPlayers.List(i) = "00" & i
        ElseIf i < 100 Then
            frmServer.lbPlayers.List(i) = "0" & i
        Else
            frmServer.lbPlayers.List(i) = i
        End If
    Next

End Sub
