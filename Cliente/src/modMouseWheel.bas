Attribute VB_Name = "modMouseWheel"
Option Explicit

'***********************************************************************************
'*****THIS SOURCE HAS BEEN OBTAINED FROM http://www.andreavb.com/tip060008.html*****
'***********************************************************************************

'************************************************************
'API
'************************************************************
Private Declare Function CallWindowProc _
                          Lib "user32.dll" _
                              Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function SetWindowLong _
                          Lib "user32.dll" _
                              Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

'************************************************************
'Constants
'************************************************************

Public Const MK_CONTROL = &H8

Public Const MK_LBUTTON = &H1

Public Const MK_RBUTTON = &H2

Public Const MK_MBUTTON = &H10

Public Const MK_SHIFT = &H4

Private Const GWL_WNDPROC = -4

Private Const WM_MOUSEWHEEL = &H20A

'************************************************************
'Variables
'************************************************************

Private hControl As Long

Private lPrevWndProc As Long

Public myHWnd As Long


Public Enum ScrollDir
    Down = 0
    Up
End Enum

'*************************************************************
'WindowProc
'*************************************************************

Private Function WindowProc(ByVal lWnd As Long, _
                            ByVal lMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

'Test if the message is WM_MOUSEWHEEL

    If lMsg = WM_MOUSEWHEEL And frmMain.hwnd = myHWnd Then

        'Add event handling code here
        'this will be universal to all forms that are 'hooked' to this code
        
        ''AddText "Window Proc", White
        
        If wParam > 0 Then    ' Moved up
            HandleMouseWheelMove ScrollDir.Down
        Else    ' Moved down
            HandleMouseWheelMove ScrollDir.Up
        End If
    End If

    'Sends message to previous procedure if not MOUSEWHEEL
    'This is VERY IMPORTANT!!!
    If lMsg <> WM_MOUSEWHEEL Then
        WindowProc = CallWindowProc(lPrevWndProc, lWnd, lMsg, wParam, lParam)
    End If

End Function

'*************************************************************
'Hook
'All forms that call this procedure must implement this procedure in their module:
'   Public Sub MouseWheelRolled()
'        <your code>
'   End Sub
'*************************************************************
Public Sub HookForMouseWheel(ByVal hControl_ As Long)

    hControl = hControl_
    lPrevWndProc = SetWindowLong(hControl, GWL_WNDPROC, AddressOf WindowProc)

End Sub

'*************************************************************
'UnHook
'*************************************************************
Public Sub UnHookMouseWheel()

    If hControl = 0 Then Exit Sub    ' Don't do anything if we haven't hooked yet
    Call SetWindowLong(hControl, GWL_WNDPROC, lPrevWndProc)

End Sub

Public Sub HandleMouseWheelMove(ByVal Dir As ScrollDir)

    FormMouseScroll Dir

End Sub






Public Sub Window_Scroll(ByVal window As GuiEnum, ByVal Dir As ScrollDir, ByRef varDown As Boolean, ByRef varUp As Boolean)

    With GUI(window)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub

        If ReInit Then Exit Sub

        If Dir = ScrollDir.Down Then  ' Down.
            varUp = True    ' Don't know why but this boolean is inverted. It works though
            varDown = False
        Else
            varDown = True    ' Don't know why but this boolean is inverted. It works though
            varUp = False
        End If

    End With

End Sub


