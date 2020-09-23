Attribute VB_Name = "MSubclass"
Option Explicit

' declares:
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const GWL_WNDPROC = (-4)


Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_iProcOld As Long

Public Property Get CurrentMessage() As Long
   CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeCantSubclass
            sText = "Can't subclass window"
        Case eeAlreadyAttached
            sText = "Message already handled by another class"
        Case eeInvalidWindow
            sText = "Invalid window"
        Case eeNoExternalWindow
            sText = "Can't modify external window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        Err.Raise e, sSource
    End If
End Sub

Sub AttachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    Dim iC As Long, bFail As Boolean
    
    If IsWindow(hwnd) = False Then ErrRaise eeInvalidWindow
    If IsWindowLocal(hwnd) = False Then ErrRaise eeNoExternalWindow

    c = GetProp(hwnd, "C" & hwnd)
    If c = 0 Then
        procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
        If procOld = 0 Then ErrRaise eeCantSubclass
        f = SetProp(hwnd, hwnd, procOld)
        Debug.Assert f <> 0
        c = 1
        f = SetProp(hwnd, "C" & hwnd, c)
    Else
        c = c + 1
        f = SetProp(hwnd, "C" & hwnd, c)
    End If
    Debug.Assert f <> 0
    
    c = GetProp(hwnd, hwnd & "#" & iMsg & "C")
    If (c > 0) Then
        For iC = 1 To c
            If (GetProp(hwnd, hwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                ErrRaise eeAlreadyAttached
                bFail = True
                Exit For
            End If
        Next iC
    End If
                
    If Not (bFail) Then
        c = c + 1
        f = SetProp(hwnd, hwnd & "#" & iMsg & "C", c)
        Debug.Assert f <> 0
        
        f = SetProp(hwnd, hwnd & "#" & iMsg & "#" & c, ObjPtr(iwp))
        Debug.Assert f <> 0
    End If
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hwnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    Dim iC As Long, iP As Long, lPtr As Long
    
    c = GetProp(hwnd, "C" & hwnd)
    If c = 1 Then
        procOld = GetProp(hwnd, hwnd)
        Debug.Assert procOld <> 0
        Call SetWindowLong(hwnd, GWL_WNDPROC, procOld)
        RemoveProp hwnd, hwnd
        RemoveProp hwnd, "C" & hwnd
    Else
        c = GetProp(hwnd, "C" & hwnd)
        c = c - 1
        f = SetProp(hwnd, "C" & hwnd, c)
    End If
    
    c = GetProp(hwnd, hwnd & "#" & iMsg & "C")
    If (c > 0) Then
        For iC = 1 To c
            If (GetProp(hwnd, hwnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                iP = iC
                Exit For
            End If
        Next iC
    
        If (iP <> 0) Then
             For iC = iP + 1 To c
                lPtr = GetProp(hwnd, hwnd & "#" & iMsg & "#" & iC)
                SetProp hwnd, hwnd & "#" & iMsg & "#" & (iC - 1), lPtr
             Next iC
        End If
        RemoveProp hwnd, hwnd & "#" & iMsg & "#" & c
        c = c - 1
        SetProp hwnd, hwnd & "#" & iMsg & "C", c
    
    End If
End Sub

Private Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) _
                            As Long
    Dim procOld As Long, pSubclass As Long, f As Long
    Dim iwp As ISubclass, iwpT As ISubclass
    Dim iPC As Long, iP As Long, bNoProcess As Long
    Dim bCalled As Boolean
    
    procOld = GetProp(hwnd, hwnd)
    Debug.Assert procOld <> 0
    bCalled = False
    iPC = GetProp(hwnd, hwnd & "#" & iMsg & "C")
    If (iPC > 0) Then
        For iP = 1 To iPC
            bNoProcess = False
            pSubclass = GetProp(hwnd, hwnd & "#" & iMsg & "#" & iP)
            If pSubclass = 0 Then
                WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                            wParam, ByVal lParam)
                bNoProcess = True
            End If
            
            If Not (bNoProcess) Then
                CopyMemory iwpT, pSubclass, 4
                Set iwp = iwpT
                CopyMemory iwpT, 0&, 4
                m_iCurrentMessage = iMsg
                m_iProcOld = procOld
                With iwp
                    If (iP = 1) Then
                        If .MsgResponse = emrPreprocess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                    WindowProc = .WindowProc(hwnd, iMsg, wParam, ByVal lParam)
                    If (iP = iPC) Then
                        If .MsgResponse = emrPostProcess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                End With
            End If
        Next iP
    Else
        WindowProc = CallWindowProc(procOld, hwnd, iMsg, _
                                    wParam, ByVal lParam)
    End If
End Function
Public Function CallOldWindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
   CallOldWindowProc = CallWindowProc(m_iProcOld, hwnd, iMsg, wParam, lParam)

End Function

Function IsWindowLocal(ByVal hwnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hwnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
'

Function GetOldProc() As Long
    GetOldProc = m_iProcOld
End Function

