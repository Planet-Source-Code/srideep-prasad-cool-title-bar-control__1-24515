Attribute VB_Name = "Runtime"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As _
    Any, Source As Any, ByVal bytes As Long)
Public Const GWL_WNDPROC = -4
Global SubObj As Callback
Global ProcOld As Long
' Returns an object given its pointer
' This function reverses the effect of the ObjPtr function

Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    ' force the value of the pointer into the temporary object variable
    CopyMemory obj, pObj, 4
    ' assign to the result (this increments the ref counter)
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function





Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hMsg As Long, PWndProc As Long, Total As Long, i As Long
        
        
    If GetProp(hWnd, CStr(hWnd) & "%" & CStr(uMsg)) = 1 Then
        WindowProc = SubObj.WindowProc(hWnd, uMsg, wParam, lParam)
    Else
        WindowProc = CallDefaultWndProc(hWnd, uMsg, wParam, lParam)
    End If
End Function


Private Function CallDefaultWndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
        
    CallDefaultWndProc = CallWindowProc(ProcOld, hWnd, uMsg, wParam, lParam)
End Function


