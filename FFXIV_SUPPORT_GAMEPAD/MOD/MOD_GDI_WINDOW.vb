Module MOD_GDI_WINDOW

#Region "WINAPI"
    Declare Function SetForegroundWindow Lib "user32.dll" (hWnd As IntPtr) As Boolean
    Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As IntPtr, ByVal hwndChildAfter As Integer, ByVal lpClassName As String, ByVal lpWindowName As String) As Integer 'クラス名、キャプションから子ウィンドウのハンドルを取得
    Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr
#End Region

    Public Sub SUB_FOREGROUND_WINDOW(ByRef prcTARGET As Process)
        Call SetForegroundWindow(prcTARGET.MainWindowHandle)
    End Sub

    Public Function FUNC_FIND_WINDOW_EX(ByVal ptrPARENT As IntPtr, ByVal ptrCHILD_AFTER As IntPtr, ByVal strCLASS_NAME As String, ByVal strWINDOW_NAME As String) As IntPtr
        Dim ptrRET As IntPtr

        ptrRET = FindWindowEx(ptrPARENT, ptrCHILD_AFTER, strCLASS_NAME, strWINDOW_NAME)

        Return ptrRET
    End Function

    Public Function FUNC_GET_FOREGROUND_WINDOW() As IntPtr
        Dim INT_RET As IntPtr
        INT_RET = GetForegroundWindow()
        Return INT_RET
    End Function

End Module
