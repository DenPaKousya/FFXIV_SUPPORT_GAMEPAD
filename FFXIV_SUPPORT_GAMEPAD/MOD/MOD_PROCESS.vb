Imports System.Runtime.InteropServices
Module MOD_PROCESS

#Region "モジュール用・定数"
    Private Const cstDEFAULT_PROCESS_NAME As String = "ffxiv_dx11"
    'Private Const cstDEFAULT_PROCESS_NAME As String = "chrome"
    'Private Const cstDEFAULT_PROCESS_NAME As String = "notepad"
    'Private Const cstDEFAULT_PROCESS_NAME As String = "Calculator"
    'Private Const cstDEFAULT_PROCESS_NAME As String = "FlowScape"
#End Region

    Public Function FUNC_GET_FFXIV_PROCESS(Optional ByVal strFFXIV_PROCESS_NAME As String = cstDEFAULT_PROCESS_NAME) As Process
        Dim p As Process
        For Each p In Process.GetProcesses()
            If (p.MainWindowHandle <> IntPtr.Zero) Then
                If (p.ProcessName = strFFXIV_PROCESS_NAME) Then
                    Return p
                End If
            End If
        Next

        Return Nothing
    End Function

End Module
