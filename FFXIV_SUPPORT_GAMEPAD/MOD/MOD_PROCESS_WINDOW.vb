Imports System.Runtime.InteropServices
Module MOD_PROCESS_WINDOW

#Region "WIN32API"
    <DllImport("user32.dll")> Private Function GetWindowRect(ByVal hwnd As IntPtr, ByRef lpRect As RECT) As Integer
    End Function
    <DllImport("user32.dll")> Private Function GetClientRect(ByVal hwnd As IntPtr, ByRef lpRect As RECT) As Integer
    End Function
#End Region

    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Public Structure RECT_WH
        Public left As Integer
        Public top As Integer
        Public width As Integer
        Public height As Integer
    End Structure

    Public Function FUNC_GET_CRIENT_RECT(ByRef prcCRIENT As Process) As RECT
        Dim srtRET As RECT

        With srtRET
            .left = 0
            .top = 0
            .right = 0
            .bottom = 0
        End With

        If prcCRIENT Is Nothing Then
            Return srtRET
        End If

        Call GetClientRect(prcCRIENT.MainWindowHandle, srtRET)

        Return srtRET
    End Function

    Public Function FUNC_GET_WINDOW_RECT(ByRef prcWINDOW As Process) As RECT
        Dim srtRET As RECT

        With srtRET
            .left = 0
            .top = 0
            .right = 0
            .bottom = 0
        End With

        If prcWINDOW Is Nothing Then
            Return srtRET
        End If

        Call GetWindowRect(prcWINDOW.MainWindowHandle, srtRET)

        Return srtRET
    End Function

    Public Function FUNC_GET_RECT_WIDTH(ByRef srtRECT As RECT) As Integer
        Dim intRET As Integer

        intRET = srtRECT.right - srtRECT.left

        Return intRET
    End Function

    Public Function FUNC_GET_RECT_HEIGHT(ByRef srtRECT As RECT) As Integer
        Dim intRET As Integer

        intRET = srtRECT.bottom - srtRECT.top

        Return intRET
    End Function

    Public Function FUNC_GET_LEFT_WINDOW(ByRef PRC_WINDOW As Process) As Integer
        Dim srtRECT As RECT
        srtRECT = FUNC_GET_WINDOW_RECT(PRC_WINDOW)

        Dim intWINDOW_LEFT As Integer
        Dim intWINDOW_TOP As Integer
        intWINDOW_LEFT = srtRECT.left
        intWINDOW_TOP = srtRECT.top

        Dim intWINDOW_WIDTH As Integer
        Dim intWINDOW_HEIGHT As Integer
        intWINDOW_WIDTH = FUNC_GET_RECT_WIDTH(srtRECT)
        intWINDOW_HEIGHT = FUNC_GET_RECT_HEIGHT(srtRECT)

        srtRECT = FUNC_GET_CRIENT_RECT(PRC_WINDOW)
        Dim intCLIENT_WIDTH As Integer
        Dim intCLIENT_HEIGHT As Integer
        intCLIENT_WIDTH = FUNC_GET_RECT_WIDTH(srtRECT)
        intCLIENT_HEIGHT = FUNC_GET_RECT_HEIGHT(srtRECT)

        Dim intWIDTH_SUB As Integer
        Dim intHEIGHT_SUB As Integer
        intWIDTH_SUB = intWINDOW_WIDTH - intCLIENT_WIDTH
        intHEIGHT_SUB = intWINDOW_HEIGHT - intCLIENT_HEIGHT

        Dim intBORDER As Integer
        intWINDOW_LEFT += intBORDER
        intWINDOW_TOP += (intHEIGHT_SUB - intBORDER)

        Return intWINDOW_LEFT
    End Function

    Public Function FUNC_GET_TOP_WINDOW(ByRef PRC_WINDOW As Process) As Integer
        Dim srtRECT As RECT
        srtRECT = FUNC_GET_WINDOW_RECT(PRC_WINDOW)

        Dim intWINDOW_LEFT As Integer
        Dim intWINDOW_TOP As Integer
        intWINDOW_LEFT = srtRECT.left
        intWINDOW_TOP = srtRECT.top

        Dim intWINDOW_WIDTH As Integer
        Dim intWINDOW_HEIGHT As Integer
        intWINDOW_WIDTH = FUNC_GET_RECT_WIDTH(srtRECT)
        intWINDOW_HEIGHT = FUNC_GET_RECT_HEIGHT(srtRECT)

        srtRECT = FUNC_GET_CRIENT_RECT(PRC_WINDOW)
        Dim intCLIENT_WIDTH As Integer
        Dim intCLIENT_HEIGHT As Integer
        intCLIENT_WIDTH = FUNC_GET_RECT_WIDTH(srtRECT)
        intCLIENT_HEIGHT = FUNC_GET_RECT_HEIGHT(srtRECT)

        Dim intWIDTH_SUB As Integer
        Dim intHEIGHT_SUB As Integer
        intWIDTH_SUB = intWINDOW_WIDTH - intCLIENT_WIDTH
        intHEIGHT_SUB = intWINDOW_HEIGHT - intCLIENT_HEIGHT

        Dim intBORDER As Integer
        intWINDOW_LEFT += intBORDER
        intWINDOW_TOP += (intHEIGHT_SUB - intBORDER)

        Return intWINDOW_TOP
    End Function
End Module
