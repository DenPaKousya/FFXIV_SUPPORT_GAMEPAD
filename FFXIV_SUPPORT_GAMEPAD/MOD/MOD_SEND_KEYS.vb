Module MOD_SEND_KEYS

#Region "WINAPI"
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As System.IntPtr, ByVal MSG As Integer, ByVal wParam As System.IntPtr, ByVal lParam As System.IntPtr) As System.IntPtr
#End Region

#Region "定数"
    Private Const CST_WAIT_ONE As Integer = 50
#End Region

#Region "列挙定数"
    Private Enum ENM_SEND_MSG
        WM_SETTEXT = &HC
        BM_CLICK = &HF5
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_COMMAND = &H111
        WM_IME_CHAR = &H286
    End Enum

    Public Enum ENM_MASK_KEYS
        NONE = 0
        SHIFT = 1
        CTRL = 2
        ALT = 3
    End Enum

    Public Enum ENM_SEND_VK
        VK_CLEAR = &HC
        VK_RETURN = &HD
        VK_SHIFT = &H10
        VK_CONTROL = &H11
        VK_MENU = &H12

        VK_SPACE = &H20
        VK_PRIOR = &H21
        VK_NEXT = &H22
        VK_END = &H23
        VK_HOME = &H24
        VK_LEFT = &H25
        VK_UP = &H26
        VK_RIGHT = &H27
        VK_DOWN = &H28

        VK_0 = &H30
        VK_1 = &H31
        VK_2 = &H32
        VK_3 = &H33
        VK_4 = &H34
        VK_5 = &H35
        VK_6 = &H36
        VK_7 = &H37
        VK_8 = &H38
        VK_9 = &H39

        VK_A = &H41
        VK_B = &H42
        VK_C = &H43
        VK_D = &H44
        VK_E = &H45
        VK_F = &H46
        VK_G = &H47
        VK_H = &H48
        VK_I = &H49
        VK_J = &H4A
        VK_K = &H4B
        VK_L = &H4C
        VK_M = &H4D
        VK_N = &H4E
        VK_O = &H4F
        VK_P = &H50
        VK_Q = &H51
        VK_R = &H52
        VK_S = &H53
        VK_T = &H54
        VK_U = &H55
        VK_V = &H56
        VK_W = &H57
        VK_X = &H58
        VK_Y = &H59
        VK_Z = &H5A

        VK_NUMPAD0 = &H60
        VK_NUMPAD1 = &H61
        VK_NUMPAD2 = &H62
        VK_NUMPAD3 = &H63
        VK_NUMPAD4 = &H64
        VK_NUMPAD5 = &H65
        VK_NUMPAD6 = &H66
        VK_NUMPAD7 = &H67
        VK_NUMPAD8 = &H68
        VK_NUMPAD9 = &H69
        VK_MULTIPLY = &H6A
        VK_ADD = &H6B
        VK_SUBTRACT = &H6D
        VK_DECIMAL = &H6E
        VK_DIVIDE = &H6F

        VK_F1 = &H70
        VK_F2 = &H71
        VK_F3 = &H72
        VK_F4 = &H73
        VK_F5 = &H74
        VK_F6 = &H75
        VK_F7 = &H76
        VK_F8 = &H77
        VK_F9 = &H78
        VK_F10 = &H79
        VK_F11 = &H7A
        VK_F12 = &H7B

        VK_SCROLL = &H91
        VK_OEM_MINUS = &HBD
        VK_OEM_2 = &HBF
        VK_OEM_7 = &HDE
    End Enum

#End Region

    Public Function FUNC_SEND_KEYS_STRING(ByRef prcTARGET As Process, ByVal strMSG As String) As Boolean
        Dim chrONE As Char
        Dim enmVK As ENM_SEND_VK
        Dim intLOOP_INDEX As Integer

        For intLOOP_INDEX = 1 To strMSG.Length
            chrONE = strMSG.Substring(intLOOP_INDEX - 1, 1)
            enmVK = FUNC_CNV_CHAR_TO_VK(chrONE)

            Dim BLN_IME As Boolean
            Select Case enmVK
                Case ENM_SEND_VK.VK_OEM_2
                    BLN_IME = False
                Case Else
                    BLN_IME = True
            End Select
            Call FUNC_SEND_KEYS(prcTARGET, enmVK,, BLN_IME)
        Next

        Return True
    End Function

    Public Function FUNC_SEND_KEYS_STRING_02(ByRef prcTARGET As Process, ByVal strMSG As String) As Boolean
        Dim chrONE As Char
        Dim intLOOP_INDEX As Integer

        For intLOOP_INDEX = 1 To strMSG.Length
            chrONE = strMSG.Substring(intLOOP_INDEX - 1, 1)
            Dim BYT_TEMP() As Byte
            BYT_TEMP = System.Text.Encoding.ASCII.GetBytes(chrONE)
            Dim INT_CODE As Integer
            INT_CODE = BYT_TEMP(0)

            Call FUNC_SEND_KEYS(prcTARGET, INT_CODE,, True)
        Next

        Return True
    End Function

    Public Sub SUB_SEND_KEYS_PGUP(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_PRIOR, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_PGDN(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_NEXT, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_END(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_END)
    End Sub

    Public Sub SUB_SEND_KEYS_LEFT(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_LEFT, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_UP(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_UP, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_RIGHT(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_RIGHT, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_DOWN(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_DOWN, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_W(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_W, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_A(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_A, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_S(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_S, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_D(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_D, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_X(ByRef prcTARGET As Process, ByVal intWAIT As Integer)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_X, intWAIT)
    End Sub

    Public Sub SUB_SEND_KEYS_SCROLL(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_SCROLL)
    End Sub

    Public Sub SUB_SEND_KEYS_1(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_1)
    End Sub

    Public Sub SUB_SEND_KEYS_2(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_2)

    End Sub
    Public Sub SUB_SEND_KEYS_3(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_3)
    End Sub

    Public Sub SUB_SEND_KEYS_4(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_4)
    End Sub

    Public Sub SUB_SEND_KEYS_5(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_5)
    End Sub

    Public Sub SUB_SEND_KEYS_6(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_6)
    End Sub

    Public Sub SUB_SEND_KEYS_7(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_7)
    End Sub

    Public Sub SUB_SEND_KEYS_8(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_8)
    End Sub

    Public Sub SUB_SEND_KEYS_9(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_9)
    End Sub

    Public Sub SUB_SEND_KEYS_0(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_0)
    End Sub

    Public Sub SUB_SEND_KEYS_MINUS(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_OEM_MINUS)
    End Sub

    Public Sub SUB_SEND_KEYS_OEM_7(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_OEM_7)
    End Sub

    Public Sub SUB_SEND_KEYS_RETURN(ByRef prcTARGET As Process)
        Call FUNC_SEND_KEYS(prcTARGET, ENM_SEND_VK.VK_RETURN)
    End Sub

#Region "内部処理"
    Private Function FUNC_CNV_CHAR_TO_VK(ByVal chrBASE As Char) As ENM_SEND_VK
        Dim enmRET As ENM_SEND_VK

        enmRET = &H0

        Select Case chrBASE
            Case "0"
                enmRET = ENM_SEND_VK.VK_0
            Case "1"
                enmRET = ENM_SEND_VK.VK_1
            Case "2"
                enmRET = ENM_SEND_VK.VK_2
            Case "3"
                enmRET = ENM_SEND_VK.VK_3
            Case "4"
                enmRET = ENM_SEND_VK.VK_4
            Case "5"
                enmRET = ENM_SEND_VK.VK_5
            Case "6"
                enmRET = ENM_SEND_VK.VK_6
            Case "7"
                enmRET = ENM_SEND_VK.VK_7
            Case "8"
                enmRET = ENM_SEND_VK.VK_8
            Case "9"
                enmRET = ENM_SEND_VK.VK_9
            Case "A", "a"
                enmRET = ENM_SEND_VK.VK_A
            Case "B", "b"
                enmRET = ENM_SEND_VK.VK_B
            Case "C", "c"
                enmRET = ENM_SEND_VK.VK_C
            Case "D", "d"
                enmRET = ENM_SEND_VK.VK_D
            Case "E", "e"
                enmRET = ENM_SEND_VK.VK_E
            Case "F", "f"
                enmRET = ENM_SEND_VK.VK_F
            Case "G", "g"
                enmRET = ENM_SEND_VK.VK_G
            Case "H", "h"
                enmRET = ENM_SEND_VK.VK_H
            Case "I", "i"
                enmRET = ENM_SEND_VK.VK_I
            Case "J", "j"
                enmRET = ENM_SEND_VK.VK_J
            Case "K", "k"
                enmRET = ENM_SEND_VK.VK_K
            Case "L", "l"
                enmRET = ENM_SEND_VK.VK_L
            Case "M", "m"
                enmRET = ENM_SEND_VK.VK_M
            Case "N", "n"
                enmRET = ENM_SEND_VK.VK_N
            Case "O", "o"
                enmRET = ENM_SEND_VK.VK_O
            Case "P", "p"
                enmRET = ENM_SEND_VK.VK_P
            Case "Q", "q"
                enmRET = ENM_SEND_VK.VK_Q
            Case "R", "r"
                enmRET = ENM_SEND_VK.VK_R
            Case "S", "s"
                enmRET = ENM_SEND_VK.VK_S
            Case "T", "t"
                enmRET = ENM_SEND_VK.VK_T
            Case "U", "u"
                enmRET = ENM_SEND_VK.VK_U
            Case "V", "v"
                enmRET = ENM_SEND_VK.VK_V
            Case "W", "w"
                enmRET = ENM_SEND_VK.VK_W
            Case "X", "x"
                enmRET = ENM_SEND_VK.VK_X
            Case "Y", "y"
                enmRET = ENM_SEND_VK.VK_Y
            Case "Z", "z"
                enmRET = ENM_SEND_VK.VK_Z
            Case " "
                enmRET = ENM_SEND_VK.VK_SPACE
            Case "/"
                enmRET = ENM_SEND_VK.VK_OEM_2
            Case Else

        End Select

        Return enmRET
    End Function

    Private Function FUNC_SEND_KEYS(ByRef prcTARGET As Process, ByVal enmVK As ENM_SEND_VK, Optional ByVal intWAIT_MSEC As Integer = CST_WAIT_ONE, Optional ByVal blnIME As Boolean = False) As Boolean
        Dim ptrHANDLE As IntPtr
        Dim ptrCHILD As Integer

        If intWAIT_MSEC <= 0 Then
            Return True
        End If

        ptrCHILD = IntPtr.Zero
        'ptrCHILD = FUNC_FIND_WINDOW_EX(prcTARGET.MainWindowHandle, 0, "Edit", "")

        If ptrCHILD > 0 Then
            ptrHANDLE = ptrCHILD
        Else
            ptrHANDLE = prcTARGET.MainWindowHandle
        End If

        If blnIME Then
            Call SendMessage(ptrHANDLE, ENM_SEND_MSG.WM_IME_CHAR, enmVK, 0)
        Else
            Call SendMessage(ptrHANDLE, ENM_SEND_MSG.WM_KEYDOWN, enmVK, 0)
        End If

        Call System.Threading.Thread.Sleep(intWAIT_MSEC)

        If Not blnIME Then
            Call SendMessage(ptrHANDLE, ENM_SEND_MSG.WM_KEYUP, enmVK, 0)
        End If

        Return True
    End Function

    Public Function FUNC_SEND_KEYS_MASK(ByRef PRC_TARGET As Process, ByVal ENM_VK As ENM_SEND_VK, ByVal ENM_MASK As ENM_MASK_KEYS, Optional ByVal intWAIT_MSEC As Integer = CST_WAIT_ONE) As Boolean
        Dim PTR_HANDLE As IntPtr

        If intWAIT_MSEC <= 0 Then
            Return True
        End If

        Dim PTR_CHILD As IntPtr
        PTR_CHILD = IntPtr.Zero
        'PTR_CHILD = FUNC_FIND_WINDOW_EX(PRC_TARGET.MainWindowHandle, 0, "Edit", "")

        If PTR_CHILD = IntPtr.Zero Then
            PTR_HANDLE = PRC_TARGET.MainWindowHandle
        Else
            PTR_HANDLE = PTR_CHILD
        End If

        If ENM_MASK <> ENM_MASK_KEYS.NONE Then
            Dim ENM_VK_MASK As ENM_SEND_VK
            ENM_VK_MASK = FUNC_GET_VK_MASK(ENM_MASK)
            Call SendMessage(PTR_HANDLE, ENM_SEND_MSG.WM_KEYDOWN, ENM_VK_MASK, 0)
        End If

        Call SendMessage(PTR_HANDLE, ENM_SEND_MSG.WM_KEYDOWN, ENM_VK, 0)
        Call System.Threading.Thread.Sleep(intWAIT_MSEC)
        Call SendMessage(PTR_HANDLE, ENM_SEND_MSG.WM_KEYUP, ENM_VK, 0)

        If ENM_MASK <> ENM_MASK_KEYS.NONE Then
            Dim ENM_VK_MASK As ENM_SEND_VK
            ENM_VK_MASK = FUNC_GET_VK_MASK(ENM_MASK)
            Call SendMessage(PTR_HANDLE, ENM_SEND_MSG.WM_KEYUP, ENM_VK_MASK, 0)
        End If

        Return True
    End Function

    Private Function FUNC_GET_VK_MASK(ByVal ENM_MASK As ENM_MASK_KEYS) As ENM_SEND_VK
        Dim ENM_RET As ENM_SEND_VK
        Select Case ENM_MASK
            Case ENM_MASK_KEYS.NONE
                ENM_RET = 0
            Case ENM_MASK_KEYS.SHIFT
                ENM_RET = ENM_SEND_VK.VK_SHIFT
            Case ENM_MASK_KEYS.CTRL
                ENM_RET = ENM_SEND_VK.VK_CONTROL
            Case ENM_MASK_KEYS.ALT
                ENM_RET = ENM_SEND_VK.VK_MENU
            Case Else
                ENM_RET = 0
        End Select

        Return ENM_RET
    End Function
#End Region

End Module
