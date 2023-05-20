Imports System.Threading

Module MOD_GAMEPAD

#Region "モジュール用・定数"
    Const CST_NUMBER_BUTTON_MAX As Integer = 16
    Const CST_BUTTONS_COUNT As Integer = 10
#End Region

#Region "モジュール用・列挙定数"
#End Region

#Region "モジュール用・構造体"
    Private Structure SRT_BUTTON_PUSH
        Public BUTTON_01 As Integer
        Public BUTTON_02 As Integer
        Public BUTTON_03 As Integer
    End Structure

    Private Structure SRT_POV
        Public UP As Boolean
        Public RIGHT As Boolean
        Public DOWN As Boolean
        Public LEFT As Boolean
    End Structure
#End Region

#Region "モジュール用・変数"
    Private SRT_BUTTONS(CST_BUTTONS_COUNT) As SRT_BUTTON_PUSH
    Private INT_MASK_EXCLUSIVE_BUTTON(2) As Integer
    Private FRM_PARENT As FRM_MAIN
    Public ING_CAPTURE_GAMEPAD_ID As Integer
#End Region

#Region "WIN32API"
    ' ***** constant value *****

    ' joyID number
    Private Const JOYSTICKID1 = 0
    Private Const JOYSTICKID2 = 1

    ' *** dwFlag of JOYINFOEX member
    Private Const JOY_RETURNZ As Long = &H4&
    Private Const JOY_RETURNY As Long = &H2&
    Private Const JOY_RETURNX As Long = &H1&
    Private Const JOY_RETURNV As Long = &H20
    Private Const JOY_RETURNU As Long = &H10
    Private Const JOY_RETURNR As Long = &H8&
    Private Const JOY_RETURNPOV As Long = &H40&
    Private Const JOY_RETURNBUTTONS As Long = &H80&
    Public Const JOY_RETURNALL As Long = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)

    ' *** Message from Joystick
    Private Const MM_JOY1BUTTONDOWN As Long = &H3B5
    Private Const MM_JOY1BUTTONUP As Long = &H3B7
    Private Const MM_JOY1MOVE As Long = &H3A0
    Private Const MM_JOY1ZMOVE As Long = &H3A2

    ' *** joyStick API MMResult
    Public Const JOYERR_BASE As Long = 160
    Public Const JOYERR_NOCANDO As Long = (JOYERR_BASE + 6)
    Public Const JOYERR_NOERROR As Long = (0)
    Public Const JOYERR_PARMS As Long = (JOYERR_BASE + 5)
    Public Const JOYERR_UNPLUGGED As Long = (JOYERR_BASE + 7)
    Public Const MMSYSERR_BASE As Long = 0
    Public Const MMSYSERR_NODRIVER As Long = (MMSYSERR_BASE + 6)
    Public Const MMSYSERR_INVALPARAM As Long = (MMSYSERR_BASE + 11)
    Public Const MMSYSERR_BADDEVICEID As Long = (MMSYSERR_BASE + 2)

    ' ***** structure definision *****
    Public Structure JOYINFOEX
        Public dwSize As Integer ' size of structure _
        Public dwFlags As Integer ' flags to indicate what to return _
        Public dwXpos As Integer ' x position _
        Public dwYpos As Integer ' y position _
        Public dwZpos As Integer ' z position _
        Public dwRpos As Integer ' rudder/4th axis position _
        Public dwUpos As Integer ' 5th axis position _
        Public dwVpos As Integer ' 6th axis position _
        Public dwButtons As Integer ' button states _
        Public dwButtonNumber As Integer ' current button number pressed _
        Public dwPOV As Integer ' point of view state _
        Public dwReserved1 As Integer ' reserved for communication between winmm driver _
        Public dwReserved2 As Integer ' reserved for future expansion _
    End Structure

    ' ***** Win32 API declair *****
    Public Declare Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDevs" () As Integer
    Public Declare Function joyGetPosEx Lib "winmm.dll" Alias "joyGetPosEx" (ByVal uJoyID As Integer, ByRef pji As JOYINFOEX) As Integer

    Public Declare Function joySetCapture Lib "winmm.dll" _
    Alias "joySetCapture" (ByVal hwnd As System.IntPtr,
                           ByVal uID As Integer,
                           ByVal uPeriod As Integer,
                           ByVal bChanged As Integer) As Integer

    Public Declare Function joyReleaseCapture Lib "winmm.dll" _
    Alias "joyReleaseCapture" (ByVal uJoyID As Integer) As Integer

#End Region

#Region "引き渡し"
    Public Sub SUB_PARAM_INIT()

        For i = 1 To (SRT_BUTTONS.Length - 1)
            With SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(i)
                SRT_BUTTONS(i).BUTTON_01 = .BUTTON_01
                SRT_BUTTONS(i).BUTTON_02 = .BUTTON_02
                SRT_BUTTONS(i).BUTTON_03 = .BUTTON_03
            End With
        Next

        INT_MASK_EXCLUSIVE_BUTTON(1) = SRT_CURRENT_SETTINGS.GAMEPAD.MASK_EXCLUSIVE_BUTTON(1)
        INT_MASK_EXCLUSIVE_BUTTON(2) = SRT_CURRENT_SETTINGS.GAMEPAD.MASK_EXCLUSIVE_BUTTON(2)

    End Sub
#End Region

#Region "外部参照"
    Public Property GAMEPAD_BUTTONS As Integer
        Get
            Return INT_PAD_PUSH_BUTTONS
        End Get
        Set(ByVal Value As Integer)
            INT_PAD_PUSH_BUTTONS = Value
        End Set
    End Property

    Public Property GAMEPAD_POV As Integer
        Get
            Return INT_PAD_PUSH_POV
        End Get
        Set(ByVal Value As Integer)
            INT_PAD_PUSH_POV = Value
        End Set
    End Property
#End Region

#Region "コントローラ関連"
    Private JoyStickInput As JOYINFOEX
    Private Captured As Boolean = False

    Public Sub SUB_GET_JOYSTICK(ByRef objOWNER As Object, ByVal INT_ID_GAMEPAD As Integer)
        ING_CAPTURE_GAMEPAD_ID = INT_ID_GAMEPAD
        JoyStickInput.dwSize = System.Runtime.InteropServices.Marshal.SizeOf(JoyStickInput)
        JoyStickInput.dwFlags = JOY_RETURNALL
        Call CaptureJoystick(objOWNER)
        FRM_PARENT = objOWNER
        Call SUB_LOOP_THREAD_START()
    End Sub

    Friend Sub SUB_END_JOYSTICK()
        If Captured Then
            Captured = False
            Call System.Threading.Thread.Sleep(100)
            Call joyReleaseCapture(ING_CAPTURE_GAMEPAD_ID)
        End If

        If Not TRD_LOOP_DOEVENTS Is Nothing Then
            'Call TRD_LOOP_DOEVENTS.Abort()
        End If

        'If Not TRD_MAIN Is Nothing Then
        '    Call TRD_MAIN.Abort()
        'End If

        'If Not TRD_EVENT Is Nothing Then
        '    Call TRD_EVENT.Abort()
        'End If
    End Sub

    ' ***** capture Joysticks *****
    Private Sub CaptureJoystick(ByRef objOWNER As Object)

        Dim Result As Integer

        If Captured Then
            'MsgBox("すでにジョイスティックの接続を確保しています")
        Else
            Dim hndTEMP As IntPtr
            hndTEMP = objOWNER.Handle
            Result = joySetCapture(CInt(hndTEMP), ING_CAPTURE_GAMEPAD_ID, 50, False)
            If Result > 0 Then
                'MsgBox("エラーだそうです。" + CStr(Result))
            Else
                Captured = True
                'MsgBox("ジョイスティックの接続を確保しました")
            End If
        End If

    End Sub

#Region "パッドの状態"
    Private INT_PAD_PUSH_BUTTONS As Integer
    Private INT_PAD_PUSH_POV As Integer
#End Region

    Private Function FUNC_GET_PAD_STATE() As Boolean
        INT_PAD_PUSH_BUTTONS = 0
        INT_PAD_PUSH_POV = 0

        Dim INT_RESULT As Integer
        INT_RESULT = joyGetPosEx(ING_CAPTURE_GAMEPAD_ID, JoyStickInput)
        Select Case INT_RESULT
            Case MMSYSERR_NODRIVER
                Return False
            Case MMSYSERR_INVALPARAM
                Return False
            Case MMSYSERR_BADDEVICEID
                Return False
            Case JOYERR_UNPLUGGED
                Return False
            Case JOYERR_PARMS
                Return False
            Case Else

        End Select

        INT_PAD_PUSH_BUTTONS = JoyStickInput.dwButtons
        INT_PAD_PUSH_POV = JoyStickInput.dwPOV

        Return True
    End Function

    Public Function FUNC_GET_JOYPOS(ByRef intBUTTONS As Integer) As Boolean
        Dim intRESULT As Integer
        intRESULT = joyGetPosEx(ING_CAPTURE_GAMEPAD_ID, JoyStickInput)

        Select Case intRESULT
            Case MMSYSERR_NODRIVER
                Return False
            Case MMSYSERR_INVALPARAM
                Return False
            Case MMSYSERR_BADDEVICEID
                Return False
            Case JOYERR_UNPLUGGED
                Return False
            Case JOYERR_PARMS
                Return False
            Case Else

        End Select

        intBUTTONS = JoyStickInput.dwButtons
        Return True
    End Function

    Private Function FUNC_GET_BUTTON_ROW(ByVal intHEX_BUTTON As Integer) As Boolean()
        Dim blnRET(CST_NUMBER_BUTTON_MAX) As Boolean
        Dim intLOOP_INDEX As Integer
        Dim intCHECK As Integer
        Dim intCALC As Integer

        blnRET(0) = False
        For intLOOP_INDEX = 0 To ((blnRET.Length - 1) - 1)
            intCHECK = 2 ^ intLOOP_INDEX
            intCALC = intHEX_BUTTON And intCHECK
            blnRET(intLOOP_INDEX + 1) = (intCALC <> 0)
        Next

        Return blnRET
    End Function

    Private Function FUNC_GET_POV_ROW(ByVal INT_HEX_POV As Integer) As SRT_POV
        Dim BLN_POV(4) As Boolean

        For intLOOP_INDEX = 0 To ((BLN_POV.Length - 1) - 1)
            Dim INT_CHECK As Integer
            INT_CHECK = 2 ^ intLOOP_INDEX

            Dim INT_CALC As Integer
            INT_CALC = INT_HEX_POV And INT_CHECK
            BLN_POV(intLOOP_INDEX + 1) = (INT_CALC <> 0)
        Next

        Dim SRT_RET As SRT_POV
        With SRT_RET
            .UP = BLN_POV(1)
            .RIGHT = BLN_POV(2)
            .DOWN = BLN_POV(3)
            .LEFT = BLN_POV(4)
        End With

        Return SRT_RET
    End Function

#End Region

#Region "スレッド"
    Private TRD_LOOP_DOEVENTS As Threading.Thread
    Private TRD_MAIN As Threading.Thread
    Private TRD_EVENT As Threading.Thread

    Private BLN_BUTTON_DOWN As Boolean

    Private Sub SUB_LOOP_THREAD_START()
        Dim cts As CancellationTokenSource
        cts = New CancellationTokenSource()


        If Not TRD_LOOP_DOEVENTS Is Nothing Then
            'TRD_LOOP_DOEVENTS.Abort()
        End If

        TRD_LOOP_DOEVENTS = Nothing
        TRD_LOOP_DOEVENTS = New Threading.Thread(New Threading.ThreadStart(AddressOf SUB_GAMEPAD_BUTTON_THREAD))
        TRD_LOOP_DOEVENTS.IsBackground = True
        TRD_LOOP_DOEVENTS.Start()
        'BLN_BUTTON_DOWN = False
        'TRD_LOOP_DOEVENTS = Nothing
        'TRD_LOOP_DOEVENTS = New Threading.Thread(New Threading.ThreadStart(AddressOf SUB_DOEVENTS_LOOP))
        'TRD_LOOP_DOEVENTS.IsBackground = True
        'TRD_LOOP_DOEVENTS.Start()

        'TRD_MAIN = Nothing
        'TRD_MAIN = New Threading.Thread(New Threading.ThreadStart(AddressOf SUB_GAMEPAD_MAIN_THREAD))
        'TRD_MAIN.IsBackground = True
        'TRD_MAIN.Start()

        'TRD_EVENT = Nothing
        'TRD_EVENT = New Threading.Thread(New Threading.ThreadStart(AddressOf SUB_GAMEPAD_EVENT_THREAD))
        'TRD_EVENT.IsBackground = True
        'TRD_EVENT.Start()
    End Sub

    Private Sub SUB_GAMEPAD_MAIN_THREAD()

        Const CST_INTERVAL As Integer = 40

        Do
            If Not Captured Then
                Exit Do
            End If

            Call FUNC_GET_PAD_STATE()

            Try
                Call System.Threading.Thread.Sleep(CST_INTERVAL)
            Catch ex As Exception
                'スルー
            End Try
        Loop
    End Sub

    Private Sub SUB_GAMEPAD_BUTTON_THREAD()

        ReDim BLN_BUTTONS_PUSH_BEFORE(CST_NUMBER_BUTTON_MAX)
        Do
            If Not Captured Then
                Exit Do
            End If

            Call SUB_CHECK_AND_DO_BUTTON()

            Try
                Const CST_INTERVAL As Integer = 20
                Call System.Threading.Thread.Sleep(CST_INTERVAL)
            Catch ex As Exception
                'スルー
            End Try
        Loop

    End Sub

    Dim BLN_BUTTONS_PUSH_BEFORE() As Boolean
    Private Sub SUB_CHECK_AND_DO_BUTTON()

        Dim INT_BUTTON__HEX As Integer
        If Not FUNC_GET_JOYPOS(INT_BUTTON__HEX) Then
            Exit Sub
        End If

        Dim BLN_BUTTONS_PUSH() As Boolean
        BLN_BUTTONS_PUSH = FUNC_GET_BUTTON_ROW(INT_BUTTON__HEX)

        Dim INT_PUSH_BUTTON_NUMBER() As Integer
        ReDim INT_PUSH_BUTTON_NUMBER(0)
        For i = 1 To (BLN_BUTTONS_PUSH.Length - 1)
            If BLN_BUTTONS_PUSH_BEFORE(i) Then '前回押されて
                If Not BLN_BUTTONS_PUSH(i) Then '今回押されてなければ
                    Dim INT_INDEX As Integer
                    INT_INDEX = INT_PUSH_BUTTON_NUMBER.Length
                    ReDim Preserve INT_PUSH_BUTTON_NUMBER(INT_INDEX)
                    INT_PUSH_BUTTON_NUMBER(INT_INDEX) = i
                End If
            End If
        Next

        For i = 1 To (INT_PUSH_BUTTON_NUMBER.Length - 1)
            Call SUB_EVENT_PUSH_BUTTON(INT_PUSH_BUTTON_NUMBER(i), BLN_BUTTONS_PUSH_BEFORE, BLN_BUTTONS_PUSH, INT_MASK_EXCLUSIVE_BUTTON(1), INT_MASK_EXCLUSIVE_BUTTON(2))
        Next

        For i = 1 To (BLN_BUTTONS_PUSH.Length - 1)
            BLN_BUTTONS_PUSH_BEFORE(i) = BLN_BUTTONS_PUSH(i)
        Next

    End Sub

    Private Sub SUB_EVENT_PUSH_BUTTON(ByVal INT_BUTTON_INDEX As Integer, ByRef BLN_DOWN_BUTTONS() As Boolean, ByRef BLN_UP_BUTTONS() As Boolean, ByVal INT_SIMULTANEOUSLY_PRESS_01 As Integer, ByVal INT_SIMULTANEOUSLY_PRESS_02 As Integer)

        For i = 1 To (SRT_BUTTONS.Length - 1)
            Dim INT_BUTTON_01 As Integer
            INT_BUTTON_01 = SRT_BUTTONS(i).BUTTON_01
            Dim INT_BUTTON_02 As Integer
            INT_BUTTON_02 = SRT_BUTTONS(i).BUTTON_02
            Dim INT_BUTTON_03 As Integer
            INT_BUTTON_03 = SRT_BUTTONS(i).BUTTON_03

            Dim INT_PUCH_CHECK_INDEX As Integer
            Dim INT_MASK_CHECK_INDEX_01 As Integer
            Dim INT_MASK_CHECK_INDEX_02 As Integer

            If INT_BUTTON_03 > 0 Then
                INT_PUCH_CHECK_INDEX = INT_BUTTON_03
                INT_MASK_CHECK_INDEX_01 = INT_BUTTON_01
                INT_MASK_CHECK_INDEX_02 = INT_BUTTON_02
            Else
                If INT_BUTTON_02 > 0 Then
                    INT_PUCH_CHECK_INDEX = INT_BUTTON_02
                    INT_MASK_CHECK_INDEX_01 = INT_BUTTON_01
                    INT_MASK_CHECK_INDEX_02 = 0
                Else
                    If INT_BUTTON_01 > 0 Then
                        INT_PUCH_CHECK_INDEX = INT_BUTTON_01
                        INT_MASK_CHECK_INDEX_01 = 0
                        INT_MASK_CHECK_INDEX_02 = 0
                    Else
                        INT_PUCH_CHECK_INDEX = 0
                    End If
                End If
            End If

            If INT_PUCH_CHECK_INDEX = 0 Then
                Continue For
            End If

            If INT_PUCH_CHECK_INDEX <> INT_BUTTON_INDEX Then
                Continue For
            End If

            Dim BLN_CHECK_ALL_BUTTON As Boolean
            BLN_CHECK_ALL_BUTTON = True
            For j = 1 To (BLN_UP_BUTTONS.Length - 1)
                If BLN_UP_BUTTONS(j) Then 'ボタンが押されている
                    Select Case True
                        Case (INT_MASK_CHECK_INDEX_01 = 0 And INT_MASK_CHECK_INDEX_02 = 0) '両マスク設定がない場合は
                            BLN_CHECK_ALL_BUTTON = False 'どのボタンも押されていてはいけない
                        Case (INT_MASK_CHECK_INDEX_01 <> 0 And INT_MASK_CHECK_INDEX_02 = 0) 'マスク1のみ設定されている場合は
                            If j <> INT_MASK_CHECK_INDEX_01 Then
                                BLN_CHECK_ALL_BUTTON = False 'マスク1と違うキーは押されてはいけない
                            End If
                        Case (INT_MASK_CHECK_INDEX_01 = 0 And INT_MASK_CHECK_INDEX_02 <> 0) 'マスク2のみ設定されている場合は
                            If j <> INT_MASK_CHECK_INDEX_02 Then
                                BLN_CHECK_ALL_BUTTON = False 'マスク1と違うキーは押されてはいけない
                            End If
                        Case (INT_MASK_CHECK_INDEX_01 <> 0 And INT_MASK_CHECK_INDEX_02 <> 0) 'マスク1、マスク2が設定されている場合
                            If (j <> INT_MASK_CHECK_INDEX_01 And j <> INT_MASK_CHECK_INDEX_02) Then
                                BLN_CHECK_ALL_BUTTON = False 'マスク1、マスク2と違うキーは押されてはいけない
                            End If
                    End Select
                End If
            Next

            If Not BLN_CHECK_ALL_BUTTON Then
                Continue For
            End If

            Dim BLN_MASK_01 As Boolean
            If INT_MASK_CHECK_INDEX_01 = 0 Then 'マスク1が非設定の場合
                BLN_MASK_01 = True
            Else
                If BLN_DOWN_BUTTONS(INT_MASK_CHECK_INDEX_01) And BLN_UP_BUTTONS(INT_MASK_CHECK_INDEX_01) Then
                    BLN_MASK_01 = True
                Else
                    BLN_MASK_01 = False
                End If
            End If

            Dim BLN_MASK_02 As Boolean
            If INT_MASK_CHECK_INDEX_02 = 0 Then 'マスク2が非設定の場合
                BLN_MASK_02 = True
            Else
                If BLN_DOWN_BUTTONS(INT_MASK_CHECK_INDEX_01) And BLN_UP_BUTTONS(INT_MASK_CHECK_INDEX_02) Then
                    BLN_MASK_02 = True
                Else
                    BLN_MASK_02 = False
                End If
            End If

            If BLN_MASK_01 And BLN_MASK_02 Then
                Call SUB_EVENT(i)
            End If
        Next

    End Sub

    Private Sub SUB_GAMEPAD_EVENT_THREAD()
        Const CST_INTERVAL As Integer = 100

        Dim INT_PAD_PUSH_BUTTONS_BF As Integer
        Dim INT_PAD_PUSH_POV_BF As Integer

        INT_PAD_PUSH_BUTTONS_BF = 0
        INT_PAD_PUSH_POV_BF = 0

        Do
            If Not Captured Then
                Exit Do
            End If

            Try
                Call System.Threading.Thread.Sleep(CST_INTERVAL)
            Catch ex As Exception
                'スルー
            End Try

            If Not FUNC_CHECK_EVENT(INT_PAD_PUSH_BUTTONS_BF, INT_PAD_PUSH_POV_BF) Then
                Continue Do
            End If

            Dim BLN_BUTTONS_PUSH() As Boolean
            BLN_BUTTONS_PUSH = FUNC_GET_BUTTON_ROW(INT_PAD_PUSH_BUTTONS)
            Dim SRT_POV_PUSH As SRT_POV
            SRT_POV_PUSH = FUNC_GET_POV_ROW(INT_PAD_PUSH_POV)
            Dim INT_EVENT As Integer
            INT_EVENT = FUNC_GET_EVENT(BLN_BUTTONS_PUSH)
            If INT_EVENT > 0 Then
                'Call SUB_EVENT(INT_EVENT)
            End If

            INT_PAD_PUSH_BUTTONS_BF = INT_PAD_PUSH_BUTTONS
            INT_PAD_PUSH_POV_BF = INT_PAD_PUSH_POV
        Loop
    End Sub

    Private Function FUNC_CHECK_EVENT(ByVal INT_PAD_PUSH_BUTTONS_BF As Integer, ByVal INT_PAD_PUSH_POV_BF As Integer) As Boolean

        If INT_PAD_PUSH_BUTTONS_BF = INT_PAD_PUSH_BUTTONS _
            And INT_PAD_PUSH_POV_BF = INT_PAD_PUSH_POV Then
            Return False
        End If

        Return True
    End Function

    Private Sub SUB_DOEVENTS_LOOP()
        Dim stwTHREAD As Stopwatch
        Dim intMSEC_THREAD As Integer
        Dim intMSEC_THREAD_SUB As Integer
        Const cstMSEC_INTERVAL_THREAD As Integer = 40

        Dim sw As System.Diagnostics.Stopwatch
        Dim intElapsed As Integer
        Dim intCHECK_TIME As Integer
        Dim blnCHECK_TIME As Boolean

        Dim blnBUTTONS_ROW(CST_NUMBER_BUTTON_MAX) As Boolean

        intCHECK_TIME = 100
        blnCHECK_TIME = True

        sw = New System.Diagnostics.Stopwatch
        Call sw.Start()

        stwTHREAD = New System.Diagnostics.Stopwatch
        Call stwTHREAD.Start()
        Call stwTHREAD.Stop()

        Dim intBUTTONS_HEX_BEFORE As Integer
        intBUTTONS_HEX_BEFORE = 0

        Do
            If Not Captured Then
                Exit Do
            End If

            Call stwTHREAD.Restart()

            Dim intBUTTONS_HEX As Integer
            If Not FUNC_GET_JOYPOS(intBUTTONS_HEX) Then
                Exit Do
            End If

            If intBUTTONS_HEX <> intBUTTONS_HEX_BEFORE Then 'ボタンの状態が変更されたら
                BLN_BUTTON_DOWN = False '次の押下を有効にする
                intBUTTONS_HEX_BEFORE = intBUTTONS_HEX
            End If

            Dim blnEVENT_ENABLED As Boolean
            blnEVENT_ENABLED = True 'デフォルト有効

            If BLN_BUTTON_DOWN Then 'すでに押下イベント中なら
                blnEVENT_ENABLED = False '無効
            End If

            If intBUTTONS_HEX = 0 Then
                blnEVENT_ENABLED = False '無効
            End If

            If blnEVENT_ENABLED Then
                Call sw.Stop()
                intElapsed = sw.ElapsedMilliseconds
                If blnCHECK_TIME = False And (intElapsed > intCHECK_TIME) Then
                    blnCHECK_TIME = True
                    Call sw.Restart()
                Else
                    Call sw.Start()
                End If
            Else
                blnCHECK_TIME = False
            End If

            If Not blnCHECK_TIME Then
                blnEVENT_ENABLED = False
            End If

            If blnEVENT_ENABLED Then
                BLN_BUTTON_DOWN = True '処理中
                blnCHECK_TIME = False

                blnBUTTONS_ROW = FUNC_GET_BUTTON_ROW(intBUTTONS_HEX)
                Dim INT_EVENT As Integer
                INT_EVENT = FUNC_GET_EVENT(blnBUTTONS_ROW)
                If INT_EVENT > 0 Then
                    Call SUB_EVENT(INT_EVENT)
                End If
            End If

            Call stwTHREAD.Stop()
            intMSEC_THREAD = stwTHREAD.ElapsedMilliseconds
            intMSEC_THREAD_SUB = (cstMSEC_INTERVAL_THREAD - intMSEC_THREAD)
            If intMSEC_THREAD_SUB > 0 Then
                Try
                    Call System.Threading.Thread.Sleep(intMSEC_THREAD_SUB)
                Catch ex As Exception
                    'スルー
                End Try
            End If
        Loop
    End Sub

    Private Function FUNC_GET_EVENT(ByRef BLN_PUSH_ROW As Boolean()) As Integer

        For i = 1 To (SRT_BUTTONS.Length - 1)
            Dim INT_BUTTON_01 As Integer
            INT_BUTTON_01 = SRT_BUTTONS(i).BUTTON_01
            Dim INT_BUTTON_02 As Integer
            INT_BUTTON_02 = SRT_BUTTONS(i).BUTTON_02
            Dim INT_BUTTON_03 As Integer
            INT_BUTTON_03 = SRT_BUTTONS(i).BUTTON_03

            Dim INT_BUTTON As Integer
            INT_BUTTON = 0
            If INT_BUTTON = 0 Then
                If INT_BUTTON_01 = 0 And INT_BUTTON_02 = 0 Then
                    INT_BUTTON = INT_BUTTON_03
                End If
            End If
            If INT_BUTTON = 0 Then
                If INT_BUTTON_01 = 0 And INT_BUTTON_03 = 0 Then
                    INT_BUTTON = INT_BUTTON_02
                End If
            End If
            If INT_BUTTON = 0 Then
                If INT_BUTTON_02 = 0 And INT_BUTTON_03 = 0 Then
                    INT_BUTTON = INT_BUTTON_01
                End If
            End If

            If INT_BUTTON > 0 Then
                If BLN_PUSH_ROW(INT_BUTTON) And FUNC_NO_PUSH_BUTTON(BLN_PUSH_ROW, INT_BUTTON) Then
                    Return i
                End If
            Else
                If INT_BUTTON_01 <= 0 Then
                    If BLN_PUSH_ROW(INT_BUTTON_02) And BLN_PUSH_ROW(INT_BUTTON_03) And FUNC_NO_PUSH_BUTTONS(BLN_PUSH_ROW, INT_BUTTON_02, INT_BUTTON_03) Then
                        Return i
                    End If
                End If
                If INT_BUTTON_02 <= 0 Then
                    If BLN_PUSH_ROW(INT_BUTTON_01) And BLN_PUSH_ROW(INT_BUTTON_03) And FUNC_NO_PUSH_BUTTONS(BLN_PUSH_ROW, INT_BUTTON_01, INT_BUTTON_03) Then
                        Return i
                    End If
                End If
                If INT_BUTTON_03 <= 0 Then
                    If BLN_PUSH_ROW(INT_BUTTON_01) And BLN_PUSH_ROW(INT_BUTTON_02) And FUNC_NO_PUSH_BUTTONS(BLN_PUSH_ROW, INT_BUTTON_01, INT_BUTTON_02) Then
                        Return i
                    End If
                End If

                If BLN_PUSH_ROW(INT_BUTTON_01) And BLN_PUSH_ROW(INT_BUTTON_02) And BLN_PUSH_ROW(INT_BUTTON_03) Then
                    Return i
                End If
            End If
        Next

        Return 0
    End Function

    Private Function FUNC_NO_PUSH_BUTTON(ByRef BLN_PUSH_ROW As Boolean(), ByVal INT_BUTTONS As Integer) As Boolean

        For i = 1 To (BLN_PUSH_ROW.Length - 1)
            If INT_BUTTONS = i Then
                Continue For
            End If

            If BLN_PUSH_ROW(i) Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Function FUNC_NO_PUSH_BUTTONS(ByRef BLN_PUSH_ROW As Boolean(), ByVal INT_BUTTONS_01 As Integer, ByVal INT_BUTTONS_02 As Integer) As Boolean

        For i = 1 To (BLN_PUSH_ROW.Length - 1)
            If INT_BUTTONS_01 = i Then
                Continue For
            End If
            If INT_BUTTONS_02 = i Then
                Continue For
            End If

            If BLN_PUSH_ROW(i) Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Sub SUB_EVENT(ByVal INT_EVENT As Integer)

        Select Case SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(INT_EVENT).KIND
            Case 0
                Dim INT_KEY_PATTERN As Integer
                INT_KEY_PATTERN = FUNC_GET_KEY_PATTERN(INT_EVENT)
                Dim INT_MASK_PATTERN As Integer
                INT_MASK_PATTERN = FUNC_GET_MASK_PATTERN(INT_EVENT)
                Call SUB_SEND_KEY(INT_KEY_PATTERN, INT_MASK_PATTERN)
            Case 1
                Dim INT_X As Integer
                INT_X = SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(INT_EVENT).MOUSE_X
                Dim INT_Y As Integer
                INT_Y = SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(INT_EVENT).MOUSE_Y
                Call SUB_SET_MOUSE_POINTER_CLIENT(INT_X, INT_Y)
            Case Else

        End Select
    End Sub

    Private Function FUNC_GET_KEY_PATTERN(ByVal INT_EVENT As Integer) As Integer
        Dim INT_RET As Integer
        With SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(INT_EVENT)
            INT_RET = .KEY_02
        End With

        Return INT_RET
    End Function

    Private Function FUNC_GET_MASK_PATTERN(ByVal INT_EVENT As Integer) As Integer
        Dim INT_RET As Integer
        With SRT_CURRENT_SETTINGS.GAMEPAD.ALLOCATION(INT_EVENT)
            INT_RET = .KEY_01
        End With

        Return INT_RET
    End Function

    Private Sub SUB_SEND_KEY(ByVal INT_KEY_PATTERN As Integer, ByVal INT_MASK_PATTERN As Integer)
        Dim ENM_SEND_KEY As ENM_SEND_VK
        ENM_SEND_KEY = FUNC_GET_SEND_VK(INT_KEY_PATTERN)

        Dim ENM_SEND_MASK As ENM_MASK_KEYS
        ENM_SEND_MASK = FUNC_GET_SEND_MASK(INT_MASK_PATTERN)

        If ENM_SEND_KEY <> 0 Then
            Call MOD_SEND_KEYS.FUNC_SEND_KEYS_MASK(PRC_TARGET, ENM_SEND_KEY, ENM_SEND_MASK)
        End If

    End Sub

    Private Function FUNC_GET_SEND_MASK(ByVal INT_MASK_PATTERN As Integer) As ENM_MASK_KEYS
        Dim ENM_RET As ENM_MASK_KEYS

        Select Case INT_MASK_PATTERN
            Case 1
                ENM_RET = ENM_MASK_KEYS.SHIFT
            Case 2
                ENM_RET = ENM_MASK_KEYS.CTRL
            Case 3
                ENM_RET = ENM_MASK_KEYS.ALT
            Case Else
                ENM_RET = ENM_MASK_KEYS.NONE
        End Select

        Return ENM_RET
    End Function

    Private Function FUNC_GET_SEND_VK(ByVal INT_KEY_PATTERN As Integer) As ENM_SEND_VK
        Dim ENM_RET As ENM_SEND_VK
        Select Case INT_KEY_PATTERN
            Case 1
                ENM_RET = ENM_SEND_VK.VK_1
            Case 2
                ENM_RET = ENM_SEND_VK.VK_2
            Case 3
                ENM_RET = ENM_SEND_VK.VK_3
            Case 4
                ENM_RET = ENM_SEND_VK.VK_4
            Case 5
                ENM_RET = ENM_SEND_VK.VK_5
            Case 6
                ENM_RET = ENM_SEND_VK.VK_6
            Case 7
                ENM_RET = ENM_SEND_VK.VK_7
            Case 8
                ENM_RET = ENM_SEND_VK.VK_8
            Case 9
                ENM_RET = ENM_SEND_VK.VK_9
            Case 10
                ENM_RET = ENM_SEND_VK.VK_0
            Case 11
                ENM_RET = ENM_SEND_VK.VK_OEM_MINUS
            Case 12
                ENM_RET = ENM_SEND_VK.VK_OEM_7
            Case 13
                ENM_RET = ENM_SEND_VK.VK_F1
            Case 14
                ENM_RET = ENM_SEND_VK.VK_F2
            Case 15
                ENM_RET = ENM_SEND_VK.VK_F3
            Case 16
                ENM_RET = ENM_SEND_VK.VK_F4
            Case 17
                ENM_RET = ENM_SEND_VK.VK_F5
            Case 18
                ENM_RET = ENM_SEND_VK.VK_F6
            Case 19
                ENM_RET = ENM_SEND_VK.VK_F7
            Case 20
                ENM_RET = ENM_SEND_VK.VK_F8
            Case 21
                ENM_RET = ENM_SEND_VK.VK_F9
            Case 22
                ENM_RET = ENM_SEND_VK.VK_F10
            Case 23
                ENM_RET = ENM_SEND_VK.VK_F11
            Case 24
                ENM_RET = ENM_SEND_VK.VK_F12
            Case 25
                ENM_RET = ENM_SEND_VK.VK_A
            Case 26
                ENM_RET = ENM_SEND_VK.VK_B
            Case 27
                ENM_RET = ENM_SEND_VK.VK_C
            Case 28
                ENM_RET = ENM_SEND_VK.VK_D
            Case 29
                ENM_RET = ENM_SEND_VK.VK_E
            Case 30
                ENM_RET = ENM_SEND_VK.VK_F
            Case 31
                ENM_RET = ENM_SEND_VK.VK_G
            Case 32
                ENM_RET = ENM_SEND_VK.VK_H
            Case 33
                ENM_RET = ENM_SEND_VK.VK_I
            Case 34
                ENM_RET = ENM_SEND_VK.VK_J
            Case 35
                ENM_RET = ENM_SEND_VK.VK_K
            Case 36
                ENM_RET = ENM_SEND_VK.VK_L
            Case 37
                ENM_RET = ENM_SEND_VK.VK_M
            Case 38
                ENM_RET = ENM_SEND_VK.VK_N
            Case 39
                ENM_RET = ENM_SEND_VK.VK_O
            Case 40
                ENM_RET = ENM_SEND_VK.VK_P
            Case 41
                ENM_RET = ENM_SEND_VK.VK_Q
            Case 42
                ENM_RET = ENM_SEND_VK.VK_R
            Case 43
                ENM_RET = ENM_SEND_VK.VK_S
            Case 44
                ENM_RET = ENM_SEND_VK.VK_T
            Case 45
                ENM_RET = ENM_SEND_VK.VK_U
            Case 46
                ENM_RET = ENM_SEND_VK.VK_V
            Case 47
                ENM_RET = ENM_SEND_VK.VK_W
            Case 48
                ENM_RET = ENM_SEND_VK.VK_X
            Case 49
                ENM_RET = ENM_SEND_VK.VK_Y
            Case 50
                ENM_RET = ENM_SEND_VK.VK_Z
            Case 51
                ENM_RET = ENM_SEND_VK.VK_NUMPAD0
            Case 52
                ENM_RET = ENM_SEND_VK.VK_NUMPAD1
            Case 53
                ENM_RET = ENM_SEND_VK.VK_NUMPAD2
            Case 54
                ENM_RET = ENM_SEND_VK.VK_NUMPAD3
            Case 55
                ENM_RET = ENM_SEND_VK.VK_NUMPAD4
            Case 56
                ENM_RET = ENM_SEND_VK.VK_NUMPAD5
            Case 57
                ENM_RET = ENM_SEND_VK.VK_NUMPAD6
            Case 58
                ENM_RET = ENM_SEND_VK.VK_NUMPAD7
            Case 59
                ENM_RET = ENM_SEND_VK.VK_NUMPAD8
            Case 60
                ENM_RET = ENM_SEND_VK.VK_NUMPAD9
            Case 61
                ENM_RET = ENM_SEND_VK.VK_MULTIPLY
            Case 62
                ENM_RET = ENM_SEND_VK.VK_ADD
            Case 63
                ENM_RET = ENM_SEND_VK.VK_SUBTRACT
            Case 64
                ENM_RET = ENM_SEND_VK.VK_DECIMAL
            Case 65
                ENM_RET = ENM_SEND_VK.VK_DIVIDE
            Case Else
                ENM_RET = 0
        End Select

        Return ENM_RET
    End Function
#End Region

End Module
