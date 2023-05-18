Imports System.ComponentModel

Public Delegate Sub WPF_SETTING_DELEGATE()
Public Delegate Sub WPF_SETTING_DELEGATE_REFRESH_TEST_GAPEPAD(ByVal STR_POV As String, ByVal STR_BUTTON As String)

Public Class WPF_SETTING

#Region "画面用構造体"
    Public Structure SRT_GAMEPAD_ALLOCATION_CONTROLS
        Public BUTTON() As System.Windows.Controls.ComboBox
        Public KIND As System.Windows.Controls.ComboBox
        Public KEY() As System.Windows.Controls.ComboBox
        Public MOUSE_X As System.Windows.Controls.TextBox
        Public MOUSE_Y As System.Windows.Controls.TextBox
    End Structure
#End Region

#Region "画面用変数"
    Public SRT_GAC() As SRT_GAMEPAD_ALLOCATION_CONTROLS
#End Region

#Region "プロパティ用変数"
    Private BLN_PROPERTY_CHECK_CLOSED As Boolean = False
    Private INT_PROPERTY_TEST_GAMEPAD_ID As Integer
#End Region

#Region "プロパティ"
    Public Property CHECK_CLOSED As Boolean
        Get
            Return BLN_PROPERTY_CHECK_CLOSED
        End Get
        Set(ByVal value As Boolean)
            BLN_PROPERTY_CHECK_CLOSED = value
        End Set
    End Property

    Public Property TEST_GAMEPAD_ID As Integer
        Get
            Return INT_PROPERTY_TEST_GAMEPAD_ID
        End Get
        Set(ByVal value As Integer)
            INT_PROPERTY_TEST_GAMEPAD_ID = value
        End Set
    End Property


#End Region

#Region "初期化・終了処理"
    Private Sub SUB_CTRL_NEW_INIT()
        'Me.Icon = FUNC_GET_IMAGESOURCE(My.Resources.SUPPORT_GAMEPAD_FFXIV)

        'ウィンドウをマウスのドラッグで移動できるようにする
        AddHandler Me.MouseLeftButtonDown, Sub(sender, e) Me.DragMove()

        Call SUB_INIT_GAMEPAD_ALLOCATION_CONTROLS()
        Call SUB_REFRESH_KIND(Me)
        Call SUB_REFRESH_KEY_MASK(Me)
        Call SUB_REFRESH_KEY(Me)
        Call SUB_GET_ITEM_TEST_GANMEPAD_ID()
        Call SUB_START_TIMER()
    End Sub

    Private Sub SUB_CTRL_VIEW_INIT()

    End Sub

    Private Sub SUB_CTRL_VALUE_INIT()

        TXT_TEST_SEND_STR.Text = "/grouppose"

        TXT_TEST_OPERATION_MOUSE_X.Text = CStr(100)
        TXT_TEST_OPERATION_MOUSE_Y.Text = CStr(100)
    End Sub
#End Region

#Region "実行処理群"
    Private Sub SUB_OK()
        SRT_CURRENT_SETTINGS = FUNC_GET_SETTINGS_WPF(Me)
        Call MOD_GAMEPAD.SUB_PARAM_INIT()

        Call FUNC_SET_SETTING()

        Call Me.Close()
    End Sub

    Private Sub SUB_CANCEL()
        Call Me.Close()
    End Sub

    Private Sub SUB_INIT_SETTING()
        Dim SRT_INI As SRT_SETTINGS
        SRT_INI = FUNC_GET_INITAL_SETTINGS()
        Call SUB_SET_SETTINGS_WPF(Me, SRT_INI)
    End Sub

    Private Sub SUB_TEST_SEND_KEY()

        If PRC_TARGET Is Nothing Then
            Exit Sub
        End If

        'Call SUB_FOREGROUND_WINDOW(PRC_TARGET)
        'Call System.Threading.Thread.Sleep(100)
        'Call FUNC_SEND_KEYS_MASK(PRC_TARGET, ENM_SEND_VK.VK_S, ENM_MASK_KEYS.CTRL)
        'Call FUNC_SEND_KEYS_MASK(PRC_TARGET, ENM_SEND_VK.VK_S, ENM_MASK_KEYS.ALT)
        'Call FUNC_SEND_KEYS_MASK(PRC_TARGET, ENM_SEND_VK.VK_F4, ENM_MASK_KEYS.NONE)

        'Call FUNC_SEND_KEYS_STRING_02(PRC_TARGET, "abcdef")
        Call SUB_SEND_KEYS_1(PRC_TARGET)
    End Sub

    Private Sub SUB_TEST_SEND_TEXT()
        If PRC_TARGET Is Nothing Then
            Exit Sub
        End If

        Dim STR_COMMAND As String
        STR_COMMAND = TXT_TEST_SEND_STR.Text

        If STR_COMMAND = "" Then
            Exit Sub
        End If

        Call SUB_FOREGROUND_WINDOW(PRC_TARGET)

        Call SUB_SEND_KEYS_RETURN(PRC_TARGET)

        Call FUNC_SEND_KEYS_STRING_02(PRC_TARGET, STR_COMMAND)

        Call SUB_SEND_KEYS_RETURN(PRC_TARGET)
    End Sub

    Private Sub SUB_TEST_OPERATION_MOUSE()
        Dim INT_X As Integer
        INT_X = FUNC_VALUE_CONVERT_NUMERIC_INT(TXT_TEST_OPERATION_MOUSE_X.Text)
        Dim INT_Y As Integer
        INT_Y = FUNC_VALUE_CONVERT_NUMERIC_INT(TXT_TEST_OPERATION_MOUSE_Y.Text)

        Call SUB_SET_MOUSE_POINTER_CLIENT(INT_X, INT_Y)
    End Sub

    Private Sub SUB_CHANGE_KIND(ByVal INT_INDEX As Integer)
        Dim INT_KIND As Integer
        INT_KIND = Me.SRT_GAC(INT_INDEX).KIND.SelectedIndex

        Dim BLN_ENABLED_KEY As Boolean
        Dim BLN_ENABLED_MOUSE As Boolean

        Select Case INT_KIND
            Case 0
                BLN_ENABLED_KEY = True
                BLN_ENABLED_MOUSE = False
            Case 1
                BLN_ENABLED_KEY = False
                BLN_ENABLED_MOUSE = True
            Case Else
                BLN_ENABLED_KEY = False
                BLN_ENABLED_MOUSE = False
        End Select

        For i = 1 To (Me.SRT_GAC(INT_INDEX).KEY.Length - 1)
            Me.SRT_GAC(INT_INDEX).KEY(i).IsEnabled = BLN_ENABLED_KEY
        Next
        Me.SRT_GAC(INT_INDEX).MOUSE_X.IsEnabled = BLN_ENABLED_MOUSE
        Me.SRT_GAC(INT_INDEX).MOUSE_Y.IsEnabled = BLN_ENABLED_MOUSE

    End Sub
#End Region

#Region "その他・内部処理"
    Private Sub SUB_INIT_GAMEPAD_ALLOCATION_CONTROLS()

        ReDim SRT_GAC(CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT)

        For i = 1 To (SRT_GAC.Length - 1)
            With SRT_GAC(i)
                ReDim .BUTTON(3)
                ReDim .KEY(2)
                Select Case i
                    Case 1
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_01_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_01_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_01_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_01_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_01_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_01_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_01_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_01_MOUSE_Y
                    Case 2
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_02_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_02_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_02_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_02_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_02_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_02_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_02_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_02_MOUSE_Y
                    Case 3
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_03_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_03_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_03_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_03_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_03_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_03_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_03_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_03_MOUSE_Y
                    Case 4
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_04_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_04_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_04_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_04_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_04_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_04_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_04_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_04_MOUSE_Y
                    Case 5
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_05_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_05_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_05_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_05_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_05_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_05_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_05_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_05_MOUSE_Y
                    Case 6
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_06_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_06_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_06_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_06_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_06_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_06_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_06_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_06_MOUSE_Y
                    Case 7
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_07_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_07_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_07_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_07_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_07_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_07_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_07_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_07_MOUSE_Y
                    Case 8
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_08_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_08_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_08_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_08_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_08_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_08_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_08_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_08_MOUSE_Y
                    Case 9
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_09_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_09_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_09_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_09_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_09_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_09_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_09_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_09_MOUSE_Y
                    Case 10
                        .BUTTON(1) = CMB_GAMEPAD_ALLOCATION_10_BUTTON_01
                        .BUTTON(2) = CMB_GAMEPAD_ALLOCATION_10_BUTTON_02
                        .BUTTON(3) = CMB_GAMEPAD_ALLOCATION_10_BUTTON_03
                        .KIND = CMB_GAMEPAD_ALLOCATION_10_KIND
                        .KEY(1) = CMB_GAMEPAD_ALLOCATION_10_KEY_01
                        .KEY(2) = CMB_GAMEPAD_ALLOCATION_10_KEY_02
                        .MOUSE_X = TXT_GAMEPAD_ALLOCATION_10_MOUSE_X
                        .MOUSE_Y = TXT_GAMEPAD_ALLOCATION_10_MOUSE_Y
                End Select
            End With
        Next

    End Sub

    Private Sub SUB_SET_TEST_GAMEPAD_ID()
        Dim OBJ_TEMP As Object
        OBJ_TEMP = CMB_TEST_GAMEPAD_ID.SelectedItem
        If IsNumeric(OBJ_TEMP) Then
            Me.TEST_GAMEPAD_ID = CInt(OBJ_TEMP)
        Else
            Me.TEST_GAMEPAD_ID = 0
        End If
    End Sub

    Private Sub SUB_GET_ITEM_TEST_GANMEPAD_ID()

        Call CMB_TEST_GAMEPAD_ID.Items.Clear()

        Dim INT_NUM_DEVS As Integer
        INT_NUM_DEVS = joyGetNumDevs

        For i = 0 To INT_NUM_DEVS
            Dim SRT_INFO As JOYINFOEX
            SRT_INFO.dwSize = System.Runtime.InteropServices.Marshal.SizeOf(SRT_INFO)
            SRT_INFO.dwFlags = JOY_RETURNALL
            Dim INT_RETURN As Integer
            INT_RETURN = joyGetPosEx(i, SRT_INFO)
            If Not INT_RETURN = JOYERR_NOERROR Then
                Continue For
            End If

            Call CMB_TEST_GAMEPAD_ID.Items.Add(CStr(i))
        Next

        Call SUB_SET_COMBO_KIND_CODE_FIRST(CMB_TEST_GAMEPAD_ID)
    End Sub

    Public Sub SUB_SET_COMBO_KIND_CODE_FIRST(ByRef CMB_CONTROL As System.Windows.Controls.ComboBox)

        With CMB_CONTROL

            If .Items.Count <= 0 Then
                Exit Sub
            End If

            If .SelectedIndex = 0 Then
                Exit Sub
            End If
            .SelectedIndex = -1
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub SUB_REFRESH_TEST_GAPEPAD(ByVal STR_POV As String, ByVal STR_BUTTON As String)
        LBL_TEST_GAMEPAD_POV.Content = STR_POV
        LBL_TEST_GAMEPAD_BUTTON.Content = STR_BUTTON
    End Sub

    Private Function FUNC_GET_BIT_ROW(ByVal INT_VALUE As Integer) As Boolean()
        Dim BLN_RET() As Boolean

        ReDim BLN_RET(16)
        For i = 1 To (BLN_RET.Length - 1)
            BLN_RET(i) = FUNC_GET_BIT(INT_VALUE, i)
        Next

        Return BLN_RET
    End Function

    Private Function FUNC_GET_BIT(ByVal INT_VALUE As Integer, ByVal INT_BIT As Integer) As Boolean

        For i = 1 To 16
            If i = INT_BIT Then
                Dim INT_BEKI As Integer
                INT_BEKI = 2 ^ (i - 1)

                Dim BLN_RET As Boolean
                BLN_RET = INT_BEKI And INT_VALUE

                Return BLN_RET
            End If
        Next

        Return False
    End Function
#End Region

#Region "NEW"
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Call SUB_CTRL_NEW_INIT()
    End Sub
#End Region

#Region "イベント-セレクションチェンジ"
    Private Sub CMB_KIND_GAMEPAD_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_KIND_GAMEPAD.SelectionChanged
        Call MOD_SETTING.SUB_REFRESH_GAMEPAD_BUTTON_WPF(Me)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_01_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_01_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(1)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_02_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_02_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(2)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_03_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_03_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(3)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_04_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_04_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(4)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_05_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_05_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(5)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_06_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_06_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(6)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_07_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_07_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(7)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_08_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_08_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(8)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_09_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_09_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(9)
    End Sub

    Private Sub CMB_GAMEPAD_ALLOCATION_10_KIND_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles CMB_GAMEPAD_ALLOCATION_10_KIND.SelectionChanged
        Call SUB_CHANGE_KIND(10)
    End Sub
#End Region

#Region "イベント-ボタンクリック"
    Private Sub BTN_OK_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_OK.Click
        Call SUB_OK()
    End Sub

    Private Sub BTN_CANCEL_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_CANCEL.Click
        Call SUB_CANCEL()
    End Sub

    Private Sub BTN_INIT_SETTINGS_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_INIT_SETTINGS.Click
        Call SUB_INIT_SETTING()
    End Sub

    Private Sub BTN_TEST_SEND_KEY_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_TEST_SEND_KEY.Click
        Call SUB_TEST_SEND_KEY()
    End Sub

    Private Sub BTN_TEST_SEND_STR_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_TEST_SEND_STR.Click
        Call SUB_TEST_SEND_TEXT()
    End Sub

    Private Sub BTN_TEST_OPERATION_MOUSE_Click(sender As Object, e As Windows.RoutedEventArgs) Handles BTN_TEST_OPERATION_MOUSE.Click
        Call SUB_TEST_OPERATION_MOUSE()
    End Sub
#End Region

#Region "イベント-タイマー"

#Region "変数"
    Private TIM_GAMEPAD As System.Timers.Timer
#End Region

#Region "開始・終了"
    Private Sub SUB_START_TIMER()
        TIM_GAMEPAD = New System.Timers.Timer
        TIM_GAMEPAD.Interval = 500
        TIM_GAMEPAD.AutoReset = True

        AddHandler TIM_GAMEPAD.Elapsed, New System.Timers.ElapsedEventHandler(AddressOf SUB_GAMEPAD_TIMER)
        Call TIM_GAMEPAD.Start()
    End Sub
    Private Sub SUB_STOP_TIMER()
        Call TIM_GAMEPAD.Stop()
    End Sub
#End Region

    Private Sub SUB_GAMEPAD_TIMER(sender As Object, e As System.Timers.ElapsedEventArgs)

        Dim iarINVOKE As IAsyncResult
        Dim myArray(1) As Object
        myArray(0) = ""
        myArray(1) = ""
        iarINVOKE = FRM_PARENT.BeginInvoke(New WPF_SETTING_DELEGATE_REFRESH_TEST_GAPEPAD(AddressOf Me.SUB_REFRESH_TEST_GAPEPAD), myArray)
        Call FRM_PARENT.EndInvoke(iarINVOKE)

        iarINVOKE = FRM_PARENT.BeginInvoke(New WPF_SETTING_DELEGATE(AddressOf Me.SUB_SET_TEST_GAMEPAD_ID))
        Call FRM_PARENT.EndInvoke(iarINVOKE)

        Dim INT_ID As Integer
        INT_ID = Me.TEST_GAMEPAD_ID

        Dim SRT_INFO As JOYINFOEX
        SRT_INFO.dwSize = System.Runtime.InteropServices.Marshal.SizeOf(SRT_INFO)
        SRT_INFO.dwFlags = JOY_RETURNALL

        Dim INT_RETURN As Integer
        INT_RETURN = joyGetPosEx(INT_ID, SRT_INFO)
        If Not INT_RETURN = JOYERR_NOERROR Then
            Exit Sub
        End If

        myArray(0) = "X:" & CStr(SRT_INFO.dwXpos) & " " & "Y:" & CStr(SRT_INFO.dwYpos)
        Dim BLN_BUTTONS() As Boolean
        BLN_BUTTONS = FUNC_GET_BIT_ROW(SRT_INFO.dwButtons)
        Dim STR_TEMP As String
        STR_TEMP = ""
        For i = 1 To (BLN_BUTTONS.Length - 1)
            STR_TEMP &= If(BLN_BUTTONS(i), "○", "×")
        Next
        myArray(1) = STR_TEMP
        iarINVOKE = FRM_PARENT.BeginInvoke(New WPF_SETTING_DELEGATE_REFRESH_TEST_GAPEPAD(AddressOf Me.SUB_REFRESH_TEST_GAPEPAD), myArray)
        Call FRM_PARENT.EndInvoke(iarINVOKE)
    End Sub
#End Region


    Private Sub WPF_SETTING_Loaded(sender As Object, e As Windows.RoutedEventArgs) Handles Me.Loaded
        Call SUB_CTRL_VIEW_INIT()
        Call SUB_CTRL_VALUE_INIT()
    End Sub

    Private Sub WPF_SETTING_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.CHECK_CLOSED = True
        Call SUB_STOP_TIMER()
    End Sub

End Class
