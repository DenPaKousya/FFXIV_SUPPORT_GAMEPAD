Module MOD_SETTING

#Region "モジュール用・定数"
    Public Const CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT = 10
#End Region

#Region "モジュール用・列挙定数"
    Private Enum ENM_KIND_GAMEPAD
        NONE = 0
        NORMAL = 1
        DS4 = 2
        EDGE301 = 3
        DUAL_SENSE = 4
    End Enum

#End Region

#Region "モジュール用・構造体"
    Public Structure SRT_SETTINGS
        Public GAMEPAD As SRT_SETTINGS_GAMEPAD
    End Structure

    Public Structure SRT_SETTINGS_GAMEPAD
        Public KIND_GAME_PAD As Integer
        Public ALLOCATION() As SRT_SETTINGS_GAMEPAD_ALLOCATION
    End Structure

    Public Structure SRT_SETTINGS_GAMEPAD_ALLOCATION
        Public BUTTON_01 As Integer
        Public BUTTON_02 As Integer
        Public BUTTON_03 As Integer
        Public KIND As Integer
        Public KEY_01 As Integer
        Public KEY_02 As Integer
        Public MOUSE_X As Integer
        Public MOUSE_Y As Integer
    End Structure

#End Region

#Region "モジュール用変数"
    Friend FRM_PARENT As FRM_MAIN
    Private WPF_SETTING As WPF_SETTING
#End Region

    Public Sub SUB_CALL_SETTING(ByRef FRM_ARG As FRM_MAIN)
        If FUNC_GET_STATE() Then
            Call SUB_CLOSE_SETTING()
        Else
            Call SUB_VIEW_SETTING(FRM_ARG)
        End If
    End Sub

    Friend Sub SUB_SET_SETTINGS_WPF(ByRef WPF_SET As WPF_SETTING, ByRef SRT_SET As SRT_SETTINGS)
        With WPF_SET
            .CMB_KIND_GAMEPAD.SelectedIndex = SRT_SET.GAMEPAD.KIND_GAME_PAD

            For i = 1 To (SRT_SET.GAMEPAD.ALLOCATION.Length - 1)
                .SRT_GAC(i).BUTTON(1).SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).BUTTON_01
                .SRT_GAC(i).BUTTON(2).SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).BUTTON_02
                .SRT_GAC(i).BUTTON(3).SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).BUTTON_03
                .SRT_GAC(i).KIND.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).KIND
                .SRT_GAC(i).KEY(1).SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).KEY_01
                .SRT_GAC(i).KEY(2).SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(i).KEY_02
                .SRT_GAC(i).MOUSE_X.Text = SRT_SET.GAMEPAD.ALLOCATION(i).MOUSE_X
                .SRT_GAC(i).MOUSE_Y.Text = SRT_SET.GAMEPAD.ALLOCATION(i).MOUSE_Y
            Next
            'Dim INT_INDEX As Integer
            'INT_INDEX = 0

            'INT_INDEX += 1
            '.CMB_GAMEPAD_ALLOCATION_01_BUTTON_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_01
            '.CMB_GAMEPAD_ALLOCATION_01_BUTTON_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_02
            '.CMB_GAMEPAD_ALLOCATION_01_BUTTON_03.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_03
            '.CMB_GAMEPAD_ALLOCATION_01_KIND.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KIND
            '.CMB_GAMEPAD_ALLOCATION_01_KEY_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_01
            '.CMB_GAMEPAD_ALLOCATION_01_KEY_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_02
            '.TXT_GAMEPAD_ALLOCATION_01_MOUSE_X.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_X
            '.TXT_GAMEPAD_ALLOCATION_01_MOUSE_Y.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_Y

            'INT_INDEX += 1
            '.CMB_GAMEPAD_ALLOCATION_02_BUTTON_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_01
            '.CMB_GAMEPAD_ALLOCATION_02_BUTTON_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_02
            '.CMB_GAMEPAD_ALLOCATION_02_BUTTON_03.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_03
            '.CMB_GAMEPAD_ALLOCATION_02_KIND.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KIND
            '.CMB_GAMEPAD_ALLOCATION_02_KEY_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_01
            '.CMB_GAMEPAD_ALLOCATION_02_KEY_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_02
            '.TXT_GAMEPAD_ALLOCATION_02_MOUSE_X.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_X
            '.TXT_GAMEPAD_ALLOCATION_02_MOUSE_Y.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_Y

            'INT_INDEX += 1
            '.CMB_GAMEPAD_ALLOCATION_03_BUTTON_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_01
            '.CMB_GAMEPAD_ALLOCATION_03_BUTTON_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_02
            '.CMB_GAMEPAD_ALLOCATION_03_BUTTON_03.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_03
            '.CMB_GAMEPAD_ALLOCATION_03_KIND.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KIND
            '.CMB_GAMEPAD_ALLOCATION_03_KEY_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_01
            '.CMB_GAMEPAD_ALLOCATION_03_KEY_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_02
            '.TXT_GAMEPAD_ALLOCATION_03_MOUSE_X.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_X
            '.TXT_GAMEPAD_ALLOCATION_03_MOUSE_Y.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_Y

            'INT_INDEX += 1
            '.CMB_GAMEPAD_ALLOCATION_04_BUTTON_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_01
            '.CMB_GAMEPAD_ALLOCATION_04_BUTTON_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_02
            '.CMB_GAMEPAD_ALLOCATION_04_BUTTON_03.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).BUTTON_03
            '.CMB_GAMEPAD_ALLOCATION_04_KIND.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KIND
            '.CMB_GAMEPAD_ALLOCATION_04_KEY_01.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_01
            '.CMB_GAMEPAD_ALLOCATION_04_KEY_02.SelectedIndex = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).KEY_02
            '.TXT_GAMEPAD_ALLOCATION_04_MOUSE_X.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_X
            '.TXT_GAMEPAD_ALLOCATION_04_MOUSE_Y.Text = SRT_SET.GAMEPAD.ALLOCATION(INT_INDEX).MOUSE_Y

        End With
    End Sub

    Public Function FUNC_GET_SETTINGS_WPF(ByRef CTL_SETTINGS As WPF_SETTING) As SRT_SETTINGS
        Dim SRT_RET As SRT_SETTINGS

        With SRT_RET.GAMEPAD
            .KIND_GAME_PAD = CTL_SETTINGS.CMB_KIND_GAMEPAD.SelectedIndex

            ReDim .ALLOCATION(CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT)
            For i = 1 To (.ALLOCATION.Length - 1)
                .ALLOCATION(i).BUTTON_01 = CTL_SETTINGS.SRT_GAC(i).BUTTON(1).SelectedIndex
                .ALLOCATION(i).BUTTON_02 = CTL_SETTINGS.SRT_GAC(i).BUTTON(2).SelectedIndex
                .ALLOCATION(i).BUTTON_03 = CTL_SETTINGS.SRT_GAC(i).BUTTON(3).SelectedIndex
                .ALLOCATION(i).KIND = CTL_SETTINGS.SRT_GAC(i).KIND.SelectedIndex
                .ALLOCATION(i).KEY_01 = CTL_SETTINGS.SRT_GAC(i).KEY(1).SelectedIndex
                .ALLOCATION(i).KEY_02 = CTL_SETTINGS.SRT_GAC(i).KEY(2).SelectedIndex
                .ALLOCATION(i).MOUSE_X = CTL_SETTINGS.SRT_GAC(i).MOUSE_X.Text
                .ALLOCATION(i).MOUSE_Y = CTL_SETTINGS.SRT_GAC(i).MOUSE_Y.Text
            Next
            'Dim INT_INDEX As Integer
            'INT_INDEX = 0

            'INT_INDEX += 1


            'INT_INDEX += 1
            '.ALLOCATION(INT_INDEX).BUTTON_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_BUTTON_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_BUTTON_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_03 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_BUTTON_03.SelectedIndex
            '.ALLOCATION(INT_INDEX).KIND = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_KIND.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_KEY_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_02_KEY_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).MOUSE_X = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_02_MOUSE_X.Text
            '.ALLOCATION(INT_INDEX).MOUSE_Y = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_02_MOUSE_Y.Text

            'INT_INDEX += 1
            '.ALLOCATION(INT_INDEX).BUTTON_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_BUTTON_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_BUTTON_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_03 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_BUTTON_03.SelectedIndex
            '.ALLOCATION(INT_INDEX).KIND = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_KIND.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_KEY_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_03_KEY_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).MOUSE_X = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_03_MOUSE_X.Text
            '.ALLOCATION(INT_INDEX).MOUSE_Y = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_03_MOUSE_Y.Text

            'INT_INDEX += 1
            '.ALLOCATION(INT_INDEX).BUTTON_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_BUTTON_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_BUTTON_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).BUTTON_03 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_BUTTON_03.SelectedIndex
            '.ALLOCATION(INT_INDEX).KIND = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_KIND.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_01 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_KEY_01.SelectedIndex
            '.ALLOCATION(INT_INDEX).KEY_02 = CTL_SETTINGS.CMB_GAMEPAD_ALLOCATION_04_KEY_02.SelectedIndex
            '.ALLOCATION(INT_INDEX).MOUSE_X = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_04_MOUSE_X.Text
            '.ALLOCATION(INT_INDEX).MOUSE_Y = CTL_SETTINGS.TXT_GAMEPAD_ALLOCATION_04_MOUSE_Y.Text
        End With

        Return SRT_RET
    End Function

    Public Function FUNC_GET_INITAL_SETTINGS() As SRT_SETTINGS
        Dim SRT_RET As SRT_SETTINGS

        With SRT_RET.GAMEPAD
            .KIND_GAME_PAD = 1
        End With

        ReDim SRT_RET.GAMEPAD.ALLOCATION(CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT)
        For i = 1 To (SRT_RET.GAMEPAD.ALLOCATION.Length - 1)
            With SRT_RET.GAMEPAD.ALLOCATION(i)
                .BUTTON_01 = 0
                .BUTTON_02 = 0
                .BUTTON_03 = 0
                .KIND = 0
                .KEY_01 = 0
                .KEY_02 = 0
                .MOUSE_X = 0
                .MOUSE_Y = 0
            End With
        Next

        Return SRT_RET
    End Function

    Public Sub SUB_REFRESH_GAMEPAD_BUTTON_WPF(ByRef WPF_ARG As WPF_SETTING)

        With WPF_ARG
            Dim ENM_KIND As ENM_KIND_GAMEPAD
            ENM_KIND = .CMB_KIND_GAMEPAD.SelectedIndex

            Dim STR_ITEMS() As String
            Select Case ENM_KIND
                Case ENM_KIND_GAMEPAD.NONE
                    ReDim STR_ITEMS(16)
                    For i = 0 To STR_ITEMS.Length - 1
                        STR_ITEMS(i) = "　"
                    Next
                Case ENM_KIND_GAMEPAD.NORMAL
                    ReDim STR_ITEMS(16)
                    STR_ITEMS(0) = "未使用"
                    For i = 1 To STR_ITEMS.Length - 1
                        STR_ITEMS(i) = "ボタン" & i
                    Next
                Case ENM_KIND_GAMEPAD.DS4
                    ReDim STR_ITEMS(14)
                    STR_ITEMS(0) = "未使用"
                    STR_ITEMS(1) = "□"
                    STR_ITEMS(2) = "×"
                    STR_ITEMS(3) = "〇"
                    STR_ITEMS(4) = "△"
                    STR_ITEMS(5) = "L1"
                    STR_ITEMS(6) = "R1"
                    STR_ITEMS(7) = "L2"
                    STR_ITEMS(8) = "R2"
                    STR_ITEMS(9) = "SHARE"
                    STR_ITEMS(10) = "OPTIONS"
                    STR_ITEMS(11) = "L3"
                    STR_ITEMS(12) = "R3"
                    STR_ITEMS(13) = "PS"
                    STR_ITEMS(14) = "PAD"
                Case ENM_KIND_GAMEPAD.EDGE301
                    ReDim STR_ITEMS(10)
                    STR_ITEMS(0) = "未使用"
                    STR_ITEMS(1) = "A"
                    STR_ITEMS(2) = "B"
                    STR_ITEMS(3) = "X"
                    STR_ITEMS(4) = "Y"
                    STR_ITEMS(5) = "LB"
                    STR_ITEMS(6) = "RB"
                    STR_ITEMS(7) = "BACK"
                    STR_ITEMS(8) = "START"
                    STR_ITEMS(9) = "LS"
                    STR_ITEMS(10) = "RS"
                Case ENM_KIND_GAMEPAD.DUAL_SENSE
                    ReDim STR_ITEMS(14)
                    STR_ITEMS(0) = "未使用"
                    STR_ITEMS(1) = "□"
                    STR_ITEMS(2) = "×"
                    STR_ITEMS(3) = "〇"
                    STR_ITEMS(4) = "△"
                    STR_ITEMS(5) = "L1"
                    STR_ITEMS(6) = "R1"
                    STR_ITEMS(7) = "L2"
                    STR_ITEMS(8) = "R2"
                    STR_ITEMS(9) = "SHARE"
                    STR_ITEMS(10) = "OPTIONS"
                    STR_ITEMS(11) = "L3"
                    STR_ITEMS(12) = "R3"
                    STR_ITEMS(13) = "PS"
                    STR_ITEMS(14) = "PAD"
                Case Else
                    ReDim STR_ITEMS(16)
                    For i = 0 To STR_ITEMS.Length - 1
                        STR_ITEMS(i) = "　"
                    Next
            End Select

            For i = 1 To (.SRT_GAC.Length - 1)
                For j = 1 To (.SRT_GAC(i).BUTTON.Length - 1)
                    Call SUB_REFRESH_COMBO_ITEM_WPF(.SRT_GAC(i).BUTTON(j), STR_ITEMS)
                Next
            Next

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_BUTTON_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_BUTTON_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_BUTTON_03, STR_ITEMS)

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_BUTTON_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_BUTTON_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_BUTTON_03, STR_ITEMS)

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_BUTTON_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_BUTTON_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_BUTTON_03, STR_ITEMS)

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_BUTTON_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_BUTTON_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_BUTTON_03, STR_ITEMS)
        End With
    End Sub

    Public Sub SUB_REFRESH_KIND(ByRef WPF_ARG As WPF_SETTING)
        Dim STR_ITEMS() As String

        ReDim STR_ITEMS(1)
        STR_ITEMS(0) = "キー"
        STR_ITEMS(1) = "マウス"

        With WPF_ARG
            For i = 1 To (.SRT_GAC.Length - 1)
                Call SUB_REFRESH_COMBO_ITEM_WPF(.SRT_GAC(i).KIND, STR_ITEMS)
            Next
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_KIND, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_KIND, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_KIND, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_KIND, STR_ITEMS)
        End With
    End Sub

    Public Sub SUB_REFRESH_KEY_MASK(ByRef WPF_ARG As WPF_SETTING)
        Dim STR_ITEMS() As String

        ReDim STR_ITEMS(3)
        STR_ITEMS(0) = "未指定"
        STR_ITEMS(1) = "Alt"
        STR_ITEMS(2) = "Ctrl"
        STR_ITEMS(3) = "Shift"

        With WPF_ARG
            For i = 1 To (.SRT_GAC.Length - 1)
                Call SUB_REFRESH_COMBO_ITEM_WPF(.SRT_GAC(i).KEY(1), STR_ITEMS)
            Next

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_KEY_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_KEY_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_KEY_01, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_KEY_01, STR_ITEMS)
        End With
    End Sub

    Public Sub SUB_REFRESH_KEY(ByRef WPF_ARG As WPF_SETTING)
        Dim STR_ITEMS() As String

        ReDim STR_ITEMS(65)
        STR_ITEMS(0) = "未指定"
        STR_ITEMS(1) = "1"
        STR_ITEMS(2) = "2"
        STR_ITEMS(3) = "3"
        STR_ITEMS(4) = "4"
        STR_ITEMS(5) = "5"
        STR_ITEMS(6) = "6"
        STR_ITEMS(7) = "7"
        STR_ITEMS(8) = "8"
        STR_ITEMS(9) = "9"
        STR_ITEMS(10) = "0"
        STR_ITEMS(11) = "-"
        STR_ITEMS(12) = "^"
        STR_ITEMS(13) = "F1"
        STR_ITEMS(14) = "F2"
        STR_ITEMS(15) = "F3"
        STR_ITEMS(16) = "F4"
        STR_ITEMS(17) = "F5"
        STR_ITEMS(18) = "F6"
        STR_ITEMS(19) = "F7"
        STR_ITEMS(20) = "F8"
        STR_ITEMS(21) = "F9"
        STR_ITEMS(22) = "F10"
        STR_ITEMS(23) = "F11"
        STR_ITEMS(24) = "F12"
        STR_ITEMS(25) = "A"
        STR_ITEMS(26) = "B"
        STR_ITEMS(27) = "C"
        STR_ITEMS(28) = "D"
        STR_ITEMS(29) = "E"
        STR_ITEMS(30) = "F"
        STR_ITEMS(31) = "G"
        STR_ITEMS(32) = "H"
        STR_ITEMS(33) = "I"
        STR_ITEMS(34) = "J"
        STR_ITEMS(35) = "K"
        STR_ITEMS(36) = "L"
        STR_ITEMS(37) = "M"
        STR_ITEMS(38) = "N"
        STR_ITEMS(39) = "O"
        STR_ITEMS(40) = "P"
        STR_ITEMS(41) = "Q"
        STR_ITEMS(42) = "R"
        STR_ITEMS(43) = "S"
        STR_ITEMS(44) = "T"
        STR_ITEMS(45) = "U"
        STR_ITEMS(46) = "V"
        STR_ITEMS(47) = "W"
        STR_ITEMS(48) = "X"
        STR_ITEMS(49) = "Y"
        STR_ITEMS(50) = "Z"
        STR_ITEMS(51) = "テンキー【0】"
        STR_ITEMS(52) = "テンキー【1】"
        STR_ITEMS(53) = "テンキー【2】"
        STR_ITEMS(54) = "テンキー【3】"
        STR_ITEMS(55) = "テンキー【4】"
        STR_ITEMS(56) = "テンキー【5】"
        STR_ITEMS(57) = "テンキー【6】"
        STR_ITEMS(58) = "テンキー【7】"
        STR_ITEMS(59) = "テンキー【8】"
        STR_ITEMS(60) = "テンキー【9】"
        STR_ITEMS(61) = "テンキー【*】"
        STR_ITEMS(62) = "テンキー【+】"
        STR_ITEMS(63) = "テンキー【-】"
        STR_ITEMS(64) = "テンキー【.】"
        STR_ITEMS(65) = "テンキー【/】"


        With WPF_ARG
            For i = 1 To (.SRT_GAC.Length - 1)
                Call SUB_REFRESH_COMBO_ITEM_WPF(.SRT_GAC(i).KEY(2), STR_ITEMS)
            Next

            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_01_KEY_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_02_KEY_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_03_KEY_02, STR_ITEMS)
            'Call SUB_REFRESH_COMBO_ITEM_WPF(.CMB_GAMEPAD_ALLOCATION_04_KEY_02, STR_ITEMS)
        End With
    End Sub

#Region "内部処理"
    Public Function FUNC_GET_STATE() As Boolean
        'If Not (FRM_SETTING Is Nothing) Then
        '    If FRM_SETTING.Visible Then
        '        Return True
        '    End If
        'End If

        If Not (WPF_SETTING Is Nothing) Then
            If Not WPF_SETTING.CHECK_CLOSED Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Sub SUB_CLOSE_SETTING()
        If Not FUNC_GET_STATE() Then
            Exit Sub
        End If

        'Call FRM_SETTING.Close()

        Try
            Call WPF_SETTING.Close()
            WPF_SETTING = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SUB_VIEW_SETTING(ByRef FRM_ARG As FRM_MAIN)

        If FUNC_GET_STATE() Then
            Exit Sub
        End If

        FRM_PARENT = FRM_ARG

        'FRM_SETTING = New FRM_SETTING
        'Call SUB_SET_SETTINGS(FRM_SETTING, SRT_CURRENT_SETTINGS)
        'Call FRM_SETTING.Show()

        WPF_SETTING = New WPF_SETTING
        Call SUB_SET_SETTINGS_WPF(WPF_SETTING, SRT_CURRENT_SETTINGS)
        Call WPF_SETTING.ShowDialog()
    End Sub

    Private Sub SUB_REFRESH_COMBO_ITEM(ByRef CMB_REFRESH As System.Windows.Forms.ComboBox, ByRef STR_ROW() As String)
        Dim intSELECTED_INDEX As Integer

        intSELECTED_INDEX = CMB_REFRESH.SelectedIndex '退避
        Call CMB_REFRESH.Items.Clear()

        For i = 0 To (STR_ROW.Length - 1)
            Call CMB_REFRESH.Items.Add(STR_ROW(i))
        Next

        If CMB_REFRESH.Items.Count > intSELECTED_INDEX Then
            CMB_REFRESH.SelectedIndex = intSELECTED_INDEX
        End If

    End Sub


    Private Sub SUB_REFRESH_COMBO_ITEM_WPF(ByRef CMB_REFRESH As System.Windows.Controls.ComboBox, ByRef STR_ROW() As String)
        Dim intSELECTED_INDEX As Integer

        intSELECTED_INDEX = CMB_REFRESH.SelectedIndex '退避
        Call CMB_REFRESH.Items.Clear()

        For i = 0 To (STR_ROW.Length - 1)
            Call CMB_REFRESH.Items.Add(STR_ROW(i))
        Next

        If CMB_REFRESH.Items.Count > intSELECTED_INDEX Then
            CMB_REFRESH.SelectedIndex = intSELECTED_INDEX
        End If

    End Sub

#End Region

End Module
