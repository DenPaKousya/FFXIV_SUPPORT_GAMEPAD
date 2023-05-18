Module MOD_INIT_FINAL
    Public Function FUNC_INIT() As Boolean
        If Not FUNC_INIT_SETTING() Then
            Return False
        End If

        If Not FUNC_GET_SETTING() Then
            Return False
        End If

        Return True
    End Function

    Public Function FUNC_FINALIZE() As Boolean
        Call FUNC_SET_SETTING()
        Return True
    End Function

#Region "内部処理"
    Private Function FUNC_INIT_SETTING()
        Dim SRT_INIT As SRT_SETTINGS

        SRT_INIT = FUNC_GET_INITAL_SETTINGS()

        SRT_CURRENT_SETTINGS = SRT_INIT

        Return True
    End Function

    Private Function FUNC_GET_SETTING()
        Dim STR_TEMP As String

        Try
            STR_TEMP = FUNC_GET_APP_SETTINGS(CST_CONFIG_GAMEPAD_ID_GAMEPAD)
            If Not STR_TEMP = "" Then
                INT_SET_ID_GAMEPAD = CInt(STR_TEMP)
            Else
                INT_SET_ID_GAMEPAD = 0
            End If
        Catch ex As Exception

        End Try

        With SRT_CURRENT_SETTINGS.GAMEPAD
            Try

                STR_TEMP = FUNC_GET_APP_SETTINGS(CST_CONFIG_GAMEPAD_KIND_GAMEPAD)
                If Not STR_TEMP = "" Then
                    .KIND_GAME_PAD = CInt(STR_TEMP)
                End If

                ReDim .ALLOCATION(CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT)
                For i = 1 To (.ALLOCATION.Length - 1)
                    Dim STR_KEY_BASE As String
                    STR_KEY_BASE = ""
                    STR_KEY_BASE &= CST_CONFIG_GAMEPAD_ALLOCATION
                    STR_KEY_BASE &= "_" & MOD_CODE_TOOL.Format(i, "00")

                    Dim STR_KEY As String

                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_01
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).BUTTON_01 = CInt(STR_TEMP)
                    End If
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_02
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).BUTTON_02 = CInt(STR_TEMP)
                    End If
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_03
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).BUTTON_03 = CInt(STR_TEMP)
                    End If

                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KIND
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).KIND = CInt(STR_TEMP)
                    End If

                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KEY_01
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).KEY_01 = CInt(STR_TEMP)
                    End If
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KEY_02
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).KEY_02 = CInt(STR_TEMP)
                    End If

                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_MOUSE_X
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).MOUSE_X = CInt(STR_TEMP)
                    End If
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_MOUSE_Y
                    STR_TEMP = FUNC_GET_APP_SETTINGS(STR_KEY)
                    If Not STR_TEMP = "" Then
                        .ALLOCATION(i).MOUSE_Y = CInt(STR_TEMP)
                    End If
                Next
            Catch ex As Exception

            End Try
        End With

        Return True
    End Function

    Public Function FUNC_SET_SETTING()
        Dim STR_TEMP As String

        With SRT_CURRENT_SETTINGS.GAMEPAD
            Try
                STR_TEMP = CStr(.KIND_GAME_PAD)
                Call FUNC_WRITE_APP_CONFIG(CST_CONFIG_GAMEPAD_KIND_GAMEPAD, STR_TEMP)

                For i = 1 To CST_SETTINGS_GAMEPAD_ALLOCATION_COUNT
                    Dim STR_KEY_BASE As String
                    STR_KEY_BASE = ""
                    STR_KEY_BASE &= CST_CONFIG_GAMEPAD_ALLOCATION
                    STR_KEY_BASE &= "_" & MOD_CODE_TOOL.Format(i, "00")

                    Dim STR_KEY As String

                    STR_TEMP = CStr(.ALLOCATION(i).BUTTON_01)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_01
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).BUTTON_02)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_02
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).BUTTON_03)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_BUTTON_03
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).KIND)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KIND
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).KEY_01)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KEY_01
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).KEY_02)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_KEY_02
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).MOUSE_X)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_MOUSE_X
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)

                    STR_TEMP = CStr(.ALLOCATION(i).MOUSE_Y)
                    STR_KEY = STR_KEY_BASE & "_" & CST_CONFIG_GAMEPAD_ALLOCATION_MOUSE_Y
                    Call FUNC_WRITE_APP_CONFIG(STR_KEY, STR_TEMP)
                Next

            Catch ex As Exception

            End Try
        End With

        Return True
    End Function
#End Region

End Module
