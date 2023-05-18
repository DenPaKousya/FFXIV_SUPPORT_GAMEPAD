Module MOD_MAIN

#Region "モジュール用・変数"
    Private FRM_OWNER As FRM_MAIN
    Private TIM_PRIVATE_TIMER As System.Timers.Timer
#End Region

    Public Sub aMain()

        If Not FUNC_INIT() Then
            Exit Sub
        End If

        FRM_OWNER = New FRM_MAIN

        Call MOD_GAMEPAD.SUB_GET_JOYSTICK(FRM_OWNER, INT_SET_ID_GAMEPAD)
        Call MOD_GAMEPAD.SUB_PARAM_INIT()
        Call SUB_GET_PROCESS()
        Call SUB_START_TIMER()

        Call System.Windows.Forms.Application.Run()
    End Sub

    Public Sub SUB_MAIN_END()

        'Call FRM_OWNER.Close()
        '        Call System.Windows.Forms.Application.Exit()
    End Sub

#Region "タイマー開始・終了"
    Friend Sub SUB_START_TIMER()

        TIM_PRIVATE_TIMER = New System.Timers.Timer
        TIM_PRIVATE_TIMER.Interval = 5000
        TIM_PRIVATE_TIMER.AutoReset = True

        AddHandler TIM_PRIVATE_TIMER.Elapsed, New System.Timers.ElapsedEventHandler(AddressOf SUB_TIMER)
        Call TIM_PRIVATE_TIMER.Start()
    End Sub

    Friend Sub SUB_STOP_TIMER()
        Call TIM_PRIVATE_TIMER.Stop()
    End Sub

#End Region

#Region "タイマー実行"
    Private Sub SUB_TIMER(sender As Object, e As System.Timers.ElapsedEventArgs)
        Call SUB_GET_PROCESS()
    End Sub
#End Region

#Region "プロセス検知・確認"
    Friend Sub SUB_GET_PROCESS()

        PRC_TARGET = FUNC_GET_FFXIV_PROCESS()

        If PRC_TARGET Is Nothing Then 'プロセスが無い場合は
            Exit Sub
        End If
    End Sub

#End Region

End Module
