Module MOD_OPERATION

    'FFXIVクライアント左上端から設定の位置へマウスポインタを移動する
    Public Sub SUB_SET_MOUSE_POINTER_CLIENT(ByVal INT_X As Integer, ByVal INT_Y As Integer)

        If PRC_TARGET Is Nothing Then
            Exit Sub
        End If

        Dim INT_LEFT As Integer
        INT_LEFT = FUNC_GET_LEFT_WINDOW(PRC_TARGET)

        Dim INT_TOP As Integer
        INT_TOP = FUNC_GET_TOP_WINDOW(PRC_TARGET)

        Dim INT_X_SET As Integer
        INT_X_SET = INT_X + INT_LEFT
        Dim INT_Y_SET As Integer
        INT_Y_SET = INT_Y + INT_TOP

        Call SUB_MOUSE_MOVE(INT_X_SET, INT_Y_SET)
    End Sub
End Module
