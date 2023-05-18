Imports System.Runtime.InteropServices
Module MOD_SEND_INPUT
    ' マウスイベント(mouse_eventの引数と同様のデータ)
    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Public Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As UInteger
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    ' キーボードイベント(keybd_eventの引数と同様のデータ)
    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Public Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    ' ハードウェアイベント
    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Public Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure

    ' 各種イベント(SendInputの引数データ)
    '<StructLayout(LayoutKind.Explicit)>
    'Public Structure INPUT
    '    <FieldOffset(0)> Public type As Integer
    '    <FieldOffset(4)> Public mi As MOUSEINPUT
    '    <FieldOffset(4)> Public ki As KEYBDINPUT
    '    <FieldOffset(4)> Public hi As HARDWAREINPUT
    'End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure INPUT
        <FieldOffset(0)> Public type As Integer
        <FieldOffset(8)> Public mi As MOUSEINPUT
        <FieldOffset(8)> Public ki As KEYBDINPUT
        <FieldOffset(8)> Public hi As HARDWAREINPUT
    End Structure

    ' キー操作、マウス操作をシミュレート(擬似的に操作する)
    <DllImport("user32.dll")>
    Private Sub SendInput(
        ByVal nInputs As Integer, ByRef pInputs As INPUT, ByVal cbsize As Integer)
    End Sub

    ' 仮想キーコードをスキャンコードに変換
    <DllImport("user32.dll", EntryPoint:="MapVirtualKeyA")>
    Private Function MapVirtualKey(
        ByVal wCode As Integer, ByVal wMapType As Integer) As Integer
    End Function

    Private Const INPUT_MOUSE = 0                   ' マウスイベント
    Private Const INPUT_KEYBOARD = 1                ' キーボードイベント
    Private Const INPUT_HARDWARE = 2                ' ハードウェアイベント

    Private Const MOUSEEVENTF_MOVE = &H1            ' マウスを移動する
    Private Const MOUSEEVENTF_ABSOLUTE = &H8000     ' 絶対座標指定
    Private Const MOUSEEVENTF_LEFTDOWN = &H2        ' 左　ボタンを押す
    Private Const MOUSEEVENTF_LEFTUP = &H4          ' 左　ボタンを離す
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8       ' 右　ボタンを押す
    Private Const MOUSEEVENTF_RIGHTUP = &H10        ' 右　ボタンを離す
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20     ' 中央ボタンを押す
    Private Const MOUSEEVENTF_MIDDLEUP = &H40       ' 中央ボタンを離す
    Private Const MOUSEEVENTF_WHEEL = &H800         ' ホイールを回転する
    Private Const WHEEL_DELTA = 120                 ' ホイール回転値

    Private Const KEYEVENTF_KEYDOWN = &H0           ' キーを押す
    Private Const KEYEVENTF_KEYUP = &H2             ' キーを離す
    Private Const KEYEVENTF_EXTENDEDKEY = &H1       ' 拡張コード
    Private Const VK_SHIFT = &H10                   ' SHIFTキー
    Private Const VK_LCONTROL = &HA2 '左Ctrl

    Public Sub SUB_MOUSE_MOVE(ByVal INT_X As Integer, ByVal INT_Y As Integer)
        ' マウス操作実行用のデータ
        Const num As Integer = 1
        Dim inp As INPUT() = New INPUT(num - 1) {}

        'マウスカーソルを移動する
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE
        inp(0).mi.dx = INT_X * (65535 / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width)
        inp(0).mi.dy = INT_Y * (65535 / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウス操作実行
        Call SendInput(num, inp(0), Marshal.SizeOf(inp(0)))
    End Sub

    Public Sub SUB_MOUSE_MOVE_RALATIVE(ByVal INT_X As Integer, ByVal INT_Y As Integer)
        Dim INT_X_CUR As Integer
        Dim INT_Y_CUR As Integer

        Dim PNT_XY As System.Drawing.Point
        PNT_XY = System.Windows.Forms.Cursor.Position
        INT_X_CUR = PNT_XY.X
        INT_Y_CUR = PNT_XY.Y

        Call SUB_MOUSE_MOVE(INT_X_CUR + INT_X, INT_Y_CUR + INT_Y)
    End Sub

    Public Sub SUB_MOUSE_LEFT_CLICK()
        ' マウス操作実行用のデータ
        Const num As Integer = 2
        Dim inp As INPUT() = New INPUT(num - 1) {}

        'マウスのボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        'マウスのボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウス操作実行
        Call SendInput(num, inp(0), Marshal.SizeOf(inp(0)))
    End Sub

    Public Sub SUB_MOUSE_LEFT_DOUBLE_CLICK()
        ' マウス操作実行用のデータ
        Const num As Integer = 4
        Dim inp As INPUT() = New INPUT(num - 1) {}

        'マウスのボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        'マウスのボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        'マウスのボタンを押す
        inp(2).type = INPUT_MOUSE
        inp(2).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(2).mi.dx = 0
        inp(2).mi.dy = 0
        inp(2).mi.mouseData = 0
        inp(2).mi.dwExtraInfo = 0
        inp(2).mi.time = 0

        'マウスのボタンを離す
        inp(3).type = INPUT_MOUSE
        inp(3).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(3).mi.dx = 0
        inp(3).mi.dy = 0
        inp(3).mi.mouseData = 0
        inp(3).mi.dwExtraInfo = 0
        inp(3).mi.time = 0

        ' マウス操作実行
        Call SendInput(num, inp(0), Marshal.SizeOf(inp(0)))
    End Sub

    Public Sub SUB_MOUSE_WHEEL_DELTA(ByVal INT_SCROL_ROW As Integer, ByVal BLN_WHEEL_FORWARD As Boolean)
        ' マウス操作実行用のデータ
        Const num As Integer = 1
        Dim inp As INPUT() = New INPUT(num - 1) {}

        'マウスホイールを前方(近づく方向)へ回転する
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_WHEEL
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        Dim INT_FORWARD As Integer
        If BLN_WHEEL_FORWARD Then
            INT_FORWARD = -1
        Else
            INT_FORWARD = 1
        End If
        inp(0).mi.mouseData = INT_FORWARD * INT_SCROL_ROW * WHEEL_DELTA
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウス操作実行
        Call SendInput(num, inp(0), Marshal.SizeOf(inp(0)))

    End Sub

    Public Sub SUB_KEYBOARD_INPUT(ByVal KEY_INPUT As System.Windows.Forms.Keys)

        ' キーボード操作実行用のデータ
        Const num As Integer = 2
        Dim inp As INPUT() = New INPUT(num - 1) {}

        'キーボードを押す
        inp(0).type = INPUT_KEYBOARD
        inp(0).ki.wVk = KEY_INPUT
        inp(0).ki.wScan = MapVirtualKey(inp(0).ki.wVk, 0)
        inp(0).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYDOWN
        inp(0).ki.dwExtraInfo = 0
        inp(0).ki.time = 0

        'キーボードを離す
        inp(1).type = INPUT_KEYBOARD
        inp(1).ki.wVk = KEY_INPUT
        inp(1).ki.wScan = MapVirtualKey(inp(1).ki.wVk, 0)
        inp(1).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP
        inp(1).ki.dwExtraInfo = 0
        inp(1).ki.time = 0

        ' キーボード操作実行
        SendInput(num, inp(0), Marshal.SizeOf(inp(0)))
    End Sub

    Public Sub SUB_KEYBOARD_INPUT_CONTINUE(ByVal STR_PUT As String)

        For i = 0 To STR_PUT.Length - 1
            Dim CHR_ONE As Char
            CHR_ONE = STR_PUT.Substring(i, 1)
            Dim KEY_CODE As System.Windows.Forms.Keys
            KEY_CODE = FUNC_CONVERT_CHAR_TO_KEY(CHR_ONE)

            If KEY_CODE = 0 Then
                Continue For
            End If

            Call SUB_KEYBOARD_INPUT(KEY_CODE)
            Call System.Threading.Thread.Sleep(100)
        Next

    End Sub

    Public Sub SUB_KEYBOARD_INPUT_CTRL_C()

        ' キーボード操作実行用のデータ
        Const num As Integer = 4
        Dim inp As INPUT() = New INPUT(num - 1) {}

        ' キーボード(CTRL)を押す
        inp(0).type = INPUT_KEYBOARD
        inp(0).ki.wVk = VK_LCONTROL
        inp(0).ki.wScan = MapVirtualKey(inp(0).ki.wVk, 0)
        inp(0).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYDOWN
        inp(0).ki.dwExtraInfo = 0
        inp(0).ki.time = 0

        'キーボードを押す
        inp(1).type = INPUT_KEYBOARD
        inp(1).ki.wVk = System.Windows.Forms.Keys.C
        inp(1).ki.wScan = MapVirtualKey(inp(1).ki.wVk, 0)
        inp(1).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYDOWN
        inp(1).ki.dwExtraInfo = 0
        inp(1).ki.time = 0

        'キーボードを離す
        inp(2).type = INPUT_KEYBOARD
        inp(2).ki.wVk = System.Windows.Forms.Keys.C
        inp(2).ki.wScan = MapVirtualKey(inp(2).ki.wVk, 0)
        inp(2).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP
        inp(2).ki.dwExtraInfo = 0
        inp(2).ki.time = 0

        'キーボード(Ctrl)を離す
        inp(3).type = INPUT_KEYBOARD
        inp(3).ki.wVk = VK_LCONTROL
        inp(3).ki.wScan = MapVirtualKey(inp(3).ki.wVk, 0)
        inp(3).ki.dwFlags = KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP
        inp(3).ki.dwExtraInfo = 0
        inp(3).ki.time = 0

        ' キーボード操作実行
        SendInput(num, inp(0), Marshal.SizeOf(inp(0)))
    End Sub

    Private Function FUNC_CONVERT_CHAR_TO_KEY(ByVal CHR_ONE As Char) As System.Windows.Forms.Keys
        Dim KEY_RET As System.Windows.Forms.Keys
        Select Case CHR_ONE
            Case "0"
                KEY_RET = System.Windows.Forms.Keys.D0
            Case "1"
                KEY_RET = System.Windows.Forms.Keys.D1
            Case "2"
                KEY_RET = System.Windows.Forms.Keys.D2
            Case "3"
                KEY_RET = System.Windows.Forms.Keys.D3
            Case "4"
                KEY_RET = System.Windows.Forms.Keys.D4
            Case "5"
                KEY_RET = System.Windows.Forms.Keys.D5
            Case "6"
                KEY_RET = System.Windows.Forms.Keys.D6
            Case "7"
                KEY_RET = System.Windows.Forms.Keys.D7
            Case "8"
                KEY_RET = System.Windows.Forms.Keys.D8
            Case "9"
                KEY_RET = System.Windows.Forms.Keys.D9
            Case "a"
                KEY_RET = System.Windows.Forms.Keys.A
            Case "b"
                KEY_RET = System.Windows.Forms.Keys.B
            Case "c"
                KEY_RET = System.Windows.Forms.Keys.C
            Case "d"
                KEY_RET = System.Windows.Forms.Keys.D
            Case "e"
                KEY_RET = System.Windows.Forms.Keys.E
            Case "f"
                KEY_RET = System.Windows.Forms.Keys.F
            Case "g"
                KEY_RET = System.Windows.Forms.Keys.G
            Case "h"
                KEY_RET = System.Windows.Forms.Keys.H
            Case "i"
                KEY_RET = System.Windows.Forms.Keys.I
            Case "j"
                KEY_RET = System.Windows.Forms.Keys.J
            Case "k"
                KEY_RET = System.Windows.Forms.Keys.K
            Case "l"
                KEY_RET = System.Windows.Forms.Keys.L
            Case "m"
                KEY_RET = System.Windows.Forms.Keys.M
            Case "n"
                KEY_RET = System.Windows.Forms.Keys.N
            Case "o"
                KEY_RET = System.Windows.Forms.Keys.O
            Case "p"
                KEY_RET = System.Windows.Forms.Keys.P
            Case "q"
                KEY_RET = System.Windows.Forms.Keys.Q
            Case "r"
                KEY_RET = System.Windows.Forms.Keys.R
            Case "s"
                KEY_RET = System.Windows.Forms.Keys.S
            Case "t"
                KEY_RET = System.Windows.Forms.Keys.T
            Case "u"
                KEY_RET = System.Windows.Forms.Keys.U
            Case "v"
                KEY_RET = System.Windows.Forms.Keys.V
            Case "w"
                KEY_RET = System.Windows.Forms.Keys.W
            Case "x"
                KEY_RET = System.Windows.Forms.Keys.X
            Case "y"
                KEY_RET = System.Windows.Forms.Keys.Y
            Case "z"
                KEY_RET = System.Windows.Forms.Keys.Z
            Case Else
                KEY_RET = 0
        End Select

        Return KEY_RET
    End Function
End Module
