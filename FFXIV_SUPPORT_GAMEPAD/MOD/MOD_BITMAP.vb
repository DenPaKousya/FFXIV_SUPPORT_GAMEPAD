Imports System.Runtime.InteropServices
Imports System.Windows

Module MOD_BITMAP

    Public Function FUNC_GET_COLOR_ARRAY(ByRef bmpBASE As System.Drawing.Bitmap) As Byte()
        Dim intSTRIDE As Integer
        Dim srcBitmapData As System.Drawing.Imaging.BitmapData
        Dim intSIZE As Integer

        If bmpBASE Is Nothing Then
            Return Nothing
        End If
        Dim rec As System.Drawing.Rectangle
        rec = New System.Drawing.Rectangle(0, 0, bmpBASE.Width, bmpBASE.Height)
        Dim pf As System.Drawing.Imaging.PixelFormat
        pf = bmpBASE.PixelFormat
        If bmpBASE Is Nothing Then
            Return Nothing
        End If
        Try
            srcBitmapData = bmpBASE.LockBits(rec, System.Drawing.Imaging.ImageLockMode.WriteOnly, pf)
        Catch ex As Exception
            Return Nothing
        End Try

        intSTRIDE = srcBitmapData.Stride
        intSIZE = srcBitmapData.Stride * srcBitmapData.Height
        Dim bytRET(intSIZE - 1) As Byte
        'ReDim bytRET(intSIZE - 1)
        Call System.Runtime.InteropServices.Marshal.Copy(srcBitmapData.Scan0, bytRET, 0, intSIZE)

        Try
            Call bmpBASE.UnlockBits(srcBitmapData)
        Catch ex As Exception
            Return Nothing
        End Try

        Return bytRET
    End Function

    'ImageFormatで指定されたImageCodecInfoを探して返す
    Public Function FUNC_GET_ENCODER_INFO(ByVal f As System.Drawing.Imaging.ImageFormat) As System.Drawing.Imaging.ImageCodecInfo
        Dim encs As System.Drawing.Imaging.ImageCodecInfo() = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
        Dim enc As System.Drawing.Imaging.ImageCodecInfo
        For Each enc In encs
            If enc.FormatID = f.Guid Then
                Return enc
            End If
        Next
        Return Nothing
    End Function

    Public Declare Function DeleteObject Lib "gdi32.dll" Alias "DeleteObject" (ByVal hObject As System.IntPtr) As Boolean

    Public Function FUNC_GET_IMAGESOURCE(ByRef bmpBASE As System.Drawing.Bitmap) As System.Windows.Media.ImageSource
        Dim intHANDLE As IntPtr
        Dim imgRET As System.Windows.Media.ImageSource

        If bmpBASE Is Nothing Then
            Return Nothing
        End If

        intHANDLE = bmpBASE.GetHbitmap
        Try
            imgRET = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(intHANDLE, IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions())
        Catch ex As Exception
            Return Nothing
        Finally
            Call DeleteObject(intHANDLE)
        End Try

        Return imgRET
    End Function

    Public Function FUNC_GET_BITMAP(ByRef bmpBASE As System.Windows.Media.Imaging.BitmapSource) As System.Drawing.Bitmap
        Dim intWIDTH As Integer
        Dim intHEIGHT As Integer
        Dim intSTRIDE As Integer
        Dim intHANDLE As IntPtr
        Dim bmpRET As System.Drawing.Bitmap
        Dim pftBASE As System.Drawing.Imaging.PixelFormat

        intWIDTH = bmpBASE.PixelWidth
        intHEIGHT = bmpBASE.PixelHeight

        intSTRIDE = intWIDTH * ((bmpBASE.Format.BitsPerPixel + 7) / 8)

        intHANDLE = IntPtr.Zero

        Try
            intHANDLE = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(intHEIGHT * intSTRIDE)
            bmpBASE.CopyPixels(New System.Windows.Int32Rect(0, 0, intWIDTH, intHEIGHT), intHANDLE, intHEIGHT * intSTRIDE, intSTRIDE)
            pftBASE = System.Drawing.Imaging.PixelFormat.Format24bppRgb
            bmpRET = New System.Drawing.Bitmap(intWIDTH, intHEIGHT, intSTRIDE, pftBASE, intHANDLE)

        Catch ex As Exception
            Return Nothing
        Finally
            If intHANDLE <> IntPtr.Zero Then
                System.Runtime.InteropServices.Marshal.FreeCoTaskMem(intHANDLE)
            End If
        End Try

        Return bmpRET
    End Function

End Module
