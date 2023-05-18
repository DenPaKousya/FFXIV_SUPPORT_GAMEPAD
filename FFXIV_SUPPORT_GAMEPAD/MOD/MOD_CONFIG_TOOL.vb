Module MOD_CONFIG_TOOL
    Public Function FUNC_GET_APP_SETTINGS(ByVal strKEY As String) As String
        Dim objTEMP As Object
        Dim strRET As String

        Try
            objTEMP = System.Configuration.ConfigurationManager.AppSettings(strKEY)
        Catch ex As Exception
            Return ""
        End Try

        If objTEMP Is Nothing Then
            Return ""
        End If

        strRET = objTEMP.ToString

        Return strRET
    End Function
End Module
