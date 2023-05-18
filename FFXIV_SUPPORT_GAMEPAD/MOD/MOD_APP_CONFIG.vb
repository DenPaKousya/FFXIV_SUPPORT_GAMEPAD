Module MOD_APP_CONFIG

#Region "モジュール用・定数"
    Private Const cstCONFIG_EXTENT As String = "config"
    Private Const cstCONFIE_KEY_MAIN As String = "configuration"
    Private Const cstCONFIE_KEY_SUB As String = "appSettings"
    Private Const cstCONFIG_KEY_ATTRIBUTE_ADD As String = "add"
    Private Const cstCONFIG_KEY_PROPERTY_KEY As String = "key"
    Private Const cstCONFIG_KEY_PROPERTY_VALUE As String = "value"
#End Region
    Public Function FUNC_WRITE_APP_CONFIG(ByVal strKEY_NAME As String, ByVal strVALUE As String)
        Dim STR_NAME_EXE As String
        STR_NAME_EXE = System.IO.Path.GetFileName(System.Windows.Forms.Application.ExecutablePath)

        Dim STR_NAME_APPL As String
        STR_NAME_APPL = STR_NAME_EXE.Replace(".exe", "")

        Dim STR_NAME_DLL As String
        STR_NAME_DLL = STR_NAME_APPL & ".dll"

        Dim ASM_EXEC As System.Reflection.Assembly
        ASM_EXEC = System.Reflection.Assembly.GetExecutingAssembly()

        Dim STR_PATH_CONFIG_APP As String
        STR_PATH_CONFIG_APP = System.IO.Path.GetDirectoryName(ASM_EXEC.Location) & "\" & STR_NAME_DLL & "." & cstCONFIG_EXTENT

        Dim doc As System.Xml.XmlDocument
        doc = New System.Xml.XmlDocument

        Try
            Call doc.Load(STR_PATH_CONFIG_APP)
        Catch ex As Exception
            doc = Nothing
            Return False
        End Try

        Dim node As System.Xml.XmlNode
        node = doc(cstCONFIE_KEY_MAIN)(cstCONFIE_KEY_SUB)

        Dim blnMOD As Boolean
        blnMOD = False
        For Each n In doc(cstCONFIE_KEY_MAIN)(cstCONFIE_KEY_SUB)
            If n.Name = cstCONFIG_KEY_ATTRIBUTE_ADD Then
                If n.Attributes.GetNamedItem(cstCONFIG_KEY_PROPERTY_KEY).Value = strKEY_NAME Then
                    n.Attributes.GetNamedItem(cstCONFIG_KEY_PROPERTY_VALUE).Value = strVALUE
                    blnMOD = True
                End If
            End If
        Next

        If Not blnMOD Then
            Dim newNode As System.Xml.XmlElement
            newNode = doc.CreateElement(cstCONFIG_KEY_ATTRIBUTE_ADD)
            newNode.SetAttribute(cstCONFIG_KEY_PROPERTY_KEY, strKEY_NAME)
            newNode.SetAttribute(cstCONFIG_KEY_PROPERTY_VALUE, strVALUE)
            node.AppendChild(newNode)
        End If

        doc.Save(STR_PATH_CONFIG_APP)
        doc = Nothing

        Call System.Configuration.ConfigurationManager.RefreshSection(cstCONFIE_KEY_SUB)

        Return True
    End Function

    Public Function FUNC_DELETE_APP_CONFIG(ByVal strKEY_NAME As String)

        Dim asm As System.Reflection.Assembly
        Dim appConfigPath As String
        Dim doc As System.Xml.XmlDocument
        Dim node As System.Xml.XmlNode
        Dim strNAME_APPL As String

        strNAME_APPL = System.IO.Path.GetFileName(System.Windows.Forms.Application.ExecutablePath)
        asm = System.Reflection.Assembly.GetExecutingAssembly()
        appConfigPath = System.IO.Path.GetDirectoryName(asm.Location) & "\" & strNAME_APPL & "." & cstCONFIG_EXTENT

        doc = New System.Xml.XmlDocument

        Try
            Call doc.Load(appConfigPath)
        Catch ex As Exception
            doc = Nothing
            Return False
        End Try

        node = doc(cstCONFIE_KEY_MAIN)(cstCONFIE_KEY_SUB)

        For Each n In doc(cstCONFIE_KEY_MAIN)(cstCONFIE_KEY_SUB)
            If n.Name = cstCONFIG_KEY_ATTRIBUTE_ADD Then
                If n.Attributes.GetNamedItem(cstCONFIG_KEY_PROPERTY_KEY).Value = strKEY_NAME Then
                    doc(cstCONFIE_KEY_MAIN)(cstCONFIE_KEY_SUB).RemoveChild(n)
                End If
            End If
        Next

        doc.Save(appConfigPath)
        doc = Nothing

        Call System.Configuration.ConfigurationManager.RefreshSection(cstCONFIE_KEY_SUB)

        Return True
    End Function

    Public Function FUNC_SAVE_APP_CONFIG(ByVal strKEY_NAME As String, ByVal strVALUE As String)
        Dim config As System.Configuration.Configuration

        config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None)
        Call config.AppSettings.Settings.Add(strKEY_NAME, strVALUE)
        Call config.Save(System.Configuration.ConfigurationSaveMode.Modified)
        Call System.Configuration.ConfigurationManager.RefreshSection(cstCONFIE_KEY_SUB)

        Return True
    End Function
End Module
