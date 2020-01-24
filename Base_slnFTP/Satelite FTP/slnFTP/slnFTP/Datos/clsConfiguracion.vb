Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class clsConfiguracion

    Public Property EnviarNotificaciones As Boolean
    Public Property ServidorDeCorreo As String
    Public Property Puerto As String
    Public Property RequiereAutenticacion As Boolean
    Public Property SSL As Boolean
    Public Property CorreoRemitente As String
    Public Property UsuarioMail As String
    Public Property ClaveMail As String
    Public Property CorreosNotificaciones As String

    Public Property ServidorFTP As String
    Public Property PuertoFTP As String
    Public Property UsuarioFTP As String
    Public Property ClaveFTP As String
    Public Property RutaFTP As String
    Public Property RutaArchivoLocal As String



    ''' <summary>
    ''' Se usa en diferentes puntos del software para cargar los parametros del sistema, si se agrega una nueva variable es necesario modificar el codigo
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub cargarValores()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        Try
            sqlComando.Connection = sqlConexion
            sqlComando.CommandType = CommandType.StoredProcedure
            sqlComando.CommandText = "sp_Propiedades_Select"
            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.Fill(ds)


            For Each Parametro As DataRow In ds.Tables(0).Rows
                If Parametro.Item("nombrePropiedad").ToString = "EnviarNotificaciones" Then
                    EnviarNotificaciones = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ServidorDeCorreo" Then
                    ServidorDeCorreo = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "Puerto" Then
                    Puerto = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RequiereAutenticacion" Then
                    RequiereAutenticacion = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "SSL" Then
                    SSL = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CorreoRemitente" Then
                    CorreoRemitente = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "UsuarioMail" Then
                    UsuarioMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ClaveMail" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CorreosNotificaciones" Then
                    CorreosNotificaciones = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ServidorFTP" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "PuertoFTP" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "UsuarioFTP" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ClaveFTP" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RutaFTP" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RutaArchivoLocal" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                End If
            Next

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub


End Class