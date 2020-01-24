'Using System;
'Using System.Collections.Generic;
'Using System.ComponentModel;
'Using System.Data;
'Using System.Drawing;
'Using System.Linq;
'Using System.Text;
'Using System.Windows.Forms;
Imports System.Data.SqlClient
Public Class clsConsultasSQL

    Public Function LeerConsulta(ByVal Nit As String, ByVal Perfil As Integer) As DataSet

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_Validar_Usuario_Portal"

        sqlComando.Parameters.AddWithValue("@Nit", Nit)
        sqlComando.Parameters.AddWithValue("@Perfil", Perfil)

        sqlAdaptador.SelectCommand = sqlComando

        Try
            sqlAdaptador.Fill(ds)

            Return ds

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Function

    Public Sub ejecutarQuery(ByVal Nit As String, ByVal Perfil As Integer)

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_Validar_Usuario_Portal"

        sqlComando.Parameters.AddWithValue("@Nit", Nit)
        sqlComando.Parameters.AddWithValue("@Perfil", Perfil)

        sqlAdaptador.SelectCommand = sqlComando

        Try
            sqlComando.Connection.Open()
            sqlComando.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

End Class
