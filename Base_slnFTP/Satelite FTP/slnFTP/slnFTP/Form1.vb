Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim objFTP As New clsFTP
            objFTP.SubirArchivoFTP()
            'objFTP.DescargarArchivoFTP()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
