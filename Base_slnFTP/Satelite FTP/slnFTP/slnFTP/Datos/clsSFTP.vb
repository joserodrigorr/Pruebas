Imports Renci.SshNet

'Se requiere la dll Renci, Para que funcionen los Metodos en SFTP
Public Class clsSFTP

    Dim objConfiguracion As New clsConfiguracion

    'CODIGO ORIGINAL PARA CONECTARSE A UNA FTP POR PROTOCOLO SFTP
    Public Sub Codigooriginal()
        Dim RutaClavePrivada As String = "C:\Ruta\id_rsa_nombre"
        Dim FrasePaso As String = "xxxxxxxxxxxxxxxx"
        Dim ClavePrivada As Renci.SshNet.PrivateKeyFile = New Renci.SshNet.PrivateKeyFile(RutaClavePrivada, FrasePaso)
        Dim conexion = New PrivateKeyConnectionInfo("xxx.xxx.xxx.xxx", "xxxxxxxxx", ProxyTypes.Socks5, "socks.xxxx.es", 1080, "", "", ClavePrivada)
        Dim Ftp As Renci.SshNet.SftpClient = New Renci.SshNet.SftpClient(conexion)
        Ftp.Connect()

        ' Descargar un Fichero.
        Dim fsd As System.IO.Stream = New System.IO.FileStream("c:\archivo.txt", System.IO.FileMode.Create)
        Ftp.DownloadFile("/.ssh/authorized_keys2", fsd)
        fsd.Close()

        ' Upload de fichero.
        Dim fs As System.IO.Stream = System.IO.File.OpenRead("c:\archivo.txt")
        Ftp.UploadFile(fs, "/path/archivo.txt", True)
        fs.Close()

        ' Renombrar un fichero.
        Ftp.RenameFile("/path/archivo.txt", "/path/archivo_renombrado.txt")

        ' Eliminar un fichero.
        Ftp.DeleteFile("/path/archivo_renombrado.txt")

        Ftp.Disconnect()
        Ftp.Dispose()

    End Sub

    'Ejemplo del Codigo Original Aplicado
    Public Sub metodosSFTP()

        objConfiguracion.cargarValores()

        Dim Ftp As Renci.SshNet.SftpClient = New Renci.SshNet.SftpClient("179.50.4.57", 22, "euro1", "3ur012015##**!!")

        Try

            Ftp.Connect()

            ''Upload de fichero.
            'Dim fs As System.IO.Stream = System.IO.File.OpenRead("C:\Users\ELKIN\Documents\Prueba.txt")
            'Ftp.UploadFile(fs, "/TRANSACCIONES/Prueba.txt", True)
            'fs.Close()

            'Dim fsd As System.IO.Stream = New System.IO.FileStream("C:\Users\ELKIN\Documents\Archivo_descargado_FTP.txt", System.IO.FileMode.Create)
            'Ftp.DownloadFile("/TRANSACCIONES/Transacciones07042015.txt", fsd)
            'fsd.Close()



            ' Renombrar un fichero.
            'Ftp.RenameFile("ftp://192.168.1.150/TRANSACCIONES/Transacciones07042015.txt", "ftp://192.168.1.150/TRANSACCIONES/Transacciones07042015AAAAAAA.txt")

            'Dim fsd As System.IO.Stream = New System.IO.FileStream("C:\Users\ELKIN\Desktop\Transacciones07042015.txt", System.IO.FileMode.CreateNew)
            'Ftp.DownloadFile("C:\Users\ELKIN\Desktop\Transacciones07042015.txt", fsd)
            'fsd.Close()

            '================================================================================================================================================
            '=================================================================================================================================================

            ' Descargar un Fichero.
            'Dim fsd As System.IO.Stream = New System.IO.FileStream("C:\Users\ELKIN\Documents\Archivo_descargado_FTP.txt", System.IO.FileMode.Create)
            'Ftp.DownloadFile("/TRANSACCIONES/Transacciones07042015.txt", fsd)
            'fsd.Close()

            ' Upload de fichero.
            'Dim fs As System.IO.Stream = System.IO.File.OpenRead("C:\Users\ELKIN\Documents\Prueba.txt")
            'Ftp.UploadFile(fs, "/TRANSACCIONES/PruebasenFTP.txt", True)
            'fs.Close()

            '' Renombrar un fichero.
            'Ftp.RenameFile("/TRANSACCIONES/PruebasenFTP.txt", "/TRANSACCIONES/PruebasenFTP_CambioNombre.txt")

            '' Eliminar un fichero.
            'Ftp.DeleteFile("/TRANSACCIONES/PruebasenFTP_CambioNombre.txt")

            MsgBox("Exito")

        Catch ex As Exception
            MsgBox(ex.ToString)

        Finally
            Ftp.Disconnect()
            Ftp.Dispose()
        End Try


    End Sub


    Public Sub SubirArchivosSFTP()

        objConfiguracion.cargarValores()

        Dim Ftp As Renci.SshNet.SftpClient = New Renci.SshNet.SftpClient("179.50.4.57", 22, "NombreUsuario", "123")

        Try
            Ftp.Connect()
            'Upload de fichero.
            Dim fs As System.IO.Stream = System.IO.File.OpenRead("E:\EjemploCodigos\Archivo.Txt")
            Ftp.UploadFile(fs, "Transacciones/Transaccionesm", True)
            'MsgBox("Se envio el archivo al servidor")
            fs.Close()


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Ftp.Disconnect()
        Ftp.Dispose()
    End Sub

End Class
