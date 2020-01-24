'Nombre	Descripción
'AppendFile	Representa el método de protocolo FTP APPE, utilizado para anexar un archivo a un archivo existente en un servidor FTP.
'DeleteFile	Representa el método de protocolo FTP DELE, utilizado para eliminar un archivo en un servidor FTP.
'DownloadFile	Representa el método de protocolo FTP RETR, utilizado para descargar un archivo de un servidor FTP.
'GetDateTimestamp	Representa el método de protocolo FTP MDTM, (supongo que para recuperar la fecha y hora del fichero)
'GetFileSize	Representa el método de protocolo FTP SIZE, utilizado para recuperar el tamaño de un archivo en un servidor FTP.
'ListDirectory	Representa el método de protocolo FTP NLIST, que obtiene una lista corta de los archivos de un servidor FTP.
'ListDirectoryDetails	Representa el método de protocolo FTP LIST, que obtiene una lista detallada de los archivos de un servidor FTP.
'MakeDirectory	Representa el método de protocolo FTP MKD, que crea un directorio en un servidor FTP.
'PrintWorkingDirectory	Representa el método de protocolo FTP PWD, que imprime el nombre del directorio de trabajo actual.
'RemoveDirectory	Representa el método de protocolo FTP RMD, que quita un directorio.
'Rename	Representa el método de protocolo FTP RENAME, que cambia el nombre de un directorio.
'UploadFile	Representa el método de protocolo FTP STOR, que carga un archivo a un servidor FTP.
'UploadFileWithUniqueName	Representa el protocolo FTP STOU, que carga un archivo con un nombre único a un servidor FTP.

'Para el StreamReader
Imports System.IO

'Para las credenciales d la FTP
Imports System.Net
Public Class clsFTP

    Dim objConfiguracion As New clsConfiguracion

    Public Sub SubirArchivoFTP()

        objConfiguracion.cargarValores()

        My.Computer.Network.UploadFile(objConfiguracion.RutaArchivoLocal, objConfiguracion.RutaFTP, objConfiguracion.UsuarioFTP, objConfiguracion.ClaveFTP, True, 300)
        MsgBox("Se envio el archivo al servidor")


        'CODIGO EJEMPLO
        '============================================================================================================================
        ''Metodo Tradicional Con la ruta directamente
        'Dim RutaLocalArchivo As String = "E:/CargaFTP_Prueba.txt"
        'Dim RutaFTP As String = "ftp://192.168.10.1:21/SubCarpeta/" & "CargaFTP_Prueba.txt"
        'Dim UsuarioFTP As String = "elkinarturo1@hotmail.com"
        'Dim PasswordFTP As String = "gloria123456"

        'My.Computer.Network.UploadFile(RutaLocalArchivo, RutaFTP, UsuarioFTP, PasswordFTP, True, 300)
        'MsgBox("Se envio el archivo al servidor")
        '============================================================================================================================

    End Sub

    Public Sub DescargarArchivoFTP()

        objConfiguracion.cargarValores()

        'Metodo Tradicional Con la ruta directamente
        My.Computer.Network.DownloadFile(New Uri(objConfiguracion.RutaFTP), objConfiguracion.RutaArchivoLocal, objConfiguracion.UsuarioFTP, objConfiguracion.ClaveFTP, True, 300, True)
        MsgBox("Se descargo el archivo del servidor")


        'CODIGO EJEMPLO
        '============================================================================================================================
        ''Metodo Tradicional Con la ruta directamente
        'Dim RutaLocalArchivo As String = "E:\PruebaFTP\CargaFTP_Prueba.txt"
        'Dim RutaFTP As String = "ftp://192.168.10.1:21/SubCarpeta/" & "CargaFTP_Prueba.txt"
        'Dim UsuarioFTP As String = "elkinarturo1@hotmail.com"
        'Dim PasswordFTP As String = "gloria123456"

        'My.Computer.Network.DownloadFile(New Uri(RutaFTP), RutaLocalArchivo, UsuarioFTP, PasswordFTP, True, 300, True)
        'MsgBox("Se descargo el archivo del servidor")
        '============================================================================================================================

    End Sub

    Public Sub MetodosFTP()
        'Metodo usando un objeto  FtpWebRequest
        Dim objFTP As FtpWebRequest = WebRequest.Create("ftp://179.50.4.57/Transacciones/")
        Dim cred As New Net.NetworkCredential("euro1", "3ur012015##**!!")
        objFTP.Credentials = cred

        objFTP.Method = WebRequestMethods.Ftp.AppendFile
        objFTP.Method = WebRequestMethods.Ftp.UploadFile
        objFTP.Method = WebRequestMethods.Ftp.UploadFileWithUniqueName
        objFTP.Method = WebRequestMethods.Ftp.DownloadFile
        objFTP.Method = WebRequestMethods.Ftp.DeleteFile
        objFTP.Method = WebRequestMethods.Ftp.ListDirectory
        objFTP.Method = WebRequestMethods.Ftp.ListDirectoryDetails
        objFTP.Method = WebRequestMethods.Ftp.RemoveDirectory
        objFTP.Method = WebRequestMethods.Ftp.Rename

        Dim ftpRespuesta As FtpWebResponse = objFTP.GetResponse
        Dim Sr As New StreamReader(ftpRespuesta.GetResponseStream)
        Dim lst = Sr.ReadToEnd
    End Sub

End Class
