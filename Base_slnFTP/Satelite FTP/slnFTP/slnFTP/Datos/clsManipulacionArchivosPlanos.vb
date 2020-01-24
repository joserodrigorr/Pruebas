'Se usa para el StringBuilder
Imports System.Text
Public Class clsManipulacionArchivosPlanos

    Dim objConfiguracion As New clsConfiguracion
    Dim objSQL As New clsConsultasSQL

    Public Sub ContarArchivos()

        objConfiguracion.cargarValores()

        'Contar el numero de archivos en la ruta de archivos planos
        Dim numArchivo As Integer = 0
        Dim counter = My.Computer.FileSystem.GetFiles(objConfiguracion.RutaArchivoLocal)
        numArchivo = CStr(counter.Count)

    End Sub

    Public Sub CargarRutaArchivo()

        objConfiguracion.cargarValores()

        Dim Ruta As String = ""
        Dim rutaArchivoOrigen As String

        Dim Examinar As New OpenFileDialog
        Try
            Examinar.InitialDirectory = objConfiguracion.RutaArchivoLocal
            Examinar.Filter = "Text files (*.txt)|*.txt|CSV|*.csv"
            Examinar.ShowDialog()
            rutaArchivoOrigen = Examinar.FileName
            Ruta = rutaArchivoOrigen
            MsgBox(rutaArchivoOrigen)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function LeerArchivoPlano(ByVal RutaPlano As String) As String

        objConfiguracion.cargarValores()

        Dim sr As New System.IO.StreamReader(RutaPlano)
        Dim LineaOriginal As String
        Dim Arreglo As Array

        Try
            Do
                LineaOriginal = sr.ReadLine()

                If Not LineaOriginal Is Nothing Then
                    Arreglo = LineaOriginal.Split("|")
                End If

            Loop Until LineaOriginal Is Nothing

            Return "Texto leido OK"

        Catch ex As Exception
            Return "Error al leer el archivo plano : " & ex.Message
        End Try
    End Function

    'Metodo 1 para generar archivos planos
    Public Function GenerarTxT(ByVal ruta As String, ByVal nombreArchivo As String, ByVal dvDatos As DataView) As String

        objConfiguracion.cargarValores()

        Dim strLinea As New StringBuilder

        Try

            For Each ItemEncabezado As DataRow In dvDatos.ToTable.Rows
                strLinea.AppendLine(ItemEncabezado.Item("HEAD").ToString & Chr(9) &
                ItemEncabezado.Item("Sender").ToString & Chr(9) &
                ItemEncabezado.Item("Receiver").ToString & Chr(9) &
                ItemEncabezado.Item("Timestamps").ToString & Chr(9) &
                ItemEncabezado.Item("DistributorStoreld").ToString)
            Next

            My.Computer.FileSystem.WriteAllText(ruta, strLinea.ToString, True)

            Return "-" & nombreArchivo & " Exportado exitosamente en la ruta: " & ruta & vbNewLine

        Catch ex As Exception
            Return "Error en lectura y creacion de TxT : " & ex.Message
        End Try

    End Function

    'Metodo 2 para generar archivos planos
    Public Sub ExportarArchivoPlano()

        objConfiguracion.cargarValores()

        Dim ds As DataSet = objSQL.LeerConsulta("", "")
        Dim archivo As Object
        Dim obj As Object
        Dim campo As String = ""
        Dim RutaArchivo As String = objConfiguracion.RutaArchivoLocal

        'objeto para generar archivo plano
        obj = CreateObject("Scripting.FileSystemObject")
        archivo = obj.CreateTextFile(RutaArchivo)

        For Each Linea As DataRow In ds.Tables(0).Rows
            campo = Linea("Plano").ToString
            archivo.WriteLine(campo)
        Next
        archivo.close()

    End Sub

End Class
