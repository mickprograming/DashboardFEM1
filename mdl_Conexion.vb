Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.SqlClient


'Instalar el NuGet: System.Data.Oledb para quitar el error al declarar OleDbConnection
Module mdl_Conexion
    Public Conexion As OleDb.OleDbConnection

    Public Sub AbrirConexionAces()
        Try
            Conexion = New OleDb.OleDbConnection
            Conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\crcofn01\SHARE\Metodos_Capacitacion\P_InfraestructuraGestion\Publica\Bases de Datos\Estadisticas\BDE1.accdb;"
            Conexion.Open()

        Catch ex As Exception
            MsgBox("Error en la conexion a la base de datos : " & ex.Message)
        End Try
    End Sub

    Public Sub CerrarConexionAcces()
        Conexion.Close()
    End Sub

    Public conn As SqlConnection = New SqlConnection
    Public Sub AbrirConexionSQL()
        Dim conString As String = "Initial Catalog=Backflush;Data Source=USTCAS83;Integrated Security=SSPI;" 'KCDF BackFlush
        conn = New SqlConnection(conString)
        'Dim myCmd As SqlCommand
        'Dim myReader As SqlDataReader
        Dim adp As New SqlDataAdapter()
        Dim ds As New DataSet
        Dim dt As New DataTable

        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Error al conectar al servidor" + ex.ToString)
        End Try

    End Sub

    Public Sub CerrarConexionSQL()
        conn.Close()
    End Sub
End Module
