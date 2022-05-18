Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Module Conexion_SQL
    Public ConexionSQL_Muestras As SqlConnection

    Public servidor_SQL_Muestras As String = "NDCA0WPAP138.MX.CKSNNET.COM\ALPHAMTY"
    Public DB_SQL_Muestras As String = "AlteaDB"
    Public usuario_Muestras As String = "SA"
    Public pass_Muestras As String = "Alpha123"

    Public conexion_a_sql_Muestras As New OleDbConnection
    Public comando_a_sql_Muestras As New SqlCommand
    Public adaptador_a_sql_Muestras As New SqlDataAdapter
    Public registro_a_sql_Muestras As New DataSet

    Public estadosql_Muestras As String

    Sub Conexion_SQL_Retention()
        Try
            Dim cadena_Retention = "data source =" & servidor_SQL_Muestras & "; initial catalog = " & DB_SQL_Muestras & "; user id = " & usuario_Muestras & "; password = " & pass_Muestras
            'Dim cadena_OEEdata = "data source =" & servidor_OEEdata & "; initial catalog =" & bd_OEEdata & "; Integrated Security = SSPI"

            ConexionSQL_Muestras = New SqlConnection(cadena_Retention)

            ConexionSQL_Muestras.Open()
            estadosql_Muestras = "Online_OEEdata"
            'MsgBox("Conexión establecida")
            'estado_rete = "Online_SQL"
            'ConexionSQL_Muestras.Close()
            'estadosql = "Offline"
            'MsgBox("Conexión cerrada")

        Catch ex As Exception

            estadosql_Muestras = "Failure offline_SQL"
            'MsgBox("Error al tratar de conectar a la base de datos: " & ex.Message)
            'estado_rete = "Offline"
            MsgBox("Error al tratar de conectar con la base de datos de ITEMS " + ex.Message)
        End Try
    End Sub


End Module
