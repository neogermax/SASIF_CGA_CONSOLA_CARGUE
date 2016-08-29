Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class BDClass

    Dim vg_S_Proveedor As String = "provider=SQLOLEDB;"

    ''' <summary>
    ''' funcion para ejecuta procedimientos almacenados
    ''' </summary>
    ''' <param name="vp_S_conexion"></param>
    ''' <param name="vp_S_StrQuery"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EjecProcedimientos(ByVal vp_S_conexion As String, ByVal vp_S_StrQuery As String)

        'inicializamos conexiones a la BD
        Dim objcmd As OleDbCommand = Nothing
        Dim objConexBD As OleDbConnection = Nothing
        Dim vl_S_processUpdate As String

        Try
            objConexBD = New OleDbConnection(vg_S_Proveedor & vp_S_conexion)
            objConexBD.ConnectionString = vg_S_Proveedor & vp_S_conexion

            'abrimos conexion
            objConexBD.Open()
            'cargamos la carga y la conexion
            objcmd = New OleDbCommand(vp_S_StrQuery, objConexBD)
            'ejecutamos la carga
            objcmd.ExecuteNonQuery()
            'cerramos conexiones
            objConexBD.Close()

            vl_S_processUpdate = "OK"

        Catch ex As Exception
            vl_S_processUpdate = "ERROR"
        End Try
        Return vl_S_processUpdate

    End Function

    ''' <summary>
    ''' funcion generica para consultas de un solo resultado tipo integer
    ''' </summary>
    ''' <param name="vp_S_StrQuery"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryResultado(ByVal vp_S_conexion As String, ByVal vp_S_StrQuery As String)

        'inicializamos conexiones a la BD
        Dim objcmd As OleDbCommand = Nothing
        Dim objConexBD As OleDbConnection = Nothing
        objConexBD = New OleDbConnection(vg_S_Proveedor & vp_S_conexion)
        Dim ReadConsulta As OleDbDataReader = Nothing

        objcmd = objConexBD.CreateCommand
        Dim resultQuery As String = ""

        objConexBD.Open()
        objcmd.CommandText = vp_S_StrQuery

        ReadConsulta = objcmd.ExecuteReader()


        While ReadConsulta.Read
            resultQuery = ReadConsulta.GetValue(0)
        End While

        ReadConsulta.Close()
        objConexBD.Close()

        Return resultQuery

    End Function


End Class
