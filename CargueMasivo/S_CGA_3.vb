Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class S_CGA_3

#Region "CRUD"

    ''' <summary>
    ''' actualizacion de llaves en la tabla temporal
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update_Keys(ByVal vp_S_Conexion As String)

        Console.ForegroundColor = ConsoleColor.DarkYellow
        Console.WriteLine("")
        Console.WriteLine("--> Actualizando llaves...")

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC UPDATE_KEYS_TEMP_SABANA3")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        If Result = "OK" Then

            sql = New StringBuilder
            StrQuery = ""

            Console.WriteLine("")
            Console.WriteLine("--> Formateando campos ...")

            sql.Append("EXEC FORMAT_CAMPOS_DATE_S3")
            StrQuery = sql.ToString

            Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        End If

        Return Result

    End Function


#End Region

End Class
