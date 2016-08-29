Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class S_CGA_1

#Region "CRUD"

    ''' <summary>
    ''' Carga inicial
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertGlobal(ByVal vp_S_Conexion As String)

        'averiguamos la cantidad de registros
        Dim ResultCount As String = Count_inicial(vp_S_Conexion)

        If ResultCount = "0" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("")
            Console.WriteLine("el achivo de Ecxel no tiene datos")
            Console.ReadLine()
            Return "VACIO"
            Exit Function
        End If

        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("- PASO 2.                                                              -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("- Realizando primera carga en la tabla S_CGA...")
        Console.WriteLine("------------------------------------------------------------------------")

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC INICIAL_CHARGE")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        Return Result

    End Function

    ''' <summary>
    ''' INSERTA LOS REGISTROS NUEVOS DE LA TABLA TEMP_SABANA
    ''' </summary>
    ''' <param name="vp_S_conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert_CGA(ByVal vp_S_conexion As String)

        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("")
        Console.WriteLine("Realizando Insercion en la tabla CGA...")

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC INSERT_S_CGA")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_conexion, StrQuery)

        Return Result

    End Function

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
        sql.Append("EXEC UPDATE_KEYS_TEMP_SABANA")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        If Result = "OK" Then

            sql = New StringBuilder
            StrQuery = ""

            Console.WriteLine("")
            Console.WriteLine("--> Formateando campos fecha...")

            sql.Append("EXEC FORMAT_CAMPOS_DATE")
            StrQuery = sql.ToString

            Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        End If

        Return Result

    End Function

    ''' <summary>
    ''' ACTUALIZACION EN TABLA S_CGA
    ''' </summary>
    ''' <param name="vp_S_Conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update_CGA(ByVal vp_S_Conexion As String)

        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("")
        Console.WriteLine("--> Realizando actualizacion en la tabla CGA...")

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC UPDATE_S_CGA")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_Conexion, StrQuery)

        Return Result

    End Function

    ''' <summary>
    ''' FUNCION QUE LIMPIA LA TABLA TEMPORAL
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DEL_TEMP(ByVal vp_S_conexion As String)

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("DELETE TEMP_SABANA")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_conexion, StrQuery)

        Return Result

    End Function

    ''' <summary>
    ''' ELIMINA LOS REGISTROS YA INSERTADOS DE LA TABLA TEMP_SABANA EN LA TABLA S_CGA
    ''' </summary>
    ''' <param name="vp_S_conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DEL_TEMP_Update(ByVal vp_S_conexion As String)

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC DELETE_TEMP_SABANA_VS_S_CGA")
        StrQuery = sql.ToString

        Result = Conexion.EjecProcedimientos(vp_S_conexion, StrQuery)

        Return Result

    End Function

#End Region

#Region "CONTADORES"

    ''' <summary>
    ''' funcion si valida si la tabla S_CGA esta vacia o no
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CargaInicial(ByVal vp_S_Conexion As String)

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("SELECT COUNT(1) FROM S_CGA")
        StrQuery = sql.ToString

        Result = Conexion.QueryResultado(vp_S_Conexion, StrQuery)
        Return Result

    End Function

    ''' <summary>
    ''' DICE LA CANTIDAD DE REGISTROS QUE SE REALIZA EN LA CARGA INICIAL
    ''' </summary>
    ''' <param name="vp_S_Conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Count_inicial(ByVal vp_S_Conexion As String)

        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("")
        Console.WriteLine("revisando registros...")

        Dim Conexion As New BDClass
        Dim Result As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder
        sql.Append("EXEC COUNT_INICIAL_CHARGE")
        StrQuery = sql.ToString
        Result = Conexion.QueryResultado(vp_S_Conexion, StrQuery)

        Console.WriteLine("la cantidad de registros son: " & Result)

        Return Result
    End Function

    ''' <summary>
    ''' DICE LA CANTIDAD DE REGISTROS QUE SE REALIZA EN EL UPDATE E INSERT
    ''' </summary>
    ''' <param name="vp_S_Conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Count_updateInsert(ByVal vp_S_Conexion As String, ByVal vp_S_Type As String)

        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("")
        Console.WriteLine("--> revisando registros...")

        Dim Conexion As New BDClass
        Dim Result, vl_s_Mensaje As String

        Dim StrQuery As String = ""
        Dim sql As New StringBuilder

        If vp_S_Type = "1" Then
            sql.Append("EXEC COUNT_UPDATE_S_CGA")
            vl_s_Mensaje = "Actualizar"
        Else
            sql.Append("EXEC COUNT_INSERT_S_CGA")
            vl_s_Mensaje = "Insertar"
        End If
        StrQuery = sql.ToString
        Result = Conexion.QueryResultado(vp_S_Conexion, StrQuery)
        Console.ForegroundColor = ConsoleColor.DarkYellow
        Console.WriteLine("    la cantidad de registros a " & vl_s_Mensaje & " son: " & Result)

        Return Result

    End Function

#End Region


End Class
