Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.Data.OleDb
Imports System.Data.SqlClient


Module cargue

    Public S_CGA_1 As New S_CGA_1
    Public S_CGA_3 As New S_CGA_3

    ''' <summary>
    ''' main que ejecuta la carga masiva
    ''' </summary>
    ''' <remarks></remarks>
    Sub Main()

        Dim BD As ModelDataContext = New ModelDataContext()
        Dim vl_S_path As String = My.Application.Info.DirectoryPath
        'llamamos encabesado de la CONSOLA
        ENTRADA()

        'llamamos captura de ruta de coneccion a la BD
        Dim vl_S_Conexion As String = CapturaRutas(vl_S_path, "0")
        If vl_S_Conexion = "ERROR" Then
            Exit Sub
        End If

        'LIMPIAMOS LA TABLA TEMPORAL
        S_CGA_1.DEL_TEMP(vl_S_Conexion)

        'llamamos captura de ruta archivo EXCEL SABANA_1
        Dim vl_S_Ruta As String = CapturaRutas(vl_S_path, "1")
        If vl_S_Ruta = "ERROR" Then
            Exit Sub
        End If

        'iniciamos el proceso de la SABANA_1
        Dim SABANA_1 As String = ENT_SABANA_1(vl_S_Ruta, vl_S_Conexion)
        Select Case SABANA_1
            Case "ERROR"
                Exit Sub
            Case "0"
         End Select

        'llamamos captura de ruta archivo EXCEL SABANA_3
        Dim vl_S_Ruta_S3 As String = CapturaRutas(vl_S_path, "2")
        If vl_S_Ruta_S3 = "ERROR" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Revise si el archivo existe o fue modificado el nombre (Excel_ruta_S3.txt)")
            Console.ReadLine()
        Else
            'iniciamos el proceso de la SABANA_3
            Dim SABANA_3 As String = ENT_SABANA_3(vl_S_Ruta_S3, vl_S_Conexion)
            Select Case SABANA_3
                Case "ERROR"
                    Exit Sub
                Case "0"

                Case "PASO"
                    Exit Select
            End Select
        End If

        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("Proceso terminado...")


    End Sub

#Region "PROCESO SABANAS"

    ''' <summary>
    ''' verifica y trae las rutas de los archivos txt
    ''' </summary>
    ''' <param name="vp_S_Path"></param>
    ''' <param name="vp_S_TipoRuta"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CapturaRutas(ByVal vp_S_Path As String, ByVal vp_S_TipoRuta As String)

        Dim vl_S_Result As String = ""

        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("------------------------------------------------------------------------")

        Select Case vp_S_TipoRuta
            Case "0"
                'llamamos la funcion de busqueda de la ruta
                Dim vl_S_Conexion As String = Readruta(vp_S_Path, vp_S_TipoRuta)
                'validamos si esta el txt.CONECCION
                If vl_S_Conexion = "ERROR" Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Revise si el archivo existe o fue modificado el nombre (STR_Connection.txt)")
                    Console.ReadLine()
                    vl_S_Result = vl_S_Conexion
                    GoTo SALIDA
                Else
                    Console.ForegroundColor = ConsoleColor.Gray
                    Console.WriteLine(" Ruta de conexion: " & vl_S_Conexion)
                    vl_S_Result = vl_S_Conexion
                End If
            Case "1"
                'llamamos la funcion de busqueda de la ruta
                Dim vl_S_Ruta As String = Readruta(vp_S_Path, vp_S_TipoRuta)
                'validamos si esta el txt.ECXEL
                If vl_S_Ruta = "ERROR" Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Revise si el archivo existe o fue modificado el nombre (Excel_ruta.txt)")
                    Console.ReadLine()
                    vl_S_Result = vl_S_Ruta
                    GoTo SALIDA
                Else
                    Console.ForegroundColor = ConsoleColor.Gray
                    Console.WriteLine(" Ruta del archivo EXCEL SABANA_1: " & vl_S_Ruta)
                    vl_S_Result = vl_S_Ruta
                End If

            Case "2"
                'llamamos la funcion de busqueda de la ruta
                Dim vl_S_Ruta As String = Readruta(vp_S_Path, vp_S_TipoRuta)
                'validamos si esta el txt.ECXEL
                If vl_S_Ruta = "ERROR" Then
                    vl_S_Result = vl_S_Ruta
                    GoTo SALIDA
                Else
                    Console.ForegroundColor = ConsoleColor.Gray
                    Console.WriteLine(" Ruta del archivo EXCEL SABANA_3: " & vl_S_Ruta)
                    vl_S_Result = vl_S_Ruta
                End If
        End Select

SALIDA:
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("------------------------------------------------------------------------")

        Return vl_S_Result

    End Function

    ''' <summary>
    ''' FUNCION QUE REALIZA EL PROCESO DEL CARGUE SABANA_1
    ''' </summary>
    ''' <param name="vp_S_Ruta"></param>
    ''' <param name="vp_S_Conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SABANA_1(ByVal vp_S_Ruta As String, ByVal vp_S_Conexion As String)

        Dim vl_S_Result As String = "OK"

        'llamamos la funcion copia del excel a la tabla temporal
        Dim vl_S_copy As String = copyTemporal(vp_S_Ruta, vp_S_Conexion, "SABANA_1")
        'validamos si la copia fue exitosa 
        If vl_S_copy = "ERROR" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("")
            Console.WriteLine("Procedimiento BulkCopy fallo el achivo Excel es diferente a la data maestra o el archivo (SABANA1_data.xls) esta abierto!")
            Console.ReadLine()
            vl_S_Result = vl_S_copy
            GoTo SALIDA
        End If

        'llamamos la funcion que actualiza las llaves en la tabla temporal
        Dim UpdateKey As String = S_CGA_1.Update_Keys(vp_S_Conexion)
        'validamos si la actualizacion de llaves fue exitosa 
        If UpdateKey = "ERROR" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("")
            Console.WriteLine("fallo al actualizar las llaves consulte con admin Sasif S.A.")
            Console.ReadLine()
            vl_S_Result = UpdateKey
            GoTo SALIDA
        End If

        'llamamos la funcion que valida si es la primera carga
        Dim Carga As String = S_CGA_1.CargaInicial(vp_S_Conexion)
        'validamos si la actualizacion de llaves fue exitosa 
        If Carga = "0" Then
            'llamamos la funcion que inserta si es la primera carga
            Dim CInicial As String = S_CGA_1.InsertGlobal(vp_S_Conexion)
            'validamos si la insercion de la primera carga fue exitosa 
            If CInicial = "ERROR" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("fallo al insertar la carga inicial a la tabla S_CGA consulte con admin Sasif S.A.")
                Console.ReadLine()
                vl_S_Result = CInicial
                GoTo SALIDA
            End If
        Else

            Console.ForegroundColor = ConsoleColor.Cyan
            Console.WriteLine("------------------------------------------------------------------------")
            Console.WriteLine("- PASO 2.                                                              -")
            Console.WriteLine("------------------------------------------------------------------------")
            Console.WriteLine("- Preparando Update en la tabla S_CGA...")
            Console.WriteLine("------------------------------------------------------------------------")

            'contamos cuantos registros debe actualizar
            Dim CountUpdate As String = S_CGA_1.Count_updateInsert(vp_S_Conexion, "1")
            If CountUpdate = "0" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("El achivo de Ecxel no actualizaciones pendientes")
                vl_S_Result = CountUpdate
                GoTo SALIDA
            End If
            'llamamos la funcion que cuenta cuantos registros debe actualizar
            Dim updateCGA As String = S_CGA_1.Update_CGA(vp_S_Conexion)
            If updateCGA = "ERROR" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("fallo al actualizar en la tabla S_CGA consulte con admin Sasif S.A.")
                Console.ReadLine()
                vl_S_Result = updateCGA
                GoTo SALIDA
            End If
            'llamamos la funcion que borra los registros ya actualizados
            S_CGA_1.DEL_TEMP_Update(vp_S_Conexion)
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.WriteLine("------------------------------------------------------------------------")
            Console.WriteLine("- PASO 3.                                                              -")
            Console.WriteLine("------------------------------------------------------------------------")
            Console.WriteLine("- Preparando Insert en la tabla S_CGA...")
            Console.WriteLine("------------------------------------------------------------------------")
            'contamos cuantos registros debe insertar
            Dim Countinsert As String = S_CGA_1.Count_updateInsert(vp_S_Conexion, "0")
            If Countinsert = "0" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("El achivo de Ecxel no tiene nuevos registros")
                vl_S_Result = Countinsert
                GoTo SALIDA
            End If
            'llamamos la funcion que cuenta cuantos registros debe insertat
            Dim InsertCGA As String = S_CGA_1.Insert_CGA(vp_S_Conexion)
            If InsertCGA = "ERROR" Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("fallo al insertar en la tabla S_CGA consulte con admin Sasif S.A.")
                Console.ReadLine()
                vl_S_Result = InsertCGA
                GoTo SALIDA
            End If

        End If


SALIDA:
        Return vl_S_Result

    End Function

    Public Function SABANA_3(ByVal vp_S_Ruta As String, ByVal vp_S_Conexion As String)

        Dim vl_S_Result As String = "OK"

        'llamamos la funcion copia del excel a la tabla temporal
        Dim vl_S_copy As String = copyTemporal(vp_S_Ruta, vp_S_Conexion, "SABANA_3")
        'validamos si la copia fue exitosa 
        Select Case vl_S_copy
            Case "ERROR"
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("")
                Console.WriteLine("Procedimiento BulkCopy fallo el achivo Excel es diferente a la data maestra o el archivo (SABANA3_data.xls) esta abierto!")
                Console.ReadLine()
                vl_S_Result = vl_S_copy
                GoTo SALIDA

            Case "PASO"
                vl_S_Result = vl_S_copy
                GoTo SALIDA

        End Select


        'llamamos la funcion que actualiza las llaves en la tabla temporal
        Dim UpdateKey As String = S_CGA_3.Update_Keys(vp_S_Conexion)
        'validamos si la actualizacion de llaves fue exitosa 
        If UpdateKey = "ERROR" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("")
            Console.WriteLine("fallo al actualizar las llaves consulte con admin Sasif S.A.")
            Console.ReadLine()
            vl_S_Result = UpdateKey
            GoTo SALIDA
        End If

SALIDA:

        Return vl_S_Result

    End Function

#End Region

#Region "FUNCIONES"

    ''' <summary>
    ''' TRAE EL ENCABEZADO DEL DESARROLLO
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ENTRADA()

        Dim RESULT As String = "OK"

        Console.Title = "Consola de Cargue masivo"

        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("-               CONSOLA OERACIONAL PARA SABANAS CGA                    -")
        Console.WriteLine("-                    DESARROLLADO POR SASIF S.A                        -")
        Console.WriteLine("-                               2016                                   -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("                   " & DateTime.Now & "                        ")
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("------------------------------------------------------------------------")

        Return RESULT

    End Function

    ''' <summary>
    ''' construye el encabesado de SABANA 1 y ejecuta el proceso SABANA 1
    ''' </summary>
    ''' <param name="vp_S_Ruta"></param>
    ''' <param name="vp_S_Conexion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ENT_SABANA_1(ByVal vp_S_Ruta As String, ByVal vp_S_Conexion As String)

        Dim vl_S_Result As String = ""

        Console.WriteLine("")
        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("-                              SABANA 1                                -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("- PASO 1.                                                              -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("--> Abriendo archivo Excel...")
        'llamamos el proceso de SABANA 1
        vl_S_Result = SABANA_1(vp_S_Ruta, vp_S_Conexion)
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("------------------------------------------------------------------------")

        Return vl_S_Result

    End Function

    Public Function ENT_SABANA_3(ByVal vp_S_Ruta As String, ByVal vp_S_Conexion As String)

        Dim vl_S_Result As String = ""

        Console.WriteLine("")
        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("-                              SABANA 3                                -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("- PASO 1.                                                              -")
        Console.WriteLine("------------------------------------------------------------------------")
        Console.WriteLine("")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("--> Abriendo archivo Excel...")
        'llamamos el proceso de SABANA 1
        vl_S_Result = SABANA_3(vp_S_Ruta, vp_S_Conexion)
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.WriteLine("------------------------------------------------------------------------")

        Return vl_S_Result

    End Function

    ''' <summary>
    ''' leer ruta del txt
    ''' </summary>
    ''' <param name="vp_S_RutaAPP"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Readruta(ByVal vp_S_RutaAPP As String, ByVal vp_S_type As String)

        Dim vl_S_strRuta As String = ""
        Dim vl_S_Name As String = ""

        Select Case vp_S_type
            Case "0"
                vl_S_Name = "\STR_Connection.txt"
            Case "1"
                vl_S_Name = "\Excel_ruta.txt"
            Case "2"
                vl_S_Name = "\Excel_ruta_S3.txt"
        End Select

        Dim vl_S_PathArchivo As String = vp_S_RutaAPP & vl_S_Name

     
        Try
            Dim objLeer As New StreamReader(vl_S_PathArchivo)
            Dim sLinea As String = ""
            Dim arrText As New ArrayList()

            Do
                sLinea = objLeer.ReadLine()
                If Not sLinea Is Nothing Then
                    arrText.Add(sLinea)
                End If
            Loop Until sLinea Is Nothing
            objLeer.Close()

            For Each sLinea In arrText
                vl_S_strRuta = sLinea
            Next

        Catch ex As Exception
            vl_S_strRuta = "ERROR"
        End Try
        Return vl_S_strRuta

    End Function

    ''' <summary>
    ''' hace bulkcopy del excel la tablatemporal SABANA 1
    ''' </summary>
    ''' <param name="vp_S_NameExcel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function copyTemporal(ByVal vp_S_NameExcel As String, ByVal vp_S_conexion As String, ByVal vp_S_Sabana As String)

        'abrimos coneccion OleBD para el archio excel
        Dim On_conex As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & vp_S_NameExcel & "; Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";")
        Dim Result As String = ""
        Dim vl_S_selectHoja As String = ""
        Dim vl_S_Tabla As String = ""

        Try

            Select Case vp_S_Sabana

                Case "SABANA_1"
                    vl_S_selectHoja = "select * from [Sabana 1$B9:BP1048575]"
                    vl_S_Tabla = "TEMP_SABANA"

                Case "SABANA_3"
                    vl_S_selectHoja = "select * from [Sabana 3$A4:CH1048575]"
                    vl_S_Tabla = "TEMP_SABANA3"

            End Select

            'Despues de conectarse al archivo excel seleccionamos los datos de la hoja por el nombre

            Dim On_cmd As New OleDbCommand(vl_S_selectHoja, On_conex)

            On_conex.Open()
            'leemos el excel
            Dim odr_Ecxel As OleDbDataReader = On_cmd.ExecuteReader()
            Console.WriteLine("")
            Console.WriteLine("--> leyendo archivo Excel...")

            Using SqlBulk As SqlBulkCopy = New SqlBulkCopy(vp_S_conexion)

                Console.WriteLine("")
                Console.ForegroundColor = ConsoleColor.DarkYellow
                Console.WriteLine("--> Copiando archivo...")

                SqlBulk.DestinationTableName = vl_S_Tabla
                'copiamos los datos en la tabla
                SqlBulk.WriteToServer(odr_Ecxel)

                Result = "OK"
            End Using

        Catch ex As Exception
            If vp_S_Sabana = "SABANA_3" Then
                Result = "PASO"
            Else
                Result = "ERROR"

            End If
        End Try

        Return Result

    End Function

#End Region

End Module
