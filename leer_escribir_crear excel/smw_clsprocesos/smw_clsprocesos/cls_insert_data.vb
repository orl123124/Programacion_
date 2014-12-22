

Imports System.IO

Public Class cls_insert_data
    Dim fecha As Date = Now
    Dim clshelper As New Helper.ClsHelperOra
    'Subrutina que crea un archivo .txt como log del Sistema
    Public Sub CreaLogFile(ByVal contenido As String)
        Dim dia As String = fecha.Day & fecha.Month & fecha.Year
        Dim fso As New System.IO.FileStream("C:\smw_log\log_smw" & dia & ".txt", FileMode.Append, FileAccess.Write)

        Dim w As New StreamWriter(fso)

        w.WriteLine(" ")
        w.WriteLine(contenido)

        w.Close()
        fso = Nothing
        w = Nothing
    End Sub

    Public Sub CreaLogFile_clsprocesos(ByVal contenido As String)
        Dim dia As String = fecha.Day & fecha.Month & fecha.Year
        Dim fso As New System.IO.FileStream("C:\smw_log\log_smw_clsprocesos" & dia & ".txt", FileMode.Append, FileAccess.Write)

        Dim w As New StreamWriter(fso)

        w.WriteLine(" ")
        w.WriteLine(contenido)

        w.Close()
        fso = Nothing
        w = Nothing
    End Sub

    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_Turnos    

    Public Function Insert_HCalculadas(ByVal value_empleado As Long, ByVal value_fechamarc As String, ByVal value_hentry1 As String, ByVal value_hsalida2 As String, ByVal value_tipodia As String, ByVal value_turnos As String, ByVal value_hextras As Double, ByVal value_henfermedad As Double, ByVal value_hpermisop As Double, ByVal value_hajustes As Double, ByVal value_usuario As String, ByVal value_supervisor As Integer, ByVal value_mercado As Integer, ByVal value_reloj As Integer, ByVal value_aut_hextra As String, ByVal value_henfer_css As Double) As Integer

        'ByVal value_hentry2 As String, ByVal value_hsalida1 As String,
        Dim arreglo_paramvalue(16) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fechamarc")
            arreglo_paramname.Add("p_hentry1")
            'arreglo_paramname.Add("p_hentry2")
            'arreglo_paramname.Add("p_hsalida1")
            arreglo_paramname.Add("p_hsalida2")
            arreglo_paramname.Add("p_tipodia")
            arreglo_paramname.Add("p_turnos")
            arreglo_paramname.Add("p_hextras")
            arreglo_paramname.Add("p_henfermedad")
            arreglo_paramname.Add("p_hpermisop")
            arreglo_paramname.Add("p_hajustes")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_supervisor")
            arreglo_paramname.Add("p_mercado")
            arreglo_paramname.Add("p_reloj")
            arreglo_paramname.Add("p_autoriza_hextra")
            arreglo_paramname.Add("p_henfer_css")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fechamarc
            arreglo_paramvalue(3) = value_hentry1
            'arreglo_paramvalue(4) = value_hentry2
            'arreglo_paramvalue(5) = value_hsalida1
            arreglo_paramvalue(4) = value_hsalida2
            arreglo_paramvalue(5) = value_tipodia
            arreglo_paramvalue(6) = value_turnos
            arreglo_paramvalue(7) = value_hextras
            arreglo_paramvalue(8) = value_henfermedad
            arreglo_paramvalue(9) = value_hpermisop
            arreglo_paramvalue(10) = value_hajustes
            arreglo_paramvalue(11) = value_usuario
            arreglo_paramvalue(12) = value_supervisor
            arreglo_paramvalue(13) = value_mercado
            arreglo_paramvalue(14) = value_reloj
            arreglo_paramvalue(15) = value_aut_hextra
            arreglo_paramvalue(16) = value_henfer_css





            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_inserthcalculada", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            'lblError.Text = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            ' objproc.CreaLogFile("En la Pantalla de Modificacion de Marcaciones - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)
            '            mensaje1 = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Insert_HCalculadas, ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return rows_affected

    End Function

    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_Turnos    
    Public Function InsertTurnos(ByVal value_empleado As Integer, ByVal value_SEMANA As Integer, _
    ByVal value_JORNADA_LUNES As String, ByVal value_JORNADA_MARTES As String, ByVal value_JORNADA_MIERCOLES As String, _
    ByVal value_JORNADA_JUEVES As String, ByVal value_JORNADA_VIERNES As String, ByVal value_JORNADA_SABADO As String, _
    ByVal value_JORNADA_DOMINGO As String, ByVal value_TIPO_LUNES As String, ByVal value_TIPO_MARTES As String, _
    ByVal value_TIPO_MIERCOLES As String, ByVal value_TIPO_JUEVES As String, ByVal value_TIPO_VIERNES As String, _
    ByVal value_TIPO_SABADO As String, ByVal value_TIPO_DOMINGO As String, ByVal value_MODIFICADO_USER As String, _
    ByVal value_CREACION_FECHA As Date, ByVal value_PERIODO As Integer) As Integer


        'sp_InsertTurnos

        Dim arreglo_paramvalue(20) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try


            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist

            arreglo_paramname.Add("p_EMPLEADO")
            arreglo_paramname.Add("p_SEMANA")
            arreglo_paramname.Add("p_JORNADA_LUNES")
            arreglo_paramname.Add("p_JORNADA_MARTES")
            arreglo_paramname.Add("p_JORNADA_MIERCOLES")
            arreglo_paramname.Add("p_JORNADA_JUEVES")
            arreglo_paramname.Add("p_JORNADA_VIERNES")
            arreglo_paramname.Add("p_JORNADA_SABADO")
            arreglo_paramname.Add("p_JORNADA_DOMINGO")
            arreglo_paramname.Add("p_TIPO_LUNES")
            arreglo_paramname.Add("p_TIPO_MARTES")
            arreglo_paramname.Add("p_TIPO_MIERCOLES")
            arreglo_paramname.Add("p_TIPO_JUEVES")
            arreglo_paramname.Add("p_TIPO_VIERNES")
            arreglo_paramname.Add("p_TIPO_SABADO")
            arreglo_paramname.Add("p_TIPO_DOMINGO")
            arreglo_paramname.Add("p_MODIFICADO_USER")
            arreglo_paramname.Add("p_MODIFICADO_FECHA")
            arreglo_paramname.Add("p_CREACION_FECHA")
            arreglo_paramname.Add("p_PERIODO")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object



            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_SEMANA
            arreglo_paramvalue(3) = value_JORNADA_LUNES
            arreglo_paramvalue(4) = value_JORNADA_MARTES
            arreglo_paramvalue(5) = value_JORNADA_MIERCOLES
            arreglo_paramvalue(6) = value_JORNADA_JUEVES
            arreglo_paramvalue(7) = value_JORNADA_VIERNES
            arreglo_paramvalue(8) = value_JORNADA_SABADO
            arreglo_paramvalue(9) = value_JORNADA_DOMINGO
            arreglo_paramvalue(10) = value_TIPO_LUNES
            arreglo_paramvalue(11) = value_TIPO_MARTES
            arreglo_paramvalue(12) = value_TIPO_MIERCOLES
            arreglo_paramvalue(13) = value_TIPO_JUEVES
            arreglo_paramvalue(14) = value_TIPO_VIERNES
            arreglo_paramvalue(15) = value_TIPO_SABADO
            arreglo_paramvalue(16) = value_TIPO_DOMINGO
            arreglo_paramvalue(17) = value_MODIFICADO_USER
            arreglo_paramvalue(18) = value_CREACION_FECHA
            arreglo_paramvalue(19) = value_CREACION_FECHA
            arreglo_paramvalue(20) = value_PERIODO


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertTurnos", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - InsertTurnos, ocurrió lo siguiente: " & " Proc-InsertTurnos" & ex.ToString)

        End Try

        Return rows_affected



    End Function

    Public Function InsertBitacora_Acceso(ByVal value_ACCION As String, ByVal value_DESCRIPCCION As String, _
    ByVal value_USUARIO_ROL As String, ByVal value_NOMBRE_PAG As String, ByVal value_ID_EMP_MOD As String) As Integer



        'sp_InsertTurnos
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer
        Dim value_FECHA As Date = Now

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_FECHA")
            arreglo_paramname.Add("p_ACCION")
            arreglo_paramname.Add("p_DESCRIPCCION")
            arreglo_paramname.Add("p_USUARIO_ROL")
            arreglo_paramname.Add("p_NOMBRE_PAG")
            arreglo_paramname.Add("p_ID_EMP_MOD")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_FECHA
            arreglo_paramvalue(2) = value_ACCION
            arreglo_paramvalue(3) = value_DESCRIPCCION
            arreglo_paramvalue(4) = value_USUARIO_ROL
            arreglo_paramvalue(5) = value_NOMBRE_PAG
            arreglo_paramvalue(6) = value_ID_EMP_MOD


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_Insert_Bitacora_Acceso", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Insert_Bitacora_Acceso, ocurrió lo siguiente: " & " Proc-sp_Insert_Bitacora_Acceso" & ex.ToString)
        End Try

        Return rows_affected

    End Function

    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_Pantallas
    Public Function InsertPantallas(ByVal value_nombre As String, ByVal value_descripcion As String, ByVal value_parentnode As String) As Integer

        'sp_InsertTurnos

        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist

            arreglo_paramname.Add("p_nombre")
            arreglo_paramname.Add("p_descrip")
            arreglo_paramname.Add("p_parentnode")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_nombre
            arreglo_paramvalue(2) = value_descripcion
            arreglo_paramvalue(3) = value_parentnode


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertPantallas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - InsertPantallas, ocurrió lo siguiente: " & " Proc-InsertTurnos" & ex.ToString)
        End Try


        Return rows_affected


    End Function

    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_Pantallas
    Public Function InsertRolesPantallas(ByVal value_rol As Integer, ByVal value_pantalla As Integer) As Integer

        'sp_InsertTurnos

        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist

            arreglo_paramname.Add("p_id_rol")
            arreglo_paramname.Add("p_id_pantalla")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_rol
            arreglo_paramvalue(2) = value_pantalla


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertRolesPantallas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - InsertRolesPantallas, ocurrió lo siguiente: " & " Proc-InsertTurnos" & ex.ToString)
        End Try


        Return rows_affected


    End Function


    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_Roles
    Public Function InsertRoles(ByVal value_rol As String, ByVal value_permiso As String) As Integer

        'sp_InsertTurnos

        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist

            arreglo_paramname.Add("p_rol")
            arreglo_paramname.Add("p_permiso")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_rol
            arreglo_paramvalue(2) = value_permiso



            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertRoles", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - InsertRoles, ocurrió lo siguiente: " & " Proc-InsertTurnos" & ex.ToString)
        End Try


        Return rows_affected


    End Function

    'Funcion que se utiliza para insertar un registro nuevo en la Tabla smw_relojes_sitios
    Public Function Insert_SitiosReloj(ByVal value_id_sitio As Long, ByVal value_descripcion As String) As Integer

        Dim listar_relojes As New DataSet
        Dim rows_affected As Integer



        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")
            arreglo_paramname.Add("p_id_sitio")
            arreglo_paramname.Add("p_descripcion")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_id_sitio
            arreglo_paramvalue(2) = value_descripcion



            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertSitios", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_SitiosReloj ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try
        Return rows_affected

    End Function

    Public Function sp_InsertUsuarios(ByVal value_id_usuario As Long, ByVal value_nombre As String, ByVal value_id_sito As Long, ByVal value_status As String, ByVal value_id_roles As Long, ByVal value_id_reloj As Long) As Integer





        Dim listar_relojes As New DataSet
        Dim rows_affected As Integer
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")



            arreglo_paramname.Add("p_id_usuario")
            arreglo_paramname.Add("p_nombre")
            arreglo_paramname.Add("p_id_sitio")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_id_roles")
            arreglo_paramname.Add("p_id_reloj")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_id_usuario
            arreglo_paramvalue(2) = value_nombre
            arreglo_paramvalue(3) = value_id_sito
            arreglo_paramvalue(4) = value_status
            arreglo_paramvalue(5) = value_id_roles
            arreglo_paramvalue(6) = value_id_reloj


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertUsuarios", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion inserta usuraios ocurrió lo siguiente: " & ex.ToString)
        Finally

        End Try

        Return rows_affected





    End Function



    Public Function sp_InsertDiasFreiados(ByVal value_periodo As Long, ByVal value_mes As Long, ByVal value_dia As Long, ByVal value_descripcion As String, ByVal value_dia_puente As Long) As Integer





        Dim listar_relojes As New DataSet
        Dim rows_affected As Integer



        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")


            arreglo_paramname.Add("p_periodo")
            arreglo_paramname.Add("p_mes")
            arreglo_paramname.Add("p_dia")
            arreglo_paramname.Add("p_descripcion")
            arreglo_paramname.Add("p_dia_puente")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_periodo
            arreglo_paramvalue(2) = value_mes
            arreglo_paramvalue(3) = value_dia
            arreglo_paramvalue(4) = value_descripcion
            arreglo_paramvalue(5) = value_dia_puente



            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertDiasFeriados", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion inserta dias feriados ocurrió lo siguiente: " & ex.ToString)

        Finally
        End Try

        Return rows_affected
    End Function
    Public Function InsertHorasJornadas(ByVal v_codigo As String, _
                                        ByVal v_hentry1 As String, _
                                        ByVal v_salida1 As String, _
                                        ByVal v_tipoturno As String, _
                                        ByVal v_tothrreg As Double, _
                                        ByVal v_hrpagar As Double, _
                                        ByVal v_pagadif As String, _
                                        ByVal v_usuario As String, _
                                        ByVal v_descanso As String)

        'ByVal v_hentry2 As String, _
        'ByVal v_salida2 As String, _

        Dim arreglo_paramvalue(9) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")


            arreglo_paramname.Add("p_codigo")
            arreglo_paramname.Add("p_DH1")
            'arreglo_paramname.Add("p_DH2")
            arreglo_paramname.Add("p_HH1")
            'arreglo_paramname.Add("p_HH2")
            arreglo_paramname.Add("p_tipoturno")
            arreglo_paramname.Add("p_tothrreg")
            arreglo_paramname.Add("p_hrpagar")
            arreglo_paramname.Add("p_pagadif")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_tdescanso")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = v_codigo
            arreglo_paramvalue(2) = v_hentry1
            'arreglo_paramvalue(3) = v_hentry2
            arreglo_paramvalue(3) = v_salida1
            'arreglo_paramvalue(5) = v_salida2
            arreglo_paramvalue(4) = v_tipoturno
            arreglo_paramvalue(5) = v_tothrreg
            arreglo_paramvalue(6) = v_hrpagar
            arreglo_paramvalue(7) = v_pagadif
            arreglo_paramvalue(8) = v_usuario
            arreglo_paramvalue(9) = v_descanso


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_insertHorasJornadas", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion InsertHorasJornadas ocurrió lo siguiente: " & ex.ToString)

        Finally
        End Try

        Return rows_affected


    End Function

    Public Function Insertar_horas_marcaciones(ByVal val_empresa As Integer, _
                                        ByVal val_empleado As Integer, _
                                        ByVal val_tiempo_reloj As String, _
                                        ByVal val_mes As Integer, _
                                        ByVal val_dia As Integer, _
                                        ByVal val_ano As Integer) As Integer

        Dim horas_marcacion As Integer
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            arreglo_paramname.Add("p_empresa")
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_tiempo_reloj")
            arreglo_paramname.Add("p_mes")
            arreglo_paramname.Add("p_dia")
            arreglo_paramname.Add("p_ano")

            arreglo_paramvalue(1) = val_empresa
            arreglo_paramvalue(2) = val_empleado
            arreglo_paramvalue(3) = val_tiempo_reloj
            arreglo_paramvalue(4) = val_mes
            arreglo_paramvalue(5) = val_dia
            arreglo_paramvalue(6) = val_ano

            horas_marcacion = clshelper.Ejecutar("smw_pkg_helper.sp_insertar_horas_marcaciones", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Insertar_horas_marcaciones ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return horas_marcacion

    End Function

    Public Function sp_InsertHorasPeriodos _
      (ByVal value_fecha1_pago As String, ByVal value_fecha2_pago As String, _
      ByVal value_fecha1_marc As String, ByVal value_fecha2_marc As String, _
      ByVal value_ultimoPeridos As Integer, ByVal value_usuario_cambio As String, _
      ByVal value_estatus As String) As Integer


        Dim listar_relojes As New DataSet
        Dim rows_affected As Integer



        Dim arreglo_paramvalue(7) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")


            arreglo_paramname.Add("p_fecha1_pago")
            arreglo_paramname.Add("p_fecha2_pago")
            arreglo_paramname.Add("p_fecha1_marc")
            arreglo_paramname.Add("p_fecha2_marc")
            arreglo_paramname.Add("p_UltimoPerido")
            arreglo_paramname.Add("p_usuario_cambio")
            arreglo_paramname.Add("p_estatus")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_fecha1_pago
            arreglo_paramvalue(2) = value_fecha2_pago
            arreglo_paramvalue(3) = value_fecha1_marc
            arreglo_paramvalue(4) = value_fecha2_marc
            arreglo_paramvalue(5) = value_ultimoPeridos
            arreglo_paramvalue(6) = value_usuario_cambio
            arreglo_paramvalue(7) = value_estatus


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_insertHorasPeridos", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion inserta horas periodos ocurrió lo siguiente: " & ex.ToString)

        Finally
        End Try

        Return rows_affected
    End Function

    Public Function Insert_DiasMartenidad _
      (ByVal value_empleado As Integer, ByVal value_fecha As String, _
ByVal value_horas As Integer, ByVal value_status As String, _
ByVal value_anulacion As String) As Integer

        Dim rows_affected As Integer
        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")


            arreglo_paramname.Add("p_employee")
            arreglo_paramname.Add("p_fecha")
            arreglo_paramname.Add("p_horas")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_anulacion")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fecha
            arreglo_paramvalue(3) = value_horas
            arreglo_paramvalue(4) = value_status
            arreglo_paramvalue(5) = value_anulacion


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertDiasMaternidad", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion inserta dias maternidad ocurrió lo siguiente: " & ex.ToString)

        Finally
        End Try

        Return rows_affected
    End Function

    Public Function ejecuta_proceso(ByVal val_empleado As Integer) As Integer

        Dim ejecucion As Integer
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            arreglo_paramname.Add("p_empleado")

            arreglo_paramvalue(1) = val_empleado

            ejecucion = clshelper.Ejecutar("smw_pkg_helper.sp_ejecuta_procesos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion ejecuta_proceso ocurrió lo siguiente: " & ex.ToString)
        End Try

    End Function

    Public Function insert_faltantes_sarweb(ByVal val_mercado As String, ByVal val_documento_id As String, ByVal val_empleado_id As Integer, _
                                           ByVal val_jobcode As String, ByVal val_job_description As String, ByVal val_empleado_nombre As String, _
                                           ByVal val_fecha_faltante As Date, ByVal val_monto As Double, ByVal val_descripcion As String, _
                                           ByVal val_tipo_faltante As Integer, ByVal val_empleado_preparo As Integer, ByVal val_empleado_aprobo As Integer, _
                                           ByVal val_usuario As String, ByVal val_aplicadogl As String, _
                                           ByVal val_aplicadopr As String, ByVal val_cod_des As String)


        Dim rows_affected As Integer
        Dim arreglo_paramvalue(16) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_mercado")
            arreglo_paramname.Add("p_documento_id")
            arreglo_paramname.Add("p_empleado_id")
            arreglo_paramname.Add("p_jobcode")
            arreglo_paramname.Add("p_job_description")
            arreglo_paramname.Add("p_empleado_nombre")
            arreglo_paramname.Add("p_fecha_faltante")
            arreglo_paramname.Add("p_monto")
            arreglo_paramname.Add("p_descripcion")
            arreglo_paramname.Add("p_tipo_faltante")
            arreglo_paramname.Add("p_empleado_preparo")
            arreglo_paramname.Add("p_empleado_aprobo")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_aplicadogl")
            arreglo_paramname.Add("p_aplicadopr")
            arreglo_paramname.Add("p_cod_des")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = val_mercado
            arreglo_paramvalue(2) = val_documento_id
            arreglo_paramvalue(3) = val_empleado_id
            arreglo_paramvalue(4) = val_jobcode
            arreglo_paramvalue(5) = val_job_description
            arreglo_paramvalue(6) = val_empleado_nombre
            arreglo_paramvalue(7) = val_fecha_faltante
            arreglo_paramvalue(8) = val_monto
            arreglo_paramvalue(9) = val_descripcion
            arreglo_paramvalue(10) = val_tipo_faltante
            arreglo_paramvalue(11) = val_empleado_preparo
            arreglo_paramvalue(12) = val_empleado_aprobo
            arreglo_paramvalue(13) = val_usuario
            arreglo_paramvalue(14) = val_aplicadogl
            arreglo_paramvalue(15) = val_aplicadopr
            arreglo_paramvalue(16) = val_cod_des

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertFaltantesSw", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion insert_faltantes_sarweb ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function insert_bitacora_sarweb(ByVal val_bitacoraid As String, _
                                           ByVal val_descripcion As String, _
                                           ByVal val_valanterior As String, _
                                           ByVal val_usuario As String, _
                                           ByVal val_mercado As String, _
                                           ByVal val_documentoid As String, _
                                           ByVal val_tipo As Integer)


        Dim rows_affected As Integer
        Dim arreglo_paramvalue(7) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_bitacoraid")
            arreglo_paramname.Add("p_descripcion")
            arreglo_paramname.Add("p_valanterior")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_mercado")
            arreglo_paramname.Add("p_documentoid")
            arreglo_paramname.Add("p_tipo")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = val_bitacoraid
            arreglo_paramvalue(2) = val_descripcion
            arreglo_paramvalue(3) = val_valanterior
            arreglo_paramvalue(4) = val_usuario
            arreglo_paramvalue(5) = val_mercado
            arreglo_paramvalue(6) = val_documentoid
            arreglo_paramvalue(7) = val_tipo

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertBitacoraSw", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion insert_faltantes_sarweb ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function insert_temp(ByVal p_empleado As Integer, _
                                ByVal p_fecha_ini As String, _
                                ByVal p_tipo As String, _
                                ByVal p_usuario As String)

        Dim rows_affected As Integer
        Dim arreglo_paramvalue(4) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_usuario")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = p_empleado
            arreglo_paramvalue(2) = p_fecha_ini
            arreglo_paramvalue(3) = p_tipo
            arreglo_paramvalue(4) = p_usuario

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.SP_INSERT_TEMP", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion insert_temp ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    'agregado por echavez 29/07/2011
    Public Function insert_PeriodoDecimo(ByVal p_check_date As String, ByVal p_start_date As String, _
                                        ByVal p_end_date As String, ByVal p_status As Integer, ByVal p_periodo As Integer)


        Dim rows_affected As Integer
        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_check_date")
            arreglo_paramname.Add("p_start_date")
            arreglo_paramname.Add("p_end_date")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_periodo")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = p_check_date
            arreglo_paramvalue(2) = p_start_date
            arreglo_paramvalue(3) = p_end_date
            arreglo_paramvalue(4) = p_status
            arreglo_paramvalue(5) = p_periodo


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertPeriodoDecimo", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion sp_InsertPeriodoDecimo ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    'agregado por echavez 4/08/2011 isr
    Public Function insert_ISR(ByVal p_importe_lim As Double, ByVal p_porcentaje As Double, ByVal p_start_date As String, _
                              ByVal p_end_date As String, ByVal p_status As Integer, ByVal p_type As String)


        Dim rows_affected As Integer
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_importe_lim")
            arreglo_paramname.Add("p_porcentaje")
            arreglo_paramname.Add("p_start_date")
            arreglo_paramname.Add("p_end_date")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_type")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = p_importe_lim
            arreglo_paramvalue(2) = p_porcentaje
            arreglo_paramvalue(3) = p_start_date
            arreglo_paramvalue(4) = p_end_date
            arreglo_paramvalue(5) = p_status
            arreglo_paramvalue(6) = p_type


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_InsertISR", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion sp_InsertISR ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function insert_liq(ByVal p_empleado As Integer, _
                                ByVal p_fecha_liq As String, _
                                ByVal p_tipo As String, _
                                ByVal p_contrato As String, _
                                ByVal p_extra As String, _
                                ByVal p_usuario As String)

        Dim rows_affected As Integer
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha_liq")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_contrato")
            arreglo_paramname.Add("p_extra")
            arreglo_paramname.Add("p_usuario")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = p_empleado
            arreglo_paramvalue(2) = p_fecha_liq
            arreglo_paramvalue(3) = p_tipo
            arreglo_paramvalue(4) = p_contrato
            arreglo_paramvalue(5) = p_extra
            arreglo_paramvalue(6) = p_usuario

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.SP_INSERT_LIQ", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion insert_LIQ ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

#Region "ORBARRIA"
    Public Function insert_to_csscorrelativo(ByVal tabla As Hashtable) As Integer
        Dim arreglo_paramvalue(4) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_cedula")
            arreglo_paramname.Add("p_seguro")
            arreglo_paramname.Add("p_nombre")
            arreglo_paramname.Add("p_consecutivo")

            arreglo_paramvalue(1) = tabla("CEDULA")
            arreglo_paramvalue(2) = tabla("SEGURO")
            arreglo_paramvalue(3) = tabla("NOMBRE")
            arreglo_paramvalue(4) = tabla("CONSECUTIVO")

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_insert_to_csscorrelativo", arreglo_paramname, arreglo_paramvalue)
        Catch ex As Exception
            CreaLogFile_clsprocesos("ERROR >>clsprocesos >>funcion  insert_to_csscorrelativo : " & ex.Message.ToString)
            Throw New Exception(ex.Message)
        End Try
        Return rows_affected
    End Function

    Public Function inst_to_mastercorrelativo() As Integer

        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            rows_affected = clshelper.Ejecutar("smw_pkg_helper.SP_INST_CSSCORR", arreglo_paramname, arreglo_paramvalue)
        Catch ex As Exception
            CreaLogFile_clsprocesos("ERROR >>clsprocesos >>funcion  inst_to_mastercorrelativo : " & ex.Message.ToString)
        End Try
        Return rows_affected
    End Function

#End Region

End Class
