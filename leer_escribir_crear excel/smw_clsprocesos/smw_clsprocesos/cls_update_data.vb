Imports Helper

Public Class cls_update_data
    Dim clsinsertdata As New cls_insert_data
    Dim clshelper As New Helper.ClsHelperOra



    Public Function Actualiza_HorasCalculadas(ByVal value_hentry1 As String, _
                                              ByVal value_hentry2 As String, _
                                              ByVal value_hsalida1 As String, _
                                              ByVal value_hsalida2 As String, _
                                              ByVal value_tipodia As String, _
                                              ByVal value_turnos As String, _
                                              ByVal value_hextras As Double, _
                                              ByVal value_henfermedad As Double, _
                                              ByVal value_fecha As String, _
                                              ByVal value_empleado As Integer, _
                                              ByVal value_usuario As String, _
                                              ByVal value_supervisor As String, _
                                              ByVal value_henfermedadnp As Double, _
                                              ByVal value_hajustes As Double, _
                                              ByVal value_aut_hextra As String, _
                                              ByVal value_henfer_css As Double, _
                                              ByVal value_status As String) As Integer

        'ByVal value_hentry1 As String, _
        'ByVal value_hentry2 As String, _
        'ByVal value_hsalida1 As String, _
        'ByVal value_hsalida2 As String, _

        Dim arreglo_paramvalue(17) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_hentry1")
            arreglo_paramname.Add("p_hentry2")
            arreglo_paramname.Add("p_hsalida1")
            arreglo_paramname.Add("p_hsalida2")
            arreglo_paramname.Add("p_tipodia")
            arreglo_paramname.Add("p_turnos")
            arreglo_paramname.Add("p_hextras")
            arreglo_paramname.Add("p_henfermedad")
            arreglo_paramname.Add("p_fecha")
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_supervisor")
            arreglo_paramname.Add("p_hpermisop")
            arreglo_paramname.Add("p_hajustes")
            arreglo_paramname.Add("p_autoriza_hextra")
            arreglo_paramname.Add("p_henfer_css")
            arreglo_paramname.Add("p_status")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_hentry1
            arreglo_paramvalue(2) = value_hentry2
            arreglo_paramvalue(3) = value_hsalida1
            arreglo_paramvalue(4) = value_hsalida2
            arreglo_paramvalue(5) = value_tipodia
            arreglo_paramvalue(6) = value_turnos
            arreglo_paramvalue(7) = value_hextras
            arreglo_paramvalue(8) = value_henfermedad
            arreglo_paramvalue(9) = value_fecha
            arreglo_paramvalue(10) = value_empleado
            arreglo_paramvalue(11) = value_usuario
            arreglo_paramvalue(12) = value_supervisor
            arreglo_paramvalue(13) = value_henfermedadnp
            arreglo_paramvalue(14) = value_hajustes
            arreglo_paramvalue(15) = value_aut_hextra
            arreglo_paramvalue(16) = value_henfer_css
            arreglo_paramvalue(17) = value_status

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_actualizahcalculada", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            'lblError.Text = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            ' objproc.CreaLogFile("En la Pantalla de Modificacion de Marcaciones - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)
            '            mensaje1 = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return rows_affected

    End Function

    Public Function Actualiza_TurnoMarcacion(ByVal value_hentry1 As String, _
                                             ByVal value_hsalida2 As String, _
                                             ByVal value_tipodia As String, _
                                             ByVal value_turnos As String, _
                                             ByVal value_fecha As String, _
                                             ByVal value_empleado As Integer, _
                                             ByVal value_usuario As String, _
                                             ByVal value_tipoemp As String, _
                                             ByVal value_mercado As Integer) As Integer
        'ByVal value_hentry2 As String, _
        'ByVal value_hsalida1 As String, _

        'Dim arreglo_paramvalue(10) As Object
        Dim arreglo_paramvalue(9) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_hentry1")
            'arreglo_paramname.Add("p_hentry2")
            'arreglo_paramname.Add("p_hsalida1")
            arreglo_paramname.Add("p_hsalida2")
            arreglo_paramname.Add("p_tipodia")
            arreglo_paramname.Add("p_turnos")
            arreglo_paramname.Add("p_fecha")
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_tipoemp")
            arreglo_paramname.Add("p_mercado")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_hentry1
            'arreglo_paramvalue(2) = value_hentry2
            'arreglo_paramvalue(3) = value_hsalida1
            arreglo_paramvalue(2) = value_hsalida2
            arreglo_paramvalue(3) = value_tipodia
            arreglo_paramvalue(4) = value_turnos
            arreglo_paramvalue(5) = value_fecha
            arreglo_paramvalue(6) = value_empleado
            arreglo_paramvalue(7) = value_usuario
            arreglo_paramvalue(8) = value_tipoemp
            arreglo_paramvalue(9) = value_mercado

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_actualizaturno_marcacion", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            'lblError.Text = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            ' objproc.CreaLogFile("En la Pantalla de Modificacion de Marcaciones - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)
            '            mensaje1 = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return rows_affected

    End Function
    Public Function UpdateTurnos(ByVal VALUE_NEMPLEADO As Long, ByVal value_NSEMANA As Long, ByVal value_JORNADA_LUNES As String, ByVal value_JORNADA_MARTES As String, ByVal value_JORNADA_MIERCOLES As String, ByVal value_JORNADA_JUEVES As String, ByVal value_JORNADA_VIERNES As String, ByVal value_JORNADA_SABADO As String, ByVal value_JORNADA_DOMINGO As String, ByVal value_TIPO_LUNES As String, ByVal value_TIPO_MARTES As String, ByVal value_TIPO_MIERCOLES As String, ByVal value_TIPO_JUEVES As String, ByVal value_TIPO_VIERNES As String, ByVal value_TIPO_SABADO As String, ByVal value_TIPO_DOMINGO As String, ByVal value_usuario As String, ByVal value_periodo As Integer) As Integer

        Dim arreglo_paramvalue(18) As Object
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
            arreglo_paramname.Add("p_PERIODO")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = VALUE_NEMPLEADO
            arreglo_paramvalue(2) = value_NSEMANA
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
            arreglo_paramvalue(17) = value_usuario
            arreglo_paramvalue(18) = value_periodo



            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_UpdateTurnos", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion UpdateTurnos, ocurrió lo siguiente: " & ex.ToString)

        End Try

        Return rows_affected



    End Function
    Public Function sp_UpdateHorasPeriodos _
      (ByVal value_numero As Integer, ByVal value_fecha1_pago As String, ByVal value_fecha2_pago As String, _
      ByVal value_fecha1_marc As String, ByVal value_fecha2_marc As String, _
      ByVal value_ultimoPeridos As Integer, ByVal value_usuario_cambio As String, _
      ByVal value_estatus As String) As Integer


        Dim listar_relojes As New DataSet
        Dim rows_affected As Integer

        Dim arreglo_paramvalue(8) As Object
        Dim arreglo_paramname As New ArrayList



        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_numero")
            arreglo_paramname.Add("p_fecha1_pago")
            arreglo_paramname.Add("p_fecha2_pago")
            arreglo_paramname.Add("p_fecha1_marc")
            arreglo_paramname.Add("p_fecha2_marc")
            arreglo_paramname.Add("p_ultimoperido")
            arreglo_paramname.Add("p_usuario_cambio")
            arreglo_paramname.Add("p_estatus")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_numero
            arreglo_paramvalue(2) = value_fecha1_pago
            arreglo_paramvalue(3) = value_fecha2_pago
            arreglo_paramvalue(4) = value_fecha1_marc
            arreglo_paramvalue(5) = value_fecha2_marc
            arreglo_paramvalue(6) = value_ultimoPeridos
            arreglo_paramvalue(7) = value_usuario_cambio
            arreglo_paramvalue(8) = value_estatus


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_actualizaPeriodos", arreglo_paramname, arreglo_paramvalue)



        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Update horas periodos ocurrió lo siguiente: " & ex.ToString)

        Finally
        End Try

        Return rows_affected
    End Function

    Public Function UpdateTurnos_Marcaciones(ByVal VALUE_NEMPLEADO As Long, ByVal value_NSEMANA As Long, ByVal value_JORNADA_NAME As String, ByVal value_JORNADA As String, ByVal value_TIPO As String, ByVal value_periodo As String, ByVal value_usuario As String) As Integer



        Dim arreglo_paramvalue(7) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_EMPLEADO")
            arreglo_paramname.Add("p_SEMANA")
            arreglo_paramname.Add("p_JORNADA_NAME")
            arreglo_paramname.Add("p_JORNADA")
            arreglo_paramname.Add("p_TIPO")
            arreglo_paramname.Add("p_PERIODO")
            arreglo_paramname.Add("p_MODIFICADO_USER")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = VALUE_NEMPLEADO
            arreglo_paramvalue(2) = value_NSEMANA
            arreglo_paramvalue(3) = value_JORNADA_NAME
            arreglo_paramvalue(4) = value_JORNADA
            arreglo_paramvalue(5) = value_TIPO
            arreglo_paramvalue(6) = value_periodo
            arreglo_paramvalue(7) = value_usuario


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_UpdateTurnos_Marcaciones", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion UpdateTurnos, ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return rows_affected
    End Function
    Public Function UpdateHorasJornadas(ByVal v_codigo As String, _
                                        ByVal v_hentry1 As String, _
                                        ByVal v_salida1 As String, _
                                        ByVal v_descanso As String, _
                                        ByVal v_tipoturno As String, _
                                        ByVal v_tothrreg As Double, _
                                        ByVal v_hrpagar As Double, _
                                        ByVal v_pagadif As String, _
                                        ByVal v_usuario As String, _
                                        ByVal v_estatus As String)

        'ByVal v_hentry2 As String, _
        'ByVal v_salida2 As String, _

        Dim arreglo_paramvalue(10) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_codigo")
            arreglo_paramname.Add("p_DH1")
            'arreglo_paramname.Add("p_DH2")
            arreglo_paramname.Add("p_HH1")
            'arreglo_paramname.Add("p_HH2")
            arreglo_paramname.Add("p_descanso")
            arreglo_paramname.Add("p_tipoturno")
            arreglo_paramname.Add("p_tothrreg")
            arreglo_paramname.Add("p_hrpagar")
            arreglo_paramname.Add("p_pagadif")
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_estatus")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = v_codigo
            arreglo_paramvalue(2) = v_hentry1
            'arreglo_paramvalue(3) = v_hentry2
            arreglo_paramvalue(3) = v_salida1
            'arreglo_paramvalue(5) = v_salida2
            arreglo_paramvalue(4) = v_descanso
            arreglo_paramvalue(5) = v_tipoturno
            arreglo_paramvalue(6) = v_tothrreg
            arreglo_paramvalue(7) = v_hrpagar
            arreglo_paramvalue(8) = v_pagadif
            arreglo_paramvalue(9) = v_usuario
            arreglo_paramvalue(10) = v_estatus

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_actualizaHorasJornadas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion UpdateHorasJornadas, ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return rows_affected


    End Function
    'Funcion que actualiza los Parametros generales del sistema de Marcaciones Web
    Public Function Update_Parametros(ByVal CODIGO_REGULAR As String, ByVal CODIGO_DOMINGO As String, _
                                            ByVal CODIGO_COMPENSATORIO As String, ByVal CODIGO_FERIADO As String, _
                                            ByVal CODIGO_AUSENCIA As String, ByVal CODIGO_AJUSTE As String, _
                                            ByVal MINIMO_EXTRA As String, ByVal HORAMAXHEXTRA As String, _
                                            ByVal MAXIMO_EXTRA_DIARIO As String, ByVal MAXIMO_EXTRA_SEMANA As String, _
                                            ByVal HORAMAXHENFERMEDAD As String, ByVal HORAS_ENF_POR_DIA As String, _
                                            ByVal HORAS_ENF_A_PAGAR As String, ByVal CODIGO_ENFERMEDAD As String, _
                                            ByVal CODIGO_ENFERMEDAD_NEGATIVA As String, ByVal COD_ENF_DOM_REY As String, ByVal COD_ENF_FIESNAC_REY As String, _
                                            ByVal CODIGO_TARDANZA_DESCTO As String, ByVal CODIGO_TARDANZA_SIN_DESCTO As String, _
                                            ByVal TARDANZA_MIN_VALOR As String, ByVal MINIMO_TARDANZA As String, _
                                            ByVal TARDANZA_RANGO1 As String, ByVal TARDANZA_RANGO2 As String, _
                                            ByVal TARDANZA_RANGO3 As String, ByVal TARDANZA_RANGO4 As String, _
                                            ByVal TARDANZA_VALOR1 As String, ByVal TARDANZA_VALOR2 As String, _
                                            ByVal TARDANZA_VALOR3 As String, ByVal TARDANZA_VALOR4 As String) As Integer



        Dim arreglo_paramvalue(29) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            'Parametros_Marcaciones()
            arreglo_paramname.Add("p_CODIGO_REGULAR")
            arreglo_paramname.Add("p_CODIGO_DOMINGO")
            arreglo_paramname.Add("p_CODIGO_COMPENSATORIO")
            arreglo_paramname.Add("p_CODIGO_FERIADO")
            arreglo_paramname.Add("p_CODIGO_AUSENCIA")
            arreglo_paramname.Add("p_CODIGO_AJUSTE")
            'Parametros_Horas_Extras()
            arreglo_paramname.Add("p_MINIMO_EXTRA")
            arreglo_paramname.Add("p_HORAMAXHEXTRA")
            arreglo_paramname.Add("p_MAXIMO_EXTRA_DIARIO")
            arreglo_paramname.Add("p_MAXIMO_EXTRA_SEMANA")
            'Parametros_Enfermedad()
            arreglo_paramname.Add("p_HORAMAXHENFERMEDAD")
            arreglo_paramname.Add("p_HORAS_ENF_POR_DIA")
            arreglo_paramname.Add("p_HORAS_ENF_A_PAGAR")
            arreglo_paramname.Add("p_CODIGO_ENFERMEDAD")
            arreglo_paramname.Add("p_CODIGO_ENFERMEDAD_NEGATIVA")
            arreglo_paramname.Add("p_COD_ENF_DOM_REY")
            arreglo_paramname.Add("p_COD_ENF_FIESNAC_REY")
            'Parametros_Tardanza()
            arreglo_paramname.Add("p_CODIGO_TARDANZA_DESCTO")
            arreglo_paramname.Add("p_CODIGO_TARDANZA_SIN_DESCTO")
            arreglo_paramname.Add("p_TARDANZA_MIN_VALOR")
            arreglo_paramname.Add("p_MINIMO_TARDANZA")
            arreglo_paramname.Add("p_TARDANZA_RANGO1")
            arreglo_paramname.Add("p_TARDANZA_RANGO2")
            arreglo_paramname.Add("p_TARDANZA_RANGO3")
            arreglo_paramname.Add("p_TARDANZA_RANGO4")
            arreglo_paramname.Add("p_TARDANZA_VALOR1")
            arreglo_paramname.Add("p_TARDANZA_VALOR2")
            arreglo_paramname.Add("p_TARDANZA_VALOR3")
            arreglo_paramname.Add("p_TARDANZA_VALOR4")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            'Parametros_Marcaciones()
            arreglo_paramvalue(1) = CODIGO_REGULAR
            arreglo_paramvalue(2) = CODIGO_DOMINGO
            arreglo_paramvalue(3) = CODIGO_COMPENSATORIO
            arreglo_paramvalue(4) = CODIGO_FERIADO
            arreglo_paramvalue(5) = CODIGO_AUSENCIA
            arreglo_paramvalue(6) = CODIGO_AJUSTE
            'Parametros_Horas_Extras()
            arreglo_paramvalue(7) = MINIMO_EXTRA
            arreglo_paramvalue(8) = HORAMAXHEXTRA
            arreglo_paramvalue(9) = MAXIMO_EXTRA_DIARIO
            arreglo_paramvalue(10) = MAXIMO_EXTRA_SEMANA
            'Parametros_Enfermedad()
            arreglo_paramvalue(11) = HORAMAXHENFERMEDAD
            arreglo_paramvalue(12) = HORAS_ENF_POR_DIA
            arreglo_paramvalue(13) = HORAS_ENF_A_PAGAR
            arreglo_paramvalue(14) = CODIGO_ENFERMEDAD
            arreglo_paramvalue(15) = CODIGO_ENFERMEDAD_NEGATIVA
            arreglo_paramvalue(16) = COD_ENF_DOM_REY
            arreglo_paramvalue(17) = COD_ENF_FIESNAC_REY
            'Parametros_Tardanza()
            arreglo_paramvalue(18) = CODIGO_TARDANZA_DESCTO
            arreglo_paramvalue(19) = CODIGO_TARDANZA_SIN_DESCTO
            arreglo_paramvalue(20) = TARDANZA_MIN_VALOR
            arreglo_paramvalue(21) = MINIMO_TARDANZA
            arreglo_paramvalue(22) = TARDANZA_RANGO1
            arreglo_paramvalue(23) = TARDANZA_RANGO2
            arreglo_paramvalue(24) = TARDANZA_RANGO3
            arreglo_paramvalue(25) = TARDANZA_RANGO4
            arreglo_paramvalue(26) = TARDANZA_VALOR1
            arreglo_paramvalue(27) = TARDANZA_VALOR2
            arreglo_paramvalue(28) = TARDANZA_VALOR3
            arreglo_paramvalue(29) = TARDANZA_VALOR4


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_updateparametros", arreglo_paramname, arreglo_paramvalue)

            'smw_pkg_helper.sp_updateparametros(p_codigo_regular => :p_codigo_regular,
            '                       p_codigo_domingo => :p-codigo_domingo,
            '                       p_codigo_compensatorio => :p_codigo_compensatorio,
            '                       p_codigo_feriado => :p_codigo_feriado,
            '                       p_codigo_ausencia => :p_codigo_ausencia,
            '                       p_codigo_ajuste => :p_codigo_ajuste,
            '                       p_minimo_extra => :p_minimo_extra,
            '                       p_horamaxhextra => :p_horamaxhextra,
            '                       p_maximo_extra_diario => :p_maximo_extra_diario,
            '                       p_maximo_extra_semana => :p_maximo_extra_semana,
            '                       p_horamaxhenfermedad => :p_horamaxhenfermedad,
            '                       p_horas_enf_por_dia => :p_horas_enf_por_dia,
            '                       p_horas_enf_a_pagar => :p_horas_enf_a_pagar,
            '                       p_codigo_enfermedad => :p_codigo_enfermedad,
            '                       p_codigo_enfermedad_negativa => :p_codigo_enfermedad_negativa,
            '                       p_cod_enf_dom_rey => :p_cod_enf_dom_rey,
            '                       p_cod_enf_fiesnac_rey => :p_cod_enf_fiesnac_rey,
            '                       p_codigo_tardanza_descto => :p_codigo_tardanza_descto,
            '                       p_codigo_tardanza_sin_descto => :p_codigo_tardanza_sin_descto,
            '                       p_tardanza_min_valor => :p_tardanza_min_valor,
            '                       p_minimo_tardanza => :p_minimo_tardanza,
            '                       p_tardanza_rango1 => :p_tardanza_rango1,
            '                       p_tardanza_rango2 => :p_tardanza_rango2,
            '                       p_tardanza_rango3 => :p_tardanza_rango3,
            '                       p_tardanza_rango4 => :p_tardanza_rango4,
            '                       p_tardanza_valor1 => :p_tardanza_valor1,
            '                       p_tardanza_valor2 => :p_tardanza_valor2,
            '                       p_tardanza_valor3 => :p_tardanza_valor3,
            '                       p_tardanza_valor4 => :p_tardanza_valor4);

        Catch ex As Exception
            rows_affected = 0
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos[cls_update_data.vb] - Seccion Update_Parametros, ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return rows_affected
    End Function

    Public Function RecalcularSP() 'funcion que devuelve si se reproceso algun datos en sp_calculahora

        Dim rows_affected As Integer
        '-------- no se enviara parametro por eso no se asigna valores a arreglos

        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        '------------------

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")



            rows_affected = clshelper.Ejecutar("smw_pkg_calculos.sp_calculahoras2", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Recalcular Horas, ocurrió lo siguiente: " & " RecalcularSP" & ex.ToString)

        End Try
        Return rows_affected

    End Function


    Public Function CrearArchivoPago()
        'funcion que ejecuta el procedure SMW_SP_CALCULAPR530 que crea un 
        'archivo con las horas laboradas para el pago de la quincena.

        Dim rows_affected As Integer
        '-------- no se enviara parametro por eso no se asigna valores a arreglos

        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        '------------------

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")



            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_pase_lawson", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - pf_CrearArchivoPago, ocurrió lo siguiente: " & " pf_CrearArchivoPago" & ex.ToString)

        End Try
        Return rows_affected

    End Function

    Public Function Ejecuta_CierreMercado(ByVal p_empresa As Integer)
        'funcion que ejecuta en el paquete smw_pkg_calculos, el procedure sp_cierre_mercado

        Dim rows_affected As Integer



        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        arreglo_paramname.Add("p_mercado")
        arreglo_paramvalue(1) = p_empresa

        '------------------

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'este procedimiento esta dentro del paquete de calculos (contenedor de los procedure de calculos internos de Marca Web)

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_cierre_mercado", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Ejecuta_CierreMercado, ocurrió lo siguiente: " & p_empresa & ex.ToString)

        End Try
        Return rows_affected

    End Function

    Public Function Ejecuta_CierrePeriodo()
        'funcion que ejecuta en el paquete smw_pkg_calculos, el procedure sp_smw_calculahoras3 que 
        'marca con estatus de procesados los registros de la tabla smw_horascalculadas y luego modifica 
        'el periodo activo de pago por el siguiente periodo que entra en vigencia.


        Dim rows_affected As Integer

        '-------- no se enviara parametro por eso no se asigna valores a arreglos
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        '------------------

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'este procedimiento esta dentro del paquete de calculos (contenedor de los procedure de calculos internos de Marca Web)

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_cierre_periodo", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Ejecuta_CierrePeriodo, ocurrió lo siguiente: " & " pf_CrearArchivoPago" & ex.ToString)

        End Try
        Return rows_affected

    End Function

    Public Function Update_usuario(ByVal value_id_usuario As Long, ByVal value_nombre As String, ByVal value_id_sitio As Long, ByVal value_status As String, ByVal value_id_roles As Long, ByVal value_id_reloj As Long) As Integer
        Dim actdatos As Integer
        Dim array_paramvalue(6)
        Dim array_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            array_paramname.Add("p_id_usuario")
            array_paramname.Add("p_nombre")
            array_paramname.Add("p_id_sitio")
            array_paramname.Add("p_status")
            array_paramname.Add("p_id_roles")
            array_paramname.Add("p_id_reloj")

            array_paramvalue(1) = value_id_usuario
            array_paramvalue(2) = value_nombre
            array_paramvalue(3) = value_id_sitio
            array_paramvalue(4) = value_status
            array_paramvalue(5) = value_id_roles
            array_paramvalue(6) = value_id_reloj

            actdatos = clshelper.Ejecutar("smw_pkg_helper.sp_actualizaUsuarios", array_paramname, array_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Update_usuario, ocurrió lo siguiente: " & ex.ToString)

        End Try

        Return actdatos

    End Function

    Public Function Update_HExtras(ByVal val_empleado As Integer, _
                                   ByVal val_fecha_marc As Date, _
                                   ByVal val_hor_salida As String, _
                                   ByVal val_h_salida As String, _
                                   ByVal val_autoriza_hextras As String, _
                                   ByVal val_h_excedentes As Double, _
                                   ByVal val_tipo_dia As String, _
                                   ByVal val_turno As String, _
                                   ByVal val_usuario As String)

        Dim rows_affected As Integer
        Dim array_paramvalue(9)
        Dim array_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            array_paramname.Add("p_empleado")
            array_paramname.Add("p_fecha_marcacion")
            array_paramname.Add("p_hor_salida")
            array_paramname.Add("p_h_salida")
            array_paramname.Add("p_autoriza_horasextras")
            array_paramname.Add("p_horas_excedentes")
            array_paramname.Add("p_tipo_dia")
            array_paramname.Add("p_turno")
            array_paramname.Add("p_usuario")

            array_paramvalue(1) = val_empleado
            array_paramvalue(2) = val_fecha_marc
            array_paramvalue(3) = val_hor_salida
            array_paramvalue(4) = val_h_salida
            array_paramvalue(5) = val_autoriza_hextras
            array_paramvalue(6) = val_h_excedentes
            array_paramvalue(7) = val_tipo_dia
            array_paramvalue(8) = val_turno
            array_paramvalue(9) = val_usuario

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_updins_rechextra", array_paramname, array_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Update_HExtras, ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return rows_affected
    End Function

    Public Function update_faltantes_sarweb(ByVal val_mercado As String, ByVal val_documento_id As String, _
                                           ByVal val_jobcode As String, ByVal val_job_description As String, ByVal val_empleado_nombre As String, _
                                           ByVal val_fecha_faltante As Date, ByVal val_monto As Double, ByVal val_descripcion As String, _
                                           ByVal val_tipo_faltante As Integer, ByVal val_empleado_preparo As Integer, ByVal val_empleado_aprobo As Integer, _
                                           ByVal val_usuario As String, ByVal val_aplicadogl As String, _
                                           ByVal val_aplicadopr As String, ByVal val_cod_des As String)

        Dim rows_affected As Integer
        Dim arreglo_paramvalue(15) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_mercado")
            arreglo_paramname.Add("p_documento_id")
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
            arreglo_paramvalue(3) = val_jobcode
            arreglo_paramvalue(4) = val_job_description
            arreglo_paramvalue(5) = val_empleado_nombre
            arreglo_paramvalue(6) = val_fecha_faltante
            arreglo_paramvalue(7) = val_monto
            arreglo_paramvalue(8) = val_descripcion
            arreglo_paramvalue(9) = val_tipo_faltante
            arreglo_paramvalue(10) = val_empleado_preparo
            arreglo_paramvalue(11) = val_empleado_aprobo
            arreglo_paramvalue(12) = val_usuario
            arreglo_paramvalue(13) = val_aplicadogl
            arreglo_paramvalue(14) = val_aplicadopr
            arreglo_paramvalue(15) = val_cod_des

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_UpdateFaltantesSw", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion update_faltantes_sarweb ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function


    'agregado por echavez 29/07/2011
    Public Function Update_PeriodoDecimo(ByVal p_id As Integer, ByVal p_check_date As String, ByVal p_start_date As String, _
                                ByVal p_end_date As String, ByVal p_status As Integer, ByVal p_periodo As Integer)
        Dim rows_affected As Integer
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_id")
            arreglo_paramname.Add("p_check_date")
            arreglo_paramname.Add("p_start_date")
            arreglo_paramname.Add("p_end_date")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_periodo")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_id
            arreglo_paramvalue(2) = p_check_date
            arreglo_paramvalue(3) = p_start_date
            arreglo_paramvalue(4) = p_end_date
            arreglo_paramvalue(5) = p_status
            arreglo_paramvalue(6) = p_periodo

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_UpdatePeriodoDecimo", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion sp_UpdatePeriodoDecimo ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function Update_ISR(ByVal p_id As Integer, ByVal p_end_date As String, ByVal p_status As Integer, _
                               ByVal p_type As String)
        Dim rows_affected As Integer
        Dim arreglo_paramvalue(4) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_id")
            arreglo_paramname.Add("p_end_date")
            arreglo_paramname.Add("p_status")
            arreglo_paramname.Add("p_type")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_id
            arreglo_paramvalue(2) = p_end_date
            arreglo_paramvalue(3) = p_status
            arreglo_paramvalue(4) = p_type

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_UpdateISR", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion sp_UpdateISR ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function Actualiza_Vacaciones(ByVal p_employee As Integer, ByVal p_seq As Integer, _
                                         ByVal p_status As Integer)
        Dim rows_affected As Integer
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_employee")
            arreglo_paramname.Add("p_seq")
            arreglo_paramname.Add("p_status")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee
            arreglo_paramvalue(2) = p_seq
            arreglo_paramvalue(3) = p_status


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.SP_ACTUALIZA_VACACIONES", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion ACTUALIZA_VACACIONES ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function





#Region "ORBARRIA"
    Public Function fun_up_csscorr() As Integer
        Dim row As New Integer
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            row = clshelper.Ejecutar("smw_pkg_helper.SP_UP_CSSCORR", arreglo_paramname, arreglo_paramvalue)
        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos fun_up_csscorr  : " & ex.ToString)
        End Try
        Return row
    End Function
#End Region

End Class
