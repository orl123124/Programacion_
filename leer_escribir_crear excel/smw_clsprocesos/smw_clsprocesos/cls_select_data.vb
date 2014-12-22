Imports Helper


Public Class cls_select_data
    Dim clsinsertdata As New cls_insert_data
    Dim clshelper As New Helper.ClsHelperOra
    'Función que permite obtener un valor númerico
    Public Function GetValue(ByVal name_result As String, ByVal value_campo As String, ByVal value_tabla As String, ByVal value_tipo As Integer, ByVal value_parametro As String, ByVal value_valor As String) As Integer

        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_campo")
            arreglo_paramname.Add("p_tabla")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_parametro")
            arreglo_paramname.Add("p_valor")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_campo
            arreglo_paramvalue(2) = value_tabla
            arreglo_paramvalue(3) = value_tipo
            arreglo_paramvalue(4) = value_parametro
            arreglo_paramvalue(5) = value_valor

            GetValue = clshelper.TraerValor(name_result, "smw_pkg_helper.sp_getvalue", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion GetValue ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return GetValue

    End Function

    'Funcion que permite obtener un grupo de datos que serán utilizados para poblar los ComboBox

    Public Function llena_combos(ByVal value_campos As String, ByVal value_tablas As String, ByVal value_tipo As Integer, ByVal value_order As String, ByVal value_tpo_dia As String) As DataSet
        Dim ds_combos As New DataSet
        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_campos")
            arreglo_paramname.Add("p_tablas")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_order")
            arreglo_paramname.Add("p_tpo_dia")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_campos
            arreglo_paramvalue(2) = value_tablas
            arreglo_paramvalue(3) = value_tipo
            arreglo_paramvalue(4) = value_order
            arreglo_paramvalue(5) = value_tpo_dia


            ds_combos = clshelper.TraerDataset("smw_pkg_helper.sp_cargar_datos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Llena_Combos ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try

        Return ds_combos

    End Function

    'Función que permite obtener datos de la Tabla SMW_HORAS_JORNADA
    Public Function Obtener_HorasJornada(ByVal value_codigo As String) As DataSet

        Dim ds_hjornada As New DataSet
        Dim arreglo_paramname As New ArrayList
        Dim arreglo_paramvalue(1) As Object
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_codigo")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_codigo

            ds_hjornada = clshelper.TraerDataset("smw_pkg_helper.sp_SeleccionaHJornada", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Obtener_HorasJornada ocurrió lo siguiente: " & ex.ToString)

        End Try

        Return ds_hjornada

    End Function

    'Función que permite obtener datos de la Tabla SMW_HORAS_JORNADA
    Public Function BusquedaJornadas(ByVal var_estado As String) As DataSet

        Dim ds_hjornada As New DataSet
        Dim arreglo_paramname As New ArrayList
        Dim arreglo_paramvalue(1) As Object
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_estado")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = var_estado

            ds_hjornada = clshelper.TraerDataset("smw_pkg_helper.sp_BuscaJornada", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion BusquedaJornadas ocurrió lo siguiente: " & ex.ToString)

        End Try

        Return ds_hjornada

    End Function

    'Funcion que permite obtener datos de la Tabla Employee
    Public Function Selecciona_HCalculadas(ByVal value_empleado As Integer, ByVal value_fechaini As String, ByVal value_fechafin As String) As DataSet
        ', ByVal tipo As Integer
        'Dim arreglo_paramvalue(tipo) As Object

        Dim ds As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fechaini")
            'If tipo = 3 Then
            arreglo_paramname.Add("p_fechafin")
            'End If


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fechaini
            arreglo_paramvalue(3) = value_fechafin

            ds = clshelper.TraerDataset("smw_pkg_helper.sp_seleccion_HorasCaluladas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Selecciona_HCalculadas ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return ds
    End Function

    'Funcion que permite obtener datos de la Tabla Employee
    Public Function BusquedaEmployee(ByVal value_empleado As Integer) As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empleado


            ds = clshelper.TraerDataset("smw_pkg_helper.sp_BuscaEmployee", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion BusquedaEmployee ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return ds
    End Function

    'Funcion que permite obtener datos de la Tabla smw_horasperiodos
    Public Function BusquedaPeriodos(ByVal value_periodo As Integer) As DataSet
        Dim ds As New DataSet
        ds.Clear()
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_periodo")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_periodo


            ds = clshelper.TraerDataset("smw_pkg_helper.sp_BuscaPeriodos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion BusquedaPeriodos ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return ds
    End Function

    'Funcion que permite obtener datos de aquellos empleados que sean supervisores de la Tabla Employee 
    Public Function selecciona_supervisor(ByVal value_mercado As Integer, ByVal value_empleado As Integer) As DataSet
        Dim listar_supervisor As New DataSet

        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_mercado")
            arreglo_paramname.Add("p_employee")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_mercado
            arreglo_paramvalue(2) = value_empleado


            listar_supervisor = clshelper.TraerDataset("smw_pkg_helper.sp_selecciona_supervisor", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_supervisor ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try

        Return listar_supervisor

    End Function

    'Funcion que permite obtener datos de las pantallas admitidas por cada Rol
    Public Function selecciona_pantallas(ByVal value_roles As Integer) As DataSet
        Dim listar_pantallas As New DataSet

        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_roles")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_roles


            listar_pantallas = clshelper.TraerDataset("smw_pkg_helper.sp_SeleccionaPantallas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_supervisor ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try

        Return listar_pantallas

    End Function

    'Funcion que permite selecciona el valor maximo(campo numerico) en horas permitido para horas extras y horas de enfermedad.
    Public Function selecciona_maxhoras(ByVal value_campos As String, ByVal value_tablas As String, ByVal value_tipo As Integer, ByVal value_order As String, ByVal value_tpo_dia As String) As DataSet
        Dim valor_maxhoras As New DataSet
        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_campos")
            arreglo_paramname.Add("p_tablas")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_order")
            arreglo_paramname.Add("p_tpo_dia")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_campos
            arreglo_paramvalue(2) = value_tablas
            arreglo_paramvalue(3) = value_tipo
            arreglo_paramvalue(4) = value_order
            arreglo_paramvalue(5) = value_tpo_dia
            valor_maxhoras = clshelper.TraerDataset("smw_pkg_helper.sp_cargar_datos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_hextra ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try

        Return valor_maxhoras

    End Function

    'Funcion que permite selecciona el valor maximo en horas de enfermedad permitido para cada colaborador.
    Public Function selecciona_henfermedad(ByVal value_empleado As Integer, ByVal value_fecha_ini As String, ByVal value_tipo As Integer) As DataSet
        Dim valor_henfer As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_tipo")
            'arreglo_paramname.Add("p_mercado")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fecha_ini
            arreglo_paramvalue(3) = value_tipo
            'arreglo_paramvalue(3) = value_mercado

            valor_henfer = clshelper.TraerDataset("smw_pkg_helper.sp_selecciona_henfermedad", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_hextra ocurrió lo siguiente: " & ex.ToString)

        Finally

        End Try

        Return valor_henfer

    End Function

    'Funcion que obtiene el turno creado al empleado, segun la semana

    'Public Function sp_SeleccionaTurnos(ByVal VALUE_NEMPLEADO As Long, ByVal VALUE_NSEMANA As Long, ByVal VALUE_NPERIODO As Long) As DataSet
    Public Function sp_SeleccionaTurnos(ByVal value_empleado As Long, ByVal value_fecha As Date) As DataSet



        Dim lista_turnos As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            'arreglo_paramname.Add("p_EMPLEADO")
            'arreglo_paramname.Add("p_SEMANA")
            'arreglo_paramname.Add("p_PERIODO")


            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            'arreglo_paramvalue(1) = VALUE_NEMPLEADO
            'arreglo_paramvalue(2) = VALUE_NSEMANA
            'arreglo_paramvalue(3) = VALUE_NPERIODO

            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fecha


            lista_turnos = clshelper.TraerDataset("smw_pkg_helper.sp_SeleccionaTurnos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion SeleccionaTurnos ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_turnos

    End Function

    Public Function num_semana(ByVal value_fecha As Date) As String

        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList
        Dim semana, periodo As String

        Try

            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist

            arreglo_paramname.Add("p_fecha")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_fecha

            semana = clshelper.TraerValor("p_semana", "smw_pkg_helper.sp_determina_semana_periodo", arreglo_paramname, arreglo_paramvalue)
            periodo = clshelper.TraerValor("p_periodo", "smw_pkg_helper.sp_determina_semana_periodo", arreglo_paramname, arreglo_paramvalue)

            num_semana = semana & "," & periodo

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion SeleccionaTurnos ocurrió lo siguiente: " & ex.ToString)
            semana = Nothing
            periodo = Nothing
            num_semana = Nothing
        End Try

        Return num_semana

    End Function


    Public Function valida_user(ByVal value_idempleado As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_usuario")
            'arreglo_paramname.Add("p_pantalla")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_idempleado
            'arreglo_paramvalue(2) = value_pantalla

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_Valida_Pantalla", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion valida_user  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function get_datos_user(ByVal value_idempleado As String) As DataSet
        Dim lista_datos As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_usuario")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_idempleado

            lista_datos = clshelper.TraerDataset("smw_pkg_helper.sp_Usuario_Mercado", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion sp_Usuario_Mercado  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_datos

    End Function

    'Funcion que permite obtener datos de la Tabla Marcaciones
    Public Function Obtener_Marcaciones(ByVal value_empresa As String, _
                                        ByVal value_empleado As String, _
                                        ByVal value_fecha As String) As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empresa")
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empresa
            arreglo_paramvalue(2) = value_empleado
            arreglo_paramvalue(3) = value_fecha

            ds = clshelper.TraerDataset("smw_pkg_helper.sp_obtener_marcaciones", arreglo_paramname, arreglo_paramvalue)
            'smw_pkg_helper.sp_obtener_marcaciones(p_empresa => :p_empresa,
            '                            p_empleado => :p_empleado,
            '                            p_fecha => :p_fecha,
            '                            p_cursor => :p_cursor);

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion BusquedaEmployee ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return ds
    End Function

    'Funcion que obtiene el turno creado al empleado, segun la semana

    'Funcion que permite obtener las horas extras de la Tabla Horas Calculadas
    Public Function Selecciona_HExtras(ByVal value_empleado As Integer, ByVal value_fechaini As String, ByVal value_fechafin As String) As DataSet
        ', ByVal tipo As Integer
        'Dim arreglo_paramvalue(tipo) As Object

        Dim ds As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fechaini")
            arreglo_paramname.Add("p_fechafin")
            'End If


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fechaini
            arreglo_paramvalue(3) = value_fechafin

            ds = clshelper.TraerDataset("smw_pkg_helper.sp_Selecciona_Hextras", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Selecciona_HExtras ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return ds
    End Function

    Public Function selecciona_usuario(ByVal value_id_usuario As Integer) As DataSet
        Dim seleccion As New DataSet
        Dim array_paramvalue(1) As Object
        Dim array_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            array_paramvalue(1) = value_id_usuario
            array_paramname.Add("p_id_usuario")

            seleccion = clshelper.TraerDataset("smw_pkg_helper.sp_Selectusuario", array_paramname, array_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_usuario ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return seleccion
    End Function

    Public Function selecciona_codigo(ByVal value_campos As String, ByVal value_tablas As String, ByVal value_tipo As Integer, ByVal value_order As String, ByVal value_tpo_dia As String) As DataSet

        Dim codsel As New DataSet
        Dim arreglo_paramvalue(5) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_campos")
            arreglo_paramname.Add("p_tablas")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_order")
            arreglo_paramname.Add("p_tpo_dia")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            arreglo_paramvalue(1) = value_campos
            arreglo_paramvalue(2) = value_tablas
            arreglo_paramvalue(3) = value_tipo
            arreglo_paramvalue(4) = value_order
            arreglo_paramvalue(5) = value_tpo_dia

            codsel = clshelper.TraerDataset("smw_pkg_helper.sp_cargar_datos", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion selecciona_hextra ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return codsel
    End Function

    Public Function seleccion_recalchextra(ByVal empleado As Integer, ByVal fecha As String) As DataSet
        Dim salida As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fecha")

            arreglo_paramvalue(1) = empleado
            arreglo_paramvalue(2) = fecha

            salida = clshelper.TraerDataset("smw_pkg_helper.sp_seleccion_recalchextra", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion seleccion_recalchextra ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return salida
    End Function

    Public Function seleccion_recalculahextra(ByVal value_empleado As Integer, ByVal value_fechaini As String, ByVal value_fechafin As String) As DataSet

        Dim ds As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_fechaini")
            arreglo_paramname.Add("p_fechafin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_empleado
            arreglo_paramvalue(2) = value_fechaini
            arreglo_paramvalue(3) = value_fechafin

            ds = clshelper.TraerDataset("smw_pkg_helper.sp_recalculahextra_MW", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion seleccion_recalculahextra ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return ds
    End Function

    Public Function Get_Menu(ByVal p_usuario As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_usuario")
            'arreglo_paramname.Add("p_pantalla")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_usuario
            'arreglo_paramvalue(2) = value_pantalla

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_Menu", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Get_Menu  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Get_Sub_Menu(ByVal p_usuario As String, ByVal p_parentnode As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_usuario")
            arreglo_paramname.Add("p_parentnode")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_usuario
            arreglo_paramvalue(2) = p_parentnode

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_Sub_Menu", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Get_Sub_Menu  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Lista_Razon(ByVal p_tipo As Integer) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_tipo")
            'arreglo_paramname.Add("p_parentnode")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_tipo
            'arreglo_paramvalue(2) = p_parentnode

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_Lista_Razon", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Lista_Razon  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function GeneraReporte03(ByVal p_razon_social As String, ByVal p_fecha_ini As String, _
                                    ByVal p_fecha_fin As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_razon_social")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_razon_social
            arreglo_paramvalue(2) = p_fecha_ini
            arreglo_paramvalue(3) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_XR215", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion GeneraReporte03  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Lista_Pagos(ByVal p_pago As Integer) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_pago")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_pago


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_Lista_Pagos", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Lista_Pagos  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Calcula_Decimo(ByVal p_fecha_pago As String, _
                                 ByVal p_fecha_ini As String, _
                                 ByVal p_fecha_fin As String) As DataSet
        ''ByVal p_payclass As String, _
        ''ByVal p_cierre As Integer

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS")

            ''arreglo_paramname.Add("p_payclass")
            arreglo_paramname.Add("p_fecha_pago")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")
            ''arreglo_paramname.Add("p_cierre")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            '' arreglo_paramvalue(1) = p_payclass
            arreglo_paramvalue(1) = p_fecha_pago
            arreglo_paramvalue(2) = p_fecha_ini
            arreglo_paramvalue(3) = p_fecha_fin
            '' arreglo_paramvalue(4) = p_cierre

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_calcula_decimo", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Calcula_Decimo ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function


    Public Function Calcula_Vacaciones(ByVal p_usuario As String) As DataSet
        'ByVal p_empleados As String, _
        '                               ByVal p_fecha_ini As String, _
        '                               ByVal p_tipo As String
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS")

            'arreglo_paramname.Add("p_empleados")
            'arreglo_paramname.Add("p_fecha_ini")
            'arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_usuario")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object

            'arreglo_paramvalue(1) = p_empleados
            'arreglo_paramvalue(2) = p_fecha_ini
            'arreglo_paramvalue(3) = p_tipo
            arreglo_paramvalue(1) = p_usuario


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_calcula_vacaciones", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Calcula_Vacaciones ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    'echavez modificacion para la el anexo 03 13/06/2011
    Public Function Anexo03(ByVal p_razon_social As String, ByVal p_fecha_ini As String, _
                                    ByVal p_fecha_fin As String) As DataSet
        Dim dsDatos As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_razon_social")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_razon_social
            arreglo_paramvalue(2) = p_fecha_ini
            arreglo_paramvalue(3) = p_fecha_fin

            dsDatos = clshelper.TraerDataset("smw_pkg_helper.SP_ANEXO_03", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Pre_Elaborada  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return dsDatos
    End Function

    Public Function Busca_Empleado(ByVal p_employee As Integer, ByVal p_fecha_ini As String, _
                                   ByVal p_tipo As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_employee")
            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_tipo")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee
            arreglo_paramvalue(2) = p_fecha_ini
            arreglo_paramvalue(3) = p_tipo

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_EMPLEADO", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Empleado  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    'modificacion echavez 29/07/2011
    Public Function Busca_PeriodoDecimos(ByVal p_periodo As Integer, ByVal p_status As Integer) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            ' clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4


            arreglo_paramname.Add("p_periodo")
            arreglo_paramname.Add("p_status")
            'arreglo_paramname.Add("p_fecha_fin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_periodo
            arreglo_paramvalue(2) = p_status
            'arreglo_paramvalue(3) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_PERIODOSDECIMOS", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Decimos  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Busca_Decimos() As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            ' clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4


            'arreglo_paramname.Add("p_fecha_pago")
            'arreglo_paramname.Add("p_fecha_ini")
            'arreglo_paramname.Add("p_fecha_fin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            'arreglo_paramvalue(1) = p_fecha_pago
            'arreglo_paramvalue(2) = p_fecha_ini
            'arreglo_paramvalue(3) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_DECIMOS", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Decimos  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Busca_Años() As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4


            'arreglo_paramname.Add("p_fecha_pago")
            'arreglo_paramname.Add("p_fecha_ini")
            'arreglo_paramname.Add("p_fecha_fin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            'arreglo_paramvalue(1) = p_fecha_pago
            'arreglo_paramvalue(2) = p_fecha_ini
            'arreglo_paramvalue(3) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_ANOS", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Años  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Reporte_XR171(ByVal p_fecha_ini As String, _
                                 ByVal p_fecha_fin As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_fecha_ini
            arreglo_paramvalue(2) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_XR171", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Reporte XR171 ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Reporte_XR140(ByVal p_usuario As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_usuario")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_usuario


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_XR140", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Reporte XR140 ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Historial_vacaciones(ByVal p_employee As Integer) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_employee")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_historial_vacaciones", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion  Historial Vacaciones ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function


    '---buscar los porcentaje del impuesto sobre la renta
    '-- echavez 04/08/2011
    Public Function Busca_ISR(ByVal p_tipo As String, ByVal p_status As Integer) As DataSet

        Dim dsLista As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            ' clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4


            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_status")



            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_tipo
            arreglo_paramvalue(2) = p_status


            dsLista = clshelper.TraerDataset("smw_pkg_helper.sp_buscar_isr", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca ISR  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return dsLista

    End Function

    'Public Function Calcula_Liquidacion(ByVal p_employee As Integer, _
    '                           ByVal p_tipo As String, _
    '                           ByVal p_fecha_liq As String, _
    '                           ByVal p_ind_especial As String, _
    '                           ByVal p_contrato As String, _
    '                           ByVal p_usuario As String) As DataSet


    '    Dim lista_pantalla As New DataSet
    '    Dim arreglo_paramvalue(6) As Object
    '    Dim arreglo_paramname As New ArrayList

    '    Try
    '        clshelper.inicia("oracle", "LAWSON\\SMW")
    '        'clshelper.inicia("oracle", "LAWSON\\LWS")

    '        arreglo_paramname.Add("p_employee")
    '        arreglo_paramname.Add("p_tipo")
    '        arreglo_paramname.Add("p_fecha_liq")
    '        arreglo_paramname.Add("p_ind_especial")
    '        arreglo_paramname.Add("p_contrato")
    '        arreglo_paramname.Add("p_usuario")


    '        'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
    '        arreglo_paramvalue(1) = p_employee
    '        arreglo_paramvalue(2) = p_tipo
    '        arreglo_paramvalue(3) = p_fecha_liq
    '        arreglo_paramvalue(4) = p_ind_especial
    '        arreglo_paramvalue(5) = p_contrato
    '        arreglo_paramvalue(6) = p_usuario

    '        lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_calcula_liquidacion", arreglo_paramname, arreglo_paramvalue)

    '    Catch ex As Exception
    '        clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Calcula_Liquidacion ocurrió lo siguiente: " & ex.ToString)
    '    End Try

    '    Return lista_pantalla

    'End Function


    'Public Function Busca_Liquidacion(ByVal p_employee As Integer) As DataSet
    '    Dim lista_pantalla As New DataSet
    '    Dim arreglo_paramvalue(1) As Object
    '    Dim arreglo_paramname As New ArrayList


    '    Try
    '        clshelper.inicia("oracle", "LAWSON\\SMW")
    '        'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

    '        'Adiciona los nombres de los parametros del Procedure en un arraylist
    '        arreglo_paramname.Add("p_employee")


    '        'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
    '        arreglo_paramvalue(1) = p_employee


    '        lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_LIQUIDACION", arreglo_paramname, arreglo_paramvalue)


    '    Catch ex As Exception
    '        clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Liquidacion  ocurrió lo siguiente: " & ex.ToString)
    '    End Try

    '    Return lista_pantalla

    'End Function


    Public Function verificar_reloj(ByVal p_reloj As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_reloj")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_reloj


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_verificar_reloj", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos -Seccion que verfica los relojes  ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return lista_pantalla
    End Function

    Public Function generar_xr212(ByVal anio As Integer, _
                                  ByVal mes As Integer, _
                                  ByVal id_razon_social As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_Periodo")
            arreglo_paramname.Add("p_Mes")
            arreglo_paramname.Add("p_Id_Razon_Social")




            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = anio
            arreglo_paramvalue(2) = mes
            arreglo_paramvalue(3) = id_razon_social


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_XR212", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos -Seccion que verfica los relojes  ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return lista_pantalla
    End Function
    ''reporte xr2212

    Public Function reporte_xr212(ByVal anio As Integer, _
                                  ByVal mes As Integer, _
                                  ByVal id_razon_social As String, _
                                  ByVal cia_planilla As Integer) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(4) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_Periodo")
            arreglo_paramname.Add("p_Mes")
            arreglo_paramname.Add("p_Id_Razon_Social")
            arreglo_paramname.Add("p_Cia_planilla")




            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = anio
            arreglo_paramvalue(2) = mes
            arreglo_paramvalue(3) = id_razon_social
            arreglo_paramvalue(4) = cia_planilla


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.reporte_XR212", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos -Seccion que verfica los relojes  ocurrió lo siguiente: " & ex.ToString)
        End Try
        Return lista_pantalla
    End Function

    Public Function Calcula_Liquidacion(ByVal p_employee As Integer, _
                               ByVal p_tipo As String, _
                               ByVal p_fecha_liq As String, _
                               ByVal p_ind_especial As String, _
                               ByVal p_contrato As String, _
                               ByVal p_usuario As String) As DataSet


        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(6) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS")

            arreglo_paramname.Add("p_employee")
            arreglo_paramname.Add("p_tipo")
            arreglo_paramname.Add("p_fecha_liq")
            arreglo_paramname.Add("p_ind_especial")
            arreglo_paramname.Add("p_contrato")
            arreglo_paramname.Add("p_usuario")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee
            arreglo_paramvalue(2) = p_tipo
            arreglo_paramvalue(3) = p_fecha_liq
            arreglo_paramvalue(4) = p_ind_especial
            arreglo_paramvalue(5) = p_contrato
            arreglo_paramvalue(6) = p_usuario

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_calcula_liquidacion", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Calcula_Liquidacion ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Busca_Liquidacion(ByVal p_employee As Integer) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_employee")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_BUSCA_LIQUIDACION", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Busca Liquidacion  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Genera_Liquidacion(ByVal p_employee As String) As DataSet
        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_employee")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_employee


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_GENERA_LIQUIDACION", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Genera Liquidacion  ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Reporte_XR141(ByVal p_usuario As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_usuario")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_usuario


            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_XR141", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Reporte XR141 ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Reporte_XR213(ByVal p_fecha_ini As String, _
                              ByVal p_fecha_fin As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_fecha_ini
            arreglo_paramvalue(2) = p_fecha_fin

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_CALCULA_CESANTIA", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Reporte XR213 ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Carta_trabajo(ByVal p_colaborador As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_colaborador")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_colaborador

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_SELECT_CARTA_TRABAJO", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Carta Trabajo ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Certificacion(ByVal p_razon As String, _
                                  ByVal p_fechaini As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_razon")
            arreglo_paramname.Add("p_fechaini")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_razon
            arreglo_paramvalue(2) = p_fechaini

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_CERTIFICACION_PLANILLA", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Carta Certificacion de Salarios ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function Reporte_Dec_Detalle(ByVal p_fecha_ini As String, _
                                  ByVal p_fecha_fin As String, _
                                  ByVal p_process_level As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(3) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_fecha_ini")
            arreglo_paramname.Add("p_fecha_fin")
            arreglo_paramname.Add("p_process_level")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_fecha_ini
            arreglo_paramvalue(2) = p_fecha_fin
            arreglo_paramvalue(3) = p_process_level

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.SP_REPORTE_DEC_DETALLE", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Reporte Dec Detalle ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function

    Public Function SubReporte_XR141() As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            ds = clshelper.TraerDataset("smw_pkg_planilla.sp_registros_tiempo", arreglo_paramname, arreglo_paramvalue)
        Catch ex As Exception

            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - SubReporte_XR141 ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return ds
    End Function

    Public Function Selecciona_Saldo(ByVal value_empleado As String) As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            Me.clshelper.Inicia("oracle", "LAWSON\\SMW", "", "")
            arreglo_paramname.Add("p_colaborador")
            arreglo_paramvalue(1) = value_empleado
            ds = Me.clshelper.TraerDataset("smw_pkg_planilla.sp_Select_CartaSaldo", arreglo_paramname, arreglo_paramvalue)
        Catch exception1 As Exception

            Me.clsinsertdata.CreaLogFile_clsprocesos(("En la clase clsprocesos - Seccion Carta de Saldo ocurrió lo siguiente: " & exception1.ToString))

        End Try
        Return ds
    End Function
    Public Function Selecciona_Certificacion(ByVal value_empleado As String) As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            Me.clshelper.Inicia("oracle", "LAWSON\\SMW", "", "")
            arreglo_paramname.Add("p_colaborador")
            arreglo_paramvalue(1) = value_empleado
            ds = Me.clshelper.TraerDataset("smw_pkg_planilla.sp_Select_CartaCert", arreglo_paramname, arreglo_paramvalue)
        Catch exception1 As Exception

            Me.clsinsertdata.CreaLogFile_clsprocesos(("En la clase clsprocesos - Seccion Carta de Certificacion ocurrió lo siguiente: " & exception1.ToString))

        End Try
        Return ds
    End Function

    Public Function Desglose_Pago(ByVal p_empleado As String, _
                                 ByVal p_MES As String) As DataSet

        Dim lista_pantalla As New DataSet
        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            'clshelper.inicia("oracle", "LAWSON\\LWS") 'se coloco otro registro para accesar lawson4

            arreglo_paramname.Add("p_empleado")
            arreglo_paramname.Add("p_MES")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_empleado
            arreglo_paramvalue(2) = p_MES

            lista_pantalla = clshelper.TraerDataset("smw_pkg_helper.sp_desgloses_pagos", arreglo_paramname, arreglo_paramvalue)


        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion Carta Certificacion de Salarios ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return lista_pantalla

    End Function
    
#Region "ORBARRIA"
    Public Function fun_get_csscorr() As DataSet
        Dim ds As New DataSet
        Dim arreglo_paramvalue(0) As Object
        Dim arreglo_paramname As New ArrayList
        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")
            ds = clshelper.TraerDataset("smw_pkg_helper.SP_GET_CSSCORR", arreglo_paramname, arreglo_paramvalue)
        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - CLS_SELECT_DATA -  fun_get_csscorr  : " & ex.ToString)
        End Try
        Return ds
    End Function


#End Region
End Class
