Imports Helper

Public Class cls_delete_data
    Dim clsinsertdata As New cls_insert_data
    Dim clshelper As New Helper.ClsHelperOra

    Public Function Delete_RolesPantallas(ByVal value_idrol As Integer) As Integer

        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_id_rol")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = value_idrol

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_DeleteRolesPantallas", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            'lblError.Text = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            ' objproc.CreaLogFile("En la Pantalla de Modificacion de Marcaciones - Seccion Actualiza Datos, ocurrió lo siguiente: " & ex.ToString)
            '            mensaje1 = "Ha ocurrido un error en la sección Actualiza Datos. Por favor llamar al Administrador del Sistema."
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion DELETE_ROLESPANTALLAS, ocurrió lo siguiente: " & ex.ToString)

        End Try
        Return rows_affected

    End Function

    Public Function delete_faltantes_sarweb(ByVal val_documentoid As String, ByVal val_mercado As String) As Integer

        Dim arreglo_paramvalue(2) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer

        Try
            clshelper.Inicia("oracle", "LAWSON\\SMW")

            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_documentoid")
            arreglo_paramname.Add("p_mercado")

            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = val_documentoid
            arreglo_paramvalue(2) = val_mercado

            rows_affected = clshelper.Ejecutar("smw_pkg_helper.sp_DeleteFaltantesSw", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion delete_faltantes_sarweb, ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

    Public Function delete_temp(ByVal p_usuario As String) As Integer

        Dim arreglo_paramvalue(1) As Object
        Dim arreglo_paramname As New ArrayList
        Dim rows_affected As Integer


        Try
            'clshelper.inicia("oracle", "LAWSON\\LWS")
            clshelper.Inicia("oracle", "LAWSON\\SMW")


            'Adiciona los nombres de los parametros del Procedure en un arraylist
            arreglo_paramname.Add("p_usuario")


            'Adiciona los valores de los parametros del Procedure en un arreglo de tipo object
            arreglo_paramvalue(1) = p_usuario


            rows_affected = clshelper.Ejecutar("smw_pkg_helper.SP_DELETE_TEMP", arreglo_paramname, arreglo_paramvalue)

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("En la clase clsprocesos - Seccion delete_temp, ocurrió lo siguiente: " & ex.ToString)
        End Try

        Return rows_affected

    End Function

End Class
