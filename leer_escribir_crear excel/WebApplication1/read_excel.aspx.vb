Imports clsprocesos
Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports OfficeOpenXml
Imports System.Reflection
Imports OfficeOpenXml.Style
Imports System.Drawing

Partial Public Class read_excel
    Inherits System.Web.UI.Page

    Dim condic_ As Boolean
    Dim ruta As String
    Dim error_ As Integer = 0
    Dim fila As Integer = 1
    Dim array(,) As String
    Dim list_ As List(Of Object)
    Dim ds_corr As DataSet

    Dim obj_update As cls_update_data
    Dim obj_insert As cls_insert_data
    Dim obj_select As cls_select_data

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Lbl_resp.Text = ""
        btn_bajar.Visible = False
    End Sub

    Protected Sub btn_ejecutar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_ejecutar.Click
        Dim row_m As Integer = 0
        Dim row_update As Integer = 0
        Try


            ruta = FileUpload1.PostedFile.FileName
            Dim extn_ = Path.GetExtension(FileUpload1.PostedFile.FileName.ToString).ToUpper

            If extn_.Equals(".XLS") Or extn_.Equals(".XLSX") Then
                condic_ = True

            Else
                condic_ = False

            End If



            If condic_ Then
                read_file(ruta)

                obj_update = New cls_update_data
                obj_insert = New cls_insert_data
                'ds_corr = obj_select.fun_csscorrelativo_emp
                row_update = obj_update.fun_up_csscorr()
                row_m = obj_insert.inst_to_mastercorrelativo
                If row_m Then
                    'Lbl_resp.Text = "los registros se procesaron correctamente ..."
                    

                    If (Not ClientScript.IsStartupScriptRegistered("alert")) Then
                        Page.ClientScript.RegisterStartupScript _
                        (Me.GetType(), "alert", "alertMe(1);", True)
                    End If
                    btn_bajar.Visible = True
                Else
                    If (Not ClientScript.IsStartupScriptRegistered("alert")) Then
                        Page.ClientScript.RegisterStartupScript _
                        (Me.GetType(), "alert", "alertMe(0);", True)
                    End If
                    'Lbl_resp.Text = "error en el proceso...."
                End If
            End If

        Catch ex As Exception
            obj_insert = New cls_insert_data
            obj_insert.CreaLogFile("error mfrm_corelativo_css : funcion read_file" & ex.Message)
        End Try
    End Sub


    Public Function read_file(ByVal ruta As String) As Boolean
        Dim wbook As ExcelWorkbook
        Dim currentWorksheet As ExcelWorksheet
        Dim col As Hashtable
        Dim existingFile As System.IO.FileInfo
        Try
            If ruta IsNot Nothing Then
                existingFile = New FileInfo(ruta)
                Using package = New ExcelPackage(existingFile)
                    'get work book in the file
                    wbook = package.Workbook

                    If wbook IsNot Nothing Then
                        If wbook.Worksheets.Count > 0 Then
                            'get first sheet
                            currentWorksheet = wbook.Worksheets.First()
                            'read some data 
                            col = New Hashtable
                            col.Add("CEDULA", currentWorksheet.Cells(fila, 1).Value)
                            col.Add("SEGURO", currentWorksheet.Cells(fila, 2).Value)
                            col.Add("NOMBRE", currentWorksheet.Cells(fila, 3).Value)
                            col.Add("CONSECUTIVO", currentWorksheet.Cells(fila, 4).Value)
                            'CEDULA	SEGURO	NOMBRE	CONSECUTIVO

                            If Not col("CEDULA").Equals("CEDULA") Then
                                error_ = error_ + 1
                            End If
                            If Not col("SEGURO").Equals("SEGURO") Then
                                error_ = error_ + 1
                            End If
                            If Not col("NOMBRE").Equals("NOMBRE") Then
                                error_ = error_ + 1
                            End If
                            If Not col("CONSECUTIVO").Equals("CONSECUTIVO") Then
                                error_ = error_ + 1
                            End If
                            'currentWorksheet.Dimension.End.Row
                            Dim col_count As Integer

                            col_count = col.Count

                            If error_ < 1 Then
                                list_ = New List(Of Object)
                                list_.ToArray()
                                Dim count As Integer = currentWorksheet.Dimension.End.Row
                                For i As Integer = 1 + fila To count
                                    For y As Integer = 1 To col.Count
                                        list_.Add(currentWorksheet.Cells(i, y).Value.ToString)
                                        'list_.ToArray()
                                        'list_(i) = currentWorksheet.Cells(i, y).Value.ToString


                                    Next
                                    to_insert(list_)
                                    list_.Clear()
                                Next
                            End If
                        End If
                    End If


                End Using
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return 0
        End Try
        Return 1

    End Function


    Public Function to_insert(ByVal tabla_ As List(Of Object)) As Integer
        Dim row As Integer
        Dim h_tabla As New Hashtable
        Try

            obj_insert = New cls_insert_data
            obj_select = New cls_select_data

            ''CEDULA	SEGURO	NOMBRE	CONSECUTIVO

            h_tabla.Add("CEDULA", tabla_(0))
            h_tabla.Add("SEGURO", tabla_(1))
            h_tabla.Add("NOMBRE", tabla_(2))
            h_tabla.Add("CONSECUTIVO", tabla_(3))
            row = obj_insert.insert_to_csscorrelativo(h_tabla)


            h_tabla.Clear()
            'tabla_.Clear()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return row
    End Function


    ' create excel 
    Public Function export_to_excel(ByVal dt As DataTable) As ExcelPackage
        Dim pkg As ExcelPackage
        Dim wsheet As ExcelWorksheet

        Dim dt_cedula, dt_seguro, dt_nombre, dt_consecutivo, dt_employee, dt_fecha, dt_estatus As New DataTable
        Try


            pkg = New ExcelPackage
            wsheet = pkg.Workbook.Worksheets.Add("sheet_1")

            'dt_cedula = dt.Clone 
            dt_cedula.Columns.Add(New DataColumn("CEDULA", GetType(String)))
            dt_seguro.Columns.Add(New DataColumn("SEGURO", GetType(String)))
            dt_nombre.Columns.Add(New DataColumn("NOMBRE", GetType(String)))
            dt_consecutivo.Columns.Add(New DataColumn("CONSECUTIVO", GetType(String)))
            dt_employee.Columns.Add(New DataColumn("EMPLOYEE", GetType(String)))
            dt_fecha.Columns.Add(New DataColumn("FECHA", GetType(String)))
            dt_estatus.Columns.Add(New DataColumn("ESTATUS", GetType(String)))
            For x As Integer = 0 To dt.Rows.Count - 1
                'CEDULA	SEGURO	NOMBRE	CONSECUTIVO
                If dt.Rows(x).Item("CEDULA").ToString IsNot Nothing Then dt_cedula.Rows.Add(dt.Rows(x).Item("CEDULA"))
                If dt.Rows(x).Item("SEGURO").ToString IsNot Nothing Then dt_seguro.Rows.Add(dt.Rows(x).Item("SEGURO"))
                If dt.Rows(x).Item("NOMBRE").ToString IsNot Nothing Then dt_nombre.Rows.Add(dt.Rows(x).Item("NOMBRE"))
                If dt.Rows(x).Item("CONSECUTIVO").ToString IsNot Nothing Then dt_consecutivo.Rows.Add(dt.Rows(x).Item("CONSECUTIVO"))
                If dt.Rows(x).Item("EMPLOYEE_LAWSON").ToString IsNot Nothing Then dt_employee.Rows.Add(dt.Rows(x).Item("EMPLOYEE_LAWSON"))
                If dt.Rows(x).Item("FECHA").ToString IsNot Nothing Then dt_fecha.Rows.Add(dt.Rows(x).Item("FECHA"))
                If dt.Rows(x).Item("ESTATUS").ToString IsNot Nothing Then dt_estatus.Rows.Add(dt.Rows(x).Item("ESTATUS"))

            Next
            ' INSERT TO COLUMN FROM DT ROWS
            wsheet.Cells("A1").LoadFromDataTable(dt_cedula, True)
            wsheet.Cells("B1").LoadFromDataTable(dt_seguro, True)
            wsheet.Cells("C1").LoadFromDataTable(dt_nombre, True)
            wsheet.Cells("D1").LoadFromDataTable(dt_consecutivo, True)
            wsheet.Cells("E1").LoadFromDataTable(dt_employee, True)
            wsheet.Cells("F1").LoadFromDataTable(dt_fecha, True)
            wsheet.Cells("G1").LoadFromDataTable(dt_estatus, True)
            Dim count_column_format As Integer = dt.Columns.Count
            'SET FORMAT CENTER AT COLUMN
            For num = 1 To count_column_format 'wsheet.Cells.Count
                wsheet.Column(num).Width = 19
                wsheet.Column(num).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            Next

            Dim rango1 As ExcelRange = wsheet.Cells("A1:D1")
            rango1.Merge = False
            rango1.Style.Font.Bold = True
            rango1.Style.Fill.PatternType = ExcelFillStyle.Solid
            rango1.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
            rango1.Style.Font.Color.SetColor(Color.White)
            rango1.Style.Border.BorderAround(ExcelBorderStyle.None)

            Dim rango2 As ExcelRange = wsheet.Cells("E1:G1")
            rango2.Merge = False
            rango2.Style.Font.Bold = True
            rango2.Style.Fill.PatternType = ExcelFillStyle.Solid
            rango2.Style.Fill.BackgroundColor.SetColor(Color.DarkGray)
            rango2.Style.Font.Color.SetColor(Color.White)
            rango2.Style.Border.BorderAround(ExcelBorderStyle.None)
            'Write it back to the client
            Dim file_name_ As String = "correlativo_css_" & "_" & Today.Day & _
              "_" & Today.Month & "_" & Today.Year & "__ " & Today.Second.ToString

            'Response.Clear()
            'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            'Response.AddHeader("content-disposition", "attachment;  filename=" & file_name_ & ".xlsx")
            'Response.BinaryWrite(pkg.GetAsByteArray())
            'Response.End()

            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;  filename= " & file_name_ & ".xlsx")
            Dim stream As MemoryStream = New MemoryStream(pkg.GetAsByteArray())

            Response.OutputStream.Write(stream.ToArray(), 0, stream.ToArray().Length)

            Response.Flush()

            Response.Close()
          

        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try
        Return pkg
    End Function

    'import to excel


    'Public Sub export_to_excel(ByVal dt As System.Data.DataTable)

    '    Dim _excel As Application  'Excel.Application
    '    Dim wBook As Workbook
    '    Dim wSheet As Worksheet

    '    _excel = New Application
    '    wBook = _excel.Workbooks.Add()
    '    wSheet = wBook.ActiveSheet()

    '    Dim rg As Microsoft.Office.Interop.Excel.Range
    '    Dim rpt As Microsoft.Office.Interop.Excel.Application
    '    Dim WorkBook As Microsoft.Office.Interop.Excel.Workbook
    '    Dim Sheet As Microsoft.Office.Interop.Excel.Worksheet
    '    Dim misValue As Object = System.Reflection.Missing.Value

    '    Dim dc As System.Data.DataColumn
    '    Dim dr As System.Data.DataRow

    '    Dim col_Index As Integer = 0 'col
    '    Dim row_Index As Integer = 0 'fila

    '    Dim v1 As String

    '    Try
    '        'creacion de columnas
    '        Dim rows_ As String = dt.Rows.Count
    '        For Each col As System.Data.DataColumn In dt.Columns
    '            col_Index = col_Index + 1
    '            wSheet.Cells(1, col_Index) = col.ColumnName
    '            Dim nom As String = wSheet.Cells(1, col_Index).ToString

    '        Next
    '        'creacion de filas
    '        For Each dr In dt.Rows
    '            row_Index = row_Index + 1
    '            col_Index = 0
    '            For Each col2 As System.Data.DataColumn In dt.Columns
    '                col_Index = col_Index + 1
    '                wSheet.Cells(row_Index + 1, col_Index) = dr(col2.ColumnName)
    '            Next
    '        Next
    '        '--------------------------------------------------------------------------------------------
    '        obj_adm_control = New cls_admin_controls
    '        Dim path_files As String


    '        'C:\correlativo_css
    '        path_files = "C:\correlativo_css\file" & "_" & Today.Day & _
    '           "_" & Today.Month & "_" & Today.Year & ".xlsx".ToString
    '        '----
    '        wSheet.Columns.AutoFit()
    '        Dim strFileName As String = path_files '"C:\inetpub\wwwroot\smw\correlativo_css\datatable.xlsx" '"K:\datatable.xlsx"
    '        'If System.IO.File.Exists(strFileName) Then
    '        '    System.IO.File.Delete(strFileName)
    '        'End If
    '        wBook.SaveAs(strFileName)




    '        ReleaseObject(wSheet)
    '        wBook.Close(False)
    '        ReleaseObject(wBook)
    '        _excel.Quit()
    '        ReleaseObject(_excel)
    '        GC.Collect()

    '        Dim str_script As String = "window.open ('" & path_files & "');"
    '        lnk_file.Enabled = True
    '        lnk_file.Text = path_files
    '        lnk_file.OnClientClick = str_script

    '        'lbl_resp.Text = "ha finalizado el proceso correctamente.."
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '        obj_insert = New cls_insert_data
    '        obj_insert.CreaLogFile("error mfrm_corelativo_css : funcion read_file" & ex.Message)
    '    End Try

    'End Sub

    'Private Sub ReleaseObject(ByVal o As Object)
    '    Try
    '        While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
    '        End While
    '    Catch
    '    Finally
    '        o = Nothing
    '    End Try
    'End Sub

     
    Protected Sub btn_bajar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_bajar.Click
        Dim dtcorr_ As System.Data.DataTable
        obj_select = New cls_select_data
        ds_corr = obj_select.fun_get_csscorr()
        dtcorr_ = ds_corr.Tables(0)
        If dtcorr_.Rows.Count > 0 Then
            export_to_excel(dtcorr_)
        Else
            If (Not ClientScript.IsStartupScriptRegistered("alert")) Then
                Page.ClientScript.RegisterStartupScript _
                (Me.GetType(), "alert", "alertMe(3);", True)
            End If
        End If
        btn_bajar.Visible = False
    End Sub
End Class