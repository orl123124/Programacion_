Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Excel
Imports clsprocesos


Partial Public Class _Default
    Inherits System.Web.UI.Page
    Dim result As String
    Dim ruta As String
    Dim cadena As String
    Dim excel As Application
    Dim w As Workbook
    Dim sheet As Worksheet
    Dim range As Range
    Dim array(,) As Object
    Dim error_ As Integer
    Dim tabla As Hashtable
    Dim condic_ As Boolean
    Dim ds As DataSet
    Dim ds_corr As DataSet
    Dim obj_insert As cls_insert_data
    Dim obj_select As cls_select_data
    Dim obj_adm_control As cls_admin_controls


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    Protected Function read_file(ByVal dir As String)
        error_ = 0
        tabla = New Hashtable
        Dim row_m As Integer
        Try
            excel = New Application
            w = excel.Workbooks.Open(dir)
            'read all sheet
            For i As Integer = 1 To w.Sheets.Count
                'get sheet 
                sheet = w.Sheets(i)
                'get range
                range = sheet.UsedRange
                'load all cell into 2d array
                array = range.Value(XlRangeValueDataType.xlRangeValueDefault) 'format(fila,column)
                'scan the cells
                If array IsNot Nothing Then
                    Dim column As Integer 'colum
                    Dim fila As Integer 'fila 
                    fila = array.GetUpperBound(0)
                    column = array.GetUpperBound(1)
                    For fil As Integer = 2 To fila 'column
                        'For col As Integer = 1 To column 'fila
                        If column = 4 Then
                            If (array(1, 1)).ToString.ToUpper <> "CEDULA" Then
                                error_ = 1
                            End If

                            If (array(1, 2)).ToString.ToUpper <> "SEGURO" Then
                                error_ = 1
                            End If
                            If (array(1, 3)).ToString.ToUpper <> "NOMBRE" Then
                                error_ = 1
                            End If

                            If (array(1, 4)).ToString.ToUpper <> "CONSECUTIVO" Then
                                error_ = 1
                            End If
                            'CEDULA	SEGURO	NOMBRE	CONSECUTIVO
                            If (error_ <= 0) Then
                                tabla.Add("CEDULA", array(fil, 1))
                                tabla.Add("SEGURO", array(fil, 2))
                                tabla.Add("NOMBRE", array(fil, 3))
                                tabla.Add("CONSECUTIVO", array(fil, 4))
                                to_insert(tabla)
                            End If
                        End If
                        'Next
                    Next
                End If
            Next
            w.Close()
            '-------------------------
            obj_insert = New cls_insert_data
            obj_select = New cls_select_data

            ds_corr = obj_select.fun_get_csscorr

            row_m = obj_insert.inst_to_mastercorrelativo

            'export to excel-------------------------------------

            Dim dtcorr_ As System.Data.DataTable
            dtcorr_ = ds_corr.Tables(0)
            If dtcorr_.Rows.Count > 0 Then
                export_to_excel(dtcorr_)
            End If





        Catch ex As Exception
            obj_insert = New cls_insert_data
            obj_insert.CreaLogFile("error mfrm_corelativo_css : funcion read_file() :" & ex.Message)
        End Try
        Return ""
    End Function
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Try

      
            ruta = FileUpload1.PostedFile.FileName
            Dim extn_ = Path.GetExtension(FileUpload1.PostedFile.FileName.ToString).ToUpper

            If extn_.Equals(".XLS") Or extn_.Equals(".XLSX") Then
                condic_ = True
            Else
                condic_ = False
            End If


            System.Threading.Thread.Sleep(5000)
            If condic_ Then read_file(ruta)
        Catch ex As Exception
            obj_insert = New cls_insert_data
            obj_insert.CreaLogFile("error mfrm_corelativo_css : funcion read_file" & ex.Message)
        End Try
    End Sub

    Public Function to_insert(ByVal tabla As Hashtable) As Hashtable
        'Dim cedula As String
        'Dim seguro As String
        'Dim nombre As String
        'Dim consecutivo As String
        Dim row As Integer

        'cedula = tabla("cedula")
        'seguro = tabla("seguro")
        'nombre = tabla("nombre")
        'consecutivo = tabla("consecutivo")
        obj_insert = New cls_insert_data
        obj_select = New cls_select_data

        row = obj_insert.insert_to_csscorrelativo(tabla)
        

        tabla.Clear()
        Return tabla
    End Function

    Protected Sub bt_exportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles bt_exportar.Click
        Dim dtcorr_ As DataTable
        dtcorr_ = ds_corr.Tables(0)


    End Sub


    Public Sub export_to_excel(ByVal dt As System.Data.DataTable)

        Dim _excel As Application  'Excel.Application
        Dim wBook As Workbook
        Dim wSheet As Worksheet

        _excel = New Application
        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim rg As Microsoft.Office.Interop.Excel.Range
        Dim rpt As Microsoft.Office.Interop.Excel.Application
        Dim WorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim Sheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow

        Dim col_Index As Integer = 0 'col
        Dim row_Index As Integer = 0 'fila

        Dim v1 As String

        Try
            'creacion de columnas
            Dim rows_ As String = dt.Rows.Count
            For Each col As System.Data.DataColumn In dt.Columns
                col_Index = col_Index + 1
                wSheet.Cells(1, col_Index) = col.ColumnName
                Dim nom As String = wSheet.Cells(1, col_Index).ToString

            Next
            'creacion de filas
            For Each dr In dt.Rows
                row_Index = row_Index + 1
                col_Index = 0
                For Each col2 As System.Data.DataColumn In dt.Columns
                    col_Index = col_Index + 1
                    wSheet.Cells(row_Index + 1, col_Index) = dr(col2.ColumnName)
                Next
            Next
            '--------------------------------------------------------------------------------------------
            obj_adm_control = New cls_admin_controls
            Dim path_files As String


            'C:\correlativo_css
            path_files = "C:\correlativo_css\file" & "_" & Today.Day & _
               "_" & Today.Month & "_" & Today.Year & ".xlsx".ToString
            '----
            wSheet.Columns.AutoFit()
            Dim strFileName As String = path_files '"C:\inetpub\wwwroot\smw\correlativo_css\datatable.xlsx" '"K:\datatable.xlsx"
            'If System.IO.File.Exists(strFileName) Then
            '    System.IO.File.Delete(strFileName)
            'End If
            wBook.SaveAs(strFileName)

            
            

            ReleaseObject(wSheet)
            wBook.Close(False)
            ReleaseObject(wBook)
            _excel.Quit()
            ReleaseObject(_excel)
            GC.Collect()

            Dim str_script As String = "window.open ('" & path_files & "');"
            lnk_exp.Text = path_files
            lnk_exp.OnClientClick = str_script

            lbl_resp.Text = "ha finalizado el proceso correctamente.."
        Catch ex As Exception
            Throw New Exception(ex.Message)
            obj_insert = New cls_insert_data
            obj_insert.CreaLogFile("error mfrm_corelativo_css : funcion read_file" & ex.Message)
        End Try

    End Sub

    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
End Class