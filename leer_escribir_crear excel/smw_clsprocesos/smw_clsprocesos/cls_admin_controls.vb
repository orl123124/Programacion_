Imports System.IO
Imports Optimizing.CryptoLib
Imports System.Security.Permissions.FileIOPermission
Imports Helper
Imports System.Data
Imports System.Threading.ThreadStateException
Imports System
Imports System.Net


Public Class cls_admin_controls
    Dim clsinsertdata As New cls_insert_data
    Dim clsselectdata As New cls_select_data
    Dim hlpCore As New ClsHelperCore


    Public Function FormateaFecha(ByVal valor As Integer, ByVal tipo As Integer) As String
        FormateaFecha = ""
        Try
            If tipo = 1 Then
                If Len(Trim(valor)) = 4 Then
                    FormateaFecha = Mid(valor, 3, 4)
                Else
                    FormateaFecha = Convert.ToString(valor)

                End If
            Else
                If Len(Trim(valor)) < 2 Then
                    FormateaFecha = "0" & Convert.ToString(valor)
                Else
                    FormateaFecha = Convert.ToString(valor)
                End If
            End If

        Catch e As Exception
            clsinsertdata.CreaLogFile_clsprocesos("ERROR: " & e.ToString)
        End Try

    End Function
   
    Public Function valida_isNull(ByVal valor As Object) As String

        If IsDBNull(valor) And Not IsNumeric(valor) Then
            Return ""
        Else
            Return valor
        End If
    End Function


    Public Function valida_pdoactivo(ByVal fecha As Date, ByVal tipo As String) As Boolean

        Dim ds As DataSet

        Dim result As Boolean = False

        Try

            ds = clsselectdata.llena_combos("FECHA_FIN_PAGO, FECHA_INI_MARCACION, FECHA_FIN_MARCACION", "SMW_PDOACTIVO", 0, "", "")
            Select Case tipo
                Case "m" 'marcaciones
                    If fecha >= CDate(ds.Tables(0).Rows(0).Item("FECHA_INI_MARCACION")) And fecha <= CDate(ds.Tables(0).Rows(0).Item("FECHA_FIN_PAGO")) Then
                        result = True
                    End If

                Case "t" 'turnos
                    'para crear y actualizar turnos, la fecha de turno debe ser mayor a la fecha del periodo ya pagado
                    'If fecha >= CDate(ds.Tables(0).Rows(0).Item("FECHA_INI_MARCACION")) Then

                    'If Microsoft.VisualBasic.DatePart("ww", fecha) >= Microsoft.VisualBasic.DatePart("ww", CDate(ds.Tables(0).Rows(0).Item("FECHA_INI_MARCACION"))) Then
                    If Microsoft.VisualBasic.DatePart("ww", fecha) >= Microsoft.VisualBasic.DatePart("ww", CDate(ds.Tables(0).Rows(0).Item("FECHA_INI_MARCACION"))) Or fecha >= CDate(ds.Tables(0).Rows(0).Item("FECHA_INI_MARCACION")) Then

                        result = True
                    End If
            End Select

        Catch ex As Exception
            clsinsertdata.CreaLogFile_clsprocesos("Ha ocurrido un error en cls_admin_controls - Seccion valida_pdoactivo. Detalle del error:" & ex.ToString)


        End Try

        Return result

    End Function


    ''Echavez Para la creacion del Batch para anexo03
    Public Function EjecucionBatch(ByVal parametrosbd As String, ByVal file As String, ByVal name As String, ByVal user_pass As String) As Integer

        Try
            crearSQL(parametrosbd, file, name) '*listo
            crearBat(user_pass, file, name)
            Enviar_FTP(file, name & ".cmd")

        Catch ex As Exception

            clsinsertdata.CreaLogFile(ex.Message & "envioFTP funcion envioFTP")
        End Try
        Return 1
    End Function


    Public Sub crearSQL(ByVal parametrosbd As String, ByVal file As String, ByVal name As String)

        Try

            Dim fso As New System.IO.FileStream(file & name & ".sql", _
            IO.FileMode.Create, FileAccess.Write)

            Dim w As New StreamWriter(fso)
            w.BaseStream.Seek(0, SeekOrigin.End)
            w.WriteLine("set echo on;")
            w.WriteLine("exec " & parametrosbd)
            w.WriteLine("exit;")
            w.Flush()
            w.Close()
            fso = Nothing
            w = Nothing

        Catch ex As Exception
            clsinsertdata.CreaLogFile(ex.Message & "envioFTP, crearConnFTP")
        End Try
    End Sub

    Public Sub crearBat(ByVal basedatos As String, _
                            ByVal filename As String, _
                            ByVal name As String)

        Try

            Dim fso As New System.IO.FileStream(String.Concat(filename, name & ".cmd"), _
            IO.FileMode.Create, FileAccess.Write)
            Dim w As New StreamWriter(fso)
            w.BaseStream.Seek(0, SeekOrigin.End)
            w.WriteLine(String.Concat("cmd/C sqlplus.exe " & basedatos & " @" & filename & name & ".sql"))
            w.Flush()
            w.Close()
            fso = Nothing
            w = Nothing
        Catch ex As Exception
            clsinsertdata.CreaLogFile(ex.Message & "envioFTP, crearBatFTP")


        End Try
    End Sub

    Public Sub Enviar_FTP(ByVal filebat As String, ByVal BatName As String)
        Try
            'Dim ProgId As Integer
            Shell(String.Concat(filebat, BatName), AppWinStyle.Hide, True, 10)

        Catch ex As Exception
            clsinsertdata.CreaLogFile(ex.Message & "envioFTP, Enviar_FTP")
        End Try
    End Sub


    Public Function FileExists(ByVal FileFullPath As String) As Boolean

        Dim f As New IO.FileInfo(FileFullPath)
        Return f.Exists

    End Function

End Class
