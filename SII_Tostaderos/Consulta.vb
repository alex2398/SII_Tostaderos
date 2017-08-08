Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class Consulta

    Private cadenaConexion As String = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.0.221)(PORT=1521)))(CONNECT_DATA=(SID=V153)(SERVER=DEDICATED)));User Id=system;Password=manager;"
    Private bindingSource1 As New BindingSource()
    Private dataAdapter_emitidas As New OleDbDataAdapter()
    Private WithEvents reloadButton As New Button()
    Private WithEvents submitButton As New Button()
    Dim todos_marcados As Boolean
    Private Sub Consulta_emitidas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        If Me.Tag = 1 Then Me.Text = "Consulta facturas emitidas" Else Me.Text = "Consulta facturas recibidas"

        Me.CenterToScreen()
        'TODO: esta línea de código carga datos en la tabla 'DataSetEmitidas.WSII_EMITIDAS' Puede moverla o quitarla según sea necesario.
        Me.DateTimePicker1.Value = Now
        Me.DateTimePicker2.Value = Now
        Label4.Text = ""
        ' Vaciamos y volvemos a cargar el datasource
        Me.bindingSource1.DataSource = DBNull.Value
        Me.DataGridView1.DataSource = Me.bindingSource1

        todos_marcados = False


    End Sub
    


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        obtener_datos()
        


    End Sub

    'Private Sub DataGridView1_RowMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick

    'Dim factura As String
    ' Dim comando As String
    ' Dim conn As New OleDbConnection
    '
    '        factura = DataGridView1.Rows(e.RowIndex).Cells(3).Value
    '    If Me.Tag = 1 Then
    '            comando = "DELETE FROM GESTION5.WSII_EMITIDAS WHERE NUM_FAC= '" & factura & "'"
    '    Else
    '            comando = "DELETE FROM GESTION5.WSII_RECIBIDAS WHERE NUM_FAC= '" & factura & "'"
    '    End If
    '
    '    Dim result As Integer = MessageBox.Show("Se borrará el registro, ¿está seguro?", "SII", MessageBoxButtons.YesNo)
    '    If result = DialogResult.Yes Then
    '            ' Borramos el registro de la base de datos
    '
    '            conn.ConnectionString = cadenaConexion
    '    Dim comm = New OleDbCommand(comando, conn)
    '    Try
    '                conn.Open()
    '    ' Insert code to process data.
    '    Catch ex As Exception
    '                MessageBox.Show("Error al conectar con la base de datos")
    '    End Try
    '
    '   Try
    '               comm.ExecuteNonQuery()
    '               MessageBox.Show("Registro borrado correctamente")
    '               obtener_datos()
    '   Catch ex As Exception
    '               MessageBox.Show(ex.Message.ToString)


    'End Try

    'End If


    'End Sub



    Private Sub obtener_datos()
        Dim comando As String
        Dim fecha_inicial As String
        Dim fecha_final As String
        Dim conn As New OleDbConnection

        fecha_inicial = DateTimePicker1.Value.Date.ToShortDateString()
        fecha_final = DateTimePicker2.Value.Date.ToShortDateString()

        Me.bindingSource1.DataSource = DBNull.Value
        Me.DataGridView1.DataSource = Me.bindingSource1
        DataGridView1.Columns.Clear()
        ' Create a new data adapter based on the specified query.
        ' TODO: Modify the connection string and include any
        ' additional required properties for your database.
        conn.ConnectionString = cadenaConexion

        Try
            conn.Open()
            ' Insert code to process data.
        Catch ex As Exception
            MessageBox.Show("Error al conectar con la base de datos")
        End Try


        If Me.Tag = 1 Then
            comando = "SELECT 0 as ID,TIPO_FAC,INV_TIP,NUM_FAC,FECHA_FAC,FECHA_CRE,FECHA_PRE from GESTION5.WSII_EMITIDAS WHERE FECHA_FAC BETWEEN TO_DATE ('" & fecha_inicial & "', 'dd/mm/yyyy') AND TO_DATE ('" & fecha_final & "', 'dd/mm/yyyy') ORDER BY FECHA_FAC, NUM_FAC"
        Else
            comando = "SELECT 0 as ID,TIPO_FAC,INV_TIP,NUM_FAC_PRV,NUM_FAC,FECHA_EXP,FECHA_OPE,FECHA_CRE,FECHA_PRE from GESTION5.WSII_RECIBIDAS WHERE FECHA_CRE BETWEEN TO_DATE ('" & fecha_inicial & "', 'dd/mm/yyyy') AND TO_DATE ('" & fecha_final & "', 'dd/mm/yyyy') ORDER BY FECHA_CRE, NUM_FAC"
        End If

        'comando.CommandText = "select * from GESTION5.WSII_EMITIDAS"
        Me.dataAdapter_emitidas = New OleDbDataAdapter(comando, conn)


        ' Create a command builder to generate SQL update, insert, and
        ' delete commands based on selectCommand. These are used to
        ' update the database.
        Dim commandBuilder As New OleDbCommandBuilder(Me.dataAdapter_emitidas)


        ' Populate a new data table and bind it to the BindingSource.
        Dim tabla As New DataTable()

        Me.dataAdapter_emitidas.Fill(tabla)

        tabla.Columns.Add(New DataColumn("CheckBox", GetType(System.Boolean)))

        'tabla.Columns("CheckBoxMap").ColumnMapping = MappingType.Hidden
        tabla.Columns("CheckBox").SetOrdinal(0)
        tabla.Columns("CheckBox").DefaultValue = False
        For Each row As DataRow In tabla.Rows
            row("CheckBox") = False
        Next
        tabla.AcceptChanges()



        If tabla IsNot Nothing And tabla.Rows.Count > 0 Then
            Me.bindingSource1.DataSource = tabla

            ' Resize the DataGridView columns to fit the newly loaded content.
            'Me.DataGridView1.AutoResizeColumns( _
            '    DataGridViewAutoSizeColumnsMode.DisplayedCells)

            If Me.Tag = "1" Then
                Me.DataGridView1.Columns(0).HeaderCell.Value = ""
                Me.DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(1).HeaderCell.Value = "ID"
                Me.DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(2).HeaderCell.Value = "FAC"
                Me.DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(3).HeaderCell.Value = "TIPO"
                Me.DataGridView1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(4).HeaderCell.Value = "NUM. FAC"
                Me.DataGridView1.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(5).HeaderCell.Value = "FECHA OPERACION"
                Me.DataGridView1.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Me.DataGridView1.Columns(5).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(6).HeaderCell.Value = "FECHA CREACION"
                Me.DataGridView1.Columns(6).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(6).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Me.DataGridView1.Columns(7).HeaderCell.Value = "FECHA GENERACION"
                Me.DataGridView1.Columns(7).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(7).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Else
                Me.DataGridView1.Columns(0).HeaderCell.Value = ""
                Me.DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(1).HeaderCell.Value = "ID"
                Me.DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(2).HeaderCell.Value = "FAC"
                Me.DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(3).HeaderCell.Value = "TIPO"
                Me.DataGridView1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(4).HeaderCell.Value = "FAC. PRV"
                Me.DataGridView1.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(5).HeaderCell.Value = "NUM. FAC"
                Me.DataGridView1.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                Me.DataGridView1.Columns(6).HeaderCell.Value = "FECHA OPERACION"
                Me.DataGridView1.Columns(6).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Me.DataGridView1.Columns(6).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(7).HeaderCell.Value = "FECHA REGISTRO"
                Me.DataGridView1.Columns(7).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Me.DataGridView1.Columns(7).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(8).HeaderCell.Value = "FECHA CREACION"
                Me.DataGridView1.Columns(8).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                Me.DataGridView1.Columns(8).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(9).HeaderCell.Value = "FECHA GENERACION"
                Me.DataGridView1.Columns(9).Resizable = DataGridViewTriState.False
                Me.DataGridView1.Columns(9).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If

            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                row.Cells(1).Value = row.Index + 1
            Next


            Me.DataGridView1.AllowUserToAddRows = False
            Label4.Text = "TOTAL: " + Me.DataGridView1.RowCount.ToString + " registros"


        Else
            Me.bindingSource1.DataSource = DBNull.Value
            Me.DataGridView1.DataSource = Me.bindingSource1
            Label4.Text = "No existen registros"
        End If



    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.RowCount > 0 Then

            Dim comando As String
            Dim conn As New OleDbConnection


            If Me.Tag = 1 Then
                comando = "DELETE FROM GESTION5.WSII_EMITIDAS WHERE NUM_FAC IN ("
            Else
                comando = "DELETE FROM GESTION5.WSII_RECIBIDAS WHERE NUM_FAC IN ("
            End If


            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim isSelected As Boolean = Convert.ToBoolean(row.Cells(0).Value)
                If isSelected Then
                    comando = comando + "'" + row.Cells(5).Value.ToString() + "',"

                End If
            Next
            comando = Strings.Left(comando, comando.Length - 1)
            comando = comando + ")"
            'MessageBox.Show(comando)


            Dim result As Integer = MessageBox.Show("Se borrarán los registros seleccionados, ¿está seguro?", "SII", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                ' Borramos el registro de la base de datos

                conn.ConnectionString = cadenaConexion
                Dim comm = New OleDbCommand(comando, conn)
                Try
                    conn.Open()
                    ' Insert code to process data.
                Catch ex As Exception
                    MessageBox.Show("Error al conectar con la base de datos")
                End Try

                Try
                    comm.ExecuteNonQuery()
                    MessageBox.Show("Registro borrado correctamente")
                    obtener_datos()
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString)


                End Try
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex <> -1 Then
            If Convert.ToBoolean(DataGridView1.Rows(e.RowIndex).Cells(0).Value) = False Then
                DataGridView1.Rows(e.RowIndex).Cells(0).Value = True
            Else
                DataGridView1.Rows(e.RowIndex).Cells(0).Value = False
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        If e.RowIndex = -1 And e.ColumnIndex = 0 Then

            If todos_marcados = False Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Index >= 0 Then
                        DataGridView1.Rows(row.Index).Cells(0).Value = True
                    End If
                Next
                todos_marcados = True
            Else
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Index >= 0 Then
                        DataGridView1.Rows(row.Index).Cells(0).Value = False
                    End If
                Next
                todos_marcados = False
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        Dim archivo As String

        Cursor.Current = Cursors.WaitCursor

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Hoja1")
        Dim valor As String

        For i = 1 To DataGridView1.RowCount
            For j = 1 To DataGridView1.ColumnCount - 1
                For k As Integer = 1 To DataGridView1.Columns.Count - 1
                    xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k).HeaderText
                    valor = DataGridView1(j, i - 1).Value.ToString()
                    xlWorkSheet.Cells(i + 1, j) = valor


                Next
            Next
        Next
        xlWorkSheet.Columns.AutoFit()
        If Me.Tag = 1 Then
            xlWorkSheet.Columns(5).NumberFormat = "mm\/dd\/yyyy"
            xlWorkSheet.Columns(6).NumberFormat = "mm\/dd\/yyyy"
            xlWorkSheet.Columns(7).NumberFormat = "mm\/dd\/yyyy"
        Else
            xlWorkSheet.Columns(6).NumberFormat = "mm\/dd\/yyyy"
            xlWorkSheet.Columns(7).NumberFormat = "mm\/dd\/yyyy"
            xlWorkSheet.Columns(8).NumberFormat = "mm\/dd\/yyyy"
            xlWorkSheet.Columns(9).NumberFormat = "mm\/dd\/yyyy"
        End If
        Cursor.Current = Cursors.Default

        Dim sd As New SaveFileDialog 'declare save file dialog
        sd.Filter = "Excel File|*.xlsx"
        sd.Title = "Guardar como"


        If sd.ShowDialog = Windows.Forms.DialogResult.OK Then 'check if save file dialog was close after selecting a path
            If sd.FileName <> "" Then
                Cursor.Current = Cursors.WaitCursor
                archivo = sd.FileName
                xlWorkSheet.SaveAs(archivo) 'sd.filename reurns save file dialog path
                xlWorkBook.Close()
                xlApp.Quit()
            End If
        End If

            releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
        Cursor.Current = Cursors.Default
        MsgBox("Archivo exportado correctamente a " & sd.FileName)

        ' Da error al volver al programa y además deja procesos en memoria
        'Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        'startInfo.FileName = "EXCEL.EXE"
        'startInfo.Arguments = archivo
        'Process.Start(startInfo)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class