Imports System.Data.Odbc
Imports System.Data.OleDb


Public Class Principal
    Private version As String


    Public Function conectarBBDD(ByVal cadenaConexion As String) As OleDbConnection

        Dim conn As New OleDbConnection
        conn.ConnectionString = cadenaConexion

        Try
            conn.Open()
        Catch ex As Exception
            MessageBox.Show("Error al conectar con la base de datos")
        End Try

        Return conn


    End Function

    Public Function ejecutarSQL(ByVal cadenaSQL As String, ByVal conexion As OleDbConnection, ByVal tipo As Integer) As OleDbDataReader

        Dim dr As OleDbDataReader
        Dim comm = New OleDbCommand(cadenaSQL, conexion)


        If tipo = 1 Then

            Try
                dr = comm.ExecuteReader()
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString)

            End Try
        ElseIf tipo = 2 Then
            Try
                comm.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString)


            End Try
        End If

        Return dr

    End Function

    Public Sub Generar_csv_emitidas(ByVal datos As OleDbDataReader, ByVal fichero As String)

        Dim dr_iva As OleDbDataReader
        Dim cadenaSQL As String
        Dim conexion As OleDbConnection
        Dim odbc_cadena As String

        Dim campo As DataRow
        Dim propiedad As DataColumn
        Dim tabla As New DataTable
        Dim tabla_iva As New DataTable
        Dim resultado As String
        Dim fila() As DataRow

        Dim razon_social As String
        Dim NIF_empresa As String
        Dim tipo_comunicacion As String
        Dim ejercicio As String
        Dim periodo_mes As String

        Dim signo As String
        Dim tipo_factura_sage As String

        Dim nombre_receptor As String
        Dim nif_receptor As String
        Dim codigo_pais As String
        Dim tipo_identificador As String
        Dim fecha_expedicion As String
        Dim fecha_operacion As String
        Dim regimen_especial As String
        Dim descripcion As String
        Dim importe As String
        Dim serie_factura As String
        Dim numero_factura As String
        Dim numero_factura_fin As String
        Dim tipo_factura As String

        Dim tipo_no_exenta As String
        Dim base_no_exenta As String
        Dim porcentaje_no_exenta As String
        Dim cuota_no_exenta As String
        Dim porcentaje_RE_no_exenta As String
        Dim cuota_RE_no_exenta As String

        Dim divisa As Double

        Dim tipo_iva_1 As String
        Dim tipo_iva_2 As String
        Dim tipo_iva_3 As String
        Dim tipo_iva_4 As String

        Dim porc_iva_1 As String
        Dim porc_iva_2 As String
        Dim porc_iva_3 As String
        Dim porc_iva_4 As String

        Dim base_exenta As String
        Dim tipo_exenta As String

        Dim base_no_sujeta As String
        Dim tipo_no_sujeta As String
        Dim importe_no_sujeta As String
        Dim importe_TAI_no_sujeta As String

        Dim tipo_desglose_base2 As String
        Dim base_base2 As String
        Dim cuota_base2 As String
        Dim porcentaje_RE_base2 As String
        Dim cuota_RE_base2 As String

        Dim tipo_desglose_base3 As String
        Dim base_base3 As String
        Dim cuota_base3 As String
        Dim porcentaje_RE_base3 As String
        Dim cuota_RE_base3 As String

        Dim tipo_desglose_base4 As String
        Dim base_base4 As String
        Dim cuota_base4 As String
        Dim porcentaje_RE_base4 As String
        Dim cuota_RE_base4 As String

        Dim tipo_rectificativa As String
        Dim factura_origen As String
        Dim base_rectificativa As String
        Dim porc_rectificativa
        Dim cuota_rectificativa As String
        Dim fecha_factura_rectificativa As String

        Dim dua As String
        Dim fecha_dua As String

        Dim situacion As String
        Dim ref_catastral As String

        Dim intracomunitario As String
        Dim signo_factura As String

        Dim tipo_desglose As String
        Dim tipo_operacion As String
        Dim regimen As String
        Dim exenta As String
        Dim marca_tercero As String
        Dim modalidad As String
        Dim clave_reg_adicional1 As String
        Dim clave_reg_adicional2 As String
        Dim importe_recargo_equivalencia As String
        Dim file As System.IO.StreamWriter

        file = My.Computer.FileSystem.OpenTextFileWriter(fichero, False)

        ' OBTENEMOS LA TABLA DE IVAS

        conexion = conectarBBDD(Variables_globales.oledb_cadena)
        dr_iva = ejecutarSQL("SELECT t.VAT_0,t.VATRAT_0 FROM GESTION5.TABRATVAT t " & _
        " JOIN (SELECT VAT_0, MAX(STRDAT_0) as fecha FROM GESTION5.TABRATVAT GROUP BY VAT_0) m " & _
        " ON  m.VAT_0 = t.VAT_0 AND m.fecha = t.STRDAT_0 ORDER BY t.VAT_0", conexion, 1)
        tabla_iva.Load(dr_iva)
        dr_iva.Close()

        ' CABECERA

        NIF_empresa = "A18074922"
        razon_social = "TOSTADEROS SOL DE ALBA S.A."
        tipo_comunicacion = "A0"


        ' Pasamos el datareader a una datatable
        tabla.Load(datos)


        Dim i As Integer = 0
        For Each MiDataRow As DataRow In tabla.Rows()
            i = i + 1

            ' Inicializamos los valores

            ejercicio = ""
            periodo_mes = ""
            tipo_factura = ""
            nombre_receptor = ""
            nif_receptor = ""
            codigo_pais = ""
            tipo_identificador = ""
            serie_factura = ""
            numero_factura = ""
            numero_factura_fin = ""
            fecha_expedicion = ""
            fecha_operacion = ""
            regimen_especial = ""
            descripcion = ""
            importe = ""
            tipo_no_exenta = ""
            base_no_exenta = ""
            porcentaje_no_exenta = ""
            cuota_no_exenta = ""
            porcentaje_RE_no_exenta = ""
            cuota_RE_no_exenta = ""
            base_exenta = ""
            tipo_exenta = ""
            base_no_sujeta = ""
            tipo_no_sujeta = ""
            importe_no_sujeta = ""
            importe_TAI_no_sujeta = ""
            tipo_desglose_base2 = ""
            porc_iva_2 = ""
            base_base2 = ""
            cuota_base2 = ""
            porcentaje_RE_base2 = ""
            cuota_RE_base2 = ""
            tipo_desglose_base3 = ""
            porc_iva_3 = ""
            base_base3 = ""
            cuota_base3 = ""
            porcentaje_RE_base3 = ""
            cuota_RE_base3 = ""
            tipo_desglose_base4 = ""
            porc_iva_4 = ""
            base_base4 = ""
            cuota_base4 = ""
            porcentaje_RE_base4 = ""
            cuota_RE_base4 = ""
            tipo_rectificativa = ""
            factura_origen = ""
            base_rectificativa = ""
            porc_rectificativa = ""
            cuota_rectificativa = ""
            fecha_factura_rectificativa = ""
            dua = ""
            fecha_dua = ""
            situacion = ""
            ref_catastral = ""
            tipo_desglose = ""
            tipo_operacion = ""
            regimen = ""
            exenta = ""
            marca_tercero = ""
            modalidad = ""
            clave_reg_adicional1 = ""
            clave_reg_adicional2 = ""
            importe_recargo_equivalencia = ""

            periodo_mes = Format(Month(MiDataRow("ACCDAT_0")), "00")
            ejercicio = Year(MiDataRow("ACCDAT_0"))

            ' DESGLOSE IVA
            tipo_iva_1 = Trim((MiDataRow("TAX_0")))
            tipo_iva_2 = Trim((MiDataRow("TAX_1")))
            tipo_iva_3 = Trim((MiDataRow("TAX_2")))
            tipo_iva_4 = Trim((MiDataRow("TAX_3")))

            ' DEPENDE DE SEIDOR (SI SE INFORMAN LOS CAMPOS EN VACIO O CON CERO)
            If tipo_iva_1 = "0" Then tipo_iva_1 = ""
            If tipo_iva_2 = "0" Then tipo_iva_2 = ""
            If tipo_iva_3 = "0" Then tipo_iva_3 = ""
            If tipo_iva_4 = "0" Then tipo_iva_4 = ""


            ' DATOS FACTURA
            tipo_factura_sage = Trim(MiDataRow("SIVTYP_0"))
            If tipo_factura_sage = "FAC" Or tipo_factura_sage = "FAL" Then tipo_factura = "F1"
            If tipo_factura_sage = "ABC" Then tipo_factura = "R1"


            nombre_receptor = Trim(MiDataRow("BPRNAM_0"))
            nif_receptor = Trim(MiDataRow("CRN_0"))

            codigo_pais = Trim(MiDataRow("CRY_0"))
            intracomunitario = Trim(MiDataRow("EECNUM_0"))

            ' Si tiene nif y el pais es España, será codigo 1
            ' Si tiene codigo intracomunitario, será codigo 2
            ' Si tiene nif y el pais no es España, será extranjero, codigo 3

            ' Español
            If nif_receptor <> "" And codigo_pais = "ES" Then
                tipo_identificador = "01"
                tipo_desglose = "F"
                tipo_operacion = ""
                ' Intracomunitario
            ElseIf intracomunitario <> "" Then
                tipo_identificador = "02"
                nif_receptor = intracomunitario
                tipo_desglose = "O"
                tipo_operacion = "E"
                ' Extranjero
            ElseIf codigo_pais <> "ES" And nif_receptor <> "" Then
                tipo_identificador = "06"
                tipo_desglose = "O"
                tipo_operacion = "E"
            End If

            'serie_factura = Strings.Left(MiDataRow("NUM_0"), 9)
            'numero_factura = Strings.Mid(MiDataRow("NUM_0"), 11, Len(MiDataRow("NUM_0")))
            numero_factura = Trim(MiDataRow("NUM_0"))
            fecha_expedicion = Replace(MiDataRow("ACCDAT_0"), "/", "-")
            fecha_operacion = Replace(MiDataRow("ACCDAT_0"), "/", "-")

            ' Régimen general
            If tipo_iva_1 = "030" Then
                regimen_especial = "02"
            ElseIf tipo_iva_1 = "040" Then
                regimen_especial = "02" ' IGIC
            ElseIf tipo_factura_sage = "FAL" Then
                regimen_especial = "12" ' Alquileres
            Else
                regimen_especial = "01"
            End If

            descripcion = Trim(MiDataRow("DES_0") + MiDataRow("DES_1") + MiDataRow("DES_2"))

            If tipo_factura_sage = "ABC" And Trim(descripcion) = "" Then
                descripcion = "DEVOLUCION DE CLIENTE"
            ElseIf Trim(descripcion) = "" Then
                descripcion = "VENTA A CLIENTE"
            End If

            If tipo_factura_sage = "FAL" Then
                descripcion = "ALQUILER"
            End If

            signo = MiDataRow("SNS_0")

            ' TRANSFORMACION DIVISA CAMPO RATCUR_0
            divisa = Format(MiDataRow("RATCUR_0"), "#0.#0")
            If divisa = 0 Then divisa = 1

            importe = Format(MiDataRow("AMTATI_0"), "#0.#0")
            importe = importe * divisa
            If signo = "-1" Then importe = importe * -1
            If signo = "-1" Then signo_factura = "-" Else signo_factura = ""

            porc_iva_1 = ""
            porc_iva_2 = ""
            porc_iva_3 = ""
            porc_iva_4 = ""

            If Not String.IsNullOrEmpty(tipo_iva_1) Then
                ' Consultamos la tabla y cogemos la segunda columna para obtener el % de iva aplicable
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_1 & "'")
                porc_iva_1 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_2) Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_2 & "'")
                porc_iva_2 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_3) Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_3 & "'")
                porc_iva_3 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_4) Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_4 & "'")
                porc_iva_4 = fila(0)(1).ToString
            End If

            If tipo_iva_1 = "999" And porc_iva_1 = "0" Then porc_iva_1 = ""
            If tipo_iva_2 = "999" And porc_iva_2 = "0" Then porc_iva_2 = ""
            If tipo_iva_3 = "999" And porc_iva_3 = "0" Then porc_iva_3 = ""
            If tipo_iva_4 = "999" And porc_iva_4 = "0" Then porc_iva_4 = ""

            ' SUJETA NO EXENTA
            If tipo_iva_1 = "001" Or tipo_iva_1 = "002" Or tipo_iva_1 = "004" Then
                exenta = "N"
                tipo_no_exenta = "S1"
                base_no_exenta = signo_factura + Format(MiDataRow("BASTAX_0"), "#0.#0")
                porcentaje_no_exenta = porc_iva_1
                cuota_no_exenta = signo_factura + Format(MiDataRow("AMTTAX_0"), "#0.#0")
                porcentaje_RE_no_exenta = ""
                cuota_RE_no_exenta = ""
                base_exenta = ""
                tipo_exenta = ""
                base_no_sujeta = ""
                tipo_no_sujeta = ""
                importe_no_sujeta = ""
                importe_TAI_no_sujeta = ""
            End If

            ' SUJETA EXENTA
            If tipo_iva_1 = "000" Or tipo_iva_1 = "020" Or tipo_iva_1 = "030" Or tipo_iva_1 = "040" Then
                exenta = "S"
                tipo_no_exenta = ""
                base_no_exenta = ""
                porcentaje_no_exenta = ""
                cuota_no_exenta = ""
                porcentaje_RE_no_exenta = ""
                cuota_RE_no_exenta = ""
                base_exenta = signo_factura + Format(MiDataRow("BASTAX_0"), "#0.#0")

                If tipo_iva_1 = "000" Then tipo_exenta = "E6"
                If tipo_iva_1 = "020" Then tipo_exenta = "E5"
                If tipo_iva_1 = "030" Then tipo_exenta = "E2"
                If tipo_iva_1 = "040" Then tipo_exenta = "E2"

                base_no_sujeta = ""
                tipo_no_sujeta = ""
                importe_no_sujeta = ""
                importe_TAI_no_sujeta = ""
            End If

            ' NO SUJETA --- NO LAS DECLARAMOS EXCLUIR DE LA CONSULTA

            'If tipo_iva_1 = "999" Then ' Tipo de IVA para operaciones no sujetas?? preguntar a Jose Angel
            ' tipo_no_exenta = ""
            ' base_no_exenta = "0"
            ' porcentaje_no_exenta = "0"
            ' cuota_no_exenta = "0"
            ' porcentaje_RE_no_exenta = "0"
            ' cuota_RE_no_exenta = "0"
            ' base_exenta = "0"
            ' tipo_exenta = "0"
            ' base_no_sujeta = Format(MiDataRow("BASTAX_0"), "#0.#0")
            ' tipo_no_sujeta = "NS"
            ' importe_no_sujeta = "0"  '??
            ' importe_TAI_no_sujeta = "0" '??
            ' End If

            ' BASE 2
            If tipo_iva_2 = "001" Or tipo_iva_2 = "002" Or tipo_iva_2 = "003" Then
                tipo_desglose_base2 = "S1"
            ElseIf tipo_iva_2 = "000" Then
                tipo_desglose_base2 = "E6"
            ElseIf tipo_iva_2 = "020" Then
                tipo_desglose_base2 = "E2"
            ElseIf tipo_iva_2 = "030" Then
                tipo_desglose_base2 = "E2"
            Else : tipo_desglose_base2 = ""
            End If


            If MiDataRow("BASTAX_1") <> "0" Then base_base2 = signo_factura + Format(MiDataRow("BASTAX_1"), "#0.#0") Else base_base2 = ""
            If MiDataRow("AMTTAX_1") <> "0" Then cuota_base2 = signo_factura + Format(MiDataRow("AMTTAX_1"), "#0.#0") Else cuota_base2 = ""
            porcentaje_RE_base2 = ""
            cuota_RE_base2 = ""

            ' BASE 3
            If tipo_iva_3 = "001" Or tipo_iva_3 = "002" Or tipo_iva_3 = "003" Then
                tipo_desglose_base3 = "S1"
            ElseIf tipo_iva_3 = "000" Then
                tipo_desglose_base3 = "E6"
            ElseIf tipo_iva_3 = "020" Then
                tipo_desglose_base3 = "E2"
            ElseIf tipo_iva_3 = "030" Then
                tipo_desglose_base3 = "E2"
            Else : tipo_desglose_base3 = ""
            End If


            If MiDataRow("BASTAX_2") <> 0 Then base_base2 = signo_factura + Format(MiDataRow("BASTAX_2"), "#0.#0") Else base_base3 = ""
            If MiDataRow("AMTTAX_2") <> 0 Then cuota_base2 = signo_factura + Format(MiDataRow("AMTTAX_2"), "#0.#0") Else cuota_base3 = ""
            porcentaje_RE_base3 = ""
            cuota_RE_base3 = ""

            ' BASE 4
            If tipo_iva_4 = "001" Or tipo_iva_4 = "002" Or tipo_iva_4 = "003" Then
                tipo_desglose_base4 = "S1"
            ElseIf tipo_iva_4 = "000" Then
                tipo_desglose_base4 = "E6"
            ElseIf tipo_iva_4 = "020" Then
                tipo_desglose_base4 = "E2"
            ElseIf tipo_iva_4 = "030" Then
                tipo_desglose_base4 = "E2"
            Else : tipo_desglose_base4 = ""
            End If


            If MiDataRow("BASTAX_3") <> 0 Then base_base4 = signo_factura + Format(MiDataRow("BASTAX_3"), "#0.#0") Else base_base4 = ""
            If MiDataRow("AMTTAX_3") <> 0 Then cuota_base4 = signo_factura + Format(MiDataRow("AMTTAX_3"), "#0.#0") Else cuota_base4 = ""
            porcentaje_RE_base4 = ""
            cuota_RE_base4 = ""

            ' FACTURA RECTIFICATIVA
            If tipo_factura = "R1" Then tipo_rectificativa = "I" Else tipo_rectificativa = ""
            factura_origen = ""
            base_rectificativa = ""
            porc_rectificativa = ""
            cuota_rectificativa = ""
            fecha_factura_rectificativa = ""

            ' ADUANA
            dua = ""
            fecha_dua = ""

            If tipo_factura_sage = "FAL" Then
                ' INMUEBLE
                situacion = "1"
                ref_catastral = Trim(MiDataRow("DES_0") + MiDataRow("DES_1") + MiDataRow("DES_2"))
            Else
                situacion = ""
                ref_catastral = ""
            End If

            ' INFO ADICIONAL
            tipo_desglose = tipo_desglose
            tipo_operacion = tipo_operacion
            regimen = ""
            exenta = ""
            marca_tercero = ""
            modalidad = ""
            clave_reg_adicional1 = ""
            clave_reg_adicional2 = ""
            importe_recargo_equivalencia = ""

            ' CAMBIAMOS LA COMA POR PUNTO EN LOS IMPORTES

            importe = Replace(importe, ",", ".")
            base_no_exenta = Replace(base_no_exenta, ",", ".")
            porcentaje_no_exenta = Replace(porcentaje_no_exenta, ",", ".")
            cuota_no_exenta = Replace(cuota_no_exenta, ",", ".")
            porcentaje_RE_no_exenta = Replace(porcentaje_RE_no_exenta, ",", ".")
            cuota_RE_no_exenta = Replace(cuota_RE_no_exenta, ",", ".")
            base_no_exenta = Replace(base_no_exenta, ",", ".")
            porcentaje_no_exenta = Replace(porcentaje_no_exenta, ",", ".")
            cuota_no_exenta = Replace(cuota_no_exenta, ",", ".")
            porcentaje_RE_no_exenta = Replace(porcentaje_RE_no_exenta, ",", ".")
            cuota_RE_no_exenta = Replace(cuota_RE_no_exenta, ",", ".")
            base_exenta = Replace(base_exenta, ",", ".")
            tipo_exenta = Replace(tipo_exenta, ",", ".")
            base_no_sujeta = Replace(base_no_sujeta, ",", ".")
            tipo_no_sujeta = Replace(tipo_no_sujeta, ",", ".")
            importe_no_sujeta = Replace(importe_no_sujeta, ",", ".")
            importe_TAI_no_sujeta = Replace(importe_TAI_no_sujeta, ",", ".")
            porc_iva_2 = Replace(porc_iva_2, ",", ".")
            base_base2 = Replace(base_base2, ",", ".")
            cuota_base2 = Replace(cuota_base2, ",", ".")
            porcentaje_RE_base2 = Replace(porcentaje_RE_base2, ",", ".")
            cuota_RE_base2 = Replace(cuota_RE_base2, ",", ".")
            porc_iva_3 = Replace(porc_iva_3, ",", ".")
            base_base3 = Replace(base_base3, ",", ".")
            cuota_base3 = Replace(cuota_base3, ",", ".")
            porcentaje_RE_base3 = Replace(porcentaje_RE_base3, ",", ".")
            cuota_RE_base3 = Replace(cuota_RE_base3, ",", ".")
            porc_iva_4 = Replace(porc_iva_4, ",", ".")
            base_base4 = Replace(base_base4, ",", ".")
            cuota_base4 = Replace(cuota_base4, ",", ".")
            porcentaje_RE_base4 = Replace(porcentaje_RE_base4, ",", ".")
            cuota_RE_base4 = Replace(cuota_RE_base4, ",", ".")



            ' GRABACION A FICHERO

            file.WriteLine(razon_social & ";" _
                           & NIF_empresa & ";" _
                           & tipo_comunicacion & ";" _
                           & ejercicio & ";" _
                           & periodo_mes & ";" _
                           & tipo_factura & ";" _
                           & nombre_receptor & ";" _
                           & nif_receptor & ";" _
                           & codigo_pais & ";" _
                           & tipo_identificador & ";" _
                           & serie_factura & ";" _
                           & numero_factura & ";" _
                           & numero_factura_fin & ";" _
                           & fecha_expedicion & ";" _
                           & fecha_operacion & ";" _
                           & regimen_especial & ";" _
                           & descripcion & ";" _
                           & importe & ";" _
                           & tipo_no_exenta & ";" _
                           & base_no_exenta & ";" _
                           & porcentaje_no_exenta & ";" _
                           & cuota_no_exenta & ";" _
                           & porcentaje_RE_no_exenta & ";" _
                           & cuota_RE_no_exenta & ";" _
                           & base_exenta & ";" _
                           & tipo_exenta & ";" _
                           & base_no_sujeta & ";" _
                           & tipo_no_sujeta & ";" _
                           & importe_no_sujeta & ";" _
                           & importe_TAI_no_sujeta & ";" _
                           & tipo_desglose_base2 & ";" _
                           & porc_iva_2 & ";" _
                           & base_base2 & ";" _
                           & cuota_base2 & ";" _
                           & porcentaje_RE_base2 & ";" _
                           & cuota_RE_base2 & ";" _
                           & tipo_desglose_base3 & ";" _
                           & porc_iva_3 & ";" _
                           & base_base3 & ";" _
                           & cuota_base3 & ";" _
                           & porcentaje_RE_base3 & ";" _
                           & cuota_RE_base3 & ";" _
                           & tipo_desglose_base4 & ";" _
                           & porc_iva_4 & ";" _
                           & base_base4 & ";" _
                           & cuota_base4 & ";" _
                           & porcentaje_RE_base4 & ";" _
                           & cuota_RE_base4 & ";" _
                           & tipo_rectificativa & ";" _
                           & factura_origen & ";" _
                           & base_rectificativa & ";" _
                           & porc_rectificativa & ";" _
                           & cuota_rectificativa & ";" _
                           & fecha_factura_rectificativa & ";" _
                           & dua & ";" _
                           & fecha_dua & ";" _
                           & situacion & ";" _
                           & ref_catastral & ";" _
                           & tipo_desglose & ";" _
                           & tipo_operacion & ";" _
                           & regimen & ";" _
                           & exenta & ";" _
                           & marca_tercero & ";" _
                           & modalidad & ";" _
                           & clave_reg_adicional1 & ";" _
                           & clave_reg_adicional2 & ";" _
                           & importe_recargo_equivalencia
)


            ' GRABAMOS EN LA TABLA WSII LAS FACTURAS QUE HEMOS GENERADO

            Dim hoy As String
            hoy = DateTime.Now.ToString("dd/MM/yyyy")

            ' VER AL FINAL PARA LLEVAR EL CONTROL DE LO GENERADO, NO ACTIVAR EN PRUEBAS
            cadenaSQL = "INSERT INTO GESTION5.WSII_EMITIDAS VALUES ('" & tipo_factura_sage & "','" & Trim(MiDataRow("INVTYP_0")) & "','" & Trim(MiDataRow("NUM_0")) & "',TO_DATE ('" & fecha_expedicion & "', 'dd/mm/yyyy'),TO_DATE ('" & fecha_operacion & "', 'dd/mm/yyyy'),TO_DATE ('" & hoy & "', 'dd/mm/yyyy'))"
            ejecutarSQL(cadenaSQL, conexion, 2)


        Next
        conexion.Close()
        file.Close()



    End Sub

    Public Sub Generar_csv_recibidas(ByVal datos As OleDbDataReader, ByVal fichero As String)

        Dim dr_iva As OleDbDataReader
        Dim cadenaSQL As String
        Dim conexion As OleDbConnection
        Dim odbc_cadena As String

        Dim campo As DataRow
        Dim propiedad As DataColumn
        Dim tabla As New DataTable
        Dim tabla_iva As New DataTable
        Dim resultado As String
        Dim fila() As DataRow

        Dim razon_social As String
        Dim NIF_empresa As String
        Dim tipo_comunicacion As String
        Dim ejercicio As String
        Dim periodo_mes As String

        Dim signo As String
        Dim tipo_factura_sage As String

        Dim nombre_proveedor As String
        Dim nif_proveedor As String
        Dim codigo_pais As String
        Dim tipo_identificador As String
        Dim fecha_expedicion As String
        Dim fecha_operacion As String
        Dim fecha_reg_contable As String
        Dim regimen_especial As String
        Dim descripcion As String
        Dim importe As String
        Dim cuota_deducible As String
        Dim serie_factura As String
        Dim numero_factura As String
        Dim numero_factura_fin As String
        Dim tipo_factura As String

        Dim tipo_no_exenta As String
        Dim base_no_exenta As String
        Dim porcentaje_no_exenta As String
        Dim cuota_no_exenta As String
        Dim porcentaje_RE_no_exenta As String
        Dim cuota_RE_no_exenta As String

        Dim divisa As Double

        Dim tipo_iva_1 As String
        Dim tipo_iva_2 As String
        Dim tipo_iva_3 As String
        Dim tipo_iva_4 As String

        Dim porc_iva_1 As String
        Dim porc_iva_2 As String
        Dim porc_iva_3 As String
        Dim porc_iva_4 As String

        Dim base_exenta As String
        Dim tipo_exenta As String

        Dim base_no_sujeta As String
        Dim tipo_no_sujeta As String
        Dim importe_no_sujeta As String
        Dim importe_TAI_no_sujeta As String

        Dim inversion_sujeto_pasivo_1 As String
        Dim base_base1 As String
        Dim cuota_base1 As String
        Dim porcentaje_RE_base1 As String
        Dim cuota_RE_base1 As String

        Dim inversion_sujeto_pasivo_2 As String
        Dim base_base2 As String
        Dim cuota_base2 As String
        Dim porcentaje_RE_base2 As String
        Dim cuota_RE_base2 As String

        Dim inversion_sujeto_pasivo_3 As String
        Dim base_base3 As String
        Dim cuota_base3 As String
        Dim porcentaje_RE_base3 As String
        Dim cuota_RE_base3 As String

        Dim inversion_sujeto_pasivo_4 As String
        Dim base_base4 As String
        Dim cuota_base4 As String
        Dim porcentaje_RE_base4 As String
        Dim cuota_RE_base4 As String

        Dim base1 As Double
        Dim base2 As Double
        Dim base3 As Double
        Dim base4 As Double
        Dim cuota1 As Double
        Dim cuota2 As Double
        Dim cuota3 As Double
        Dim cuota4 As Double

        Dim tipo_rectificativa As String
        Dim factura_origen As String
        Dim base_rectificativa As String
        Dim porc_rectificativa
        Dim cuota_rectificativa As String
        Dim fecha_factura_rectificativa As String

        Dim dua As String
        Dim fecha_dua As String

        Dim situacion As String
        Dim ref_catastral As String

        Dim tipo_desglose As String
        Dim tipo_operacion As String
        Dim regimen As String
        Dim exenta As String
        Dim marca_tercero As String
        Dim modalidad As String
        Dim clave_reg_adicional1 As String
        Dim clave_reg_adicional2 As String

        Dim importe_cal As Double
        Dim importe_calc2 As Double
        Dim cuota_cal As Double
        Dim signo_factura As String

        Dim intracomunitario As String




        Dim file As System.IO.StreamWriter

        file = My.Computer.FileSystem.OpenTextFileWriter(fichero, False)

        ' OBTENEMOS LA TABLA DE IVAS

        conexion = conectarBBDD(Variables_globales.oledb_cadena)
        dr_iva = ejecutarSQL("SELECT t.VAT_0,t.VATRAT_0 FROM GESTION5.TABRATVAT t " & _
        " JOIN (SELECT VAT_0, MAX(STRDAT_0) as fecha FROM GESTION5.TABRATVAT GROUP BY VAT_0) m " & _
        " ON  m.VAT_0 = t.VAT_0 AND m.fecha = t.STRDAT_0 ORDER BY t.VAT_0", conexion, 1)
        tabla_iva.Load(dr_iva)
        dr_iva.Close()

        ' CABECERA

        NIF_empresa = "A18074922"

        razon_social = "TOSTADEROS SOL DE ALBA S.A."
        tipo_comunicacion = "A0"

        ' Pasamos el datereader a una datatable
        tabla.Load(datos)

        Dim i As Integer = 0
        For Each MiDataRow As DataRow In tabla.Rows()
            i = i + 1

            ' INICIALIZAMOS LOS CAMPOS
            
            ejercicio = ""
            periodo_mes = ""
            tipo_factura = ""
            nombre_proveedor = ""
            nif_proveedor = ""
            codigo_pais = ""
            tipo_identificador = ""
            serie_factura = ""
            numero_factura = ""
            numero_factura_fin = ""
            fecha_expedicion = ""
            fecha_operacion = ""
            fecha_reg_contable = ""
            regimen_especial = ""
            descripcion = ""
            importe = ""
            cuota_deducible = ""
            inversion_sujeto_pasivo_1 = ""
            base_base1 = ""
            porc_iva_1 = ""
            cuota_base1 = ""
            porcentaje_RE_base1 = ""
            cuota_RE_base1 = ""
            inversion_sujeto_pasivo_2 = ""
            base_base2 = ""
            porc_iva_2 = ""
            cuota_base2 = ""
            porcentaje_RE_base2 = ""
            cuota_RE_base2 = ""
            inversion_sujeto_pasivo_3 = ""
            base_base3 = ""
            porc_iva_3 = ""
            cuota_base3 = ""
            porcentaje_RE_base3 = ""
            cuota_RE_base3 = ""
            inversion_sujeto_pasivo_4 = ""
            base_base4 = ""
            porc_iva_4 = ""
            cuota_base4 = ""
            porcentaje_RE_base4 = ""
            cuota_RE_base4 = ""
            tipo_rectificativa = ""
            factura_origen = ""
            base_rectificativa = ""
            porc_rectificativa = ""
            cuota_rectificativa = ""
            fecha_factura_rectificativa = ""
            dua = ""
            fecha_dua = ""
            clave_reg_adicional1 = ""
            clave_reg_adicional2 = ""

            ' Tomamos los valores de la fecha de operación
            periodo_mes = Format(Month(MiDataRow("ACCDAT_0")), "00")
            ejercicio = Year(MiDataRow("ACCDAT_0"))

            ' DESGLOSE IVA
            tipo_iva_1 = Trim((MiDataRow("TAX_0")))
            tipo_iva_2 = Trim((MiDataRow("TAX_1")))
            tipo_iva_3 = Trim((MiDataRow("TAX_2")))
            tipo_iva_4 = Trim((MiDataRow("TAX_3")))

            ' DEPENDE DE SEIDOR (SI SE INFORMAN LOS CAMPOS EN VACIO O CON CERO)
            If tipo_iva_1 = "0" Then tipo_iva_1 = ""
            If tipo_iva_2 = "0" Then tipo_iva_2 = ""
            If tipo_iva_3 = "0" Then tipo_iva_3 = ""
            If tipo_iva_4 = "0" Then tipo_iva_4 = ""


            ' DATOS FACTURA

            ' TIPO FACTURA


            tipo_factura_sage = Trim(MiDataRow("PIVTYP_0"))

            ' FACTURAS F1
            If tipo_iva_1 = "000" Or tipo_iva_1 = "001" Or tipo_iva_1 = "002" Or tipo_iva_1 = "004" Or tipo_iva_1 = "021" Or tipo_iva_1 = "022" Or tipo_iva_1 = "023" Then tipo_factura = "F1" ' Factura normal
            If tipo_iva_1 = "R11" Or tipo_iva_1 = "R12" Or tipo_iva_1 = "R13" Or tipo_iva_1 = "R14" Or tipo_iva_1 = "R19" Or tipo_iva_1 = "R22" Then tipo_factura = "F1" ' Factura normal
            If tipo_iva_1 = "024" Or tipo_iva_1 = "025" Or tipo_iva_1 = "026" Then tipo_factura = "F1" ' Factura normal
            ' FACTURAS F5
            If tipo_iva_1 = "031" Or tipo_iva_1 = "032" Or tipo_iva_1 = "033" Then tipo_factura = "F5" ' Importaciones
            ' FACTURAS R1
            If tipo_factura_sage = "ABP" Then tipo_factura = "R1" ' Abonos


            ' CAMPO CRN_0 DE BPARTNER = NIF ESPAÑOL
            ' CAMPO EECNUM_0 DE BPADDRESS = NIF INTRACOMUNITARIO
            ' Y CRY_0 NO ES "ES" Y TIENE NIF ES EXTRANJERO
            ' NIF ESPAÑOL TIPO 1
            ' CEE TIPO 2
            ' EXT TIPO 6

            ' Si es una factura de aduana, el nif_proveedor y el nombre es Tostaderos
            If tipo_factura = "F5" Then
                nombre_proveedor = "TOSTADEROS SOL DE ALBA S.A."
                nif_proveedor = "A18074922"
            Else
                nombre_proveedor = Trim(MiDataRow("BPRNAM_0"))
                nif_proveedor = Trim(MiDataRow("CRN_0"))
            End If

            codigo_pais = Trim(MiDataRow("CRY_0"))
            intracomunitario = Trim(MiDataRow("EECNUM_0"))

            ' Si tiene nif y el pais es España, será codigo 1
            ' Si tiene codigo intracomunitario, será codigo 2
            ' Si tiene nif y el pais no es España, será extranjero, codigo 3

            If nif_proveedor <> "" And codigo_pais = "ES" Then
                tipo_identificador = "01"
            ElseIf intracomunitario <> "" Then
                tipo_identificador = "02"
                nif_proveedor = intracomunitario
            ElseIf codigo_pais <> "ES" And nif_proveedor <> "" Then
                tipo_identificador = "06"
            End If

            ' numero de factura de proveedor el campo es BPRVCR_0 EN PRODUCCION

            serie_factura = ""
            numero_factura = Trim(MiDataRow("BPRVCR_0"))

            fecha_expedicion = Replace(MiDataRow("BPRDAT_0"), "/", "-")
            fecha_operacion = Replace(MiDataRow("ACCDAT_0"), "/", "-")
            fecha_reg_contable = Replace(MiDataRow("CREDAT_0"), "/", "-")

            ' Régimen especial

            If tipo_factura = "F5" Then
                regimen_especial = "01"
            ElseIf intracomunitario <> "" Then
                regimen_especial = "09"
            Else
                regimen_especial = "01"
            End If

            descripcion = Trim(MiDataRow("DES_0"))

            If descripcion = "" And (tipo_factura_sage = "ABP" Or tipo_factura_sage = "ADB") Then
                descripcion = "DEVOLUCION A PROVEEDOR"
            ElseIf descripcion = "" Then
                descripcion = "COMPRA A PROVEEDOR"
            End If


            signo = Trim(MiDataRow("SNS_0"))


            If Trim(MiDataRow("DES_2")) = "ALQUILER" Then regimen_especial = "12"

            ' TRANSFORMACION DIVISA CAMPO RATCUR_0
            divisa = Format(MiDataRow("RATCUR_0"), "#0.#0")

            ' CALCULAMOS A MANO EL IMPORTE (SUMA DE LAS BASES + LAS CUOTAS) Y LA CUOTA (SUMA DE LAS DEDUCCIONES)
            ' OMITIMOS LAS REDUCCIONES Y LAS EXENCIONES EXTRANJERAS

            If tipo_iva_1.Contains("R") Or tipo_iva_1 = "030" Then
                base1 = 0
                cuota1 = 0
            Else
                base1 = MiDataRow("BASTAX_0")
                cuota1 = MiDataRow("AMTTAX_0")
            End If

            If tipo_iva_2.Contains("R") Or tipo_iva_2 = "030" Then
                base2 = 0
                cuota2 = 0
            Else
                base2 = MiDataRow("BASTAX_1")
                cuota2 = MiDataRow("AMTTAX_1")
            End If

            If tipo_iva_3.Contains("R") Or tipo_iva_3 = "030" Then
                base3 = 0
                cuota3 = 0
            Else
                base3 = MiDataRow("BASTAX_2")
                cuota3 = MiDataRow("AMTTAX_2")
            End If

            If tipo_iva_4.Contains("R") Or tipo_iva_4 = "030" Then
                base4 = 0
                cuota4 = 0
            Else
                base4 = MiDataRow("BASTAX_3")
                cuota4 = MiDataRow("AMTTAX_3")
            End If


            importe_cal = base1 + base2 + base3 + base4
            cuota_cal = cuota1 + cuota2 + cuota3 + cuota4

            importe_cal = importe_cal + cuota_cal
            importe_cal = importe_cal * divisa

            ' PARA INTRACOMUNITARIOS
            If intracomunitario <> "" Then importe_cal = MiDataRow("AMTATI_0")

            If signo = "-1" Then importe_cal = importe_cal * -1
            If signo = "-1" Then cuota_cal = cuota_cal * -1
            If signo = "-1" Then signo_factura = "-" Else signo_factura = ""

            importe = Format(importe_cal, "#0.#0")

            cuota_deducible = Format(cuota_cal, "#0.#0")

            porc_iva_1 = ""
            porc_iva_2 = ""
            porc_iva_3 = ""
            porc_iva_4 = ""

            If Not String.IsNullOrEmpty(tipo_iva_1) And Not tipo_iva_1.Contains("R") And tipo_iva_1 <> "030" Then
                ' Consultamos la tabla y cogemos la segunda columna para obtener el % de iva aplicable
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_1 & "'")
                porc_iva_1 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_2) And Not tipo_iva_2.Contains("R") And tipo_iva_2 <> "030" Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_2 & "'")
                porc_iva_2 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_3) And Not tipo_iva_3.Contains("R") And tipo_iva_3 <> "030" Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_3 & "'")
                porc_iva_3 = fila(0)(1).ToString
            End If
            If Not String.IsNullOrEmpty(tipo_iva_4) And Not tipo_iva_4.Contains("R") And tipo_iva_4 <> "030" Then
                fila = tabla_iva.Select("VAT_0='" & tipo_iva_4 & "'")
                porc_iva_4 = fila(0)(1).ToString
            End If


            ' Calculamos el importe de la factura manualmente para las ADUANAS
            ' En este caso solo se informará una línea por factura como máximo

            If tipo_iva_1 = "031" Or tipo_iva_1 = "032" Or tipo_iva_1 = "033" Then
                importe = Format(((MiDataRow("BASTAX_0") * 100) / porc_iva_1), "#0.#0")
                cuota_deducible = Format((MiDataRow("BASTAX_0")), "#0.#0")
            End If

            ' BASE 1
            If Not String.IsNullOrEmpty(tipo_iva_1) Then
                If tipo_iva_1 = "024" Or tipo_iva_1 = "025" Or tipo_iva_1 = "026" Then inversion_sujeto_pasivo_1 = "Y" Else inversion_sujeto_pasivo_1 = "N"
            Else : inversion_sujeto_pasivo_1 = ""
            End If
            If tipo_iva_1 <> "031" And tipo_iva_1 <> "032" And tipo_iva_1 <> "033" Then
                If base1 <> 0 Then
                    base_base1 = signo_factura + Format(base1, "#0.#0")
                ElseIf tipo_iva_1 <> "" Then
                    base_base1 = "0.00"
                Else
                    base_base1 = ""
                End If
                If cuota1 <> 0 Then
                    cuota_base1 = signo_factura + Format(cuota1, "#0.#0")
                ElseIf tipo_iva_1 <> "" Then
                    cuota_base1 = "0.00"
                Else
                    cuota_base1 = ""
                End If
            Else
                base_base1 = signo_factura + Format(((MiDataRow("BASTAX_0") * 100) / porc_iva_1), "#0.#0")
                cuota_base1 = signo_factura + Format((MiDataRow("BASTAX_0")), "#0.#0")
            End If
            'If cuota_base1 = "0,00" Then cuota_base1 = ""
            porcentaje_RE_base1 = ""
            cuota_RE_base1 = ""

            ' REDUCCIONES
            If tipo_iva_1.StartsWith("R") Then
                base_base1 = "0.00"
                porc_iva_1 = "0.00"
                cuota_base1 = "0.00"
                porcentaje_RE_base1 = "0.00"
                cuota_RE_base1 = "0.00"
            End If

            ' BASE 2
            If Not String.IsNullOrEmpty(tipo_iva_2) Then
                If tipo_iva_2 = "024" Or tipo_iva_2 = "025" Or tipo_iva_2 = "026" Then inversion_sujeto_pasivo_2 = "Y" Else inversion_sujeto_pasivo_2 = "N"
            Else : inversion_sujeto_pasivo_2 = ""
            End If
            If base2 <> 0 Then
                base_base2 = signo_factura + Format(base2, "#0.#0")
            ElseIf Trim(tipo_iva_2) <> "" Then
                base_base2 = "0.00"
            Else
                base_base2 = ""
            End If
            If cuota2 <> 0 Then
                cuota_base2 = signo_factura + Format(cuota2, "#0.#0")
            ElseIf Trim(tipo_iva_2) <> "" Then
                cuota_base2 = "0.00"
            Else
                cuota_base2 = ""
            End If
            'If cuota_base2 = "0,00" Then cuota_base2 = ""
            porcentaje_RE_base2 = ""
            cuota_RE_base2 = ""

            ' REDUCCIONES (13/07/17)
            If tipo_iva_2.StartsWith("R") Then
                base_base2 = "0.00"
                porc_iva_2 = "0.00"
                cuota_base2 = "0.00"
                porcentaje_RE_base2 = "0.00"
                cuota_RE_base2 = "0.00"
            End If

            ' BASE 3
            If Not String.IsNullOrEmpty(tipo_iva_3) Then
                If tipo_iva_3 = "024" Or tipo_iva_3 = "025" Or tipo_iva_3 = "026" Or tipo_iva_3.Contains("R%") Then inversion_sujeto_pasivo_3 = "Y" Else inversion_sujeto_pasivo_3 = "N"
            Else : inversion_sujeto_pasivo_3 = ""
            End If

            If base3 <> 0 Then
                base_base3 = signo_factura + Format(base3, "#0.#0")
            ElseIf Trim(tipo_iva_3) <> "" Then
                base_base3 = "0.00"
            Else
                base_base3 = ""
            End If
            If cuota3 <> 0 Then
                cuota_base3 = signo_factura + Format(cuota3, "#0.#0")
            ElseIf Trim(tipo_iva_3) <> "" Then
                cuota_base3 = "0.00"
            Else
                cuota_base3 = ""
            End If
            'If cuota_base3 = "0,00" Then cuota_base3 = ""
            porcentaje_RE_base3 = ""
            cuota_RE_base3 = ""

            ' REDUCCIONES (13/07/17)
            If tipo_iva_3.StartsWith("R") Then
                base_base3 = "0.00"
                porc_iva_3 = "0.00"
                cuota_base3 = "0.00"
                porcentaje_RE_base3 = "0.00"
                cuota_RE_base3 = "0.00"
            End If


            ' BASE 4
            If Not String.IsNullOrEmpty(tipo_iva_4) Then
                If tipo_iva_4 = "024" Or tipo_iva_4 = "025" Or tipo_iva_4 = "026" Or tipo_iva_4.Contains("R%") Then inversion_sujeto_pasivo_4 = "Y" Else inversion_sujeto_pasivo_4 = "N"
            Else : inversion_sujeto_pasivo_4 = ""
            End If

            If base4 <> 0 Then
                base_base4 = signo_factura + Format(base4, "#0.#0")
            ElseIf Trim(tipo_iva_4) <> "" Then
                base_base4 = "0.00"
            Else
                base_base4 = ""
            End If
            If cuota4 <> 0 Then
                cuota_base4 = signo_factura + Format(cuota4, "#0.#0")
            ElseIf Trim(tipo_iva_4) <> "" Then
                cuota_base4 = "0.00"
            Else
                cuota_base4 = ""
            End If
            'If cuota_base4 = "0,00" Then cuota_base4 = ""
            porcentaje_RE_base4 = ""
            cuota_RE_base4 = ""

            ' REDUCCIONES (13/07/17)
            If tipo_iva_4.StartsWith("R") Then
                base_base4 = "0.00"
                porc_iva_4 = "0.00"
                cuota_base4 = "0.00"
                porcentaje_RE_base1 = "0.00"
                cuota_RE_base1 = "0.00"
            End If

            ' FACTURA RECTIFICATIVA
            If tipo_factura = "R1" Then tipo_rectificativa = "I" Else tipo_rectificativa = ""
            factura_origen = ""
            base_rectificativa = ""
            porc_rectificativa = ""
            cuota_rectificativa = ""
            fecha_factura_rectificativa = ""

            ' ADUANA
            dua = Trim(MiDataRow("XDUA_0"))
            If Trim(dua) <> "" Then
                fecha_dua = Replace(MiDataRow("XFDUA_0"), "/", "-")
                ' EL CÓDIGO DUA COMPLETO SE INFORMARÁ EN EL NÚMERO DE FACTURA DEL PROVEEDOR POR TAMAÑO DEL CAMPO
                dua = Trim(MiDataRow("BPRVCR_0"))
            Else : fecha_dua = ""
            End If

            ' INFO ADICIONAL
            clave_reg_adicional1 = ""
            clave_reg_adicional2 = ""

            ' CAMBIAMOS LA COMA POR PUNTO EN LOS IMPORTES
            importe = Replace(importe, ",", ".")
            cuota_deducible = Replace(cuota_deducible, ",", ".")

            porc_iva_1 = Replace(porc_iva_1, ",", ".")
            base_base1 = Replace(base_base1, ",", ".")
            cuota_base1 = Replace(cuota_base1, ",", ".")
            porcentaje_RE_base1 = Replace(porcentaje_RE_base1, ",", ".")
            cuota_RE_base1 = Replace(cuota_RE_base1, ",", ".")

            porc_iva_2 = Replace(porc_iva_2, ",", ".")
            base_base2 = Replace(base_base2, ",", ".")
            cuota_base2 = Replace(cuota_base2, ",", ".")
            porcentaje_RE_base2 = Replace(porcentaje_RE_base2, ",", ".")
            cuota_RE_base2 = Replace(cuota_RE_base2, ",", ".")

            porc_iva_3 = Replace(porc_iva_3, ",", ".")
            base_base3 = Replace(base_base3, ",", ".")
            cuota_base3 = Replace(cuota_base3, ",", ".")
            porcentaje_RE_base3 = Replace(porcentaje_RE_base3, ",", ".")
            cuota_RE_base3 = Replace(cuota_RE_base3, ",", ".")

            porc_iva_4 = Replace(porc_iva_4, ",", ".")
            base_base4 = Replace(base_base4, ",", ".")
            cuota_base4 = Replace(cuota_base4, ",", ".")
            porcentaje_RE_base4 = Replace(porcentaje_RE_base4, ",", ".")
            cuota_RE_base4 = Replace(cuota_RE_base4, ",", ".")


            ' GRABACION A FICHERO
            file.WriteLine(razon_social & ";" _
                           & NIF_empresa & ";" _
                           & tipo_comunicacion & ";" _
                           & ejercicio & ";" _
                           & periodo_mes & ";" _
                           & tipo_factura & ";" _
                           & nombre_proveedor & ";" _
                           & nif_proveedor & ";" _
                           & codigo_pais & ";" _
                           & tipo_identificador & ";" _
                           & serie_factura & ";" _
                           & numero_factura & ";" _
                           & numero_factura_fin & ";" _
                           & fecha_expedicion & ";" _
                           & fecha_operacion & ";" _
                           & fecha_reg_contable & ";" _
                           & regimen_especial & ";" _
                           & descripcion & ";" _
                           & importe & ";" _
                           & cuota_deducible & ";" _
                           & inversion_sujeto_pasivo_1 & ";" _
                           & base_base1 & ";" _
                           & porc_iva_1 & ";" _
                           & cuota_base1 & ";" _
                           & porcentaje_RE_base1 & ";" _
                           & cuota_RE_base1 & ";" _
                           & inversion_sujeto_pasivo_2 & ";" _
                           & base_base2 & ";" _
                           & porc_iva_2 & ";" _
                           & cuota_base2 & ";" _
                           & porcentaje_RE_base2 & ";" _
                           & cuota_RE_base2 & ";" _
                           & inversion_sujeto_pasivo_3 & ";" _
                           & base_base3 & ";" _
                           & porc_iva_3 & ";" _
                           & cuota_base3 & ";" _
                           & porcentaje_RE_base3 & ";" _
                           & cuota_RE_base3 & ";" _
                           & inversion_sujeto_pasivo_4 & ";" _
                           & base_base4 & ";" _
                           & porc_iva_4 & ";" _
                           & cuota_base4 & ";" _
                           & porcentaje_RE_base4 & ";" _
                           & cuota_RE_base4 & ";" _
                           & tipo_rectificativa & ";" _
                           & factura_origen & ";" _
                           & base_rectificativa & ";" _
                           & porc_rectificativa & ";" _
                           & cuota_rectificativa & ";" _
                           & fecha_factura_rectificativa & ";" _
                           & dua & ";" _
                           & fecha_dua & ";" _
                           & clave_reg_adicional1 & ";" _
                           & clave_reg_adicional2
)


            ' GRABAMOS EN LA TABLA WSII LAS FACTURAS QUE HEMOS GENERADO
            Dim hoy As String
            hoy = DateTime.Now.ToString("dd/MM/yyyy")


            cadenaSQL = "INSERT INTO GESTION5.WSII_RECIBIDAS VALUES ('" & tipo_factura_sage & "','" & Trim(MiDataRow("INVTYP_0")) & "','" & Trim(MiDataRow("NUM_0")) & "',TO_DATE ('" & fecha_expedicion & "', 'dd/mm/yyyy'),TO_DATE ('" & fecha_operacion & "', 'dd/mm/yyyy'), TO_DATE('" & fecha_reg_contable & "', 'dd/mm/yyyy'), TO_DATE('" & hoy & "', 'dd/mm/yyyy'), '" & numero_factura & "')"
            ejecutarSQL(cadenaSQL, conexion, 2)


        Next
        conexion.Close()
        file.Close()



    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim conn As OleDbConnection
        Dim dr As OleDbDataReader
        Dim sql As String
        Dim fecha_desde As String
        Dim fecha_hasta As String
        Dim fecha_SII As String
        Dim fichero As String

        fecha_desde = DateTimePicker1.Value.Date.ToShortDateString()
        fecha_hasta = DateTimePicker2.Value.Date.ToShortDateString()
        fecha_SII = "01/07/2017"



        ' COMPROBACIÓN DE CAMPOS

        If TextBox2.Text = "" Then
            MessageBox.Show("Falta seleccionar el archivo de salida", "SII", MessageBoxButtons.OK)
            Exit Sub
        End If

        If CDate(fecha_desde) > CDate(fecha_hasta) Then
            MessageBox.Show("Error en rango de fechas", "SII", MessageBoxButtons.OK)
            Exit Sub
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TextBox2.Text) Then
            MessageBox.Show("No existe el directorio de destino", "SII", MessageBoxButtons.OK)
            Exit Sub
        End If

        ' FIN COMPROBACIÓN DE CAMPOS

        Cursor = Cursors.WaitCursor

        fichero = TextBox2.Text + "\"

        If RadioButton1.Checked = True Then fichero = fichero & "fichero_SII_emitidas.csv"
        If RadioButton2.Checked = True Then fichero = fichero & "fichero_SII_recibidas.csv"

        conn = conectarBBDD(Variables_globales.oledb_cadena)

        ' TABLAS(SAGE)
        '
        ' SINVOICE   =   CABECERA FACTURA DE VENTAS
        ' PINVOICE   =   CABECERA FACTURA DE COMPRAS

        ' EMITIDAS
        If RadioButton1.Checked = True Then

            ' TEST
            'sql = "SELECT SIVTYP_0,INVTYP_0,NUM_0,BPRVCR_0,CRN_0, EECNUM_0,BPARTNER.BPRNAM_0,BPARTNER.CRY_0,SINVOICE.CREDAT_0,ACCDAT_0,SNS_0,RATCUR_0,AMTATI_0,TAX_0,TAX_1,TAX_2,TAX_3,BASTAX_0,BASTAX_1,BASTAX_2,BASTAX_3,AMTTAX_0,AMTTAX_1,AMTTAX_2,AMTTAX_3,AMTTAX_4,DES_0,DES_1,DES_2 " & _
            '   " FROM GESTION5.SINVOICE, GESTION5.BPARTNER  " & _
            '   " WHERE BPRNUM_0=BPR_0 AND NUM_0 in ('FAC-TFS17-00765','FAC-TFS17-00818') AND CPY_0='TSA' AND NUM_0 NOT IN (SELECT NUM_FAC FROM GESTION5.WSII_EMITIDAS) AND SIVTYP_0 IN ('ABC','FAC','FAL') AND TAX_0<>'999'"

            ' REAL

            sql = "SELECT SIVTYP_0,INVTYP_0,NUM_0,BPRVCR_0,CRN_0, EECNUM_0,BPARTNER.BPRNAM_0,BPARTNER.CRY_0,SINVOICE.CREDAT_0,ACCDAT_0,SNS_0,RATCUR_0,AMTATI_0,TAX_0,TAX_1,TAX_2,TAX_3,BASTAX_0,BASTAX_1,BASTAX_2,BASTAX_3,AMTTAX_0,AMTTAX_1,AMTTAX_2,AMTTAX_3,AMTTAX_4,DES_0,DES_1,DES_2 " & _
               " FROM GESTION5.SINVOICE, GESTION5.BPARTNER  " & _
               " WHERE BPRNUM_0=BPR_0 AND SINVOICE.ACCDAT_0 BETWEEN TO_DATE ('" & fecha_desde & "', 'dd/mm/yyyy') AND TO_DATE ('" & fecha_hasta & "', 'dd/mm/yyyy') AND SINVOICE.ACCDAT_0 >=TO_DATE ('" & fecha_SII & "', 'dd/mm/yyyy') " & _
                " AND CPY_0='TSA' AND SIVTYP_0 IN ('ABC','FAC','FAL') AND TAX_0<>'999' AND STA_0=3 " & _
                " AND NUM_0 NOT IN (SELECT NUM_FAC FROM GESTION5.WSII_EMITIDAS) ORDER BY ACCDAT_0, NUM_0"



            dr = ejecutarSQL(sql, conn, 1)
            If Not dr.HasRows Then
                MessageBox.Show("No existen datos en las fechas indicadas", "SII", MessageBoxButtons.OK)
                Cursor = Cursors.Default
                Exit Sub
            End If

            Generar_csv_emitidas(dr, fichero)

        End If

        ' RECIBIDAS
        If RadioButton2.Checked = True Then

            ' TEST
            'sql = "SELECT PIVTYP_0,INVTYP_0,NUM_0,CRN_0, EECNUM_0,BPARTNER.BPRNAM_0,BPARTNER.CRY_0,PINVOICE.CREDAT_0,ACCDAT_0,BPRDAT_0,SNS_0,RATCUR_0,AMTATI_0,TAX_0,TAX_1,TAX_2,TAX_3,BASTAX_0,BASTAX_1,BASTAX_2,BASTAX_3,AMTTAX_0,AMTTAX_1,AMTTAX_2,AMTTAX_3,AMTTAX_4,DES_0,DES_1,DES_2,XDUA_0,XFDUA_0" & _
            '   " FROM GESTION5.PINVOICE, GESTION5.BPARTNER " & _
            '   " WHERE BPRNUM_0=BPR_0 AND PINVOICE.CREDAT_0 BETWEEN TO_DATE ('" & fecha_desde & "', 'dd/mm/yyyy') AND TO_DATE ('" & fecha_hasta & "', 'dd/mm/yyyy') AND CPY_0='TSA' AND NUM_0 NOT IN (SELECT NUM_FAC FROM GESTION5.WSII_RECIBIDAS) AND PIVTYP_0 IN ('ABP','FAP') and TAX_0<>'030'"

            ' REAL

            sql = "SELECT PIVTYP_0,INVTYP_0,NUM_0,BPRVCR_0,CRN_0, EECNUM_0,BPARTNER.BPRNAM_0,BPARTNER.CRY_0,PINVOICE.CREDAT_0,ACCDAT_0,BPRDAT_0,SNS_0,RATCUR_0,AMTATI_0,TAX_0,TAX_1,TAX_2,TAX_3,BASTAX_0,BASTAX_1,BASTAX_2,BASTAX_3,AMTTAX_0,AMTTAX_1,AMTTAX_2,AMTTAX_3,AMTTAX_4,DES_0,DES_1,DES_2,XDUA_0,XFDUA_0" & _
                " FROM GESTION5.PINVOICE, GESTION5.BPARTNER " & _
                " WHERE BPRNUM_0=BPR_0 AND PINVOICE.CREDAT_0 BETWEEN TO_DATE ('" & fecha_desde & "', 'dd/mm/yyyy') AND TO_DATE ('" & fecha_hasta & "', 'dd/mm/yyyy') AND PINVOICE.ACCDAT_0 >=TO_DATE ('" & fecha_SII & "', 'dd/mm/yyyy') " & _
                " AND CPY_0='TSA' AND PIVTYP_0 IN ('ABP','FAP','FCC','ADB') AND STA_0 = 3 " & _
                " AND NUM_0 NOT IN (SELECT NUM_FAC FROM GESTION5.WSII_RECIBIDAS) AND NUM_0 NOT IN (SELECT NUM_0 FROM GESTION5.PINVOICE WHERE TAX_0='030' AND TAX_1=' ') ORDER BY CREDAT_0, NUM_0"

            dr = ejecutarSQL(sql, conn, 1)
            If Not dr.HasRows Then
                MessageBox.Show("No existen datos en las fechas indicadas", "SII", MessageBoxButtons.OK)
                Cursor = Cursors.Default
                Exit Sub
            End If


            Generar_csv_recibidas(dr, fichero)
        End If

        Cursor = Cursors.Default

        MessageBox.Show("Fichero generado correctamente en " & fichero, "SII", MessageBoxButtons.OK)

        dr.Close()
        conn.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TextBox2.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
        Me.CenterToScreen()
        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\SII"
        Me.Text = "SII Tostaderos Sol de Alba " + Variables_globales.version
        Me.DateTimePicker1.Value = Now()
        Me.DateTimePicker2.Value = Now()
    End Sub


    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub RecibidasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecibidasToolStripMenuItem.Click
        Consulta.Tag = 2
        Consulta.Show()

    End Sub

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub


    Private Sub EmitidasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmitidasToolStripMenuItem.Click
        Consulta.Tag = 1
        Consulta.Show()

    End Sub

    Private Sub SalirToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub


End Class
