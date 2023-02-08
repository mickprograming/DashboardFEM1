Module mdl_Consultas

    Dim Mes As String
    Public Sub CargarGrid()
        mdl_Conexion.AbrirConexionAces()
        Dim Cuadrilla As String
        'Dim Ano As Integer
        Dim Stsql As String

        mdl_Consultas.ValorMes()

        Stsql = "Select Dia, Sum (Totales) As Totales_, Sum(Buenas) As Buenas_, Sum (Horas) As Horas_, Round(AVG(OEE)) As OEE,
                Round(IIF(((Totales_-Buenas_)*100) = 0, 1,((Totales_-Buenas_)*100))/IIF(Totales_= 0,1,Totales_),2)As Waste,
                Round(Buenas_/(Horas_*1000),1) as MDH
                from Produccion where Maquina = '" & Form1.cmb_Maquina.Text & "' And Mes = '" & Mes & "' And Año = " & Form1.cmb_Año.Text & ""

        If Form1.cmb_Cuadrilla.Text <> "General" Then
            Cuadrilla = " And Cuadrilla = '" & Form1.cmb_Cuadrilla.Text & "' Group by Dia"
        Else
            Cuadrilla = " Group By Dia"

        End If
        Stsql = Stsql & Cuadrilla

        Dim da As New OleDb.OleDbDataAdapter(Stsql, Conexion)
        Dim ds As New DataSet
        da.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            Form1.dgv_TendenciaOEE.DataSource = ds.Tables(0)
        Else
            Form1.dgv_TendenciaOEE.DataSource = Nothing
        End If
        mdl_Conexion.CerrarConexionAcces()

        'Carga del grafico
        'Borrar y pasar a función aparte 

        Module1.CargarGraficasOEE()
        Module1.CargarGraficasObjetivos()

    End Sub

    Public Sub Promedios()
        Dim ds As New DataSet
        mdl_Conexion.AbrirConexionAces()
        Dim Cuadrilla As String
        'Dim Ano As Integer
        Dim Stsql As String
        mdl_Consultas.ValorMes()
        Stsql = "Select Format((AVG(OEE)/100), '##.##%') as OEE,
                Format(((SUM(Totales)-Sum(Buenas))/SUM(Totales)), '##.##%') as Waste,
                Format(((SUM(Buenas)/SUM(Horas))/1000),'##.##') As MDH,
                Format((Sum(Buenas)/1000),'####### su') as Volumen
                from Produccion where Maquina = '" & Form1.cmb_Maquina.Text & "' And Mes = '" & Mes & "' ANd Año = " & Form1.cmb_Año.Text & ""

        If Form1.cmb_Cuadrilla.Text <> "General" Then
            Cuadrilla = " And Cuadrilla = '" & Form1.cmb_Cuadrilla.Text & "'"
        End If
        Stsql = Stsql & Cuadrilla


        Dim da As New OleDb.OleDbDataAdapter(Stsql, Conexion)

        da.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 Then
            Form1.dgv_TendenciaOEE.DataSource = ds.Tables(0)
        Else
            Form1.dgv_TendenciaOEE.DataSource = Nothing
        End If

        Form1.lblOEE.Text = Form1.dgv_TendenciaOEE(0, 0).Value
        Form1.lblWaste.Text = Form1.dgv_TendenciaOEE(1, 0).Value
        Form1.lblMDH.Text = Form1.dgv_TendenciaOEE(2, 0).Value
        Form1.lbl_Volumen.Text = Form1.dgv_TendenciaOEE(3, 0).Value
        mdl_Conexion.CerrarConexionAcces()
    End Sub

    Public Sub ValorMes()
        Select Case Form1.cmb_Mes.Text
            Case = "Enero"
                Mes = "M01"
            Case = "Febrero"
                Mes = "M02"
            Case = "Marzo"
                Mes = "M03"
            Case = "Abril"
                Mes = "M04"
            Case = "Mayo"
                Mes = "M05"
            Case = "Junio"
                Mes = "M06"
            Case = "Julio"
                Mes = "M07"
            Case = "Agosto"
                Mes = "M08"
            Case = "Septiembre"
                Mes = "M09"
            Case = "Octubre"
                Mes = "M10"
            Case = "Noviembre"
                Mes = "M11"
            Case = "Diciembre"
                Mes = "M12"

        End Select
    End Sub
    Public Sub CargarObjetivos()

        mdl_Conexion.AbrirConexionAces()
        'Dim Maquina, Mes, Dia, Cuadrilla As String
        'Dim Ano As Integer
        Dim Stsql As String
        mdl_Consultas.ValorMes()
        Stsql = "Select Format(OEE,'##.##') As OEE,  Format(Waste,'##.##') As Waste, MPH, Format(Volumen,'## su')
                from Objetivos where Maquina = '" & Form1.cmb_Maquina.Text & "' And Mes = '" & Mes & "' ANd Año = " & Form1.cmb_Año.Text & ""

        Dim da As New OleDb.OleDbDataAdapter(Stsql, Conexion)
        Dim ds As New DataSet
        da.Fill(ds)


        If ds.Tables(0).Rows.Count > 0 Then
            Form1.dgv_Objetivos.DataSource = ds.Tables(0)
        Else
            Form1.dgv_Objetivos.DataSource = Nothing
        End If

        Form1.lbl_ObjOEE.Text = Form1.dgv_Objetivos(0, 0).Value
        Form1.lbl_ObjWaste.Text = Form1.dgv_Objetivos(1, 0).Value
        Form1.lbl_ObjMDH.Text = Form1.dgv_Objetivos(2, 0).Value
        Form1.lbl_ObjVolumen.Text = Form1.dgv_Objetivos(3, 0).Value
        mdl_Conexion.CerrarConexionAcces()
    End Sub
End Module
