
Module Module1

    Public Sub CargarGraficasOEE()

        Dim col, fil, i As Integer
        col = Form1.dgv_TendenciaOEE.Rows.Count
        On Error Resume Next
        Form1.chr_OEE.Series("Dia").Points.Clear()
        Form1.chr_Waste.Series("Dia").Points.Clear()
        Form1.chr_MDH.Series("Dia").Points.Clear()
        Form1.chr_OEE.Series("Rojo").Points.Clear()
        Form1.chr_Waste.Series("Rojo").Points.Clear()
        Form1.chr_MDH.Series("Rojo").Points.Clear()



        Do While i <= col
            If Form1.dgv_TendenciaOEE(4, i).Value > Form1.dgv_Objetivos(0, 0).Value Then
                Form1.chr_OEE.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(4, i).Value)
                Form1.chr_OEE.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            Else
                Form1.chr_OEE.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(4, i).Value)
                Form1.chr_OEE.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            End If
            i = i + 1
        Loop
        i = 0
        Do While i <= col
            If Form1.dgv_TendenciaOEE(5, i).Value < Form1.dgv_Objetivos(1, 0).Value Then
                Form1.chr_Waste.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(5, i).Value)
                Form1.chr_Waste.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            Else
                Form1.chr_Waste.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(5, i).Value)
                Form1.chr_Waste.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            End If
            i = i + 1
        Loop
        i = 0
        Do While i <= col
            If Form1.dgv_TendenciaOEE(6, i).Value > Form1.dgv_Objetivos(2, 0).Value Then
                Form1.chr_MDH.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(6, i).Value)
                Form1.chr_MDH.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            Else
                Form1.chr_MDH.Series("Rojo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(6, i).Value)
                Form1.chr_MDH.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, " ")
            End If
            i = i + 1
        Loop







        'Do While i <= col
        '    Form1.chr_Waste.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(5, i).Value)
        '    i = i + 1
        'Loop
        'i = 0
        'Do While i <= col
        '    Form1.chr_MDH.Series("Dia").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_TendenciaOEE(6, i).Value)
        '    i = i + 1
        'Loop


    End Sub

    Public Sub CargarGraficasObjetivos()

        Dim col, fil, i As Integer
        col = Form1.dgv_TendenciaOEE.Rows.Count
        On Error Resume Next
        Form1.chr_OEE.Series("Objetivo").Points.Clear()
        Form1.chr_Waste.Series("Objetivo").Points.Clear()
        Form1.chr_MDH.Series("Objetivo").Points.Clear()

        Do While i <= col
            Form1.chr_OEE.Series("Objetivo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_Objetivos(0, 0).Value)
            i = i + 1
        Loop
        i = 0
        Do While i <= col
            Form1.chr_Waste.Series("Objetivo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_Objetivos(1, 0).Value)
            i = i + 1
        Loop
        i = 0
        Do While i <= col
            Form1.chr_MDH.Series("Objetivo").Points.AddXY(Form1.dgv_TendenciaOEE(0, i).Value, Form1.dgv_Objetivos(2, 0).Value)
            i = i + 1
        Loop

    End Sub
End Module
