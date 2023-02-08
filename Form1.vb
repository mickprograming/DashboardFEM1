Public Class Form1
    'Paleta de colores Arquitectura gótica
    '#6C6B74 GRIS
    '#2E303E  AZUL OSCURO   46,48,62
    '#9199BE  MORADO CLARO  
    '#54678F  AZUL CLARO
    '#212624  NEGRO CLARO
    'https://fontawesome.com/search?o=r&s=thin  consultar nombre de iconos en Font Aweson.com
    'Instalar Plugin FontAwosen.Sharp para los iconos
    'Logo Principal https://looka.com/s/109377868
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fecha As Date
        fecha = Today
        Select Case fecha.Month
            Case 1
                Me.cmb_Mes.Text = "Enero"
            Case 2
                Me.cmb_Mes.Text = "Febrero"
            Case 3
                Me.cmb_Mes.Text = "Marzo"
            Case 4
                Me.cmb_Mes.Text = "Abril"
            Case 5
                Me.cmb_Mes.Text = "Mayo"
            Case 6
                Me.cmb_Mes.Text = "Junio"
            Case 7
                Me.cmb_Mes.Text = "Julio"
            Case 8
                Me.cmb_Mes.Text = "Agosto"
            Case 9
                Me.cmb_Mes.Text = "Septiembre"
            Case 10
                Me.cmb_Mes.Text = "Octubre"
            Case 11
                Me.cmb_Mes.Text = "Noviembre"
            Case 12
                Me.cmb_Mes.Text = "Diciembre"
        End Select
        Me.cmb_Año.Text = fecha.Year

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub IconButton2_Click(sender As Object, e As EventArgs) Handles IconButton2.Click
        mdl_Consultas.CargarObjetivos()
        mdl_Consultas.CargarGrid()
        mdl_Consultas.Promedios()


    End Sub


    Private Sub dgv_TendenciaOEE_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_TendenciaOEE.CellContentClick

    End Sub

    Private Sub IconButton3_Click(sender As Object, e As EventArgs) Handles IconButton3.Click
        mdl_Consultas.Promedios()
    End Sub
End Class
