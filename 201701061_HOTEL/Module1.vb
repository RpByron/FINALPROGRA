Module Module1
    Public Const costoestandar = 250
    Public Const costoAC = 290
    Public Const Costojacuzzi = 370

    Public tipo_habitacion(7)
    Public estandar(7)
    Public AC(7)
    Public jacuzzi(7)
    Public nombre(7)
    Public nit(7)
    Public cantpersonas(7)
    Public pagoparcial(7)
    Public pagototal(7)
    Public habitacion(7)


    Public vector As Byte = 0

    Sub mostrarvectores()
        Form1.DataGridView1.Rows.Clear()
        For I = 0 To 6
            If (nombre(I) <> Nothing) Then
                Form1.DataGridView1.Rows.Add(nombre(I), (nit(I)), (cantpersonas(I)), (habitacion(I), (pagoparcial(I)), (pagototal(I))))
            Else

                Continue For
            End If
        Next

    End Sub

    Sub limpiar_entradas()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.TextBox3.Clear()
        Form1.ComboBox1.Text = ""
        Form1.TextBox1.Focus()
    End Sub
End Module
