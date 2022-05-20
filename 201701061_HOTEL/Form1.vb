Public Class Form1


    Private Sub LimpiarVectorresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LimpiarVectorresToolStripMenuItem.Click
        DataGridView1.Rows.Clear()

        vector = 0

        For I = 0 To 6

            nombre(vector) = Nothing
            nit(vector) = Nothing
            cantpersonas(vector) = Nothing
            tipo_habitacion(vector) = Nothing
            estandar(vector) = Nothing
            AC(vector) = Nothing
            jacuzzi(vector) = Nothing
            pagoparcial(vector) = Nothing
            pagototal(vector) = Nothing

        Next I
    End Sub

    Private Sub EliminarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminarToolStripMenuItem.Click
        Dim I As Byte

        nombre(vector) = Nothing
        nit(vector) = Nothing
        cantpersonas(vector) = Nothing
        tipo_habitacion(vector) = Nothing
        estandar(vector) = Nothing
        AC(vector) = Nothing
        jacuzzi(vector) = Nothing
        pagoparcial(vector) = Nothing
        pagototal(vector) = Nothing

        For I = vector To 6

            nombre(I) = nombre(I + 1)
            nit(I) = nit(I + 1)
            cantpersonas(I) = cantpersonas(I + 1)
            tipo_habitacion(I) = tipo_habitacion(I + 1)
            estandar(I) = estandar(vector)
            AC(I) = AC(I + 1)
            jacuzzi(I) = jacuzzi(I + 1)
            pagoparcial(I) = pagoparcial(I + 1)
            pagoparcial(I) = pagoparcial(I + 1)
        Next I
        MsgBox("Registro Eliminado exitosamente")

        nombre(vector) = Nothing
        nit(vector) = Nothing
        cantpersonas(vector) = Nothing
        tipo_habitacion(vector) = Nothing
        estandar(vector) = Nothing
        AC(vector) = Nothing
        jacuzzi(vector) = Nothing
        pagoparcial(vector) = Nothing
        pagototal(vector) = Nothing
        limpiar_entradas()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub AceptarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AceptarToolStripMenuItem.Click
        If vector <= 6 Then
            If Val(TextBox1.Text <> "") Then
                nombre(vector) = Val(TextBox1.Text)
            Else
                MsgBox("Error, ingresar nombre")
                TextBox1.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox2.Text) And Val(TextBox2.Text <> "")) Then
                nit(vector) = Val(TextBox2.Text)
            Else
                MsgBox("Error, ingresar NIT")
                TextBox2.Focus() : Exit Sub
            End If
            If (IsNumeric(TextBox3.Text) And Val(TextBox3.Text <> "")) Then
                cantpersonas(vector) = Val(TextBox3.Text)
            Else
                MsgBox("Ingresar la cantidad de personas se hospedan")
                TextBox2.Focus() : Exit Sub
            End If


            cantpersonas(vector) = TextBox2.Text


            Select Case ComboBox1.SelectedIndex
                Case 0
                    estandar(vector) = costoestandar
                Case 1
                    AC(vector) = costoAC
                Case 2
                    jacuzzi(vector) = Costojacuzzi

                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select

            tipo_habitacion(vector) = ComboBox1.Text


            Select Case ComboBox1.SelectedIndex
                Case 0
                    pagototal(vector) = costoestandar
                Case 1
                    pagototal(vector) = costoAC
                Case 2
                    pagototal(vector) = Costojacuzzi
                Case Else
                    MsgBox("Error, no se seleciono ningun tipo")
                    TextBox1.Focus()
            End Select

        End If
    End Sub

    Private Sub MostrarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MostrarToolStripMenuItem.Click
        mostrarvectores()
    End Sub

    Private Sub ConsultarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarToolStripMenuItem.Click
        Dim existe As Boolean = False

        Dim I As Byte
        I = 0
        While (I <= 6) And Not (existe)

            If (nombre(I) = Val(TextBox1.Text)) Then
                existe = True
            Else
                I = I + 1
            End If
        End While

        If (existe) Then
            MsgBox("Registro Encontrado exitosamente")

            TextBox1.Text = Str(nombre(I))
            TextBox2.Text = Str(nit(I))
            TextBox3.Text = Str(cantpersonas(I))
            ComboBox1.Text = tipo_habitacion(I)

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(nombre(I), (nit(I)), (cantpersonas(I)), (tipo_habitacion(I), (pagoparcial(I)), (pagototal(I))))
            vector = I
        Else
            MsgBox("Dato no encontrado")
            TextBox1.Focus()
        End If
    End Sub

    Private Sub ActualizarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarToolStripMenuItem.Click
        nombre(vector) = Val(TextBox1.Text)
        tipo_habitacion(vector) = ComboBox1.Text
        nit(vector) = Val(TextBox2.Text)
        cantpersonas(vector) = Val(TextBox3.Text)



        MsgBox("Registro actualizado correctamente")

        vector = 6
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox(" ¿ Desea salir del programa ? ", vbQuestion + vbYesNo, "salir") = vbYes Then
            End
        End If

    End Sub
End Class
