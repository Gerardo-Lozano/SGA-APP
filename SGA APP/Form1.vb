Imports System.Data.SqlClient

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Conexion_SQL_Retention()

        Try

            'Leer tabla de proveedores

            Dim consulta_batch As String
            Dim lista_batch As Byte

            consulta_batch = "SELECT *  FROM [GHS].[dbo].[SGA] WHERE Codigo LIKE '%" & TextBox6.Text & "%' OR Nombre LIKE '%" & TextBox6.Text & "%'"

            adaptador_a_sql_Muestras = New SqlDataAdapter(consulta_batch, Conexion_SQL.ConexionSQL_Muestras)
            Conexion_SQL.registro_a_sql_Muestras = New DataSet
            adaptador_a_sql_Muestras.Fill(registro_a_sql_Muestras, "Tabla1")
            lista_batch = registro_a_sql_Muestras.Tables("Tabla1").Rows.Count
            DataGridView1.DataSource = registro_a_sql_Muestras.Tables("Tabla1")
        Catch

        End Try

        Button7.Visible = False
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index

        TextBox1.Text = DataGridView1.Item(0, i).Value().ToString
        TextBox2.Text = DataGridView1.Item(1, i).Value().ToString
        TextBox3.Text = DataGridView1.Item(3, i).Value().ToString
        TextBox4.Text = DataGridView1.Item(4, i).Value().ToString
        TextBox5.Text = DataGridView1.Item(2, i).Value().ToString
        ComboBox1.Text = DataGridView1.Item(5, i).Value().ToString
        ComboBox2.Text = DataGridView1.Item(6, i).Value().ToString
        ComboBox3.Text = DataGridView1.Item(7, i).Value().ToString
        ComboBox4.Text = DataGridView1.Item(8, i).Value().ToString

        Button5.Visible = True
        Button7.Visible = True
        Button2.Visible = False
        Button3.Visible = False

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        Button3.Visible = False
        Button5.Visible = True
        Button2.Visible = False

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button5.Visible = False
        Button7.Visible = False
        Button2.Visible = False
        Button3.Visible = True
        Button4.Visible = False

        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        TextBox1.Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text.Trim = String.Empty Or TextBox2.Text.Trim = String.Empty Then
            MsgBox("Por favor capture mínimo los campos de No. Item y Descripción", MsgBoxStyle.OkOnly)
            Return
        End If

        'Pregunta si quiere grabar
        Dim strMsg As String
        Dim iResponse As Integer

        ' Texto en el cuadro de pregunta 
        strMsg = "Desea agregar este código?" & Chr(10)
        strMsg = strMsg & "Haga Click en Si para agregar o NO para cancelar"
        ' Mensaje de pregutna. 
        iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Agregar codigo")
        ' Checar respuesta 
        If iResponse = vbNo Then
            'Undo the change. 
            'DoCmd.RunCommand acCmdUndo 
            Return
            'Cancel the update. 
            'Cancel = True
        End If

        Conexion_SQL_Retention()

        Try

            Dim Query As String

            Query = "INSERT INTO [GHS].[dbo].[SGA] ([Codigo], [Nombre ], [Palabra de advertencia], [Indicador de riesgo], [Indicador de precaucion ], [Simbolo1], [Simbolo2], [Simbolo3] ,Simbolo4) VALUES
                   ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox5.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & ComboBox4.Text & "')"


            Dim comando_save As SqlCommand
            comando_save = New SqlCommand(Query, Conexion_SQL.ConexionSQL_Muestras)
            comando_save.ExecuteNonQuery()

            MsgBox("Agregado correctamente")
        Catch ex As Exception
            MsgBox("Error, No se grabo la información" & ex.Message)

        End Try


        Button5.Visible = True
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = True

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False

        Button1_Click(sender, e)

        ConexionSQL_Muestras.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text.Trim = String.Empty Or TextBox2.Text.Trim = String.Empty Then
            MsgBox("Por favor capture mínimo los campos de No. Item y Descripción", MsgBoxStyle.OkOnly)
            Return
        End If

        'Pregunta si quiere grabar
        Dim strMsg As String
        Dim iResponse As Integer

        ' Texto en el cuadro de pregunta 
        strMsg = "Desea actualizar la información de este código?" & Chr(10)
        strMsg = strMsg & "Haga Click en Si para modificar o NO para cancelar"
        ' Mensaje de pregutna. 
        iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Modificar información del código")
        ' Checar respuesta 
        If iResponse = vbNo Then
            'Undo the change. 
            'DoCmd.RunCommand acCmdUndo 
            Return
            'Cancel the update. 
            'Cancel = True
        End If

        Conexion_SQL_Retention()

        Try

            Dim Query As String

            Query = "UPDATE [GHS].[dbo].[SGA] " &
                "SET [Codigo] = '" & TextBox1.Text & "',
                [Nombre ] = '" & TextBox2.Text & "',
                [Palabra de advertencia] = '" & TextBox5.Text & "',
                [Indicador de riesgo] = '" & TextBox3.Text & "',
                [Indicador de precaucion ] = '" & TextBox4.Text & "',
                [Simbolo1] = '" & ComboBox1.Text & "',
                [Simbolo2] = '" & ComboBox2.Text & "',
                [Simbolo3] = '" & ComboBox3.Text & "',
                [Simbolo4] = '" & ComboBox4.Text & "'
                WHERE [Codigo] = '" & TextBox1.Text & "'"

            Dim comando_save As SqlCommand
            comando_save = New SqlCommand(Query, Conexion_SQL.ConexionSQL_Muestras)
            comando_save.ExecuteNonQuery()

            MsgBox("Actualizado correctamente")
        Catch ex As Exception
            MsgBox("Error, No se grabo la información" & ex.Message)
        End Try

        Button5.Visible = True
        Button7.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = True

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False

        Button1_Click(sender, e)

        ConexionSQL_Muestras.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        If TextBox1.Text.Trim = String.Empty Then
            MsgBox("Por favor seleccione un Codigo a eliminar", MsgBoxStyle.OkOnly)
            Return
        End If

        'Pregunta si quiere grabar
        Dim strMsg As String
        Dim iResponse As Integer

        ' Texto en el cuadro de pregunta 
        strMsg = "Desea eliminar este codigo?" & Chr(10)
        strMsg = strMsg & "Haga Click en Si para eliminar o NO para cancelar"
        ' Mensaje de pregutna. 
        iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Eliminar codigo")
        ' Checar respuesta 
        If iResponse = vbNo Then
            'Undo the change. 
            'DoCmd.RunCommand acCmdUndo 
            Return
            'Cancel the update. 
            'Cancel = True
        End If


        Conexion_SQL_Retention()

        Try

            Dim Query As String

            Query = "DELETE FROM [GHS].[dbo].[SGA] WHERE [Codigo] = '" & TextBox1.Text & "'"

            Dim comando_save As SqlCommand
            comando_save = New SqlCommand(Query, Conexion_SQL.ConexionSQL_Muestras)
            comando_save.ExecuteNonQuery()

            MsgBox("Eliminado correctamente")
        Catch ex As Exception
            MsgBox("Error, No se grabo la información" & ex.Message)

        End Try


        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False

        Button1_Click(sender, e)

        ConexionSQL_Muestras.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True

        TextBox2.Select()

        Button7.Visible = False
        Button3.Visible = False
        Button2.Visible = True
        Button4.Visible = False

    End Sub
End Class
