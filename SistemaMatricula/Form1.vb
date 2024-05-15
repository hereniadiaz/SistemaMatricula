Imports System.Data.OleDb
Public Class Form1
    Private Sub btnConectar_Click(sender As Object, e As EventArgs) Handles btnConectar.Click
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.Oledb.4.0; Data Source=G:\1 CLASES 2024\PROGRAMACION 12C\II Parcial\Sistema Matricula 2024\Base de Datos\matricula(2002-2003).mdb")
        Try
            con.Open()
            MsgBox("conectado")

            Dim query = "Select * from estudiante"
            Dim adap As New OleDbDataAdapter(query, con)
            Dim dt As New DataTable
            adap.Fill(dt)
            DataGridView1.DataSource = dt

        Catch ex As Exception
            MsgBox("No se conectó por: " & ex.Message)
        End Try
    End Sub
End Class
