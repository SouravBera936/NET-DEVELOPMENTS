Public Class EmpDet
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Txt File (*.txt)|*.txt|All Files (*.*)|*.*"
        saveFileDialog.FilterIndex = 1
        saveFileDialog.RestoreDirectory = True
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim fileName As String = saveFileDialog.FileName
            Using sw As New System.IO.StreamWriter(fileName)
                For Each item As Object In ListBox1.Items
                    sw.WriteLine(item.ToString())
                Next
            End Using
        End If
    End Sub
End Class