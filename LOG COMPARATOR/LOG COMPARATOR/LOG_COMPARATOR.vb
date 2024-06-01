Imports System.ComponentModel
Imports System.Data.Common
Imports System.IO
Imports System.Threading

Public Class LOG_COMPARATOR
    Private WithEvents backgroundWorker1 As New BackgroundWorker()
    Private WithEvents backgroundWorker2 As New BackgroundWorker()
    Private WithEvents backgroundWorker3 As New BackgroundWorker()
    Dim FinishedWork As Integer = 0
    Private Sub LOG_COMPARATOR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button3.Visible = False
        Dim ProgrammeName = FuncLib.GetProgrammeName(TextBox1.Text, Configs.TextBox3.Text, Configs.TextBox4.Text)
        If ProgrammeName.Item1 IsNot Nothing Then
            Dim FoundMasterFile As String = FuncLib.SearchForMasterFile(Configs.TextBox1.Text, ProgrammeName.Item1)
            TextBox2.Text = FoundMasterFile
        End If
        DataGridView1.Columns.Add("RowIndex", "Row Index")
        DataGridView1.Columns.Add("ColumnIndex", "Column Index")
        DataGridView1.Columns.Add("ExpectedValue", "Expected Value")
        DataGridView1.Columns.Add("ActualValue", "Actual Value")
        backgroundWorker1.WorkerSupportsCancellation = True
        backgroundWorker1.RunWorkerAsync()
        backgroundWorker2.WorkerSupportsCancellation = True
        backgroundWorker2.RunWorkerAsync()
        backgroundWorker3.WorkerSupportsCancellation = True
        backgroundWorker3.RunWorkerAsync()
    End Sub
    Private Sub backgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles backgroundWorker1.DoWork
        Invoke(Sub()
                   ListBox1.Items.Clear()
                   Button1.BackColor = Color.Yellow
                   Dim Col1FromLog As List(Of String) = FuncLib.LoadValuesFromCSVcol1(TextBox1.Text, 1, 1)
                   Dim Col1FromMaster As List(Of String) = FuncLib.LoadValuesFromCSVcol1(TextBox2.Text, 1, 1)
                   Dim setFromLog As New HashSet(Of String)(Col1FromLog)
                   Dim setFromMaster As New HashSet(Of String)(Col1FromMaster)
                   Dim differences As New List(Of Tuple(Of Integer, Integer, String, String))()
                   ListBox1.Items.Clear()
                   For rowIndex As Integer = 0 To Math.Max(Col1FromLog.Count, Col1FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col1FromLog.Count, Col1FromLog(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col1FromMaster.Count, Col1FromMaster(rowIndex), "")
                       ListBox1.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 1, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For rowIndex As Integer = 0 To Math.Max(Col1FromLog.Count, Col1FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col1FromMaster.Count, Col1FromMaster(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col1FromLog.Count, Col1FromLog(rowIndex), "")
                       ListBox1.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 1, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For Each diff As Tuple(Of Integer, Integer, String, String) In differences
                       Dim rowIndex As Integer = diff.Item1 + 1
                       Dim columnIndex As Integer = diff.Item2
                       Dim expectedValue As String = diff.Item3
                       Dim actualValue As String = diff.Item4
                       Dim existingRowIndex As Integer = -1
                       For Each row As DataGridViewRow In DataGridView1.Rows
                           If row.Cells("RowIndex").Value IsNot Nothing AndAlso CInt(row.Cells("RowIndex").Value) = rowIndex AndAlso row.Cells("ColumnIndex").Value IsNot Nothing AndAlso CInt(row.Cells("ColumnIndex").Value) = columnIndex Then
                               existingRowIndex = row.Index
                               Exit For
                           End If
                       Next
                       If existingRowIndex <> -1 Then
                           DataGridView1.Rows(existingRowIndex).Cells("ExpectedValue").Value = expectedValue
                           DataGridView1.Rows(existingRowIndex).Cells("ActualValue").Value = actualValue
                       Else
                           DataGridView1.Rows.Add(rowIndex, columnIndex, expectedValue, actualValue)
                       End If
                   Next
                   If differences.Count = 0 Then
                       Button1.BackColor = Color.SeaGreen
                   Else
                       Button1.BackColor = Color.Crimson
                   End If
               End Sub)
    End Sub
    Private Sub backgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        FinishedWork += 1
    End Sub
    Private Sub backgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles backgroundWorker2.DoWork
        Invoke(Sub()
                   ListBox2.Items.Clear()
                   Button2.BackColor = Color.Yellow
                   Dim Col2FromLog As List(Of String) = FuncLib.LoadValuesFromCSVcol2(TextBox1.Text, 1, 2)
                   Dim Col2FromMaster As List(Of String) = FuncLib.LoadValuesFromCSVcol2(TextBox2.Text, 1, 2)
                   Dim setFromLog As New HashSet(Of String)(Col2FromLog)
                   Dim setFromMaster As New HashSet(Of String)(Col2FromMaster)
                   Dim differences As New List(Of Tuple(Of Integer, Integer, String, String))()
                   For rowIndex As Integer = 0 To Math.Max(Col2FromLog.Count, Col2FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col2FromLog.Count, Col2FromLog(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col2FromMaster.Count, Col2FromMaster(rowIndex), "")
                       ListBox2.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 2, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For rowIndex As Integer = 0 To Math.Max(Col2FromLog.Count, Col2FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col2FromMaster.Count, Col2FromMaster(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col2FromLog.Count, Col2FromLog(rowIndex), "")
                       ListBox2.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 2, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For Each diff As Tuple(Of Integer, Integer, String, String) In differences
                       Dim rowIndex As Integer = diff.Item1 + 1
                       Dim columnIndex As Integer = diff.Item2
                       Dim expectedValue As String = diff.Item3
                       Dim actualValue As String = diff.Item4
                       Dim existingRowIndex As Integer = -1
                       For Each row As DataGridViewRow In DataGridView1.Rows
                           If row.Cells("RowIndex").Value IsNot Nothing AndAlso CInt(row.Cells("RowIndex").Value) = rowIndex AndAlso row.Cells("ColumnIndex").Value IsNot Nothing AndAlso CInt(row.Cells("ColumnIndex").Value) = columnIndex Then
                               existingRowIndex = row.Index
                               Exit For
                           End If
                       Next
                       If existingRowIndex <> -1 Then
                           DataGridView1.Rows(existingRowIndex).Cells("ExpectedValue").Value = expectedValue
                           DataGridView1.Rows(existingRowIndex).Cells("ActualValue").Value = actualValue
                       Else
                           DataGridView1.Rows.Add(rowIndex, columnIndex, expectedValue, actualValue)
                       End If
                   Next
                   If differences.Count = 0 Then
                       Button2.BackColor = Color.SeaGreen
                   Else
                       Button2.BackColor = Color.Crimson
                   End If
               End Sub)
    End Sub
    Private Sub backgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles backgroundWorker2.RunWorkerCompleted
        FinishedWork += 1
    End Sub
    Private Sub backgroundWorker3_DoWork(sender As Object, e As DoWorkEventArgs) Handles backgroundWorker3.DoWork
        Invoke(Sub()
                   ListBox3.Items.Clear()
                   Button4.BackColor = Color.Yellow
                   Dim Col3FromLog As List(Of String) = FuncLib.LoadValuesFromCSVcol3(TextBox1.Text, 1, 3)
                   Dim Col3FromMaster As List(Of String) = FuncLib.LoadValuesFromCSVcol3(TextBox2.Text, 1, 3)
                   Dim setFromLog As New HashSet(Of String)(Col3FromLog)
                   Dim setFromMaster As New HashSet(Of String)(Col3FromMaster)
                   Dim differences As New List(Of Tuple(Of Integer, Integer, String, String))()
                   For rowIndex As Integer = 0 To Math.Max(Col3FromLog.Count, Col3FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col3FromLog.Count, Col3FromLog(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col3FromMaster.Count, Col3FromMaster(rowIndex), "")
                       ListBox3.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 3, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For rowIndex As Integer = 0 To Math.Max(Col3FromLog.Count, Col3FromMaster.Count) - 1
                       Dim expectedValue As String = If(rowIndex < Col3FromMaster.Count, Col3FromMaster(rowIndex), "")
                       Dim actualValue As String = If(rowIndex < Col3FromLog.Count, Col3FromLog(rowIndex), "")
                       ListBox3.Items.Add($"Compairing {expectedValue} with {actualValue}")
                       If expectedValue <> actualValue Then
                           differences.Add(Tuple.Create(rowIndex, 3, expectedValue, actualValue))
                           Exit For
                       End If
                   Next
                   For Each diff As Tuple(Of Integer, Integer, String, String) In differences
                       Dim rowIndex As Integer = diff.Item1 + 1
                       Dim columnIndex As Integer = diff.Item2
                       Dim expectedValue As String = diff.Item3
                       Dim actualValue As String = diff.Item4
                       Dim existingRowIndex As Integer = -1
                       For Each row As DataGridViewRow In DataGridView1.Rows
                           If row.Cells("RowIndex").Value IsNot Nothing AndAlso CInt(row.Cells("RowIndex").Value) = rowIndex AndAlso row.Cells("ColumnIndex").Value IsNot Nothing AndAlso CInt(row.Cells("ColumnIndex").Value) = columnIndex Then
                               existingRowIndex = row.Index
                               Exit For
                           End If
                       Next
                       If existingRowIndex <> -1 Then
                           DataGridView1.Rows(existingRowIndex).Cells("ExpectedValue").Value = expectedValue
                           DataGridView1.Rows(existingRowIndex).Cells("ActualValue").Value = actualValue
                       Else
                           DataGridView1.Rows.Add(rowIndex, columnIndex, expectedValue, actualValue)
                       End If
                   Next
                   If differences.Count = 0 Then
                       Button4.BackColor = Color.SeaGreen
                   Else
                       Button4.BackColor = Color.Crimson
                   End If
               End Sub)
    End Sub
    Private Sub backgroundWorker3_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles backgroundWorker3.RunWorkerCompleted
        FinishedWork += 1
        Thread.Sleep(3000)
        If FinishedWork = 3 Then
            'FuncLib.SendMail()
        End If
    End Sub

End Class
