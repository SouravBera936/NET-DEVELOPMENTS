Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.CompilerServices
Public Class FolderWatcher
    Private watcher As FileSystemWatcher
    Private Sub FolderWatcher_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        System.Threading.Thread.Sleep(2000)
        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ListBox1.Invoke(Sub()
                            Try
                                ListBox1.Items.Add("Loading Confriguations.....")
                                Dim Loadini As String = FuncLib.LoadIniFile()
                                If Loadini = "True" Then
                                    ListBox1.Items.Add($"Ini File Loading Status :{Loadini}")
                                Else
                                    ListBox1.Items.Add($"Error Occured While Loading Confriguration File :{Loadini}")
                                    Dim iret As Object = MsgBox($"Error Occured While Loading Confriguration File.{Environment.NewLine}Error Code:{Loadini}", vbCritical + vbRetryCancel, "FuncLib.LoadIniFile()")
                                    If iret = vbRetry Then
                                        Application.Restart()
                                    ElseIf iret = vbCancel Then
                                        Application.Exit()
                                    End If
                                End If
                            Catch ex As Exception
                                ListBox1.Items.Add($"Error Occured While Loading Confriguration File:{ex.Message}")
                                Dim iret As Object = MsgBox($"Error Occured While Loading Confriguration File.{Environment.NewLine}Error Code:{ex.Message}", vbCritical + vbRetryCancel, "BackgroundWorker1_DoWork()")
                                If iret = vbRetry Then
                                    Application.Restart()
                                ElseIf iret = vbCancel Then
                                    Application.Exit()
                                End If
                            End Try
                        End Sub)
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            ListBox1.Items.Add($"Error Occured While Executing BackgroundWorker1_RunWorkerCompleted():{e.Error}")
            Dim iret As Object = MsgBox($"Error Occured While Loading Confriguration File.{Environment.NewLine}Error Code:{e.Error}", vbCritical + vbOKOnly, "BackgroundWorker1_RunWorkerCompleted()")
            If iret = vbOK Then
                ListBox1.Invoke(Sub()
                                    Application.Exit()
                                End Sub)
            End If
        Else
            BackgroundWorker2.WorkerSupportsCancellation = True
            BackgroundWorker2.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        ListBox1.Invoke(Sub()
                            Try
                                ListBox1.Items.Add($"Starting Foldeer Watcher For : {Configs.TextBox2.Text.ToString}")
                                watcher = New FileSystemWatcher()
                                watcher.Path = Configs.TextBox2.Text.ToString
                                watcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite
                                watcher.Filter = "*.*"
                                AddHandler watcher.Created, AddressOf OnCreated
                                watcher.EnableRaisingEvents = True
                            Catch ex As Exception
                                ListBox1.Items.Add($"Error Occured While Starting Folder Watcher:{ex.Message}")
                                Dim iret As Object = MsgBox($"Error Occured While Starting Folder Watcher :{ex.Message}", vbRetryCancel + vbCritical, "Folder Watcher")
                                If iret = vbRetry Then
                                    Application.Restart()
                                ElseIf iret = vbCancel Then
                                    Application.Exit()
                                End If
                            End Try
                        End Sub)
    End Sub
    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            ListBox1.Items.Add($"Error Occured While Executing Backgroundorker2_Dowork():{e.Error}")
            Dim iret As Object = MsgBox($"Error Occured While Executing Backgroundorker2_Dowork():{e.Error}", vbOKOnly + vbCritical, "Backgroundworker2_RunworkerCompleted")
            If iret = vbOK Then
                Application.Exit()
            End If
        Else
            Invoke(Sub()
                       ListBox1.Items.Add($"Folder Watcher Started At :{DateTime.Now.ToString("G")}")
                       Me.Hide()
                   End Sub)
        End If
    End Sub
    Private Sub OnCreated(source As Object, e As FileSystemEventArgs)
        ListBox1.Invoke(Sub()
                            Dim fileExtension As String = Path.GetExtension(e.FullPath).ToLower()
                            If fileExtension = ".csv" OrElse fileExtension = ".xlsx" OrElse fileExtension = ".xls" Then
                                ListBox1.Items.Add($"New Log Found :{e.FullPath}")
                                LOG_COMPARATOR.TextBox1.Text = e.FullPath
                                LOG_COMPARATOR.TextBox1.Enabled = False
                                LOG_COMPARATOR.TextBox2.Clear()
                                LOG_COMPARATOR.TextBox2.Enabled = False
                                LOG_COMPARATOR.ShowDialog()
                            End If
                        End Sub)
    End Sub
    Private Sub FolderWatcher_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        watcher.Dispose()
    End Sub
End Class