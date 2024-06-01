Imports System.Net
Imports MadMilkman.Ini
Imports System.Configuration
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Globalization
Imports OfficeOpenXml
Imports System.IO
Imports System.Diagnostics.Eventing
Imports Utilities
Imports System.Runtime.Remoting
Imports OfficeOpenXml.FormulaParsing.Excel.Functions
Imports System.Net.Mail
Module FuncLib
    Dim Inipath As String = ConfigurationManager.AppSettings("IniFile")
    Dim ini As New IniFile
    Public Function LoadIniFile() As String
        Try
            ini.Load(Inipath)
            Configs.TextBox1.Text = ini.Sections("BASIC").Keys("Master_File").Value
            Configs.TextBox2.Text = ini.Sections("BASIC").Keys("Report_Path").Value
            Configs.TextBox3.Text = ini.Sections("Programme_Name").Keys("Row").Value
            Configs.TextBox4.Text = ini.Sections("Programme_Name").Keys("Column").Value
            Configs.TextBox5.Text = ini.Sections("BASIC").Keys("Algorithm_Perct").Value
            Configs.ComboBox1.Items.Clear()
            Configs.ComboBox1.Text = Nothing
            Configs.ComboBox1.Items.Add("Automatic")
            Configs.ComboBox1.Items.Add("Manual")
            If ini.Sections("BASIC").Keys("Run_Mode").Value = "Automatic" Then
                Configs.ComboBox1.SelectedItem = Configs.ComboBox1.Items(0)
            ElseIf ini.Sections("BASIC").Keys("Run_Mode").Value = "Manual" Then
                Configs.ComboBox1.SelectedItem = Configs.ComboBox1.Items(1)
            End If
            Return "True"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function
    Public Function GetProgrammeName(filePath As String, row As Integer, column As Integer) As Tuple(Of String, String)
        Dim cellValue As String = String.Empty
        Try
            Using reader As New StreamReader(filePath)
                For i As Integer = 1 To row
                    Dim line As String = reader.ReadLine()
                    If i = row Then
                        Dim values As String() = line.Split(","c)
                        If column >= 1 AndAlso column <= values.Length Then
                            Dim a1 As String = values(column - 1).TrimStart()
                            Dim a2 As String = a1.TrimEnd
                            Dim a3 As String = a2.Trim
                            cellValue = a3
                            Exit For
                        End If
                    End If
                Next
            End Using
        Catch ex As Exception
            Return Tuple.Create(Of String, String)(Nothing, ex.Message)
        End Try
        Return Tuple.Create(cellValue, String.Empty)
    End Function
    Public Function SearchForMasterFile(folderPath As String, fileName As String) As String
        Dim filePath As String = String.Empty
        Dim maxMatchPercentage As Integer = Configs.TextBox5.Text
        Try
            Dim files As String() = Directory.GetFiles(folderPath)
            For Each file As String In files
                Dim nameWithoutExtension As String = Path.GetFileNameWithoutExtension(file)
                Dim matchPercentage As Integer = CalculateMatchPercentage(nameWithoutExtension, fileName)
                If matchPercentage >= maxMatchPercentage Then
                    filePath = file
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try
        Return filePath
    End Function
    Private Function CalculateMatchPercentage(str1 As String, str2 As String) As Integer
        Dim matchCount As Integer = 0
        Dim maxLength As Integer = Math.Max(str1.Length, str2.Length)
        For i As Integer = 0 To Math.Min(str1.Length, str2.Length) - 1
            If str1(i) = str2(i) Then
                matchCount += 1
            End If
        Next
        Return CInt((matchCount / maxLength) * 100)
    End Function
    Public Function LoadValuesFromCSVcol1(filePath As String, startRow As Integer, column As Integer) As List(Of String)
        Dim cellValues As New List(Of String)()

        Try
            Using reader As New StreamReader(filePath)
                Dim currentRow As Integer = 0
                While Not reader.EndOfStream
                    currentRow += 1
                    Dim line As String = reader.ReadLine()

                    If currentRow >= startRow Then
                        Dim values As String() = line.Split(","c)
                        If column >= 1 AndAlso column <= values.Length Then
                            cellValues.Add(values(column - 1).Trim()) ' Read from specified column
                        End If
                    End If
                End While
            End Using
        Catch ex As Exception
            ' Handle exceptions if needed
        End Try

        Return cellValues
    End Function
    Public Function LoadValuesFromCSVcol2(filePath As String, startRow As Integer, column As Integer) As List(Of String)
        Dim cellValues As New List(Of String)()

        Try
            Using reader As New StreamReader(filePath)
                Dim currentRow As Integer = 0
                While Not reader.EndOfStream
                    currentRow += 1
                    Dim line As String = reader.ReadLine()

                    If currentRow >= startRow Then
                        Dim values As String() = line.Split(","c)
                        If column >= 1 AndAlso column <= values.Length Then
                            cellValues.Add(values(column - 1).Trim()) ' Read from specified column
                        End If
                    End If
                End While
            End Using
        Catch ex As Exception
            ' Handle exceptions if needed
        End Try

        Return cellValues
    End Function
    Public Function LoadValuesFromCSVcol3(filePath As String, startRow As Integer, column As Integer) As List(Of String)
        Dim cellValues As New List(Of String)()

        Try
            Using reader As New StreamReader(filePath)
                Dim currentRow As Integer = 0
                While Not reader.EndOfStream
                    currentRow += 1
                    Dim line As String = reader.ReadLine()

                    If currentRow >= startRow Then
                        Dim values As String() = line.Split(","c)
                        If column >= 1 AndAlso column <= values.Length Then
                            cellValues.Add(values(column - 1).Trim()) ' Read from specified column
                        End If
                    End If
                End While
            End Using
        Catch ex As Exception
            ' Handle exceptions if needed
        End Try

        Return cellValues
    End Function
    Public Function SendMail()
        Dim smtpServer As New SmtpClient("smtp.gmail.com")
        smtpServer.Port = 587 ' Gmail SMTP port
        smtpServer.Credentials = New System.Net.NetworkCredential("souravbera1097@gmail.com", "berasourav936@gmail.com")
        smtpServer.EnableSsl = True ' Enable SSL

        ' Create a new email message
        Dim mail As New MailMessage()

        ' Add sender
        mail.From = New MailAddress("souravbera1097@gmail.com")

        ' Add recipients (multiple addresses)
        mail.To.Add("berasourav936@gmail.com")
        ' Add more recipients if needed

        ' Subject and body of the email
        mail.Subject = "Your subject"
        mail.Body = "Your message body"

        ' Send the email
        Try
            smtpServer.Send(mail)
            MsgBox("success")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
End Module
