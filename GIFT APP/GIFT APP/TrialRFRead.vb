Imports System.IO.Ports
Public Class TrialRFRead
    Private WithEvents serialPort As New SerialPort()
    Private Sub TrialRFRead_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        serialPort.PortName = "COM5"
        serialPort.BaudRate = 9600
        Try
            serialPort.Open()
            Console.WriteLine($"Serial port {serialPort.PortName} opened successfully.")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim button As Button = DirectCast(sender, Button)
        If button.Text = "Start Reading" Then
            serialPort.WriteLine("START") ' Send any command to start reading
            button.Text = "Stop Reading"
            TextBox4.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox1.Clear()
        Else
            serialPort.WriteLine("STOP")
            button.Text = "Start Reading"
        End If
    End Sub
    Private Sub SerialPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        serialPort.DiscardInBuffer()
        Dim hexData As String = serialPort.ReadLine()
        Dim extractedString As String = hexData.Substring(1, hexData.Length - 2)
        TextBox1.Invoke(Sub() TextBox1.Text = extractedString.ToString())
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim decimalValue As Long = Convert.ToInt64(TextBox1.Text.ToString, 16)
        TextBox4.Text = decimalValue
        Dim hexadecimalString As String = Convert.ToString(decimalValue, 16)
        Dim last8Digits As String = hexadecimalString.Substring(hexadecimalString.Length - 8)
        Dim EsdCardNumber As Integer = Convert.ToInt32(last8Digits, 16)
        Dim EsdCardNumberString As String = EsdCardNumber.ToString("D10")
        TextBox2.Text = EsdCardNumberString
        Dim last4digit As String = hexadecimalString.Substring(hexadecimalString.Length - 4)
        Dim CardNumber As Integer = Convert.ToInt32(last4digit, 16)
        Dim CardNumberstring As String = CardNumber.ToString("D5")
        TextBox3.Text = CardNumberstring
    End Sub
End Class