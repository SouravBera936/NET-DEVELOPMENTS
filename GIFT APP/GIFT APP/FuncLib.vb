Imports System.Net
Imports MadMilkman.Ini
Imports System.Data.OleDb
Imports System.Configuration
Imports Bunifu.Framework.UI
Imports ADOX
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Globalization
Imports Guna.UI2.WinForms.Suite
Imports OfficeOpenXml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Bunifu.UI.WinForms
Imports System.Diagnostics.Eventing
Imports Utilities
Imports Org.BouncyCastle.Crypto.Engines
Imports System.Runtime.Remoting
Imports OfficeOpenXml.FormulaParsing.Excel.Functions
Module FuncLib
    Dim Inipath As String = ConfigurationManager.AppSettings("IniFile")
    Dim ini As New IniFile
    Dim hostName As String = Dns.GetHostName()
    Dim ValidUser As Boolean = False
    Public Function LoadIniFile() As String
        Try
            ini.Load(Inipath)
            Return Nothing
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function
    Public Class WriteLog
        Public Shared Sub WriteErrorLog(ErrMsg As String)
            ini.Load(Inipath)
            If ini.Sections("BASIC").Keys("ErrorLoging_State").Value = "True" Then
                Dim path As String = ini.Sections("BASIC").Keys("Error_Log").Value
                Using file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                    file.WriteLine($"{DateTime.Now} ON {hostName} {ErrMsg}")
                End Using
            Else
                'donothing
            End If
        End Sub
        Public Shared Sub WriteAppLog(AppLog As String)
            ini.Load(Inipath)
            If ini.Sections("BASIC").Keys("AppLoging_State").Value = "True" Then
                Dim path As String = ini.Sections("BASIC").Keys("Application_Log").Value
                Using file As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                    file.WriteLine($"{DateTime.Now} ON {hostName} {AppLog}")
                End Using
            Else
                'donothing
            End If
        End Sub
    End Class
    Public Class DataBaseOperations
        Public Shared Function InitializeDB() As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    If connection.State = ConnectionState.Open Then
                        Return Nothing
                    Else
                        Return "Error Connecting To Database!"
                    End If
                End Using
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function CheckUserValidation(UserName As String) As String
            ini.Load(Inipath)
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT * FROM Employee_DB WHERE (NT_ID = @UserName)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@UserName", UserName)
                        Using reader As OleDbDataReader = command.ExecuteReader()
                            If reader.Read() Then
                                MainWindow.ListBox1.Items.Add(reader("Employee_ID"))
                                MainWindow.ListBox1.Items.Add(reader("NT_ID"))
                                MainWindow.ListBox1.Items.Add(reader("Scanner_5D"))
                                MainWindow.ListBox1.Items.Add(reader("Scanner_ID"))
                                MainWindow.ListBox1.Items.Add(reader("Employee_Name"))
                                MainWindow.ListBox1.Items.Add(reader("Gender"))
                                MainWindow.ListBox1.Items.Add(reader("Celebration_Date"))
                                MainWindow.ListBox1.Items.Add(reader("Buisness_Area"))
                                MainWindow.ListBox1.Items.Add(reader("Functions"))
                                MainWindow.ListBox1.Items.Add(reader("Leave_Approver"))
                                MainWindow.ListBox1.Items.Add(reader("Created_By"))
                                MainWindow.ListBox1.Items.Add(reader("Created_At"))
                                FuncLib.EnhanceFeatures.IconbuildafterLogin(True, reader("Employee_Name"))
                                ValidUser = True
                            Else
                                ValidUser = False
                            End If
                        End Using
                    End Using
                    Dim query1 As String = "SELECT * FROM User_DB WHERE (NT_ID = @UserName)"
                    Using command1 As New OleDbCommand(query1, connection)
                        command1.Parameters.AddWithValue("@UserName", UserName)
                        Using reader As OleDbDataReader = command1.ExecuteReader()
                            If reader.HasRows() Then
                                reader.Read()
                                MainWindow.ListBox1.Items.Add(reader("Modified_By"))
                                MainWindow.ListBox1.Items.Add(reader("Modified_At"))
                                MainWindow.ListBox1.Items.Add(reader("Issue_Gift"))
                                MainWindow.ListBox1.Items.Add(reader("Events_Addition"))
                                MainWindow.ListBox1.Items.Add(reader("Admn_Access"))
                                reader.Close()
                            Else
                                MainWindow.ListBox1.Items.Add("NA")
                                MainWindow.ListBox1.Items.Add("NA")
                                MainWindow.ListBox1.Items.Add("False")
                                MainWindow.ListBox1.Items.Add("False")
                                MainWindow.ListBox1.Items.Add("False")
                                reader.Close()
                            End If
                        End Using
                    End Using
                End Using
                Return ValidUser.ToString()
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function AddToFunctionDB(DB As String, Employee_ID As String, NT_ID As String, Scanner_ID As String, Scanner_5D As String, Employee_Name As String, Gender As String, BirthDate As Date, Buisness_Area As String, Functions As String, Leave_Approver As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"INSERT INTO {DB} (Employee_ID,NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Created_At)" &
                                       "VALUES (@Employee_ID,@NT_ID,@Scanner_ID,@Scanner_5D,@Employee_Name,@Gender,@Birth_Date,@Buisness_Area,@Functions,@Leave_Approver,@Created_At)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        command.Parameters.AddWithValue("@NT_ID", NT_ID)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        command.Parameters.AddWithValue("@Scanner_5D", Scanner_5D)
                        command.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        command.Parameters.AddWithValue("@Gender", Gender)
                        command.Parameters.AddWithValue("@Birth_Date", BirthDate)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Functions", Functions)
                        command.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        command.Parameters.AddWithValue("@Created_At", DateTime.Now.ToString("G"))
                        command.ExecuteNonQuery()
                    End Using
                End Using
                Return Nothing
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function CheckifDataAlreadyExist(DB As String, Employee_ID As String, NT_ID As String, Connection As OleDbConnection) As String
            Try
                Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID AND NT_ID = @NT_ID) ORDER BY Created_At DESC"
                Using command As New OleDbCommand(query, Connection)
                    command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                    command.Parameters.AddWithValue("@NT_ID", NT_ID)
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        While reader.Read
                            Dim StartDate As Date = reader("Created_At")
                            Dim EndDate As Date = DateTime.Now.ToString("G")
                            Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(StartDate, EndDate)
                            If daysDifference <= Convert.ToInt32(ini.Sections("REPEATSPANTIME").Keys(DB).Value) Then
                                Return "True"
                            Else
                                Return "False"
                            End If
                        End While
                    Else
                        Return "False"
                    End If
                End Using
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function CheckIFGiftShared(DB As String, Employee_ID As String, NT_ID As String, Connection As OleDbConnection) As String
            Try
                Dim query As String = $"SELECT * FROM Gift_History WHERE (Employee_ID = @Employee_ID AND NT_ID = @NT_ID AND Gift_Type= @Gift_Type) ORDER BY Shared_At DESC"
                Using command As New OleDbCommand(query, Connection)
                    command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                    command.Parameters.AddWithValue("@NT_ID", NT_ID)
                    command.Parameters.AddWithValue("@Gift_Type", DB)
                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        While reader.Read
                            Dim StartDate As Date = reader("Shared_At")
                            Dim EndDate As Date = DateTime.Now.ToString("G")
                            Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(StartDate, EndDate)
                            If daysDifference <= Convert.ToInt32(ini.Sections("REPEATSPANTIME").Keys(DB).Value) Then
                                Return "True"
                            Else
                                Return "False"
                            End If
                        End While
                    Else
                        Return "False"
                    End If
                End Using
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
    End Class
    Public Class Gifts
        Public Shared Function GetOcaasioList() As List(Of String)
            Dim result As New List(Of String)()
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Dim tables As DataTable = New DataTable()
                Dim excludedTables As New List(Of String) From {"MSys", "Employee_DB", "Gift_History", "User_DB"}
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    tables = connection.GetSchema("Tables")
                End Using
                For Each row As DataRow In tables.Rows
                    Dim tableName As String = row("TABLE_NAME").ToString()
                    If Not (excludedTables.Any(Function(t) tableName.StartsWith(t))) Then
                        result.Add(tableName)
                    End If
                Next
                Return result
            Catch ex As Exception
                Return New List(Of String) From {ex.Message}
            End Try
        End Function
        Public Shared Function CheckTypeOfOcassion(Occasion As String) As String
            Dim OccassionType As String
            Try
                ini.Load(Inipath)
                OccassionType = ini.Sections("TYPE").Keys(Occasion).Value
                If OccassionType IsNot Nothing Then
                    If OccassionType = "Constant" Then
                        Return OccassionType
                    ElseIf OccassionType = "Variable" Then
                        Return OccassionType
                    End If
                Else
                    Return "Undefined Type"
                End If
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function VarableDateQuery(Occasion As String) As List(Of Date)
            Dim dates As List(Of DateTime) = New List(Of DateTime)()
            Try
                ini.Load(Inipath)
                Dim dateSection As IniSection = ini.Sections("DATE")
                If dateSection IsNot Nothing Then
                    For Each key As IniKey In dateSection.Keys
                        If key.Name = Occasion Then
                            For Each dateValue As String In key.Value.Split(","c)
                                dates.Add(DateTime.ParseExact(dateValue, "MM/dd/yyyy", Nothing))
                            Next
                        End If
                    Next
                    Return dates
                End If
            Catch ex As Exception
                dates.Clear()
                dates.Add(DateTime.MinValue)
            End Try
        End Function
        Public Shared Function CheckEmployeeEligibilityForConstantFunctions(DB As String, Employee_ID As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = '{Employee_ID}')"
                    Using command As New OleDbCommand(query, connection)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference1 <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RecordEmployeeGiftDistributionForConstantFunction(DB As String, Employee_ID As String, NT_ID As String, Employee_Name As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Celebration_Date As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        Dim count As Integer = CInt(command.ExecuteScalar())
                        If count > 1 Then
                            Dim deleteQuery As String = $"DELETE FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = @Employee_ID)"
                            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                deleteCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                                deleteCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                    Dim insertQuery As String = $"INSERT INTO Gift_History (Employee_ID,NT_ID,Employee_Name,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Employee_Name,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Gift_Type,@GivenBy_ID,@GivenBy_NT,@GivenBy_Name,@Shared_At)"
                    Using insertCommand As New OleDbCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Gift_Type", DB)
                        insertCommand.Parameters.AddWithValue("@GivenBy_ID", MainWindow.ListBox1.Items(0).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_NT", MainWindow.ListBox1.Items(1).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_Name", MainWindow.ListBox1.Items(4).ToString)
                        insertCommand.Parameters.AddWithValue("@Shared_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CheckEmployeeEligibilityForVariableFunctions(DB As String, Employee_ID As String, Celebration_Date As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = @Employee_ID AND Celebration_Date=@Celebration_Date)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        command.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("G")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference1 <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RecordEmployeeGiftDistributionForVariableFunction(DB As String, Employee_ID As String, NT_ID As String, Employee_Name As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Celebration_Date As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID) AND (Celebration_Date = @Celebration_Date)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        command.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        Dim count As Integer = CInt(command.ExecuteScalar())
                        If count > 1 Then
                            Dim deleteQuery As String = $"DELETE FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = @Employee_ID AND Celebration_Date = @Celebration_Date)"
                            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                deleteCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                                deleteCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                                deleteCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                    Dim insertQuery As String = $"INSERT INTO Gift_History (Employee_ID,NT_ID,Employee_Name,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Employee_Name,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Gift_Type,@GivenBy_ID,@GivenBy_NT,@GivenBy_Name,@Shared_At)"
                    Using insertCommand As New OleDbCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Gift_Type", DB)
                        insertCommand.Parameters.AddWithValue("@GivenBy_ID", MainWindow.ListBox1.Items(0).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_NT", MainWindow.ListBox1.Items(1).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_Name", MainWindow.ListBox1.Items(4).ToString)
                        insertCommand.Parameters.AddWithValue("@Shared_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function FetchSharedGiftsHistory(Employee_ID As String, NT_ID As String) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID, NT_ID, Employee_Name, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Gift_Type, Shared_At FROM Gift_History WHERE (GivenBy_ID = '{Employee_ID}') AND (GivenBy_NT = '{NT_ID}')"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function CheckEmployeeEligibilityForBulkDistributionConstant(DB As String, Scanner_ID As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Scanner_ID = @Scanner_ID)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("G")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference1 <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RecordEmployeeEligibilityForBulkDistributionConstant(DB As String, Employee_ID As String, NT_ID As String, Employee_Name As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Celebration_Date As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        Dim count As Integer = CInt(command.ExecuteScalar())
                        If count > 0 Then
                            Dim deleteQuery As String = $"DELETE FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = @Employee_ID)"
                            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                deleteCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                                deleteCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                    Dim insertQuery As String = $"INSERT INTO Gift_History (Employee_ID,NT_ID,Employee_Name,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Employee_Name,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Gift_Type,@GivenBy_ID,@GivenBy_NT,@GivenBy_Name,@Shared_At)"
                    Using insertCommand As New OleDbCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Gift_Type", DB)
                        insertCommand.Parameters.AddWithValue("@GivenBy_ID", MainWindow.ListBox1.Items(0).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_NT", MainWindow.ListBox1.Items(1).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_Name", MainWindow.ListBox1.Items(4).ToString)
                        insertCommand.Parameters.AddWithValue("@Shared_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CheckEmployeeEligibilityForBulkDistributionvariable(DB As String, Scanner_ID As String, Celebration_Date As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Scanner_ID = @Scanner_ID AND Celebration_Date=@Celebration_Date)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        command.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RecordEmployeeEligibilityForBulkDistributionvariable(DB As String, Employee_ID As String, NT_ID As String, Employee_Name As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Celebration_Date As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID) AND (Celebration_Date = @Celebration_Date)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        command.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        Dim count As Integer = CInt(command.ExecuteScalar())
                        If count > 0 Then
                            Dim deleteQuery As String = $"DELETE FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Employee_ID = @Employee_ID AND Celebration_Date = @Celebration_Date)"
                            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                deleteCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                                deleteCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                                deleteCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                    Dim insertQuery As String = $"INSERT INTO Gift_History (Employee_ID,NT_ID,Employee_Name,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Employee_Name,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Gift_Type,@GivenBy_ID,@GivenBy_NT,@GivenBy_Name,@Shared_At)"
                    Using insertCommand As New OleDbCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Gift_Type", DB)
                        insertCommand.Parameters.AddWithValue("@GivenBy_ID", MainWindow.ListBox1.Items(0).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_NT", MainWindow.ListBox1.Items(1).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_Name", MainWindow.ListBox1.Items(4).ToString)
                        insertCommand.Parameters.AddWithValue("@Shared_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function IdentifyEmployee(RFIDNumber As String, Method As String) As String
            If RFIDNumber IsNot Nothing Then
                Dim decimalValue As Long = Convert.ToInt64(RFIDNumber, 16)
                Dim hexadecimalString As String = Convert.ToString(decimalValue, 16)
                Dim last8Digits As String = hexadecimalString.Substring(hexadecimalString.Length - 8)
                Dim EsdCardNumber As Integer = Convert.ToInt32(last8Digits, 16)
                Dim EsdCardNumberString As String = EsdCardNumber.ToString("D10")
                Dim last4digit As String = hexadecimalString.Substring(hexadecimalString.Length - 4)
                Dim CardNumber As Integer = Convert.ToInt32(last4digit, 16)
                Dim CardNumberstring As String = CardNumber.ToString("D5")
                If Method = "BY 10D" Then
                    Return EsdCardNumberString
                ElseIf Method = "BY 5D" Then
                    Return CardNumberstring
                End If
            End If
        End Function
        Public Shared Function ConvertESDNumberToCardNumber(ESDNumber As Integer) As String
            Dim hexadecimalString As String = Convert.ToString(ESDNumber, 16)
            Dim last4digit As String = hexadecimalString.Substring(hexadecimalString.Length - 4)
            Dim CardNumber As Integer = Convert.ToInt32(last4digit, 16)
            Dim CardNumberstring As String = CardNumber.ToString("D5")
            Return CardNumberstring
        End Function
        Public Shared Function CheckEmployeeEligibilityForConstantBy5D(DB As String, Employee_ID As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Scanner_5D = '{Employee_ID}')"
                    Using command As New OleDbCommand(query, connection)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CheckEmployeeEligibilityForvariableBy5D(DB As String, Employee_ID As String, Celebration_Date As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE ID = (SELECT MAX(ID) FROM {DB} WHERE Scanner_5D = @Employee_ID AND Celebration_Date=@Celebration_Date)"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        command.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Celebration_Date")
                                Dim CreationDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference1 As Integer = EnhanceFeatures.CalculateDateDifferenceForEligibility(StartDate, EndDate)
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(CreationDate, EndDate)
                                If daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) AndAlso daysDifference <= Convert.ToInt32(ini.Sections("ELIGIBILITY").Keys(DB).Value) Then
                                    If reader("Employee_ID") = MainWindow.ListBox1.Items(0) Then
                                        Return "Same"
                                        Exit Function
                                    End If
                                    MainWindow.BunifuTextBox2.Text = reader("Employee_ID")
                                    MainWindow.BunifuTextBox3.Text = reader("NT_ID")
                                    MainWindow.BunifuTextBox4.Text = reader("Employee_Name")
                                    MainWindow.BunifuTextBox5.Text = reader("Buisness_Area")
                                    MainWindow.BunifuTextBox6.Text = reader("Functions")
                                    MainWindow.BunifuTextBox7.Text = reader("Leave_Approver")
                                    MainWindow.BunifuTextBox8.Text = reader("Celebration_Date")
                                    Return "True"
                                Else
                                    Return "TLE"
                                End If
                            End While
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
    End Class
    Public Class Events
        Public Shared Function GetVariableocassions() As List(Of String)
            ini.Load(Inipath)
            Dim variableOccasions As New List(Of String)()
            Dim typeSection As IniSection = ini.Sections("TYPE")
            If typeSection IsNot Nothing Then
                For Each key As IniKey In typeSection.Keys
                    If key.Value = "Variable" Then
                        variableOccasions.Add(key.Name)
                    End If
                Next
            End If
            Return variableOccasions
        End Function
        Public Shared Function CompareDateMatcOfFunction(Ocassion As String, Ocassion_Date As DateTime) As String
            Try
                ini.Load(Inipath)
                Dim dateSection As IniSection = ini.Sections("DATE")
                If dateSection IsNot Nothing Then
                    Dim dates As List(Of DateTime) = New List(Of DateTime)()
                    For Each key As IniKey In dateSection.Keys
                        If key.Name = Ocassion Then
                            For Each dateValue As String In key.Value.Split(","c)
                                dates.Add(DateTime.ParseExact(dateValue, "MM/dd/yyyy", Nothing))
                            Next
                        End If
                    Next
                    Dim selectedDate As DateTime = Ocassion_Date.Date
                    If dates.Contains(selectedDate) Then
                        Return "True"
                    Else
                        Return "False"
                    End If
                Else
                    Return "False"
                End If
            Catch ex As Exception
                Return ex.Message
            End Try

        End Function
        Public Shared Function addFunctionDateToFunction(Ocassion As String, NewDate As DateTime) As String
            Try
                Dim iniData As New IniFile()
                iniData.Load(Inipath)
                Dim dateSection As IniSection = iniData.Sections("DATE")
                If dateSection IsNot Nothing Then
                    Dim key As IniKey = dateSection.Keys(Ocassion)
                    If key IsNot Nothing Then
                        Dim existingDates As String = key.Value.TrimEnd(","c)
                        Dim formattedDate As String = NewDate.ToString("MM/dd/yyyy")
                        If Not String.IsNullOrEmpty(existingDates) Then
                            existingDates += ","
                        End If
                        existingDates += formattedDate
                        key.Value = existingDates
                        iniData.Save(Inipath)
                    End If
                End If
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function FetchInternalDbForNewVariableOccasion() As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID,NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Buisness_Area,Functions,Leave_Approver FROM Employee_DB"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function ImportDataIntoVariableFunction(DB As String, Employee_ID As String, NT_ID As String, Scanner_5D As String, Scanner_ID As String, Employee_Name As String, Gender As String, Celebration_Date As Date, Buisness_Area As String, Functions As String, Leave_Approver As String, Created_At As DateTime) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"INSERT INTO {DB} (Employee_ID,NT_ID,Scanner_5D,Scanner_ID,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Created_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Scanner_5D,@Scanner_ID,@Employee_Name,@Gender,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Created_At)"
                    Using insertCommand As New OleDbCommand(query, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Scanner_5D", Scanner_5D)
                        insertCommand.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Gender", Gender)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Created_At", Created_At)
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function VerifyExcel(FilePath As String, requiredColumns As List(Of String)) As String
            Try
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                Using package As New ExcelPackage(New System.IO.FileInfo(FilePath))
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                    For Each cell As ExcelRangeBase In worksheet.Cells
                        cell.Value = cell.Text
                    Next
                    Dim headerRow As ExcelRangeBase = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                    Dim columnNames As New List(Of String)
                    For Each cell As ExcelRangeBase In headerRow
                        columnNames.Add(cell.Text.Trim())
                    Next
                    Dim extraColumns As New List(Of String)
                    Dim PresentColumns As New List(Of String)
                    For Each columnName As String In columnNames
                        If Not requiredColumns.Contains(columnName) Then
                            extraColumns.Add(columnName)
                        Else
                            PresentColumns.Add(columnName)
                        End If
                    Next
                    If extraColumns.Count > 0 OrElse PresentColumns.Count <> requiredColumns.Count Then
                        Return "False"
                    End If
                    Return "True"
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function Saveasstringformatting(FilePath As String, columnNames As List(Of String)) As Boolean
            Try
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                Using package As New ExcelPackage(New System.IO.FileInfo(FilePath))
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                    For Each columnName As String In columnNames
                        Dim columnIndex As Integer = GetColumnIndex(worksheet, columnName)
                        If columnIndex <> -1 Then
                            For Each cell As ExcelRangeBase In worksheet.Cells(2, columnIndex, worksheet.Dimension.End.Row, columnIndex)
                                cell.Style.Numberformat.Format = "@"
                            Next
                        End If
                    Next
                    package.Save()
                End Using
                Return True
            Catch ex As Exception
                Return ex.Message
            End Try

        End Function
        Private Shared Function GetColumnIndex(worksheet As ExcelWorksheet, columnName As String) As Integer
            For col As Integer = 1 To worksheet.Dimension.End.Column
                If worksheet.Cells(1, col).Text.Trim().Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                    Return col
                End If
            Next
            Return -1
        End Function
        Public Shared Function VerifyNullDataforevent(ByVal filePath As String, Optional ByVal sheetName As String = Nothing) As String
            Try
                Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;HDR=YES;'"
                Dim dataTable As New DataTable()
                Using connection As New OleDbConnection(connectionString)
                    If String.IsNullOrEmpty(sheetName) Then
                        connection.Open()
                        Dim dtSheets As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                        sheetName = dtSheets.Rows(0)("TABLE_NAME").ToString()
                        connection.Close()
                    End If
                    Dim query As String = $"SELECT [Emp No],[Display Name],[Gender],[Birth Date],[Business Area],[Function],[10 Digit Card Number] FROM [{sheetName}]"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(dataTable)
                    End Using
                End Using
                For Each row As DataRow In dataTable.Rows
                    For Each columnName In {"Emp No", "Display Name", "Gender", "Birth Date", "Business Area", "Function", "10 Digit Card Number"}
                        Dim cellValue As Object = row(columnName)
                        If cellValue Is Nothing OrElse IsDBNull(cellValue) OrElse String.IsNullOrEmpty(cellValue.ToString().Trim()) Then
                            Return "False"
                        End If
                    Next
                Next
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function VerifyDuplicateColumnsForEvent(ByVal filePath As String, Celebration_Date As Date, Optional ByVal sheetName As String = Nothing) As Tuple(Of DataTable, String)
            Try
                Dim totalRowCount As Integer = 0
                Dim duplicateRowCount As Integer = 0
                Dim importedRowCount As Integer = 0
                Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;HDR=YES;'"
                Dim dataTable As New DataTable()
                Using connection As New OleDbConnection(connectionString)
                    If String.IsNullOrEmpty(sheetName) Then
                        connection.Open()
                        Dim dtSheets As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                        sheetName = dtSheets.Rows(0)("TABLE_NAME").ToString()
                        connection.Close()
                    End If
                    Dim query As String = $"SELECT * FROM [{sheetName}]"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(dataTable)
                    End Using
                    Dim dict As New Dictionary(Of String, DataRow)()
                    Dim uniqueDataTable As DataTable = dataTable.Clone()
                    Dim latestDuplicatesDataTable As DataTable = dataTable.Clone()
                    For Each row As DataRow In dataTable.Rows
                        Dim key1 As String = $"{row.Item(0)}"
                        Dim key3 As String = $"{row.Item(2)}"
                        If dict.ContainsKey(key1) OrElse dict.ContainsKey(key3) Then
                            latestDuplicatesDataTable.ImportRow(row)
                        Else
                            uniqueDataTable.ImportRow(row)
                            dict(key1) = row
                            dict(key3) = row
                        End If
                    Next
                    Dim mergedDataTable As DataTable = uniqueDataTable.Clone()
                    For Each column As DataColumn In latestDuplicatesDataTable.Columns
                        If Not mergedDataTable.Columns.Contains(column.ColumnName) Then
                            mergedDataTable.Columns.Add(column.ColumnName, column.DataType)
                        End If
                    Next
                    For Each row As DataRow In uniqueDataTable.Rows
                        mergedDataTable.ImportRow(row)
                    Next
                    For Each row As DataRow In latestDuplicatesDataTable.Rows
                        Dim key1 As String = $"{row.Item(0)}"
                        Dim key3 As String = $"{row.Item(2)}"
                        If Not mergedDataTable.AsEnumerable().Any(Function(r) $"{r("Employee_ID")}" = key1 OrElse
                                                                  $"{r("Scanner_ID")}" = key3) Then
                            mergedDataTable.ImportRow(row)
                        End If
                    Next
                    Dim additionalColumnNames As String() = {"Scanner_5D", "Celebration_Date", "Created_At"}
                    For Each columnName As String In additionalColumnNames
                        If Not mergedDataTable.Columns.Contains(columnName) Then
                            Select Case columnName
                                Case "Scanner_5D"
                                    mergedDataTable.Columns.Add(columnName, GetType(String))
                                Case "Celebration_Date", "Created_At"
                                    mergedDataTable.Columns.Add(columnName, GetType(DateTime))
                            End Select
                        End If
                    Next
                    For Each row As DataRow In mergedDataTable.Rows
                        row("Scanner_5D") = FuncLib.Gifts.ConvertESDNumberToCardNumber(row("Scanner_ID"))
                        row("Celebration_Date") = Celebration_Date
                        row("Created_At") = DateTime.Now.ToString("G")
                    Next
                    Return Tuple.Create(mergedDataTable, "")
                End Using
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function ImplementDataToGrid(FilePath As String, CelebrationDate As Date) As Tuple(Of DataTable, String)
            Using package As New ExcelPackage(New System.IO.FileInfo(FilePath))
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                Dim dt As New DataTable()
                dt.Columns.Add("Employee_ID", GetType(String))
                dt.Columns.Add("NT_ID", GetType(String))
                dt.Columns.Add("Scanner_ID", GetType(String))
                dt.Columns.Add("Scanner_5D", GetType(String))
                dt.Columns.Add("Employee_Name", GetType(String))
                dt.Columns.Add("Gender", GetType(String))
                dt.Columns.Add("Celebration_Date", GetType(Date))
                dt.Columns.Add("Buisness_Area", GetType(String))
                dt.Columns.Add("Functions", GetType(String))
                dt.Columns.Add("Leave_Approver", GetType(String))
                dt.Columns.Add("Created_At", GetType(DateTime))
                For row As Integer = 2 To worksheet.Dimension.End.Row
                    Dim Employee_ID As String = worksheet.Cells(row, 1).Text
                    Dim NT_ID As String = worksheet.Cells(row, 2).Text
                    Dim Scanner_ID As String = worksheet.Cells(row, 3).Text
                    Dim Scanner_5D As String = FuncLib.Gifts.ConvertESDNumberToCardNumber(Scanner_ID)
                    Dim Employee_Name As String = worksheet.Cells(row, 4).Text
                    Dim Gender As String = worksheet.Cells(row, 5).Text
                    Dim Celebration_Date As Date = CelebrationDate
                    Dim Buisness_Area As String = worksheet.Cells(row, 6).Text
                    Dim Functions As String = worksheet.Cells(row, 7).Text
                    Dim Leave_Approver As String = worksheet.Cells(row, 8).Text
                    Dim Created_At As String = DateTime.Now.ToString("G")
                    dt.Rows.Add(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                Next
                Return Tuple.Create(dt, "")
            End Using
        End Function
        Public Shared Function ImportDataIntoVariableFromDG(DB As String, datas As DataGridView) As Tuple(Of String, Integer)
            Try
                ini.Load(Inipath)
                Dim Count As Integer = 0
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    For Each row As DataGridViewRow In datas.Rows
                        If Not row.IsNewRow Then
                            Dim query As String = $"INSERT INTO {DB} (Employee_ID,NT_ID,Scanner_5D,Scanner_ID,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Created_At)" &
                                                "VALUES (@Employee_ID,@NT_ID,@Scanner_5D,@Scanner_ID,@Employee_Name,@Gender,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Created_At)"
                            Using insertCommand As New OleDbCommand(query, connection)
                                insertCommand.Parameters.AddWithValue("@Employee_ID", row.Cells("EMP ID").Value)
                                insertCommand.Parameters.AddWithValue("@NT_ID", row.Cells("NT ID").Value)
                                insertCommand.Parameters.AddWithValue("@Scanner_5D", row.Cells("SCAN 5D").Value)
                                insertCommand.Parameters.AddWithValue("@Scanner_ID", row.Cells("SCAN ID").Value)
                                insertCommand.Parameters.AddWithValue("@Employee_Name", row.Cells("NAME").Value)
                                insertCommand.Parameters.AddWithValue("@Gender", row.Cells("GENDER").Value)
                                insertCommand.Parameters.AddWithValue("@Celebration_Date", row.Cells("DATE").Value)
                                insertCommand.Parameters.AddWithValue("@Buisness_Area", row.Cells("BUISNESS").Value)
                                insertCommand.Parameters.AddWithValue("@Function", row.Cells("DEPARTMENT").Value)
                                insertCommand.Parameters.AddWithValue("@Leave_Approver", row.Cells("REPORTING MANAGER").Value)
                                insertCommand.Parameters.AddWithValue("@Created_At", row.Cells("TIME STAMP").Value)
                                insertCommand.ExecuteNonQuery()
                                Count += 1
                            End Using
                        End If
                    Next
                End Using
                Return Tuple.Create("True", Count)
            Catch ex As Exception
                Return Tuple.Create(Of String, Integer)(ex.Message, 0)
            End Try
        End Function
        Public Shared Function CheckIfOcassionAleareadyExist(Ocassion As String) As String
            Try
                Dim BoolFlag As Boolean = False
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Dim tables As DataTable = New DataTable()
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    tables = connection.GetSchema("Tables")
                End Using
                For Each row As DataRow In tables.Rows
                    Dim tableName As String = row("TABLE_NAME").ToString()
                    If tableName = Ocassion Then
                        BoolFlag = True
                        Exit For
                    End If
                Next
                Return BoolFlag
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CreateNewOcassion(Function_title As String, Function_Type As String, Function_Date As Date) As String
            Try
                Dim iniData As New IniFile()
                iniData.Load(Inipath)
                Dim RepeatSpantime As IniSection = iniData.Sections("REPEATSPANTIME")
                Dim Types As IniSection = iniData.Sections("TYPE")
                Dim dates As IniSection = iniData.Sections("DATE")
                Dim Eligibility As IniSection = iniData.Sections("ELIGIBILITY")
                If Function_Type = "Repeat Every Year" Then
                    RepeatSpantime.Keys.Add(Function_title, 335)
                    Types.Keys.Add(Function_title, "Variable")
                    dates.Keys.Add(Function_title, Function_Date)
                    Eligibility.Keys.Add(Function_title, 30)
                ElseIf Function_Type = "Only This Time" Then
                    Types.Keys.Add(Function_title, "Constant")
                    dates.Keys.Add(Function_title, Function_Date)
                    Eligibility.Keys.Add(Function_title, 30)
                End If
                iniData.Save(Inipath)
                Return "True"
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function CreateNewoccasionTable(table_name As String) As String
            ini.Load(Inipath)
            Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
            Dim tableName As String = table_name
            Dim columns As String = "[ID] AUTOINCREMENT, " &
                        "[Employee_ID] TEXT(50), " &
                        "[NT_ID] TEXT(50), " &
                        "[Scanner_ID] TEXT(50), " &
                        "[Scanner_5D] TEXT(50), " &
                        "[Employee_Name] TEXT(50), " &
                        "[Gender] TEXT(50), " &
                        "[Celebration_Date] DATETIME, " &
                        "[Buisness_Area] TEXT(50), " &
                        "[Functions] TEXT(50), " &
                        "[Leave_Approver] TEXT(50), " &
                        "[Created_At] DATETIME"
            Dim createTableSQL As String = $"CREATE TABLE {tableName} ({columns}, PRIMARY KEY (ID))"
            Using connection As New OleDbConnection(connectionString)
                Try
                    connection.Open()
                    Using command As New OleDbCommand(createTableSQL, connection)
                        command.ExecuteNonQuery()
                    End Using
                    Return "True"
                Catch ex As Exception
                    Return ex.Message.ToString
                End Try
            End Using
        End Function
        Public Shared Function FetchAllDataForConstantFunctionDeletion(DB As String) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT  Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver,Created_At FROM {DB} "
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function DeleteDataTableForConstantFunction(DB As String)
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DROP TABLE [{DB}]"
                    Using command As New OleDbCommand(commandText, connection)
                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception

            End Try
            Dim iniData As New IniFile()
            iniData.Load(Inipath)
            RemoveKeyFromSection(iniData.Sections("TYPE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("DATE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("ELIGIBILITY"), $"{DB}")
            iniData.Save(Inipath)
        End Function
        Public Shared Function DeleteDataTableForVariableFunction(DB As String)
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DROP TABLE [{DB}]"
                    Using command As New OleDbCommand(commandText, connection)
                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception

            End Try
            Dim iniData As New IniFile()
            iniData.Load(Inipath)
            RemoveKeyFromSection(iniData.Sections("REPEATSPANTIME"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("TYPE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("DATE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("ELIGIBILITY"), $"{DB}")
            iniData.Save(Inipath)
        End Function
        Public Shared Sub RemoveKeyFromSection(section As IniSection, keyName As String)
            If section IsNot Nothing Then
                Dim key As IniKey = section.Keys(keyName)
                If key IsNot Nothing Then
                    section.Keys.Remove(key)
                End If
            End If
        End Sub
        Public Shared Function FetchAllDataForvariableDeletion(DB As String, Celebration_Date As DateTime) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Celebration_Date = #{Celebration_Date}#)"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function DeleteDataFromvariableFunction(DB As String, Celebration_Date As DateTime)
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DELETE FROM {DB} WHERE (Celebration_Date = #{Celebration_Date}#)"
                    Using command As New OleDbCommand(commandText, connection)
                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception

            End Try
            Dim iniData As New IniFile()
            iniData.Load(Inipath)
            RemoveKeyFromSection(iniData.Sections("REPEATSPANTIME"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("TYPE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("DATE"), $"{DB}")
            RemoveKeyFromSection(iniData.Sections("ELIGIBILITY"), $"{DB}")
            iniData.Save(Inipath)
        End Function
        Public Shared Function ImplementDataToGridforconstant(FilePath As String) As Tuple(Of DataTable, String)
            Using package As New ExcelPackage(New System.IO.FileInfo(FilePath))
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                Dim dt As New DataTable()
                dt.Columns.Add("Employee_ID", GetType(String))
                dt.Columns.Add("NT_ID", GetType(String))
                dt.Columns.Add("Scanner_ID", GetType(String))
                dt.Columns.Add("Scanner_5D", GetType(String))
                dt.Columns.Add("Employee_Name", GetType(String))
                dt.Columns.Add("Gender", GetType(String))
                dt.Columns.Add("Celebration_Date", GetType(Date))
                dt.Columns.Add("Buisness_Area", GetType(String))
                dt.Columns.Add("Functions", GetType(String))
                dt.Columns.Add("Leave_Approver", GetType(String))
                dt.Columns.Add("Created_At", GetType(DateTime))
                For row As Integer = 2 To worksheet.Dimension.End.Row
                    Dim Employee_ID As String = worksheet.Cells(row, 1).Text
                    Dim NT_ID As String = worksheet.Cells(row, 2).Text
                    Dim Scanner_ID As String = worksheet.Cells(row, 3).Text
                    Dim Scanner_5D As String = FuncLib.Gifts.ConvertESDNumberToCardNumber(Scanner_ID)
                    Dim Employee_Name As String = worksheet.Cells(row, 4).Text
                    Dim Gender As String = worksheet.Cells(row, 5).Text
                    Dim Celebration_Date = worksheet.Cells(row, 6).Text
                    Dim Buisness_Area As String = worksheet.Cells(row, 7).Text
                    Dim Functions As String = worksheet.Cells(row, 8).Text
                    Dim Leave_Approver As String = worksheet.Cells(row, 9).Text
                    Dim Created_At As String = DateTime.Now.ToString("G")
                    dt.Rows.Add(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                Next
                Return Tuple.Create(dt, "")
            End Using
        End Function
        Public Shared Function ImplementDataToGridforVariable(FilePath As String, CelebratioDate As Date) As Tuple(Of DataTable, String)
            Using package As New ExcelPackage(New System.IO.FileInfo(FilePath))
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                Dim dt As New DataTable()
                dt.Columns.Add("Employee_ID", GetType(String))
                dt.Columns.Add("NT_ID", GetType(String))
                dt.Columns.Add("Scanner_ID", GetType(String))
                dt.Columns.Add("Scanner_5D", GetType(String))
                dt.Columns.Add("Employee_Name", GetType(String))
                dt.Columns.Add("Gender", GetType(String))
                dt.Columns.Add("Celebration_Date", GetType(Date))
                dt.Columns.Add("Buisness_Area", GetType(String))
                dt.Columns.Add("Functions", GetType(String))
                dt.Columns.Add("Leave_Approver", GetType(String))
                dt.Columns.Add("Created_At", GetType(DateTime))
                For row As Integer = 2 To worksheet.Dimension.End.Row
                    Dim Employee_ID As String = worksheet.Cells(row, 1).Text
                    Dim NT_ID As String = worksheet.Cells(row, 2).Text
                    Dim Scanner_ID As String = worksheet.Cells(row, 3).Text
                    Dim Scanner_5D As String = FuncLib.Gifts.ConvertESDNumberToCardNumber(Scanner_ID)
                    Dim Employee_Name As String = worksheet.Cells(row, 4).Text
                    Dim Gender As String = worksheet.Cells(row, 5).Text
                    Dim Celebration_Date = CelebratioDate
                    Dim Buisness_Area As String = worksheet.Cells(row, 6).Text
                    Dim Functions As String = worksheet.Cells(row, 7).Text
                    Dim Leave_Approver As String = worksheet.Cells(row, 8).Text
                    Dim Created_At As String = DateTime.Now.ToString("G")
                    dt.Rows.Add(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                Next
                Return Tuple.Create(dt, "")
            End Using
        End Function
        Public Shared Function CheckIFEmloyeeExist(DB As String, Emp_ID As String, Scanner_ID As String) As Boolean
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Employee_ID = @Employee_ID AND Scanner_ID = @Scanner_ID) ORDER BY Created_At DESC"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Emp_ID)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Created_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(StartDate, EndDate)
                                If daysDifference <= Convert.ToInt32(ini.Sections("REPEATSPANTIME").Keys(DB).Value) Then
                                    Return True
                                Else
                                    Return False
                                End If
                            End While
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function CheckIfGiftAlreadyShared(DB As String, Emp_ID As String, Scanner_ID As String) As Boolean
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM Gift_History WHERE (Employee_ID = @Employee_ID AND Scanner_ID = @Scanner_ID AND Gift_Type= @Gift_Type) ORDER BY Shared_At DESC"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Employee_ID", Emp_ID)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        command.Parameters.AddWithValue("@Gift_Type", DB)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Dim StartDate As Date = reader("Shared_At")
                                Dim EndDate As Date = DateTime.Now.ToString("g")
                                Dim daysDifference As Integer = EnhanceFeatures.CalculateDateDifference(StartDate, EndDate)
                                If daysDifference <= Convert.ToInt32(ini.Sections("REPEATSPANTIME").Keys(DB).Value) Then
                                    Return True
                                Else
                                    Return False
                                End If
                            End While
                        End If
                    End Using
                End Using
            Catch ex As Exception

            End Try
        End Function
        Public Shared Function AddEmployeeToConstantDatabase(DB As String, Datas As DataGridView) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    For Each row As DataGridViewRow In Datas.Rows
                        If Not row.IsNewRow Then
                            Dim query As String = $"INSERT INTO {DB} (Employee_ID,NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Created_At)" &
                                                     "VALUES (@Employee_ID,@NT_ID,@Scanner_ID,@Scanner_5D,@Employee_Name,@Gender,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Created_At)"
                            Using insertCommand As New OleDbCommand(query, connection)
                                If Not CheckIFEmloyeeExist(DB, row.Cells("EMP ID").Value, row.Cells("SCAN ID").Value) AndAlso Not CheckIfGiftAlreadyShared(DB, row.Cells("EMP ID").Value, row.Cells("SCAN ID").Value) Then
                                    insertCommand.Parameters.AddWithValue("@Employee_ID", row.Cells("EMP ID").Value)
                                    If row.Cells("NT ID").Value IsNot Nothing OrElse Not IsDBNull(row.Cells("NT ID").Value) Then
                                        insertCommand.Parameters.AddWithValue("@NT_ID", row.Cells("NT ID").Value)
                                    Else
                                        insertCommand.Parameters.AddWithValue("@NT_ID", "")
                                    End If
                                    insertCommand.Parameters.AddWithValue("@Scanner_ID", row.Cells("SCAN ID").Value)
                                    insertCommand.Parameters.AddWithValue("@Scanner_5D", row.Cells("SCAN 5D").Value)
                                    insertCommand.Parameters.AddWithValue("@Employee_Name", row.Cells("NAME").Value)
                                    insertCommand.Parameters.AddWithValue("@Gender", row.Cells("GENDER").Value)
                                    insertCommand.Parameters.AddWithValue("@Celebration_Date", row.Cells("DATE").Value)
                                    insertCommand.Parameters.AddWithValue("@Buisness_Area", row.Cells("BUISNESS").Value)
                                    insertCommand.Parameters.AddWithValue("@Functions", row.Cells("DEPARTMENT").Value)
                                    insertCommand.Parameters.AddWithValue("@Leave_Approver", row.Cells("REPORTING MANAGER").Value)
                                    insertCommand.Parameters.AddWithValue("@Created_At", row.Cells("TIME STAMP").Value)
                                    insertCommand.ExecuteNonQuery()
                                End If
                            End Using
                        End If
                    Next
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RemoveSelectedEmployeeFromConstantDB(DB As String, Emp_ID As String, Scanner_ID As String, Creationdate As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DELETE FROM {DB} WHERE (Employee_ID = @Emoloyee_ID AND Scanner_ID = @Scanner_ID AND Created_At = @Created_At)"
                    Using command As New OleDbCommand(commandText, connection)
                        command.Parameters.AddWithValue("@Emoloyee_ID", Emp_ID)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        command.Parameters.AddWithValue("@Created_At", Creationdate)
                        command.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function RemoveSelectedEmployeeFromVariableFunction(DB As String, Emp_ID As String, Scanner_ID As String, Creationdate As Date, CelebrationDate As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DELETE FROM {DB} WHERE (Employee_ID = @Emoloyee_ID AND Scanner_ID = @Scanner_ID AND Created_At = @Created_At AND Celebration_Date=@Celebration_Date)"
                    Using command As New OleDbCommand(commandText, connection)
                        command.Parameters.AddWithValue("@Emoloyee_ID", Emp_ID)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        command.Parameters.AddWithValue("@Created_At", Creationdate)
                        command.Parameters.AddWithValue("@Celebration_Date", CelebrationDate)
                        command.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function DeletevariableFunctionDate(DB As String, CelebrationDate As Date) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim commandText As String = $"DELETE FROM {DB} WHERE (Celebration_Date=@Celebration_Date)"
                    Using command As New OleDbCommand(commandText, connection)
                        command.Parameters.AddWithValue("@Celebration_Date", CelebrationDate)
                        command.ExecuteNonQuery()
                    End Using
                End Using
                Dim inidata As New IniFile
                inidata.Load(Inipath)
                Dim section As IniSection = inidata.Sections("DATE")
                Dim keyvalue As String = section.Keys($"{DB}").Value
                Dim datesList As New List(Of String)(keyvalue.Split(","c))
                Dim found As Boolean = False
                For Each dateString As String In datesList
                    Dim dateValue As DateTime
                    If DateTime.TryParse(dateString, dateValue) Then
                        If dateValue = CelebrationDate.Date Then
                            datesList.Remove(dateString)
                            found = True
                            Exit For
                        End If
                    End If
                Next
                If found Then
                    Dim updatedValue As String = String.Join(",", datesList)
                    section.Keys($"{DB}").Value = updatedValue
                    inidata.Save(Inipath)
                End If
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function DeleteVariableFunctionDateFromIni(DB As String, Celebration_Date As Date) As String
            Try
                Dim inidata As New IniFile
                inidata.Load(Inipath)
                Dim section As IniSection = inidata.Sections("DATE")
                Dim keyvalue As String = section.Keys($"{DB}").Value
                Dim datesList As New List(Of String)(keyvalue.Split(","c))
                Dim found As Boolean = False
                For Each dateString As String In datesList
                    Dim dateValue As DateTime
                    If DateTime.TryParse(dateString, dateValue) Then
                        If dateValue = Celebration_Date.Date Then
                            datesList.Remove(dateString)
                            found = True
                            Exit For
                        End If
                    End If
                Next
                If found Then
                    Dim updatedValue As String = String.Join(",", datesList)
                    section.Keys($"{DB}").Value = updatedValue
                    inidata.Save(Inipath)
                End If
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CheckAndReturnEmployeeDetailsByEmployeeID(Emp_ID As String, Celebration_Date As Date) As Tuple(Of String, List(Of String), DataTable)
            Dim EmployeeFound As New DataTable()
            Dim EmployeeNotFound As New List(Of String)()
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Buisness_Area, Functions, Leave_Approver FROM Employee_DB WHERE Employee_ID='{Emp_ID}'"
                    Using command As New OleDbCommand(query, connection)
                        'command.Parameters.AddWithValue("@EmployeeID", Emp_ID)
                        Using reader As OleDbDataReader = command.ExecuteReader()
                            If reader.Read() Then
                                Using adapter As New OleDbDataAdapter(query, connection)
                                    adapter.Fill(EmployeeFound)
                                End Using
                            Else
                                EmployeeNotFound.Add(Emp_ID)
                            End If
                        End Using
                    End Using
                End Using
                Dim additionalColumnNames As String() = {"Celebration_Date", "Created_At"}
                For Each columnName As String In additionalColumnNames
                    If Not EmployeeFound.Columns.Contains(columnName) Then
                        Select Case columnName
                            Case "Celebration_Date"
                                EmployeeFound.Columns.Add(columnName, GetType(Date))
                            Case "Created_At"
                                EmployeeFound.Columns.Add(columnName, GetType(DateTime))
                        End Select
                    End If
                Next
                For Each row As DataRow In EmployeeFound.Rows
                    row("Celebration_Date") = Celebration_Date
                    row("Created_At") = DateTime.Now.ToString("G")
                Next
                Return Tuple.Create("True", EmployeeNotFound, EmployeeFound)
            Catch ex As Exception
                Return Tuple.Create(Of String, List(Of String), DataTable)(ex.Message, Nothing, Nothing)
            End Try
        End Function
    End Class

    Public Class History
        Public Shared Function FetchAllHistory() As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID,NT_ID,Employee_Name,Buisness_Area,Functions,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At FROM Gift_History"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function FilterDateWise(StartDate As Date, EndDate As Date) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            ini.Load(Inipath)
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID,NT_ID,Employee_Name,Buisness_Area,Functions,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At FROM Gift_History WHERE Shared_At >= #{StartDate}# AND Shared_At <= #{EndDate}#"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function SaveasExcel(dataGridView As DataGridView, filePath As String) As String
            Try
                Dim directoryPath As String = Path.GetDirectoryName(filePath)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                Dim sb As New System.Text.StringBuilder()
                Dim columns As Integer = dataGridView.Columns.Count
                For i As Integer = 0 To columns - 1
                    sb.Append(dataGridView.Columns(i).HeaderText)
                    If i < columns - 1 Then
                        sb.Append(",")
                    End If
                Next
                sb.AppendLine()
                For Each row As DataGridViewRow In dataGridView.Rows
                    For i As Integer = 0 To columns - 1
                        sb.Append(row.Cells(i).Value)
                        If i < columns - 1 Then
                            sb.Append(",")
                        End If
                    Next
                    sb.AppendLine()
                Next
                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                    Dim filePath1 As String = $"{filePath}.csv"
                    System.IO.File.WriteAllText(filePath, sb.ToString())
                Else
                    Dim filePath1 As String = $"{filePath}.csv"
                    System.IO.File.WriteAllText(filePath, sb.ToString())
                End If
                Return "True"
            Catch ex As Exception
                Return ex.Message.ToString
            End Try

        End Function
        Public Shared Function SaveAsPdf(dataGridView As DataGridView, filePath As String) As String
            Try
                Dim directoryPath As String = Path.GetDirectoryName(filePath)
                Dim document As New iTextSharp.text.Document()
                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                    Dim writer As PdfWriter = PdfWriter.GetInstance(document, New FileStream(filePath, FileMode.Create))
                    document.Open()
                    Dim table As New PdfPTable(dataGridView.Columns.Count)
                    For Each column As DataGridViewColumn In dataGridView.Columns
                        table.AddCell(column.HeaderText)
                    Next
                    For Each row As DataGridViewRow In dataGridView.Rows
                        For Each cell As DataGridViewCell In row.Cells
                            If cell.Value IsNot Nothing Then
                                table.AddCell(cell.Value.ToString())
                            Else
                                table.AddCell("") ' or any default value you want to use
                            End If
                        Next
                    Next
                    document.Add(table)
                    document.Close()
                Else
                    Dim writer As PdfWriter = PdfWriter.GetInstance(document, New FileStream(filePath, FileMode.Create))
                    document.Open()
                    Dim table As New PdfPTable(dataGridView.Columns.Count)
                    For Each column As DataGridViewColumn In dataGridView.Columns
                        table.AddCell(column.HeaderText)
                    Next
                    For Each row As DataGridViewRow In dataGridView.Rows
                        For Each cell As DataGridViewCell In row.Cells
                            If cell.Value IsNot Nothing Then
                                table.AddCell(cell.Value.ToString())
                            Else
                                table.AddCell("") ' or any default value you want to use
                            End If
                        Next
                    Next
                    document.Add(table)
                    document.Close()
                End If
                Return "True"
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function FetchAllPendingDataFromDBforConstant(DB As String) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID,NT_ID,Employee_Name,Buisness_Area,Functions,Created_At FROM {DB}"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function FetchAllPendingDataFromDBforVariable(DB As String, FunctioDate As Date) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT Employee_ID,NT_ID,Employee_Name,Buisness_Area,Functions,Created_At FROM {DB} WHERE Celebration_Date=#{FunctioDate}#"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function FilterDateWiseforPendingHistory(DB As String, StartDate As Date, EndDate As Date) As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM {DB} WHERE (Created_At>=#{StartDate}#) AND (Created_At<=#{EndDate}#) "
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
    End Class
    Public Class Setup
        Public Shared Function VerifyNewDataBase(Provier As String, DataSource As String, Password As String) As String
            Try
                Dim connectionString As String = $"Provider={Provier};Data Source={DataSource};Jet OLEDB:Database Password={Password}"

                ' Dictionary to store table names and their column names
                Dim requiredTables As New Dictionary(Of String, List(Of String))
                requiredTables.Add("BirthDay", New List(Of String) From {"ID", "Employee_ID", "Scanner_ID", "Scanner_5D", "Employee_Name", "Gender", "Celebration_Date", "Buisness_Area", "Functions", "Leave_Approver", "Created_At"})
                requiredTables.Add("Employee_DB", New List(Of String) From {"ID", "Employee_ID", "NT_ID", "Scanner_ID", "Scanner_5D", "Employee_Name", "Gender", "Celebration_Date", "Buisness_Area", "Functions", "Leave_Approver", "Created_By", "Created_At"})
                requiredTables.Add("Gift_History", New List(Of String) From {"ID", "Employee_ID", "NT_ID", "Employee_Name", "Celebration_Date", "Buisness_Area", "Functions", "Leave_Approver", "Gift_Type", "GivenBy_ID", "GivenBy_NT", "GivenBy_Name", "Shared_At"})
                requiredTables.Add("User_DB", New List(Of String) From {"ID", "NT_ID", "Scanner_ID", "Scanner_5D", "Employee_Name", "Gender", "Buisness_Area", "Functions", "Leave_Approver", "Modified_By", "Modified_At", "Issue_Gift", "Events_Addition", "Admn_Access"})

                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim columnExists As Boolean = False
                    For Each tables In requiredTables
                        Dim tableName As String = tables.Key
                        Dim columns As List(Of String) = tables.Value
                        For Each columnName In columns
                            Using cmd As New OleDbCommand($"SELECT TOP 1 [{columnName}] FROM [{tableName}]", connection)
                                Try
                                    cmd.ExecuteNonQuery()
                                    columnExists = True
                                Catch
                                    columnExists = False
                                    Exit For
                                End Try
                            End Using
                        Next
                    Next
                    If Not columnExists Then
                        Return "False"
                    Else
                        Return "True"
                    End If
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function ForamttingOFDataabse(Provier As String, DataSource As String, Password As String) As String
            Try
                Dim connectionString As String = $"Provider={Provier};Data Source={DataSource};Jet OLEDB:Database Password={Password};"
                Dim tablesToCreate As New Dictionary(Of String, List(Of String))
                tablesToCreate.Add("BirthDay", New List(Of String) From {"ID AUTOINCREMENT", "Employee_ID TEXT(50)", "Scanner_ID TEXT(50)", "Scanner_5D TEXT(50)", "Employee_Name TEXT(50)", "Gender TEXT(50)", "Celebration_Date DATETIME", "Buisness_Area TEXT(50)", "Functions TEXT(50)", "Leave_Approver TEXT(50)", "Created_At DATETIME"})
                tablesToCreate.Add("Employee_DB", New List(Of String) From {"ID AUTOINCREMENT", "Employee_ID TEXT(50)", "NT_ID TEXT(50)", "Scanner_ID TEXT(50)", "Scanner_5D TEXT(50)", "Employee_Name TEXT(50)", "Gender TEXT(50)", "Celebration_Date DATETIME", "Buisness_Area TEXT(50)", "Functions TEXT(50)", "Leave_Approver TEXT(50)", "Created_By TEXT(50)", "Created_At DATETIME"})
                tablesToCreate.Add("Gift_History", New List(Of String) From {"ID AUTOINCREMENT", "Employee_ID TEXT(50)", "NT_ID TEXT(50)", "Employee_Name TEXT(50)", "Celebration_Date DATETIME", "Buisness_Area TEXT(50)", "Functions TEXT(50)", "Leave_Approver TEXT(50)", "Gift_Type TEXT(50)", "GivenBy_ID TEXT(50)", "GivenBy_NT TEXT(50)", "GivenBy_Name TEXT(50)", "Shared_At DATETIME"})
                tablesToCreate.Add("User_DB", New List(Of String) From {"ID AUTOINCREMENT", "NT_ID TEXT(50)", "Scanner_ID TEXT(50)", "Scanner_5D TEXT(50)", "Employee_Name TEXT(50)", "Gender TEXT(50)", "Buisness_Area TEXT(50)", "Functions TEXT(50)", "Leave_Approver TEXT(50)", "Modified_By TEXT(50)", "Modified_At DATETIME", "Issue_Gift YESNO", "Events_Addition YESNO", "Admn_Access YESNO"})
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim dropTablesQuery As String = "DROP TABLE {0};"
                    For Each tableName In tablesToCreate.Keys
                        If TableExists(connection, tableName) Then
                            Using cmd As New OleDbCommand(String.Format(dropTablesQuery, tableName), connection)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If
                    Next
                    Dim createTablesQuery As String = "CREATE TABLE {0} ({1});"
                    For Each tables In tablesToCreate
                        Dim tableName As String = tables.Key
                        Dim columns As String = String.Join(", ", tables.Value)
                        Using cmd As New OleDbCommand(String.Format(createTablesQuery, tableName, columns), connection)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Private Shared Function TableExists(connection As OleDbConnection, tableName As String) As Boolean
            Dim schemaTable As DataTable = connection.GetSchema("Tables")
            Return schemaTable.AsEnumerable().Any(Function(row) row.Field(Of String)("TABLE_NAME") = tableName)
        End Function
        Public Shared Function CreateNewDataBase(Provier As String, DataSource As String, password As String) As String
            Try
                Dim cat As New ADOX.Catalog()
                cat.Create($"Provider={Provier};Data Source={DataSource};Jet OLEDB:Database Password={password};")
                cat = Nothing
                Dim connectionString As String = $"Provider={Provier};Data Source={DataSource};Jet OLEDB:Database Password={password};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    connection.Close()
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function ValidateAndSaveDtaabseIntoIni(Provider As String, DataSource As String, Password As String) As String
            Dim connectionString As String = $"Provider={Provider};Data Source={DataSource};Jet OLEDB:Database Password={Password}"
            Try
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    If connection.State = ConnectionState.Open Then
                        Dim Ini1 As New IniFile
                        Ini1.Load(Inipath)
                        Ini1.Sections("DATABASE").Keys("Provider").Value = Provider
                        Ini1.Sections("DATABASE").Keys("Source").Value = DataSource
                        Ini1.Sections("DATABASE").Keys("Password").Value = Password
                        Ini1.Save(Inipath)
                    Else
                        Return "False"
                    End If
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function VerifyNullData(ByVal filePath As String, ByVal sheetName As String) As String
            Try
                Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;HDR=YES;'"
                Dim dataTable As New DataTable()
                Using connection As New OleDbConnection(connectionString)
                    Dim query As String = $"SELECT Employee_ID,Scanner_ID,Employee_Name,Gender,Birth_Date,Buisness_Area,Functions,Leave_Approver FROM [{sheetName}$]"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(dataTable)
                    End Using
                End Using
                For Each row As DataRow In dataTable.Rows
                    For Each columnName In {"Employee_ID", "Scanner_ID", "Employee_Name", "Gender", "Birth_Date", "Buisness_Area", "Functions", "Leave_Approver"}
                        Dim cellValue As Object = row(columnName)
                        If cellValue Is Nothing OrElse IsDBNull(cellValue) OrElse String.IsNullOrEmpty(cellValue.ToString().Trim()) Then
                            Return "False"
                        End If
                    Next
                Next
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function Populatdatatodatagridafterextractingduplicate(ByVal filePath As String, ByVal sheetName As String) As Tuple(Of DataTable, String, Integer, Integer, Integer)
            Try
                Dim totalRowCount As Integer = 0
                Dim duplicateRowCount As Integer = 0
                Dim importedRowCount As Integer = 0
                Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;HDR=YES;'"
                Dim dataTable As New DataTable()
                Using connection As New OleDbConnection(connectionString)
                    Dim query As String = $"SELECT * FROM [{sheetName}$]"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(dataTable)
                    End Using
                End Using
                totalRowCount = dataTable.Rows.Count
                Dim dict As New Dictionary(Of String, DataRow)()
                Dim uniqueDataTable As DataTable = dataTable.Clone()
                Dim latestDuplicatesDataTable As DataTable = dataTable.Clone()
                For Each row As DataRow In dataTable.Rows
                    Dim key1 As String = $"{row.Item(0)}"
                    Dim key3 As String = $"{row.Item(2)}"
                    If dict.ContainsKey(key1) OrElse dict.ContainsKey(key3) Then
                        'latestDuplicatesDataTable.Rows.Remove(dict(key))
                        latestDuplicatesDataTable.ImportRow(row)
                        duplicateRowCount += 1
                    Else
                        uniqueDataTable.ImportRow(row)
                        dict(key1) = row
                        dict(key3) = row
                        importedRowCount += 1
                    End If
                Next
                Dim mergedDataTable As DataTable = uniqueDataTable.Clone()
                For Each column As DataColumn In latestDuplicatesDataTable.Columns
                    If Not mergedDataTable.Columns.Contains(column.ColumnName) Then
                        mergedDataTable.Columns.Add(column.ColumnName, column.DataType)
                    End If
                Next
                For Each row As DataRow In uniqueDataTable.Rows
                    mergedDataTable.ImportRow(row)
                Next
                For Each row As DataRow In latestDuplicatesDataTable.Rows
                    Dim key1 As String = $"{row.Item(0)}"
                    Dim key3 As String = $"{row.Item(2)}"
                    If Not mergedDataTable.AsEnumerable().Any(Function(r) $"{r("Employee_ID")}" = key1 OrElse
                                                              $"{r("Scanner_ID")}" = key3) Then
                        mergedDataTable.ImportRow(row)
                    End If
                Next
                Dim additionalColumnNames As String() = {"Scanner_5D", "Created_By", "Created_At"}
                For Each columnName As String In additionalColumnNames
                    If Not mergedDataTable.Columns.Contains(columnName) Then
                        Select Case columnName
                            Case "Scanner_5D", "Created_By"
                                mergedDataTable.Columns.Add(columnName, GetType(String))
                            Case "Created_At"
                                mergedDataTable.Columns.Add(columnName, GetType(DateTime))
                        End Select
                    End If
                Next
                For Each row As DataRow In mergedDataTable.Rows
                    row("Scanner_5D") = FuncLib.Gifts.ConvertESDNumberToCardNumber(row("Scanner_ID"))
                    row("Created_By") = $"{Environment.UserName}"
                    row("Created_At") = DateTime.Now.ToString("G")
                Next
                Return Tuple.Create(mergedDataTable, "", totalRowCount, duplicateRowCount, importedRowCount)
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String, Integer, Integer, Integer)(Nothing, ex.Message, 0, 0, 0)
            End Try
        End Function
        Public Shared Function CHeckIfEmployeeAlreadyExist(Emp_ID As String, Scanner_ID As String) As Boolean
            ini.Load(Inipath)
            Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
            Using connection As New OleDbConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT COUNT(*) FROM Employee_DB WHERE Employee_ID = @EmployeeID OR Scanner_ID = @ScannerID"
                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("@EmployeeID", Emp_ID)
                    command.Parameters.AddWithValue("@ScannerID", Scanner_ID)
                    Dim count As Integer = CInt(command.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        End Function
        Public Shared Function AddEmployeesToEmployeeDB(Employee_ID As String, NT_ID As String, Scanner_ID As String, Scanner_5D As String, Employee_Name As String, Gender As String, Celebration_Date As Date, Buisness_Area As String, Functions As String, Leave_Approver As String, Created_By As String, Created_At As DateTime) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"INSERT INTO Employee_DB (Employee_ID,NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver, Created_By, Created_At)" &
                                         "VALUES (@Employee_ID,@NT_ID,@Scanner_ID,@Scanner_5D,@Employee_Name,@Gender,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver, @Created_By, @Created_At)"
                    Using insertCommand As New OleDbCommand(query, connection)
                        If Not CHeckIfEmployeeAlreadyExist(Employee_ID, NT_ID) Then
                            insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                            insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                            insertCommand.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                            insertCommand.Parameters.AddWithValue("@Scanner_5D", Scanner_5D)
                            insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                            insertCommand.Parameters.AddWithValue("@Gender", Gender)
                            insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                            insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                            insertCommand.Parameters.AddWithValue("@Functions", Functions)
                            insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                            insertCommand.Parameters.AddWithValue("@Created_By", Created_By)
                            insertCommand.Parameters.AddWithValue("@Created_At", Created_At)
                            insertCommand.ExecuteNonQuery()
                            Return "True"
                        Else
                            Return "False"
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function ReplaceMaterFile() As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM Employee_DB"
                    Using command As New OleDbCommand(query, connection)
                        Dim count As Integer = CInt(command.ExecuteScalar())
                        If count > 0 Then
                            Dim deleteQuery As String = $"DELETE * FROM  Employee_DB "
                            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                deleteCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function ExportDataFromEmployeeDB() As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM Employee_DB"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function ExportDataFromUser_DB() As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM User_DB"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function RemoveSelectedEmployessFromDB(Emp_ID As String) As String
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Dim Scanner_ID As String
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim Query As String = $"SELECT * FROM Employee_DB WHERE Employee_ID=@Employee_ID"
                    Using Command As New OleDbCommand(Query, connection)
                        Command.Parameters.AddWithValue("@Employee_ID", Emp_ID)
                        Using reader As OleDbDataReader = Command.ExecuteReader()
                            If reader.Read Then
                                Scanner_ID = reader("Scanner_ID")
                            End If
                        End Using
                    End Using
                    Dim Query2 As String = $"SELECT * FROM User_DB WHERE Scanner_ID=@Scanner_ID"
                    Using Command2 As New OleDbCommand(Query2, connection)
                        Command2.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        Using reader As OleDbDataReader = Command2.ExecuteReader()
                            If reader.HasRows Then
                                Dim Query3 As String = $"DELETE FROM User_DB WHERE Scanner_ID=@Scanner_ID"
                                Using Command3 As New OleDbCommand(Query3, connection)
                                    Command3.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                                    Command3.ExecuteNonQuery()
                                End Using
                            Else
                                'donothing
                            End If
                        End Using
                    End Using
                    Dim Query1 As String = $"DELETE FROM Employee_DB WHERE Employee_ID=@Employee_ID"
                    Using deleteCommand As New OleDbCommand(Query1, connection)
                        deleteCommand.Parameters.AddWithValue("@Employee_ID", Emp_ID)
                        deleteCommand.ExecuteNonQuery()
                    End Using
                End Using
                    Return "True"
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
        Public Shared Function ExportDataFromUserDB() As Tuple(Of DataTable, String)
            Dim Grid As New DataTable()
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"SELECT * FROM User_DB"
                    Using adapter As New OleDbDataAdapter(query, connection)
                        adapter.Fill(Grid)
                    End Using
                End Using
                Return Tuple.Create(Grid, "")
            Catch ex As Exception
                Return Tuple.Create(Of DataTable, String)(Nothing, ex.Message)
            End Try
        End Function
        Public Shared Function UpdateEmployeeData(Nt_ID As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Issue_Gift As Boolean, Events_Addition As Boolean, Admn_Access As Boolean) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"UPDATE User_DB SET Buisness_Area=@Buisness_Area, Functions=@Functions, Leave_Approver=@Leave_Approver, Modified_By=@Modified_By, Modified_At=@Modified_At,Issue_Gift=@Issue_Gift, Events_Addition=@Events_Addition, Admn_Access=@Admn_Access WHERE NT_ID=@NT_ID"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Functions", Functions)
                        command.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        command.Parameters.AddWithValue("@Modified_By", $"{Environment.UserName}")
                        command.Parameters.AddWithValue("@Modified_At", $"{DateTime.Now.ToString("G")}")
                        command.Parameters.AddWithValue("@Issue_Gift", Issue_Gift)
                        command.Parameters.AddWithValue("@Events_Addition", Events_Addition)
                        command.Parameters.AddWithValue("@Admn_Access", Admn_Access)
                        command.Parameters.AddWithValue("@NT_ID", Nt_ID)
                        command.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function UpdateEmployeeDataOfEmpDB(Employee_ID As String, NT_ID As String, Scanner_ID As String, Scanner_5D As String, Employee_Name As String, Gender As String, Celebration_Date As Date, Buisness_Area As String, Functions As String, Leave_Approver As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"UPDATE Employee_DB SET Employee_ID=@Employee_ID, NT_ID=@NT_ID, Scanner_ID=@Scanner_ID, Scanner_5D=@Scanner_5D, Employee_Name=@Employee_Name, Gender=@Gender, Celebration_Date=@Celebration_Date, Buisness_Area=@Buisness_Area, Functions=@Functions, Leave_Approver=@Leave_Approver, Created_By=@Created_By, Created_At=@Created_At WHERE Employee_ID=@Employee_ID"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        command.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                    End Using
                End Using
            Catch ex As Exception

            End Try
        End Function
        Public Shared Function FetchCriterialCount(Ocassion As String) As Tuple(Of Integer, Integer, String)
            Try
                ini.Load(Inipath)
                If ini.Sections("TYPE").Keys($"{Ocassion}").Value = "Variable" Then
                    Dim repeatspantime As Integer = ini.Sections("REPEATSPANTIME").Keys($"{Ocassion}").Value
                    Dim eligibilityspantime As Integer = ini.Sections("ELIGIBILITY").Keys($"{Ocassion}").Value
                    Return Tuple.Create(repeatspantime, eligibilityspantime, "")
                ElseIf ini.Sections("TYPE").Keys($"{Ocassion}").Value = "Constant" Then
                    Dim eligibilityspantime As Integer = ini.Sections("ELIGIBILITY").Keys($"{Ocassion}").Value
                    Return Tuple.Create(0, eligibilityspantime, "")
                End If
            Catch ex As Exception
                Return Tuple.Create(0, 0, ex.Message)
            End Try
        End Function
        Public Shared Function UpdateCriteriaFunction(Occassion As String, Repeatspantime As Integer, eligibilitytime As Integer) As String
            Try
                Dim inidata As New IniFile
                inidata.Load(Inipath)
                If ini.Sections("TYPE").Keys($"{Occassion}").Value = "Variable" Then
                    inidata.Sections("REPEATSPANTIME").Keys($"{Occassion}").Value = Repeatspantime
                    inidata.Sections("ELIGIBILITY").Keys($"{Occassion}").Value = eligibilitytime
                ElseIf ini.Sections("TYPE").Keys($"{Occassion}").Value = "Constant" Then
                    inidata.Sections("ELIGIBILITY").Keys($"{Occassion}").Value = eligibilitytime
                End If
                inidata.Save(Inipath)
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function FetchEmployeeDataForExceptionGiftSharing(Scanner_ID As String) As Tuple(Of String, String, String, String, String, String, String)
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT * FROM Employee_DB WHERE Scanner_ID = @Scanner_ID"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            While reader.Read
                                Return Tuple.Create("", reader("Employee_ID").ToString, reader("Employee_Name").ToString, reader("NT_ID").ToString, reader("Leave_Approver").ToString, reader("Functions").ToString, reader("Buisness_Area").ToString)
                            End While
                        Else
                            Return Tuple.Create("Invalid", "", "", "", "", "", "")
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return Tuple.Create(ex.Message, "", "", "", "", "", "")
            End Try
        End Function
        Public Shared Function ShareExceptionGift(Employee_ID As String, NT_ID As String, Employee_Name As String, Celebration_Date As Date, Buisness_Area As String, Functions As String, Leave_Approver As String, Gift_Type As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim insertQuery As String = $"INSERT INTO Gift_History (Employee_ID,NT_ID,Employee_Name,Celebration_Date,Buisness_Area,Functions,Leave_Approver,Gift_Type,GivenBy_ID,GivenBy_NT,GivenBy_Name,Shared_At)" &
                                          "VALUES (@Employee_ID,@NT_ID,@Employee_Name,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver,@Gift_Type,@GivenBy_ID,@GivenBy_NT,@GivenBy_Name,@Shared_At)"
                    Using insertCommand As New OleDbCommand(insertQuery, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        If NT_ID IsNot Nothing OrElse Not IsDBNull(NT_ID) Then
                            insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        Else
                            insertCommand.Parameters.AddWithValue("@NT_ID", "")
                        End If
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Gift_Type", Gift_Type)
                        insertCommand.Parameters.AddWithValue("@GivenBy_ID", MainWindow.ListBox1.Items(0).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_NT", MainWindow.ListBox1.Items(1).ToString)
                        insertCommand.Parameters.AddWithValue("@GivenBy_Name", MainWindow.ListBox1.Items(4).ToString)
                        insertCommand.Parameters.AddWithValue("@Shared_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try

        End Function
        Public Shared Function ClearHistory() As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using Connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = "DELETE * FROM Gift_History"
                    Using insertCommand As New OleDbCommand(query, connection)
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function CheckifEmployeeExistforaddnewemp(Scanner_ID As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = "SELECT * FROM Employee_DB WHERE Scanner_ID = @Scanner_ID"
                    Using command As New OleDbCommand(query, connection)
                        command.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        Dim reader As OleDbDataReader = command.ExecuteReader()
                        If reader.HasRows Then
                            Return "True"
                        Else
                            Return "False"
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function AddNewEmployeeToEmpDB(Employee_ID As String, NT_ID As String, Scanner_ID As String, Scanner_5D As String, Employee_Name As String, Gender As String, Celebration_Date As Date, Buisness_Area As String, Functions As String, Leave_Approver As String) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"INSERT INTO Employee_DB (Employee_ID,NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Celebration_Date,Buisness_Area,Functions,Leave_Approver, Created_By, Created_At)" &
                                                 "VALUES (@Employee_ID,@NT_ID,@Scanner_ID,@Scanner_5D,@Employee_Name,@Gender,@Celebration_Date,@Buisness_Area,@Functions,@Leave_Approver, @Created_By, @Created_At)"
                    Using insertCommand As New OleDbCommand(query, connection)
                        insertCommand.Parameters.AddWithValue("@Employee_ID", Employee_ID)
                        If NT_ID IsNot Nothing OrElse Not IsDBNull(NT_ID) Then
                            insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        Else
                            insertCommand.Parameters.AddWithValue("@NT_ID", "")
                        End If
                        insertCommand.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        insertCommand.Parameters.AddWithValue("@Scanner_5D", Scanner_5D)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Gender", Gender)
                        insertCommand.Parameters.AddWithValue("@Celebration_Date", Celebration_Date)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Created_By", Environment.UserName)
                        insertCommand.Parameters.AddWithValue("@Created_At", DateTime.Now.ToString("G"))
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Public Shared Function AddNewEmployeeToUserDatabase(NT_ID As String, Scanner_ID As String, Scanner_5D As String, Employee_Name As String, Gender As String, Buisness_Area As String, Functions As String, Leave_Approver As String, Issue_Gift As Boolean, Events_Addition As Boolean, Admn_Access As Boolean) As String
            Try
                ini.Load(Inipath)
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim query As String = $"INSERT INTO User_DB (NT_ID,Scanner_ID,Scanner_5D,Employee_Name,Gender,Buisness_Area,Functions,Leave_Approver, Modified_By, Modified_At,Issue_Gift,Events_Addition,Admn_Access)" &
                                                 "VALUES (@NT_ID,@Scanner_ID,@Scanner_5D,@Employee_Name,@Gender,@Buisness_Area,@Functions,@Leave_Approver, @Modified_By, @Modified_At, @Issue_Gift, @Events_Addition, @Admn_Access)"
                    Using insertCommand As New OleDbCommand(query, connection)
                        insertCommand.Parameters.AddWithValue("@NT_ID", NT_ID)
                        insertCommand.Parameters.AddWithValue("@Scanner_ID", Scanner_ID)
                        insertCommand.Parameters.AddWithValue("@Scanner_5D", Scanner_5D)
                        insertCommand.Parameters.AddWithValue("@Employee_Name", Employee_Name)
                        insertCommand.Parameters.AddWithValue("@Gender", Gender)
                        insertCommand.Parameters.AddWithValue("@Buisness_Area", Buisness_Area)
                        insertCommand.Parameters.AddWithValue("@Functions", Functions)
                        insertCommand.Parameters.AddWithValue("@Leave_Approver", Leave_Approver)
                        insertCommand.Parameters.AddWithValue("@Modified_By", Environment.UserName)
                        insertCommand.Parameters.AddWithValue("@Modified_At", DateTime.Now.ToString("G"))
                        insertCommand.Parameters.AddWithValue("@Issue_Gift", Issue_Gift)
                        insertCommand.Parameters.AddWithValue("@Events_Addition", Events_Addition)
                        insertCommand.Parameters.AddWithValue("@Admn_Access", Admn_Access)
                        insertCommand.ExecuteNonQuery()
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
    End Class
    Public Class EnhanceFeatures
        Inherits Form
        Public Shared Sub IconbuildafterLogin(Flag As Boolean, EmployeeName As String)
            Dim iconButton As New Bunifu.UI.WinForms.BunifuButton.BunifuButton()
            If Flag = True Then
                MainWindow.BunifuDropdown1.Text = EmployeeName
                Dim rnd As New Random()
                Dim red As Integer = rnd.Next(256)
                Dim green As Integer = rnd.Next(256)
                Dim blue As Integer = rnd.Next(256)
                iconButton.Size = New Size(32, 32)
                iconButton.AllowAnimations = True
                iconButton.AllowMouseEffects = True
                iconButton.AllowToggling = False
                iconButton.AnimationSpeed = 200
                iconButton.AutoGenerateColors = True
                iconButton.AutoRoundBorders = False
                iconButton.AutoSize = False
                iconButton.ForeColor = Color.Black
                iconButton.IdleBorderRadius = 30
                iconButton.IdleBorderThickness = 1
                iconButton.IdleBorderColor = Color.Cyan
                iconButton.TextAlign = ContentAlignment.MiddleCenter
                iconButton.Text = Nothing
                iconButton.Text = String.Join("", EmployeeName.Split(" "c).Select(Function(word) word.Substring(0, 1))).ToUpper()
                Dim iconButtonX As Integer = 3
                iconButton.Location = New Point(iconButtonX, 3)
                iconButton.IdleFillColor = Color.FromArgb(red, green, blue)
                iconButton.IdleBorderColor = Color.Black
                MainWindow.BunifuDropdown1.Controls.Clear()
                MainWindow.BunifuDropdown1.Controls.Add(iconButton)
                MainWindow.BunifuDropdown1.BackgroundColor = Color.Transparent
                MainWindow.BunifuDropdown1.BorderColor = Color.FromArgb(red, green, blue)
                iconButton.Cursor = Cursors.Hand
                Dim toolTip As New Bunifu.UI.WinForms.BunifuToolTip
                AddHandler iconButton.MouseEnter, Sub(sender As Object, e As EventArgs)
                                                      toolTip.SetToolTip(iconButton, $"Name:{EmployeeName}{Environment.NewLine}Employee ID : {MainWindow.ListBox1.Items(0)}{Environment.NewLine}Buisness Area: {MainWindow.ListBox1.Items(7)}{Environment.NewLine}Function: {MainWindow.ListBox1.Items(8)}{Environment.NewLine}NT ID : {MainWindow.ListBox1.Items(1)}")
                                                  End Sub
                AddHandler iconButton.MouseLeave, Sub(sender As Object, e As EventArgs)
                                                      toolTip.Hide()
                                                  End Sub
            ElseIf Flag = False Then
                MainWindow.BunifuDropdown1.Controls.Remove(iconButton)
            End If
        End Sub
        Public Shared Function CalculateDateDifference(startDate As Date, endDate As Date) As Integer
            Dim startDayOfYear As Integer = startDate.DayOfYear
            Dim endDayOfYear As Integer = endDate.DayOfYear
            Dim daysDifference As Integer
            If startDate.Year = endDate.Year Then
                daysDifference = Math.Abs(endDayOfYear - startDayOfYear)
            Else
                Dim startYearDays As Integer = If(Date.IsLeapYear(startDate.Year), 366, 365) - startDayOfYear
                Dim endYearDays As Integer = endDayOfYear
                daysDifference = startYearDays + endYearDays
            End If
            Return daysDifference
        End Function
        Public Shared Function CalculateDateDifferenceForEligibility(startDate As Date, endDate As Date) As Integer
            Dim startDayOfYear As Integer = startDate.DayOfYear
            Dim endDayOfYear As Integer = endDate.DayOfYear
            Dim daysDifference As Integer
            If startDayOfYear <= endDayOfYear Then
                daysDifference = endDayOfYear - startDayOfYear
            Else
                Dim daysInStartYear As Integer = If(Date.IsLeapYear(startDate.Year), 366, 365)
                daysDifference = daysInStartYear - startDayOfYear + endDayOfYear
            End If
            Return daysDifference
        End Function
    End Class
    Public Class BirthdayCardForm
        Inherits Form
        Public Shared Function InitializeBirthdayCards() As String
            Dim BunifuSnackbar1 As New Bunifu.UI.WinForms.BunifuSnackbar
            Try
                Dim connectionString As String = $"Provider={ini.Sections("DATABASE").Keys("Provider").Value};Data Source={ini.Sections("DATABASE").Keys("Source").Value};Jet OLEDB:{ini.Sections("DATABASE").Keys("Jet_OLEDB").Value} Password={ini.Sections("DATABASE").Keys("Password").Value};"
                Using connection As New OleDbConnection(connectionString)
                    connection.Open()
                    Dim cardSize As New Size(150, 170)
                    Dim cardSpacing As New Size(20, 20)
                    Dim cardsPerRow As Integer = 7
                    Dim currentRow As Integer = 0
                    Dim currentColumn As Integer = 0
                    Dim query As String = $"SELECT * FROM Employee_DB WHERE DatePart('d', [Celebration_Date]) = DatePart('d', Date()) AND Datepart('m',[Celebration_Date])=DatePart('m',Date())"
                    Using command As New OleDbCommand(query, connection)
                        Using reader As OleDbDataReader = command.ExecuteReader()
                            While reader.Read()
                                Dim EmpID As String = reader("Employee_ID")
                                Dim NT_ID As String = reader("NT_ID")
                                Dim EmployeeName As String = reader("Employee_Name")
                                Dim birthday As Date = reader("Celebration_Date")
                                Dim Buisness_Area As String = reader("Buisness_Area")
                                Dim Func As String = reader("Functions")
                                Dim Scanner_ID As String = reader("Scanner_ID")
                                Dim Scanner_5D As String = reader("Scanner_5D")
                                Dim Leave_Approver As String = reader("Leave_Approver")
                                Dim Gender As String = reader("Gender")
                                If FuncLib.DataBaseOperations.CheckifDataAlreadyExist("BirthDay", EmpID, NT_ID, connection) = "True" Then
                                    If FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "True" Then
                                        'FuncLib.DataBaseOperations.AddToFunctionDB("BirthDay", EmpID, NT_ID, Scanner_ID, EmployeeName, Gender, birthday, Buisness_Area, Func, Leave_Approver)
                                    ElseIf FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "False" Then
                                        'FuncLib.DataBaseOperations.AddToFunctionDB("BirthDay", EmpID, NT_ID, Scanner_ID, EmployeeName, Gender, birthday, Buisness_Area, Func, Leave_Approver)
                                    End If
                                ElseIf FuncLib.DataBaseOperations.CheckifDataAlreadyExist("BirthDay", EmpID, NT_ID, connection) = "False" Then
                                    If FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "True" Then
                                        'FuncLib.DataBaseOperations.AddToFunctionDB("BirthDay", EmpID, NT_ID, Scanner_ID, EmployeeName, Gender, birthday, Buisness_Area, Func, Leave_Approver)
                                    ElseIf FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "False" Then
                                        FuncLib.DataBaseOperations.AddToFunctionDB("BirthDay", EmpID, NT_ID, Scanner_ID, Scanner_5D, EmployeeName, Gender, birthday, Buisness_Area, Func, Leave_Approver)
                                    End If
                                End If
                                Dim rnd As New Random()
                                Dim red As Integer = rnd.Next(256)
                                Dim green As Integer = rnd.Next(256)
                                Dim blue As Integer = rnd.Next(256)
                                Dim card As New BunifuCards()
                                card.Size = cardSize
                                card.BackColor = Color.FromArgb(red, green, blue)
                                card.ShadowDepth = 40
                                card.BorderRadius = 20
                                card.Location = New Point(currentColumn * (cardSize.Width + cardSpacing.Width) + 20, currentRow * (cardSize.Height + cardSpacing.Height) + 20)

                                Dim iconButton As New Bunifu.UI.WinForms.BunifuButton.BunifuButton()
                                iconButton.Size = New Size(50, 50)
                                iconButton.AllowAnimations = True
                                iconButton.AllowMouseEffects = True
                                iconButton.AllowToggling = False
                                iconButton.AnimationSpeed = 200
                                iconButton.AutoGenerateColors = True
                                iconButton.AutoRoundBorders = False
                                iconButton.AutoSize = False
                                iconButton.ForeColor = Color.White
                                iconButton.IdleBorderRadius = 47
                                iconButton.IdleBorderThickness = 1
                                iconButton.TextAlign = ContentAlignment.MiddleCenter
                                Dim trimmedName As String = EmployeeName.TrimEnd()
                                If Not String.IsNullOrWhiteSpace(trimmedName) Then
                                    If trimmedName.EndsWith(" ") Then
                                        trimmedName &= ""
                                    End If
                                    Dim initials = trimmedName.Split(" "c).
                                                    Where(Function(word) Not String.IsNullOrWhiteSpace(word)).
                                                    Select(Function(word) If(word.Length > 0, word.Substring(0, 1), "")).
                                                    ToArray()
                                    iconButton.Text = String.Join("", initials).ToUpper()
                                End If
                                iconButton.Location = New Point((card.Width - iconButton.Width) \ 2, 10)
                                iconButton.IdleFillColor = Color.FromArgb(red, green, blue)
                                iconButton.Cursor = Cursors.Hand
                                iconButton.onHoverState.FillColor = Color.FromArgb(green, blue, red)
                                Dim toolTip As New Bunifu.UI.WinForms.BunifuToolTip
                                AddHandler iconButton.MouseEnter, Sub(sender As Object, e As EventArgs)
                                                                      toolTip.SetToolTip(iconButton, $"Name: {EmployeeName}{Environment.NewLine}Employee ID : {EmpID}{Environment.NewLine}Buisness Area: {Buisness_Area}{Environment.NewLine}Function: {Func}{Environment.NewLine}NT ID : {NT_ID}")
                                                                  End Sub
                                AddHandler iconButton.MouseLeave, Sub(sender As Object, e As EventArgs)
                                                                      toolTip.Hide()
                                                                  End Sub
                                card.Controls.Add(iconButton)

                                Dim nameLabel As New Windows.Forms.Label()
                                nameLabel.Text = $"Happy Birthday, {Environment.NewLine}{EmployeeName.Split(" "c)(0)}!"
                                nameLabel.AutoSize = False
                                nameLabel.Size = New Size(card.Width - 40, 30)
                                nameLabel.Location = New Point(20, iconButton.Bottom + 10)
                                nameLabel.TextAlign = ContentAlignment.MiddleCenter
                                card.Controls.Add(nameLabel)

                                Dim age As Integer = Date.Now.Year - birthday.Year
                                Dim ageLabel As New Windows.Forms.Label()
                                ageLabel.Text = $"Turning {age} years old"
                                ageLabel.AutoSize = False
                                ageLabel.Size = New Size(card.Width - 40, 30)
                                ageLabel.Location = New Point(20, nameLabel.Bottom + 5)
                                ageLabel.TextAlign = ContentAlignment.MiddleCenter
                                card.Controls.Add(ageLabel)

                                Dim GiftButton As New Bunifu.UI.WinForms.BunifuButton.BunifuButton()
                                GiftButton.Size = New Size(90, 20)
                                GiftButton.Location = New Point((card.Width - GiftButton.Width) \ 2, ageLabel.Bottom + 5)
                                GiftButton.Text = "Issue Gift"
                                GiftButton.AllowAnimations = True
                                GiftButton.AllowMouseEffects = True
                                GiftButton.AllowToggling = False
                                GiftButton.AnimationSpeed = 200
                                GiftButton.AutoGenerateColors = True
                                GiftButton.AutoRoundBorders = False
                                GiftButton.AutoSize = False
                                GiftButton.ForeColor = Color.White
                                GiftButton.IdleBorderRadius = 15
                                GiftButton.IdleBorderThickness = 1
                                GiftButton.TextAlign = ContentAlignment.MiddleCenter
                                GiftButton.IdleFillColor = Color.FromArgb(red, green, blue)
                                GiftButton.IdleBorderColor = Color.FromArgb(red, green, blue)
                                GiftButton.onHoverState.FillColor = Color.FromArgb(green, blue, red)
                                GiftButton.Cursor = Cursors.Hand
                                AddHandler GiftButton.MouseClick, Sub(sender As Object, e As EventArgs)
                                                                      Dim CheckEligibilityforconstant = FuncLib.Gifts.CheckEmployeeEligibilityForConstantFunctions("BirthDay", EmpID)
                                                                      If CheckEligibilityforconstant = "True" Then
                                                                          Dim DistributeonConstantFunction As String = FuncLib.Gifts.RecordEmployeeGiftDistributionForConstantFunction("BirthDay", EmpID, NT_ID, EmployeeName, Buisness_Area, Func, Leave_Approver, birthday)
                                                                          If DistributeonConstantFunction = "True" Then
                                                                              BunifuSnackbar1.Show(MainWindow, $"BirtDay Gift Shared To {EmployeeName}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                                                                              card.Controls.Remove(GiftButton)
                                                                          Else
                                                                              BunifuSnackbar1.Show(MainWindow, $"Error In Sharing The Gift. {Environment.NewLine} Error : {DistributeonConstantFunction}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                                                              FuncLib.WriteLog.WriteErrorLog($"Error Occured While Sharing Gift From Cards : {DistributeonConstantFunction}")
                                                                              Exit Sub
                                                                          End If
                                                                      ElseIf CheckEligibilityforconstant = "TLE" Then
                                                                          BunifuSnackbar1.Show(MainWindow, $"{EmployeeName} is Not Allowed For BirthDay Gift. {Environment.NewLine} Reason : Time Limit Excedded", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                                                                      ElseIf CheckEligibilityforconstant Is Nothing Then
                                                                          BunifuSnackbar1.Show(MainWindow, $"{EmployeeName} is Not Allowed For BirthDay Gift. {Environment.NewLine} Reason : name Not Found In The BirthDay List", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                                                                      ElseIf CheckEligibilityforconstant = "Same" Then
                                                                          BunifuSnackbar1.Show(MainWindow, $"{EmployeeName} Can Not Be Shared Gift.{Environment.NewLine} Reason : You Cannot Share Gift To Yourself ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                                                                      Else

                                                                          BunifuSnackbar1.Show(MainWindow, $"Error occured While Sharing Gift.{Environment.NewLine} Error : {CheckEligibilityforconstant}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                                                                          FuncLib.WriteLog.WriteErrorLog($"Error Occured While Sharing Gift From Cards : {CheckEligibilityforconstant}")
                                                                          Exit Sub
                                                                      End If
                                                                  End Sub
                                If FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "True" Then
                                    card.Controls.Remove(GiftButton)
                                ElseIf FuncLib.DataBaseOperations.CheckIFGiftShared("BirthDay", EmpID, NT_ID, connection) = "False" Then
                                    If MainWindow.ListBox1.Items(14) = True Or MainWindow.ListBox1.Items(16) = True Then
                                        card.Controls.Add(GiftButton)
                                    Else
                                        card.Controls.Remove(GiftButton)
                                    End If
                                End If
                                If MainWindow.ListBox1.Items(0).ToString = EmpID Then
                                    'donothing
                                Else
                                    MainWindow.TabPage1.Controls.Add(card)
                                    currentColumn += 1
                                    If currentColumn >= cardsPerRow Then
                                        currentColumn = 0
                                        currentRow += 1
                                    End If
                                End If
                            End While
                        End Using
                    End Using
                End Using
                Return "True"
            Catch ex As Exception
                Return ex.Message.ToString
            End Try
        End Function
    End Class
End Module
