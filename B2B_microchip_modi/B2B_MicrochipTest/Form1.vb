Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports Renci.SshNet

Public Class Form1
    Private Detail As DataTable 'Variable to store captured data
    Private ReadOnly sqlHandler As SQLHandler
    Dim EmailSubject As String
    Dim em As New EmailHandler
    Dim dsEmail As DataSet = em.GetMailRecipients(141)
    Dim filePaths As String ' Declare filePath variable at class level

    Public Sub New()
        InitializeComponent()
        Dim connectionString As String = "Server=MSDYNAMICS-DB\AXDB;Database=MES_ATEC;User Id=sa;Password=p@ssw0rd;" 'need to change Credentials
        sqlHandler = New SQLHandler(connectionString)
    End Sub

    Private Sub All_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim folderPath As String = "F:\" 'Live
        Dim folderPath As String = "C:\xml\de_xml\" 'Test Path

        Try
            Dim query As String = "usp_TRN_Microchip_Recieve '2023-06-15 16:25:34.563', '2023-06-15 16:25:34.563'"
            Detail = sqlHandler.ExecuteQuery(query)
            If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
                ' Call CreateCSVAndCleanup and pass folderPath as a parameter
                CreateCSVAndCleanup(folderPath)
                UploadFileToSFTP(filePaths)
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
        Me.Close()
    End Sub

    Private Sub CreateCSVAndCleanup(folderPath As String)
        Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
        filePaths = Path.Combine(folderPath, fileName) ' Assign value to filePath

        Try
            If Detail.Rows.Count = 0 Then
                MessageBox.Show("No data to export.")
                Return
            End If

            Dim newDataRows As New List(Of DataRow)()

            For Each row As DataRow In Detail.Rows
                If Not DataExistsInDatabase(row) Then
                    newDataRows.Add(row)
                End If
            Next

            If newDataRows.Count = 0 Then
                MessageBox.Show("No new data to export.")
                Return
            End If

            Using writer As New StreamWriter(filePaths)
                Dim headerLine As String = String.Join(",", Detail.Columns.Cast(Of DataColumn)().Select(Function(column) column.ColumnName))
                writer.WriteLine(headerLine)

                For Each row As DataRow In newDataRows
                    Dim dataLine As String = String.Join(",", row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), "")))
                    writer.WriteLine(dataLine)
                Next
            End Using

            'MessageBox.Show($"CSV file created successfully at: {filePaths}")
            SaveData(newDataRows) 'Save Data
            SendEmailWithExcelData() 'Email Sending
            Detail.Clear() 'Clear Detail DataTable
        Catch ex As Exception
            MessageBox.Show($"Error creating CSV file: {ex.Message}")
        End Try
    End Sub

    Private Sub UploadFileToSFTP(filePath As String)
        Dim localDirectory As String = "C:\xml\de_xml\" ' Source directory where files are located
        Dim backupDirectory As String = "C:\xml\backup_relocate\" ' relocate directory files
        Dim sftpDirectory As String = "/Processed/" ' Remote SFTP directory path

        Try
            ' Initialize SFTP client
            Using client As New SftpClient("sftp.microchip.com", "ATIC", "h]?(TYN,2cMv5=V-")
                client.Connect()

                Dim files As String() = Directory.GetFiles(localDirectory)

                For Each filePaths As String In files
                    Dim fileInfo As New FileInfo(filePath)
                    Dim fileName As String = fileInfo.Name
                    Using fileStream As New FileStream(filePath, FileMode.Open)
                        client.UploadFile(fileStream, sftpDirectory & fileName)
                    End Using

                    Dim destinationFile As String = Path.Combine(backupDirectory, fileName)
                    fileInfo.MoveTo(destinationFile)
                Next

                client.Disconnect()
            End Using

            'Console.WriteLine("Files uploaded and moved to backup directory successfully!")
        Catch ex As Exception
            Console.WriteLine("An error occurred: " & ex.Message)
        End Try
        Environment.Exit(0)
    End Sub


    Private Function DataExistsInDatabase(row As DataRow) As Boolean
        Try
            Dim query As String = "SELECT COUNT(*) FROM TRN_Microchip_csv_logs WHERE " &
                      "Supplier_Name = @Supplier_Name AND " &
                      "Stage = @Stage AND " &
                      "Step = @Step AND " &
                      "PartCode = @PartCode AND " &
                      "Bin = @Bin AND " &
                      "Ship_To = @Ship_To AND " &
                      "Lot_Number = @Lot_Number AND " &
                      "Lot_wafer_Qty = @Lot_wafer_Qty AND " &
                      "Lot_Qty = @Lot_Qty AND " &
                      "Wafer_ID = @Wafer_ID AND " &
                      "Action_Code = @Action_Code AND " &
                      "Child_lot = @Child_lot AND " &
                      "Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND " &
                      "Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND " &
                      "Transaction_Date = @Transaction_Date AND " &
                      "Lot_Type = @Lot_Type AND " &
                      "Customer_Code = @Customer_Code"

            Dim parameters As New List(Of SqlParameter)()
            parameters.Add(New SqlParameter("@Supplier_Name", row("Supplier_Name")))
            parameters.Add(New SqlParameter("@Stage", row("Stage")))
            parameters.Add(New SqlParameter("@Step", row("Step")))
            parameters.Add(New SqlParameter("@PartCode", row("PartCode")))
            parameters.Add(New SqlParameter("@Bin", row("Bin")))
            parameters.Add(New SqlParameter("@Ship_To", row("Ship_To")))
            parameters.Add(New SqlParameter("@Lot_Number", row("Lot_Number")))
            parameters.Add(New SqlParameter("@Lot_wafer_Qty", row("Lot_wafer_Qty")))
            parameters.Add(New SqlParameter("@Lot_Qty", row("Lot_Qty")))
            parameters.Add(New SqlParameter("@Wafer_ID", row("Wafer_ID")))
            parameters.Add(New SqlParameter("@Action_Code", row("Action_Code")))
            parameters.Add(New SqlParameter("@Child_lot", row("Child_lot")))
            parameters.Add(New SqlParameter("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty")))
            parameters.Add(New SqlParameter("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID")))
            parameters.Add(New SqlParameter("@Transaction_Date", row("Transaction_Date")))
            parameters.Add(New SqlParameter("@Lot_Type", row("Lot_Type")))
            parameters.Add(New SqlParameter("@Customer_Code", row("Customer_Code")))

            Dim rowCount As Integer = Convert.ToInt32(sqlHandler.ExecuteScalar(query, parameters.ToArray()))

            Return rowCount > 0
        Catch ex As Exception
            MessageBox.Show("Error checking database: " & ex.Message)
            Return True
        End Try
    End Function

    Private Sub SaveData(newDataRows As List(Of DataRow))
        Try
            Dim connectionString As String = "Server=MSDYNAMICS-DB\AXDB;Database=MES_ATEC;User Id=sa;Password=p@ssw0rd;" ' Replace with your actual connection string

            Using connection As New SqlConnection(connectionString)
                connection.Open()

                For Each row As DataRow In newDataRows

                    Dim existsQuery As String = "SELECT * FROM TRN_Microchip_csv_logs WHERE Supplier_Name = @Supplier_Name AND Stage = @Stage AND Step = @Step AND PartCode = @PartCode AND Bin = @Bin AND Ship_To = @Ship_To AND Lot_Number = @Lot_Number AND Lot_wafer_Qty = @Lot_wafer_Qty AND Lot_Qty = @Lot_Qty AND Wafer_ID = @Wafer_ID AND Action_Code = @Action_Code AND Child_lot = @Child_lot AND Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND Transaction_Date = @Transaction_Date AND Lot_Type = @Lot_Type AND Customer_Code = @Customer_Code"
                    Using existsCommand As New SqlCommand(existsQuery, connection)

                        existsCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
                        existsCommand.Parameters.AddWithValue("@Stage", row("Stage"))
                        existsCommand.Parameters.AddWithValue("@Step", row("Step"))
                        existsCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
                        existsCommand.Parameters.AddWithValue("@Bin", row("Bin"))
                        existsCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
                        existsCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
                        existsCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
                        existsCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
                        existsCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
                        existsCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
                        existsCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
                        existsCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
                        existsCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
                        existsCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
                        existsCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
                        existsCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

                        Using reader As SqlDataReader = existsCommand.ExecuteReader()

                            If Not reader.HasRows Then
                                reader.Close()

                                Dim insertQuery As String = "INSERT INTO TRN_Microchip_csv_logs (Supplier_Name, Stage, Step, PartCode, Bin, Ship_To, Lot_Number, Lot_wafer_Qty, Lot_Qty, Wafer_ID, Action_Code, Child_lot, Child_Lot_wafer_Qty, Child_Lot_Wafer_ID, Transaction_Date, Lot_Type, Customer_Code) VALUES (@Supplier_Name, @Stage, @Step, @PartCode, @Bin, @Ship_To, @Lot_Number, @Lot_wafer_Qty, @Lot_Qty, @Wafer_ID, @Action_Code, @Child_lot, @Child_Lot_wafer_Qty, @Child_Lot_Wafer_ID, @Transaction_Date, @Lot_Type, @Customer_Code)"
                                Using insertCommand As New SqlCommand(insertQuery, connection)

                                    insertCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
                                    insertCommand.Parameters.AddWithValue("@Stage", row("Stage"))
                                    insertCommand.Parameters.AddWithValue("@Step", row("Step"))
                                    insertCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
                                    insertCommand.Parameters.AddWithValue("@Bin", row("Bin"))
                                    insertCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
                                    insertCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
                                    insertCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
                                    insertCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
                                    insertCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
                                    insertCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
                                    insertCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
                                    insertCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
                                    insertCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
                                    insertCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
                                    insertCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
                                    insertCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

                                    insertCommand.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End Using
                Next

            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub SendEmailWithExcelData()
        Try
            ' Construct email subject
            Dim Datenow As DateTime = DateTime.Now
            EmailSubject = "TEST " & Datenow.ToString()

            Dim message As New StringBuilder()
            message.AppendLine("<html><body>")

            Dim folderPath As String = "F:\" ' Test Path
            Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
            Dim filePath As String = Path.Combine(folderPath, fileName)

            em.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
        Catch ex As Exception
            MessageBox.Show($"Error sending email: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class


'Imports System.Data.SqlClient
'Imports System.IO
'Imports System.Text
'Imports Renci.SshNet

'Public Class Form1
'    Private Detail As DataTable 'Variable to store captured data
'    Private ReadOnly sqlHandler As SQLHandler
'    Dim EmailSubject As String
'    Dim em As New EmailHandler
'    Dim dsEmail As DataSet = em.GetMailRecipients(141)

'    Public Sub New()
'        InitializeComponent()
'        Dim connectionString As String = "Server=MSDYNAMICS-DB\AXDB;Database=MES_ATEC;User Id=sa;Password=p@ssw0rd;" 'need to change Credentials
'        sqlHandler = New SQLHandler(connectionString)
'    End Sub

'    Private Sub All_Load(sender As Object, e As EventArgs) Handles MyBase.Load

'        'Dim folderPath As String = "F:\" 'Live
'        Dim folderPath As String = "D:\" 'Test Path

'        Try
'            'Dim fromDate As DateTime = DateTime.Now.AddDays(-1) ' Set the start date
'            'Dim toDate As DateTime = DateTime.Now ' Set the end date
'            'Live Sending
'            'Dim query As String = $"usp_TRN_Microchip_csv '{fromDate.ToString("yyyy-MM-dd HH:mm:ss")}', '{toDate.ToString("yyyy-MM-dd HH:mm:ss")}'"


'            'Manual Sending
'            'Dim query As String = "usp_TRN_Microchip_csv '2023-04-01 09:09:01.843', '2023-06-15 16:26:42.127'"
'            Dim query As String = "usp_TRN_Microchip_Recieve '2023-06-15 16:25:34.563', '2023-06-15 16:25:34.563'"

'            Detail = sqlHandler.ExecuteQuery(query)
'            If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
'                CreateCSVAndCleanup(folderPath)
'                UploadFileToSFTP(filePath)
'            End If
'        Catch ex As Exception
'            'MessageBox.Show("Error: " & ex.Message)
'        End Try
'        Me.Close()
'    End Sub

'    Private Sub CreateCSVAndCleanup(folderPath As String)
'        Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
'        Dim filePath As String = Path.Combine(folderPath, fileName)

'        Try
'            If Detail.Rows.Count = 0 Then
'                'MessageBox.Show("No data to export.")
'                Return
'            End If

'            Dim newDataRows As New List(Of DataRow)()

'            For Each row As DataRow In Detail.Rows
'                If Not DataExistsInDatabase(row) Then
'                    newDataRows.Add(row)
'                End If
'            Next

'            If newDataRows.Count = 0 Then
'                'MessageBox.Show("No new data to export.")
'                Return
'            End If

'            Using writer As New StreamWriter(filePath)
'                Dim headerLine As String = String.Join(",", Detail.Columns.Cast(Of DataColumn)().Select(Function(column) column.ColumnName))
'                writer.WriteLine(headerLine)

'                For Each row As DataRow In newDataRows
'                    Dim dataLine As String = String.Join(",", row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), "")))
'                    writer.WriteLine(dataLine)
'                Next
'            End Using

'            'MessageBox.Show($"CSV file created successfully at: {filePath}")
'            SaveData(newDataRows) 'Save Data
'            SendEmailWithExcelData() 'Email Sending
'            Detail.Clear() 'Clear Detail DataTable
'        Catch ex As Exception
'            'MessageBox.Show($"Error creating CSV file: {ex.Message}")
'        End Try
'    End Sub

'    Private Sub UploadFileToSFTP(filePath As String)
'        Dim sftpHost As String = "sftp.microchip.com"
'        Dim sftpPort As Integer = 22
'        Dim sftpUsername As String = "ATIC"
'        Dim sftpPassword As String = "h]?(TYN,2cMv5=V-"

'        Using client As New SftpClient(sftpHost, sftpPort, sftpUsername, sftpPassword)
'            client.Connect()

'            Dim remoteDirectoryPath As String = "/Processed"

'            Using fileStream As New FileStream(filePath, FileMode.Open)
'                client.UploadFile(fileStream, $"{remoteDirectoryPath}/{Path.GetFileName(filePath)}")
'            End Using

'            client.Disconnect()
'        End Using
'    End Sub



'    Private Function DataExistsInDatabase(row As DataRow) As Boolean
'        Try
'            Dim query As String = "SELECT COUNT(*) FROM TRN_Microchip_csv_logs WHERE " &
'                      "Supplier_Name = @Supplier_Name AND " &
'                      "Stage = @Stage AND " &
'                      "Step = @Step AND " &
'                      "PartCode = @PartCode AND " &
'                      "Bin = @Bin AND " &
'                      "Ship_To = @Ship_To AND " &
'                      "Lot_Number = @Lot_Number AND " &
'                      "Lot_wafer_Qty = @Lot_wafer_Qty AND " &
'                      "Lot_Qty = @Lot_Qty AND " &
'                      "Wafer_ID = @Wafer_ID AND " &
'                      "Action_Code = @Action_Code AND " &
'                      "Child_lot = @Child_lot AND " &
'                      "Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND " &
'                      "Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND " &
'                      "Transaction_Date = @Transaction_Date AND " &
'                      "Lot_Type = @Lot_Type AND " &
'                      "Customer_Code = @Customer_Code"

'            Dim parameters As New List(Of SqlParameter)()
'            parameters.Add(New SqlParameter("@Supplier_Name", row("Supplier_Name")))
'            parameters.Add(New SqlParameter("@Stage", row("Stage")))
'            parameters.Add(New SqlParameter("@Step", row("Step")))
'            parameters.Add(New SqlParameter("@PartCode", row("PartCode")))
'            parameters.Add(New SqlParameter("@Bin", row("Bin")))
'            parameters.Add(New SqlParameter("@Ship_To", row("Ship_To")))
'            parameters.Add(New SqlParameter("@Lot_Number", row("Lot_Number")))
'            parameters.Add(New SqlParameter("@Lot_wafer_Qty", row("Lot_wafer_Qty")))
'            parameters.Add(New SqlParameter("@Lot_Qty", row("Lot_Qty")))
'            parameters.Add(New SqlParameter("@Wafer_ID", row("Wafer_ID")))
'            parameters.Add(New SqlParameter("@Action_Code", row("Action_Code")))
'            parameters.Add(New SqlParameter("@Child_lot", row("Child_lot")))
'            parameters.Add(New SqlParameter("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty")))
'            parameters.Add(New SqlParameter("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID")))
'            parameters.Add(New SqlParameter("@Transaction_Date", row("Transaction_Date")))
'            parameters.Add(New SqlParameter("@Lot_Type", row("Lot_Type")))
'            parameters.Add(New SqlParameter("@Customer_Code", row("Customer_Code")))

'            Dim rowCount As Integer = Convert.ToInt32(sqlHandler.ExecuteScalar(query, parameters.ToArray()))

'            Return rowCount > 0
'        Catch ex As Exception
'            'MessageBox.Show("Error checking database: " & ex.Message)
'            Return True
'        End Try
'    End Function

'    Private Sub SaveData(newDataRows As List(Of DataRow))
'        Try
'            Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

'            Using connection As New SqlConnection(connectionString)
'                connection.Open()

'                For Each row As DataRow In newDataRows

'                    Dim existsQuery As String = "SELECT * FROM TRN_Microchip_csv_logs WHERE Supplier_Name = @Supplier_Name AND Stage = @Stage AND Step = @Step AND PartCode = @PartCode AND Bin = @Bin AND Ship_To = @Ship_To AND Lot_Number = @Lot_Number AND Lot_wafer_Qty = @Lot_wafer_Qty AND Lot_Qty = @Lot_Qty AND Wafer_ID = @Wafer_ID AND Action_Code = @Action_Code AND Child_lot = @Child_lot AND Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND Transaction_Date = @Transaction_Date AND Lot_Type = @Lot_Type AND Customer_Code = @Customer_Code"
'                    Using existsCommand As New SqlCommand(existsQuery, connection)

'                        existsCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
'                        existsCommand.Parameters.AddWithValue("@Stage", row("Stage"))
'                        existsCommand.Parameters.AddWithValue("@Step", row("Step"))
'                        existsCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
'                        existsCommand.Parameters.AddWithValue("@Bin", row("Bin"))
'                        existsCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
'                        existsCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
'                        existsCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
'                        existsCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
'                        existsCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
'                        existsCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
'                        existsCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
'                        existsCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
'                        existsCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
'                        existsCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
'                        existsCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
'                        existsCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

'                        Using reader As SqlDataReader = existsCommand.ExecuteReader()

'                            If Not reader.HasRows Then
'                                reader.Close()

'                                Dim insertQuery As String = "INSERT INTO TRN_Microchip_csv_logs (Supplier_Name, Stage, Step, PartCode, Bin, Ship_To, Lot_Number, Lot_wafer_Qty, Lot_Qty, Wafer_ID, Action_Code, Child_lot, Child_Lot_wafer_Qty, Child_Lot_Wafer_ID, Transaction_Date, Lot_Type, Customer_Code) VALUES (@Supplier_Name, @Stage, @Step, @PartCode, @Bin, @Ship_To, @Lot_Number, @Lot_wafer_Qty, @Lot_Qty, @Wafer_ID, @Action_Code, @Child_lot, @Child_Lot_wafer_Qty, @Child_Lot_Wafer_ID, @Transaction_Date, @Lot_Type, @Customer_Code)"
'                                Using insertCommand As New SqlCommand(insertQuery, connection)

'                                    insertCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
'                                    insertCommand.Parameters.AddWithValue("@Stage", row("Stage"))
'                                    insertCommand.Parameters.AddWithValue("@Step", row("Step"))
'                                    insertCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
'                                    insertCommand.Parameters.AddWithValue("@Bin", row("Bin"))
'                                    insertCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
'                                    insertCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
'                                    insertCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
'                                    insertCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
'                                    insertCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
'                                    insertCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
'                                    insertCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
'                                    insertCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
'                                    insertCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
'                                    insertCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
'                                    insertCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
'                                    insertCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

'                                    insertCommand.ExecuteNonQuery()
'                                End Using
'                            End If
'                        End Using
'                    End Using
'                Next

'            End Using
'        Catch ex As Exception
'            'MessageBox.Show("Error: " & ex.Message)
'        End Try
'    End Sub

'    Private Sub SendEmailWithExcelData()
'        Try
'            ' Construct email subject
'            Dim Datenow As DateTime = DateTime.Now
'            EmailSubject = "TEST " & Datenow.ToString()

'            Dim message As New StringBuilder()
'            message.AppendLine("<html><body>")

'            Dim folderPath As String = "F:\" ' Test Path
'            Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
'            Dim filePath As String = Path.Combine(folderPath, fileName)

'            em.SendEmail(EmailSubject, message.ToString(), filePath, dsEmail)
'        Catch ex As Exception
'            MessageBox.Show($"Error sending email: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End Try
'    End Sub

'End Class
