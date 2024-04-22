Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class Form1
    Private Detail As DataTable 'Variable to store captured data
    Private ReadOnly sqlHandler As SQLHandler
    Dim EmailSubject As String
    Dim em As New EmailHandler
    Dim dsEmail As DataSet = em.GetMailRecipients(141)

    Public Sub New()
        InitializeComponent()
        Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" 'need to change Credentials
        sqlHandler = New SQLHandler(connectionString)
    End Sub

    Private Sub All_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Dim folderPath As String = "F:\" 'Live
        Dim folderPath As String = "F:\" 'Test Path

        Try
            Dim fromDate As DateTime = DateTime.Now.AddDays(-1) ' Set the start date
            Dim toDate As DateTime = DateTime.Now ' Set the end date
            'Live Sending
            Dim query As String = $"usp_TRN_Microchip_csv '{fromDate.ToString("yyyy-MM-dd HH:mm:ss")}', '{toDate.ToString("yyyy-MM-dd HH:mm:ss")}'"


            'Manual Sending
            'Dim query As String = "usp_TRN_Microchip_csv '2023-04-01 09:09:01.843', '2023-06-15 16:26:42.127'"
            'Dim query As String = "usp_TRN_Microchip_Recieve '2023-06-15 16:25:34.563', '2023-06-15 16:25:34.563'"

            Detail = sqlHandler.ExecuteQuery(query)
            If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
                CreateCSVAndCleanup(folderPath)
            End If
        Catch ex As Exception
            'MessageBox.Show("Error: " & ex.Message)
        End Try
        Me.Close()
    End Sub

    Private Sub CreateCSVAndCleanup(folderPath As String)
        Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
        Dim filePath As String = Path.Combine(folderPath, fileName)

        Try
            If Detail.Rows.Count = 0 Then
                'MessageBox.Show("No data to export.")
                Return
            End If

            Dim newDataRows As New List(Of DataRow)()

            For Each row As DataRow In Detail.Rows
                If Not DataExistsInDatabase(row) Then
                    newDataRows.Add(row)
                End If
            Next

            If newDataRows.Count = 0 Then
                'MessageBox.Show("No new data to export.")
                Return
            End If

            Using writer As New StreamWriter(filePath)
                Dim headerLine As String = String.Join(",", Detail.Columns.Cast(Of DataColumn)().Select(Function(column) column.ColumnName))
                writer.WriteLine(headerLine)

                For Each row As DataRow In newDataRows
                    Dim dataLine As String = String.Join(",", row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), "")))
                    writer.WriteLine(dataLine)
                Next
            End Using

            'MessageBox.Show($"CSV file created successfully at: {filePath}")
            SaveData(newDataRows) 'Save Data
            SendEmailWithExcelData() 'Email Sending
            Detail.Clear() 'Clear Detail DataTable
        Catch ex As Exception
            'MessageBox.Show($"Error creating CSV file: {ex.Message}")
        End Try
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
            'MessageBox.Show("Error checking database: " & ex.Message)
            Return True
        End Try
    End Function

    Private Sub SaveData(newDataRows As List(Of DataRow))
        Try
            Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

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
            'MessageBox.Show("Error: " & ex.Message)
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


'--------------Test 3--------------
'Imports System.Data.SqlClient
'Imports System.IO
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView

'Public Class Form1
'    Private Detail As DataTable 'Variable to store captured data
'    'Dim em As New EmailHandler
'    'Dim dsEmail As DataSet = em.GetMailRecipients(141)

'    Private Sub All_Load(sender As Object, e As EventArgs) Handles MyBase.Load
'        Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string
'        Dim folderPath As String = "F:\" ' Specify the folder path where you want to save the CSV file

'        'Dim Datenow As DateTime = DateTime.Now

'        Try
'            Using connection As New SqlConnection(connectionString)
'                connection.Open()

'                ' Execute the query to retrieve data
'                'Live Sending
'                'Dim query As String = "usp_TRN_Microchip_csv '', ''"

'                'Manual Sending

'                'Dim query As String = "usp_TRN_Microchip_csv '2023-04-01 09:09:01.843', '2023-06-15 16:26:42.127'"

'                Dim query As String = "usp_TRN_Microchip_Recieve '2023-06-15 16:25:34.563', '2023-06-15 16:25:34.563'"


'                Using command As New SqlCommand(query, connection)
'                    Using reader As SqlDataReader = command.ExecuteReader()
'                        Dim dataTable As New DataTable()
'                        dataTable.Load(reader)

'                        Detail = dataTable
'                    End Using
'                End Using
'            End Using
'        Catch ex As Exception
'            'MessageBox.Show("Error: " & ex.Message)
'        End Try

'        If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
'            CreateCSVAndCleanup()
'        End If

'        Me.Close()
'    End Sub

'    Private Sub SaveData()
'        Try
'            Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

'            Using connection As New SqlConnection(connectionString)
'                connection.Open()

'                For Each row As DataRow In Detail.Rows

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

'                                    insertCommand.ExecuteNonQuery() ' Execute the insert command
'                                End Using
'                            End If
'                        End Using
'                    End Using
'                Next

'                'MessageBox.Show("Data saved successfully.")
'            End Using
'        Catch ex As Exception
'            MessageBox.Show("Error: " & ex.Message)
'        End Try
'    End Sub

'    Private Sub CreateCSVAndCleanup()
'        Dim fileName As String = $"ATIC_TXN_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
'        Try
'            If Detail.Rows.Count = 0 Then
'                'MessageBox.Show("No data to export.")
'                Return
'            End If

'            ' Check if data already exists in the database
'            If DataExistsInDatabase() Then
'                'Email Message no file is created
'                'MessageBox.Show("Data already exists in the database. No need to create CSV file.")
'                Return
'            End If

'            If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
'                SaveData()
'            End If

'            Dim filePath As String = Path.Combine("F:\", fileName)
'            Using writer As New StreamWriter(filePath)

'                Dim headerLine As String = String.Join(",", Detail.Columns.Cast(Of DataColumn)().Select(Function(column) column.ColumnName))
'                writer.WriteLine(headerLine)

'                For Each row As DataRow In Detail.Rows
'                    Dim dataLine As String = String.Join(",", row.ItemArray.Select(Function(item) If(item IsNot Nothing, item.ToString(), "")))
'                    writer.WriteLine(dataLine)
'                Next
'            End Using

'            'MessageBox.Show($"CSV file created successfully at: {filePath}")

'            Detail.Clear()
'            'EMAIL SENDING
'            'MessageBox.Show("Detail DataTable cleared successfully.")
'        Catch ex As Exception
'            'MessageBox.Show("Error: " & ex.Message)
'        End Try
'    End Sub

'    Private Function DataExistsInDatabase() As Boolean
'        Try
'            Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J" ' Replace with your actual connection string

'            Using connection As New SqlConnection(connectionString)
'                connection.Open()

'                For Each row As DataRow In Detail.Rows
'                    Dim query As String = "SELECT * FROM TRN_Microchip_csv_logs WHERE " &
'                                      "Supplier_Name = @Supplier_Name AND " &
'                                      "Stage = @Stage AND " &
'                                      "Step = @Step AND " &
'                                      "PartCode = @PartCode AND " &
'                                      "Bin = @Bin AND " &
'                                      "Ship_To = @Ship_To AND " &
'                                      "Lot_Number = @Lot_Number AND " &
'                                      "Lot_wafer_Qty = @Lot_wafer_Qty AND " &
'                                      "Lot_Qty = @Lot_Qty AND " &
'                                      "Wafer_ID = @Wafer_ID AND " &
'                                      "Action_Code = @Action_Code AND " &
'                                      "Child_lot = @Child_lot AND " &
'                                      "Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND " &
'                                      "Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND " &
'                                      "Transaction_Date = @Transaction_Date AND " &
'                                      "Lot_Type = @Lot_Type AND " &
'                                      "Customer_Code = @Customer_Code"

'                    Using command As New SqlCommand(query, connection)
'                        command.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
'                        command.Parameters.AddWithValue("@Stage", row("Stage"))
'                        command.Parameters.AddWithValue("@Step", row("Step"))
'                        command.Parameters.AddWithValue("@PartCode", row("PartCode"))
'                        command.Parameters.AddWithValue("@Bin", row("Bin"))
'                        command.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
'                        command.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
'                        command.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
'                        command.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
'                        command.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
'                        command.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
'                        command.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
'                        command.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
'                        command.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
'                        command.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
'                        command.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
'                        command.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

'                        Dim rowCount As Integer = Convert.ToInt32(command.ExecuteScalar())

'                        If rowCount > 0 Then
'                            Return True
'                        End If
'                    End Using
'                Next
'                Return False
'            End Using
'        Catch ex As Exception
'            ' Handle any exceptions
'            MessageBox.Show("Error checking database: " & ex.Message)
'            Return False
'        End Try
'    End Function



'    Private Function CreateMsgBody(ByVal PONo As String, ByVal DNNo As String, ByVal PONumber As String, ByVal RejectQty As String,
'                               ByVal BarcodeOutput As String, ByVal Type As String) As String

'        CreateMsgBody = ""
'        CreateMsgBody = "<html><body><pre>"
'        CreateMsgBody &= "</pre></body></html>"
'    End Function
'End Class



'Imports System.Data.SqlClient
'Imports System.IO

'Public Class Form1
'    Private Detail As DataTable 'Variable to store captured data

'    Private Sub All_Load(sender As Object, e As EventArgs) Handles MyBase.Load
'        Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string
'        Dim folderPath As String = "F:\" ' Specify the folder path where you want to save the CSV file

'        Try
'            Using connection As New SqlConnection(connectionString)
'                connection.Open()

'                ' Execute the query to retrieve data
'                Dim query As String = "usp_TRN_Microchip_csv '2023-04-01 09:09:01.843', '2023-06-15 16:26:42.127'"
'                Using command As New SqlCommand(query, connection)
'                    Using reader As SqlDataReader = command.ExecuteReader()
'                        Dim dataTable As New DataTable()
'                        dataTable.Load(reader)

'                        Detail = dataTable
'                    End Using
'                End Using
'            End Using
'        Catch ex As Exception
'            'MessageBox.Show("Error: " & ex.Message)
'        End Try

'        If Detail IsNot Nothing AndAlso Detail.Rows.Count > 0 Then
'            SaveData()
'        End If

'        Me.Close()
'    End Sub

'    '--------------Test 1--------------
'    'Private Sub SaveData()
'    '    Try
'    '        ' Check if the Detail DataTable contains any rows
'    '        If Detail.Rows.Count = 0 Then
'    '            MessageBox.Show("No data to save.")
'    '            Return
'    '        End If

'    '        Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

'    '        Using connection As New SqlConnection(connectionString)
'    '            connection.Open()

'    '            ' Create a SqlBulkCopy object
'    '            Using bulkCopy As New SqlBulkCopy(connection)
'    '                ' Set the destination table name
'    '                bulkCopy.DestinationTableName = "TRN_Microchip_csv_logs"

'    '                ' Map the columns from the DataTable to the destination table columns
'    '                For Each column As DataColumn In Detail.Columns
'    '                    bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName)
'    '                Next

'    '                ' Write the data from the Detail DataTable to the database
'    '                bulkCopy.WriteToServer(Detail)
'    '            End Using

'    '            MessageBox.Show("Data saved successfully.")
'    '        End Using
'    '    Catch ex As Exception
'    '        MessageBox.Show("Error: " & ex.Message)
'    '    End Try
'    'End Sub

'    '--------------TEST 2'--------------
'    'Private Sub SaveData()
'    '    Try
'    '        Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

'    '        Using connection As New SqlConnection(connectionString)
'    '            connection.Open()

'    '            ' Loop through each row in the Detail DataTable
'    '            For Each row As DataRow In Detail.Rows
'    '                ' Check if the data already exists in the TRN_Microchip_csv_logs table
'    '                Dim existsQuery As String = "SELECT COUNT(*) FROM TRN_Microchip_csv_logs WHERE Supplier_Name = @Supplier_Name AND Stage = @Stage AND Step = @Step AND PartCode = @PartCode AND Bin = @Bin AND Ship_To = @Ship_To AND Lot_Number = @Lot_Number AND Lot_wafer_Qty = @Lot_wafer_Qty AND Lot_Qty = @Lot_Qty AND Wafer_ID = @Wafer_ID AND Action_Code = @Action_Code AND Child_lot = @Child_lot AND Child_Lot_wafer_Qty = @Child_Lot_wafer_Qty AND Child_Lot_Wafer_ID = @Child_Lot_Wafer_ID AND Transaction_Date = @Transaction_Date AND Lot_Type = @Lot_Type AND Customer_Code = @Customer_Code"
'    '                Using existsCommand As New SqlCommand(existsQuery, connection)
'    '                    ' Set parameter values
'    '                    existsCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
'    '                    existsCommand.Parameters.AddWithValue("@Stage", row("Stage"))
'    '                    existsCommand.Parameters.AddWithValue("@Step", row("Step"))
'    '                    existsCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
'    '                    existsCommand.Parameters.AddWithValue("@Bin", row("Bin"))
'    '                    existsCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
'    '                    existsCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
'    '                    existsCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
'    '                    existsCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
'    '                    existsCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
'    '                    existsCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
'    '                    existsCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
'    '                    existsCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
'    '                    existsCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
'    '                    existsCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
'    '                    existsCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
'    '                    existsCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

'    '                    ' Execute the query to check if the data exists
'    '                    Dim rowCount As Integer = Convert.ToInt32(existsCommand.ExecuteScalar())

'    '                    ' If rowCount > 0, data already exists, so skip insertion for this row
'    '                    If rowCount > 0 Then
'    '                        Continue For
'    '                    End If
'    '                End Using

'    '                ' If the data does not exist, the row is inserted into the TRN_Microchip_csv_logs table
'    '                ' Create a SqlCommand object to insert the data
'    '                Dim insertQuery As String = "INSERT INTO TRN_Microchip_csv_logs (Supplier_Name, Stage, Step, PartCode, Bin, Ship_To, Lot_Number, Lot_wafer_Qty, Lot_Qty, Wafer_ID, Action_Code, Child_lot, Child_Lot_wafer_Qty, Child_Lot_Wafer_ID, Transaction_Date, Lot_Type, Customer_Code) VALUES (@Supplier_Name, @Stage, @Step, @PartCode, @Bin, @Ship_To, @Lot_Number, @Lot_wafer_Qty, @Lot_Qty, @Wafer_ID, @Action_Code, @Child_lot, @Child_Lot_wafer_Qty, @Child_Lot_Wafer_ID, @Transaction_Date, @Lot_Type, @Customer_Code)"
'    '                Using insertCommand As New SqlCommand(insertQuery, connection)
'    '                    ' Set parameter values
'    '                    insertCommand.Parameters.AddWithValue("@Supplier_Name", row("Supplier_Name"))
'    '                    insertCommand.Parameters.AddWithValue("@Stage", row("Stage"))
'    '                    insertCommand.Parameters.AddWithValue("@Step", row("Step"))
'    '                    insertCommand.Parameters.AddWithValue("@PartCode", row("PartCode"))
'    '                    insertCommand.Parameters.AddWithValue("@Bin", row("Bin"))
'    '                    insertCommand.Parameters.AddWithValue("@Ship_To", row("Ship_To"))
'    '                    insertCommand.Parameters.AddWithValue("@Lot_Number", row("Lot_Number"))
'    '                    insertCommand.Parameters.AddWithValue("@Lot_wafer_Qty", row("Lot_wafer_Qty"))
'    '                    insertCommand.Parameters.AddWithValue("@Lot_Qty", row("Lot_Qty"))
'    '                    insertCommand.Parameters.AddWithValue("@Wafer_ID", row("Wafer_ID"))
'    '                    insertCommand.Parameters.AddWithValue("@Action_Code", row("Action_Code"))
'    '                    insertCommand.Parameters.AddWithValue("@Child_lot", row("Child_lot"))
'    '                    insertCommand.Parameters.AddWithValue("@Child_Lot_wafer_Qty", row("Child_Lot_wafer_Qty"))
'    '                    insertCommand.Parameters.AddWithValue("@Child_Lot_Wafer_ID", row("Child_Lot_Wafer_ID"))
'    '                    insertCommand.Parameters.AddWithValue("@Transaction_Date", row("Transaction_Date"))
'    '                    insertCommand.Parameters.AddWithValue("@Lot_Type", row("Lot_Type"))
'    '                    insertCommand.Parameters.AddWithValue("@Customer_Code", row("Customer_Code"))

'    '                    insertCommand.ExecuteNonQuery() ' Execute the insert command
'    '                End Using
'    '            Next

'    '            MessageBox.Show("Data saved successfully.")
'    '        End Using
'    '    Catch ex As Exception
'    '        MessageBox.Show("Error: " & ex.Message)
'    '    End Try
'    'End Sub

'    Private Sub SaveData()
'        Try
'            Dim connectionString As String = "Server=DESKTOP-6E9LU1F\SQLEXPRESS;Database=MES_ATEC;User Id=sa;Password=18Bz23efBd0J;" ' Replace with your actual connection string

'            Using connection As New SqlConnection(connectionString)
'                connection.Open()


'                For Each row As DataRow In Detail.Rows

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

'                                    insertCommand.ExecuteNonQuery() ' Execute the insert command
'                                End Using
'                            End If
'                        End Using
'                    End Using
'                Next

'                MessageBox.Show("Data saved successfully.")
'            End Using
'        Catch ex As Exception
'            MessageBox.Show("Error: " & ex.Message)
'        End Try
'    End Sub


'End Class



