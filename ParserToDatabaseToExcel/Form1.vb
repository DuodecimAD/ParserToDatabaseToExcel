Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types


Public Class Form1

    Dim userId = "TESTDOCKER"
    Dim password = "TESTDOCKER"
    Dim dataSource = "localhost:49161/xe"

#If False Then

    Dim conn As New OracleConnection

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            conn.ConnectionString = "User Id=" + userId + ";Password=" + password + ";Data Source=" + dataSource + ";"
            conn.Open()
            TextBox1.Text = "Connected"
        Catch ex As Exception
            TextBox1.Text = "Error: " & ex.Message
        End Try

        Dim cmd As New OracleCommand("Select CITY_NAME from CITY", conn)
        Dim answer As OracleDataReader = cmd.ExecuteReader()

        While answer.Read()

            Dim cityName As String = answer.GetString(0)
            TextBox1.AppendText(cityName)

        End While
        conn.dispose()
    End Sub


    Dim excelApp As New Application()

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' Specify the path to the Excel file (XLSX) you want to check
        Dim filePath As String = "D:\VisualBasic\databaseConnectVB\xlsx\test.xlsx"
        ' Create a new Excel Application and open a workbook


        Label1.Text = ""


        ' Try to open the workbook. If it's already open, it will throw an exception.
        Dim workbook As Workbook = excelApp.Workbooks.Open(filePath)
        Try
            'workbook = excelApp.Workbooks.Open(filePath)

            ' Access a specific worksheet, e.g., "Sheet1"
            Dim worksheet As Worksheet = workbook.Sheets("Sheet1")

            ' Find the next available row in column 1 (A)
            Dim lastRow As Integer = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row

            ' Write data to the next available row in column 1 (A)
            worksheet.Cells(lastRow + 1, 1).Value = TextBox2.Text

            Label1.Text = "Added " + TextBox2.Text + " to the excel file in the cell A" + lastRow.ToString

            ' Print a message to the console
            Console.WriteLine(lastRow)

            ' Save the workbook with the changes
            workbook.Save()

            ' Close and release Excel objects
            workbook.Close()
            'excelApp.Quit()
            Marshal.ReleaseComObject(workbook)
            'Marshal.ReleaseComObject(excelApp)
        Catch ex As Exception
            ' The workbook is already open or couldn't be opened.
            Console.WriteLine("Excel workbook is already open or cannot be opened.")
            Return
        End Try



        ' Write data to a specific cell, e.g., cell A1
        ' worksheet.Cells(1, 1).Value = TextBox2.Text

        ''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''

    End Sub

#End If

    Private Sub OpenFileToParseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileToParseToolStripMenuItem.Click
        ' Create an instance of OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()

        ' Set properties of the OpenFileDialog
        openFileDialog.Title = "Select a JSON File"
        openFileDialog.Filter = "JSON Files|*.json|All Files|*.*" ' Specify file filters for JSON
        openFileDialog.InitialDirectory = Environment.CurrentDirectory
        ' Set the initial directory

        ' Show the OpenFileDialog and check if the user clicked OK
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file name
            Dim selectedFileName As String = openFileDialog.FileName
            Dim fileContent As String = File.ReadAllText(selectedFileName)

            ProcessFileContent(fileContent)
        End If
    End Sub

    Private Sub ProcessFileContent(jsonString As String)
        TextBox1.Text = jsonString

        ' Deserialize the JSON string
        Dim jsonArray As JArray = JsonConvert.DeserializeObject(jsonString)

        ' Convert the JArray to a string and remove square brackets and curly braces
        Dim cleanJsonString As String = jsonArray.ToString().Replace("[", "").Replace("]", "").Replace("{", "").Replace("}", "")
        TextBox2.Text = cleanJsonString
    End Sub


    Dim conn As New OracleConnection

    Private Sub SendDB_Click(sender As Object, e As EventArgs) Handles SendDB.Click
        Try
            conn.ConnectionString = "User Id=" + userId + ";Password=" + password + ";Data Source=" + dataSource + ";"
            conn.Open()
            info.Text = "info : Connected"
        Catch ex As Exception
            info.Text = "info : Error: " & ex.Message
            Return
        End Try

        ' Assuming you have a function to create the JSON string from your data
        Dim jsonString As String = TextBox1.Text

        ' Call the stored procedure with the JSON parameter
        Try
            Using cmd As New OracleCommand("InsertJsonData", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("p_json_string", OracleDbType.Clob).Value = jsonString
                cmd.ExecuteNonQuery()
            End Using

            info.Text = "info : JSON data sent successfully."
        Catch ex As Exception
            info.Text = "info : Error: " & ex.Message
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
