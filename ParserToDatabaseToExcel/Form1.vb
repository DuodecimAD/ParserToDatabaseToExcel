Imports System.Data.SqlClient
Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types



Public Class Form1

    Public Class Animal
        Public Property Name As String
        Public Property Size_cm As Integer
        Public Property Country As String
    End Class

    Private Sub OpenFileToParseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileToParseToolStripMenuItem.Click
        ' Create an instance of OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()

        ' Set properties of the OpenFileDialog
        openFileDialog.Title = "Select a JSON File"
        openFileDialog.Filter = "JSON Files|*.json|All Files|*.*" ' Specify file filters for JSON
        'openFileDialog.InitialDirectory = Environment.CurrentDirectory
        openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        ' Set the initial directory

        ' Show the OpenFileDialog and check if the user clicked OK
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file name
            Dim selectedFileName As String = openFileDialog.FileName
            Dim fileContent As String = File.ReadAllText(selectedFileName)

            TextBox1.Text = fileContent

            'ProcessFileContent(fileContent)
        End If
    End Sub


    Dim userId = "TESTDOCKER"
    Dim password = "TESTDOCKER"
    Dim dataSource = "localhost:49161/xe"

    Dim conn As New OracleConnection

    Private Sub SendDB_Click(sender As Object, e As EventArgs) Handles SendDB.Click
        ' Assuming you have a function to create the JSON string from your data
        Dim jsonString As String = TextBox1.Text

        Dim animalsList As List(Of Animal) = JsonConvert.DeserializeObject(Of List(Of Animal))(jsonString)

        Try
            ' Initialize the connection
            conn.ConnectionString = "User Id=" + userId + ";Password=" + password + ";Data Source=" + dataSource + ";"
            conn.Open()
            info.Text = "info : Connected"

            For Each animal In animalsList
                Dim insertCommand As New OracleCommand("INSERT INTO Animals (Name, Size_cm, Country) VALUES (:Name, :Size_cm, :Country)", conn)

                ' Add parameters
                insertCommand.Parameters.Add("Name", OracleDbType.Varchar2).Value = animal.Name
                insertCommand.Parameters.Add("Size_cm", OracleDbType.Int32).Value = animal.Size_cm
                insertCommand.Parameters.Add("Country", OracleDbType.Varchar2).Value = animal.Country

                ' Execute the command
                insertCommand.ExecuteNonQuery()
            Next
            info.Text = "info : Connected : json has been written in the database"
        Catch ex As Exception
            info.Text = "info : Error: " & ex.Message
            Return
        Finally
            ' Close the connection in the finally block to ensure it's closed even if an exception occurs
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub GetDB_Click(sender As Object, e As EventArgs) Handles GetDB.Click

        Try
            conn.ConnectionString = "User Id=" + userId + ";Password=" + password + ";Data Source=" + dataSource + ";"
            conn.Open()
            info.Text = "Connected"
        Catch ex As Exception
            info.Text = "Error: " & ex.Message
        End Try

        Dim cmd As New OracleCommand("SELECT * from Animals", conn)
        Dim answer As OracleDataReader = cmd.ExecuteReader()

        While answer.Read()

            Dim names As String = answer.GetString(0)
            Dim size As Integer = answer.GetInt32(1)
            Dim country As String = answer.GetString(2)
            TextBox2.AppendText($"{names}, {size}, {country}{vbCrLf}")

        End While
        conn.Dispose()

    End Sub

    Dim excelApp As New Application()

    Private Sub toExcel_Click(sender As Object, e As EventArgs) Handles toExcel.Click

        ' Specify the path to the Excel file (XLSX) you want to check
        Dim filePath As String = "D:\VisualBasic\ParserToDatabaseToExcel\xlsx\test.xlsx"
        ' Create a new Excel Application and open a workbook


        ' Try to open the workbook. If it's already open, it will throw an exception.
        Dim workbook As Workbook = excelApp.Workbooks.Open(filePath)
        Try
            'workbook = excelApp.Workbooks.Open(filePath)

            ' Access a specific worksheet, e.g., "Sheet1"
            Dim worksheet As Worksheet = workbook.Sheets("Sheet1")

            ' Find the next available row in column 1 (A)
            Dim lastRow As Integer = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row
            lastRow = lastRow + 1

            ' Write data to the next available row in column 1 (A)
            'worksheet.Cells(lastRow + 1, 1).Value = TextBox2.Text

            ' Split the input string into lines
            Dim lines() As String = TextBox2.Text.Split(vbCrLf)

            TextBox3.Text = "Added : " + vbCrLf

            Dim columnName = ""
            worksheet.Cells(1, 1).Value = "Name"
            worksheet.Cells(1, 2).Value = "Size"
            worksheet.Cells(1, 3).Value = "Country"

            ' Write each line to a separate row in the Excel sheet
            For i As Integer = 0 To lines.Length - 2
                ' Split each line into individual elements based on commas
                Dim elements() As String = lines(i).Split(",")

                ' Write each element to a separate cell in the same row
                For j As Integer = 0 To elements.Length - 1



                    ' Use j index for elements array, not i
                    worksheet.Cells(lastRow + i, j + 1).Value = elements(j).Trim()

                    Select Case j
                        Case 0
                            columnName = "A"
                        Case 1
                            columnName = "B"
                        Case 2
                            columnName = "C"
                        Case Else
                            columnName = "add more letters in the switch case"
                    End Select

                    TextBox3.AppendText(elements(j).Trim() + " in the cell " + columnName + (lastRow + i).ToString + ", ")
                Next
                TextBox3.AppendText(vbCrLf)
            Next

            'TextBox3.Text = "Added " + TextBox2.Text + " to the excel file in the cell A" + lastRow.ToString

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

    End Sub

End Class


