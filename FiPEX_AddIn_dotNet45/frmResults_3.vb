Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CatalogUI
Imports System.IO
Imports System.Windows.Forms
Imports x = Microsoft.Office.Interop.Excel
Imports System.Linq


Public Class frmResults_3

    Private Sub cmdExportCSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExportXLS.Click

        ' Prompt the user for directory
        ' Check write permissions to directory
        ' Prompt the user for tablename
        ' Check that tablename does not exist

        ' export datagridview


        Dim sDirectory As String

        Dim fdlg As FolderBrowserDialog = New FolderBrowserDialog
        fdlg.Description = "Browse to Output Directory"
        fdlg.RootFolder = System.Environment.SpecialFolder.MyComputer
        If fdlg.ShowDialog = Windows.Forms.DialogResult.OK Then

            ' Check that the user currently has file permissions to write to 
            ' this directory
            Dim bPermissionCheck As Boolean
            bPermissionCheck = FileWriteDeleteCheck(fdlg.SelectedPath)
            If bPermissionCheck = False Then
                MsgBox("File / folder permission check: " & Str(bPermissionCheck))
                MsgBox("It appears you do not have write permission to the Output Directory.  Write permission to this directory is needed.")
                sDirectory = ""
                Exit Sub
            End If

            ' Check that the user currently has file permissions to write to 
            ' this directory
            'Dim pc As New CheckPerm
            'pc.Permission = "Modify"
            'If Not pc.CheckPerm(fdlg.SelectedPath) Then
            '    MsgBox("You don't currently have permission to write files to that directory.")
            '    Exit Sub
            'End If

            ' Set the path in the text dialogue to save to extension stream
            sDirectory = fdlg.SelectedPath
        End If


        ' GET THE TABLE NAME AND CHECK NAME
        Dim defaultValue As String
        Dim myTableName As Object
        Dim sTableName As String
        Dim Message As String = "Please choose table name"

        defaultValue = ""
        myTableName = InputBox(Message, "Choose Table Name", defaultValue)


        sTableName = myTableName.ToString
        If sTableName = "" Then
            MsgBox("You must type a table name")
            myTableName = InputBox(Message, "Choose Table Name", defaultValue)
            If sTableName = "" Then
                Exit Sub
            End If
            sTableName = myTableName.ToString
        End If

        ' check name exists.  
        Dim sFileName As String
        sFileName = sDirectory + "/" + sTableName + ".xls"
      
        Dim iColumnCount As Integer = 0
        Dim iRowCount As Integer = 0

        Try
            iColumnCount = DataGridView1.Columns.Count
        Catch ex As Exception
            MsgBox("Could not get column count")
        End Try

        Try
            iRowCount = DataGridView1.Rows.Count
        Catch ex As Exception
            MsgBox("Could not get row count")
        End Try

     


        ' http://support.microsoft.com/kb/306022
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object
        Try
            'Start a new workbook in Excel.
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add

            'Add data to cells of the first worksheet in the new workbook.
            oSheet = oBook.Worksheets(1)

            For i = 0 To iRowCount - 1
                For j = 0 To iColumnCount - 1
                    oSheet.Cells(i + 1, j + 1).Value = DataGridView1.Rows(i).Cells(j).Value
                Next
            Next

            'oSheet.Range("B1").Value = "First Name"
            'oSheet.Range("A1:B1").Font.Bold = True
            'oSheet.Range("A2").Value = "Doe"
            'oSheet.Range("B2").Value = "John"

            'Save the Workbook and quit Excel.
            oBook.SaveAs(sFileName)
            oSheet = Nothing
            oBook = Nothing
            oExcel.Quit()
            oExcel = Nothing
        Catch ex As Exception
            MsgBox("Problem exporting to excel." + ex.Message)
        End Try


        MsgBox("Successfully exported. Please ignore warning on open with Excel.")




        '' http://www.codeproject.com/Tips/472706/Export-DataGridView-to-Excel

        ''get all visible columns in display index order
        'Dim ColNames As List(Of String) = (From col As DataGridViewColumn _
        '                                   In DataGridView1.Columns.Cast(Of DataGridViewColumn)() _
        '                                   Where (col.Visible = True) _
        '                                   Order By col.DisplayIndex _
        '                                   Select col.Name).ToList
        'Dim colcount = 0
        'Try
        '    Dim excelApp As New Microsoft.Office.Interop.Excel.Application
        '    Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        '    xlWorkBook = excelApp.Workbooks.Add(1)
        '    Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        '    xlWorkSheet = CType(xlWorkBook.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)

        'Catch ex As Exception
        '    MsgBox("Error trying to create excel workbook" + ex.Message)
        'End Try



        'For Each s In ColNames
        '    colcount += 1
        '    xlWorkSheet.Cells(1, colcount) = DataGridView1.Columns.Item(s).HeaderText
        'Next




        'With SaveExcelFileDialog
        '    .Filter = "Excel|*.xlsx"
        '    .Title = "Save griddata in Excel"
        '    If .ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
        '        Dim o As New ExcelExporter
        '        Dim b = o.Export(DataGridView1, .FileName)
        '    End If
        '    .Dispose()
        'End With





        '' Create File and output
        'Dim sw As StreamWriter
        'sw = File.CreateText(sDirectory + "/" + sTableName + ".csv")

    


        'sw.WriteLine("data;")
        'sw.Close()





    End Sub
    Public Function FileWriteDeleteCheck(ByVal sDCIOutputDir As String) As Boolean

        Dim FILE_NAME As String = "FiPExPermTEST1.txt"
        If File.Exists(sDCIOutputDir + "\" + FILE_NAME) Then
            MsgBox("tempmsg: this is the file name tested: " + sDCIOutputDir + FILE_NAME)
            MsgBox("test file already exists in DCI output directory")
        End If

        Try
            Dim path As String = sDCIOutputDir + "\" + FILE_NAME
            Dim sw As StreamWriter = File.CreateText(path)
            sw.Close()

            ' Ensure that the target does not exist.
            File.Delete(path)

            Return True

        Catch e As Exception
            MsgBox("The following exception was found: " & e.Message)
            Return False
        End Try

    End Function
End Class
