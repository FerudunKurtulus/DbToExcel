Public Class GetDatas
    'My database connection string
    Dim connection As New SqlClient.SqlConnection("server=Sen-ERP;Database=SENSAC_ERP_DB;Integrated Security=True;Connect Timeout=30")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Create excel file
        Dim xlApp As New Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        xlBook = xlApp.Workbooks.Add()
        'Get all tables name from database.If you dont want to include some tables you can use other code and add table names
        Dim adp As New SqlClient.SqlDataAdapter("select name from sys.tables", connection)
        'Dim adp As New SqlClient.SqlDataAdapter("select name from sys.tables where name not in ('models','users')", connection)
        Dim f As New DataTable
        adp.Fill(f)

        For i As Integer = 0 To f.Rows.Count - 1
            adp.Dispose()
            Dim g As New DataTable
            'Get table datas from database.I want just 10 rows from tables.If you want all rows you can remove 'top 10' from sqlstring
            adp = New SqlClient.SqlDataAdapter("select top 10* from [" & f.Rows(i)("name").ToString & "]", connection)
            adp.Fill(g)
            'Add page to excel workbook and name it with table name
            xlSheet = xlBook.Sheets.Add()
            xlSheet.Name = f.Rows(i)("name").ToString
            For k As Integer = 0 To g.Columns.Count - 1
                xlSheet.Cells(1, k + 1) = g.Columns(k).ColumnName.ToString
                For j As Integer = 0 To g.Rows.Count - 1
                    xlSheet.Cells(j + 2, k + 1) = g.Rows(j)(k).ToString
                Next
            Next
            'Give data columns some back color,font size etc.
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, g.Columns.Count)).Interior.Color = Color.FromArgb(253, 233, 217)
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, g.Columns.Count)).Font.Color = Color.Black
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, g.Columns.Count)).Font.Name = "Calibri"
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, g.Columns.Count)).Font.FontStyle = FontStyle.Bold
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, g.Columns.Count)).Font.Size = 13
            'Give data rows some back color,font size etc.
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).Interior.Color = Color.White
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).Font.Color = Color.Black
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).Font.Name = "Calibri"
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).Font.FontStyle = FontStyle.Regular
            xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).Font.Size = 11
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).EntireColumn.AutoFit()
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).HorizontalAlignment = Excel.Constants.xlCenter
            xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(g.Rows.Count + 1, g.Columns.Count)).VerticalAlignment = Excel.Constants.xlCenter
        Next
        'Save dialog 
        Dim sfd As New SaveFileDialog
        With sfd
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            .FileName = "databaseExcel"
            .Filter = "Excel Dosyaları |*.xls|Excel Dosyaları |*.xlsx"
            .Title = "Save File"
        End With
        'Call save dialog and check if user terminated.If user end it dialog code stops there.
        Dim rslt = sfd.ShowDialog()
        If rslt = DialogResult.No Or rslt = DialogResult.Cancel Then Exit Sub
        'Get file location.
        Dim save As String = sfd.FileName
        'Save excel workbook and close excel
        xlBook.SaveAs(save)
        xlBook.Close()
        xlApp.Quit()
        MsgBox("Excel file created.")

    End Sub
End Class