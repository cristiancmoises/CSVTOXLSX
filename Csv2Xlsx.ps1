gci "C:\Example1.csv", "C:\Users\=.csv" | %{ #Location of your files or folder
$Path = $_.DirectoryName
$filename = $_.BaseName

#Define locations and delimiter
$csv = $_.FullName #Location of the source file
#$xlsx = "$Path/$filename.xlsx" # Names & saves Excel file same name/location as CSV
$xlsx = "C:\$filename.xlsx" # Names Excel file same name as CSV

$delimiter = ";" #Specify the delimiter used in the file

# Create a new Excel workbook with one empty sheet
$excel = New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)


# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()

#Save and Close the Workbook as XLSX
$excel.DisplayAlerts = $False
$Workbook.SaveAs($xlsx,51)
$excel.Quit()

}