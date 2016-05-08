# Name: PS to SQL (PowerShell to SQL) 
# Author: LJ
# Date: 07-May-2016
# URL: https://github.com/ljnotes/PStoSQL.git
# Description: PS script to make connection to SQL Server database, read data and write data back based on the passed in configuration. 
#              The script also includes the functionality to save the data to Excel file 

# Uncomment the following if the script should be run in admin mode
# #Requires -RunAsAdministrator

param
(
    ## 
    ## Database information
    ## 

    # Set the Server name here 
    # Named instances can be set like servername\instancename 
    [string]$SQLServer = "MyServer\MyInstance"

    # Set the database name
    ,[string]$SQLDatabase = "master" 

    # Set the connection string
    ,[string]$SQLConnectionString = $null

    # Set the SQL query here 
    ,[string]$SQLQuery = "SELECT TOP 33 * FROM dbo.syscomments(NOLOCK)"

    ## 
    ## File and Directory information
    ##  

    # Should save as file
    ,[bool]$ResultToExcel = $true 

    # Directory where the file should be saved to
    ,[string]$ResultDirectory = 'C:\Temp\'

    # File name 
    ,[string]$ResultFileName = "MyDataSet"
    
) 

## SQL Connection, Command, Adapter and dataset object 
$SQLConnectionObj = $null
$SQLCommandObj = $null
$SQLDataAdapterObj = $null
$DatasetObj = $null
$DatatableObj = $null

# Make connection to the SQL Server and get the data
try
{
    # Create Sql Connection object
    $SQLConnectionObj = New-Object System.Data.SqlClient.SqlConnection

    # Set the connection string; formulate if not provided
    if(![string]::IsNullOrEmpty($SQLConnectionString))
    {
        $SQLConnectionObj.ConnectionString = $SQLConnectionString
    }
    elseif (![string]::IsNullOrEmpty($SQLServer) -and ![string]::IsNullOrEmpty($SQLDatabase))
    {
        $SQLConnectionString = "Server=$SQLServer;Database=$SQLDatabase;Integrated Security=True"
        $SQLConnectionObj.ConnectionString = $SQLConnectionString
    }

    # Create SQL command object
    $SQLCommandObj = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommandObj.CommandText = $SQLQuery;
    $SQLCommandObj.Connection = $SQLConnectionObj

    # Create adapter to the database 
    $SQLDataAdapterObj = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLDataAdapterObj.SelectCommand = $SQLCommandObj;

    # Create Data-set to store the data and get the data using adapter
    $DatasetObj = New-Object System.Data.DataSet
    $SQLDataAdapterObj.Fill($DatasetObj)
    $DatatableObj = $DatasetObj.Tables[0]

    # Clean-up
    $SQLConnectionObj.Close()
}
catch 
{
    'An error occured while connecting to the Server - '  + $SQLServer + ' - with connection string - ' + $SQLConnectionString + '; Exception: - ' + $_.Exception.Message
}

# Export the data to Excel 
if($ResultToExcel) 
{
    try
    {
        ## Variables 
        $ExcelObj = $null
        $ExcelWorkBook = $null
        $ExcelWorkSheet = $null
        $ExcelWorkSheetCurrent = 1
        $ExcelFilePath = $null
        $ExcelDataColumns = $null
        $ExcelRowOffset = 1
        $ExcelColOffset = 1

        # Create Excel Application object 
        $ExcelObj = New-Object -ComObject Excel.Application
        # 1 / True = Visible; 0 / False = No
        $ExcelObj.Visible = $false

        # Create the Work book and sheet 
        $ExcelWorkBook = $ExcelObj.Workbooks.Add() 
        $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item($ExcelWorkSheetCurrent)

        # Set the font and formatting
        $ExcelWorkSheet.Cells.Font.Name = "Calibri"
        $ExcelWorkSheet.Cells.Font.Size = 10

        ## Write the data to the sheet

        # Create the header with columns in the result set 
        $ExcelDataColumns = $DatatableObj.Columns
        foreach($DataColumn in $ExcelDataColumns)
        {
            # Set the font weight to bold and write the column name in the cell
            $ExcelWorkSheet.Cells($ExcelRowOffset, $ExcelColOffset).Font.Bold = $true
            $ExcelWorkSheet.Cells($ExcelRowOffset, $ExcelColOffset) = $DataColumn.ColumnName
            $ExcelColOffset++
        }

        # Write the data rows
        $ExcelDataRows = $DatatableObj.Rows
        foreach($DataRow in $ExcelDataRows)
        {
            # Set row and column
            $ExcelRowOffset++
            $ExcelColOffset = 1

            foreach($item in $DataRow.ItemArray)
            {
                $ExcelWorkSheet.Cells($ExcelRowOffset, $ExcelColOffset) = $item.ToString()
                $ExcelColOffset++
            }
        }

        ## Save the file

        $ExcelFilePath = "$ResultDirectory$ResultFileName.xlsx"

        # Delete the file if it already exists 
        if(Test-Path $ExcelFilePath)
        {
            Remove-Item $ExcelFilePath
        }
        $ExcelWorkBook.SaveAs($ExcelFilePath)

        ## Clean-up

        $ExcelWorkBook.Close()
        $ExcelObj.Quit()
    }
    catch
    {
        'An error occured while saving the file - ' + $_.Exception.Message
    }
}
