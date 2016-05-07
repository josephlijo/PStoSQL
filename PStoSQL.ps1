# Name: PS to SQL (PowerShell to SQL) 
# Author: LJ
# Date: 07-May-2016
# URL: https://github.com/ljnotes/PStoSQL.git
# Description: PS script to make connection to SQL Server database, read data and write data back based on the passed in configuration. 

# Uncomment the following if the script should be run in admin mode
# #Requires -RunAsAdministrator

# Variables - begins

# Set the Server name here 
# Named instances can be set like servername\instancename 
$SQLServer = "MyServer\MyInstance"

# Set the database name 
$SQLDatabase = "master"

# Set the connection string
$SQLConnectionString = $null

# Set the SQL query here
$SQLQuery = "SELECT * FROM dbo.syscomments(NOLOCK)"

# SQL Connection, Command, Adapter and dataset object 
$SQLConnectionObj = $null
$SQLCommandObj = $null
$SQLDataAdapterObj = $null
$SQLDatasetObj = $null

# Variables - ends

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

    Write-Host 'Connection string is: ' $SQLConnectionString

    # Create SQL command object
    $SQLCommandObj = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommandObj.CommandText = $SQLQuery;
    $SQLCommandObj.Connection = $SQLConnectionObj

    # Create adapter to the database 
    $SQLDataAdapterObj = New-Object System.Data.SqlClient.SqlDataAdapter
    $SQLDataAdapterObj.SelectCommand = $SQLCommandObj;

    # Create Data-set to store the data and get the data using adapter
    $SQLDatasetObj = New-Object System.Data.DataSet
    $SQLDataAdapterObj.Fill($SQLDatasetObj)

    # Clean-up
    $SQLConnectionObj.Close()
}
catch 
{
    'An error occured while connecting to the Server - '  + $SQLServer + ' - with connection string - ' + $SQLConnectionString + '; Exception: - ' + $_.Exception.Message
}