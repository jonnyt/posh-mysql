<#
    http://www.jaredperry.ca/2009/06/powershell-and-mysql/
    vwiki.co.uk/MySQL_and_PowerShell
#>

Function LoadAssembly
{
    # We need to know the default installation location for 32-bit or 64-bit
    $query = Get-WmiObject -query "Select AddressWidth from Win32_Processor" | Select-Object -First 1

    Try
    {
        if($query.AddressWidth -eq 64)
        {
            $rootFolder = 'C:\Program Files (x86)\MySQL'
        }
        else
        {
            $rootFolder = 'C:\Program Files\MySQL\'
        }
        $library = $rootFolder | Get-ChildItem -Filter 'MySql.Data.dll' -Recurse | Select-Object -First 1
        [void][system.reflection.Assembly]::LoadFrom($($library.FullName))

    }
    Catch [Exception]
    {
        Throw "Could not load the MySQL connector for DotNet.  Please make sure it is installed in the default location.  You can download the installer from http://dev.mysql.com/downloads/connector/net/"
    }
}

Function Invoke-NonQuery
{
    Param(
        [Parameter(Mandatory=$TRUE)][string]$serverName,
        [Parameter(Mandatory=$TRUE)][string]$userName,
        [Parameter(Mandatory=$TRUE)][string]$password,
        [Parameter(Mandatory=$TRUE)][string]$dbName,
        [Parameter(Mandatory=$TRUE)][string]$query
    )


    LoadAssembly
    $dbconnect = New-Object MySql.Data.MySqlClient.MySqlConnection
    $dbconnect.ConnectionString = "server=$serverName;user id=$userName;password=$password;database=$dbName;pooling=false"

    Try
    {
        $dbconnect.Open()
        $sql = New-Object MySql.Data.MySqlClient.MySqlCommand
        $sql.Connection = $dbconnect
        $sql.CommandText = $query
        return $sql.ExecuteNonQuery()
    }
    Catch
    {
        Throw
    }
    Finally
    {
        if($dbconnect -ne $null)
        {
            $dbconnect.Close()
        }
    }
}

Export-ModuleMember -Function Invoke-NonQuery