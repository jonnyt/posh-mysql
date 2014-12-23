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

Function Invoke-Query
{
    Param(
        [Parameter(Mandatory=$TRUE)][string]$serverName,
        [Parameter(Mandatory=$TRUE)][string]$userName,
        [Parameter(Mandatory=$TRUE)][string]$password,
        [Parameter(Mandatory=$TRUE)][string]$dbName,
        [Parameter(Mandatory=$TRUE)][string]$query,
        [Parameter(Mandatory=$FALSE)][string]$queryTimeoutSeconds=60
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
        $sql.CommandTimeout = $queryTimeoutSeconds
        return $sql.ExecuteReader()
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

Function Get-PSObjectFromDataRecords
{
    Param(
        [Parameter(Mandatory=$TRUE)]$records
    )

    if($records.GetType().BaseType.Name -eq "DbDataRecord" -or $records.count -gt 0)
    {
        # property bag for column names
        $properties = @{}
    
        # add properties to the bag
        if($records.GetType().BaseType.Name -eq "DbDataRecord")
        {
            # single record
            for($i=0;$i -lt $records.FieldCount;$i++)
            {
                $properties.Add($records.GetName($i),$null)
            }   
        }
        else
        {
            # multiple records
            for($i=0;$i -lt $records[0].FieldCount;$i++)
            {
                $properties.Add($records[0].GetName($i),$null)
            }
        }

        # iterate results, set up the customobject, pop to the pipeline
        foreach($result in $records)
        {
            $resultObject = New-Object -TypeName PSObject -Property $properties
            for($i=0;$i -lt $result.FieldCount;$i++)
            {
                $resultObject.($result.GetName($i)) = $result.GetValue($i)
            }
            $resultObject
        }
    }
}

Export-ModuleMember -Function Invoke-NonQuery
Export-ModuleMember -Function Invoke-Query
Export-ModuleMember -Function Get-PSObjectFromDataRecords