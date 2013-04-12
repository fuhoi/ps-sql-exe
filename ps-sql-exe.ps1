# FILE: ps-sql-exe.ps1
# AUTHOR: Adam Lincoln
# DATE: 2013-Apr-13
# PURPOSE: Command line wrapper for PowerShell ODBC connections
# NOTES:
# Command line parameters
# Sample usage: ps-sql-exe "DSN=MYOB_ODBC_SYS" "SELECT * FROM Sales"
# Sample ODBC connection string: "DSN=MYOB_ODBC_SYS; TYPE=MYOB; ACCESS_TYPE=READ; DRIVER_COMPLETION=DRIVER_NOPROMPT; NETWORK_PROTOCOL=NONET;"

# Param(
#     [Parameter(Mandatory=$True, Position=1)]
#     [string]$connection_string,
#     
#     [Parameter(Mandatory=$True, Position=2)]
#     [string]$sql_command
# )

# Variables
$dir_working = Split-Path $MyInvocation.MyCommand.Path
$file_current = $MyInvocation.MyCommand.Name
$file_this = Join-Path $dir_working $file_current

# Helpers
Set-Alias ps64 "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe" 
Set-Alias ps32 "$env:windir\syswow64\WindowsPowerShell\v1.0\powershell.exe"

Function Get-ProcessorArchitecture {
    Return $env:Processor_Architecture
}

# Support for 32-bit drivers on a 64-bit os
Function Set-ProcessorArchitecture {
    Param([string]$bitness_required = "x86")
    $bitness_current = Get-ProcessorArchitecture
    If ($bitness_required -eq $bitness_current) {
        Write-Host "Processor Architecture set"
    } Else {
        If ($bitness_required -eq "AMD64") {
            Write-Host "Attempting to launch 64bit process..."
            ps64 -file $file_this
        } ElseIf ($bitness_required -eq "x86") {
            Write-Host "Attempting to launch 32bit process..."
            ps32 -file $file_this
        } Else {
            Write-Error "Not Supported"
        }
    }
}

Function Execute-Command {
    Param(
        [Parameter(Mandatory=$True, Position=1)]
        [string]$connection_string,
        
        [Parameter(Mandatory=$True, Position=2)]
        [string]$sql_command
    )

    Write-Host "Creating connection..."
    $db_conn = New-Object System.Data.Odbc.OdbcConnection
    $db_conn.ConnectionString = $connection_string
    $db_conn.Open()
    
    Write-Host "Creating command..."
    $db_cmd = New-Object System.Data.Odbc.OdbcCommand
    $db_cmd.Connection = $db_conn
    $db_cmd.CommandTimeout = 300
    $db_cmd.CommandText = $sql_command
    
    Write-Host "Creating data handlers..."
    $da = New-Object System.Data.Odbc.OdbcDataAdapter($db_cmd)
    # $ds = New-Object System.Data.DataSet
    $dt = New-Object System.Data.DataTable
    
    Write-Host "Executing SQL..."
    # [void]$da.Fill($ds)
    [void]$da.Fill($dt)
    
    Write-Host "Displaying data..."
    $dt | Format-Table -AutoSize -Wrap
    
    Write-Host "Closing connection..."
    $db_conn.close()
}

Get-ProcessorArchitecture

If ($(Get-ProcessorArchitecture) -ne "x86") {
    Set-ProcessorArchitecture "x86"
} Else {
    Execute-Command
}
