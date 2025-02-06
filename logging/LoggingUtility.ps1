function WriteLogsToFile {
    param(
        [Parameter(Mandatory=$false)]
        [string]$LogDirectoryPath,
        [Parameter(Mandatory=$true)]
        [string]$LogLabel,
        [Parameter(Mandatory=$true)]
        [scriptblock]$Script
    )

    $now = Get-Date

    $dateString = "$($now.Year)-$($now.Month)-$($now.Day)"
    $timeString = "$($now.Hour):$($now.Minute):$($now.Second)"
    $logDir = if ($LogDirectoryPath -eq $null -or $LogDirectoryPath -eq "") { $env:PS_DEFAULT_LOG_DIR } else { $LogDirectoryPath }    
    

    $logPath = Join-path $logDir "PowerShell-$($dateString).log"

@"
###########################################################
 ---------------------------------------------------------
###########################################################

LOG FOR LABEL -: $LogLabel :-

TIME OF RUN $timeString

===========================================================
 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
===========================================================

"@ >> $logPath

    $logs = ""
    
    & $Script *>&1 | 
    # Tee-Object -FilePath $logPath -Append 
    %{
        $val = $_ 
        $res = $null
        switch ($_.GetType().FullName) {
            'System.Management.Automation.PSObject' {           # Success/Output (1)
                if ($val.ForegroundColor -ne $null) {
                    Write-Host $val -ForegroundColor $val.ForegroundColor
                    "Info: $val" >> $logPath
                }
                else {
                    "Output: $val" >> $logPath
                    $res = $val
                }
            }
            'System.Management.Automation.ErrorRecord' {        # Error (2)
                "Error: $val" >> $logPath
                Write-Error $val
            }
            'System.Management.Automation.WarningRecord' {      # Warning (3)
                "Warning: $val" >> $logPath
                Write-Warning $val
            }
            'System.Management.Automation.VerboseRecord' {      # Verbose (4)
                "Verbose: $val" >> $logPath
                Write-Verbose $val
            }
            'System.Management.Automation.DebugRecord' {        # Debug (5)
                "Debug: $val" >> $logPath
                Write-Debug $val
            }
            'System.Management.Automation.InformationRecord' {  # Information (6)
                "Info: $val" >> $logPath
                if ($val.ForegroundColor -ne $null) { Write-Host $val -ForegroundColor $val.ForegroundColor }
                else { Write-Host $val }
            }
            default {  # Catch any other types
                "Other: $val`n" >> $logPath
                $res = $val
            }
        }

        $res
    }

}
