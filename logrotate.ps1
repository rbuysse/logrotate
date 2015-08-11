# Sort of an analogue to Linux's logrotate
# Recursively zips and deletes files from an arbitrary path
# Works best running once per day as a scheduled task
# Originally written to manage IIS logfiles


$logfile = "f:\path\to\logfile.txt" # Output logfile location for this script
$machinefile = "F:\path\to\serverlist.txt" # Path to the list of servers with logs you want to manage

try {
    $machinelist = get-content $machinefile
} catch {
    write-output "Unable to open $machinefile"
}

$foo = date
write-output "" > $logfile
write-output "-------------" >> $logfile
write-output "starting to work at $foo " >> $logfile

ForEach ($server in $machinelist) {
    $results = ""
    write-output "" >> $logfile
    write-output "working on $server" >> $logfile 
    try {
        Test-WSMan $server -ErrorAction Stop
    } catch {
        Write-Output "Unable to connect to $server" >> $logfile
        continue
    }
    $results = Invoke-Command -ComputerName $server -ScriptBlock {
        $logDays = -2
        $compressDays = -14

        # IsFileLocked takes one argument, a file path, and returns true or false depending on the file's locked state.
        function IsFileLocked([string] $path) {
            If ([string]::IsNullOrEmpty($path) -eq $true){
                Throw "The path must be specified."
            }
            [bool] $fileExists = Test-Path $path    
            If ($fileExists -eq $false) {
                Throw "File does not exist (" + $path + ")"
            }    
            [bool] $isFileLocked = $true
            $file = $null
            Try {
                $file = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
                $isFileLocked = $false
            }

            Catch [IO.IOException] {
                If ($_.Exception.Message.EndsWith("it is being used by another process.") -eq $false) {
                    Throw $_.Exception
                }
            }
            Finally {
                If ($file -ne $null) {
                    $file.Close()
                }
            }
            return $isFileLocked
        }
        # Out-zip function takes one argument, a file object, and adds it to a compressed archive of the same name. it kinda sucks.
        function Out-Zip($item) {
            if ($item -ne $null) {
                $zipFileName = $item.fullname + ".zip"
                Set-Content $zipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) -Force

                $zipFile = Get-Item $zipFileName
                $zipFile.IsReadOnly = $false

                $shell = New-Object -ComObject 'shell.application'
                $zip = $shell.NameSpace($zipFileName)
                $zip.CopyHere(($item.fullname))
                start-sleep -Seconds 1
                while (IsFileLocked($item.fullname) -eq $true) {
                    Start-Sleep -Seconds 2
                    write-host "$item.fullname IS LOCKED"
                }
            }
        }

        $path = "f:\logs\"
            $logObjects = Get-ChildItem -recurse $path| ?{$_.psiscontainer -eq $False} 
            ForEach($file in $logObjects) {
                # Look for uncompressed logfiles older than $logDays and first compress and then delete the original.
                if (($file.LastWriteTime -lt (Get-Date).AddDays($logDays)) -and ($file.Extension -ne ".zip")) {
                    Out-Zip $file
                    try {
                        rm $file.fullname
                    } catch {
                        Write-Output "Unable to delete $($file.fullname)"
                    }
                }
                elseif (($file.LastWriteTime -lt (Get-Date).AddDays($compressDays)) -and ($file.Extension -eq ".zip")) {
                    $file.IsReadOnly = $false
                    try {
                        rm $file.fullname
                    } catch {
                        Write-Output "Unable to delete $($file.fullname)"
                    }
                }
            }
        }
    write-output $results >> $logfile
}

