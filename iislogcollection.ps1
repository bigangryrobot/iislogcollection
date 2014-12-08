while($true){
    $core ={
        $DirBase = "\\somefileserver"
        $FileAgeLimit = (Get-Date).AddDays(-5)
        $ServersFile =  "\\somefileserver.somecompany\servers.txt"
        $SysLogfile = $DirBase + "\syslog\syslog.log"
        $IISLogLocation = $DirBase + "\temp"
        $ProcessedIISLogLocation = $DirBase + "\processed\*.log"
        $ArchiveLocation     = $DirBase + "\archive\"
        $ArchiveSysLogfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesyslog.rar"
        $ArchiveSysLogTempfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesyslog.txt"
        $ArchiveIISLogfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesIIS.rar"
        $WinZip = $DirBase + "mechanism\7-Zip\7z.exe"
        $Debug = 0
        $TotalFilesProcessed = 0
        $MaxLogSizeInMB=20
        $SleepTimer=100
        $threadTimeout=500
        $MaxThreads=25
        $RunSummary=$null
        $TotalServers=0
        $TotalFiles=0
        $DeleteBacklogFiles=$true

        while($true)
        {
            
            Function SortRandom {
                process {
                    [array]$x = $x + $_
                }
                end {
                    $x | sort-object {(new-object Random).next()}
                }
            }
            
            Function LogWrite {
                Param ([string]$Logstring,[string]$threadID,[string]$LogLevel,[string]$TimeTaken)
                If ($siteName -eq $null) {
                    $siteName = "    "
                }
                If ($threadID -eq $null) {
                    $threadID = "    "
                }
                $SizeMb="{0:N2}" -f ((Get-ChildItem $SysLogfile| Measure-Object -property length -sum).sum / 1MB) + "MB"
                If ($SizeMb -ge $MaxLogSizeInMB) {
                    Move-Item $SysLogfile $ArchiveSysLogTempfile -force
                    $CommandOuput = Invoke-Expression "$($WinZip ) a -dw $ArchiveSysLogfile $ArchiveSysLogTempfile" 2>&1
                    } Else {
                    $MaxThreads = "{0:000}" -f [int]((Get-Process powershell).count)
                    $TotalFilesProcessed = "{0:0000}" -f [int] (get-content "C:\Temp\TotalFilesProcessed" -EA SilentlyContinue)
                    $elementsprocessed = "{0:00000000}" -f [int] (get-content "C:\Temp\elementsprocessed" -EA SilentlyContinue)
                    $threadID = "{0:000}" -f [int] ($threadID)
                    $runtime =  "{0:0000.00}" -f [int] ('{0:N3}' -f ((Get-Date)-$ScriptStartDTM).totalseconds)
                    $TotalServers = "{0:000}" -f [int] (get-content "C:\Temp\TotalFiles" -EA SilentlyContinue)
                    $TotalFiles= "{0:000}" -f [int] (get-content "C:\Temp\TotalServers" -EA SilentlyContinue)
                    $FilesToServerRatio=00.00

                    
                    If ($Debug -eq 1) {
                        Write-Host $SysLogfile "$(get-date -f MM-dd-yy"  "hh:mm:ss)`t [$($threadID)] `t [$($loglevel)] `t $("{0:N0}" -f ((Get-Date)-$ScriptStartDTM).totalseconds)s`t $($sitename)`t [$TotalFilesProcessed] `t[$elementsprocessed] `t [$MaxThreads]`t $($logstring)"
                        } Else {
                        $Stoploop = $false
                        [int]$Retrycount = "0"
                        while($Stoploop -eq $false) {
                            try {
                                Add-content -force $SysLogfile -value "$(get-date -f MM-dd-yy"  "hh:mm:ss) [$($threadID)]`t[$($loglevel)] $($runtime)s`t$($sitename.padright(24,' '))`t [$TotalFilesProcessed] `t[$elementsprocessed] `t [$MaxThreads]`t [$FilesToServerRatio]`t$($logstring)" -ErrorAction Stop
                                $Stoploop = $true
                            }
                            catch {
                                if ($Retrycount -gt 3){
                                    $Stoploop = $true
                                    } else {
                                    Start-Sleep -Milliseconds 30
                                    $Retrycount = $Retrycount + 1
                                }
                            }
                        }
                    }
                }
            }
            
            $process ={
                
                param ($Object)
                $ThreadStartDTM = (Get-Date)
                $DirBase = "\\somefileserver\windowsiislogs"
                $SysLogfile = $DirBase + "\syslog\syslog.log"
                $FileAgeLimit = (Get-Date).AddDays(-5)
                $IISLogLocation = $DirBase + "\temp"
                $ProcessedIISLogLocation = $DirBase + "\processed\*.log"
                $ArchiveLocation     = $DirBase + "\archive\"
                $ArchiveSysLogfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesyslog.rar"
                $ArchiveSysLogTempfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesyslog.txt"
                $ArchiveIISLogfile = $DirBase + "\archive\$(get-date -f MM-dd-hh-mm)archivesIIS.rar"
                $WinZip = $DirBase + "mechanism\7-Zip\7z.exe"
                $Debug = 0
                $TotalFilesProcessed = 0
                $MaxLogSizeInMB=5
                $SleepTimer=40
                $threadTimeout=400
                $ProtectedThreadID=Get-Random -maximum 500
                $DeleteBacklogFiles=$true
                
                Function LogWrite {
                    Param ([string]$Logstring,[string]$threadID,[string]$sitename,[string]$LogLevel,[string]$TimeTaken)
                    If(!(Test-Path -Path $SysLogfile)) {
                        new-item -Path $SysLogfile -Value "new file" –itemtype file
                    }
                    $SizeMb="{0:N2}" -f ((Get-ChildItem $SysLogfile| Measure-Object -property length -sum).sum / 1MB) + "MB"
                    If ($SizeMb -ge $MaxLogSizeInMB) {
                        Move-Item $SysLogfile $ArchiveSysLogTempfile -force
                        $CommandOuput = Invoke-Expression "$($WinZip ) a -dw $ArchiveSysLogfile $ArchiveSysLogTempfile" 2>&1
                        } Else {
                        
                        $MaxThreads = "{0:000}" -f [int]((Get-Process powershell).count)
                        $TotalFilesProcessed = "{0:0000}" -f [int] (get-content "C:\Temp\TotalFilesProcessed"-ErrorAction SilentlyContinue)
                        $elementsprocessed = "{0:00000000}" -f [int] (get-content "C:\Temp\elementsprocessed"-ErrorAction SilentlyContinue)
                        $threadID = "{0:000}" -f [int] ($threadID)
                        $runtime =  "{0:0000.00}" -f [int] ('{0:N3}' -f ((Get-Date)-$ThreadStartDTM).totalseconds)
                        $TotalServers = "{0:000}" -f [int] (get-content "C:\Temp\TotalFiles"-ErrorAction SilentlyContinue)
                        $TotalFiles= "{0:000}" -f [int] (get-content "C:\Temp\TotalServers"-ErrorAction SilentlyContinue)
                        $FilesToServerRatio=00.00
                        
                        If ($Debug -eq 1) {
                            Write-Host $SysLogfile "$(get-date -f MM-dd-yy"  "hh:mm:ss)`t [$($threadID)] `t [$($loglevel)] `t $("{0:N0}" -f ((Get-Date)-$ThreadStartDTM).totalseconds)s `t [$TotalFilesProcessed] `t[$elementsprocessed] `t [$MaxThreads]`t $($logstring)"
                            } Else {
                            $Stoploop = $false
                            [int]$Retrycount = "0"
                            while($Stoploop -eq $false) {
                                try {
                                    Add-content -force $SysLogfile -value "$(get-date -f MM-dd-yy"  "hh:mm:ss)`t [$($threadID)] `t [$($loglevel)] `t $($runtime)s`t$($sitename.padright(24,' '))`t [$TotalFilesProcessed] `t[$elementsprocessed] `t [$MaxThreads]`t [$FilesToServerRatio] `t $($logstring)" -ErrorAction Stop
                                    $Stoploop = $true
                                }
                                catch {
                                    if ($Retrycount -gt 3){
                                        $Stoploop = $true
                                        } else {
                                        Start-Sleep -Milliseconds 30
                                        $Retrycount = $Retrycount + 1
                                    }
                                }
                            }
                        }
                    }
                }
                
                Function MoveFiles {
                    Param ([string]$Path,[string]$Server,[string]$Site)
                    $ProtectedVarPath =$Path
                    $ProtectedVarServer =$Server
                    $ProtectedVarServerShort=($ProtectedVarServer).Replace(".somecompany", "")
                    $ProtectedVarSite= $Site
                    $ArchiveIISLogfile = $DirBase + "\archive\$(get-date -f MM-dd-hh)\$($Site)\$($ProtectedVarServerShort)\"
                    
                    If ($ProtectedVarSite -eq 'akami') {
                        LogWrite "Unpacking akami rar" "0" "INFO"
                        Get-ChildItem $parent -Recurse -Filter "*.rar" | ForEach-Object {
                            $CommandOuput = Invoke-Expression "$($Unrar) x -y $($_.FullName)" 2>&1
                            LogWrite "Command Ouput: $CommandOuput" "0" "INFO"
                        }
                        #build the custom header
                        $originalfilename = $_.FullName
                        $newfilename = $_.FullName + $(get-date).ToString("yyyyMMddthhmmss") + ".log"
$header=@"
#Software: Microsoft Internet Information Services 7.0
#Version: 1.0
#Date: $((Get-Date).ToString("yyyy-MM-dd hh:mm:ss"))
#Fields: date time c-ip cs-method s-sitename cs-uri-stem cs-uri-query sc-status sc-bytes time-taken cs(Referer) cs(User-Agent) cs(Cookie)
"@
                        $header > $newfilename
                        Get-Content $ProtectedVarPath | ForEach-Object {
                            $stemremoval = '/' + ($_.split("`t")[4]).Split('/')[1]
                            $cleanstring = $(($($_.split("`t")[0..1])+$($_.split("`t")[2])+$($_.split("`t")[3])+$(($_.split("`t")[4]).Split('/')[1])+$(($_.split("`t")[4]).Replace($stemremoval,"").Split("?")[0])+$((($_.split("`t")[4]).Replace($stemremoval,"").Split("?")[1]))+$($_.split("`t")[5])+$($_.split("`t")[6])+$($_.split("`t")[7])+$($_.split("`t")[8])+$(($_.split("`t")[9]).replace(" ","+"))+$($_.split("`t")[10]))) -replace '\s+', ' ' -replace "`t",' ' -replace "  `n","`n"  -replace "`""
                            If ($cleanstring -ne $null) {
                                write-output "$cleanstring" >> $newfilename
                            }
                        }
                        LogWrite "Generation of akami logs completed" "0" "INFO"
                        Remove-Item $originalfilename
                        
                    }
                    
                    $ProtectedVarNumberOfFiles=(Get-ChildItem $ProtectedVarPath -filter "*.log").Count
                    If ($ProtectedVarNumberOfFiles -eq $null) {
                        LogWrite "No files found to be processed, moving to next in list" $ProtectedThreadID $ProtectedVarSite "INFO"
                        return
                        }If ($ProtectedVarNumberOfFiles -eq 1) {
                        LogWrite "Only one file present and its locked, so moving to next in list" $ProtectedThreadID $ProtectedVarSite "WARN"
                        return
                    }
                    If ($ProtectedVarNumberOfFiles -gt 300 -and $DeleteBacklogFiles -eq $true) {
                        $ItemsDeleted = ""
                        $ItemsDeleted += "Pathname $ProtectedVarPath `t"
                        $ItemsDeleted += "Servername $ProtectedVarServer `t"
                        $ItemsDeleted += "Site $ProtectedVarSite `t"
                        $numbertodelete = $ProtectedVarNumberOfFiles - 20
                        $ItemsDeleted += "deleting $numbertodelete total files"
                        LogWrite "Starting emergency deletion of logs as $ProtectedVarNumberOfFiles files were found and that is greater than the limit of 300" $ProtectedThreadID $ProtectedVarSite "ERROR"
                        Get-ChildItem $ProtectedVarPath -filter "*.log" | sort LastWriteTime -Descending | select -Last $numbertodelete | Foreach-Object{
                            Remove-Item $_.fullname -force
                            $ItemsDeleted += "$($_.fullname) `t"
                            LogWrite "Emergency deletion of log $($_.fullname) completed" $ProtectedThreadID $ProtectedVarSite "ERROR"
                        }
                        Send-MailMessage -smtpServer 'mail.somecompany' -from 'iislogcollection@somecompany.com' -to 'blarg@somecompany.com' -subject "Emergency deletion of $numbertodelete logs for server $ProtectedVarServer" -body "$ItemsDeleted" -priority High
                    }
                    If ($ProtectedVarNumberOfFiles -gt 300 -and $DeleteBacklogFiles -eq $false) {
                        $threadTimeout = $threadTimeout + 100
                    }
                    LogWrite "$ProtectedVarNumberOfFiles files found within $ProtectedVarPath, begining run on each file" $ProtectedThreadID $ProtectedVarSite "INFO"
                    [int] $TotalFiles = get-content "C:\Temp\TotalFiles"
                    $TotalFiles ++
                    write-output $TotalFiles  > "C:\Temp\TotalFiles"
                    
                    #$RunSummary = get-content "C:\Temp\RunSummary"
                    #$RunSummary += "$($ProtectedVarSite) : $ProtectedVarNumberOfFiles files found within $ProtectedVarPath `t"
                    #write-output $RunSummary >> "C:\Temp\RunSummary"
                    
                    If(!(Test-Path -Path $ArchiveIISLogfile)) {
                        $ProtectedCommandOuput = Invoke-Expression "Mkdir -p $ArchiveIISLogfile" 2>&1
                        LogWrite "Mkdir -p $ArchiveIISLogfile" $ProtectedThreadID $ProtectedVarSite "INFO"
                        } Else {
                        LogWrite "$ArchiveIISLogfile exists" $ProtectedThreadID $ProtectedVarSite "INFO"
                    }
                    Get-ChildItem $ProtectedVarPath -filter "*.log" | sort LastWriteTime -Descending | Foreach-Object{
                        $SmatterStatsArchivefile = $DirBase + "\logsforsmatterstats\$($ProtectedVarSite)\$($ProtectedVarServer)_$(get-date -f MM-dd).zip"
                        If(!(Test-Path -Path "$($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\")) {
                            $ProtectedCommandOuput = Invoke-Expression "Mkdir -p $($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\" 2>&1
                            LogWrite "Mkdir -p $($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\" $ProtectedThreadID $ProtectedVarSite "INFO"
                            } Else {
                            LogWrite "$($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\ exists" $ProtectedThreadID $ProtectedVarSite "INFO"
                        }
                        $SmatterStatsLogfile = $DirBase + "\logsforsmatterstats\$($ProtectedVarSite)\$($ProtectedVarServer)_$(get-date -f MM-dd)_$($_.Name).log"
                        If(!(Test-Path -Path "$($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\")) {
                            $ProtectedCommandOuput = Invoke-Expression "Mkdir -p $($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\" 2>&1
                            LogWrite "Mkdir -p $($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\" $ProtectedThreadID $ProtectedVarSite "INFO"
                            } Else {
                            LogWrite "$($DirBase)\logsforsmatterstats\$($ProtectedVarSite)\ exists" $ProtectedThreadID $ProtectedVarSite "INFO"
                        }
                        LogWrite "Checking $($_.FullName) for file lock" $ProtectedThreadID $ProtectedVarSite "INFO"
                        $logFileName=$_.FullName
                        Try{
                            Move-Item  $_.FullName $_.FullName -force  -ErrorAction Stop
                        }
                        Catch {
                            LogWrite "Log file locked moving to next file : $($_)" $ProtectedThreadID $ProtectedVarSite "WARN"
                            #Try{
                                #    LogWrite "Trying workarround to read locked file via logparser and checkpoints" $ProtectedThreadID $ProtectedVarSite "INFO"
                                #    $ProtectedCommandOuput = &"C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select s-computername as server_name, c-ip as client_ip, s-ip as server_ip , count(*) as rpm,TO_LOCALTIME(QUANTIZE(TO_TIMESTAMP(date, time), 60)) as datestamp into PerServerRPMStats from `'$logFileName`' group by TO_LOCALTIME(QUANTIZE(TO_TIMESTAMP(date,time), 60)), server_name,server_ip,client_ip`" -server:somesqlserver.somecompany -database:IT_Operations -username:somedatabase -password:somepassword -i:IISW3C -ignoreIdCols:ON -o:SQL -transactionRowCount:-1 -iCheckPoint" 2>&1
                                #    LogWrite "$ProtectedVarSite ROLL-OFF DATA logparser query import: $ProtectedCommandOuput" $ProtectedThreadID $ProtectedVarSite "INFO"
                                
                                #    $ProtectedCommandOuput = &"C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select TO_LOCALTIME(TO_TIMESTAMP(date, time)) as date_time, c-ip as client_ip, `'$($ProtectedVarSite)`' as site_name, s-computername as computer_name, s-ip as server_ip, s-port as server_port, cs-method as method, cs-uri-stem as uri_stem, cs-uri-query as uri_query, sc-status as status, sc-substatus as substatus, cs-host as host, cs(User-Agent) as user_agent, cs(Referer) as referer, sc-bytes as bytes_out,cs-bytes as bytes_in, time-taken as time-taken, sc-win32-status as w32status into events_clark from `'$logFileName`' WHERE cs-uri-stem like `'%asp%`' OR cs-uri-stem like `'%htm%`' OR cs-uri-stem like `'%asm%`' OR cs-uri-stem like `'%svc%`' OR cs-uri-stem like `'%axd%' OR cs-uri-stem like `'%/`' OR cs-uri-stem like `'%js%`' or cs-uri-stem like `'%/directory%`' or cs-uri-stem like `'%/name%`' or cs-uri-stem like `'%ashx`'`" -server:somesqlserver.somecompany -database:somedatabase -username:somedatabase -password:somepassword -i:IISW3C -ignoreIdCols:ON -o:SQL -transactionRowCount:-1 -iCheckPoint" 2>&1
                                #    LogWrite "$ProtectedVarSite MAIN LOG DATA logparser query import: $ProtectedCommandOuput" $ProtectedThreadID $ProtectedVarSite "INFO"
                                #    return
                            #}
                            #Catch {
                                #    LogWrite "Logparser error : $($_)" $ProtectedThreadID "ERROR"
                                #    LogWrite "Error trying workarround to read locked file via logparser and checkpoints" $ProtectedThreadID "ERROR"
                                #    return
                            #}
                            return
                        }
                        
                        $CopyStartDTM = (Get-Date)
                        Try{
                            #$ProtectedCommandOuput = &"C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select `'$($ProtectedVarServer)`' as server_name, c-ip as client_ip, s-ip as server_ip , count(*) as rpm,TO_LOCALTIME(QUANTIZE(TO_TIMESTAMP(date, time), 60)) as datestamp into PerServerRPMStats from `'$($_.FullName)`' group by TO_LOCALTIME(QUANTIZE(TO_TIMESTAMP(date,time), 60)),server_ip,client_ip`" -server:somesqlserver.somecompany -database:IT_Operations -username:somedatabase -password:somedatabase -i:IISW3C -ignoreIdCols:ON -o:SQL -transactionRowCount:-1" 2>&1
                            #LogWrite "ROLL-OFF DATA : logparser query import: $ProtectedCommandOuput" $ProtectedThreadID $ProtectedVarSite "INFO"
                            If ($ProtectedVarSite -eq 'akami') {
                                $ProtectedCommandOuput = &"C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select TO_LOCALTIME(TO_TIMESTAMP(date, time)) as date_time, c-ip as client_ip, cs-method as method, cs-uri-stem as uri_stem, cs-uri-query as uri_query, sc-status as status, cs(User-Agent) as user_agent, cs(Referer) as referer, sc-bytes as bytes, time-taken as time-taken into akami_events from `'\\somefileserver\test.20140627P030441.log`'`" -server:somesqlserver.somecompany -database:somedb -username:someusername -password:somepassword -i:IISW3C -ignoreIdCols:ON -o:SQL-e 50"
                                } Else {
                                $ProtectedCommandOuput = &"C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select TO_LOCALTIME(TO_TIMESTAMP(date, time)) as date_time, c-ip as client_ip, `'$($ProtectedVarSite)`' as site_name, `'$($ProtectedVarServer)`' as computer_name, s-ip as server_ip, s-port as server_port, cs-method as method, cs-uri-stem as uri_stem, cs-uri-query as uri_query, sc-status as status, sc-substatus as substatus, cs-host as host, cs(User-Agent) as user_agent, cs(Referer) as referer, sc-bytes as bytes_out,cs-bytes as bytes_in, time-taken as time-taken, sc-win32-status as w32status into events from `'$($_.FullName)`' WHERE cs-uri-stem like `'%asp%`' OR cs-uri-stem like `'%htm%`' OR cs-uri-stem like `'%asm%`' OR cs-uri-stem like `'%svc%`' OR cs-uri-stem like `'%axd%' OR cs-uri-stem like `'%/`' OR cs-uri-stem like `'%js%`' or cs-uri-stem like `'%/directory%`' or cs-uri-stem like `'%/name%`' or cs-uri-stem like `'%ashx`'`" -server:somesqlserver.somecompany -database:somedatabase -username:someusername -password:somepassword -i:IISW3C -ignoreIdCols:ON -o:SQL" 2>&1

                                #write-host "C:\Users\admin.clark\Desktop\mechanism\logparser.exe" "`"select TO_LOCALTIME(TO_TIMESTAMP(date, time)) as date_time, c-ip as client_ip, `'$($ProtectedVarSite)`' as site_name, `'$($ProtectedVarServer)`' as computer_name, s-ip as server_ip, s-port as server_port, cs-method as method, cs-uri-stem as uri_stem, cs-uri-query as uri_query, sc-status as status, sc-substatus as substatus, cs-host as host, cs(User-Agent) as user_agent, cs(Referer) as referer, sc-bytes as bytes_out,cs-bytes as bytes_in, time-taken as time-taken, sc-win32-status as w32status into events from `'$($_.FullName)`' WHERE cs-uri-stem like `'%asp%`' OR cs-uri-stem like `'%htm%`' OR cs-uri-stem like `'%asm%`' OR cs-uri-stem like `'%svc%`' OR cs-uri-stem like `'%axd%' OR cs-uri-stem like `'%/`' OR cs-uri-stem like `'%js%`' or cs-uri-stem like `'%/directory%`' or cs-uri-stem like `'%/name%`' or cs-uri-stem like `'%ashx`'`" -server:somesqlserver.somecompany -database:somedatabase -username:someusername -password:somepassword -i:IISW3C -ignoreIdCols:ON -o:SQL"
                            }
                            LogWrite "logparser output :$ProtectedCommandOuput" $ProtectedThreadID $ProtectedVarSite "INFO"
                        }
                        Catch {
                            LogWrite "Logparser error : $($_)" $ProtectedThreadID $ProtectedVarSite "ERROR"
                            return
                        }
                        [string]$outputstring = $ProtectedCommandOuput
                        $cleanedoutput = $outputstring.replace(" Statistics: ","").replace("-","").replace(" Elements processed: ","").replace(" Execution time:","").replace(" seconds ","").replace(" Elements output: ","").replace("  ","") -replace "\(([^\)]+)\)"," "
                        
                        [int] $elementsprocessed = get-content "C:\Temp\elementsprocessed"-ErrorAction SilentlyContinue
                        #$elementsoutput = $cleanedoutput.split(" ")[1]
                        $elementsprocessed += [int] $cleanedoutput.split(" ")[1]
                        write-output $elementsprocessed > "C:\Temp\elementsprocessed" -ErrorAction SilentlyContinue
                        
                        $insertrate = $cleanedoutput.split(" ")[2]
                        [int] $MaxThreads = (Get-Process powershell).count
                        if ([int]$insertrate -gt 35.00) {
                            if ($MaxThreads -le 10) {
                                $MaxThreads = 10
                                LogWrite "Insert speed of $insertrate would warrent reducing threads but min thread count reached of $MaxThreads" $ProtectedThreadID $ProtectedVarSite "INFO"
                                }else {
                                $MaxThreads --
                                LogWrite "Thread count reduced due to insert speed of $insertrate, thread count now at $MaxThreads" $ProtectedThreadID $ProtectedVarSite "INFO"
                            }
                            } elseif ([int]$insertrate -lt 30.00) {
                            if ($MaxThreads -ge 70) {
                                $MaxThreads = 70
                                LogWrite "Insert speed of $insertrate would warrent increasing threads but max thread count reached of $MaxThreads" $ProtectedThreadID $ProtectedVarSite "INFO"
                                }else {
                                $MaxThreads ++
                                LogWrite "Thread count increased due to insert speed of $insertrate, thread count now at $MaxThreads" $ProtectedThreadID $ProtectedVarSite "INFO"
                            }
                        }
                        write-output $MaxThreads > "C:\Temp\MaxThreads"  -ErrorAction SilentlyContinue
                        Try{
                            #if ($ProtectedVarSite -eq 'pf2') {
                                #    $CommandOuput = Invoke-Expression "$($WinZip) a $SmatterStatsArchivefile$($_.FullName)" 2>&1
                                #    LogWrite "Command Ouput: $CommandOuput" $ProtectedThreadID $ProtectedVarSite "INFO"
                                #}Else {
                                Copy-Item  $_.FullName $SmatterStatsLogfile -force  -ErrorAction Stop
                            #}
                            Move-Item $logFileName $ArchiveIISLogfile -force  -ErrorAction Stop
                        }
                        Catch {
                            LogWrite "Error moving $logFileName to $ArchiveIISLogfile : $($_)" $ProtectedThreadID $ProtectedVarSite "WARN"
                            Try{
                                LogWrite "Archive failed, trying to remove file from the server" $ProtectedThreadID $ProtectedVarSite "WARN"
                                Remove-Item -force $logFileName -ErrorAction Stop
                                If(!(Test-Path -Path $logFileName)) {
                                    LogWrite "Deletion of $logFileName success" $ProtectedThreadID $ProtectedVarSite "INFO"
                                    }Else {
                                    LogWrite "Deletion of $logFileName failed, expect duplicate inserts" $ProtectedThreadID $ProtectedVarSite "ERROR"
                                }
                            }
                            Catch {
                                If(!(Test-Path -Path $logFileName)) {
                                    LogWrite "Deletion of $logFileName success" $ProtectedThreadID $ProtectedVarSite "INFO"
                                    }Else {
                                    LogWrite "Deletion of $logFileName failed, expect duplicate inserts" $ProtectedThreadID $ProtectedVarSite "ERROR"
                                }
                            }
                            return
                        }
                        $CopyEndDTM = (Get-Date)
                        LogWrite "process completed against $($_.FullName) in $(($CopyEndDTM-$CopyStartDTM).totalseconds) seconds" $ProtectedThreadID $ProtectedVarSite "INFO"
                        $FilesProcessed ++
                        [int] $TotalFilesProcessed = get-content "C:\Temp\TotalFilesProcessed"-ErrorAction SilentlyContinue
                        $TotalFilesProcessed ++
                        write-output $TotalFilesProcessed  > "C:\Temp\TotalFilesProcessed"-ErrorAction SilentlyContinue
                        If (((Get-Date)-$ThreadStartDTM).totalseconds -gt $threadTimeout) {
                            LogWrite "I... need you... to pick up that paperclip. Then go away forever. Thread $ProtectedThreadID death eminent" $ProtectedThreadID $ProtectedVarSite "WARN"
                            break
                        }
                    }
                    If ($FilesProcessed -ne $ProtectedVarNumberOfFiles){
                        LogWrite "Number of files processed does not equal number of files found : $($FilesProcessed) of $ProtectedVarNumberOfFiles" $ProtectedThreadID $ProtectedVarSite "WARN"
                    }
                    LogWrite "Import of $($ProtectedVarPath) completed with $($FilesProcessed) files processed" $ProtectedThreadID $ProtectedVarSite "NOTE"
                    #$RunSummary = get-content "C:\Temp\RunSummary"
                    #$RunSummary += "$($ProtectedVarSite) : $($FilesProcessed) files processed within $ProtectedVarPath `t"
                    #write-output $RunSummary >> "C:\Temp\RunSummary"
                    $FilesProcessed = 0
                }

                # ¦¦¦¦¦¦+ ¦¦¦¦¦¦+  ¦¦¦¦¦¦+  ¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+    ¦¦¦¦¦¦¦+¦¦+   ¦¦+¦¦¦+   ¦¦+ ¦¦¦¦¦¦+¦¦¦¦¦¦¦¦+¦¦+ ¦¦¦¦¦¦+ ¦¦¦+   ¦¦+¦¦¦¦¦¦¦+
                # ¦¦+--¦¦+¦¦+--¦¦+¦¦+---¦¦+¦¦+----+¦¦+----+¦¦+----+¦¦+----+    ¦¦+----+¦¦¦   ¦¦¦¦¦¦¦+  ¦¦¦¦¦+----++--¦¦+--+¦¦¦¦¦+---¦¦+¦¦¦¦+  ¦¦¦¦¦+----+
                # ¦¦¦¦¦¦++¦¦¦¦¦¦++¦¦¦   ¦¦¦¦¦¦     ¦¦¦¦¦+  ¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+    ¦¦¦¦¦+  ¦¦¦   ¦¦¦¦¦+¦¦+ ¦¦¦¦¦¦        ¦¦¦   ¦¦¦¦¦¦   ¦¦¦¦¦+¦¦+ ¦¦¦¦¦¦¦¦¦¦+
                # ¦¦+---+ ¦¦+--¦¦+¦¦¦   ¦¦¦¦¦¦     ¦¦+--+  +----¦¦¦+----¦¦¦    ¦¦+--+  ¦¦¦   ¦¦¦¦¦¦+¦¦+¦¦¦¦¦¦        ¦¦¦   ¦¦¦¦¦¦   ¦¦¦¦¦¦+¦¦+¦¦¦+----¦¦¦
                # ¦¦¦     ¦¦¦  ¦¦¦+¦¦¦¦¦¦+++¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦    ¦¦¦     +¦¦¦¦¦¦++¦¦¦ +¦¦¦¦¦+¦¦¦¦¦¦+   ¦¦¦   ¦¦¦+¦¦¦¦¦¦++¦¦¦ +¦¦¦¦¦¦¦¦¦¦¦¦¦
                # +-+     +-+  +-+ +-----+  +-----++------++------++------+    +-+      +-----+ +-+  +---+ +-----+   +-+   +-+ +-----+ +-+  +---++------+
                # *Process functions
                # these are the meat of the program enabling the transfer of data into sql server and archives and more
                #
                LogWrite "New thread started" $ProtectedThreadID $($Object.Sitename) "INIT"
                [int] $TotalServers = get-content "C:\Temp\TotalServers" -ErrorAction SilentlyContinue
                $TotalServers ++
                write-output $TotalServers  > "C:\Temp\TotalServers" -ErrorAction SilentlyContinue
                
                LogWrite "Looking at $($Object.Pathname) and server $($Object.Servername) and sitename $($Object.Sitename)" $ProtectedThreadID $($Object.Sitename) "INFO"
                If(!(Test-Connection -Cn $Object.Servername -BufferSize 16 -Count 1 -ea 0  -ThrottleLimit 1000 )) {
                    LogWrite "Server $($Object.Servername) does not exist on server so skipping" $ProtectedThreadID $($Object.Sitename) "WARN"
                    } Else {
                    If(!(Test-Path -Path $Object.Pathname)) {
                        LogWrite "Path $($Object.Pathname) does not exist on server so skipping" $ProtectedThreadID $($Object.Sitename) "WARN"
                        } Else {
                        MoveFiles $Object.Pathname $Object.Servername $Object.Sitename
                    }
                }
            }# end process block
            
            # ¦¦¦¦¦¦+ ¦¦¦¦¦¦+  ¦¦¦¦¦¦+  ¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+    ¦¦+    ¦¦+¦¦¦¦¦¦+  ¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦¦+¦¦¦¦¦¦+ 
            # ¦¦+--¦¦+¦¦+--¦¦+¦¦+---¦¦+¦¦+----+¦¦+----+¦¦+----+¦¦+----+    ¦¦¦    ¦¦¦¦¦+--¦¦+¦¦+--¦¦+¦¦+--¦¦+¦¦+--¦¦+¦¦+----+¦¦+--¦¦+
            # ¦¦¦¦¦¦++¦¦¦¦¦¦++¦¦¦   ¦¦¦¦¦¦     ¦¦¦¦¦+  ¦¦¦¦¦¦¦+¦¦¦¦¦¦¦+    ¦¦¦ ¦+ ¦¦¦¦¦¦¦¦¦++¦¦¦¦¦¦¦¦¦¦¦¦¦¦++¦¦¦¦¦¦++¦¦¦¦¦+  ¦¦¦¦¦¦++
            # ¦¦+---+ ¦¦+--¦¦+¦¦¦   ¦¦¦¦¦¦     ¦¦+--+  +----¦¦¦+----¦¦¦    ¦¦¦¦¦¦+¦¦¦¦¦+--¦¦+¦¦+--¦¦¦¦¦+---+ ¦¦+---+ ¦¦+--+  ¦¦+--¦¦+
            # ¦¦¦     ¦¦¦  ¦¦¦+¦¦¦¦¦¦+++¦¦¦¦¦¦+¦¦¦¦¦¦¦+¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦    +¦¦¦+¦¦¦++¦¦¦  ¦¦¦¦¦¦  ¦¦¦¦¦¦     ¦¦¦     ¦¦¦¦¦¦¦+¦¦¦  ¦¦¦
            # +-+     +-+  +-+ +-----+  +-----++------++------++------+     +--++--+ +-+  +-++-+  +-++-+     +-+     +------++-+  +-+
                                                                                                                       
            # *process wrapper
            # this maintains the controll and health of the threads that move the logs. Its job is to call process code blocks into job statements to the max thread count.
            # it also does some in flight cleanup and as well some after flight procedure.
            #            
            LogWrite '' "0" "INIT"
            LogWrite 'ooooo                                    .oooooo.             oooo  oooo                          .    o8o                                           .o    .oooo.   ' "0" "INIT"
            LogWrite '`888"                                   d8P"  `Y8b            `888  `888                        .o8    `""                                         o888   d8P"`Y8b  ' "0" "INIT"
            LogWrite ' 888          .ooooo.   .oooooooo      888           .ooooo.   888   888   .ooooo.   .ooooo.  .o888oo oooo   .ooooo.  ooo. .oo.        oooo    ooo  888  888    888 ' "0" "INIT"
            LogWrite ' 888         d88" `88b 888" `88b       888          d88" `88b  888   888  d88" `88b d88" `"Y8   888   `888  d88" `88b `888P"Y88b        `88.  .8"   888  888    888 ' "0" "INIT"
            LogWrite ' 888         888   888 888   888       888          888   888  888   888  888ooo888 888         888    888  888   888  888   888         `88..8"    888  888    888 ' "0" "INIT"
            LogWrite ' 888       o 888   888 `88bod8P"       `88b    ooo  888   888  888   888  888    .o 888   .o8   888 .  888  888   888  888   888          `888"     888  `88b  d88" ' "0" "INIT"
            LogWrite 'o888ooooood8 `Y8bod8P" `8oooooo.        `Y8bood8P"  `Y8bod8P" o888o o888o `Y8bod8P" `Y8bod8P"   "888" o888o `Y8bod8P" o888o o888o          `8"     o888o  `Y8bd8P"  ' "0" "INIT"
            LogWrite '                       d"     YD                                                                                                                                    ' "0" "INIT"
            LogWrite '                       "Y88888P"                                                                                                                                    ' "0" "INIT"
            LogWrite '' "0" "INIT"
            LogWrite "Script init from locaton \\somefileserver\windowsiislogs\mechanism\" "0" "INIT"
            LogWrite "Documentation located at \\somefileserver\windowsiislogs\mechanism\" "0" "INIT"
            LogWrite "Cleaning up from last run, this can take up to $threadTimeout seconds to complete" "0" "INIT"
            $currentwait = 0
            While ($(Get-Job -state running | Where-Object {$_.Name.Contains("logcollection")} ).count -ne $null){
                Get-Job -State Completed | Where-Object {$_.Name.Contains("logcollection")}  | Stop-job
                Get-Job -State Completed  | Where-Object {$_.Name.Contains("logcollection")} | Remove-Job
                Get-Job -State Stopped | Where-Object {$_.Name.Contains("logcollection")} | Remove-Job
                Start-Sleep 1
                $currentwait ++
                if ($currentwait -gt $threadTimeout) {
                    LogWrite "hit max wait time, ending all current threads now" "0" "INIT"
                    Get-Job  | Where-Object {$_.Name.Contains("logcollection")} | remove-job -f
                }
            }
            Get-Job  | Where-Object {$_.Name.Contains("logcollection")} | remove-job -f
            LogWrite "All stale  threads cleaned" "0" "INIT"
            $TotalServers=0
            $TotalFiles=0
            write-output $TotalServers > "C:\Temp\TotalServers"  -ErrorAction SilentlyContinue
            write-output $TotalFiles > "C:\Temp\TotalFiles"  -ErrorAction SilentlyContinue
            #write-output $RunSummary > "C:\Temp\RunSummary"
            write-output $MaxThreads > "C:\Temp\MaxThreads"  -ErrorAction SilentlyContinue
            $ScriptStartDTM = (Get-Date)
            LogWrite "Listing all settings" "0" "INIT"
            get-variable -include ser*,m*,arch*,*iis*,fie*,processed* | Foreach-Object {
                LogWrite "       $($_.name)=$($_.value)" "0" "INIT"
            }
            Import-CSV $ServersFile -Header Sitename,Servername,Pathname,Flag | SortRandom | SortRandom | Foreach-Object {
                if ($_.flag -eq "keep"){
                    While ((Get-Process powershell).count -ge (get-content "C:\Temp\MaxThreads")){
                        Get-Job | Where-Object {$_.Name.Contains("logcollection")} | Receive-Job
                        Get-Job -State Completed | Where-Object {$_.Name.Contains("logcollection")}  | Stop-job
                        Get-Job -State Completed  | Where-Object {$_.Name.Contains("logcollection")} | Remove-Job
                        Get-Job -State Stopped | Where-Object {$_.Name.Contains("logcollection")} | Remove-Job
                        Start-Sleep -Milliseconds $SleepTimer
                        $lastwritetime = [datetime](Get-ItemProperty -Path "\\somefileserver.somecompany\windowsiislogs\syslog\syslog.log" -Name LastWriteTime).lastwritetime
                        if ($lastwritetime -lt ((Get-Date).AddMinutes(-15))){
                            Get-Job  | Where-Object {$_.Name.Contains("logcollection")} | remove-job -f
                            Send-MailMessage -smtpServer 'mail.somecompany' -from 'iislogcollection@somecompany.com' -to 'blarg@somecompany.com' -subject "Emergency shutdown of log collection script" -body "Emergency shutdown of log collection script" -priority High
                            Stop-Process -processname WerFault -Force -ErrorAction SilentlyContinue
                            Stop-Process -processname Powershell -Force -ErrorAction SilentlyContinue
                            Stop-Process -processname LogParser -Force -ErrorAction SilentlyContinue
                            Stop-Process -processname WerFault -Force -ErrorAction SilentlyContinue
                            return
                        }
                    }
                    $Object = $_
                    Start-Job -ScriptBlock $process -ArgumentList $Object -Name logcollection | out-null
                } else {
                    $NumberOfFiles=(Get-ChildItem $_.Pathname -filter "*.log").Count
                    write-host "No Keep flag for $($_.Pathname) and server $($_.Servername) and sitename $($_.Sitename)" "CORE" $_.Sitename "WARN"
                    $numbertodelete = $NumberOfFiles - 3
                    write-host "Starting deletion of logs given missing flag value though will leave 3 most recent logs in place" "CORE" $_.Sitename "WARN"
                    if ($numbertodelete -gt 0) {
                        Get-ChildItem $_.Pathname -filter "*.log" | sort LastWriteTime -Descending | select -Last $numbertodelete | Foreach-Object{
                            Remove-Item $_.fullname -force
                            write-host "deletion of log $($_.fullname) completed" "CORE" $($_.Sitename) "WARN"
                        }
                    }
                }                    
            }
            If ($TotalFilesProcessed -eq "" -or $TotalFilesProcessed -eq $null ){ $TotalFilesProcessed = 0 }
            LogWrite "Import of $($ServersFile) completed with $($TotalFilesProcessed) files processed" "0" "INFO"
            LogWrite "Cleaning $($DirBase)\archive of files older than $($FileAgeLimit)" "0" "INFO"
            Get-ChildItem -Path $ArchiveLocation | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $FileAgeLimit } | Foreach-Object {
                LogWrite "Removing $($_.FullName) from $ArchiveLocation" "0" "INFO"
                Remove-Item $_.FullName
            }
            LogWrite "Cleaning C:\Users\admin.clark\Documents of partial archive files" "0" "INFO"
            Get-ChildItem -Path "C:\Users\admin.clark\Documents" -filter "*rar*", "*zip*" | Foreach-Object {
                LogWrite "Removing $($_.FullName) from C:\Users\admin.clark\Documents" "0" "INFO"
                Remove-Item $_.FullName
            }
            $ScriptEndDTM = (Get-Date)
            #$RunSummary = get-content "C:\Temp\RunSummary"
            #$RunSummary | foreach {
                #     LogWrite $_  "0" "SUMMARY"
            #}
            Stop-Process -processname WerFault -Force -ErrorAction SilentlyContinue
            [int] $elementsprocessed = get-content "C:\Temp\elementsprocessed"-ErrorAction SilentlyContinue
            $TotalFilesProcessed = get-content "C:\Temp\TotalFilesProcessed"-ErrorAction SilentlyContinue
            LogWrite "total files processed $TotalFilesProcessed" "0" "SUMMARY"
            LogWrite "total elements processed $elementsprocessed" "0" "SUMMARY"
            LogWrite "script start $ScriptStartDTM"  "0" "SUMMARY"
            LogWrite "script end $ScriptEndDTM"  "0" "SUMMARY"
            $emailbody = "Script completed in $(($ScriptEndDTM-$ScriptStartDTM).totalseconds) seconds `n"
            $emailbody += "total files processed $TotalFilesProcessed `n"
            $emailbody += "total elements processed $elementsprocessed `n"
            $emailbody += "script start $ScriptStartDTM `n"
            $emailbody += "script end $ScriptEndDTM `n"
            $emailbody += ""
            Send-MailMessage -smtpServer 'mail.somecompany' -from 'iislogcollection@somecompany.com' -to 'blarg@somecompany.com' -subject "Script completed in $(($ScriptEndDTM-$ScriptStartDTM).totalseconds) seconds" -body "$emailbody" -priority Low
            LogWrite "Script completed in $(($ScriptEndDTM-$ScriptStartDTM).totalseconds) seconds"  "0" "SUMMARY"
        }# end corescript trueloop  
    }# end corescript codeblock

    #  ¦¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦¦+    ¦¦+    ¦¦+¦¦¦¦¦¦+  ¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦+ ¦¦¦¦¦¦¦+¦¦¦¦¦¦+ 
    # ¦¦+----+¦¦+---¦¦+¦¦+--¦¦+¦¦+----+    ¦¦¦    ¦¦¦¦¦+--¦¦+¦¦+--¦¦+¦¦+--¦¦+¦¦+--¦¦+¦¦+----+¦¦+--¦¦+
    # ¦¦¦     ¦¦¦   ¦¦¦¦¦¦¦¦¦++¦¦¦¦¦+      ¦¦¦ ¦+ ¦¦¦¦¦¦¦¦¦++¦¦¦¦¦¦¦¦¦¦¦¦¦¦++¦¦¦¦¦¦++¦¦¦¦¦+  ¦¦¦¦¦¦++
    # ¦¦¦     ¦¦¦   ¦¦¦¦¦+--¦¦+¦¦+--+      ¦¦¦¦¦¦+¦¦¦¦¦+--¦¦+¦¦+--¦¦¦¦¦+---+ ¦¦+---+ ¦¦+--+  ¦¦+--¦¦+
    # +¦¦¦¦¦¦++¦¦¦¦¦¦++¦¦¦  ¦¦¦¦¦¦¦¦¦¦+    +¦¦¦+¦¦¦++¦¦¦  ¦¦¦¦¦¦  ¦¦¦¦¦¦     ¦¦¦     ¦¦¦¦¦¦¦+¦¦¦  ¦¦¦
    #  +-----+ +-----+ +-+  +-++------+     +--++--+ +-+  +-++-+  +-++-+     +-+     +------++-+  +-+
    # *core wrapper
    # she maintains the overall health of the script and ensures that it stays operational. Its job is to simply call the core code block as a process when it is not already running.
    #
    If (@(Get-Job | Where { $_.State -eq "Running" } | Where-Object {$_.Name.Contains("corescript")}).Count -lt 1 -or (@(Get-Job | Where { $_.State -eq "Running" } | Where-Object {$_.Name.Contains("corescript")}).Count -eq $Null)){
        Start-Job -ScriptBlock $core -Name corescript | out-null
        Send-MailMessage -smtpServer 'mail.somecompany' -from 'iislogcollection@somecompany.com' -to 'blarg@somecompany.com' -subject "Log Collection successfully started" -body "Log Collection successfully started" -priority low
        sleep 5
    }
    Get-Job | Where-Object {$_.Name.Contains("corescript")} | Receive-Job
    Get-Job | Where { $_.State -ne "Running" } | Where-Object {$_.Name.Contains("corescript")} | remove-job -force
    Get-Job -State Completed | Where-Object {$_.Name.Contains("corescript")}  | Stop-job
    Get-Job -State Completed  | Where-Object {$_.Name.Contains("corescript")} | Remove-Job      
    sleep 1
}# end main true loop