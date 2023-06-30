# =====================================
# Récupération de la liste des serveurs
# =====================================

Import-Module ActiveDirectory

$Servers = @()
$Servers += (Get-ADComputer).DNSHostName #Example of how you cna get a list of server names

# =====================================
# Récupération de l'uptime des serveurs
# =====================================

$Maximum_allowed_reboots = 1
$Maximum_allowed_Uptime = 60
$Maximum_allowed_Uptime_FILTER1 = 7
$Maximum_allowed_Uptime_FILTER2 = 1

$sheet = @()

foreach ($Server in $Servers) {
    Write-Host -Message "######################## $Server ########################"

    Write-Debug -Message "$Server`: Attempting connection..."
    try {
        $Connection = Test-NetConnection -ComputerName "$Server" -Port 3389 -WarningAction Stop -ErrorAction Stop
        $Session = New-PSSession -SSHTransport -HostName "$Server" -Port 22 -KeyFilePath C:\Users\admin\.ssh\id_rsa -WarningAction Stop -ErrorAction Stop #Example path
        
        Write-Host -ForegroundColor Green -Message "$Server`: Connection established."

        Clear-Variable -name "Name", "Days", "SNMP" -ErrorAction SilentlyContinue

        Write-Debug -Message "$Server`: Attempting diagnosis..."
        try {
            $sheet += Invoke-Command $Session -ArgumentList $Server -ScriptBlock {
                Param($HostName)
                $Diagnostic = New-Object -TypeName "PSObject"
                
                #Keeping only needed informations
                $OS = (Get-ComputerInfo).OsName
                $OS = $OS -split " "
                $OS = $OS[1..3] #Keeps only the relevant parts
                $OS = $OS -join " "

                $Uptime = (Get-Uptime).Days

                if (Get-CimInstance -Class Win32_Service -Filter "Name='snmp'") {
                    $SNMP = (Get-CimInstance -Class Win32_Service -Filter "Name='nscp'").State
                    
                } else {
                    $SNMP = "Not_installed"

                }

                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Name" -Value $HostName
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Uptime" -Value $Uptime
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "OS" -Value $OS
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "SNMP" -Value $SNMP

                return $Diagnostic

            }
            Write-Host -ForegroundColor Green -Message "$Server`: Diagnosis completed."

        } catch {
            Write-Error -Message "$Server`: Failed to diagnose! Error message: $_"

        }     

    } catch { 
        Write-Error -Message "$Server`: Failed to connect! Error message: $_"

        Write-Debug -Message "$Server`: Attempting further diagnosis over failure..."
        try {
            $Diagnostic = New-Object PSObject
            Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Name" -Value $Server

            #Evaluating error level
            if (-not $Connection.PingSucceeded){
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "Ping KO"
                Write-Debug -Message "$Server`: Ping is KO."

            } elseif (-not $Connection.TcpTestSucceeded){
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "RDP KO"
                Write-Debug -Message "$Server`: RDP is KO."

            } elseif (-not $Session){ #Finnaly, not sure it works. Use "& ssh" test instead
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "SSH KO"
                Write-Debug -Message "$Server`: SSH is KO."

            } else {
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "Unknown error"
                Write-Debug -Message "$Server`: Unknown error."

            }

            $sheet += $Diagnostic

            Write-Host -ForegroundColor Green -Message "$Server`: Diagnosis over failure completed."

        } catch {
            Write-Error -Message "$Server`: Failed to diagnose further! Error message: $_"

        }
        
    }
    

}

Write-Host -ForegroundColor Magenta -Message "######################## All servers have been diagnosed. ########################"  

# ========================
# Redémarrage des serveurs 
# ========================

$sheet = $sheet | Sort-Object -Descending Issue, Uptime | Select-Object Name, OS, Uptime, Issue, SNMP 

foreach ($Diagnostic in  $sheet | Where-Object {$_.Issue.Count -eq 0}) {
    $HostName = $Diagnostic.Name

    Write-Host -Message "######################## $HostName ########################"
    
    Write-Debug -Message "$HostName`: Attempting to manage..."
    try {
        if ($HostName -match "filter1") {
            if ((Get-Date).DayOfWeek -eq "Saturday"){ #EXample of specific condition on filtered servers
                $Max = $Maximum_allowed_Uptime_FILTER1
                $Maximum_allowed_reboots++ #Cancel filter1 effect on daily reboots limit

            } else {
                Write-Host -ForegroundColor Yellow -Message "$hostname`: Must be rebooted on Saturdays."
                continue

            }

        } elseif ($HostName -match "filter2") {
            $Max = $Maximum_allowed_Uptime_FILTER2
            #You might don't want to cancel the effect on reboots limit but just have a different uptime limit

        } else {
            $Max = $Maximum_allowed_Uptime

        }

        Write-Debug -Message "Checking if reboot is required..."
        if ($Diagnostic.Uptime -eq ($Max - 1)) {
            Write-Host -ForegroundColor Yellow -Message "$hostname : Reboot required! Scheduled tomorrow."
            Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Reboot" -Value "Tomorrow"
        
        } elseif ($Diagnostic.Uptime -ge $Max) {
            if ($Maximum_allowed_reboots -eq 0) {
                Write-Host -ForegroundColor Yellow -Message "$hostname : Reboot required! Queued"
                Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Reboot" -Value "Queued"
            
            #Reboot !  
            } else {        
                Write-Host -ForegroundColor Red -Message "$hostname`: Reboot required! Scheduled now!"
                
                Write-Debug -Message "$hostname`: Attempting reboot!"
                try {
                    $Session = New-PSSession -SSHTransport -HostName "$hostname" -Port 22 -KeyFilePath C:\Users\admin\.ssh\id_rsa -WarningAction Stop -ErrorAction Stop #Example path
                    Invoke-Command $Session -ErrorAction Stop -ArgumentList $Server -ScriptBlock { 
                        Restart-Computer -Force 

                    } 

                    Write-Debug -Message "Pending..."
                    $failure = 0
                    $success = 0
                    while (($success -lt 5) -and ($failure -lt 30)) {
                        Start-Sleep -Seconds 5
                        $answer = & ssh $HostName echo "OK"
                        if ($answer -eq "OK") {
                            $success++

                        } else {
                            $success = 0
                            $failure++

                        }

                        Write-Debug -Message "Success: $success/5 | Failure: $failure/30"

                    } 

                    if ($success -eq 5) {
                        Write-Debug -Message "Checking success..."
                        $Session = New-PSSession -SSHTransport -HostName "$hostname" -Port 22 -KeyFilePath C:\Users\admin\.ssh\id_rsa -WarningAction Stop -ErrorAction Stop #Example path
                        $Uptime = Invoke-Command $Session  -WarningAction Stop -ErrorAction Stop -ScriptBlock {
                            $Uptime = (Get-Uptime).Days
                            return $Uptime
                            
                        }

                        Write-Debug -Message "Uptime: $Uptime Days"
                        if ($Uptime -eq 0) {
                            Write-Host -ForegroundColor Green -Message "$hostname`: Reboot succeed!"
                            $Diagnostic.Uptime = $Uptime
                            Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Reboot" -Value "Reboot succeed"

                        } else {
                            Write-Error -Message "$hostname`: Reboot incomplete!"
                            Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "Reboot incomplete"

                        }

                    } else {
                        Write-Error -Message "$hostname`: DEAD!"
                        Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "DEAD"
                        
                    }
                
                } catch {
                    Write-Error -Message "$hostname`: Reboot failed! Error message: $_"
                    Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "Reboot failed"

                }
                
                $Maximum_allowed_reboots--

            }

        } else {
            Write-Host -ForegroundColor Green -Message "$hostname`: Uptime under maximum, not further operations required."

        }

    } catch {
        Write-Error -Message "$hostname`: Management failed."
        Add-Member -InputObject $Diagnostic -MemberType NoteProperty -Name "Issue" -Value "Management failed"

    }

}

$sheet = $sheet | Sort-Object -Descending Issue, Uptime | Select-Object Name, OS, Uptime, Issue, SNMP, Reboot

# ======================
# Préparation de l'email
# ======================

$smtpServer = "ipadress"
$from = "youradress"
$to = "teamadress"
$subject = "Uptime Servers"
$encoding = "UTF8"

# ====================
# Envoi de l'email
# ====================

$Header =
@"
<style>
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
    TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #ebbd63;}
    TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@    

$sheetHTML = $sheet | ConvertTo-Html -Head $Header| out-string

$body = 
"
<html>
    <body>
        <p>
            Bonjour,
            <br />
            Voici le diagnostic des serveurs Windows :
            <br />
            <br />
            $sheetHTML
        </p>
    </body>
</html>
"

$Reasontosend = $sheet | Where-Object {($_.Issue.Count -ge 1) -or ($_.Reboot.Count -ge 1) -or ($_.SNMP.Count -ge 0)}
if ($Reasontosend.Count -ne 0) {
    Send-MailMessage -bodyasHTML -smtpserver $smtpserver -from $from -to $to -subject $subject -body $body -Encoding $encoding
    Write-Host -ForegroundColor Green "=========== Mail sent to $to. ==========="

} else {
    Write-Host -ForegroundColor Green "=========== No problem nor restart planned. Mail won't be sent. ==========="

}
