# ==============
# Configurations
# ==============
$MAXreboot = 1

class Server {
    #Properties (~ Attributes but it refers to something else in PowerShell)
    [System.Management.Automation.PSObject[]]$Ping
    [System.Management.Automation.Runspaces.PSSession]$Session
    
    [hashtable]$Debug = @{}
    
    [string]$MasterRSA = "C:\Users\admin\.ssh\id_rsa" #Example path
    [string]$Hostname
    [string]$OSname
    [string]$State
    [string]$SNMP
    
    [bool]$Issue = $false
    [bool]$Restricted
    
    [int]$Uptime
    [int]$MAX

    #Constructors
    Server([string]$Hostname) {
        $this.Hostname = $Hostname
        $this.Diagnose()
        Write-Debug -Message "$($this.Hostname)`> Registration completed."

    }

    #Methods
    [int]Connect() {
        Write-Debug -Message "$($this.Hostname)`> Attempting connection..."
        try { 
            $this.Ping = Test-Connection -ComputerName $this.Hostname -ErrorAction Stop
            if (($this.Ping).Status -ne "Success") {
                $this.Debug["230"] = "Connection failed! Ping KO."
                $this.State = "Ping KO!"
                $this.Issue = $true
                Write-Error -Message "$($this.Hostname)`> Ping KO!"

                return 230

            }
            
            $this.Session = New-PSSession -SSHTransport -HostName $this.Hostname -Port 22 -KeyFilePath $this.MasterRSA -ErrorAction Stop
            
            Write-Debug -Message "$($this.Hostname)`> Connection succeeded."
            $this.State = "Connected"

            return 0

        } catch {
            $this.Debug["404"] = $Error[0].Exception.Message
            $this.State = "Unreachable"
            $this.Issue = $true
            Write-Error -Message "$($this.Hostname)`> Connection failed!" #Must be put at the end or hastable will take this error message

            return 404

        }
    } 

    [int]Diagnose() {
        Write-Debug -Message "$($this.Hostname)`> Attempting diagnosis..."
        try {
            $this.Connect()
            
            Write-Debug -Message "$($this.Hostname)`> Retrieving SNMP state..."
            $this.SNMP = Invoke-Command $this.Session -ScriptBlock { 
                if (Get-CimInstance -Class Win32_Service -Filter "Name='snmp'") {
                    return (Get-CimInstance -Class Win32_Service -Filter "Name='nscp'").State
                    
                } else {
                    return "Not installed"

                }

            } -ErrorAction Stop


            Write-Debug -Message "$($this.Hostname)`> Retrieving uptime..."
            $this.Uptime = Invoke-Command $this.Session -ScriptBlock { 
                return (Get-Uptime).Days

            } -ErrorAction Stop

            Write-Debug -Message "$($this.Hostname)`> Retrieving OSname..."
            $this.OSname = Invoke-Command $this.Session -ScriptBlock {
                $OS = (Get-ComputerInfo).OsName
                $OS = $OS -split " "
                $OS = $OS[1..3] #Keeping the relevant part of the OS name
                $OS = $OS -join " " 
                
                return $OS

            } -ErrorAction Stop
            
            if ($this.Hostname -match "FILTER0") {
                $this.MAX = 7
                $this.Restricted = $false
                
                if ((Get-Date).DayOfWeek -ne "Saturday"){ #Example of specific conditions applied on filtered servers
                    Write-Debug -Message "$($this.Hostname)`> Date mismatch."
                    $this.State = "Date mismatch"

                }

            } elseif ($this.Hostname -match "FILTER1") {
    
    
            } elseif ($this.Hostname -match "FILTER2") {
    
    
            } elseif ($this.Hostname -match "FILTER3") {
    
    
            } else {
                $this.MAX = 60
                $this.Restricted = $true
    
            }
            
            Write-Debug -Message "$($this.Hostname)`> Diagnosis completed."
            $this.State = "Diagnosed"
    
            return 0

        } catch {
            $this.Debug["200"] = $Error[0].Exception.Message
            $this.State = "Diagnosis failed"
            $this.Issue = $true
            Write-Error -Message "$($this.Hostname)`> Diagnosis failed!"

            return 200

        }

    }

    [int]Reboot() {
        Write-Debug -Message "$($this.Hostname)`> Attempting reboot..."
        try {
            Invoke-Command $this.Session -ScriptBlock { 
                Restart-Computer -Force 
                #Write-Host -ForegroundColor Magenta -Message "SCRIPT: Fake reboot."

            }

            Write-Debug -Message "$($this.Hostname)`> Pending..."
            $failure = 0
            $success = 0
            while (($success -lt 5) -and ($failure -lt 30)) {
                Start-Sleep -Seconds 5
                $answer = & ssh $this.Hostname echo "OK"
                if ($answer -eq "OK") {
                    $success++

                } else {
                    $success = 0
                    $failure++

                }

                Write-Debug -Message "$($this.Hostname)`> Success: $success/5 | Failure: $failure/30"

            } 

            if ($success -eq 5) {
                Write-Debug -Message "$($this.Hostname)`> Checking success..."
                $this.Diagnose()

                Write-Debug -Message "$($this.Hostname)`> Uptime: $($this.Uptime) Days"
                if ($this.Uptime -eq 0) {
                    Write-Host -ForegroundColor Green -Message "SUCCESS $($this.Hostname)`> Reboot succeed!"
                    $this.State = "Rebooted"

                    return 0

                } else {
                    Write-Warning -Message "$($this.Hostname)`> Reboot failed! Uptime wasn't reset."
                    $this.Debug["700"] = "Reboot failed! Uptime wasn't reset."
                    $this.State = "Reboot failed"
                    $this.Issue = $true

                    return 700

                }

            } else {
                Write-Warning -Message "$($this.Hostname)`> SSH connection timed out!"
                $this.Debug["900"] = "Host is down! SSH connection timed out."
                $this.State = "Murdered"
                $this.Issue = $true
                
                return 900
            }

        } catch {
            $this.Debug["100"] = $Error[0].Exception.Message
            $this.State = "Reboot failed"
            $this.Issue = $true
            Write-Error -Message "$($this.Hostname)`> Reboot failed!"

            return 100

        }

    }

    [int]Manage([int]$MAXreboot) {
        Write-Debug -Message "$($this.Hostname)`> Attempting management..."
        try {
            if ($this.Uptime -ge $this.MAX) {
                if (-not $this.Restricted) {
                    Write-Warning -Message "$($this.Hostname)`> Reboot required!"
                    $this.Reboot()
                    return 0

                } elseif ($MAXreboot -ge 0) { 
                    Write-Warning -Message "$($this.Hostname)`> Reboot required!"
                    $this.Reboot()
                    return 1

                } else {
                    Write-Warning -Message "$($this.Hostname)`> Reboot planned!"
                    $this.State = "Queued"
                    return 0
                }

            } elseif ($this.Uptime -ge [int]($this.MAX * 0.75)) {
                #When uptime is more than 75% of the max
                Write-Warning -Message "$($this.Hostname)`> Incoming management."
                $this.State = "Incoming management"
                return 0


            } else {
                Write-Host -ForegroundColor Green -Message "SUCCESS: $($this.Hostname)`> Uptime under maximum, no further operations required."
                #$this.State = "Running"
                $this.State = $null
                return 0

            }

        } catch {
            $this.Debug["777"] = $Error[0].Exception.Message
            $this.State = "Management failed"
            $this.Issue = $true
            Write-Error -Message "$($this.Hostname)`> Management failed!"
            
            return 0

        }

    }

}

# =====================================
# Récupération de la liste des serveurs
# =====================================

$ServersNamesList = @()
$ServersNamesList += (Get-ADComputer).DNSHostName

# ====================
# Gestion des serveurs 
# ====================

$Servers = @()

foreach ($ServerName in $ServersNamesList) {
    Write-Debug -Message "---------------------------------------------------"
    Write-Debug -Message "$ServerName`> Attempting registration..."
    try {
        $Servers += [Server]::new($ServerName)
        Write-Host -ForegroundColor Green -Message "SUCCESS: $ServerName`> Registration succeeded."

    } catch {
        Write-Error "CRITICAL : $($Error[0].Exception.Message)"
        exit

    }

}

$Servers = $Servers | Sort-Object -Descending Issue, Uptime 

foreach ($Server in $Servers | Where-Object {(-not $_.Issue) -and ($_.State -eq "Diagnosed")}) {
    Write-Debug -Message "---------------------------------------------------"
    $MAXreboot -= $Server.Manage($MAXreboot) #Method Manage() already contains try-catch
    
}


# ======================
# Préparation de l'email
# ======================

$smtpServer = "ipadress"
$from = "youradress"
$to = "teamadress"
$subject = "Servers Reboot Management"
$encoding = "UTF8"

# ====================
# Envoi de l'email
# ====================

$Header =
@"
<style>
    TABLE {
        border-width: 1px; 
        border-style: solid; 
        border-color: #000000; 
        border-collapse: collapse;
        text-align: center;
        text-color: #000000;
        background-color: #CFE3EA; 
    }
    TH {
        border-width: 2px; 
        padding: 10px; border-style: solid; 
        border-color: #000000; 
        background-color: #4CB2E9;
    }
    TD {
        border-width: 2px; 
        padding: 10px; 
        border-style: solid; 
        border-color: #000000;
    }
</style>
"@    

$Table = $Servers | Select-Object Hostname, OSname, Uptime, SNMP, State, Issue | ConvertTo-Html -Head $Header| ForEach-Object {$_ -replace "<td>True</td>", "<td style='background-color:#FF8080'></td>"} | ForEach-Object {$_ -replace "<td>False</td>", "<td style='background-color:#39D122'></td>"}| Out-String


$body = 
"
<html>
    <body>
        <p>
            Bonjour, <br>
            <br>
            Voici le diagnostic des serveurs Windows : <br>
            <br>
            $Table
        </p>
    </body>
</html>
"

$Reasontosend = $Servers | Where-Object {($_.Issue -eq $true) -or ($_.State -notmatch "Running")} 
if ($Reasontosend.Count -ne 0) {
    Send-MailMessage -bodyasHTML -smtpserver $smtpserver -from $from -to $to -subject $subject -body $body -Encoding $encoding
    Write-Host -ForegroundColor Magenta -Message "SCRIPT: Mail sent to $to."

} else {
    Write-Host -ForegroundColor Magenta -Message "SCRIPT: No problem nor restart planned. Mail won't be sent."

}

# ========
# Débogage
# ========

<#
foreach ($Server in $Servers | Where-Object {$_.Issue}) {
    Write-Host $Server.Hostname
    Write-Host $Server.Debug

}
#>
