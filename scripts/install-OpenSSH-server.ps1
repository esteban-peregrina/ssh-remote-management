<#Arrêt du service WinRM
$WinRM = Get-Service -Name "WinRM"
if ($WinRM.Status -eq "Running") {
    try {
        Stop-Service -Name $WinRM.Name -Force -ErrorAction Stop | Out-Null
        Write-Host "WinRM à été arrêté avec succès."
    } catch {
        Write-Error "Échec de l'arrêt de WinRM : $_"
        exit
    }
}#>


Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

#Installation de PowerShell 7
$pwsh7 = "$env:ProgramFiles\PowerShell\7\pwsh.exe"
if (Test-Path $pwsh7 -PathType Leaf) {
    Write-Host -ForegroundColor DarkGreen "PowerShell 7 est déja installé."
} else {
    Write-Host "PowerShell 7 n'est pas installé. Tentative d'installation..."
    try {
        $msionline = "your-shared-folder\OpenSSH\your-install-folder\PowerShell-7.3.4-win-x64.msi" #Example path
        $local = Join-Path $env:USERPROFILE "Desktop"
        Copy-Item -Path $msionline -Destination $local -Recurse
        $msilocal = Join-Path $local "PowerShell-7.3.4-win-x64.msi"
        Start-Process "msiexec.exe" -argumentList "/package $msilocal /quiet ADD_PATH=1 REGISTER_MANIFEST=1 ENABLE_PSREMOTING=1 ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_MU=0 USE_MU=1" -Wait -ErrorAction Stop | Out-Null
        Remove-Item -Path $msilocal
        Write-Host -ForegroundColor Green "PowerShell 7 a été installé avec succès."

    } catch {
        Write-Error "Échec de l'installation PowerShell 7 : $_"
        exit

    }

}

#Installation selon le système d'exploitation
$osversion = Get-CimInstance -Class Win32_OperatingSystem | Select-Object Caption
if ($osversion.Caption -eq "Microsoft Windows Server 2012 R2 Standard") {
    #Installation d'OpenSSH
    $ssh = Get-Service -Name "sshd" -ErrorAction SilentlyContinue
    if ($null -eq $ssh) {
        Write-Host "OpenSSH server n'est pas installé. Tentative d'installation..."
        try {
            #Installation d'OpenSSH
            $opensshpath = "your-shared-folder\OpenSSH"
            Copy-Item -Path $opensshpath -Destination $env:ProgramFiles -Recurse
            Set-Location $env:ProgramFiles\OpenSSH\
            & .\install-sshd.ps1 -ErrorAction Stop | Out-Null
            Write-Host -ForegroundColor Green "OpenSHH server a été installé avec succès."

            #Ajout de la variable d'environnement système
            $envpath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ";" + "$env:ProgramFiles" + "\OpenSSH\"
            [System.Environment]::SetEnvironmentVariable('Path', $envpath, 'Machine')
            Write-Host -ForegroundColor Green "La variable d'environnement `"path`" a été mise à jour."

        } catch {
            Write-Error "Échec de l'installation d'OpenSSH server : $_"
            exit

        }

    } else {
        Write-HosT -ForegroundColor DarkGreen "OpenSSH Server est déjà installé."
    }


} elseif (($osversion.Caption -eq "Microsoft Windows Server 2019 Standard") -or ($osversion.Caption -eq "Microsoft Windows 10 Professionnel") -or ($osversion.Caption -eq "Microsoft Windows 10 Professionnel pour les Stations de travail")) {
    #Installation d'OpenSSH
    $ssh = Get-Service -Name "sshd" -ErrorAction SilentlyContinue
    if ($null -eq $ssh) {
        Write-Host "Le serveur OpenSSH n'est pas installé. Tentative d'installation..." 
        try {
            Add-WindowsCapability -Online -Name OpenSSH.server~~~~0.0.1.0 -ErrorAction Stop | Out-Null
            Write-Host -ForegroundColor Green "Le serveur OpenSHH a été installé avec succès."

        } catch {
            Write-Error "Échec de l'installation du serveur OpenSSH : $_"
            exit

        }

    } else {
        Write-Host -ForegroundColor DarkGreen "OpenSSH Server est déjà installé."

    }
     
} else {
    Write-Error "La version du système d'exploitation n'est pas prise en charge par le script."
    exit

}

#Lancement du daemon
$daemon = Get-Service -Name "sshd"
if ($daemon.Status -eq "Running") {
    Write-Host -ForegroundColor DarkGreen "Le service OpenSSH Server est déjà lancé."
} else {
    try {
        Start-Service -Name $daemon.Name -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green "Le service OpenSSH Server a été lancé avec succès."

    } catch {
        Write-Error "Échec lors du lancement du service OpenSSH : $_" 
        exit

    }
}

#Activation du démarrage automatique du daemon
if ($daemon.StartType -eq "Automatic") {
    Write-Host -ForegroundColor DarkGreen "Le lancement automatique du service OpenSSH Server est déjà mis en place."
} else {
    try {
        Set-Service -Name $daemon.Name -StartupType Automatic -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green "La mise en place du lancement automatique du service OpenSSH Server a été réalisée avec succès."

    } catch {
        Write-Error "Échec de la mise en place du lancement automatique du service OpenSSH Server : $_"
        exit

    }
}

#Mise en place du subsystem PowerShell (attention, si le subsytème n'est pas configuré, le fichier de config sera écrasé avec celui du fichier de déploiement)
$sshdConfigPath = join-path $env:ProgramData\ssh "sshd_config"
try {
    if (Select-String -Path $sshdConfigPath -Pattern "powershell" -SimpleMatch -Quiet -ErrorAction SilentlyContinue) {
        Write-Host -ForegroundColor DarkGreen "Le fichier `"sshd_config`" est déjà configuré."
    
    } else { 
        Write-Host "Le fichier `"sshd_config`" n'est pas encore configuré. Tentative de configuration..."
        try {
            $config
            if (($osversion.Caption -eq "Microsoft Windows 10 Professionnel") -or ($osversion.Caption -eq "Microsoft Windows 10 Professionnel pour les Stations de travail")) {
                $configpath = "your-shared-folder\OpenSSH\your-install-folder\sshd_config_subsystem_WIN10" #Pre-configured config files
                $config = Get-Content -Path $configpath -ErrorAction Stop 
            } else {
                $configpath = "your-shared-folder\OpenSSH\your-install-folder\sshd_config_subsystem" #Pre-configured config files
                $config = Get-Content -Path $configpath -ErrorAction Stop  
            }
            $config | Set-Content $sshdConfigPath -ErrorAction Stop 
            Write-Host -ForegroundColor Green "La configuration du fichier `"sshd_config`" a été réalisée avec succès."
        
        } catch {
            Write-Error "Échec de la configuration du fichier `"sshd_config`" : $_"
            exit

        }

    }

} catch {
    Write-Error "Échec lors de la configuration du fichier `"sshd_config`" : $_"
    exit  
}


#Création du répertoire .ssh
$sshdir = join-path $env:USERPROFILE ".ssh"
if (Test-Path $sshdir -PathType Container) {
    Write-Host -ForegroundColor DarkGreen "Le répertoire `".ssh`" existe déja."

} else {
    Write-Host "Le répertoire `".ssh`" n'existe pas. Tentative de création..."
    try {
        New-Item $sshdir -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green "Le répertoire `".ssh`" a été créé avec succès."

    } catch {
        Write-Error "Échec de la création du répertoire `".ssh`" : $_"
        exit

    }

}

#Création du fichier "authorized_keys" (noter qu'il existe un fichier authorized_keys2)
$authorizedkeys = join-path $env:USERPROFILE\.ssh "authorized_keys"
if (Test-Path "$authorizedkeys" -PathType Leaf) { 
    Write-Host -ForegroundColor DarkGreen "Le fichier `"authorized_keys`" existe déja."

} else {
    try {
        New-Item $authorizedkeys -ItemType File -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green "Le fichier `'authorized_keys`' a été créé avec succès."

    } catch {
        Write-Error "Échec de la création du fichier `"authorized_keys`" : $_"
        exit

    }

}

#Insertion de la clef rsa de vp-cp-automat02
$rsakey = "ssh-rsa key username@hostname" #Example key (better to not store it directly in the code)
if (Select-String -Path $authorizedkeys -Pattern $rsakey -SimpleMatch -Quiet) {
    Write-Host -ForegroundColor DarkGreen "Le fichier `"authorized_keys`" contient déjà la clef RSA du serveur `"hostname`"."

} else { 
    try {
        Add-Content "$authorizedkeys" -Value $rsakey -ErrorAction Stop  | Out-Null
        #Fonctionnera tant que chaque écriture de clef (qu'elle vienne de ce script ou d'un autre) inclus bien un retour à la ligne.
        
        Write-Host -ForegroundColor Green "La clef ssh du serveur `"hostname`" a été ajoutée au fichier `"authorized_keys`" avec succès."
    
    } catch {
        Write-Host "Échec de l'ajout de la clef ssh du serveur `"hostname`" au fichier `"authorized_keys`" : $_"
        exit
    
    }

}

#Ajout de la règle de pare-feu
$firewallRule = Get-NetFirewallRule -Name "sshd" -ErrorAction SilentlyContinue
if ($null -eq $firewallRule) {
    Write-Host "La règle de pare-feu n'existe pas. Tentative de création..."
    try {    
        New-NetFirewallRule -Name sshd -DisplayName "Autoriser SSH" -Enabled True -Direction Inbound -Profile Domain -Protocol TCP -Action Allow -LocalPort 22 -ErrorAction Stop | Out-Null #Not safe enough 
        Write-Host -ForegroundColor Green "La règle de pare-feu a été ajoutée avec succès."

    } catch {
        Write-Error "Échec de l'ajout de la règle de pare-feu : $_"
        exit
    }
} else { 
    Write-Host -ForegroundColor DarkGreen  "La règle de pare-feu existe déjà."

}

#Redémarrage du service
try {    
    Restart-Service -Name "sshd" -ErrorAction Stop | Out-Null
    Write-Host -ForegroundColor Green "Le service OpenSSH a été redémarré avec succès."

} catch {
    Write-Error "Échec du redémarrage du service OpenSSH : $_"
    exit

}