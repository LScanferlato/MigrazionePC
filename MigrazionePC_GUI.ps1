Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ================================
#   FUNZIONE LOG
# ================================
function Write-Log {
    param(
        [string]$Message,
        [System.Windows.Forms.TextBox]$TextBox
    )
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $TextBox.AppendText("[$timestamp] $Message`r`n")
    $TextBox.ScrollToCaret()
}

# ================================
#   TROVA PERCORSO WINGET
# ================================
function Get-WingetPath {
    $base = "C:\Program Files\WindowsApps"

    $winget = Get-ChildItem $base -Filter "Microsoft.DesktopAppInstaller_*" -Directory -ErrorAction SilentlyContinue |
              ForEach-Object { Join-Path $_.FullName "winget.exe" } |
              Where-Object { Test-Path $_ } |
              Select-Object -First 1

    return $winget
}

# ================================
#   CONTROLLO EXECUTION POLICY E ADMIN
# ================================
function Check-ExecutionPolicy {
    param([System.Windows.Forms.TextBox]$LogBox)

    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        Write-Log "Esecuzione come amministratore rilevata." $LogBox
    } else {
        Write-Log "ATTENZIONE: PowerShell NON è avviato come amministratore." $LogBox
        Write-Log "Provo a impostare ExecutionPolicy per l'utente corrente..." $LogBox

        try {
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            Write-Log "ExecutionPolicy impostata correttamente per l'utente corrente." $LogBox
        }
        catch {
            Write-Log "Impossibile impostare ExecutionPolicy. Avvia PowerShell come amministratore." $LogBox
        }
    }
}

# ================================
#   BARRA DI AVANZAMENTO
# ================================
function Start-ProgressBar {
    param($Bar)
    $Bar.Visible = $true
    $Bar.Style = 'Marquee'
}

function Stop-ProgressBar {
    param($Bar)
    $Bar.Visible = $false
}

# ================================
#   ESPORTAZIONE DATI
# ================================
function Esporta-Dati {
    param(
        [string]$Utente,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar
    )

    if (-not $Utente) {
        Write-Log "Errore: seleziona un profilo utente prima di esportare." $LogBox
        return
    }

    $origineProfilo = "C:\Users\$Utente"
    $destinazione = "C:\MigrazionePC"
    $destProfilo = "$destinazione\ProfiloUtente"

    Write-Log "Inizio esportazione per utente '$Utente'..." $LogBox

    if (-not (Test-Path $origineProfilo)) {
        Write-Log "Errore: il profilo $origineProfilo non esiste." $LogBox
        return
    }

    try {
        Write-Log "Creo cartella di migrazione: $destinazione" $LogBox
        New-Item -ItemType Directory -Force -Path $destinazione | Out-Null

        # Garantisce permessi completi alla cartella
        try {
            $acl = Get-Acl $destinazione
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                "$env:USERNAME","FullControl","ContainerInherit,ObjectInherit","None","Allow"
            )
            $acl.SetAccessRule($rule)
            Set-Acl $destinazione $acl
            Write-Log "Permessi aggiornati su $destinazione per l'utente corrente." $LogBox
        }
        catch {
            Write-Log "Impossibile aggiornare i permessi su ${destinazione}: $($_.Exception.Message)" $LogBox
        }

        # Winget export tramite cmd.exe
        $wg = Get-WingetPath
        if ($wg) {
            Write-Log "Esporto lista programmi winget..." $LogBox
            cmd.exe /c "`"$wg`" export -o `"$destinazione\programmi.json`" > `"$destinazione\winget_export.log`" 2>&1"
        } else {
            Write-Log "Winget non trovato. Salto esportazione programmi." $LogBox
        }

        Write-Log "Esporto variabili di ambiente..." $LogBox
        Get-ChildItem Env: | Export-Csv "$destinazione\variabili.csv" -NoTypeInformation

        Write-Log "Copio l'intero profilo utente (può richiedere tempo)..." $LogBox
        New-Item -ItemType Directory -Force -Path $destProfilo | Out-Null

        Start-ProgressBar -Bar $ProgressBar

        $robocopyCmd = "robocopy `"$origineProfilo`" `"$destProfilo`" /MIR /XJ /R:1 /W:1"
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $robocopyCmd" -NoNewWindow -RedirectStandardOutput "$destinazione\robocopy_export.log" -PassThru
        $process.WaitForExit()

        Stop-ProgressBar -Bar $ProgressBar
        Write-Log "Copia profilo completata. Log salvato in robocopy_export.log" $LogBox

        Write-Log "ESPORTAZIONE COMPLETATA." $LogBox
        Write-Log "Cartella pronta: $destinazione (copiala sul nuovo PC)." $LogBox
    }
    catch {
        Stop-ProgressBar -Bar $ProgressBar
        Write-Log "Errore durante l'esportazione: $($_.Exception.Message)" $LogBox
    }
}

# ================================
#   IMPORTAZIONE DATI
# ================================
function Importa-Dati {
    param(
        [string]$Utente,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar
    )

    if (-not $Utente) {
        Write-Log "Errore: seleziona un profilo utente prima di importare." $LogBox
        return
    }

    $destinazione = "C:\MigrazionePC"
    $profiloDest = "C:\Users\$Utente"

    Write-Log "Inizio importazione per utente '$Utente'..." $LogBox

    if (-not (Test-Path $destinazione)) {
        Write-Log "Errore: la cartella $destinazione non esiste. Copiala dal vecchio PC." $LogBox
        return
    }

    try {
        # Garantisce permessi completi alla cartella
        try {
            $acl = Get-Acl $destinazione
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                "$env:USERNAME","FullControl","ContainerInherit,ObjectInherit","None","Allow"
            )
            $acl.SetAccessRule($rule)
            Set-Acl $destinazione $acl
            Write-Log "Permessi aggiornati su $destinazione per l'utente corrente." $LogBox
        }
        catch {
            Write-Log "Impossibile aggiornare i permessi su ${destinazione}: $($_.Exception.Message)" $LogBox
        }

        # Winget import tramite cmd.exe
        $wg = Get-WingetPath
        if ($wg -and (Test-Path "$destinazione\programmi.json")) {
            Write-Log "Reinstallo programmi tramite winget..." $LogBox
            cmd.exe /c "`"$wg`" import -i `"$destinazione\programmi.json`" > `"$destinazione\winget_import.log`" 2>&1"
        } else {
            Write-Log "Winget non trovato o programmi.json mancante. Salto reinstallazione programmi." $LogBox
        }

        if (Test-Path "$destinazione\variabili.csv") {
            Write-Log "Ripristino variabili di ambiente..." $LogBox
            Import-Csv "$destinazione\variabili.csv" | ForEach-Object { setx $_.Name $_.Value | Out-Null }
        }

        if (Test-Path "$destinazione\ProfiloUtente") {
            Write-Log "Ripristino profilo utente (può richiedere tempo)..." $LogBox

            Start-ProgressBar -Bar $ProgressBar

            $robocopyCmd = "robocopy `"$destinazione\ProfiloUtente`" `"$profiloDest`" /MIR /XJ /R:1 /W:1"
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $robocopyCmd" -NoNewWindow -RedirectStandardOutput "$destinazione\robocopy_import.log" -PassThru
            $process.WaitForExit()

            Stop-ProgressBar -Bar $ProgressBar
            Write-Log "Ripristino profilo completato. Log salvato in robocopy_import.log" $LogBox
        }

        Write-Log "IMPORTAZIONE COMPLETATA." $LogBox
    }
    catch {
        Stop-ProgressBar -Bar $ProgressBar
        Write-Log "Errore durante l'importazione: $($_.Exception.Message)" $LogBox
    }
}

# ================================
#   COSTRUZIONE INTERFACCIA GRAFICA
# ================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Migrazione PC"
$form.Size = New-Object System.Drawing.Size(650, 460)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

# Label utente
$labelUtente = New-Object System.Windows.Forms.Label
$labelUtente.Text = "Seleziona il profilo utente:"
$labelUtente.Location = New-Object System.Drawing.Point(20, 20)
$labelUtente.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($labelUtente)

# ComboBox utenti
$comboUtenti = New-Object System.Windows.Forms.ComboBox
$comboUtenti.Location = New-Object System.Drawing.Point(220, 18)
$comboUtenti.Size = New-Object System.Drawing.Size(200, 20)
$comboUtenti.DropDownStyle = "DropDownList"
$form.Controls.Add($comboUtenti)

# Carica profili da C:\Users
Get-ChildItem "C:\Users" -Directory | ForEach-Object {
    $comboUtenti.Items.Add($_.Name)
}

# Seleziona automaticamente l’utente corrente
$currentUser = $env:USERNAME
if ($comboUtenti.Items.Contains($currentUser)) {
    $comboUtenti.SelectedItem = $currentUser
}

# Pulsante Esporta
$btnEsporta = New-Object System.Windows.Forms.Button
$btnEsporta.Text = "Esporta dal vecchio PC"
$btnEsporta.Location = New-Object System.Drawing.Point(20, 60)
$btnEsporta.Size = New-Object System.Drawing.Size(180, 35)
$form.Controls.Add($btnEsporta)

# Pulsante Importa
$btnImporta = New-Object System.Windows.Forms.Button
$btnImporta.Text = "Importa sul nuovo PC"
$btnImporta.Location = New-Object System.Drawing.Point(240, 60)
$btnImporta.Size = New-Object System.Drawing.Size(180, 35)
$form.Controls.Add($btnImporta)

# Pulsante Esci
$btnEsci = New-Object System.Windows.Forms.Button
$btnEsci.Text = "Esci"
$btnEsci.Location = New-Object System.Drawing.Point(450, 60)
$btnEsci.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($btnEsci)

# TextBox log
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20, 110)
$logBox.Size = New-Object System.Drawing.Size(550, 270)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$form.Controls.Add($logBox)

# Barra di avanzamento
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 390)
$progressBar.Size = New-Object System.Drawing.Size(550, 20)
$progressBar.Style = 'Marquee'
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Eventi pulsanti
$btnEsporta.Add_Click({
    Esporta-Dati -Utente $comboUtenti.SelectedItem -LogBox $logBox -ProgressBar $progressBar
})

$btnImporta.Add_Click({
    Importa-Dati -Utente $comboUtenti.SelectedItem -LogBox $logBox -ProgressBar $progressBar
})

$btnEsci.Add_Click({
    $form.Close()
})

# Messaggio iniziale
Check-ExecutionPolicy -LogBox $logBox
Write-Log "Benvenuto nella Migrazione PC con interfaccia grafica." $logBox
Write-Log "1) Sul vecchio PC: seleziona il profilo e clicca 'Esporta dal vecchio PC'." $logBox
Write-Log "2) Copia C:\MigrazionePC sul nuovo PC." $logBox
Write-Log "3) Sul nuovo PC: seleziona il profilo e clicca 'Importa sul nuovo PC'." $logBox

# Avvia form
[void]$form.ShowDialog()
