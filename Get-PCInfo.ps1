# flag za lazy loading
$Script:LazyLoading = $true
$Script:GeneratedTabs = @{}



$Script:AppVersion = "2025.6.3"
$Script:AppAuthor = "Ivica Rašan"

# Poboljšana logging funkcija
function Write-PCInfoLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $colors = @{ "INFO" = "Green"; "WARNING" = "Yellow"; "ERROR" = "Red"; "DEBUG" = "Cyan" }
    Write-Host "[$timestamp] $Message" -ForegroundColor $colors[$Level]
    
    # Dodaj u log datoteku
    try {
        "$timestamp [$Level] $Message" | Add-Content "$env:TEMP\Get-PCInfo.log" -ErrorAction SilentlyContinue
    } catch { }
}





# Provjera i automatsko podizanje privilegija ako skripta nije pokrenuta s administrativnim ovlastima
$CurrentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $CurrentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Ponovno pokrenite skriptu s administrativnim pravima
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}


# === SAMOPOKRETANJE MINIMIZIRANO AKO NIJE VEÆ ===
if (-not $env:GUI_LAUNCHED) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -NoLogo -NoProfile -File `"$scriptPath`""
    $psi.UseShellExecute = $false
    $psi.EnvironmentVariables["GUI_LAUNCHED"] = "1"
    $psi.WindowStyle = 'Minimized'
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}


# ========================================
# LOADING SCREEN - poèetak
# ========================================

# Uèitajte assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# KREIRANJE LOADING FORME s mega features
$loadingForm = New-Object System.Windows.Forms.Form
$loadingForm.Size = New-Object System.Drawing.Size(900, 650)  # Veæi za dodatne features
$loadingForm.StartPosition = "CenterScreen"
$loadingForm.BackColor = [System.Drawing.Color]::FromArgb(45, 50, 55)
$loadingForm.FormBorderStyle = "None"
$loadingForm.TopMost = $true
$loadingForm.Opacity = 0.0  # Poèinje nevidljivo za fade-in

# Naslovni label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Get-PCInfo Professional ver. 2025.6.3"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::LightBlue
    $titleLabel.Location = New-Object System.Drawing.Point(50, 40)
    $titleLabel.Size = New-Object System.Drawing.Size(800, 60)
    $titleLabel.TextAlign = "MiddleCenter"
    $loadingForm.Controls.Add($titleLabel)

# Subtitle
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Advanced System Management Suite"
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Italic)
$subtitleLabel.ForeColor = [System.Drawing.Color]::LightGray
$subtitleLabel.Location = New-Object System.Drawing.Point(50, 100)
$subtitleLabel.Size = New-Object System.Drawing.Size(800, 30)
$subtitleLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($subtitleLabel)

# OPIS APLIKACIJE
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "Aplikacija kombinira moguænost PowerShell-a s intuitivnim grafièkim suèeljem, omoguæujuæi sistemskim administratorima potpunu kontrolu nad lokalnim korisnièkim raèunima, hardverskim resursima, sigurnosnim komponentama i sustavskim dogaðajima."
$descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$descriptionLabel.ForeColor = [System.Drawing.Color]::White
$descriptionLabel.Location = New-Object System.Drawing.Point(80, 140)
$descriptionLabel.Size = New-Object System.Drawing.Size(740, 70)
$descriptionLabel.TextAlign = "MiddleLeft"
$loadingForm.Controls.Add($descriptionLabel)

# SYSTEM INFO PANEL
$systemInfoLabel = New-Object System.Windows.Forms.Label
$systemInfoLabel.Text = "System: $env:COMPUTERNAME | OS: Windows | User: $env:USERNAME"
$systemInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
$systemInfoLabel.ForeColor = [System.Drawing.Color]::LightGreen
$systemInfoLabel.Location = New-Object System.Drawing.Point(50, 220)
$systemInfoLabel.Size = New-Object System.Drawing.Size(800, 20)
$systemInfoLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($systemInfoLabel)

# Loading indikator s floating particles
$loadingLabel = New-Object System.Windows.Forms.Label
$loadingLabel.Text = "?"
$loadingLabel.Font = New-Object System.Drawing.Font("Segoe UI", 32, [System.Drawing.FontStyle]::Bold)
$loadingLabel.ForeColor = [System.Drawing.Color]::LightBlue
$loadingLabel.Location = New-Object System.Drawing.Point(425, 250)
$loadingLabel.Size = New-Object System.Drawing.Size(60, 60)
$loadingLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($loadingLabel)

# FLOATING PARTICLES (3 mala elementa)
$particle1 = New-Object System.Windows.Forms.Label
$particle1.Text = "•"
$particle1.Font = New-Object System.Drawing.Font("Segoe UI", 16)
$particle1.ForeColor = [System.Drawing.Color]::Orange
$particle1.Location = New-Object System.Drawing.Point(300, 270)
$particle1.Size = New-Object System.Drawing.Size(20, 20)
$loadingForm.Controls.Add($particle1)

$particle2 = New-Object System.Windows.Forms.Label
$particle2.Text = "•"
$particle2.Font = New-Object System.Drawing.Font("Segoe UI", 16)
$particle2.ForeColor = [System.Drawing.Color]::Yellow
$particle2.Location = New-Object System.Drawing.Point(580, 270)
$particle2.Size = New-Object System.Drawing.Size(20, 20)
$loadingForm.Controls.Add($particle2)

$particle3 = New-Object System.Windows.Forms.Label
$particle3.Text = "•"
$particle3.Font = New-Object System.Drawing.Font("Segoe UI", 16)
$particle3.ForeColor = [System.Drawing.Color]::Cyan
$particle3.Location = New-Object System.Drawing.Point(450, 200)
$particle3.Size = New-Object System.Drawing.Size(20, 20)
$loadingForm.Controls.Add($particle3)

# Progress bar container
$progressContainer = New-Object System.Windows.Forms.Panel
$progressContainer.BackColor = [System.Drawing.Color]::FromArgb(30, 35, 40)
$progressContainer.Location = New-Object System.Drawing.Point(80, 340)
$progressContainer.Size = New-Object System.Drawing.Size(640, 35)
$loadingForm.Controls.Add($progressContainer)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(3, 3)
$progressBar.Size = New-Object System.Drawing.Size(634, 29)
$progressBar.Style = "Continuous"
$progressBar.Value = 0
$progressBar.Maximum = 100
$progressContainer.Controls.Add($progressBar)

# Progress percentage s estimated time
$progressPercentLabel = New-Object System.Windows.Forms.Label
$progressPercentLabel.Text = "0%"
$progressPercentLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$progressPercentLabel.ForeColor = [System.Drawing.Color]::LightBlue
$progressPercentLabel.Location = New-Object System.Drawing.Point(750, 345)
$progressPercentLabel.Size = New-Object System.Drawing.Size(50, 30)
$progressPercentLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($progressPercentLabel)

# ESTIMATED TIME LABEL
$timeEstimateLabel = New-Object System.Windows.Forms.Label
$timeEstimateLabel.Text = "Procjenjujem..."
$timeEstimateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$timeEstimateLabel.ForeColor = [System.Drawing.Color]::Orange
$timeEstimateLabel.Location = New-Object System.Drawing.Point(80, 380)
$timeEstimateLabel.Size = New-Object System.Drawing.Size(200, 20)
$timeEstimateLabel.TextAlign = "MiddleLeft"
$loadingForm.Controls.Add($timeEstimateLabel)

# Status label s funny messages
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Pokretanje aplikacije..."
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusLabel.Location = New-Object System.Drawing.Point(50, 410)
$statusLabel.Size = New-Object System.Drawing.Size(800, 30)
$statusLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($statusLabel)

# TIPS LABEL
$tipsLabel = New-Object System.Windows.Forms.Label
$tipsLabel.Text = "Tip: Pritisnite ESC za izlaz iz aplikacije"
$tipsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$tipsLabel.ForeColor = [System.Drawing.Color]::LightYellow
$tipsLabel.Location = New-Object System.Drawing.Point(50, 450)
$tipsLabel.Size = New-Object System.Drawing.Size(800, 25)
$tipsLabel.TextAlign = "MiddleCenter"
$loadingForm.Controls.Add($tipsLabel)

# MINI LOG WINDOW
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Multiline = $true
$logTextBox.ReadOnly = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.BackColor = [System.Drawing.Color]::Black
$logTextBox.ForeColor = [System.Drawing.Color]::LightGreen
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
$logTextBox.Location = New-Object System.Drawing.Point(280, 480)
$logTextBox.Size = New-Object System.Drawing.Size(550, 65)
$logTextBox.Text = "[LOG] Aplikacija pokrenuta`r`n[LOG] Uèitavanje modula...`r`n"
$loadingForm.Controls.Add($logTextBox)

# Version i copyright
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "Verzija 2025.6.3 Professional Edition"
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$versionLabel.ForeColor = [System.Drawing.Color]::LightGray
$versionLabel.Location = New-Object System.Drawing.Point(50, 580)
$versionLabel.Size = New-Object System.Drawing.Size(400, 20)
$versionLabel.TextAlign = "MiddleLeft"
$loadingForm.Controls.Add($versionLabel)

$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Text = "© 2025 - Made with PowerShell"
$copyrightLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
$copyrightLabel.Location = New-Object System.Drawing.Point(50, 600)
$copyrightLabel.Size = New-Object System.Drawing.Size(400, 20)
$copyrightLabel.TextAlign = "MiddleLeft"
$loadingForm.Controls.Add($copyrightLabel)

$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Text = "© Ivica Rašan 2025"
$authorLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$authorLabel.ForeColor = [System.Drawing.Color]::Gray
$authorLabel.Location = New-Object System.Drawing.Point(50, 615)
$authorLabel.Size = New-Object System.Drawing.Size(400, 20)
$authorLabel.TextAlign = "MiddleLeft"
$loadingForm.Controls.Add($authorLabel)

# Pokažite loading formu
$loadingForm.Show()

# FADE-IN ANIMACIJA
for ($i = 0; $i -le 20; $i++) {
    $loadingForm.Opacity = $i / 20.0
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 25
}

# VARIJABLE ZA ANIMACIJE
$script:animationIndex = 0
$script:particleDirection1 = 1
$script:particleDirection2 = -1
$script:particleDirection3 = 1
$script:startTime = Get-Date
$script:tipIndex = 0

# FUNNY LOADING MESSAGES
$funnyMessages = @(
    "Budim server iz zimskog sna... ??",
    "Tražim izgubljene fajlove... ??", 
    "Uèim tabove da se ponašaju... ??",
    "Pregovaram s Windows-om... ??",
    "Èekam da se kava skuha... ?",
    "Pretvaram se da radim... ??",
    "Uèitavam super tajni kod... ???",
    "Rješavam Rubikovu kocku... ??",
    "Tražim smisao života u kodovima... ??"
)

# TIPS ARRAY
$loadingTips = @(
    "Tip: Pritisnite ESC za izlaz iz aplikacije",
    "Tip: Koristite Ctrl+F za brzu pretragu",
    "Tip: Desni klik za kontekst menu",
    "Tip: Tab se mogu reorganizovati drag & drop",
    "Tip: F5 osvježava podatke u aktivnom tab-u",
    "Tip: Ctrl+E eksportuje podatke u Excel",
    "Tip: Alt+F4 zatvara aplikaciju"
)

# MEGA TIMER s SVIM ANIMACIJAMA
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200
$timer.Add_Tick({
    # 1. MAIN LOADING SPINNER
    $spinners = @("?", "??", "???", "????", "?????", "????", "???", "??", "?")
    $colors = @([System.Drawing.Color]::LightBlue, [System.Drawing.Color]::LightGreen, 
               [System.Drawing.Color]::Orange, [System.Drawing.Color]::Yellow, 
               [System.Drawing.Color]::Cyan, [System.Drawing.Color]::LightCoral)
    
    $loadingLabel.Text = $spinners[$script:animationIndex % $spinners.Length]
    $loadingLabel.ForeColor = $colors[$script:animationIndex % $colors.Length]
    
    # 2. FLOATING PARTICLES ANIMACIJA
    # Particle 1 - horizontal movement
    $newX1 = $particle1.Location.X + (5 * $script:particleDirection1)
    if ($newX1 -le 250 -or $newX1 -ge 350) { $script:particleDirection1 *= -1 }
    $particle1.Location = New-Object System.Drawing.Point($newX1, $particle1.Location.Y)
    
    # Particle 2 - horizontal movement (opposite)
    $newX2 = $particle2.Location.X + (4 * $script:particleDirection2)
    if ($newX2 -le 550 -or $newX2 -ge 650) { $script:particleDirection2 *= -1 }
    $particle2.Location = New-Object System.Drawing.Point($newX2, $particle2.Location.Y)
    
    # Particle 3 - vertical movement
    $newY3 = $particle3.Location.Y + (3 * $script:particleDirection3)
    if ($newY3 -le 180 -or $newY3 -ge 220) { $script:particleDirection3 *= -1 }
    $particle3.Location = New-Object System.Drawing.Point($particle3.Location.X, $newY3)
    
    # 3. TITLE PULSIRANJE
    if ($script:animationIndex % 6 -eq 0) {
        $titleLabel.ForeColor = [System.Drawing.Color]::White
    } else {
        $titleLabel.ForeColor = [System.Drawing.Color]::LightBlue
    }
    
    # 4. FUNNY MESSAGES (svakih 15 ciklusa)
    if ($script:animationIndex % 15 -eq 0) {
        $randomMsg = $funnyMessages[(Get-Random -Maximum $funnyMessages.Length)]
        $statusLabel.Text = $randomMsg
    }
    
    # 5. TIPS ROTATION (svakih 25 ciklusa)
    if ($script:animationIndex % 25 -eq 0) {
        $tipsLabel.Text = $loadingTips[$script:tipIndex % $loadingTips.Length]
        $script:tipIndex++
    }
    
    $script:animationIndex++
})
$timer.Start()

# POBOLJŠANA FUNKCIJA ZA AŽURIRANJE
function Update-LoadingStatus {
    param($message, $progress = $null)
    
    # Dodaj u log
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $logTextBox.AppendText("[LOG $timestamp] $message`r`n")
    $logTextBox.SelectionStart = $logTextBox.Text.Length
    $logTextBox.ScrollToCaret()
    
    # Update status
    $statusLabel.Text = $message
    
    if ($progress -ne $null) {
        $progressValue = [Math]::Min($progress, 100)
        $progressBar.Value = $progressValue
        $progressPercentLabel.Text = "$progressValue%"
        
        # GRADIENT PROGRESS BAR COLORS
        if ($progressValue -lt 30) {
            $progressBar.ForeColor = [System.Drawing.Color]::Red
        } elseif ($progressValue -lt 70) {
            $progressBar.ForeColor = [System.Drawing.Color]::Orange
        } else {
            $progressBar.ForeColor = [System.Drawing.Color]::LightGreen
        }
        
        # ESTIMATED TIME CALCULATION
        $elapsed = (Get-Date) - $script:startTime
        if ($progressValue -gt 5) {
            $totalEstimated = ($elapsed.TotalSeconds / $progressValue) * 100
            $remaining = [Math]::Max(0, $totalEstimated - $elapsed.TotalSeconds)
            $timeEstimateLabel.Text = " $([Math]::Round($remaining, 0)) sekundi preostalo"
        }
    }
    
    [System.Windows.Forms.Application]::DoEvents()
    
    # SOUND EFFECT za milestone-ove
    if ($progress -eq 25 -or $progress -eq 50 -or $progress -eq 75 -or $progress -eq 100) {
        [Console]::Beep(800, 100)
    }
}


# ========================================
# LOADING SCREEN - kraj
# ========================================








Update-LoadingStatus "Kreiranje glavne forme..." 20

# === Glavna forma ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Get-PCInfo Professional ver. 2025.6.3 / @Ivica Rašan 2025"
$form.Size = New-Object System.Drawing.Size(1000, 850)
$form.StartPosition = "CenterScreen"
Write-PCInfoLog "=== Get-PCInfo Professional v$Script:AppVersion Starting ==="
Write-PCInfoLog "Author: $Script:AppAuthor"
Write-PCInfoLog "System: $env:COMPUTERNAME, User: $env:USERNAME"
Write-PCInfoLog "PowerShell: $($PSVersionTable.PSVersion)"


# === TabControl ===
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$tabControl.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)


# === Funkcija: Dodaj tab koji prikazuje HTML (korisnici) ===
function Add-TabWeb {
    param (
        [string]$name,
        [string]$htmlFilePath,
        [scriptblock]$GenerationFunction = $null
    )

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $name

    if ($Script:LazyLoading -and $GenerationFunction) {
        # Dodajte placeholder umjesto stvarnog sadržaja
        $placeholder = New-Object System.Windows.Forms.Label
        $placeholder.Text = "Uèitavanje... Kliknite za generiranje izvještaja"
        $placeholder.Dock = "Fill"
        $placeholder.TextAlign = "MiddleCenter"
        $placeholder.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $placeholder.BackColor = [System.Drawing.Color]::LightGray
        
        # Dodajte click event za generiranje
        $placeholder.Add_Click({
            Generate-TabContent -TabName $name -GenerationFunction $GenerationFunction -Tab $tab
        })
        
        $tab.Controls.Add($placeholder)
        
        # Saèuvajte generation function za kasnije
        $Script:GeneratedTabs[$name] = @{
            Generated = $false
            Function = $GenerationFunction
            Tab = $tab
        }
    } else {
        # Postojeæa logika za veæ generirane izvještaje
        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.Dock = "Fill"
        $browser.ScriptErrorsSuppressed = $true
        $browser.Url = (New-Object System.Uri($htmlFilePath))
        $tab.Controls.Add($browser)
    }

    $tabControl.TabPages.Add($tab)
}



#funkcija za generiranje sadržaja na zahtjev
function Generate-TabContent {
    param(
        [string]$TabName,
        [scriptblock]$GenerationFunction,
        $Tab
    )
    
    if ($Script:GeneratedTabs[$TabName].Generated) {
        return
    }
    
    # Prikaži progress
    $Tab.Controls.Clear()
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Generiranje $TabName izvještaja..."
    $progressLabel.Dock = "Fill"
    $progressLabel.TextAlign = "MiddleCenter"
    $progressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $Tab.Controls.Add($progressLabel)
    
    # Async generiranje
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    
    $ps.AddScript({
        param($FunctionScript, $OutputPath)
        & $FunctionScript -OutputPath $OutputPath
    }).AddArgument($GenerationFunction).AddArgument("$env:TEMP\$TabName.html")
    
    $async = $ps.BeginInvoke()
    
    # Timer za provjeru završetka
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick({
        if ($async.IsCompleted) {
            $timer.Stop()
            $ps.EndInvoke($async)
            $ps.Dispose()
            $runspace.Close()
            
            # Zamijeni placeholder s WebBrowser kontrolom
            $Tab.Controls.Clear()
            $browser = New-Object System.Windows.Forms.WebBrowser
            $browser.Dock = "Fill"
            $browser.ScriptErrorsSuppressed = $true
            $browser.Url = (New-Object System.Uri("$env:TEMP\$TabName.html"))
            $Tab.Controls.Add($browser)
            
            $Script:GeneratedTabs[$TabName].Generated = $true
        }
    })
    $timer.Start()
}



# === Funkcija za èišæenje HTML teksta ===
function Clean-HTMLText {
    param([string]$text)
    
    if (-not $text) { return "" }
    
    # Ukloni sve HTML tagove
    $cleaned = $text -replace '<[^>]+>', ''
    
    # Ukloni HTML entitete
    $cleaned = $cleaned -replace '&nbsp;', ' '
    $cleaned = $cleaned -replace '&amp;', '&'
    $cleaned = $cleaned -replace '&lt;', '<'
    $cleaned = $cleaned -replace '&gt;', '>'
    $cleaned = $cleaned -replace '&quot;', '"'
    $cleaned = $cleaned -replace '&#39;', "'"
    
    # Oèisti whitespace
    $cleaned = $cleaned -replace '\r?\n', ' '
    $cleaned = $cleaned -replace '\s+', ' '
    
    return $cleaned.Trim()
}

# === Funkcija za ekstraktiranje podataka iz tablice ===
function Extract-TableData {
    param([string]$html)
    
    try {
        $tableData = @()
        
        Write-Host "Debug: Tražim tablicu u HTML-u..." -ForegroundColor Yellow
        
        # Fleksibilniji regex pattern za tablicu
        $tablePattern = '(?s)<table[^>]*>.*?</table>'
        $tableMatch = [regex]::Match($html, $tablePattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        if ($tableMatch.Success) {
            Write-Host "Debug: Pronašao tablicu!" -ForegroundColor Green
            $tableContent = $tableMatch.Value
            
            # Ukloni thead/tbody/tfoot wrappere ako postoje
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Fleksibilniji pattern za redove
            $rowPattern = '(?s)<tr[^>]*>.*?</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            Write-Host "Debug: Pronašao $($rowMatches.Count) redova" -ForegroundColor Cyan
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Value
                $cells = @()
                
                # Fleksibilniji pattern za æelije
                $cellPattern = '(?s)<t[hd][^>]*>.*?</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Value
                    $cleanCell = Clean-HTMLText -text $cellContent
                    
                    # Skrati predugaèke SID-ove
                    if ($cleanCell -match '^S-1-5-21-\d+-\d+-\d+-\d+$') {
                        $cleanCell = "..." + $cleanCell.Substring($cleanCell.Length - 10, 10)
                    }
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    Write-Host "Debug: Red ima $($cells.Count) æelija: $($cells -join ' | ')" -ForegroundColor Magenta
                    $tableData += ,$cells
                }
            }
        } else {
            Write-Host "Debug: Nema tablice, pokušavam alternativni pristup..." -ForegroundColor Red
            
            # Alternativni pristup - traži strukturu podataka direktno
            $lines = $html -split "`n"
            
            foreach ($line in $lines) {
                # Provjeri sadrži li linija podatke koji izgledaju kao tablièke
                if ($line -match '\w+.*\w+.*\w+' -and $line -notmatch '<style|<script|body \{') {
                    $cleaned = Clean-HTMLText -text $line
                    if ($cleaned.Length -gt 20) {
                        # Pokušaj podijeliti podatke na stupce
                        $parts = $cleaned -split '\s{2,}' | Where-Object { $_.Trim() -ne '' }
                        if ($parts.Count -gt 3) {
                            Write-Host "Debug: Alternativni red: $($parts -join ' | ')" -ForegroundColor Yellow
                            $tableData += ,$parts
                        }
                    }
                }
            }
        }
        
        Write-Host "Debug: Ukupno pronašao $($tableData.Count) redova podataka" -ForegroundColor Green
        return $tableData
    }
    catch {
        Write-Warning "Greška pri ekstraktiranju tablice: $_"
        return @()
    }
}


# === Funkcija za lijepo formatiranje tablice ===
function Format-TableNicely {
    param([array]$tableData)
    
    try {
        if ($tableData.Count -eq 0) {
            return @("Nema podataka")
        }

        # Dodaj redni broj kao prvu kolonu
        $headers = @("Br.") + $tableData[0]
        $dataRows = @()
        
        for ($i = 1; $i -lt $tableData.Count; $i++) {
            $dataRows += ,(@($i.ToString()) + $tableData[$i])
        }

        $allRows = @($headers) + $dataRows

        # Izraèunaj optimalne širine kolona
        $colWidths = @()
        for ($col = 0; $col -lt $allRows[0].Count; $col++) {
            $maxWidth = 0
            foreach ($row in $allRows) {
                if ($col -lt $row.Count) {
                    $width = $row[$col].ToString().Length
                    if ($width -gt $maxWidth) { $maxWidth = $width }
                }
            }
            # Ogranièi širinu na maksimalno 30 znakova
            $colWidths += [Math]::Min($maxWidth, 30)
        }

        # Generiraj tablicu
        $result = @()
        
        # Gornja linija
        $topLine = "-" + (($colWidths | ForEach-Object { "¦" * ($_ + 2) }) -join "T") + "¬"
        $result += $topLine

        # Header red
        $headerRow = "-"
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $cell = $headers[$col].ToString()
            if ($cell.Length -gt $colWidths[$col]) {
                $cell = $cell.Substring(0, $colWidths[$col] - 3) + "..."
            }
            $headerRow += " " + $cell.PadRight($colWidths[$col]) + " -"
        }
        $result += $headerRow

        # Separator linija
        $sepLine = "+" + (($colWidths | ForEach-Object { "¦" * ($_ + 2) }) -join "+") + "+"
        $result += $sepLine

        # Data redovi
        foreach ($row in $dataRows) {
            $dataRow = "-"
            for ($col = 0; $col -lt $row.Count; $col++) {
                $cell = if ($col -lt $row.Count) { $row[$col].ToString() } else { "" }
                if ($cell.Length -gt $colWidths[$col]) {
                    $cell = $cell.Substring(0, $colWidths[$col] - 3) + "..."
                }
                $dataRow += " " + $cell.PadRight($colWidths[$col]) + " -"
            }
            $result += $dataRow
        }

        # Donja linija
        $bottomLine = "L" + (($colWidths | ForEach-Object { "¦" * ($_ + 2) }) -join "+") + "-"
        $result += $bottomLine

        return $result
    }
    catch {
        Write-Warning "Greška pri formatiranju: $_"
        return @("Greška pri formatiranju tablice")
    }
}


# === Alternativna funkcija za parsiranje iz vašeg formata ===
function Parse-AlternativeFormat {
    param([string]$html)
    
    try {
        Write-Host "Debug: Pokušavam alternativno parsiranje..." -ForegroundColor Yellow
        
        # Na osnovu vašeg primjera, podaci se nalaze nakon odreðenih pattern-a
        $lines = $html -split "`r?`n"
        $dataLines = @()
        
        foreach ($line in $lines) {
            $cleanLine = $line.Trim()
            
            # Preskaèemo prazne linije i one koje sadrže HTML/CSS
            if ($cleanLine -eq '' -or 
                $cleanLine -match '^<' -or 
                $cleanLine -match 'body \{' -or
                $cleanLine -match '^\}' -or
                $cleanLine -match 'background-color' -or
                $cleanLine -match 'font-family') {
                continue
            }
            
            # Pokušaj identificirati linije s podacima korisnika
            # Na osnovu vašeg primjera: "1234567 kjèkjèk -saasfdasf Yes..."
            if ($cleanLine -match '^[a-zA-Z0-9\.\-_]+\s+.*\s+(Yes|No)\s+') {
                Write-Host "Debug: Pronašao data liniju: $cleanLine" -ForegroundColor Green
                
                # Pokušaj parsirati liniju na komponente
                $parts = @()
                $tokens = $cleanLine -split '\s+'
                
                if ($tokens.Count -ge 8) {
                    $parts += $tokens[0]  # Name
                    
                    # Full Name (može biti više rijeèi)
                    $nameEnd = 1
                    while ($nameEnd -lt $tokens.Count -and $tokens[$nameEnd] -notmatch '^(Yes|No)$') {
                        $nameEnd++
                    }
                    $fullName = ($tokens[1..($nameEnd-1)] -join ' ')
                    $parts += $fullName
                    
                    # Ostali podaci
                    $remainingTokens = $tokens[$nameEnd..($tokens.Count-1)]
                    $parts += $remainingTokens[0..5]  # Account Active, Last Password Set, etc.
                    
                    $dataLines += ,$parts
                }
            }
        }
        
        if ($dataLines.Count -gt 0) {
            # Dodaj header
            $headers = @('Name', 'Full Name', 'Account Active', 'Last Password Set', 'Last Logon', 'Account Expiry', 'Account Disabled', 'Password Required')
            $result = @($headers) + $dataLines
            return $result
        }
        
        return @()
    }
    catch {
        Write-Warning "Greška u alternativnom parsiranju: $_"
        return @()
    }
}

# === Potpuno nova i poboljšana funkcija Export-TxtSmartParser ===
function Export-TxtSmartParser {
    param (
        [string]$FilePath,
        $control
    )
    Export-TxtStructured -FilePath $FilePath -control $control
}




# FUNKCIJA ZA DOHVAÆANJE STVARNOG IMENA KORISNIKA
function Get-RealUserName {
    try {
        # Metoda 1: Pokušaj dohvatiti puno ime iz WMI lokalnog korisnika
        $currentUser = $env:USERNAME
        $localUser = Get-WmiObject -Class Win32_UserAccount -Filter "Name='$currentUser' AND LocalAccount='True'" -ErrorAction SilentlyContinue
        
        if ($localUser -and $localUser.FullName -and $localUser.FullName.Trim() -ne "") {
            return $localUser.FullName.Trim()
        }
        
        # Metoda 2: Pokušaj dohvatiti iz Active Directory (ako je domena)
        try {
            $adUser = Get-WmiObject -Class Win32_UserAccount -Filter "Name='$currentUser' AND LocalAccount='False'" -ErrorAction SilentlyContinue
            if ($adUser -and $adUser.FullName -and $adUser.FullName.Trim() -ne "") {
                return $adUser.FullName.Trim()
            }
        } catch { }
        
        # Metoda 3: Pokušaj kroz .NET Framework
        try {
            $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            if ($identity.Name -match '\\(.+)$') {
                $userName = $matches[1]
            } else {
                $userName = $identity.Name
            }
            
            # Pokušaj dohvatiti DisplayName iz registra
            $userSID = $identity.User.Value
            $regPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$userSID"
            if (Test-Path $regPath) {
                $profileInfo = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                if ($profileInfo.ProfileImagePath) {
                    $profilePath = $profileInfo.ProfileImagePath
                    # Pokušaj izvuæi ime iz putanje profila
                    if ($profilePath -match '\\([^\\]+)$') {
                        $extractedName = $matches[1]
                        if ($extractedName -ne $currentUser) {
                            return $extractedName
                        }
                    }
                }
            }
        } catch { }
        
        # Metoda 4: Pokušaj kroz net user komandu
        try {
            $netUserOutput = net user $currentUser 2>$null
            if ($LASTEXITCODE -eq 0) {
                foreach ($line in $netUserOutput) {
                    if ($line -match '^Full Name\s+(.+)$') {
                        $fullName = $matches[1].Trim()
                        if ($fullName -and $fullName -ne "" -and $fullName -ne $currentUser) {
                            return $fullName
                        }
                    }
                }
            }
        } catch { }
        
        # Metoda 5: Pokušaj kroz Environment varijable
        try {
            if ($env:USERDOMAIN -and $env:USERDOMAIN -ne $env:COMPUTERNAME) {
                # Domena korisnik - pokušaj ADSI
                $domain = $env:USERDOMAIN
                $searcher = [adsisearcher]"(samaccountname=$currentUser)"
                $user = $searcher.FindOne()
                if ($user -and $user.Properties.displayname) {
                    return $user.Properties.displayname[0]
                }
            }
        } catch { }
        
        # Ako ništa ne radi, vrati username
        return $currentUser
        
    } catch {
        # U sluèaju bilo kakve greške, vrati samo username
        return $env:USERNAME
    }
}



# === Poboljšana Excel Export funkcija s vizualnim formatiranjem ===
function Export-ToExcelWithFormatting {
    param(
        [string]$htmlContent,
        [string]$fileName
    )

    try {
        Write-Host "Pokušavam Excel export s formatiranjem..." -ForegroundColor Yellow
        
        # Parsiraj HTML tablicu s dodatnim informacijama o formatiranju
        $tableData = Parse-HTMLTableWithFormatting -html $htmlContent
        
        if ($tableData.Rows.Count -eq 0) {
            throw "Nema podataka za export u Excel"
        }

        # Kreiraj Excel aplikaciju
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Kreiraj workbook i worksheet
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Izvještaj"

        # Popuni podatke s formatiranjem
        $row = 1
        foreach ($dataRow in $tableData.Rows) {
            $col = 1
            foreach ($cellInfo in $dataRow) {
                $cell = $worksheet.Cells.Item($row, $col)
                $cell.Value = $cellInfo.Text
                
                # Primijeni formatiranje na æeliju
                Apply-CellFormatting -cell $cell -formatInfo $cellInfo
                
                $col++
            }
            $row++
        }

        # Specijalno formatiranje za header red
        if ($tableData.Rows.Count -gt 0) {
            $headerRange = $worksheet.Range("A1", [char](64 + $tableData.Rows[0].Count) + "1")
            $headerRange.Font.Bold = $true
            $headerRange.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(30, 58, 138)) # #1e3a8a
            $headerRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(147, 197, 253)) # #93c5fd
            $headerRange.Borders.LineStyle = 1
            $headerRange.Borders.Weight = 2
        }

        # Formatiranje alternativnih redova (kao u HTML-u)
        for ($i = 2; $i -le $tableData.Rows.Count; $i++) {
            $rowRange = $worksheet.Range("A$i", [char](64 + $tableData.Rows[0].Count) + "$i")
            
            if ($i % 2 -eq 0) {
                # Parni redovi - plava pozadina (#f0f9ff)
                $rowRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(240, 249, 255))
            } else {
                # Neparni redovi - bijela pozadina
                $rowRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
            }
            
            # Dodaj granice
            $rowRange.Borders.LineStyle = 1
            $rowRange.Borders.Weight = 1
            $rowRange.Borders.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(221, 221, 221))
        }

        # Formatiranje za neaktivne korisnike (crvena boja)
        for ($i = 1; $i -le $tableData.Rows.Count; $i++) {
            $rowData = $tableData.Rows[$i-1]
            $isInactive = $false
            
            # Provjeri sadrži li red oznake neaktivnosti
            foreach ($cell in $rowData) {
                if ($cell.Text -match "Neaktivan|Onemoguæen|False|disabled" -or $cell.IsInactive) {
                    $isInactive = $true
                    break
                }
            }
            
            if ($isInactive) {
                $rowRange = $worksheet.Range("A$i", [char](64 + $rowData.Count) + "$i")
                $rowRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 230, 230)) # #ffe6e6
                $rowRange.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
                $rowRange.Font.Bold = $true
            }
        }

        # Auto-fit kolone
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        
        # Zamrzni prvi red (header)
        $worksheet.Range("A2").Select() | Out-Null
        $excel.ActiveWindow.FreezePanes = $true

        # Dodaj footer informacije (kao u HTML-u)
        $lastRow = $tableData.Rows.Count + 3
        $worksheet.Cells.Item($lastRow, 1) = "Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
        $worksheet.Cells.Item($lastRow + 1, 1) = "Sustav: $env:COMPUTERNAME"
        $worksheet.Cells.Item($lastRow + 2, 1) = "Izrada: Ivica Rašan, 2025"
        
        # Formatiraj footer
        $footerRange = $worksheet.Range("A$lastRow", "A$($lastRow + 2)")
        $footerRange.Font.Size = 10
        $footerRange.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Gray)
        $footerRange.Font.Italic = $true

        # Spremi datoteku
        $workbook.SaveAs($fileName, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        
        Write-Host "Excel datoteka s formatiranjem uspješno kreirana: $fileName" -ForegroundColor Green
        
        # Zatvori Excel
        $workbook.Close($false)
        $excel.Quit()
        
        # Oslobodi COM objekte
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Error "Greška pri Excel exportu s formatiranjem: $($_.Exception.Message)"
        
        # Cleanup u sluèaju greške
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        
        throw
    }
}

# === Poboljšana funkcija za parsiranje s formatiranjem ===
function Parse-HTMLTableWithFormatting {
    param([string]$html)
    
    try {
        Write-Host "Parsiram HTML tablicu s formatiranjem..." -ForegroundColor Cyan
        
        $result = @{
            Rows = @()
        }
        
        # Ukloni CSS i JavaScript, ali zadrži class atribute
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', '' -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Pronaði tablicu
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Ukloni thead/tbody/tfoot wrappere ali zadrži tr elemente
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði sve redove s atributima
            $rowPattern = '(?s)<tr([^>]*)>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            foreach ($rowMatch in $rowMatches) {
                $rowAttributes = $rowMatch.Groups[1].Value
                $rowContent = $rowMatch.Groups[2].Value
                $cells = @()
                
                # Determiniraj je li red neaktivan
                $rowIsInactive = $rowAttributes -match 'class="[^"]*inactive[^"]*"'
                
                # Pronaði sve æelije s atributima
                $cellPattern = '(?s)<(t[hd])([^>]*?)>(.*?)</\1>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellTag = $cellMatch.Groups[1].Value
                    $cellAttributes = $cellMatch.Groups[2].Value
                    $cellContent = $cellMatch.Groups[3].Value
                    
                    # Parsiraj sadržaj æelije
                    $cleanText = Parse-CellContent -content $cellContent
                    
                    # Kreiraj objekt s informacijama o æeliji
                    $cellInfo = @{
                        Text = $cleanText
                        IsHeader = ($cellTag -eq "th")
                        IsInactive = ($rowIsInactive -or $cellAttributes -match 'class="[^"]*inactive[^"]*"')
                        HasIcon = ($cellContent -match '<img[^>]*>')
                        IsFirstColumn = ($cells.Count -eq 0)
                    }
                    
                    $cells += $cellInfo
                }
                
                if ($cells.Count -gt 0) {
                    $result.Rows += ,$cells
                }
            }
        }
        
        Write-Host "Parsirano $($result.Rows.Count) redova s formatiranjem" -ForegroundColor Green
        return $result
        
    } catch {
        Write-Warning "Greška pri parsiranju HTML tablice s formatiranjem: $_"
        return @{ Rows = @() }
    }
}



# === Funkcija za parsiranje sadržaja æelije ===
function Parse-CellContent {
    param([string]$content)
    
    # Zamijeni <br> s newline
    $text = $content -replace '<br[^>]*>', "`n"
    
    # Ukloni HTML tagove ali zadrži tekst
    $text = $text -replace '<[^>]+>', ''
    
    # Ukloni HTML entitete
    $text = $text -replace '&nbsp;', ' '
    $text = $text -replace '&amp;', '&'
    $text = $text -replace '&lt;', '<'
    $text = $text -replace '&gt;', '>'
    $text = $text -replace '&quot;', '"'
    $text = $text -replace '&#39;', "'"
    
    # Oèisti whitespace
    $text = $text -replace '\r?\n\s*\r?\n', "`n"
    $text = $text -replace '^\s+|\s+$', ''
    
    return $text
}

# === Funkcija za primjenu formatiranja na Excel æeliju ===
function Apply-CellFormatting {
    param(
        $cell,
        $formatInfo
    )
    
    try {
        # Osnovno formatiranje teksta
        $cell.Font.Name = "Segoe UI"
        $cell.Font.Size = 11
        
        # Formatiranje za header æelije
        if ($formatInfo.IsHeader) {
            $cell.Font.Bold = $true
            $cell.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(30, 58, 138))
            $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(147, 197, 253))
        }
        
        # Formatiranje za prvu kolonu (sticky column iz CSS-a)
        if ($formatInfo.IsFirstColumn -and -not $formatInfo.IsHeader) {
            $cell.Font.Bold = $true
            $cell.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(107, 114, 128))
            $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(243, 244, 246))
        }
        
        # Formatiranje za neaktivne æelije
        if ($formatInfo.IsInactive) {
            $cell.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
            $cell.Font.Bold = $true
            $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 230, 230))
        }
        
        # Dodaj granice
        $cell.Borders.LineStyle = 1
        $cell.Borders.Weight = 1
        $cell.Borders.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(221, 221, 221))
        
    } catch {
        Write-Warning "Greška pri formatiranju æelije: $_"
    }
}


# === Funkcija za parsiranje HTML-a za TXT ===
function Parse-HTMLTableForTXT {
    param([string]$html)
    
    try {
        $tableData = @()
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Pronaði tablicu
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Ukloni thead/tbody wrappere
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði sve redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði sve æelije (th ili td)
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Èisti HTML i formatiraj za tekst
                    $cleanCell = Clean-CellForTXT -content $cellContent
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        }
        
        Write-Host "  Parsirano $($tableData.Count) redova za TXT" -ForegroundColor Gray
        return $tableData
        
    } catch {
        Write-Warning "Greška pri parsiranju za TXT: $_"
        return @()
    }
}

# === Funkcija za èišæenje sadržaja æelije ===
function Clean-CellForTXT {
    param([string]$content)
    
    if ([string]::IsNullOrEmpty($content)) {
        return ""
    }
    
    # Zamijeni <br> s razmakom umjesto newline
    $text = $content -replace '<br[^>]*>', ' '
    
    # Ukloni sve HTML tagove
    $text = $text -replace '<[^>]+>', ''
    
    # Ukloni HTML entitete
    $text = $text -replace '&nbsp;', ' '
    $text = $text -replace '&amp;', '&'
    $text = $text -replace '&lt;', '<'
    $text = $text -replace '&gt;', '>'
    $text = $text -replace '&quot;', '"'
    $text = $text -replace '&#39;', "'"
    
    # Oèisti whitespace
    $text = $text -replace '\r?\n', ' '  # Zamijeni newline s razmakom
    $text = $text -replace '\s+', ' '    # Smanji multiple spaces
    $text = $text.Trim()
    
    # Ogranièi duljinu za èitljivost
    if ($text.Length -gt 50) {
        $text = $text.Substring(0, 47) + "..."
    }
    
    return $text
}

# === Funkcija za formatiranje podataka kao èiste tablice ===
function Format-DataAsCleanTable {
    param([array]$data)
    
    if ($data.Count -eq 0) {
        return @("Nema podataka")
    }
    
    try {
        # Dodaj redni broj kao prvu kolonu
        $numberedData = @()
        for ($i = 0; $i -lt $data.Count; $i++) {
            if ($i -eq 0) {
                # Header red
                $numberedData += ,(@("Br.") + $data[$i])
            } else {
                # Data redovi
                $numberedData += ,(@($i.ToString()) + $data[$i])
            }
        }
        
        # Izraèunaj širine kolona
        $colWidths = @()
        $maxCols = ($numberedData | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
        
        for ($col = 0; $col -lt $maxCols; $col++) {
            $maxWidth = 0
            foreach ($row in $numberedData) {
                if ($col -lt $row.Count) {
                    $width = $row[$col].ToString().Length
                    if ($width -gt $maxWidth) { $maxWidth = $width }
                }
            }
            # Ogranièi širinu na maksimalno 25 znakova za èitljivost
            $colWidths += [Math]::Min($maxWidth + 2, 25)
        }
        
        # Generiraj èistu tablicu
        $result = @()
        
        foreach ($rowIndex in 0..($numberedData.Count - 1)) {
            $row = $numberedData[$rowIndex]
            $line = ""
            
            for ($col = 0; $col -lt $row.Count; $col++) {
                $cell = if ($col -lt $row.Count) { $row[$col].ToString() } else { "" }
                
                # Skrati ako je predugaèak
                if ($cell.Length -gt $colWidths[$col] - 2) {
                    $cell = $cell.Substring(0, $colWidths[$col] - 5) + "..."
                }
                
                $line += $cell.PadRight($colWidths[$col])
            }
            
            $result += $line.TrimEnd()
            
            # Dodaj separator nakon header reda
            if ($rowIndex -eq 0) {
                $separator = ""
                for ($col = 0; $col -lt $colWidths.Count; $col++) {
                    $separator += "-" * $colWidths[$col]
                }
                $result += $separator
            }
        }
        
        return $result
        
    } catch {
        Write-Warning "Greška pri formatiranju tablice: $_"
        return @("Greška pri formatiranju podataka")
    }
}

# === Ekstraktiranje naslova iz HTML-a ===
function Extract-HTMLTitle {
    param([string]$html)
    
    # Pokušaj dohvatiti h1 naslov
    if ($html -match '<h1[^>]*>(.*?)</h1>') {
        $title = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        return $title.Trim()
    }
    
    # Ili title tag
    if ($html -match '<title[^>]*>(.*?)</title>') {
        $title = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        return $title.Trim()
    }
    
    return $null
}

# === Ekstraktiranje opisa iz HTML-a ===
function Extract-HTMLDescription {
    param([string]$html)
    
    # Pokušaj dohvatiti prvi paragraf
    if ($html -match '<p[^>]*>(.*?)</p>') {
        $desc = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        $desc = $desc.Trim()
        
        # Samo ako nije prazan i nije prelogièan
        if ($desc.Length -gt 10 -and $desc.Length -lt 200) {
            return $desc
        }
    }
    
    return $null
}



# === Export funkcije za druge formate ===
function Export-ToPDF {
    param(
        [string]$htmlContent,
        [string]$fileName
    )

    try {
        Write-Host "Kreiram PDF dokument..." -ForegroundColor Yellow
        
        # DEFINIRAJ LOKALNO
        $reportTitle = Get-ReportTitle -htmlContent $htmlContent
        $appVersion = "2025.6.3"  # Hardkodiraj verziju
        $currentDateTime = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
        $currentYear = Get-Date -Format 'yyyy'
        
        # Kreiraj optimizirani HTML
        $tableData = Parse-HTMLTableToArray -html $htmlContent
        
        $pdfHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        @page { size: A4 landscape; margin: 0.5cm; }
        body { font-family: Arial; font-size: 8pt; margin: 0; padding: 10px; }
        .header { text-align: center; margin-bottom: 20px; }
        h1 { font-size: 16pt; color: #2c3e50; margin-bottom: 5px; }
        .subtitle { font-size: 10pt; color: #7f8c8d; margin-bottom: 15px; }
        table { width: 100%; border-collapse: collapse; font-size: 7pt; }
        th, td { border: 1px solid #ddd; padding: 2px 4px; max-width: 100px; overflow: hidden; }
        th { background: #3498db; color: white; font-weight: bold; text-align: center; }
        tr:nth-child(even) { background: #f8f9fa; }
        .footer { margin-top: 15px; font-size: 7pt; text-align: center; color: #666; }
    </style>
</head>
<body>
    <div class='header'>
        <h1>$reportTitle</h1>
        <div class='subtitle'>Generirano: $currentDateTime | Sustav: $env:COMPUTERNAME</div>
    </div>
    <table>
"@

        foreach ($row in $tableData) {
            $pdfHtml += "<tr>"
            foreach ($cell in $row) {
                $tag = if ($tableData.IndexOf($row) -eq 0) { "th" } else { "td" }
                $shortCell = if ($cell.Length -gt 25) { $cell.Substring(0, 22) + "..." } else { $cell }
                $pdfHtml += "<$tag>$shortCell</$tag>"
            }
            $pdfHtml += "</tr>"
        }

        $pdfHtml += @"
    </table>
    <div class='footer'>
        <p><strong>Get-PCInfo Professional v$appVersion</strong> | © Ivica Rašan $currentYear</p>
    </div>
</body>
</html>
"@
        
        $tempHtml = "$env:TEMP\pdf_export.html"
        [System.IO.File]::WriteAllText($tempHtml, $pdfHtml, [System.Text.Encoding]::UTF8)

        # Word -> PDF
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($tempHtml)
        $doc.PageSetup.Orientation = 1  # Landscape
        $doc.SaveAs([ref]$fileName, [ref]17)  # PDF
        $doc.Close($false)
        $word.Quit()

        # Cleanup
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()

        Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
        Write-Host "PDF uspješno kreiran!" -ForegroundColor Green

    } catch {
        Write-Error "PDF export greška: $($_.Exception.Message)"
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
        throw
    }
}


function Get-ReportTitle {
    param([string]$htmlContent)
    
    try {
        # Pokušaj dohvatiti naslov iz trenutnog taba
        if ($tabControl -and $tabControl.SelectedTab) {
            $tabName = $tabControl.SelectedTab.Text
            
            # Mapiranje naziva tabova na naslove izvještaja
            $titleMap = @{
                "Korisnici"                 = "Lokalni korisnièki raèuni"
                "Grupe"                     = "Lokalne grupe i èlanstvo"
                "Logiranja i odjave"        = "Povijest prijava i odjava korisnika"
                "Neaktivni korisnici"       = "Neaktivni korisnièki raèuni"
                "Folder Permissions"        = "Prava pristupa po folderima"
                "Korištenje USB"            = "USB aktivnost s korisnicima"
                "Kreiranje KR"              = "Datumi kreiranja korisnièkih raèuna"
                "Obrisani KR"               = "Izbrisani korisnièki raèuni"
                "Print"                     = "Printeri i povijest printanja"
                "Administratori"            = "Lokalna Administrators grupa"
                "Serijski brojevi"          = "Serijski brojevi komponenti"
                "Instalirani programi"      = "Popis instaliranih programa"
                "Info"                      = "Get-PCInfo - Informacije o aplikaciji"
            }
            
            if ($titleMap.ContainsKey($tabName)) {
                return $titleMap[$tabName]
            }
        }
        
        # Fallback: pokušaj dohvatiti iz HTML sadržaja
        if ($htmlContent -match '<h1[^>]*>(.*?)</h1>') {
            $htmlTitle = $matches[1] -replace '<[^>]+>', ''
            return $htmlTitle.Trim()
        }
        
        if ($htmlContent -match '<title[^>]*>(.*?)</title>') {
            $htmlTitle = $matches[1] -replace '<[^>]+>', ''
            return $htmlTitle.Trim()
        }
        
        # Zadnji fallback
        return "Get-PCInfo Izvještaj"
        
    } catch {
        return "Get-PCInfo Izvještaj"
    }
}

# funkcija koja direktno èita iz WebBrowser kontrole
function Get-TableDataFromWebBrowser {
    param($webBrowserControl)
    
    try {
        Write-Host "=== DIREKTNO CITANJE IZ WEBBROWSER ===" -ForegroundColor Yellow
        
        # Pokušaj direktno pristupiti DOM-u
        $document = $webBrowserControl.Document
        if (-not $document) {
            Write-Host "Document nije dostupan" -ForegroundColor Red
            return @()
        }
        
        # Pronaði tablicu u DOM-u
        $tables = $document.GetElementsByTagName("table")
        if ($tables.Count -eq 0) {
            Write-Host "Nema tablica u DOM-u" -ForegroundColor Red
            return @()
        }
        
        Write-Host "Pronasao $($tables.Count) tablica" -ForegroundColor Green
        
        # Uzmi prvu tablicu
        $table = $tables[0]
        $rows = $table.GetElementsByTagName("tr")
        
        Write-Host "Tablica ima $($rows.Count) redova" -ForegroundColor Cyan
        
        $tableData = @()
        
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $row = $rows[$r]
            $cells = @()
            
            # Dohvati sve td i th elemente
            $tds = $row.GetElementsByTagName("td")
            $ths = $row.GetElementsByTagName("th")
            
            # Kombiniraj td i th
            $allCells = @()
            for ($i = 0; $i -lt $ths.Count; $i++) { $allCells += $ths[$i] }
            for ($i = 0; $i -lt $tds.Count; $i++) { $allCells += $tds[$i] }
            
            foreach ($cell in $allCells) {
                $cellText = $cell.InnerText
                if ($cellText) {
                    $cellText = $cellText.Trim()
                    $cells += $cellText
                }
            }
            
            # Dodaj red ako ima æelije
            if ($cells.Count -gt 0) {
                $tableData += ,$cells
                Write-Host "Red ${r}: $($cells.Count) celija - $($cells[0])" -ForegroundColor Green
            }
        }
        
        # KLJUÈNO: Ukloni footer tekst iz BILO KOJE æelije
        $cleanedTableData = @()
        foreach ($row in $tableData) {
            $cleanedRow = @()
            foreach ($cell in $row) {
                # Ako æelija sadrži footer tekst, preskoèi je
                if ($cell -notmatch 'Get-PCInfo.*Professional.*v\d{4}\.\d+.*©.*Ivica.*Rašan') {
                    $cleanedRow += $cell
                } else {
                    Write-Host "PRESKOÈI æeliju s footer tekstom: $($cell.Substring(0, [Math]::Min(30, $cell.Length)))..." -ForegroundColor Red
                }
            }
            
            # Dodaj red samo ako ima æelije nakon èišæenja
            if ($cleanedRow.Count -gt 0) {
                $cleanedTableData += ,$cleanedRow
            }
        }
        
        Write-Host "=== UKUPNO: $($cleanedTableData.Count) cistih redova ===" -ForegroundColor Green
        return $cleanedTableData
        
    } catch {
        Write-Host "Greska pri direktnom citanju: $_" -ForegroundColor Red
        return @()
    }
}





# Dodatna funkcija za èišæenje HTML-a prije parsiranja
function Clean-HTMLBeforeParsing {
    param([string]$html)
    
    Write-Host "Èišæenje HTML-a prije parsiranja..." -ForegroundColor Cyan
    
    # Ukloni sve elemente koji sadrže footer tekst
    $cleanedHtml = $html -replace '(?s)<[^>]*>.*?Get-PCInfo.*?Professional.*?v\d{4}\.\d+.*?©.*?Ivica.*?Rašan.*?</[^>]*>', ''
    
    # Ukloni i tekstualne instance
    $cleanedHtml = $cleanedHtml -replace 'Get-PCInfo.*?Professional.*?v\d{4}\.\d+.*?©.*?Ivica.*?Rašan.*?\d{4}', ''
    
    # Ukloni prazne æelije koje su ostale
    $cleanedHtml = $cleanedHtml -replace '(?s)<td[^>]*>\s*</td>', ''
    $cleanedHtml = $cleanedHtml -replace '(?s)<th[^>]*>\s*</th>', ''
    
    Write-Host "HTML oèišæen" -ForegroundColor Green
    return $cleanedHtml
}


# funkcija koja æe restrukturirati tablicu ako je potrebno
function Fix-TableStructure {
    param([array]$tableData)
    
    if ($tableData.Count -eq 0) {
        return $tableData
    }
    
    Write-Host "=== RESTRUKTURIRANJE TABLICE ===" -ForegroundColor Yellow
    
    # Pronaði najèešæi broj kolona
    $columnCounts = $tableData | ForEach-Object { $_.Count } | Group-Object | Sort-Object Count -Descending
    $expectedColumns = $columnCounts[0].Name -as [int]
    
    Write-Host "Oèekivani broj kolona: $expectedColumns" -ForegroundColor Cyan
    
    $fixedData = @()
    for ($i = 0; $i -lt $tableData.Count; $i++) {
        $row = $tableData[$i]
        
        # Ako red ima premalo kolona, dodaj prazne
        while ($row.Count -lt $expectedColumns) {
            $row += ""
        }
        
        # Ako red ima previše kolona, skrati
        if ($row.Count -gt $expectedColumns) {
            $row = $row[0..($expectedColumns-1)]
        }
        
        # Provjeri da prvi red sadrži header podatke
        if ($i -eq 0 -and $row[0] -eq "") {
            $row[0] = "#"  # Dodaj broj stupca ako je prazan
        }
        
        $fixedData += ,$row
        Write-Host "Red $i`: $($row.Count) kolona - '$($row[0])'" -ForegroundColor Gray
    }
    
    Write-Host "=== TABLIÈNA STRUKTURA POPRAVLJENA ===" -ForegroundColor Green
    return $fixedData
}


# Backup funkcija za èišæenje tablice nakon parsiranja
function Clean-TableDataPostProcess {
    param([array]$tableData)
    
    if ($tableData.Count -eq 0) {
        return $tableData
    }
    
    # Oèisti prvu æeliju prvog reda ako sadrži app info
    if ($tableData[0].Count -gt 0) {
        $firstCell = $tableData[0][0]
        if ($firstCell -match 'Get-PCInfo|Professional|v\d{4}\.\d+|©.*Ivica.*Rašan') {
            Write-Host "POST-PROCESS: Èišæenje prve æelije '$firstCell'" -ForegroundColor Red
            $tableData[0][0] = "#"  # Zamijeni s jednostavnim znakom
        }
    }
    
    # Ukloni sve redove koji u cijelosti sadrže app info
    $cleanedData = @()
    foreach ($row in $tableData) {
        $rowText = ($row -join ' ')
        if ($rowText -notmatch 'Get-PCInfo.*Professional.*v\d{4}\.\d+.*©.*Ivica.*Rašan') {
            $cleanedData += ,$row
        } else {
            Write-Host "POST-PROCESS: Uklanjam red '$($rowText.Substring(0, [Math]::Min(50, $rowText.Length)))...'" -ForegroundColor Red
        }
    }
    
    return $cleanedData
}



# === EXPORT-TO-WORD FUNKCIJA ===
function Export-ToWord {
    param(
        [string]$htmlContent,
        [string]$fileName
    )

    try {
        Write-Host "=== WORD EXPORT ===" -ForegroundColor Yellow
        
        $reportTitle = Get-ReportTitle -htmlContent $htmlContent
        $appVersion = "2025.6.3"
        $currentDateTime = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
        $currentYear = Get-Date -Format 'yyyy'
        
        # Pokušaj direktno èitanje iz WebBrowser kontrole
        $tableData = @()
        
        # Dohvati aktivnu WebBrowser kontrolu
        $activeTab = $tabControl.SelectedTab
        if ($activeTab) {
            $webBrowser = $activeTab.Controls | Where-Object { $_ -is [System.Windows.Forms.WebBrowser] } | Select-Object -First 1
            if ($webBrowser) {
                Write-Host "WebBrowser kontrola pronaðena" -ForegroundColor Green
                $tableData = Get-TableDataFromWebBrowser -webBrowserControl $webBrowser
            }
        }
        
        # Fallback na HTML parsing ako direktno èitanje ne radi
        if ($tableData.Count -eq 0) {
            Write-Host "Direktno èitanje neuspješno, koristim HTML parsing..." -ForegroundColor Yellow
            $tableData = Parse-HTMLTableToArray -html $htmlContent
        }
        
        if ($tableData.Count -eq 0) {
            Write-Host "Nema podataka za export!" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("Nema podataka za Word export!", "Greška", "OK", "Warning")
            return
        }
        
        Write-Host "Konaèni broj redova: $($tableData.Count)" -ForegroundColor Green
        
        # PROVJERI I UKLONI FOOTER TEKST IZ PRVE ÆELIJE
        Write-Host "DEBUG: Provjeravam prvu æeliju..." -ForegroundColor Yellow
        if ($tableData.Count -gt 0 -and $tableData[0].Count -gt 0) {
            Write-Host "DEBUG: Prva æelija sadrži: '$($tableData[0][0])'" -ForegroundColor Cyan
            
            # Ako prva æelija sadrži footer tekst, zamijeni je s "#"
            if ($tableData[0][0] -match 'Get-PCInfo|Professional|v\d{4}\.\d+|©.*Ivica.*Rašan') {
                Write-Host "ZAMJENA: Uklanjam footer iz prve æelije" -ForegroundColor Red
                $tableData[0][0] = "#"
            }
        }

        # PROVJERI SVE OSTALE ÆELIJE U TABLICI
        for ($r = 0; $r -lt $tableData.Count; $r++) {
            for ($c = 0; $c -lt $tableData[$r].Count; $c++) {
                if ($tableData[$r][$c] -match 'Get-PCInfo.*Professional.*v\d{4}\.\d+.*©.*Ivica.*Rašan') {
                    Write-Host "ZAMJENA: Uklanjam footer iz æelije [$r,$c]" -ForegroundColor Red
                    $tableData[$r][$c] = ""  # Zamijeni s praznim tekstom
                }
            }
        }
        
        # Kreiraj Word dokument
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Add()
        
        # Landscape setup
        $doc.PageSetup.Orientation = 1
        $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(0.8)
        $doc.PageSetup.RightMargin = $word.CentimetersToPoints(0.8)
        $doc.PageSetup.TopMargin = $word.CentimetersToPoints(1)
        $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1)
        
        # Naslov
        $word.Selection.Font.Name = "Calibri"
        $word.Selection.Font.Size = 16
        $word.Selection.Font.Bold = $true
        $word.Selection.Font.Color = 0  # CRNA BOJA
        $word.Selection.ParagraphFormat.Alignment = 1
        $word.Selection.TypeText($reportTitle)
        $word.Selection.TypeParagraph()
        
        # Podnaslov
        $word.Selection.Font.Size = 10
        $word.Selection.Font.Bold = $false
        $word.Selection.Font.Color = 8421504  # Siva boja
        $word.Selection.TypeText("Generirano: $currentDateTime | Sustav: $env:COMPUTERNAME | Korisnik: $env:USERNAME")
        $word.Selection.TypeParagraph()
        $word.Selection.TypeParagraph()
        
        $word.Selection.ParagraphFormat.Alignment = 0
        
        # Kreiraj tablicu
        $rowCount = $tableData.Count
        $colCount = $tableData[0].Count
        
        Write-Host "Kreiram Word tablicu: $rowCount x $colCount" -ForegroundColor Cyan
        
        $table = $doc.Tables.Add($word.Selection.Range, $rowCount, $colCount)
        $table.Style = "Table Grid"
        
        # Popuni podatke
        for ($r = 0; $r -lt $rowCount; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) {
                $cell = $table.Cell($r + 1, $c + 1)
                $cellValue = if ($c -lt $tableData[$r].Count) { $tableData[$r][$c] } else { "" }
                
                # DODATNA PROVJERA - ako æelija još uvijek sadrži footer, ukloni ga
                if ($cellValue -match 'Get-PCInfo.*Professional.*v\d{4}\.\d+.*©.*Ivica.*Rašan') {
                    $cellValue = ""
                    Write-Host "Uklonjen footer iz Word æelije [$r,$c]" -ForegroundColor Red
                }
                
                $cell.Range.Text = $cellValue
                $cell.Range.Font.Size = 8
                $cell.Range.Font.Color = 0  # CRNA BOJA
                
                if ($r -eq 0) {  # Header
                    $cell.Range.Font.Bold = $true
                    $cell.Range.Shading.BackgroundPatternColor = -603914241
                    $cell.Range.Font.Color = 0  # Crna boja za header
                }
            }
        }
        
        $table.AutoFitBehavior(1)
        
        if ($table.PreferredWidth -gt 750) {
            for ($r = 1; $r -le $rowCount; $r++) {
                for ($c = 1; $c -le $colCount; $c++) {
                    $table.Cell($r, $c).Range.Font.Size = 7
                    $table.Cell($r, $c).Range.Font.Color = 0  # Osiguraj crnu boju
                }
            }
        }
        
        # DODAJ FOOTER TEKST NA DNO DOKUMENTA
        # Postavi kursor na kraj dokumenta (ispod tablice)
        $word.Selection.EndKey(6)  # wdStory - na kraj dokumenta
        $word.Selection.TypeParagraph()
        $word.Selection.TypeParagraph()
        
        # Formatiranje footer teksta
        $word.Selection.Font.Name = "Calibri"
        $word.Selection.Font.Size = 9
        $word.Selection.Font.Bold = $false
        $word.Selection.Font.Color = 8421504  # Siva boja
        $word.Selection.ParagraphFormat.Alignment = 1  # Centrirano
        
        # Umetni footer tekst
        $footerText = "Get-PCInfo Professional v$appVersion | © Ivica Rašan 2025"
        $word.Selection.TypeText($footerText)
        
        $doc.SaveAs([ref]$fileName, [ref]16)
        $doc.Close($false)
        $word.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        
        Write-Host "Word dokument uspješno kreiran!" -ForegroundColor Green
        
    } catch {
        Write-Error "Word export greška: $($_.Exception.Message)"
        if ($doc) { $doc.Close($false) }
        if ($word) { $word.Quit() }
        throw
    }
}

# === KREIRAJ ÈIST HTML ZA WORD ===
function Create-WordCompatibleHTML {
    param([string]$originalHtml)
    
    # Izvuci samo podatke iz tablice
    $tableData = Parse-HTMLTableToArray -html $originalHtml
    
    if ($tableData.Count -eq 0) {
        return $originalHtml
    }
    
    # Kreiraj èist HTML bez problematiènih stilova
    $cleanHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: Arial, sans-serif; 
            color: black; 
            background-color: white; 
            margin: 20px; 
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            border: 1px solid #000; 
        }
        th, td { 
            border: 1px solid #000; 
            padding: 8px; 
            text-align: left; 
            color: black; 
        }
        th { 
            background-color: #f0f0f0; 
            font-weight: bold; 
            color: black; 
        }
    </style>
</head>
<body>
    <h1 style="color: black;">$(Get-ReportTitle -htmlContent $originalHtml)</h1>
    <table>
"@
    
    # Dodaj tablicu
    for ($i = 0; $i -lt $tableData.Count; $i++) {
        $tag = if ($i -eq 0) { "th" } else { "td" }
        $cleanHtml += "<tr>"
        foreach ($cell in $tableData[$i]) {
            $cleanHtml += "<$tag>$cell</$tag>"
        }
        $cleanHtml += "</tr>"
    }
    
    $cleanHtml += @"
    </table>
    <p style="color: black; margin-top: 20px;">
        Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')<br>
        Sustav: $env:COMPUTERNAME
    </p>
</body>
</html>
"@
    
    return $cleanHtml
}



# Test funkcija za debug
function Test-TableParsing {
    try {
        Write-Host "=== TEST PARSIRANJA TABLICE ===" -ForegroundColor Yellow
        
        $activeTab = $tabControl.SelectedTab
        if ($activeTab) {
            Write-Host "Aktivni tab: $($activeTab.Text)" -ForegroundColor Green
            
            $webBrowser = $activeTab.Controls | Where-Object { $_ -is [System.Windows.Forms.WebBrowser] } | Select-Object -First 1
            if ($webBrowser) {
                Write-Host "WebBrowser pronaden" -ForegroundColor Green
                
                $tableData = Get-TableDataFromWebBrowser -webBrowserControl $webBrowser
                
                if ($tableData.Count -gt 0) {
                    Write-Host "Ukupno redova: $($tableData.Count)" -ForegroundColor Green
                    for ($i = 0; $i -lt [Math]::Min(3, $tableData.Count); $i++) {
                        Write-Host "Red ${i}: $($tableData[$i] -join ' | ')" -ForegroundColor Cyan
                    }
                } else {
                    Write-Host "Nema podataka" -ForegroundColor Red
                }
            } else {
                Write-Host "WebBrowser nije pronaden" -ForegroundColor Red
            }
        } else {
            Write-Host "Nema aktivnog taba" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "Test greska: $_" -ForegroundColor Red
    }
}




# === Poboljšana Excel Export funkcija ===
function Export-ToExcelImproved {
    param(
        [string]$htmlContent,
        [string]$fileName
    )

    try {
        Write-Host "Pokušavam Excel export..." -ForegroundColor Yellow
        
        # Parsiraj HTML tablicu u strukturirane podatke
        $tableData = Parse-HTMLTableToArray -html $htmlContent
        
        if ($tableData.Count -eq 0) {
            throw "Nema podataka za export u Excel"
        }

        # Kreiraj Excel aplikaciju
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Kreiraj workbook i worksheet
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Izvještaj"

        # Popuni podatke
        $row = 1
        foreach ($dataRow in $tableData) {
            $col = 1
            foreach ($cell in $dataRow) {
                $worksheet.Cells.Item($row, $col) = $cell
                $col++
            }
            $row++
        }

        # Formatiraj header red (prvi red)
        if ($tableData.Count -gt 0) {
            $headerRange = $worksheet.Range("A1", [char](64 + $tableData[0].Count) + "1")
            $headerRange.Font.Bold = $true
            $headerRange.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightBlue)
        }

        # Auto-fit kolone
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null

        # Spremi datoteku
        $workbook.SaveAs($fileName, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        
        Write-Host "Excel datoteka uspješno kreirana: $fileName" -ForegroundColor Green
        
        # Zatvori Excel
        $workbook.Close($false)
        $excel.Quit()
        
        # Oslobodi COM objekte
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Error "Greška pri Excel exportu: $($_.Exception.Message)"
        
        # Cleanup u sluèaju greške
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        
        throw
    }
}



# === Funkcija za parsiranje HTML tablice u array ===
function Parse-HTMLTableToArray {
    param([string]$html)
    
    try {
        Write-Host "=== NOVA Parse funkcija - ignoriram footer ===" -ForegroundColor Yellow
        
        $tableData = @()
        
        # Ukloni sve što nije tablica
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # KLJUÈNO: Pronaði SAMO tablicu izmeðu <table> i </table> tagova
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            Write-Host "? Tablica ekstraktirana" -ForegroundColor Green
            
            # Ukloni wrappere
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            Write-Host "Broj redova u tablici: $($rowMatches.Count)" -ForegroundColor Cyan
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                # KLJUÈNA PROVJERA: Preskoèi redove koji imaju manje od 3 æelije (to su footer redovi)
                if ($cellMatches.Count -lt 3) {
                    Write-Host "?? Preskaèem red s $($cellMatches.Count) æelija (vjerojatno footer)" -ForegroundColor Yellow
                    continue
                }
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Oèisti sadržaj
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\r?\n', ' '
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    $cells += $cleanCell
                }
                
                # DODATNA PROVJERA: Preskoèi redove koji sadrže aplikaciju info
                $rowText = $cells -join ' '
                if ($rowText -match 'Get-PCInfo|v\d{4}\.\d+\.\d+|Ivica Rašan|Professional') {
                    Write-Host "?? Preskoæujem footer red: $($rowText.Substring(0, [Math]::Min(50, $rowText.Length)))..." -ForegroundColor Red
                    continue
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                    Write-Host "? Dodao red s $($cells.Count) æelija" -ForegroundColor Green
                }
            }
        } else {
            Write-Host "? Tablica nije pronaðena u HTML-u" -ForegroundColor Red
        }
        
        Write-Host "=== FINALNO: $($tableData.Count) redova parsirano ===" -ForegroundColor Green
        return $tableData
        
    } catch {
        Write-Warning "Greška pri parsiranju: $_"
        return @()
    }
}

# === CSV fallback funkcija ===
function Export-ToCSVFallback {
    param(
        [string]$htmlContent,
        [string]$fileName
    )
    
    try {
        $tableData = Parse-HTMLTableToArray -html $htmlContent
        
        if ($tableData.Count -eq 0) {
            throw "Nema podataka za CSV export"
        }
        
        # Kreiraj CSV sadržaj
        $csvContent = @()
        foreach ($row in $tableData) {
            # Escape quotes i dodaj quotes oko polja koja sadrže zareze
            $escapedRow = $row | ForEach-Object {
                $field = $_ -replace '"', '""'  # Escape quotes
                if ($field -match '[,"]') {
                    "`"$field`""  # Wrap in quotes if contains comma or quote
                } else {
                    $field
                }
            }
            $csvContent += $escapedRow -join ','
        }
        
        # Promijeni ekstenziju u .csv
        $csvFileName = $fileName -replace '\.xlsx?$', '.csv'
        
        # Spremi CSV datoteku
        $csvContent | Out-File -FilePath $csvFileName -Encoding UTF8 -Force
        
        [System.Windows.Forms.MessageBox]::Show(
            "Excel nije dostupan. Podaci su izveženi u CSV format:`n$csvFileName`n`nMožete otvoriti datoteku u Excelu.", 
            "CSV Export", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
    } catch {
        throw "CSV fallback neuspješan: $_"
    }
}



function Convert-HTMLTableToCSV {
    param([string]$htmlContent)
    
    try {
        $lines = $htmlContent -split "`n"
        $csvLines = @()
        $inTable = $false
        
        foreach ($line in $lines) {
            if ($line -match '<table') {
                $inTable = $true
                continue
            }
            if ($line -match '</table>') {
                $inTable = $false
                continue
            }
            
            if ($inTable -and ($line -match '<tr' -or $line -match '<td' -or $line -match '<th')) {
                $cleanLine = $line -replace '<[^>]+>', '|' -replace '\s+', ' ' -replace '\|+', '|' -replace '^\|', '' -replace '\|$', ''
                if ($cleanLine.Trim() -ne '') {
                    $csvLines += $cleanLine.Trim() -replace '\|', ','
                }
            }
        }
        
        return $csvLines -join "`n"
    }
    catch {
        return $null
    }
}




# JEDNOSTAVNIJA FUNKCIJA ZA PARSIRANJE PRINTER IZVJEŠTAJA
function Parse-PrintersFromHTML {
    param([string]$html)
    
    try {
        Write-Host "DEBUG: Starting printer parsing..." -ForegroundColor Yellow
        
        $reportData = @{
            InstalledPrinters = @()
            PrintHistory = @()
            PrinterStatus = @()
        }
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        Write-Host "DEBUG: Looking for h3 sections..." -ForegroundColor Cyan
        
        # Pronaði h3 naslove i sadržaj koji slijedi
        $h3Pattern = '(?s)<h3[^>]*>(.*?)</h3>'
        $h3Matches = [regex]::Matches($cleanHtml, $h3Pattern)
        
        Write-Host "DEBUG: Found $($h3Matches.Count) h3 sections" -ForegroundColor Cyan
        
        if ($h3Matches.Count -eq 0) {
            Write-Host "DEBUG: No h3 sections, looking for all tables..." -ForegroundColor Yellow
            
            # Jednostavno uzmi sve tablice
            $tablePattern = '(?s)<table[^>]*>(.*?)</table>'
            $tableMatches = [regex]::Matches($cleanHtml, $tablePattern)
            
            Write-Host "DEBUG: Found $($tableMatches.Count) tables" -ForegroundColor Cyan
            
            if ($tableMatches.Count -gt 0) {
                # Prva tablica = pisaèi
                $firstTable = "<table>" + $tableMatches[0].Groups[1].Value + "</table>"
                $printerData = Extract-PrinterNames -tableHtml $firstTable
                $reportData.InstalledPrinters = $printerData
                
                # Druga tablica = povijest (ako postoji)
                if ($tableMatches.Count -gt 1) {
                    $secondTable = "<table>" + $tableMatches[1].Groups[1].Value + "</table>"
                    $historyData = Extract-PrintHistory -tableHtml $secondTable
                    $reportData.PrintHistory = $historyData
                }
                
                # Treæa tablica = status (ako postoji)
                if ($tableMatches.Count -gt 2) {
                    $thirdTable = "<table>" + $tableMatches[2].Groups[1].Value + "</table>"
                    $statusData = Extract-PrintStatus -tableHtml $thirdTable
                    $reportData.PrinterStatus = $statusData
                }
            }
        } else {
            # Obradi h3 sekcije
            for ($i = 0; $i -lt $h3Matches.Count; $i++) {
                $h3Match = $h3Matches[$i]
                $sectionTitle = $h3Match.Groups[1].Value -replace '<[^>]+>', ''
                $sectionTitle = $sectionTitle.Trim()
                
                Write-Host "DEBUG: Processing section: '$sectionTitle'" -ForegroundColor Cyan
                
                # Pronaði sadržaj sekcije
                $startPos = $h3Match.Index + $h3Match.Length
                $endPos = if ($i -lt $h3Matches.Count - 1) { 
                    $h3Matches[$i + 1].Index 
                } else { 
                    $cleanHtml.Length 
                }
                
                $sectionContent = $cleanHtml.Substring($startPos, $endPos - $startPos)
                
                # Kategoriziraj sekciju
                if ($sectionTitle -match "instaliran|printer|pisaè") {
                    Write-Host "DEBUG: Processing as Installed Printers" -ForegroundColor Green
                    if ($sectionContent -match '(?s)<table[^>]*>(.*?)</table>') {
                        $tableHtml = "<table>" + $matches[1] + "</table>"
                        $reportData.InstalledPrinters = Extract-PrinterNames -tableHtml $tableHtml
                    } else {
                        $reportData.InstalledPrinters = @(@{ Message = "Nema tablice za pisaèe" })
                    }
                }
                elseif ($sectionTitle -match "povijest|history|ispis") {
                    Write-Host "DEBUG: Processing as Print History" -ForegroundColor Green
                    if ($sectionContent -match '(?s)<table[^>]*>(.*?)</table>') {
                        $tableHtml = "<table>" + $matches[1] + "</table>"
                        $reportData.PrintHistory = Extract-PrintHistory -tableHtml $tableHtml
                    } else {
                        # Provjeri ima li tekst o tome da nema zapisa
                        $cleanText = $sectionContent -replace '<[^>]+>', ' ' -replace '\s+', ' '
                        $reportData.PrintHistory = @(@{ Message = $cleanText.Trim() })
                    }
                }
                elseif ($sectionTitle -match "status|red") {
                    Write-Host "DEBUG: Processing as Printer Status" -ForegroundColor Green
                    if ($sectionContent -match '(?s)<table[^>]*>(.*?)</table>') {
                        $tableHtml = "<table>" + $matches[1] + "</table>"
                        $reportData.PrinterStatus = Extract-PrintStatus -tableHtml $tableHtml
                    } else {
                        $cleanText = $sectionContent -replace '<[^>]+>', ' ' -replace '\s+', ' '
                        $reportData.PrinterStatus = @(@{ Message = $cleanText.Trim() })
                    }
                }
            }
        }
        
        Write-Host "DEBUG: Final result - Printers: $($reportData.InstalledPrinters.Count), History: $($reportData.PrintHistory.Count), Status: $($reportData.PrinterStatus.Count)" -ForegroundColor Green
        return $reportData
        
    } catch {
        Write-Warning "DEBUG: Error in Parse-PrintersFromHTML: $_"
        return @{
            InstalledPrinters = @(@{ Message = "Greška pri parsiranju pisaèa" })
            PrintHistory = @(@{ Message = "Greška pri parsiranju povijesti" })
            PrinterStatus = @(@{ Message = "Greška pri parsiranju statusa" })
        }
    }
}

# FUNKCIJA ZA EKSTRAKTIRANJE IMENA PISAÈA
function Extract-PrinterNames {
    param([string]$tableHtml)
    
    try {
        $printers = @()
        
        if ($tableHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue  # Preskoèi header
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                
                # Pronaði prvu æeliju (naziv pisaèa)
                if ($rowContent -match '<t[hd][^>]*>(.*?)</t[hd]>') {
                    $printerName = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                    $printerName = $printerName.Trim()
                    
                    if ($printerName -ne "") {
                        $printers += @{ Name = $printerName }
                    }
                }
            }
        }
        
        return $printers
        
    } catch {
        Write-Warning "Error extracting printer names: $_"
        return @(@{ Message = "Greška pri ekstraktiranju imena pisaèa" })
    }
}

# FUNKCIJA ZA EKSTRAKTIRANJE POVIJESTI PRINTANJA
function Extract-PrintHistory {
    param([string]$tableHtml)
    
    try {
        $history = @()
        
        if ($tableHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            if ($rowMatches.Count -le 1) {
                return @(@{ Message = "Nema zapisa o printanju" })
            }
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                    $cells += $cellContent.Trim()
                }
                
                if ($cells.Count -gt 0) {
                    $historyItem = @{}
                    for ($i = 0; $i -lt $cells.Count; $i++) {
                        $historyItem["Kolona$($i+1)"] = $cells[$i]
                    }
                    $history += $historyItem
                }
            }
        }
        
        return $history
        
    } catch {
        Write-Warning "Error extracting print history: $_"
        return @(@{ Message = "Greška pri ekstraktiranju povijesti" })
    }
}


# FUNKCIJA ZA EKSTRAKTIRANJE STATUSA PISAÈA
function Extract-PrintStatus {
    param([string]$tableHtml)
    
    try {
        $status = @()
        
        if ($tableHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            if ($rowMatches.Count -le 1) {
                return @(@{ Message = "Nema status informacija" })
            }
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', ' '
                    $cells += $cellContent.Trim()
                }
                
                if ($cells.Count -gt 0) {
                    $statusItem = @{}
                    for ($i = 0; $i -lt $cells.Count; $i++) {
                        $statusItem["Polje$($i+1)"] = $cells[$i]
                    }
                    $status += $statusItem
                }
            }
        }
        
        return $status
        
    } catch {
        Write-Warning "Error extracting print status: $_"
        return @(@{ Message = "Greška pri ekstraktiranju statusa" })
    }
}


# JEDNOSTAVNA FUNKCIJA ZA PARSIRANJE TABLICA
function Parse-SimpleTable {
    param([string]$content)
    
    try {
        $data = @()
        
        # Pronaði tablicu
        if ($content -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Pronaði redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            $headers = @()
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    $cells += $cleanCell
                }
                
                if ($isFirstRow) {
                    $headers = $cells
                    $isFirstRow = $false
                    continue
                }
                
                if ($cells.Count -gt 0) {
                    $rowData = @{}
                    for ($i = 0; $i -lt $cells.Count; $i++) {
                        $key = if ($i -lt $headers.Count -and $headers[$i] -ne "") { 
                            $headers[$i] 
                        } else { 
                            "Kolona$($i+1)" 
                        }
                        $rowData[$key] = $cells[$i]
                    }
                    $data += $rowData
                }
            }
        }
        
        return $data
        
    } catch {
        Write-Warning "Error parsing simple table: $_"
        return @()
    }
}



# === AŽURIRANA EXPORT-ACTIVETAB FUNKCIJA ===
function Export-ActiveTab {
    param ([string]$format)

    $tab = $tabControl.SelectedTab
    $control = $tab.Controls | Where-Object { $_ -is [System.Windows.Forms.WebBrowser] -or $_ -is [System.Windows.Forms.RichTextBox] }

    if (-not $control) {
        [System.Windows.Forms.MessageBox]::Show("Nema podataka za izvoz.", "Upozorenje", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $html = ""
    if ($control -is [System.Windows.Forms.WebBrowser]) {
        $html = $control.DocumentText
    } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
        $html = $control.Text
    }

    $path = New-Object System.Windows.Forms.SaveFileDialog
    $path.Filter = switch ($format.ToLower()) {
        "txt"  { "TXT files (*.txt)|*.txt" }
        "html" { "HTML files (*.html)|*.html" }
        "doc"  { "Word files (*.docx)|*.docx" }
        "xls"  { "Excel files (*.xlsx)|*.xlsx" }
        "csv"  { "CSV files (*.csv)|*.csv" }
        "pdf"  { "PDF files (*.pdf)|*.pdf" }
        default { "TXT files (*.txt)|*.txt" }
    }
    
    $path.FileName = "$($tab.Text -replace '[^\w\-_\.]', '_')_Export"

    if ($path.ShowDialog() -eq "OK") {
        try {
            switch ($format.ToLower()) {
                "txt" {
                    Export-TxtStructured -FilePath $path.FileName -control $control
                }
                "xls" {
                    Export-ToExcelWithFallback -htmlContent $html -fileName $path.FileName
                }
                "csv" {
                    Export-ToCSV -htmlContent $html -fileName $path.FileName
                }
                "html" {
                    [System.IO.File]::WriteAllText($path.FileName, $html, [System.Text.Encoding]::UTF8)
                }
                "doc" {
                    Export-ToWord -htmlContent $html -fileName $path.FileName
                }
                "pdf" {
                    Export-ToPDF -htmlContent $html -fileName $path.FileName
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show("Izvoz zavrsen: $($path.FileName)", "Uspjeh", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Greska pri izvozu: $($_.Exception.Message)", "Greska", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}



# Umjesto postojeæeg dijela s gumbovima:
$panelButtons.Controls.AddRange(@(
    $btnExportTxt,
    $btnExportCSV,        # <-- NOVI
    $btnExportDoc,
    $btnExportExcel,
    $btnExportPDF,
    $btnExportHTML,
    $btnPrint
))



# Test funkcija za debug
function Test-HTMLParsing {
    $tab = $tabControl.SelectedTab
    $control = $tab.Controls[0]
    $html = $control.DocumentText
    
    Write-Host "HTML duljina: $($html.Length)"
    Write-Host "Prvo 500 znakova:"
    Write-Host $html.Substring(0, [Math]::Min(500, $html.Length))
    
    $parsed = Parse-HTMLTableSafe -html $html
    Write-Host "Parsirano redova: $($parsed.Count)"
}

# Pozovite prije Excel exporta

function Test-ExcelEnvironment {
    try {
        $excel = New-Object -ComObject Excel.Application
        Write-Host "Excel Version: $($excel.Version)"
        
        $wb = $excel.Workbooks.Add()
        Write-Host "Workbook Created: $($wb -ne $null)"
        
        $ws = $wb.Worksheets
        Write-Host "Worksheets Collection: $($ws -ne $null)"
        Write-Host "Worksheets Count: $($ws.Count)"
        
        $sheet = $ws.Item(1)
        Write-Host "First Sheet: $($sheet -ne $null)"
        
        $wb.Close($false)
        $excel.Quit()
        Write-Host "Test završen uspješno"
        
    } catch {
        Write-Host "Test neuspješan: $($_.Exception.Message)"
    }
}



function Format-GroupsDataEnhanced {
    param([array]$groupsData)
    
    $output = @()
    $output += "GRUPNE INFORMACIJE:"
    $output += "-" * 25
    $output += ""
    
    if ($groupsData.Count -eq 0) {
        $output += "Nema pronadenih grupa."
        return $output
    }
    
    $groupNumber = 1
    $totalMembers = 0
    $groupsWithMembers = 0
    
    foreach ($group in $groupsData) {
        $output += "[$groupNumber] $($group.Name)"
        
        if (-not $group.HasMembers -or $group.Members.Count -eq 0) {
            $output += "   > Nema clanova"
        } else {
            $groupsWithMembers++
            $totalMembers += $group.Members.Count
            
            $output += "   > Clanova: $($group.Members.Count)"
            $memberNumber = 1
            
            foreach ($member in $group.Members) {
                $output += "   > [$memberNumber] $($member.Username)"
                if ($member.FullName -and $member.FullName -ne "") {
                    $output += "      - Ime: $($member.FullName)"
                }
                $memberNumber++
            }
        }
        
        $output += ""
        $groupNumber++
    }
    
    # Statistike
    $output += "STATISTIKE GRUPA:"
    $output += "-" * 20
    $output += "   Ukupno grupa: $($groupsData.Count)"
    $output += "   Grupa s clanovima: $groupsWithMembers"
    $output += "   Praznih grupa: $($groupsData.Count - $groupsWithMembers)"
    $output += "   Ukupno clanova: $totalMembers"
    $output += ""
    
    return $output
}

function Format-PrintersDataEnhanced {
    param([hashtable]$printersData)
    
    $output = @()
    
    # Instalirani pisaci
    if ($printersData.InstalledPrinters.Count -gt 0) {
        $output += "INSTALIRANI PISACI:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.InstalledPrinters[0].Message) {
            $output += "   > $($printersData.InstalledPrinters[0].Message)"
        } else {
            $printerNumber = 1
            foreach ($printer in $printersData.InstalledPrinters) {
                $output += "[$printerNumber] $($printer.Values | Select-Object -First 1)"
                if ($printer.Count -gt 1) {
                    foreach ($key in $printer.Keys) {
                        if ($printer[$key] -ne ($printer.Values | Select-Object -First 1)) {
                            $output += "   > $key : $($printer[$key])"
                        }
                    }
                }
                $output += ""
                $printerNumber++
            }
        }
        $output += ""
    }
    
    # Povijest printanja
    if ($printersData.PrintHistory.Count -gt 0) {
        $output += "POVIJEST PRINTANJA:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.PrintHistory[0].Message) {
            $output += "   > $($printersData.PrintHistory[0].Message)"
        } else {
            $historyNumber = 1
            foreach ($historyItem in $printersData.PrintHistory) {
                $output += "[$historyNumber] Print Job"
                foreach ($key in $historyItem.Keys) {
                    $output += "   > $key : $($historyItem[$key])"
                }
                $output += ""
                $historyNumber++
            }
        }
        $output += ""
    }
    
    # Status pisaca
    if ($printersData.PrinterStatus.Count -gt 0) {
        $output += "STATUS PISACA:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.PrinterStatus[0].Message) {
            $output += "   > $($printersData.PrinterStatus[0].Message)"
        } else {
            $statusNumber = 1
            foreach ($statusItem in $printersData.PrinterStatus) {
                $output += "[$statusNumber] Status Info"
                foreach ($key in $statusItem.Keys) {
                    $output += "   > $key : $($statusItem[$key])"
                }
                $output += ""
                $statusNumber++
            }
        }
        $output += ""
    }
    
    # Statistike
    $totalPrinters = if ($printersData.InstalledPrinters[0].Message) { 0 } else { $printersData.InstalledPrinters.Count }
    $totalHistory = if ($printersData.PrintHistory[0].Message) { 0 } else { $printersData.PrintHistory.Count }
    $totalStatus = if ($printersData.PrinterStatus[0].Message) { 0 } else { $printersData.PrinterStatus.Count }
    
    $output += "STATISTIKE PRINTANJA:"
    $output += "-" * 30
    $output += "Ukupno pisaca: $totalPrinters"
    $output += "Zapisa o printanju: $totalHistory"
    $output += "Status informacija: $totalStatus"
    $output += ""
    
    return $output
}


# === STRUKTURIRANI TXT EXPORT S PREGLEDNIM FORMATIRANJEM ===
function Export-TxtStructured {
    param (
        [string]$FilePath,
        $control
    )

    try {
        Write-PCInfoLog "Starting structured TXT export to: $FilePath"
        
        $html = ""
        if ($control -is [System.Windows.Forms.WebBrowser]) {
            $html = $control.DocumentText
        } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
            $html = $control.Text
        }

        if ([string]::IsNullOrEmpty($html)) {
            throw "No HTML content for export"
        }

        $output = @()
        
        # Poboljšani header
        $title = Get-TitleFromHTML -html $html
        if ($title) {
            $output += "+" + "-" * ($title.Length + 2) + "+"
            $output += "| $($title.ToUpper()) |"
            $output += "+" + "-" * ($title.Length + 2) + "+"
            $output += ""
        }

        # Metadata section
        $output += "INFORMACIJE O IZVJEŠTAJU"
        $output += "-" * 30
        $output += "Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
        $output += "Sustav: $env:COMPUTERNAME"
        $output += "Korisnik: $env:USERNAME"
        $output += "Aplikacija: Get-PCInfo v$Script:AppVersion"
        $output += "Autor: $Script:AppAuthor"
        $output += ""

        # Poboljšano parsiranje podataka
        $reportType = Detect-ReportType -title $title -html $html
        Write-PCInfoLog "Processing report type: $reportType"
        
        switch ($reportType) {
            "Groups" {
                $groupsData = Parse-GroupsFromHTML -html $html
                $output += Format-GroupsDataEnhanced -groupsData $groupsData
            }
            "Users" {
                $tableData = Parse-HTMLTableStructured -html $html
                $output += Format-UsersDataEnhanced -tableData $tableData
            }
            "Printers" {
                $printersData = Parse-PrintersFromHTML -html $html
                $output += Format-PrintersDataEnhanced -printersData $printersData
            }
            default {
                $tableData = Parse-HTMLTableStructured -html $html
                if ($tableData.Count -gt 0) {
                    $output += "PODACI IZ TABLICE:"
                    $output += "-" * 25
                    $output += ""
                    
                    $rowNumber = 1
                    foreach ($row in $tableData) {
                        $output += "[$rowNumber]"
                        for ($i = 0; $i -lt $row.Count; $i++) {
                            if ($row[$i] -and $row[$i].Trim() -ne "") {
                                $output += "   > $($i+1). $($row[$i])"
                            }
                        }
                        $output += ""
                        $rowNumber++
                    }
                } else {
                    $output += "Nema strukturiranih podataka u tablici"
                }
            }
        }

        # Poboljšani footer
        $output += ""
        $output += "+" + "-" * 48 + "+"
        $output += "|" + " " * 16 + "STATISTIKE" + " " * 16 + "|"
        $output += "+" + "-" * 48 + "+"
        $output += "| Redaka eksportiranih: $($output.Count - 20)".PadRight(46) + "|"
        $output += "| Vrijeme izvoza: $(Get-Date -Format 'HH:mm:ss')".PadRight(46) + "|"
        $output += "| Velièina: ~$([math]::Round(($output -join "`n").Length / 1024, 2)) KB".PadRight(46) + "|"
        $output += "+" + "-" * 48 + "+"

        # Spremi datoteku s boljim encoding
        $output | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        Write-PCInfoLog "TXT export completed successfully. File size: $((Get-Item $FilePath).Length) bytes"

    } catch {
        Write-PCInfoLog "Error in TXT export: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# === FUNKCIJA ZA PARSIRANJE GRUPA ===
function Parse-GroupsFromHTML {
    param([string]$html)
    
    try {
        Write-Host "Parsing groups from HTML..." -ForegroundColor Gray
        
        $groups = @()
        
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        $h3Pattern = '(?s)<h3[^>]*>(.*?)</h3>'
        $h3Matches = [regex]::Matches($cleanHtml, $h3Pattern)
        
        foreach ($h3Match in $h3Matches) {
            $groupName = $h3Match.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', ' '
            $groupName = $groupName.Trim()
            
            if ($groupName -eq '') { continue }
            
            $afterH3 = $cleanHtml.Substring($h3Match.Index + $h3Match.Length)
            $nextH3 = $afterH3.IndexOf('<h3')
            if ($nextH3 -eq -1) { $nextH3 = 2000 }
            
            $searchArea = $afterH3.Substring(0, [Math]::Min($nextH3, $afterH3.Length))
            
            if ($searchArea -match 'Nema èlanova') {
                $groups += @{
                    Name = $groupName
                    HasMembers = $false
                    Members = @()
                }
                continue
            }
            
            if ($searchArea -match '(?s)<table[^>]*>(.*?)</table>') {
                $tableContent = $matches[1]
                $members = Parse-GroupMembersFromTable -tableContent $tableContent
                
                $groups += @{
                    Name = $groupName
                    HasMembers = ($members.Count -gt 0)
                    Members = $members
                }
            } else {
                $groups += @{
                    Name = $groupName
                    HasMembers = $false
                    Members = @()
                }
            }
        }
        
        return $groups
        
    } catch {
        Write-Warning "Error parsing groups: $_"
        return @()
    }
}

# === FUNKCIJA ZA PARSIRANJE ÈLANOVA GRUPE ===
function Parse-GroupMembersFromTable {
    param([string]$tableContent)
    
    try {
        $members = @()
        
        $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
        $rowMatches = [regex]::Matches($tableContent, $rowPattern)
        
        $isFirstRow = $true
        foreach ($rowMatch in $rowMatches) {
            if ($isFirstRow) {
                $isFirstRow = $false
                continue
            }
            
            $rowContent = $rowMatch.Groups[1].Value
            $cells = @()
            
            $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
            $cellMatches = [regex]::Matches($rowContent, $cellPattern)
            
            foreach ($cellMatch in $cellMatches) {
                $cellContent = $cellMatch.Groups[1].Value
                $cleanCell = $cellContent -replace '<[^>]+>', ''
                $cleanCell = $cleanCell -replace '&nbsp;', ' '
                $cleanCell = $cleanCell -replace '&amp;', '&'
                $cleanCell = $cleanCell -replace '\s+', ' '
                $cleanCell = $cleanCell.Trim()
                $cells += $cleanCell
            }
            
            if ($cells.Count -ge 3) {
                $members += @{
                    Username = $cells[0]
                    FullName = if ($cells.Count -gt 1) { $cells[1] } else { "" }
                    Type = if ($cells.Count -gt 2) { $cells[2] } else { "" }
                    Status = if ($cells.Count -gt 3) { $cells[3] } else { "" }
                    SID = if ($cells.Count -gt 4) { $cells[4] } else { "" }
                    Created = if ($cells.Count -gt 5) { $cells[5] } else { "" }
                    LastLogon = if ($cells.Count -gt 6) { $cells[6] } else { "" }
                }
            }
        }
        
        return $members
        
    } catch {
        Write-Warning "Error parsing group members: $_"
        return @()
    }
}


# === FUNKCIJA ZA FORMATIRANJE GRUPA ===
function Format-GroupsData {
    param([array]$groupsData)
    
    $output = @()
    
    if ($groupsData.Count -eq 0) {
        $output += "Nema pronaðenih grupa."
        return $output
    }
    
    $groupNumber = 1
    $totalMembers = 0
    $groupsWithMembers = 0
    
    foreach ($group in $groupsData) {
        $output += "[$groupNumber] $($group.Name)"
        $output += ""
        
        if (-not $group.HasMembers -or $group.Members.Count -eq 0) {
            $output += "    Nema èlanova"
            $output += ""
        } else {
            $groupsWithMembers++
            $totalMembers += $group.Members.Count
            
            $output += "    Èlanovi ($($group.Members.Count)):"
            $memberNumber = 1
            
            foreach ($member in $group.Members) {
                $output += "    [$memberNumber] $($member.Username)"
                if ($member.FullName -and $member.FullName -ne "") {
                    $output += "        Puno ime: $($member.FullName)"
                }
                if ($member.Type -and $member.Type -ne "") {
                    $output += "        Tip: $($member.Type)"
                }
                if ($member.Status -and $member.Status -ne "") {
                    $output += "        Status: $($member.Status)"
                }
                if ($member.Created -and $member.Created -ne "") {
                    $output += "        Kreiran: $($member.Created)"
                }
                if ($member.LastLogon -and $member.LastLogon -ne "" -and $member.LastLogon -ne "-") {
                    $output += "        Zadnja prijava: $($member.LastLogon)"
                }
                
                $output += ""
                $memberNumber++
            }
        }
        
        $output += "-" * 50
        $output += ""
        $groupNumber++
    }
    
    # Statistike
    $output += "STATISTIKE GRUPA:"
    $output += "-" * 20
    $output += "Ukupno grupa: $($groupsData.Count)"
    $output += "Grupa s èlanovima: $groupsWithMembers"
    $output += "Praznih grupa: $($groupsData.Count - $groupsWithMembers)"
    $output += "Ukupno èlanova: $totalMembers"
    $output += ""
    
    return $output
}


# === Funkcije za strukturirani parsing ===
function Get-TitleFromHTML {
    param([string]$html)
    
    if ($html -match '<h1[^>]*>(.*?)</h1>') {
        $title = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        return $title.Trim()
    }
    return "IZVJESTAJ"
}


function Parse-HTMLTableStructured {
    param([string]$html)
    
    try {
        $tableData = @()
        
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        }
        
        return $tableData
        
    } catch {
        Write-Warning "Error in structured parsing: $_"
        return @()
    }
}

function Clean-HTMLForStructured {
    param([string]$html)
    
    $cleanText = $html -replace '(?s)<style[^>]*>.*?</style>', ''
    $cleanText = $cleanText -replace '(?s)<script[^>]*>.*?</script>', ''
    $cleanText = $cleanText -replace '<br[^>]*>', "`n"
    $cleanText = $cleanText -replace '<[^>]+>', ' '
    $cleanText = $cleanText -replace '&nbsp;', ' '
    $cleanText = $cleanText -replace '&amp;', '&'
    $cleanText = $cleanText -replace '&lt;', '<'
    $cleanText = $cleanText -replace '&gt;', '>'
    $cleanText = $cleanText -replace '&quot;', '"'
    $cleanText = $cleanText -replace '\s+', ' '
    $cleanText = $cleanText.Trim()
    
    return $cleanText
}

function Split-TextIntoBlocks {
    param([string]$text)
    
    $words = $text -split ' '
    $blocks = @()
    $currentBlock = ""
    
    foreach ($word in $words) {
        if (($currentBlock + " " + $word).Length -le 80) {
            if ($currentBlock -eq "") {
                $currentBlock = $word
            } else {
                $currentBlock += " " + $word
            }
        } else {
            if ($currentBlock -ne "") {
                $blocks += $currentBlock
            }
            $currentBlock = $word
        }
    }
    
    if ($currentBlock -ne "") {
        $blocks += $currentBlock
    }
    
    return $blocks
}

# === STRUKTURIRANI TXT EXPORT S PREGLEDNIM FORMATIRANJEM ===
function Export-TxtStructured {
    param (
        [string]$FilePath,
        $control
    )

    try {
        Write-Host "Creating structured TXT export..." -ForegroundColor Yellow
        
        $html = ""
        if ($control -is [System.Windows.Forms.WebBrowser]) {
            $html = $control.DocumentText
        } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
            $html = $control.Text
        }

        if ([string]::IsNullOrEmpty($html)) {
            throw "No HTML content for export"
        }

        $output = @()
        
        # Dohvati naslov
        $title = Get-TitleFromHTML -html $html
        if ($title) {
            $output += $title.ToUpper()
            $output += "=" * $title.Length
            $output += ""
        }

        # Odredi tip izvještaja i formatiraj prema tome
        $reportType = Detect-ReportType -title $title -html $html
        
        switch ($reportType) {
            "Groups" {
                Write-Host "Processing Groups report..." -ForegroundColor Cyan
                $groupsData = Parse-GroupsFromHTML -html $html
                $output += Format-GroupsData -groupsData $groupsData
            }
            "FolderPermissions" {
                Write-Host "Processing Folder Permissions report..." -ForegroundColor Cyan
                $foldersData = Parse-FolderPermissionsFromHTML -html $html
                $output += Format-FolderPermissionsData -foldersData $foldersData
            }
            "Users" {
                Write-Host "Processing Users report..." -ForegroundColor Cyan
                $description = Get-DescriptionFromHTML -html $html
                if ($description) {
                    $output += $description
                    $output += ""
                }

                $tableData = Parse-HTMLTableStructured -html $html
                
                if ($tableData.Count -gt 0) {
                    $userCount = 1
                    foreach ($user in $tableData) {
                        if ($user.Count -ge 4) {
                            $output += "[$userCount] $($user[0])"
                            $output += "    Puno ime: $($user[1])"
                            $output += "    Aktivan: $($user[2])"
                            $output += "    Zadnja lozinka: $($user[3])"
                            
                            if ($user.Count -gt 4) {
                                $output += "    Zadnja prijava: $($user[4])"
                            }
                            if ($user.Count -gt 5) {
                                $output += "    Istek racuna: $($user[5])"
                            }
                            if ($user.Count -gt 6) {
                                $output += "    Onemogucen: $($user[6])"
                            }
                            if ($user.Count -gt 10) {
                                $output += "    Grupe: $($user[10])"
                            }
                            
                            $output += ""
                            $userCount++
                        }
                    }
                    
                    # Statistike za korisnièke izvještaje
                    $output += "STATISTIKE:"
                    $output += "-" * 20
                    $output += "Ukupno korisnika: $($tableData.Count)"
                    
                    $activeUsers = $tableData | Where-Object { $_.Count -gt 2 -and $_[2] -eq "Yes" }
                    $inactiveUsers = $tableData | Where-Object { $_.Count -gt 2 -and $_[2] -eq "No" }
                    
                    $output += "Aktivni korisnici: $($activeUsers.Count)"
                    $output += "Neaktivni korisnici: $($inactiveUsers.Count)"
                    $output += ""
                }
            }
            default {
                Write-Host "Processing Generic report..." -ForegroundColor Cyan
                # Pokušaj prvo tablicu, zatim fallback na jednostavan tekst
                $tableData = Parse-HTMLTableStructured -html $html
                
                if ($tableData.Count -gt 0) {
                    $output += "PODACI IZ TABLICE:"
                    $output += "-" * 20
                    $output += ""
                    
                    $rowNumber = 1
                    foreach ($row in $tableData) {
                        $output += "[$rowNumber]"
                        for ($i = 0; $i -lt $row.Count; $i++) {
                            if ($row[$i] -and $row[$i].Trim() -ne "") {
                                $output += "    Kolona $($i+1): $($row[$i])"
                            }
                        }
                        $output += ""
                        $rowNumber++
                    }
                } else {
                    # Fallback na jednostavan tekst
                    $description = Get-DescriptionFromHTML -html $html
                    if ($description) {
                        $output += $description
                        $output += ""
                    }
                    
                    $output += "SADRZAJ IZVJESTAJA:"
                    $output += "-" * 20
                    $output += ""
                    
                    $cleanText = Clean-HTMLForStructured -html $html
                    $blocks = Split-TextIntoBlocks -text $cleanText
                    
                    foreach ($block in $blocks) {
                        $output += $block
                        $output += ""
                    }
                }
            }
        }

        # Footer
        $output += "=" * 60
        $output += "INFORMACIJE O IZVJESTAJU"
        $output += "=" * 60
        $output += "Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
        $output += "Sustav: $env:COMPUTERNAME"

        # Spremi datoteku
        $output | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        Write-Host "Structured TXT export completed: $FilePath" -ForegroundColor Green

    } catch {
        Write-Error "Error in structured TXT export: $($_.Exception.Message)"
        throw
    }
}


# === NOVA FUNKCIJA ZA PARSIRANJE FOLDER PERMISSIONS ===
function Parse-FolderPermissionsFromHTML {
    param([string]$html)
    
    try {
        Write-Host "Parsing folder permissions..." -ForegroundColor Gray
        
        $folders = @()
        
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    $cleanCell = $cellContent -replace '<br[^>]*>', "; "
                    $cleanCell = $cleanCell -replace '<b>', ""
                    $cleanCell = $cleanCell -replace '</b>', ""
                    $cleanCell = $cleanCell -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -ge 2) {
                    $folders += @{
                        Path = $cells[1]
                        Permissions = if ($cells.Count -gt 2) { $cells[2] } else { "N/A" }
                        Number = if ($cells.Count -gt 0) { $cells[0] } else { "" }
                    }
                }
            }
        }
        
        Write-Host "Parsed $($folders.Count) folders" -ForegroundColor Green
        return $folders
        
    } catch {
        Write-Warning "Error parsing folder permissions: $_"
        return @()
    }
}


# === FUNKCIJA ZA FORMATIRANJE FOLDER PERMISSIONS ===
function Format-FolderPermissionsData {
    param([array]$foldersData)
    
    $output = @()
    
    if ($foldersData.Count -eq 0) {
        $output += "Nema pronaðenih folder podataka."
        return $output
    }
    
    $output += "PRAVA PRISTUPA PO FOLDERIMA:"
    $output += "-" * 40
    $output += ""
    
    $folderNumber = 1
    foreach ($folder in $foldersData) {
        $output += "[$folderNumber] $($folder.Path)"
        $output += ""
        
        if ($folder.Permissions -and $folder.Permissions -ne "N/A") {
            $permissions = $folder.Permissions -split ';'
            
            if ($permissions.Count -gt 1) {
                $output += "    Dozvole:"
                foreach ($perm in $permissions) {
                    $cleanPerm = $perm.Trim()
                    if ($cleanPerm -ne "") {
                        $output += "      - $cleanPerm"
                    }
                }
            } else {
                $output += "    Dozvole: $($folder.Permissions)"
            }
        } else {
            $output += "    Dozvole: Nema informacija"
        }
        
        $output += ""
        $output += "-" * 60
        $output += ""
        $folderNumber++
    }
    
    # Statistike
    $output += "STATISTIKE FOLDERA:"
    $output += "-" * 20
    $output += "Ukupno foldera: $($foldersData.Count)"
    
    $foldersWithPermissions = $foldersData | Where-Object { $_.Permissions -and $_.Permissions -ne "N/A" }
    $output += "Foldera s dozvolama: $($foldersWithPermissions.Count)"
    $output += "Foldera bez informacija: $($foldersData.Count - $foldersWithPermissions.Count)"
    $output += ""
    
    return $output
}


# === FUNKCIJA ZA DETEKCIJU TIPA IZVJEŠTAJA ===
function Detect-ReportType {
    param([string]$title, [string]$html)
    
    Write-Host "DEBUG: Detecting report type for title: '$title'" -ForegroundColor Yellow
    
    if ($title -match "grup|Group") {
        Write-Host "DEBUG: Detected as Groups" -ForegroundColor Green
        return "Groups"
    }
    elseif ($title -match "prava|pristup|Permission|Folder") {
        Write-Host "DEBUG: Detected as FolderPermissions" -ForegroundColor Green
        return "FolderPermissions"
    }
    elseif ($title -match "pisaè|printer|Print|ispis|Popis pisaè") {
        Write-Host "DEBUG: Detected as Printers" -ForegroundColor Green
        return "Printers"
    }
    elseif ($title -match "korisnik|User|Account") {
        Write-Host "DEBUG: Detected as Users" -ForegroundColor Green
        return "Users"
    }
    else {
        Write-Host "DEBUG: Detected as Generic (no match found)" -ForegroundColor Red
        return "Generic"
    }
}

# === FUNKCIJA ZA PARSIRANJE PRINTER TABLICA ===
function Parse-PrinterTable {
    param([string]$content, [string]$type)
    
    try {
        $data = @()
        
        # Provjeri ima li paragraf s obavještenjm o nedostajanju podataka
        if ($content -match '<p[^>]*>[^<]*(?:Nema|nema|Ne moguæe|nije dostupno)[^<]*</p>') {
            return @(@{ Message = "Nema podataka za ovu sekciju" })
        }
        
        # Pronaði tablicu u sadržaju
        if ($content -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Pronaði sve redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            $headers = @()
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði sve æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    $cells += $cleanCell
                }
                
                if ($isFirstRow) {
                    $headers = $cells
                    $isFirstRow = $false
                    continue
                }
                
                if ($cells.Count -gt 0) {
                    $rowData = @{}
                    for ($i = 0; $i -lt $cells.Count; $i++) {
                        $key = if ($i -lt $headers.Count -and $headers[$i] -ne "") { 
                            $headers[$i] 
                        } else { 
                            "Column$($i+1)" 
                        }
                        $rowData[$key] = $cells[$i]
                    }
                    $data += $rowData
                }
            }
        }
        
        return $data
        
    } catch {
        Write-Warning "Error parsing printer table ($type): $_"
        return @()
    }
}

# === FUNKCIJA ZA FORMATIRANJE PRINTER PODATAKA ===
function Format-PrintersData {
    param([hashtable]$printersData)
    
    $output = @()
    
    # Instalirani pisaèi
    $output += "INSTALIRANI PISACI:"
    $output += "-" * 30
    $output += ""
    
    if ($printersData.InstalledPrinters.Count -gt 0) {
        if ($printersData.InstalledPrinters[0].Message) {
            $output += "    $($printersData.InstalledPrinters[0].Message)"
        } else {
            $printerNumber = 1
            foreach ($printer in $printersData.InstalledPrinters) {
                if ($printer.Name) {
                    $output += "[$printerNumber] $($printer.Name)"
                } else {
                    $output += "[$printerNumber] Nepoznat pisaè"
                }
                $printerNumber++
            }
        }
    } else {
        $output += "    Nema instaliranih pisaèa"
    }
    $output += ""
    
    # Povijest printanja
    $output += "POVIJEST PRINTANJA:"
    $output += "-" * 30
    $output += ""
    
    if ($printersData.PrintHistory.Count -gt 0) {
        if ($printersData.PrintHistory[0].Message) {
            $output += "    $($printersData.PrintHistory[0].Message)"
        } else {
            $historyNumber = 1
            foreach ($historyItem in $printersData.PrintHistory) {
                $output += "[$historyNumber] Print Job"
                foreach ($key in $historyItem.Keys) {
                    $output += "    $key : $($historyItem[$key])"
                }
                $output += ""
                $historyNumber++
            }
        }
    } else {
        $output += "    Nema zapisa o printanju"
    }
    $output += ""
    
    # Status pisaèa
    $output += "STATUS PISAÈA:"
    $output += "-" * 30
    $output += ""
    
    if ($printersData.PrinterStatus.Count -gt 0) {
        if ($printersData.PrinterStatus[0].Message) {
            $output += "    $($printersData.PrinterStatus[0].Message)"
        } else {
            $statusNumber = 1
            foreach ($statusItem in $printersData.PrinterStatus) {
                $output += "[$statusNumber] Status"
                foreach ($key in $statusItem.Keys) {
                    $output += "    $key : $($statusItem[$key])"
                }
                $output += ""
                $statusNumber++
            }
        }
    } else {
        $output += "    Nema status informacija"
    }
    $output += ""
    
    # Statistike
    $totalPrinters = if ($printersData.InstalledPrinters.Count -gt 0 -and -not $printersData.InstalledPrinters[0].Message) { 
        $printersData.InstalledPrinters.Count 
    } else { 
        0 
    }
    
    $output += "STATISTIKE:"
    $output += "-" * 20
    $output += "Ukupno pisaèa: $totalPrinters"
    $output += "Zapisa o printanju: $($printersData.PrintHistory.Count)"
    $output += "Status informacija: $($printersData.PrinterStatus.Count)"
    $output += ""
    
    return $output
}
	
# === FUNKCIJA ZA FORMATIRANJE PRINTER PODATAKA ===
function Format-PrintersData {
    param([hashtable]$printersData)
    
    $output = @()
    
    # Instalirani pisaèi
    if ($printersData.InstalledPrinters.Count -gt 0) {
        $output += "INSTALIRANI PISACI:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.InstalledPrinters[0].Message) {
            $output += "    $($printersData.InstalledPrinters[0].Message)"
        } else {
            $printerNumber = 1
            foreach ($printer in $printersData.InstalledPrinters) {
                $output += "[$printerNumber] $($printer.Values | Select-Object -First 1)"
                if ($printer.Count -gt 1) {
                    foreach ($key in $printer.Keys) {
                        if ($printer[$key] -ne ($printer.Values | Select-Object -First 1)) {
                            $output += "    $key : $($printer[$key])"
                        }
                    }
                }
                $output += ""
                $printerNumber++
            }
        }
        $output += ""
    }
    
    # Povijest printanja
    if ($printersData.PrintHistory.Count -gt 0) {
        $output += "POVIJEST PRINTANJA:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.PrintHistory[0].Message) {
            $output += "    $($printersData.PrintHistory[0].Message)"
        } else {
            $historyNumber = 1
            foreach ($historyItem in $printersData.PrintHistory) {
                $output += "[$historyNumber] Print Job"
                foreach ($key in $historyItem.Keys) {
                    $output += "    $key : $($historyItem[$key])"
                }
                $output += ""
                $historyNumber++
            }
        }
        $output += ""
    }
    
    # Status pisaèa
    if ($printersData.PrinterStatus.Count -gt 0) {
        $output += "STATUS PISACA:"
        $output += "-" * 30
        $output += ""
        
        if ($printersData.PrinterStatus[0].Message) {
            $output += "    $($printersData.PrinterStatus[0].Message)"
        } else {
            $statusNumber = 1
            foreach ($statusItem in $printersData.PrinterStatus) {
                $output += "[$statusNumber] Status Info"
                foreach ($key in $statusItem.Keys) {
                    $output += "    $key : $($statusItem[$key])"
                }
                $output += ""
                $statusNumber++
            }
        }
        $output += ""
    }
    
    # Statistike
    $totalPrinters = if ($printersData.InstalledPrinters[0].Message) { 0 } else { $printersData.InstalledPrinters.Count }
    $totalHistory = if ($printersData.PrintHistory[0].Message) { 0 } else { $printersData.PrintHistory.Count }
    $totalStatus = if ($printersData.PrinterStatus[0].Message) { 0 } else { $printersData.PrinterStatus.Count }
    
    $output += "STATISTIKE PRINTANJA:"
    $output += "-" * 30
    $output += "Ukupno pisaca: $totalPrinters"
    $output += "Zapisa o printanju: $totalHistory"
    $output += "Status informacija: $totalStatus"
    $output += ""
    
    return $output
}



# === Funkcije za strukturirani parsing ===
function Get-TitleFromHTML {
    param([string]$html)
    
    if ($html -match '<h1[^>]*>(.*?)</h1>') {
        $title = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        return $title.Trim()
    }
    return "IZVJEŠTAJ O KORISNIÈKIM RAÈUNIMA"
}

function Get-DescriptionFromHTML {
    param([string]$html)
    
    if ($html -match '<p[^>]*>(.*?)</p>') {
        $desc = $matches[1] -replace '<[^>]+>', '' -replace '&nbsp;', ' '
        $desc = $desc.Trim()
        if ($desc.Length -gt 10 -and $desc.Length -lt 200) {
            return $desc
        }
    }
    return $null
}

function Parse-HTMLTableStructured {
    param([string]$html)
    
    try {
        $tableData = @()
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Pronaði tablicu
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Pronaði sve redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $isFirstRow = $true
            foreach ($rowMatch in $rowMatches) {
                if ($isFirstRow) {
                    $isFirstRow = $false
                    continue  # Preskoèi header red
                }
                
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði sve æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Èisti sadržaj æelije
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        }
        
        return $tableData
        
    } catch {
        Write-Warning "Greška pri strukturiranom parsiranju: $_"
        return @()
    }
}

function Clean-HTMLForStructured {
    param([string]$html)
    
    # Ukloni style i script tagove
    $cleanText = $html -replace '(?s)<style[^>]*>.*?</style>', ''
    $cleanText = $cleanText -replace '(?s)<script[^>]*>.*?</script>', ''
    
    # Zamijeni <br> s newline
    $cleanText = $cleanText -replace '<br[^>]*>', "`n"
    
    # Ukloni ostale HTML tagove
    $cleanText = $cleanText -replace '<[^>]+>', ' '
    
    # Zamijeni HTML entitete
    $cleanText = $cleanText -replace '&nbsp;', ' '
    $cleanText = $cleanText -replace '&amp;', '&'
    $cleanText = $cleanText -replace '&lt;', '<'
    $cleanText = $cleanText -replace '&gt;', '>'
    $cleanText = $cleanText -replace '&quot;', '"'
    
    # Oèisti whitespace
    $cleanText = $cleanText -replace '\s+', ' '
    $cleanText = $cleanText.Trim()
    
    return $cleanText
}

function Split-TextIntoBlocks {
    param([string]$text)
    
    $words = $text -split ' '
    $blocks = @()
    $currentBlock = ""
    
    foreach ($word in $words) {
        # Ako dodavanje rijeèi ne prelazi 80 znakova
        if (($currentBlock + " " + $word).Length -le 80) {
            if ($currentBlock -eq "") {
                $currentBlock = $word
            } else {
                $currentBlock += " " + $word
            }
        } else {
            # Završi trenutni blok i poèni novi
            if ($currentBlock -ne "") {
                $blocks += $currentBlock
            }
            $currentBlock = $word
        }
    }
    
    # Dodaj zadnji blok
    if ($currentBlock -ne "") {
        $blocks += $currentBlock
    }
    
    return $blocks
}



# === NAJJEDNOSTAVNIJI MOGUÆI TXT EXPORT (ako se i osnovni pokvari) ===
function Export-TxtMinimal {
    param (
        [string]$FilePath,
        $control
    )

    try {
        Write-Host "Kreiram minimalni TXT export..." -ForegroundColor Yellow
        
        # Dohvati HTML
        $html = ""
        if ($control -is [System.Windows.Forms.WebBrowser]) {
            try {
                # Pokušaj dohvatiti samo tekst bez HTML-a
                $html = $control.Document.Body.InnerText
            } catch {
                # Ako ne može innerText, koristi DocumentText
                $html = $control.DocumentText
            }
        } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
            $html = $control.Text
        }

        # Najjednostavnije èišæenje
        $text = $html -replace '<[^>]+>', ''
        $text = $text -replace '&\w+;', ' '
        $text = $text -replace '\s+', ' '
        $text = $text.Trim()
        
        # Osnovni output
        $output = @(
            "IZVJEŠTAJ O KORISNIÈKIM RAÈUNIMA",
            "=" * 40,
            "",
            $text,
            "",
            "Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')",
            "Sustav: $env:COMPUTERNAME"
        )
        
        # Spremi
        $output | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        Write-Host "Minimalni TXT export završen" -ForegroundColor Green
        
    } catch {
        Write-Error "Greška u minimalnom TXT exportu: $_"
        throw
    }
}






# === Sigurnija funkcija za parsiranje s null provjere ===
function Parse-HTMLTableSafe {
    param([string]$html)
    
    try {
        if ([string]::IsNullOrEmpty($html)) {
            Write-Warning "HTML sadržaj je prazan"
            return @()
        }

        Write-Host "  Parsing HTML table (safe)..." -ForegroundColor Gray
        
        $tableData = @()
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        if ([string]::IsNullOrEmpty($cleanHtml)) {
            Write-Warning "HTML je prazan nakon èišæenja"
            return @()
        }

        Write-Host "  Tražim tablicu u HTML-u..." -ForegroundColor Gray
        
        # Pronaði tablicu
        $tablePattern = '(?s)<table[^>]*>(.*?)</table>'
        $tableMatch = [regex]::Match($cleanHtml, $tablePattern)
        
        if ($tableMatch.Success) {
            $tableContent = $tableMatch.Groups[1].Value
            
            if ([string]::IsNullOrEmpty($tableContent)) {
                Write-Warning "Sadržaj tablice je prazan"
                return @()
            }
            
            Write-Host "Tablica pronaðena" -ForegroundColor Gray
            
            # Ukloni thead/tbody wrappere
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            if ($null -eq $rowMatches -or $rowMatches.Count -eq 0) {
                Write-Warning "Nema pronaðenih redova u tablici"
                return @()
            }
            
            Write-Host "  Pronašao $($rowMatches.Count) redova" -ForegroundColor Gray
            
            foreach ($rowMatch in $rowMatches) {
                if ($null -eq $rowMatch -or $null -eq $rowMatch.Groups) {
                    continue
                }

                $rowContent = $rowMatch.Groups[1].Value
                if ([string]::IsNullOrEmpty($rowContent)) {
                    continue
                }

                $cells = @()
                
                # Pronaði æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                if ($null -ne $cellMatches) {
                    foreach ($cellMatch in $cellMatches) {
                        if ($null -eq $cellMatch -or $null -eq $cellMatch.Groups) {
                            $cells += ""
                            continue
                        }

                        $cellContent = $cellMatch.Groups[1].Value
                        
                        # Sigurno èišæenje
                        $cleanCell = if ([string]::IsNullOrEmpty($cellContent)) { "" } else {
                            $cellContent -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"' -replace '\s+', ' '
                        }
                        
                        $cleanCell = if ($null -eq $cleanCell) { "" } else { $cleanCell.Trim() }
                        $cells += $cleanCell
                    }
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        } else {
            Write-Host " Nema HTML tablice" -ForegroundColor Yellow
            return @()
        }
        
        Write-Host " Parsiranje završeno: $($tableData.Count) redova" -ForegroundColor Gray
        return $tableData
        
    } catch {
        Write-Warning "Greška pri sigurnom parsiranju: $($_.Exception.Message)"
        return @()
    }
}





# === Ultra sigurna funkcija za parsiranje ===
function Parse-HTMLTableUltraSafe {
    param([string]$html)
    
    $result = @()
    
    try {
        if ([string]::IsNullOrWhiteSpace($html)) {
            return $result
        }

        # Jednostavniji pristup - direktno tražimo tr i td/th elemente
        Write-Host "  Ultra safe parsing..." -ForegroundColor Gray
        
        # Ukloni problematiène dijelove
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Direktno traži sve tr elemente
        $trPattern = '(?s)<tr[^>]*>(.*?)</tr>'
        $trMatches = [regex]::Matches($cleanHtml, $trPattern)
        
        if ($trMatches.Count -eq 0) {
            Write-Host "  Nema <tr> elemenata" -ForegroundColor Yellow
            return $result
        }
        
        Write-Host "  Pronašao $($trMatches.Count) <tr> elemenata" -ForegroundColor Gray
        
        foreach ($trMatch in $trMatches) {
            if ($null -eq $trMatch.Groups -or $trMatch.Groups.Count -lt 2) {
                continue
            }
            
            $trContent = $trMatch.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($trContent)) {
                continue
            }
            
            # Traži td ili th elemente
            $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
            $cellMatches = [regex]::Matches($trContent, $cellPattern)
            
            if ($cellMatches.Count -eq 0) {
                continue
            }
            
            $rowData = @()
            foreach ($cellMatch in $cellMatches) {
                if ($null -eq $cellMatch.Groups -or $cellMatch.Groups.Count -lt 2) {
                    $rowData += ""
                    continue
                }
                
                $cellContent = $cellMatch.Groups[1].Value
                
                # Osnovano èišæenje
                $cleanCell = $cellContent -replace '<[^>]+>', ''
                $cleanCell = $cleanCell -replace '&\w+;', ' '
                $cleanCell = $cleanCell -replace '\s+', ' '
                $cleanCell = $cleanCell.Trim()
                
                if ($cleanCell.Length -gt 255) {
                    $cleanCell = $cleanCell.Substring(0, 252) + "..."
                }
                
                $rowData += $cleanCell
            }
            
            if ($rowData.Count -gt 0) {
                $result += ,$rowData
            }
        }
        
        Write-Host " Ultra safe parsing: $($result.Count) redova" -ForegroundColor Gray
        return $result
        
    }
    catch {
        Write-Warning "Ultra safe parsing greška: $($_.Exception.Message)"
        return @()
    }
}

# === Ažuriraj Export-ActiveTab ===
function Export-ActiveTab {
    param ([string]$format)

    $tab = $tabControl.SelectedTab
    $control = $tab.Controls | Where-Object { $_ -is [System.Windows.Forms.WebBrowser] -or $_ -is [System.Windows.Forms.RichTextBox] }

    if (-not $control) {
        [System.Windows.Forms.MessageBox]::Show("Nema podataka za izvoz.", "Upozorenje", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $html = ""
    if ($control -is [System.Windows.Forms.WebBrowser]) {
        $html = $control.DocumentText
    } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
        $html = $control.Text
    }

    $path = New-Object System.Windows.Forms.SaveFileDialog
    $path.Filter = switch ($format.ToLower()) {
        "txt"  { "TXT files (*.txt)|*.txt" }
        "html" { "HTML files (*.html)|*.html" }
        "doc"  { "Word files (*.docx)|*.docx" }
        "xls"  { "Excel files (*.xlsx)|*.xlsx" }
        "csv"  { "CSV files (*.csv)|*.csv" }
        "pdf"  { "PDF files (*.pdf)|*.pdf" }
        default { "TXT files (*.txt)|*.txt" }
    }
    
    $path.FileName = "$($tab.Text -replace '[^\w\-_\.]', '_')_Export"

    if ($path.ShowDialog() -eq "OK") {
        try {
            switch ($format.ToLower()) {
                "txt" {
                    Export-TxtStructured -FilePath $path.FileName -control $control
                }
                "xls" {
                    Export-ToExcelWithFallback -htmlContent $html -fileName $path.FileName
                }
                "csv" {
                    Export-ToCSV -htmlContent $html -fileName $path.FileName
                }
                "html" {
                    [System.IO.File]::WriteAllText($path.FileName, $html, [System.Text.Encoding]::UTF8)
                }
                "doc" {
                    Export-ToWord -htmlContent $html -fileName $path.FileName
                }
                "pdf" {
                    Export-ToPDF -htmlContent $html -fileName $path.FileName
                }
            }
            
            [System.Windows.Forms.MessageBox]::Show("Izvoz završen: $($path.FileName)", "Uspjeh", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Greška pri izvozu: $($_.Exception.Message)", "Greška", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# === 1. DIJAGNOSTIKA EXCEL PROBLEMA ===
function Diagnose-ExcelIssue {
    Write-Host "=== EXCEL DIJAGNOSTIKA ===" -ForegroundColor Yellow
    
    # Provjeri je li Excel instaliran
    Write-Host "1. Provjera Excel instalacije..." -ForegroundColor Cyan
    
    $excelPath = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles}\Microsoft Office\Office15\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\EXCEL.EXE"
    )
    
    $excelFound = $false
    foreach ($path in $excelPath) {
        if (Test-Path $path) {
            Write-Host "  ? Excel pronaðen: $path" -ForegroundColor Green
            $excelFound = $true
            break
        }
    }
    
    if (-not $excelFound) {
        Write-Host "  ? Excel NIJE pronaðen na uobièajenim lokacijama" -ForegroundColor Red
    }
    
    # Provjeri registry
    Write-Host "2. Provjera Registry COM registracije..." -ForegroundColor Cyan
    
    $regPaths = @(
        "HKEY_CLASSES_ROOT\Excel.Application",
        "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Excel.Application"
    )
    
    foreach ($regPath in $regPaths) {
        try {
            $key = Get-ItemProperty -Path "Registry::$regPath" -ErrorAction SilentlyContinue
            if ($key) {
                Write-Host "  ? Registry kljuè postoji: $regPath" -ForegroundColor Green
            } else {
                Write-Host "  ? Registry kljuè ne postoji: $regPath" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ? Greška pri èitanju: $regPath - $_" -ForegroundColor Red
        }
    }
    
    # Provjeri proces privilegije
    Write-Host "3. Provjera procesa i privilegija..." -ForegroundColor Cyan
    
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    Write-Host "  Administrator privilegije: $isAdmin" -ForegroundColor $(if($isAdmin){"Green"}else{"Yellow"})
    Write-Host "  Trenutni korisnik: $($currentUser.Name)" -ForegroundColor Gray
    Write-Host "  Proces ID: $PID" -ForegroundColor Gray
    
    # Provjeri pokrenute Excel procese
    $excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Host "  ??  Excel procesi veæ pokrenuti: $($excelProcesses.Count)" -ForegroundColor Yellow
        foreach ($proc in $excelProcesses) {
            Write-Host "    PID: $($proc.Id), StartTime: $($proc.StartTime)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  ? Nema pokrenutih Excel procesa" -ForegroundColor Green
    }
    
    Write-Host "=== KRAJ DIJAGNOSTIKE ===" -ForegroundColor Yellow
}

# === 2. ALTERNATIVNO RJEŠENJE - EXPORT PUTEM CSV I PRETVORBA ===
function Export-ToExcelViaCSV {
    param(
        [string]$htmlContent,
        [string]$fileName
    )
    
    try {
        Write-Host "Excel COM ne radi, koristim CSV -> Excel konverziju..." -ForegroundColor Yellow
        
        # 1. Kreiraj CSV datoteku
        $csvFileName = $fileName -replace '\.xlsx?$', '.csv'
        Export-ToCSV -htmlContent $htmlContent -fileName $csvFileName
        
        if (-not (Test-Path $csvFileName)) {
            throw "CSV datoteka nije kreirana"
        }
        
        Write-Host "? CSV datoteka kreirana: $csvFileName" -ForegroundColor Green
        
        # 2. Pokušaj konvertirati CSV u Excel putem PowerShell Excel modula
        try {
            # Provjeri je li ImportExcel modul dostupan
            if (-not (Get-Module -ListAvailable -Name "ImportExcel")) {
                Write-Host "ImportExcel modul nije instaliran, instaliram..." -ForegroundColor Yellow
                Install-Module -Name ImportExcel -Force -Scope CurrentUser
            }
            
            Import-Module ImportExcel
            
            # Uèitaj CSV i spremi kao Excel
            $csvData = Import-Csv $csvFileName
            $csvData | Export-Excel -Path $fileName -AutoSize -TableStyle Medium2 -BoldTopRow
            
            Write-Host "? Excel datoteka kreirana putem ImportExcel modula!" -ForegroundColor Green
            
            # Obriši privremenu CSV datoteku
            Remove-Item $csvFileName -Force -ErrorAction SilentlyContinue
            
        } catch {
            Write-Warning "ImportExcel modul neuspješan: $_"
            
            # 3. Alternativa - ostavi CSV i preimenuj u .xlsx
            Write-Host "Kreiram osnovnu Excel datoteku..." -ForegroundColor Yellow
            
            # Èitaj CSV i kreiraj jednostavan Excel XML format
            $csvContent = Get-Content $csvFileName -Encoding UTF8
            $excelXML = Convert-CSVToExcelXML -csvLines $csvContent
            
            # Spremi kao .xlsx (zapravo XML)
            $excelXML | Out-File -FilePath $fileName -Encoding UTF8
            
            Write-Host " Osnovna Excel datoteka kreirana" -ForegroundColor Green
        }
        
    } catch {
        throw "CSV -> Excel konverzija neuspješna: $_"
    }
}


function Export-ToExcelSimple {
    param(
        [array]$Data,
        [string]$FilePath
    )
    
    try {
        # Kreiraj Excel aplikaciju
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Kreiraj workbook
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Dodaj podatke
        $row = 1
        foreach ($item in $Data) {
            $col = 1
            foreach ($property in $item.PSObject.Properties) {
                if ($row -eq 1) {
                    # Dodaj zaglavlja
                    $worksheet.Cells.Item($row, $col) = $property.Name
                }
                $worksheet.Cells.Item($row + 1, $col) = $property.Value
                $col++
            }
            $row++
        }
        
        # Spremi kao Excel
        $workbook.SaveAs($FilePath, 51) # 51 = xlOpenXMLWorkbook
        $workbook.Close()
        $excel.Quit()
        
        # Cleanup
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
    } catch {
        throw "Excel export failed: $_"
    }
}

# Koristite ovako:
# Export-ToExcelSimple -Data $userData -FilePath "C:\path\to\file.xlsx"



# === 3. KONVERZIJA CSV-a U EXCEL XML FORMAT ===
function Convert-CSVToExcelXML {
    param([array]$csvLines)
    
    $xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<Worksheet name="Izvještaj">
<SheetData>
"@
    
    $rowNum = 1
    foreach ($line in $csvLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        
        $cells = $line -split ','
        $xml += "<Row r='$rowNum'>"
        
        $colNum = 1
        foreach ($cell in $cells) {
            $cleanCell = $cell.Trim('"') -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $colLetter = [char](64 + $colNum)
            $xml += "<Cell r='$colLetter$rowNum'><Value>$cleanCell</Value></Cell>"
            $colNum++
        }
        
        $xml += "</Row>"
        $rowNum++
    }
    
    $xml += @"
</SheetData>
</Worksheet>
</Workbook>
"@
    
    return $xml
}

# === 4. POBOLJŠANA EXPORT FUNKCIJA S FALLBACK OPCIJAMA ===
function Export-ToExcelWithFallback {
    param(
        [string]$htmlContent,
        [string]$fileName
    )
    
    Write-Host "=== EXCEL EXPORT S FALLBACK ===" -ForegroundColor Yellow
    
    # Opcija 1: Pokušaj standardni COM pristup
    try {
        Write-Host "Pokušavam standardni Excel COM..." -ForegroundColor Cyan
        
        # Vrlo kratki test
        $testExcel = New-Object -ComObject Excel.Application
        $testExcel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testExcel) | Out-Null
        
        # Ako test proðe, pokušaj pravi export
        Export-ToExcelSuperSafe -htmlContent $htmlContent -fileName $fileName
        Write-Host "? Standardni Excel COM uspješan!" -ForegroundColor Green
        return
        
    } catch {
        Write-Host "? Standardni Excel COM neuspješan: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Opcija 2: CSV -> Excel konverzija
    try {
        Write-Host "Pokušavam CSV -> Excel konverziju..." -ForegroundColor Cyan
        Export-ToExcelViaCSV -htmlContent $htmlContent -fileName $fileName
        Write-Host "? CSV -> Excel konverzija uspješna!" -ForegroundColor Green
        return
        
    } catch {
        Write-Host "? CSV -> Excel konverzija neuspješna: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Opcija 3: Samo CSV s upozorenjem
    try {
        Write-Host "Kreiram samo CSV datoteku..." -ForegroundColor Yellow
        $csvFileName = $fileName -replace '\.xlsx?$', '.csv'
        Export-ToCSV -htmlContent $htmlContent -fileName $csvFileName
        
        [System.Windows.Forms.MessageBox]::Show(
            "Excel COM objekti ne rade na ovom sustavu.`n`nKreirana je CSV datoteka umjesto Excel datoteke:`n$csvFileName`n`nMožete otvoriti CSV datoteku u Excelu.", 
            "Excel Fallback", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        Write-Host "? CSV datoteka kreirana kao fallback" -ForegroundColor Green
        return
        
    } catch {
        Write-Host "? I CSV fallback neuspješan: $($_.Exception.Message)" -ForegroundColor Red
        throw "Svi Excel export pristupi neuspješni"
    }
}



# === Jednostavnija funkcija za parsiranje HTML tablice ===
function Parse-HTMLTableSimple {
    param([string]$html)
    
    try {
        Write-Host "  Parsing HTML table..." -ForegroundColor Gray
        
        $tableData = @()
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        Write-Host "  Tražim tablicu u HTML-u..." -ForegroundColor Gray
        
        # Pronaði tablicu - fleksibilniji pristup
        $tablePattern = '(?s)<table[^>]*>(.*?)</table>'
        if ($cleanHtml -match $tablePattern) {
            $tableContent = $matches[1]
            Write-Host "  ? Tablica pronaðena" -ForegroundColor Gray
            
            # Ukloni thead/tbody wrappere
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            Write-Host "  Pronašao $($rowMatches.Count) redova" -ForegroundColor Gray
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Jednostavno èišæenje
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    # Ogranièi duljinu za Excel
                    if ($cleanCell.Length -gt 32767) {
                        $cleanCell = $cleanCell.Substring(0, 32760) + "..."
                    }
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        } else {
            Write-Host "  Nema HTML tablice" -ForegroundColor Yellow
            return @()
        }
        
        Write-Host "  ? Parsiranje završeno: $($tableData.Count) redova" -ForegroundColor Gray
        return $tableData
        
    } catch {
        Write-Warning "Greška pri parsiranju: $_"
        return @()
    }
}

# === Dodaj ovu funkciju u svoj kod ===
function Export-ToCSVEnhanced {
    param(
        [string]$htmlContent,
        [string]$fileName
    )

    try {
        Write-PCInfoLog "Starting enhanced CSV export to: $fileName"
        
        # Koristi vašu postojeæu Parse-HTMLTableForCSV funkciju ako postoji, 
        # ili Parse-HTMLTableToArray
        $tableData = if (Get-Command Parse-HTMLTableForCSV -ErrorAction SilentlyContinue) {
            Parse-HTMLTableForCSV -html $htmlContent
        } else {
            Parse-HTMLTableToArray -html $htmlContent
        }
        
        if ($tableData.Count -eq 0) {
            throw "No table data found for CSV export"
        }
        
        Write-PCInfoLog "Found $($tableData.Count) rows for CSV export"

        # Poboljšani CSV s metadata
        $csvContent = @()
        
        # CSV Header comments
        $csvContent += "# Get-PCInfo Professional v$Script:AppVersion"
        $csvContent += "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $csvContent += "# System: $env:COMPUTERNAME"
        $csvContent += "# User: $env:USERNAME"
        $csvContent += "# Author: $Script:AppAuthor"
        $csvContent += "#"
        
        # Data rows s poboljšanim escaping
        foreach ($row in $tableData) {
            $escapedRow = $row | ForEach-Object {
                $field = if ($null -eq $_) { "" } else { $_.ToString() }
                $field = $field -replace '"', '""'  # Escape quotes
                
                # Quote fields that contain special characters
                if ($field -match '[,";`n`r`t]' -or $field.StartsWith(' ') -or $field.EndsWith(' ')) {
                    "`"$field`""
                } else {
                    $field
                }
            }
            $csvContent += $escapedRow -join ','
        }
        
        # Spremi s UTF-8 BOM za Excel kompatibilnost
        $utf8BOM = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllLines($fileName, $csvContent, $utf8BOM)
        
        Write-PCInfoLog "CSV export completed. Rows: $($tableData.Count), Size: $((Get-Item $fileName).Length) bytes"
        
        return @{ Success = $true; Path = $fileName; Rows = $tableData.Count }
        
    } catch {
        Write-PCInfoLog "CSV export failed: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}


# === Pomoæna funkcija za parsiranje HTML tablice u CSV ===
function Parse-HTMLTableForCSV {
    param([string]$html)
    
    try {
        $tableData = @()
        
        # Ukloni CSS i JavaScript
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', '' -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Pronaði tablicu
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Ukloni thead/tbody/tfoot wrappere
            $tableContent = $tableContent -replace '(?s)</?thead[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tbody[^>]*>', ''
            $tableContent = $tableContent -replace '(?s)</?tfoot[^>]*>', ''
            
            # Pronaði sve redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði sve æelije (th ili td)
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Oèisti HTML tagove, entitete i formatiranje
                    $cleanCell = $cellContent -replace '<br[^>]*>', "`n"  # Zamijeni <br> s newline
                    $cleanCell = $cleanCell -replace '<[^>]+>', ''  # Ukloni HTML tagove
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\r?\n\s*\r?\n', "`n"  # Ukloni multiple newlines
                    $cleanCell = $cleanCell -replace '^\s+|\s+$', ''  # Trim
                    
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells
                }
            }
        }
        
        Write-Host "Parsirano $($tableData.Count) redova za CSV" -ForegroundColor Cyan
        return $tableData
        
    } catch {
        Write-Warning "Greška pri parsiranju HTML tablice za CSV: $_"
        return @()
    }
}



# === Funkcija za ispis ===
function Print-ActiveTab {
    $activeTab = $tabControl.SelectedTab
    $control = $activeTab.Controls[0]

    if ($control -is [System.Windows.Forms.WebBrowser]) {
        $control.ShowPrintDialog()
    } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms

        $printDoc = New-Object System.Drawing.Printing.PrintDocument
        $font = New-Object System.Drawing.Font("Consolas", 10)
        $lines = $control.Lines
        $lineIndex = 0

        $printDoc.add_PrintPage({
            param($sender, $e)
            $y = $e.MarginBounds.Top
            while ($lineIndex -lt $lines.Length) {
                $line = $lines[$lineIndex]
                $e.Graphics.DrawString($line, $font, [System.Drawing.Brushes]::Black, $e.MarginBounds.Left, $y)
                $y += $font.Height
                $lineIndex++
                if ($y + $font.Height -gt $e.MarginBounds.Bottom) {
                    $e.HasMorePages = $true
                    return
                }
            }
            $e.HasMorePages = $false
        })

        $dlg = New-Object System.Windows.Forms.PrintDialog
        $dlg.Document = $printDoc
        if ($dlg.ShowDialog() -eq "OK") {
            $printDoc.Print()
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ispis nije podržan za ovu vrstu kontrole.")
    }
}







































# === Funkcija: Dodaj tab koji prikazuje HTML (korisnici) ===
function Add-TabWeb {
    param (
        [string]$name,
        [string]$htmlFilePath
    )

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $name

    $browser = New-Object System.Windows.Forms.WebBrowser
    $browser.Dock = "Fill"
    $browser.ScriptErrorsSuppressed = $true

    # Kada se dokument uèita, ukloni horizontalni skroler
    $browser.DocumentCompleted += {
        try {
            $browser.Document.Body.Style = "overflow-x: hidden; margin: 0; padding: 0;"
        } catch {
            Write-Host "Ne mogu pristupiti DOM objektu."
        }
    }

    $browser.Url = (New-Object System.Uri($htmlFilePath))
    $tab.Controls.Add($browser)
    $tabControl.TabPages.Add($tab)
}

# === Funkcija da pošaljemo HTML u clipboard ===
function Set-ClipboardHtml {
    param([string]$html)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Web

    $dataObject = New-Object System.Windows.Forms.DataObject
    $dataFormat = "HTML Format"

    # Priprema HTML header za clipboard format
    $htmlHeader = @"
Version:1.0
StartHTML:00000097
EndHTML:{0}
StartFragment:00000133
EndFragment:{1}
<!DOCTYPE html>
<html><body>
<!--StartFragment-->
{2}
<!--EndFragment-->
</body></html>
"@

    $fragment = $html
    $htmlContent = [string]::Format($htmlHeader, 99999, 99999, $fragment)

    # Precizno izraèunaj offsete za clipboard format
    $startHtml = 97
    $startFragment = $htmlContent.IndexOf("<!--StartFragment-->") + 20
    $endFragment = $htmlContent.IndexOf("<!--EndFragment-->")

    $htmlContent = $htmlContent.Replace("StartHTML:00000097", "StartHTML:$("{0:D8}" -f $startHtml)")
    $htmlContent = $htmlContent.Replace("EndHTML:{0}", "EndHTML:$("{0:D8}" -f $htmlContent.Length)")
    $htmlContent = $htmlContent.Replace("EndFragment:{1}", "EndFragment:$("{0:D8}" -f $endFragment)")

    $dataObject.SetData($dataFormat, $true, $htmlContent)
    [System.Windows.Forms.Clipboard]::SetDataObject($dataObject, $true)
}

# === 5. AŽURIRANA EXPORT-ACTIVETAB FUNKCIJA ===
function Export-ActiveTabEnhanced {
    param ([string]$format)

    Write-PCInfoLog "=== Starting export process for format: $format ==="

    $tab = $tabControl.SelectedTab
    $control = $tab.Controls | Where-Object { $_ -is [System.Windows.Forms.WebBrowser] -or $_ -is [System.Windows.Forms.RichTextBox] }

    if (-not $control) {
        Write-PCInfoLog "No exportable content found in active tab" -Level "WARNING"
        [System.Windows.Forms.MessageBox]::Show("Nema podataka za izvoz.", "Upozorenje", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $html = ""
    if ($control -is [System.Windows.Forms.WebBrowser]) {
        $html = $control.DocumentText
    } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
        $html = $control.Text
    }

    # Poboljšani save dialog
    $path = New-Object System.Windows.Forms.SaveFileDialog
    $path.Title = "Export $($tab.Text) Report - Get-PCInfo v$Script:AppVersion"
    
    # Generiraj sigurno ime datoteke
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $safeName = $tab.Text -replace '[^\w\-_\.]', '_'
    
    $path.Filter = switch ($format.ToLower()) {
        "txt"  { "Text files (*.txt)|*.txt|All files (*.*)|*.*" }
        "html" { "HTML files (*.html)|*.html|All files (*.*)|*.*" }
        "doc"  { "Word files (*.docx)|*.docx|All files (*.*)|*.*" }
        "xls"  { "Excel files (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|All files (*.*)|*.*" }
        "csv"  { "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*" }
        "pdf"  { "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*" }
        default { "Text files (*.txt)|*.txt|All files (*.*)|*.*" }
    }
    
    $path.FileName = "${safeName}_${timestamp}"
    $path.DefaultExt = $format

    if ($path.ShowDialog() -eq "OK") {
        try {
            Write-PCInfoLog "Export started - File: $($path.FileName), Format: $format"
            $startTime = Get-Date
            
            # Pozovi postojeæe funkcije s poboljšanjima
            switch ($format.ToLower()) {
                "txt" {
                    # Koristi vašu poboljšanu Export-TxtStructured
                    Export-TxtStructured -FilePath $path.FileName -control $control
                }
                "csv" {
                    # Koristi novu poboljšanu CSV funkciju
                    Export-ToCSVEnhanced -htmlContent $html -fileName $path.FileName
                }
                "xls" {
    try {
        # Koristite ImportExcel modul
        if (Get-Module -ListAvailable -Name "ImportExcel") {
            # Parsiraj HTML u podatke
            $tableData = Parse-HTMLTableToArray -html $html
            
            if ($tableData.Count -gt 0) {
                # Konvertiraj u objekte
                $objects = @()
                $headers = $tableData[0]
                
                for ($i = 1; $i -lt $tableData.Count; $i++) {
                    $obj = New-Object PSObject
                    for ($j = 0; $j -lt $headers.Count; $j++) {
                        $obj | Add-Member -MemberType NoteProperty -Name $headers[$j] -Value $tableData[$i][$j]
                    }
                    $objects += $obj
                }
                
                # Izvezi u Excel
                $objects | Export-Excel -Path $path.FileName -AutoSize -TableStyle Medium2 -BoldTopRow
            }
        } else {
            # Fallback na COM objekt
            Export-ToExcelSuperSafe -htmlContent $html -fileName $path.FileName
        }
    } catch {
        Write-PCInfoLog "Excel export failed, trying CSV fallback" -Level "WARNING"
        $csvPath = $path.FileName -replace '\.xlsx?$', '.csv'
        Export-ToCSVEnhanced -htmlContent $html -fileName $csvPath
        [System.Windows.Forms.MessageBox]::Show(
            "Excel export nije uspio. Podaci su spremljeni kao CSV datoteka:`n$csvPath`n`nMožete otvoriti CSV datoteku u Excelu.",
            "Excel Export - CSV Fallback",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
}
                "html" {
                    # Spremi HTML s poboljšanim metadata
                    $enhancedHtml = Add-HTMLMetadata -originalHtml $html -tabName $tab.Text
                    [System.IO.File]::WriteAllText($path.FileName, $enhancedHtml, [System.Text.Encoding]::UTF8)
                }
                "doc" {
                    # Koristi vašu postojeæu Word funkciju
                    Export-ToWord -htmlContent $html -fileName $path.FileName
                }
                "pdf" {
                    # Koristi vašu postojeæu PDF funkciju
                    Export-ToPDF -htmlContent $html -fileName $path.FileName
                }
            }
            
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            Write-PCInfoLog "Export completed in $([math]::Round($duration, 2)) seconds"
            
            # Poboljšani success dialog
            $fileSize = if (Test-Path $path.FileName) { 
                "$([math]::Round((Get-Item $path.FileName).Length / 1KB, 2)) KB" 
            } else { 
                "Unknown" 
            }
            
            $response = [System.Windows.Forms.MessageBox]::Show(
                "Export completed successfully!`n`n" +
                "File: $($path.FileName)`n" +
                "Size: $fileSize`n" +
                "Duration: $([math]::Round($duration, 2))s`n" +
                "Format: $($format.ToUpper())`n`n" +
                "Would you like to open the file?",
                "Export Successful - Get-PCInfo v$Script:AppVersion",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process $path.FileName
                Write-PCInfoLog "Opened exported file: $($path.FileName)"
            }
            
        }
        catch {
            $errorMsg = "Export failed: $($_.Exception.Message)"
            Write-PCInfoLog $errorMsg -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        Write-PCInfoLog "Export cancelled by user"
    }
}

# === 6. RJEŠENJA ZA EXCEL COM PROBLEME ===
function Fix-ExcelCOMIssues {
    Write-Host "=== POKUŠAJ POPRAVKA EXCEL COM PROBLEMA ===" -ForegroundColor Yellow
    
    try {
        # 1. Zaustavi sve Excel procese
        Write-Host "Zaustavljam Excel procese..." -ForegroundColor Cyan
        Get-Process -Name "excel" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
        
        # 2. Oèisti COM cache
        Write-Host "Èistim COM cache..." -ForegroundColor Cyan
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # 3. Re-registriraj Excel COM komponente (zahtijeva admin)
        if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Host "Re-registriram Excel COM (admin privilegije)..." -ForegroundColor Cyan
            
            $excelPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE"
            if (-not (Test-Path $excelPath)) {
                $excelPath = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
            }
            
            if (Test-Path $excelPath) {
                Start-Process -FilePath $excelPath -ArgumentList "/regserver" -Wait -WindowStyle Hidden
                Write-Host "? Excel COM re-registriran" -ForegroundColor Green
            }
        } else {
            Write-Host "Za re-registraciju COM-a potrebne su admin privilegije" -ForegroundColor Yellow
        }
        
        Write-Host "Popravka završena, pokušajte ponovno" -ForegroundColor Green
        
    } catch {
        Write-Warning "Greška pri popravci: $_"
    }
}


# === Test funkcija za debugging (opcionalno) ===
function Test-ExcelInstallation {
    try {
        $excel = New-Object -ComObject Excel.Application
        $version = $excel.Version
        $excel.Quit()
        Write-Host "Excel je instaliran, verzija: $version" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Excel NIJE instaliran ili dostupan: $_" -ForegroundColor Red
        return $false
    }
}


# === Pomoæna funkcija za parsiranje HTML tablice s regex ===
function Parse-HTMLTableWithRegex {
    param([string]$html)
    
    try {
        $tableData = @()
        
        # Pronaði tablicu
        if ($html -match '<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Pronaði sve redove
            $rowMatches = [regex]::Matches($tableContent, '<tr[^>]*>(.*?)</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                $cells = @()
                
                # Pronaði sve æelije (th ili td)
                $cellMatches = [regex]::Matches($rowContent, '<t[hd][^>]*>(.*?)</t[hd]>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Oèisti HTML tagove i entitete
                    $cleanCell = Clean-HTMLContent -content $cellContent
                    $cells += $cleanCell
                }
                
                if ($cells.Count -gt 0) {
                    $tableData += ,$cells  # Dodaj kao array
                }
            }
        }
        
        return $tableData
    }
    catch {
        Write-Warning "Greška pri parsiranju tablice: $_"
        return @()
    }
}

# === Pomoæna funkcija za èišæenje HTML sadržaja ===
function Clean-HTMLContent {
    param([string]$content)
    
    try {
        # Ukloni HTML tagove
        $cleaned = $content -replace '<[^>]+>', ''
        
        # Ukloni HTML entitete
        $cleaned = $cleaned -replace '&nbsp;', ' '
        $cleaned = $cleaned -replace '&amp;', '&'
        $cleaned = $cleaned -replace '&lt;', '<'
        $cleaned = $cleaned -replace '&gt;', '>'
        $cleaned = $cleaned -replace '&quot;', '"'
        
        # Oèisti whitespace
        $cleaned = $cleaned -replace '\r', '' -replace '\n', ' ' -replace '\s+', ' '
        
        return $cleaned.Trim()
    }
    catch {
        return $content.ToString()
    }
}



function Format-UsersDataEnhanced {
    param([array]$tableData)
    
    $output = @()
    $output += "KORISNICKE INFORMACIJE:"
    $output += "-" * 30
    $output += ""
    
    if ($tableData.Count -gt 1) {
        $userCount = 1
        foreach ($user in $tableData[1..($tableData.Count-1)]) {
            if ($user.Count -ge 4) {
                $status = if ($user[2] -eq "Yes") { "AKTIVAN" } else { "NEAKTIVAN" }
                
                $output += "[$userCount] $($user[0])"
                $output += "   > Puno ime: $($user[1])"
                $output += "   > Status: $status"
                $output += "   > Zadnja lozinka: $($user[3])"
                
                if ($user.Count -gt 4) {
                    $output += "   > Zadnja prijava: $($user[4])"
                }
                if ($user.Count -gt 5) {
                    $output += "   > Istek racuna: $($user[5])"
                }
                if ($user.Count -gt 6) {
                    $disabled = if ($user[6] -eq "True") { "DA" } else { "NE" }
                    $output += "   > Onemogucen: $disabled"
                }
                if ($user.Count -gt 10) {
                    $output += "   > Grupe: $($user[10])"
                }
                
                $output += ""
                $userCount++
            }
        }
        
        # Statistike
        $activeUsers = ($tableData[1..($tableData.Count-1)] | Where-Object { $_.Count -gt 2 -and $_[2] -eq "Yes" }).Count
        $inactiveUsers = ($tableData[1..($tableData.Count-1)] | Where-Object { $_.Count -gt 2 -and $_[2] -eq "No" }).Count
        
        $output += "STATISTIKE:"
        $output += "-" * 15
        $output += "   Ukupno korisnika: $($tableData.Count - 1)"
        $output += "   Aktivnih: $activeUsers"
        $output += "   Neaktivnih: $inactiveUsers"
        $output += ""
    }
    
    return $output
}



function Add-HTMLMetadata {
    param([string]$originalHtml, [string]$tabName)
    
    # Dodaj metadata u HTML
    $metadata = @"
<!-- 
Generated by: Get-PCInfo Professional v$Script:AppVersion
Author: $Script:AppAuthor
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
System: $env:COMPUTERNAME
User: $env:USERNAME
Report: $tabName
-->
"@
    
    if ($originalHtml -match '<head>') {
        return $originalHtml -replace '<head>', "<head>`n$metadata"
    } else {
        return "$metadata`n$originalHtml"
    }
}



# === Pomoæna funkcija za formatiranje tablice kao tekst ===
function Format-TableAsText {
    param([array]$tableData)
    
    try {
        if ($tableData.Count -eq 0) {
            return @("Nema podataka u tablici")
        }

        # Provjeri da svi redovi imaju isti broj kolona
        $maxCols = 0
        foreach ($row in $tableData) {
            if ($row -is [array] -and $row.Count -gt $maxCols) {
                $maxCols = $row.Count
            }
        }

        if ($maxCols -eq 0) {
            return @("Greška u strukturi tablice")
        }

        # Normaliziraj sve redove da imaju isti broj kolona
        $normalizedData = @()
        foreach ($row in $tableData) {
            if ($row -is [array]) {
                $normalizedRow = @()
                for ($i = 0; $i -lt $maxCols; $i++) {
                    if ($i -lt $row.Count) {
                        $normalizedRow += $row[$i].ToString()
                    } else {
                        $normalizedRow += ""
                    }
                }
                $normalizedData += ,$normalizedRow
            }
        }

        # Dodaj broj redak kao prvu kolonu
        $headers = @("#") + $normalizedData[0]
        $dataRows = @()
        for ($i = 1; $i -lt $normalizedData.Count; $i++) {
            $dataRows += ,(@($i.ToString()) + $normalizedData[$i])
        }

        $allRows = @($headers) + $dataRows

        # Izraèunaj širinu kolona
        $colWidths = @(0) * $allRows[0].Count
        foreach ($row in $allRows) {
            for ($i = 0; $i -lt $row.Count; $i++) {
                if ($i -lt $colWidths.Count) {
                    $len = $row[$i].ToString().Length
                    if ($len -gt $colWidths[$i]) { 
                        $colWidths[$i] = $len 
                    }
                }
            }
        }

        # Generiraj tablicu
        $result = @()
        $line = ($colWidths | ForEach-Object { "-" * ($_ + 2) }) -join "+"
        $result += $line

        foreach ($i in 0..($allRows.Count - 1)) {
            $rowText = ""
            for ($j = 0; $j -lt $allRows[$i].Count; $j++) {
                if ($j -lt $colWidths.Count) {
                    $cellText = $allRows[$i][$j].ToString()
                    $rowText += "| " + $cellText.PadRight($colWidths[$j]) + " "
                }
            }
            $rowText += "|"
            $result += $rowText
            if ($i -eq 0) { $result += $line }
        }
        $result += $line

        return $result
    }
    catch {
        Write-Warning "Greška pri formatiranju tablice: $_"
        return @("Greška pri formatiranju tablice: $($_.Exception.Message)")
    }
}


# === Jednostavno dohvaæanje naslova ===
function Get-SimpleTitle {
    param([string]$html)
    
    # Pronaði h1 naslov
    if ($html -match '<h1[^>]*>(.*?)</h1>') {
        $title = $matches[1]
        # Ukloni HTML tagove
        $title = $title -replace '<[^>]+>', ''
        $title = $title -replace '&nbsp;', ' '
        $title = $title.Trim()
        return $title
    }
    
    return "Izvještaj o korisnièkim raèunima"
}

# === Jednostavno dohvaæanje opisa ===
function Get-SimpleDescription {
    param([string]$html)
    
    # Pronaði prvi paragraf
    if ($html -match '<p[^>]*>(.*?)</p>') {
        $desc = $matches[1]
        # Ukloni HTML tagove
        $desc = $desc -replace '<[^>]+>', ''
        $desc = $desc -replace '&nbsp;', ' '
        $desc = $desc.Trim()
        
        # Samo ako nije predugaèak
        if ($desc.Length -gt 5 -and $desc.Length -lt 150) {
            return $desc
        }
    }
    
    return $null
}

# === Konverzija HTML tablice u jednostavan tekst ===
function Convert-HTMLTableToSimpleText {
    param([string]$html)
    
    try {
        $result = @()
        
        # Jednostavno èišæenje HTML-a
        $cleanHtml = $html -replace '(?s)<style[^>]*>.*?</style>', ''
        $cleanHtml = $cleanHtml -replace '(?s)<script[^>]*>.*?</script>', ''
        
        # Pronaði tablicu
        if ($cleanHtml -match '(?s)<table[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Pronaði sve redove
            $rowPattern = '(?s)<tr[^>]*>(.*?)</tr>'
            $rowMatches = [regex]::Matches($tableContent, $rowPattern)
            
            $rowNumber = 1
            foreach ($rowMatch in $rowMatches) {
                $rowContent = $rowMatch.Groups[1].Value
                
                # Pronaði æelije
                $cellPattern = '(?s)<t[hd][^>]*>(.*?)</t[hd]>'
                $cellMatches = [regex]::Matches($rowContent, $cellPattern)
                
                $rowText = ""
                $cellNumber = 1
                
                # Dodaj redni broj na poèetak (osim za header)
                if ($rowNumber -gt 1) {
                    $rowText += "$($rowNumber - 1). "
                } else {
                    $rowText += "    "  # Razmak za header
                }
                
                foreach ($cellMatch in $cellMatches) {
                    $cellContent = $cellMatch.Groups[1].Value
                    
                    # Èisti sadržaj æelije
                    $cleanCell = $cellContent -replace '<[^>]+>', ''
                    $cleanCell = $cleanCell -replace '&nbsp;', ' '
                    $cleanCell = $cleanCell -replace '&amp;', '&'
                    $cleanCell = $cleanCell -replace '&lt;', '<'
                    $cleanCell = $cleanCell -replace '&gt;', '>'
                    $cleanCell = $cleanCell -replace '&quot;', '"'
                    $cleanCell = $cleanCell -replace '\s+', ' '
                    $cleanCell = $cleanCell.Trim()
                    
                    # Ogranièi duljinu
                    if ($cleanCell.Length -gt 30) {
                        $cleanCell = $cleanCell.Substring(0, 27) + "..."
                    }
                    
                    # Dodaj æeliju s razmakom
                    $rowText += $cleanCell.PadRight(32)
                }
                
                # Dodaj red u rezultat
                if ($rowText.Trim() -ne "") {
                    $result += $rowText.TrimEnd()
                    
                    # Dodaj separator nakon header reda
                    if ($rowNumber -eq 1) {
                        $result += ("-" * 80)
                    }
                }
                
                $rowNumber++
            }
        } else {
            $result += "Nema tablice u HTML sadržaju"
        }
        
        return $result
        
    } catch {
        Write-Warning "Greška pri konverziji tablice: $_"
        return @("Greška pri obradi tablice")
    }
}

# === ALTERNATIVNO - ULTRA JEDNOSTAVAN TXT EXPORT ===
function Export-TxtUltraSimple {
    param (
        [string]$FilePath,
        $control
    )

    try {
        Write-Host "Kreiram ultra jednostavan TXT..." -ForegroundColor Yellow
        
        # Dohvati HTML
        $html = ""
        if ($control -is [System.Windows.Forms.WebBrowser]) {
            $html = $control.DocumentText
        } elseif ($control -is [System.Windows.Forms.RichTextBox]) {
            $html = $control.Text
        }

        # Ultra jednostavno - samo ukloni HTML tagove
        $cleanText = $html -replace '<style[^>]*>.*?</style>', ''
        $cleanText = $cleanText -replace '<script[^>]*>.*?</script>', ''
        $cleanText = $cleanText -replace '<[^>]+>', ' '
        $cleanText = $cleanText -replace '&nbsp;', ' '
        $cleanText = $cleanText -replace '&amp;', '&'
        $cleanText = $cleanText -replace '&lt;', '<'
        $cleanText = $cleanText -replace '&gt;', '>'
        $cleanText = $cleanText -replace '&quot;', '"'
        $cleanText = $cleanText -replace '\s+', ' '
        $cleanText = $cleanText.Trim()
        
        # Podijeli u linije za èitljivost
        $lines = @()
        $words = $cleanText -split '\s+'
        $currentLine = ""
        
        foreach ($word in $words) {
            if (($currentLine + " " + $word).Length -le 80) {
                if ($currentLine -eq "") {
                    $currentLine = $word
                } else {
                    $currentLine += " " + $word
                }
            } else {
                if ($currentLine -ne "") {
                    $lines += $currentLine
                }
                $currentLine = $word
            }
        }
        
        if ($currentLine -ne "") {
            $lines += $currentLine
        }
        
        # Dodaj footer
        $lines += ""
        $lines += "Generirano: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
        $lines += "Sustav: $env:COMPUTERNAME"
        
        # Spremi
        $lines | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        Write-Host "? Ultra jednostavan TXT export završen" -ForegroundColor Green
        
    } catch {
        Write-Error "Greška u ultra jednostavnom TXT exportu: $_"
        throw
    }
}







function Export-HTMLTab {
    param([string]$content, [string]$fileName)

    $html = @"
<!DOCTYPE html>
<html>
<head><meta charset='UTF-8'><title>Export</title></head>
<body><pre>$content</pre></body>
</html>
"@
    Set-Content -Path $fileName -Value $html -Encoding UTF8
}










# === Funkcija: Lokalni korisnici ===
function Generate-UserReport {
    param([string]$OutputPath)

    $users = Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount='True'"
    $userInfo = @()
    $disabledCount = 0
    $adminCount = 0

    foreach ($user in $users) {
        $userDetails = net user $user.Name
        $groups = net user $user.Name | Where-Object { $_ -match "Local Group Memberships" } | ForEach-Object { ($_ -split "\s{2,}")[1].Trim() }
        $fullName = "N/A"
        $accountActive = "N/A"
        $lastPasswordSet = "N/A"
        $lastLogon = "N/A"
        $accountExpiry = "N/A"
        $passwordExpires = "N/A"
        $passwordChangeable = "N/A"
        $passwordRequired = "N/A"
        $userMayChangePassword = "N/A"
        $accountDisabled = $user.Disabled
        $SID = $user.SID
        $isAdmin = $false

        if ($accountDisabled -eq $true) { $disabledCount++ }
        if ($groups -and ($groups -match "Administrators|Administratori")) {
            $isAdmin = $true
            $adminCount++
        }

        foreach ($line in $userDetails) {
            if ($line -match "Full Name\s+(.+)") { $fullName = $matches[1].Trim() }
            if ($line -match "Account active\s+(.+)") { $accountActive = $matches[1].Trim() }
            if ($line -match "Password last set\s+(.+)") { $lastPasswordSet = $matches[1].Trim() }
            if ($line -match "Last logon\s+(.+)") { $lastLogon = $matches[1].Trim() }
            if ($line -match "Account expires\s+(.+)") { $accountExpiry = $matches[1].Trim() }
            if ($line -match "Password expires\s+(.+)") { $passwordExpires = $matches[1].Trim() }
            if ($line -match "Password required\s+(.+)") { $passwordRequired = $matches[1].Trim() }
            if ($line -match "Password changeable\s+(.+)") { $passwordChangeable = $matches[1].Trim() }
            if ($line -match "User may change password\s+(.+)") { $userMayChangePassword = $matches[1].Trim() }
        }

        $userInfo += [PSCustomObject]@{
            Name                = $user.Name
            FullName            = $fullName
            AccountActive       = $accountActive
            LastPasswordSet     = $lastPasswordSet
            LastLogon           = $lastLogon
            AccountExpiry       = $accountExpiry
            PasswordExpires     = $passwordExpires
            PasswordChangeable  = $passwordChangeable
            PasswordRequired    = $passwordRequired
            UserMayChangePassword = $userMayChangePassword
            AccountDisabled     = $accountDisabled
            SID                 = $SID
            Groups              = ($groups -join ", ")
            IsAdmin             = $isAdmin
        }
    }

    $i = 1
    $rows = $userInfo | ForEach-Object {
        $class = ""
        $adminBadge = ""
        $statusBadge = ""

        if ($_.AccountDisabled -eq $true) {
            $class = "user-disabled"
            $statusBadge = "<span class='status-badge status-disabled'>ONEMOGUÆEN</span>"
        } else {
            $statusBadge = "<span class='status-badge status-active'>AKTIVAN</span>"
        }

        if ($_.IsAdmin) {
            $adminBadge = "<span class='admin-badge'>ADMIN</span>"
            if ($class -eq "") { $class = "admin-user" }
        }

        # Kreiraj obojeni prikaz za Password expires
        $passwordExpiresFormatted = ""
        $expiryClass = ""
        if ($_.PasswordExpires -match "Never") {
            $expiryClass = "expiry-never"
            $passwordExpiresFormatted = "<span class='$expiryClass'>$($_.PasswordExpires)</span>"
        } elseif ($_.PasswordExpires -ne "N/A" -and $_.PasswordExpires -ne "No") {
            try {
                $expiryDate = [DateTime]::Parse($_.PasswordExpires)
                $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
                if ($daysUntilExpiry -le 7) {
                    $expiryClass = "expiry-critical"
                } elseif ($daysUntilExpiry -le 14) {
                    $expiryClass = "expiry-warning"
                } else {
                    $expiryClass = "expiry-ok"
                }
                $passwordExpiresFormatted = "<span class='$expiryClass'>$($_.PasswordExpires)</span>"
            } catch {
                $expiryClass = "expiry-unknown"
                $passwordExpiresFormatted = "<span class='$expiryClass'>$($_.PasswordExpires)</span>"
            }
        } else {
            $expiryClass = "expiry-unknown"
            $passwordExpiresFormatted = "<span class='$expiryClass'>$($_.PasswordExpires)</span>"
        }

        # Kreiraj obojeni prikaz za Account expires
        $accountExpiryFormatted = ""
        $accountExpiryClass = ""
        if ($_.AccountExpiry -match "Never") {
            $accountExpiryClass = "expiry-never"
            $accountExpiryFormatted = "<span class='$accountExpiryClass'>$($_.AccountExpiry)</span>"
        } elseif ($_.AccountExpiry -ne "N/A" -and $_.AccountExpiry -ne "No") {
            try {
                $expiryDate = [DateTime]::Parse($_.AccountExpiry)
                $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
                if ($daysUntilExpiry -le 7) {
                    $accountExpiryClass = "expiry-critical"
                } elseif ($daysUntilExpiry -le 14) {
                    $accountExpiryClass = "expiry-warning"
                } else {
                    $accountExpiryClass = "expiry-ok"
                }
                $accountExpiryFormatted = "<span class='$accountExpiryClass'>$($_.AccountExpiry)</span>"
            } catch {
                $accountExpiryClass = "expiry-unknown"
                $accountExpiryFormatted = "<span class='$accountExpiryClass'>$($_.AccountExpiry)</span>"
            }
        } else {
            $accountExpiryClass = "expiry-unknown"
            $accountExpiryFormatted = "<span class='$accountExpiryClass'>$($_.AccountExpiry)</span>"
        }

        $passwordInfo = @"
<div class='pwd-info'>
    <span><b>Account expires:</b> $accountExpiryFormatted</span>
    <span><b>Password last set:</b> $($_.LastPasswordSet)</span>
    <span><b>Password expires:</b> $passwordExpiresFormatted</span>
    <span><b>Password changeable:</b> $($_.PasswordChangeable)</span>
    <span><b>Password required:</b> $($_.PasswordRequired)</span>
    <span><b>User may change password:</b> $($_.UserMayChangePassword)</span>
</div>
"@

        "<tr class='$class'><td>$i</td><td>$($_.Name) $adminBadge</td><td>$($_.FullName)</td><td>$statusBadge</td><td>$passwordInfo</td><td>$($_.LastLogon)</td><td><code>$($_.SID)</code></td><td>$($_.Groups)</td></tr>"
        $i++
    } | Out-String

$html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Lokalni korisnièki raèuni</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 10px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.subtitle {
    color: rgba(255,255,255,0.9);
    font-size: 16px;
    margin-bottom: 30px;
    text-align: center;
    line-height: 1.5;
}

.container {
    width: 100%;
    max-width: 1600px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 700px; 
    overflow: auto; 
    margin: 20px;
    border-radius: 10px;
    position: relative;
    background: white;
}

.table-wrapper::-webkit-scrollbar {
    width: 12px;
    height: 12px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-corner {
    background: #f1f1f1;
}

table { 
    width: 100%; 
    min-width: 1290px;
    border-collapse: collapse;
    font-size: 13px;
    position: relative;
    table-layout: fixed;
}

/* Poboljšani sticky header */
thead {
    position: -webkit-sticky;
    position: sticky;
    top: 0;
    z-index: 99;
    background: linear-gradient(135deg, #667eea, #764ba2);
}

thead th { 
    position: -webkit-sticky;
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2) !important;
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 100;
    box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    border-bottom: 2px solid rgba(255,255,255,0.3);
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 10px 8px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
    word-wrap: break-word;
    overflow: hidden;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

tr.user-disabled td {
    background-color: #fef2f2 !important;
}

tr.user-disabled:hover td {
    background-color: #fee2e2 !important;
}

tr.admin-user td {
    background-color: #fef3c7 !important;
}

tr.admin-user:hover td {
    background-color: #fde68a !important;
}

/* Poboljšani sticky prvi stupac */
td:first-child, th:first-child { 
    position: -webkit-sticky;
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5 !important;
    font-weight: 600;
    color: #667eea;
    width: 40px;
    text-align: center;
    z-index: 101;
    box-shadow: 2px 0 4px rgba(0,0,0,0.2);
}

tr:hover td:first-child {
    background-color: #e8eaf6;
    box-shadow: 2px 0 4px rgba(0,0,0,0.15);
}

tr.user-disabled td:first-child {
    background-color: #fee2e2 !important;
    box-shadow: 2px 0 4px rgba(0,0,0,0.1);
}

tr.admin-user td:first-child {
    background-color: #fde68a !important;
    box-shadow: 2px 0 4px rgba(0,0,0,0.1);
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9) !important;
    z-index: 102;
    box-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    position: -webkit-sticky;
    position: sticky;
    left: 0;
    top: 0;
}

/* Status badges */
.status-badge {
    padding: 3px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
}

.status-active {
    background: linear-gradient(135deg, #10b981, #059669);
    color: white;
}

.status-disabled {
    background: linear-gradient(135deg, #ef4444, #dc2626);
    color: white;
}

.admin-badge {
    background: linear-gradient(135deg, #f59e0b, #d97706);
    color: white;
    padding: 2px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 8px;
}

/* Password info stilovi */
.pwd-info {
    display: flex;
    flex-direction: column;
    gap: 4px;
}

.pwd-info span {
    background: #f3f4f6;
    padding: 3px 6px;
    border-radius: 4px;
    font-size: 11px;
    color: #6b7280;
}

.pwd-info span:nth-child(1),
.pwd-info span:nth-child(3) {
    background: #ffffff;
    border: 1px solid #d1d5db;
    font-weight: 600;
}

.pwd-info b {
    color: #1f2937;
    font-weight: 600;
}

/* Stilovi za istijek lozinke i raèuna */
.expiry-critical {
    color: #dc2626 !important;
    font-weight: 600 !important;
    background-color: #fee2e2 !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
    font-size: 10px !important;
}

.expiry-warning {
    color: #d97706 !important;
    font-weight: 600 !important;
    background-color: #fef3c7 !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
    font-size: 10px !important;
}

.expiry-ok {
    color: #ffffff !important;
    font-weight: 700 !important;
    background-color: #059669 !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
    font-size: 10px !important;
}

.expiry-never {
    color: #1f2937 !important;
    font-weight: 700 !important;
    background-color: #e5e7eb !important;
    padding: 2px 6px !important;
    border-radius: 4px !important;
    font-size: 10px !important;
    border: 1px solid #9ca3af !important;
}

.expiry-unknown {
    color: #9ca3af !important;
    font-style: italic !important;
    font-size: 10px !important;
}

code {
    background: #f1f5f9;
    padding: 2px 6px;
    border-radius: 4px;
    font-family: 'Courier New', monospace;
    font-size: 11px;
    color: #475569;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Optimizacija širine stupaca prema sadržaju */
th:nth-child(1), td:nth-child(1) { width: 40px; min-width: 40px; max-width: 40px; }
th:nth-child(2), td:nth-child(2) { width: 120px; min-width: 100px; max-width: 150px; }
th:nth-child(3), td:nth-child(3) { width: 140px; min-width: 120px; max-width: 180px; }
th:nth-child(4), td:nth-child(4) { width: 130px; min-width: 120px; max-width: 150px; }
th:nth-child(5), td:nth-child(5) { width: 340px; min-width: 320px; max-width: 360px; }
th:nth-child(6), td:nth-child(6) { width: 130px; min-width: 110px; max-width: 160px; }
th:nth-child(7), td:nth-child(7) { width: 310px; min-width: 290px; max-width: 330px; }
th:nth-child(8), td:nth-child(8) { width: 100px; min-width: 80px; max-width: 120px; }

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px 6px; font-size: 11px; }
    .stats-box { padding: 15px 25px; }
    .pwd-info { gap: 2px; }
    .pwd-info span { font-size: 10px; }
    .container { max-width: 100%; }
}
</style>
</head>
<body>
    <h1>Status lokalnih korisnièkih raèuna</h1>
    <p class="subtitle">Crvena boja oznaèava neaktivne (onemoguæene) raèune, žuta administratore.<br>Oznaèavanje rokova valjanosti bojom: <span class="expiry-critical">Kritièno (7 dana)</span>, <span class="expiry-warning">Upozorenje (14 dana)</span>, <span class="expiry-ok">Uredu (>14 dana)</span>, <span class="expiry-never">Nikad</span></p>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Korisnièko ime</th>
                        <th>Puno ime</th>
                        <th>Status</th>
                        <th>Lozinka</th>
                        <th>Zadnja prijava</th>
                        <th>SID</th>
                        <th>Grupe</th>
                    </tr>
                </thead>
                <tbody>
                    $rows
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno korisnika:</b> $($userInfo.Count)</p>
                <p><b>Onemoguæenih:</b> $disabledCount</p>
                <p><b>Administratora:</b> $adminCount</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}


# === Funkcija: Grupe sa korisnicima ===
function Generate-GroupUserReport {
    param([string]$OutputPath)

    $localGroups = Get-LocalGroup | Sort-Object Name
    $totalMembers = 0
    $activeGroups = 0

    $groupsHtml = ""
    foreach ($group in $localGroups) {
        $groupMembers = Get-LocalGroupMember -Group $group.Name -ErrorAction SilentlyContinue | Sort-Object Name

        if (!$groupMembers -or $groupMembers.Count -eq 0) {
            $groupsHtml += @"
<div class="group-section empty-group">
    <div class="group-header">
        <h2>$($group.Name)</h2>
        <div class="member-count">0 èlanova</div>
    </div>
    <div class="empty-message">
        <p>Ova grupa nema èlanova.</p>
    </div>
</div>
"@
        } else {
            $activeGroups++
            $totalMembers += $groupMembers.Count
            
            $groupsHtml += @"
<div class="group-section">
    <div class="group-header">
        <h2>$($group.Name)</h2>
        <div class="member-count">$($groupMembers.Count) èlanova</div>
    </div>
    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th>#</th>
                    <th>Korisnièko ime</th>
                    <th>Puno ime</th>
                    <th>Tip</th>
                    <th>Status</th>
                    <th>SID</th>
                    <th>Kreiran</th>
                    <th>Zadnja prijava</th>
                    <th>Lozinka istekla</th>
                </tr>
            </thead>
            <tbody>
"@

            $i = 1
            foreach ($member in $groupMembers) {
                $userName = $member.Name -replace '^.*\\', ''
                $fullName = ""
                $type = $member.ObjectClass
                $status = "Nepoznato"
                $statusClass = "status-unknown"
                $sid = $member.SID.Value
                $created = "-"
                $lastLogon = "-"
                $pwdExpired = "-"
                $rowClass = ""

                if ($type -eq "User") {
                    $localUser = Get-LocalUser -Name $userName -ErrorAction SilentlyContinue
                    if ($localUser) {
                        $fullName = if ($localUser.FullName) { $localUser.FullName } else { "<em>Nema podataka</em>" }
                        $status = if ($localUser.Enabled) { "Aktivan" } else { "Neaktivan" }
                        $statusClass = if ($localUser.Enabled) { "status-active" } else { "status-disabled" }
                        
                        if (-not $localUser.Enabled) { 
                            $rowClass = "member-inactive" 
                        }
                        
                        $lastLogon = if ($localUser.LastLogon) { 
                            $localUser.LastLogon.ToString("dd.MM.yyyy HH:mm") 
                        } else { 
                            "<em>Nikad</em>" 
                        }
                        
                        $pwdExpired = if ($localUser.PasswordExpired) { "Da" } else { "Ne" }

                        $netUser = net user $userName 2>$null
                        $createdMatch = ($netUser | Where-Object { $_ -match 'Password last set' }) -split '\s{2,}' | Select-Object -Last 1
                        if ($createdMatch) {
                            $created = $createdMatch.Trim()
                        }
                    } else {
                        $fullName = "<em>Domenski ili servisni korisnik</em>"
                        $status = "Nepoznato"
                        $statusClass = "status-unknown"
                        $rowClass = "member-domain"
                    }
                } else {
                    $fullName = "<em>Servisni raèun ili grupa</em>"
                    $status = "N/A"
                    $statusClass = "status-unknown"
                    $rowClass = "member-domain"
                }

                $statusBadge = switch ($statusClass) {
                    "status-active" { "<span class='status-badge status-active'>AKTIVAN</span>" }
                    "status-disabled" { "<span class='status-badge status-disabled'>NEAKTIVAN</span>" }
                    default { "<span class='status-badge status-unknown'>NEPOZNATO</span>" }
                }

                $groupsHtml += @"
                <tr class='$rowClass'>
                    <td>$i</td>
                    <td>$userName</td>
                    <td>$fullName</td>
                    <td>$type</td>
                    <td>$statusBadge</td>
                    <td><code>$sid</code></td>
                    <td>$created</td>
                    <td>$lastLogon</td>
                    <td>$pwdExpired</td>
                </tr>
"@
                $i++
            }

            $groupsHtml += @"
            </tbody>
        </table>
    </div>
</div>
"@
        }
    }

$htmlContent = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Lokalne grupe i korisnici</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    padding: 40px 20px;
    color: #1f2937;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.container {
    max-width: 1400px;
    margin: 0 auto;
}

.group-section {
    margin-bottom: 30px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.2);
    overflow: hidden;
    backdrop-filter: blur(10px);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.group-section:hover {
    transform: translateY(-5px);
    box-shadow: 0 25px 70px rgba(0,0,0,0.3);
}

.group-header {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 30px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.group-header h2 {
    font-size: 20px;
    font-weight: 600;
    margin: 0;
}

.member-count {
    background: rgba(255, 255, 255, 0.2);
    padding: 8px 16px;
    border-radius: 20px;
    font-size: 14px;
    font-weight: 500;
}

.table-container {
    overflow-x: auto;
    max-height: 500px;
    overflow-y: auto;
}

.table-container::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

.table-container::-webkit-scrollbar-track {
    background: #f1f1f1;
}

.table-container::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 4px;
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}

th {
    background: #f8f9fa;
    color: #1f2937;
    font-weight: 600;
    padding: 12px 15px;
    text-align: left;
    border-bottom: 2px solid #e9ecef;
    position: sticky;
    top: 0;
    z-index: 10;
}

td {
    padding: 12px 15px;
    border-bottom: 1px solid #e9ecef;
    vertical-align: top;
}

tbody tr {
    transition: background-color 0.3s ease;
}

tbody tr:nth-child(even) {
    background-color: #f8f9fa;
}

tbody tr:hover {
    background-color: #e8eaf6;
}

tr.member-inactive td {
    background-color: #fef2f2 !important;
}

tr.member-inactive:hover td {
    background-color: #fee2e2 !important;
}

tr.member-domain td {
    background-color: #fef3c7 !important;
    font-style: italic;
}

tr.member-domain:hover td {
    background-color: #fde68a !important;
}

td:first-child, th:first-child {
    position: sticky;
    left: 0;
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 50px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

tr.member-inactive td:first-child {
    background-color: #fee2e2 !important;
}

tr.member-domain td:first-child {
    background-color: #fde68a !important;
}

/* Status badges */
.status-badge {
    padding: 3px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
}

.status-active {
    background: linear-gradient(135deg, #10b981, #059669);
    color: white;
}

.status-disabled {
    background: linear-gradient(135deg, #ef4444, #dc2626);
    color: white;
}

.status-unknown {
    background: linear-gradient(135deg, #f59e0b, #d97706);
    color: white;
}

code {
    background: #f1f5f9;
    padding: 2px 6px;
    border-radius: 4px;
    font-family: 'Courier New', monospace;
    font-size: 11px;
    color: #475569;
}

/* Empty group styling */
.empty-group .group-header {
    background: linear-gradient(135deg, #9ca3af, #6b7280);
}

.empty-message {
    padding: 40px;
    text-align: center;
    color: #6b7280;
    font-style: italic;
    background: #f9fafb;
}

.stats-container {
    display: flex;
    justify-content: center;
    margin-top: 30px;
}

.stats-box {
    background: rgba(255, 255, 255, 0.95);
    color: #1f2937;
    padding: 25px 40px;
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.2);
    text-align: center;
    backdrop-filter: blur(10px);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 25px 70px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #667eea;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.group-section {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    body { padding: 20px 10px; }
    h1 { font-size: 24px; }
    .group-header { 
        flex-direction: column; 
        gap: 10px; 
        text-align: center; 
        padding: 15px 20px;
    }
    .group-header h2 { font-size: 18px; }
    td, th { padding: 8px 10px; font-size: 11px; }
    .stats-box { padding: 20px 25px; }
}
</style>
</head>
<body>
    <div class="container">
        <h1>Korisnièke grupe s èlanstvom</h1>
        $groupsHtml
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno grupa:</b> $($localGroups.Count)</p>
                <p><b>Grupa s èlanovima:</b> $activeGroups</p>
                <p><b>Ukupno èlanova:</b> $totalMembers</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $OutputPath -Encoding utf8
}







# === Funkcija: Logiranja i odjave ===
function Get-LogonHistory {
    param (
        [int]$Days = 30
    )

    Write-Host "Dohvaæam logon/logoff dogaðaje za zadnjih $Days dana..." -ForegroundColor Green
    $startTime = (Get-Date).AddDays(-$Days)
    $results = @()

    # 1. Pokušaj s Security logom (glavni izvor)
    try {
        Write-Host "Analiziram Security log..." -ForegroundColor Yellow
        $securityParams = @{
            LogName = "Security"
            StartTime = $startTime
            FilterXPath = "*[System[EventID=4624 or EventID=4634 or EventID=4647]]"
            ErrorAction = "SilentlyContinue"
        }
        
        $securityEvents = Get-WinEvent @securityParams
        Write-Host "Pronaðeno $($securityEvents.Count) Security dogaðaja" -ForegroundColor Cyan
        
        foreach ($event in $securityEvents) {
            $xml = [xml]$event.ToXml()
            $eventData = @{}
            
            foreach ($data in $xml.Event.EventData.Data) {
                if ($data.Name) {
                    $eventData[$data.Name] = $data.'#text'
                }
            }
            
            $logonType = switch ($event.Id) {
                4624 { "Logon" }
                4634 { "LogOff" }
                4647 { "LogOff" }
            }
            
            $username = $eventData['TargetUserName']
            $domain = $eventData['TargetDomainName']
            $logonTypeNum = [int]$eventData['LogonType']
            
            # Filtriranje sistemskih raèuna
            if ($username -in @('SYSTEM', 'ANONYMOUS LOGON', 'LOCAL SERVICE', 'NETWORK SERVICE') -or 
                $username -like '*$' -or $username -eq '') {
                continue
            }
            
            # Za uspješne prijave, filtriramo tipove (samo interaktivne)
            if ($event.Id -eq 4624 -and $logonTypeNum -notin @(2, 10, 11)) {
                continue
            }
            
            $fullUser = if ($domain -and $domain -ne $env:COMPUTERNAME -and $domain -ne '') {
                "$domain\$username"
            } else {
                $username
            }
            
            $results += [PSCustomObject]@{
                Time = $event.TimeCreated
                EventType = $logonType
                User = $fullUser
            }
        }
    } catch {
        Write-Warning "Security log nedostupan: $($_.Exception.Message)"
    }

    # 2. Fallback na System log (vaš originalni pristup)
    if ($results.Count -eq 0) {
        Write-Host "Pokušavam s System logom..." -ForegroundColor Yellow
        $params = @{
            LogName = "System"
            Source = "Microsoft-Windows-WinLogon"
            After = $startTime
            ErrorAction = "SilentlyContinue"
        }

        try {
            $events = Get-EventLog @params
            Write-Host "Pronaðeno $($events.Count) System dogaðaja" -ForegroundColor Cyan

            foreach ($event in $events) {
                $logonType = switch ($event.InstanceId) {
                    7001 { "Logon" }
                    7002 { "LogOff" }
                    default { continue }
                }

                if ($event.ReplacementStrings -and $event.ReplacementStrings.Count -gt 1) {
                    $userSID = $event.ReplacementStrings[1]
                    try {
                        $ntAccount = (New-Object System.Security.Principal.SecurityIdentifier $userSID).Translate([System.Security.Principal.NTAccount])
                        $user = $ntAccount.Value.Split('\')[-1]
                    } catch {
                        $user = "Nepoznato"
                    }
                } else {
                    $user = "System-User"
                }

                $results += [PSCustomObject]@{
                    Time      = $event.TimeWritten
                    EventType = $logonType
                    User      = $user
                }
            }
        } catch {
            Write-Warning "System log nedostupan: $($_.Exception.Message)"
        }
    }

    # 3. Dodatni WMI pristup ako ništa ne radi
    if ($results.Count -eq 0) {
        Write-Host "Pokušavam WMI pristup..." -ForegroundColor Yellow
        try {
            $wmiFilter = "LogFile='Security' AND EventCode IN (528, 540, 538, 4624, 4634)"
            $wmiEvents = Get-CimInstance -ClassName Win32_NTLogEvent -Filter $wmiFilter -ErrorAction SilentlyContinue |
                        Where-Object { $_.TimeGenerated -ge $startTime }
            
            Write-Host "WMI pronašao $($wmiEvents.Count) dogaðaja" -ForegroundColor Cyan
            
            foreach ($event in $wmiEvents) {
                $eventType = switch ($event.EventCode) {
                    528 { "Logon" }
                    540 { "Logon" }
                    538 { "LogOff" }
                    4624 { "Logon" }
                    4634 { "LogOff" }
                }
                
                $username = if ($event.InsertionStrings -and $event.InsertionStrings.Count -gt 0) {
                    $event.InsertionStrings[0]
                } else {
                    "WMI-User"
                }
                
                $results += [PSCustomObject]@{
                    Time = $event.TimeGenerated
                    EventType = $eventType
                    User = $username
                }
            }
        } catch {
            Write-Warning "WMI pristup neuspješan: $($_.Exception.Message)"
        }
    }

    Write-Host "Ukupno pronaðeno: $($results.Count) dogaðaja" -ForegroundColor Green
    Write-Host "- Prijave: $(($results | Where-Object {$_.EventType -eq 'Logon'}).Count)" -ForegroundColor Green
    Write-Host "- Odjave: $(($results | Where-Object {$_.EventType -eq 'LogOff'}).Count)" -ForegroundColor Green

    return $results | Where-Object { $_.Time -is [datetime] -and $_.Time -le (Get-Date) -and $_.Time -ge $startTime } | Sort-Object Time -Descending
}

function Generate-LogonHTMLReport {
    param (
        [array]$LogonData,
        [string]$ReportPath
    )

    Write-Host "Generiranje HTML izvještaja..." -ForegroundColor Green

    # Podaci za statistike
    $byDay = $LogonData | Where-Object { $_.EventType -eq "Logon" } | Group-Object { $_.Time.Date } | Sort-Object Name
    $totalLogons = ($LogonData | Where-Object { $_.EventType -eq "Logon" }).Count
    $totalLogoffs = ($LogonData | Where-Object { $_.EventType -eq "LogOff" }).Count
    $uniqueUsers = ($LogonData | Select-Object -ExpandProperty User -Unique).Count

    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Logon/Logoff Povijest</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 10px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.subtitle {
    color: rgba(255,255,255,0.9);
    font-size: 16px;
    text-align: center;
    margin-bottom: 30px;
    font-weight: 300;
}

h2 {
    color: #1e3a8a;
    font-size: 24px;
    margin: 40px 0 20px 0;
    text-align: center;
}

.summary-cards {
    display: flex;
    justify-content: center;
    gap: 20px;
    max-width: 800px;
    margin: 0 auto 30px auto;
    flex-wrap: wrap;
}

.summary-card {
    background: rgba(255, 255, 255, 0.9);
    border-radius: 15px;
    padding: 20px;
    text-align: center;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    transition: transform 0.3s ease;
    min-width: 200px;
    flex: 1;
}

.summary-card:hover {
    transform: translateY(-5px);
}

.summary-card .stat-line {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
}

.summary-card .number {
    font-size: 36px;
    font-weight: 700;
    color: #667eea;
}

.summary-card .label {
    color: #6b7280;
    font-size: 14px;
    font-weight: 500;
}

.container {
    max-width: 1200px;
    margin: 0 auto;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 60px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Date header */
tr.date-header td {
    background: linear-gradient(90deg, #667eea, #764ba2) !important;
    color: white !important;
    font-weight: 700;
    font-size: 15px;
    padding: 15px;
    text-align: left;
    position: sticky;
    top: 48px;
    z-index: 9;
}

/* Event type stilovi */
.event-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
}

.event-badge.logon {
    background-color: #d1fae5;
    color: #065f46;
    border: 1px solid #10b981;
}

.event-badge.logoff {
    background-color: #fee2e2;
    color: #991b1b;
    border: 1px solid #ef4444;
}

.event-icon {
    width: 8px;
    height: 8px;
    border-radius: 50%;
}

.event-icon.logon {
    background-color: #10b981;
}

.event-icon.logoff {
    background-color: #ef4444;
}

/* Timeline vrijeme */
.time-display {
    font-family: 'Courier New', monospace;
    font-weight: 600;
}

/* User display */
.user-name {
    font-weight: 500;
    color: #374151;
}

/* Row striping */
tbody tr:nth-child(even) td {
    background-color: #f8fafc;
}

/* Boje za razlièite datume */
tr.date-gray td { 
    background-color: #f9fafb; 
}

tr.date-white td { 
    background-color: #ffffff; 
}

/* Statistike tablice */
.stats-tables {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 20px;
    padding: 20px;
    max-width: 900px;
    margin: 0 auto;
}

.stat-table {
    background: white;
    border-radius: 15px;
    overflow: hidden;
    box-shadow: 0 5px 15px rgba(0,0,0,0.1);
}

.stat-table table {
    width: 100%;
}

.stat-table th {
    background: #f3f4f6;
    color: #1f2937;
    padding: 12px;
    font-weight: 600;
}

.stat-table td {
    padding: 10px 12px;
}

.stat-table tr:nth-child(even) td {
    background-color: #f9fafb;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    margin: 30px auto;
    max-width: 600px;
}

.stats-box .stat-line {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 15px;
    margin: 8px 0;
    font-size: 14px;
    padding: 0 10px;
}

.stats-box .stat-line b {
    color: #fff;
    font-weight: 600;
    flex-shrink: 0;
}

.stats-box .stat-line span {
    color: #fff;
    text-align: right;
    flex-grow: 1;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .summary-cards { 
        flex-direction: column; 
        align-items: center;
    }
    .summary-card {
        min-width: 250px;
    }
    .summary-card .stat-line {
        flex-direction: column;
        gap: 5px;
    }
    .stats-tables { grid-template-columns: 1fr; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box .stat-line {
        justify-content: center;
        flex-direction: column;
        gap: 5px;
        text-align: center;
    }
    .stats-box .stat-line span {
        text-align: center;
    }
}
</style>
</head>
<body>
    <h1>Zapis aktivnosti korisnika</h1>
    <div class='subtitle'>Chronološki prikaz prijava i odjava (najnovije prvo)</div>
    
    <div class='summary-cards'>
        <div class='summary-card'>
            <div class='stat-line'>
                <div class='number'>$totalLogons</div>
                <div class='label'>Ukupno prijava</div>
            </div>
        </div>
        <div class='summary-card'>
            <div class='stat-line'>
                <div class='number'>$totalLogoffs</div>
                <div class='label'>Ukupno odjava</div>
            </div>
        </div>
        <div class='summary-card'>
            <div class='stat-line'>
                <div class='number'>$uniqueUsers</div>
                <div class='label'>Jedinstvenih korisnika</div>
            </div>
        </div>
    </div>

    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Vrijeme</th>
                        <th>Vrsta</th>
                        <th>Korisnik</th>
                    </tr>
                </thead>
                <tbody>
"@

    $i = 1
    $lastDate = $null
    $useGray = $false

    # ISPRAVNO SORTIRANJE - samo po vremenu, najnovije prvo
    foreach ($log in $LogonData | Sort-Object Time -Descending) {
        $currentDate = $log.Time.Date
        if ($currentDate -ne $lastDate) {
            $useGray = -not $useGray
            $dateClass = if ($useGray) { "date-gray" } else { "date-white" }
            $dateDisplay = $currentDate.ToString("dd.MM.yyyy.")
            $html += "<tr class='date-header'><td colspan='4'>$dateDisplay</td></tr>`n"
            $lastDate = $currentDate
        }

        # Korisnièko ime bez domene
        $username = $log.User
        if ($username -match "\\") {
            $username = $username.Split('\')[-1]
        }
        if ($username -eq "" -or $username -eq "Nepoznato") { 
            $username = "Unknown" 
        }

        # Event styling s ikonama
        $eventClass = if ($log.EventType -eq "Logon") { "logon" } else { "logoff" }
        $eventIcon = "<div class='event-icon $eventClass'></div>"
        $eventBadge = "<span class='event-badge $eventClass'>$eventIcon<span style='margin-left: 4px;'>$($log.EventType)</span></span>"
        
        # Vrijeme prikaz
        $time = "<span class='time-display'>$($log.Time.ToString('HH:mm:ss'))</span>"
        
        # Korisnik prikaz
        $userDisplay = "<span class='user-name'>$username</span>"

        $html += "<tr class='$dateClass'><td>$i</td><td>$time</td><td>$eventBadge</td><td>$userDisplay</td></tr>`n"
        $i++
    }

    $html += @"
                </tbody>
            </table>
        </div>

        <div class='stats-tables'>
            <div class='stat-table' style='max-width: 600px; margin: 0 auto;'>
                <h2>Broj prijava po korisniku</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Korisnik</th>
                            <th>Broj prijava</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    # Broj prijava po korisniku
    $logonCounts = $LogonData | Where-Object { $_.EventType -eq "Logon" } | Group-Object User | Sort-Object Count -Descending
    foreach ($group in $logonCounts) {
        $username = $group.Name
        if ($username -match "\\") {
            $username = $username.Split('\')[-1]
        }
        $html += "<tr><td>$username</td><td>$($group.Count)</td></tr>`n"
    }

    $html += @"
                    </tbody>
                </table>
            </div>

            <div class='stat-table' style='max-width: 600px; margin: 0 auto;'>
                <h2>Broj odjava po korisniku</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Korisnik</th>
                            <th>Broj odjava</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    # Broj odjava po korisniku
    $logoffCounts = $LogonData | Where-Object { $_.EventType -eq "LogOff" } | Group-Object User | Sort-Object Count -Descending
    foreach ($group in $logoffCounts) {
        $username = $group.Name
        if ($username -match "\\") {
            $username = $username.Split('\')[-1]
        }
        $html += "<tr><td>$username</td><td>$($group.Count)</td></tr>`n"
    }

    $html += @"
                    </tbody>
                </table>
            </div>
        </div>

        <div class='stat-table' style='margin: 20px auto; max-width: 600px;'>
            <h2>Prosjeèno trajanje sesije</h2>
            <table>
                <thead>
                    <tr>
                        <th>Korisnik</th>
                        <th>Prosjek trajanja</th>
                        <th>Broj parova</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Poboljšano izraèunavanje trajanja sesije
    $users = $LogonData | Group-Object User
    foreach ($userGroup in $users) {
        $user = $userGroup.Name
        $events = $userGroup.Group | Sort-Object Time
        $sessions = @()
        
        for ($j = 0; $j -lt $events.Count - 1; $j++) {
            if ($events[$j].EventType -eq "Logon" -and $events[$j+1].EventType -eq "LogOff") {
                $duration = $events[$j+1].Time - $events[$j].Time
                if ($duration.TotalMinutes -gt 0 -and $duration.TotalDays -lt 1) {
                    $sessions += $duration
                }
            }
        }
        
        if ($sessions.Count -gt 0) {
            $avg = New-TimeSpan -Seconds ($sessions | Measure-Object -Property TotalSeconds -Average).Average
            if ($user -match "\\") {
                $user = $user.Split('\')[-1]
            }
            $html += "<tr><td>$user</td><td>$($avg.ToString("hh\:mm\:ss"))</td><td>$($sessions.Count)</td></tr>`n"
        }
    }

    $html += @"
                </tbody>
            </table>
        </div>

        <div class='stats-box'>
            <div class='stat-line'><b>Generirano:</b> <span>$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</span></div>
            <div class='stat-line'><b>Sustav:</b> <span>$env:COMPUTERNAME</span></div>
            <div class='stat-line'><b>Korisnik:</b> <span>$env:USERNAME</span></div>
            <div class='stat-line'><b>Ukupno dogaðaja:</b> <span>$($LogonData.Count)</span></div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $ReportPath -Value $html -Encoding UTF8
    Write-Host "HTML izvještaj generiran: $ReportPath" -ForegroundColor Green
}






# === Funkcija: Neaktivni korisnici ===
function Generate-InactiveUsersReport {
    param (
        [string]$OutputPath
    )

    $users = Get-LocalUser | Where-Object { $_.Enabled -eq $false } | Select-Object Name, Enabled

    if ($users.Count -gt 0) {
        $i = 1
        $adminCount = 0
        $rows = $users | ForEach-Object {
            $username = $_.Name
            $passwordSet = "Nepoznato"
            $lastLogon = "Nepoznato"
            $passwordStatus = "Nepoznato"
            $isAdmin = $false
            $groupList = ""

            try {
                $netUserOutput = net user $username 2>&1
                if ($LASTEXITCODE -eq 0) {
                    # Password last set
                    $pwdLine = $netUserOutput | Where-Object { $_ -match "Password last set|Zadnja promjena lozinke" }
                    if ($pwdLine) {
                        $normalized = ($pwdLine -replace "\s{2,}", "|").Trim()
                        $fields = $normalized -split "\|"
                        if ($fields.Count -ge 2) {
                            $passwordSet = $fields[1].Trim()
                        }
                    }

                    # Last logon
                    $logonLine = $netUserOutput | Where-Object { $_ -match "Last logon|Zadnja prijava" }
                    if ($logonLine) {
                        $normalized = ($logonLine -replace "\s{2,}", "|").Trim()
                        $fields = $normalized -split "\|"
                        if ($fields.Count -ge 2) {
                            $lastLogon = $fields[1].Trim()
                        }
                    }

                    # Password status
                    $expiresLine = $netUserOutput | Where-Object { $_ -match "Password expires|Istek lozinke" }
                    $changeLine  = $netUserOutput | Where-Object { $_ -match "User may change password|Korisnik može promijeniti lozinku" }
                    $requiredLine = $netUserOutput | Where-Object { $_ -match "Password required|Potrebna lozinka" }

                    $expires = "?"
                    $canChange = "?"
                    $required = "?"

                    if ($expiresLine) {
                        $normalized = ($expiresLine -replace "\s{2,}", "|").Trim()
                        $parts = $normalized -split "\|"
                        if ($parts.Count -ge 2) {
                            $expires = $parts[1].Trim() -replace "Never", "Nikad"
                        }
                    }
                    if ($changeLine) {
                        $normalized = ($changeLine -replace "\s{2,}", "|").Trim()
                        $parts = $normalized -split "\|"
                        if ($parts.Count -ge 2) {
                            $canChange = switch ($parts[1].Trim().ToLower()) {
                                "yes" { "Da" }
                                "no"  { "Ne" }
                                default { $parts[1].Trim() }
                            }
                        }
                    }
                    if ($requiredLine) {
                        $normalized = ($requiredLine -replace "\s{2,}", "|").Trim()
                        $parts = $normalized -split "\|"
                        if ($parts.Count -ge 2) {
                            $required = switch ($parts[1].Trim().ToLower()) {
                                "yes" { "Da" }
                                "no"  { "Ne" }
                                default { $parts[1].Trim() }
                            }
                        }
                    }

                    $passwordStatus = "<div class='pwd-status'><span>Istek: <b>$expires</b></span><span>Promjena: <b>$canChange</b></span><span>Potrebna: <b>$required</b></span></div>"

                    # Grupe
                    $groupLines = $netUserOutput | Where-Object { $_ -match "Local Group Memberships|Lokalne grupe korisnika" }
                    if ($groupLines) {
                        $groupList = ($groupLines -replace ".*\*", "").Trim()
                        if ($groupList -match "Administrators|Administratori") {
                            $isAdmin = $true
                            $adminCount++
                        }
                    }
                }
            } catch {
                $passwordSet = "Greška"
                $lastLogon = "Greška"
                $passwordStatus = "Greška"
            }

            $class = if ($isAdmin) { "admin-user" } else { "" }
            $adminBadge = if ($isAdmin) { "<span class='admin-badge'>ADMIN</span>" } else { "" }

            "<tr class='$class'><td>$i</td><td>$username $adminBadge</td><td><span class='status-disabled'>Onemoguæen</span></td><td>$passwordSet</td><td>$lastLogon</td><td>$passwordStatus</td></tr>"
            $i++
        } | Out-String

        $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Neaktivni Korisnici</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

tr.admin-user td {
    background-color: #fef2f2 !important;
}

tr.admin-user:hover td {
    background-color: #fee2e2 !important;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 60px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Status stilovi */
.status-disabled {
    color: #ef4444;
    font-weight: 600;
}

.admin-badge {
    background: linear-gradient(135deg, #ef4444, #dc2626);
    color: white;
    padding: 2px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 8px;
}

/* Password status stilovi */
.pwd-status {
    display: flex;
    gap: 15px;
    flex-wrap: wrap;
}

.pwd-status span {
    background: #f3f4f6;
    padding: 4px 8px;
    border-radius: 6px;
    font-size: 12px;
    color: #6b7280;
}

.pwd-status b {
    color: #1f2937;
    font-weight: 600;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
    .pwd-status { gap: 8px; }
}
</style>
</head>
<body>
    <h1>Popis neaktivnih korisnièkih raèuna</h1>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Ime korisnika</th>
                        <th>Status</th>
                        <th>Lozinka postavljena</th>
                        <th>Zadnja aktivnost</th>
                        <th>Status lozinke</th>
                    </tr>
                </thead>
                <tbody>
                    $rows
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Neaktivnih korisnika:</b> $($users.Count)</p>
                <p><b>Od toga administratora:</b> $adminCount</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@
    } else {
        $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset="UTF-8">
    <title>Neaktivni Korisnici</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 40px 20px;
}

.message-box {
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    padding: 60px;
    text-align: center;
    backdrop-filter: blur(10px);
}

h1 { 
    color: #667eea;
    font-size: 32px;
    margin-bottom: 20px;
}

p {
    color: #6b7280;
    font-size: 18px;
}
</style>
</head>
<body>
    <div class='message-box'>
        <h1>Neaktivni korisnièki raèuni</h1>
        <p>Nema onemoguæenih lokalnih korisnika.</p>
    </div>
</body>
</html>
"@
    }

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}





# === Funkcija: Folder Permissions Report ===
function Generate-FolderPermissionsReport {
    param (
        [string]$TargetDrive = "D:\",
        [string]$OutputPath,
        [switch]$IncludeAdvancedSettings
    )

    function Get-FolderPermissions {
        param ([string]$FolderPath)
        try {
            $acl = Get-Acl -Path $FolderPath
            $owner = ($acl.Owner -split '\\')[-1]
            $inheritanceEnabled = -not $acl.AreAccessRulesProtected
            $auditRules = try { $acl.Audit.Count } catch { 0 }
            $permissionEntries = $acl.Access.Count
            $explicitPermissions = ($acl.Access | Where-Object { -not $_.IsInherited }).Count
            $inheritedPermissions = ($acl.Access | Where-Object { $_.IsInherited }).Count
            
            $permissions = $acl.Access | ForEach-Object {
                $identity = ($_.IdentityReference -split '\\')[-1]
                $rights = $_.FileSystemRights
                $type = $_.AccessControlType
                $inherited = $_.IsInherited
                $inheritedFrom = if ($inherited) { "Nasleðeno" } else { "Eksplicitno" }
                $appliesTo = $_.InheritanceFlags.ToString() + " | " + $_.PropagationFlags.ToString()
                
                # Detaljno mapiranje osnovnih prava
                $basicRights = @()
                $rightsString = $rights.ToString()
                
                # Provjeri hijerarhijski - najprije najšira prava
                if ($rightsString -match "FullControl") { 
                    $basicRights += "Full Control" 
                }
                elseif ($rightsString -match "Modify") { 
                    $basicRights += "Modify" 
                    # Modify ukljuèuje i ostala prava, ali ih eksplicitno navedemo
                    $basicRights += "Read & Execute"
                    $basicRights += "Write"
                    $basicRights += "Read"
                    $basicRights += "List Folder Contents"
                }
                else {
                    # Provjeri pojedinaèna prava
                    if ($rightsString -match "ReadAndExecute") { $basicRights += "Read & Execute" }
                    if ($rightsString -match "ListDirectory") { $basicRights += "List Folder Contents" }
                    if ($rightsString -match "Read" -and $rightsString -notmatch "ReadAndExecute") { $basicRights += "Read" }
                    if ($rightsString -match "Write" -and $rightsString -notmatch "Modify") { $basicRights += "Write" }
                    if ($rightsString -match "ExecuteFile") { $basicRights += "Execute" }
                    if ($rightsString -match "CreateFiles") { $basicRights += "Create Files" }
                    if ($rightsString -match "CreateDirectories") { $basicRights += "Create Folders" }
                    if ($rightsString -match "Delete") { $basicRights += "Delete" }
                    if ($rightsString -match "ReadPermissions") { $basicRights += "Read Permissions" }
                    if ($rightsString -match "ChangePermissions") { $basicRights += "Change Permissions" }
                    if ($rightsString -match "TakeOwnership") { $basicRights += "Take Ownership" }
                }
                
                $basicRightsString = if ($basicRights.Count -gt 0) { 
                    ($basicRights | Sort-Object -Unique) -join ", " 
                } else { 
                    $rightsString 
                }
                
                # Kreiraj detaljni prikaz direktno u HTML - POBOLJŠANO ZA KOMPOZITNA PRAVA
                $detailedRightsHTML = $rights.ToString()
                
                # Debug - dodaj originalnu vrijednost za analizu
                Write-Host "Analiziram prava za $identity : '$detailedRightsHTML'" -ForegroundColor Yellow
                Write-Host "Numerièka vrijednost: $([int]$rights)" -ForegroundColor Cyan
                
                # Za kompozitna prava, analiziraj bit po bit
                $granularRights = @()
                $rightsValue = [int]$rights
                
                # Provjeri svaki bit pojedinaèno
                if ($rightsValue -band 1) { $granularRights += "List folder / read data" }
                if ($rightsValue -band 2) { $granularRights += "Create files / write data" }
                if ($rightsValue -band 4) { $granularRights += "Create folders / append data" }
                if ($rightsValue -band 32) { $granularRights += "Read extended attributes" }
                if ($rightsValue -band 64) { $granularRights += "Write extended attributes" }
                if ($rightsValue -band 128) { $granularRights += "Traverse folder / execute file" }
                if ($rightsValue -band 256) { $granularRights += "Delete subfolders and files" }
                if ($rightsValue -band 128) { $granularRights += "Read attributes" }
                if ($rightsValue -band 256) { $granularRights += "Write attributes" }
                if ($rightsValue -band 65536) { $granularRights += "Delete" }
                if ($rightsValue -band 131072) { $granularRights += "Read permissions" }
                if ($rightsValue -band 262144) { $granularRights += "Change permissions" }
                if ($rightsValue -band 524288) { $granularRights += "Take ownership" }
                if ($rightsValue -band 1048576) { $granularRights += "Synchronize" }
                
                # Ako je pronašao granularna prava, koristi ih
                if ($granularRights.Count -gt 0) {
                    $detailedRightsHTML = ($granularRights | Sort-Object -Unique) -join "<br/>• "
                    $detailedRightsHTML = "• " + $detailedRightsHTML
                } else {
                    # Fallback na originalno mapiranje
                    $detailedRightsHTML = $detailedRightsHTML -replace "FullControl", "FullControl (sve dozvole)"
                    $detailedRightsHTML = $detailedRightsHTML -replace "Modify", "Modify (mijenjanje)"  
                    $detailedRightsHTML = $detailedRightsHTML -replace "ReadAndExecute", "ReadAndExecute (èitanje i izvršavanje)"
                    $detailedRightsHTML = $detailedRightsHTML -replace "Write", "Write (pisanje)"
                    $detailedRightsHTML = $detailedRightsHTML -replace "Read", "Read (èitanje)"
                    $detailedRightsHTML = $detailedRightsHTML -replace "Synchronize", "Synchronize (sinkronizacija)"
                    $detailedRightsHTML = $detailedRightsHTML -replace ", ", "<br/>• "
                    $detailedRightsHTML = "• " + $detailedRightsHTML
                }
                
                # Odreðivanje boje prema tipu pristupa
                $color = if ($type -eq "Allow") { "#10b981" } else { "#ef4444" }
                $typeText = if ($type -eq "Allow") { "Dopušteno" } else { "Zabranjeno" }
                $inheritanceColor = if ($inherited) { "#f59e0b" } else { "#3b82f6" }
                
                # Advanced Security Settings - KOMPAKTNI PRIKAZ
                $principal = ($_.IdentityReference -split '\\')[-1]  # Samo korisnièko ime bez raèunala
                $accessType = $type.ToString()
                $advancedInfo = @"
<div style='background-color:#f8fafc;padding:6px 8px;border-radius:4px;margin-top:6px;border-left:2px solid #cbd5e1;'>
    <div style='font-size:10px;color:#64748b;line-height:1.3;'>
        <strong style='color:#475569;'>$principal</strong> | <span style='color:#059669;'>$accessType</span> | <span style='color:#dc2626;'>$inheritedFrom</span> | <span style='color:#7c3aed;'>$appliesTo</span>
    </div>
</div>
"@
                
                "<div style='margin-bottom:8px;padding:8px;background-color:#f3f4f6;border-radius:6px;border-left:3px solid $color;'>
                    <div class='permission-header' style='display:flex;justify-content:space-between;align-items:center;padding:4px;'>
                        <div>
                            <strong style='color:#1f2937;font-size:14px;'>$identity</strong>
                            <span style='color:$color;font-weight:600;font-size:12px;background:rgba(255,255,255,0.8);padding:2px 8px;border-radius:12px;margin-left:10px;'>$typeText</span>
                        </div>
                        <div style='font-size:12px;color:#6b7280;'>
                            <strong style='color:#1e40af;'>Osnovna:</strong> $basicRightsString
                        </div>
                    </div>
                    
                    <div style='margin-top:10px;padding-top:8px;border-top:1px solid #e5e7eb;'>
                        <div style='margin-bottom:6px;'>
                            <strong style='color:#dc2626;font-size:13px;font-weight:700;background:#fef2f2;padding:2px 8px;border-radius:4px;'>Advanced Security:</strong>
                            <div style='color:#6b7280;font-size:11px;margin-left:4px;margin-top:4px;line-height:1.4;'>$detailedRightsHTML</div>
                        </div>
                        
                        <div style='display:flex;justify-content:space-between;font-size:11px;margin-bottom:6px;'>
                            <span style='color:$inheritanceColor;font-weight:500;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:3px;'>$inheritedFrom</span>
                            <span style='color:#9ca3af;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:3px;font-size:10px;'>$appliesTo</span>
                        </div>
                        
                        $advancedInfo
                    </div>
                </div>"
            }
            
            # Dodaj sveobuhvatne Advanced Security Settings na vrh - KOMPAKTNO
            $globalAdvancedSettings = @"
<div style='background-color:#e0f2fe;padding:8px 12px;border-radius:6px;margin-bottom:8px;border-left:3px solid #0288d1;'>
    <div style='font-weight:600;color:#01579b;margin-bottom:4px;font-size:13px;'>Security Overview</div>
    <div style='font-size:11px;color:#0277bd;line-height:1.4;display:flex;flex-wrap:wrap;gap:15px;'>
        <span><strong>Owner:</strong>&nbsp;$owner</span>&nbsp;&nbsp;
        <span><strong>Entries:</strong>&nbsp;$permissionEntries</span>&nbsp;&nbsp;
        <span><strong>Explicit:</strong>&nbsp;$explicitPermissions</span>&nbsp;&nbsp;
        <span><strong>Inherited:</strong>&nbsp;$inheritedPermissions</span>&nbsp;&nbsp;
        <span><strong>Inheritance:</strong>&nbsp;$(if ($inheritanceEnabled) { "ON" } else { "OFF" })</span>&nbsp;&nbsp;
        <span><strong>Audit:</strong>&nbsp;$auditRules</span>
    </div>
</div>
"@
            
            $permissionsString = $globalAdvancedSettings + ($permissions -join "")
        } catch {
            $permissionsString = "<span style='color:#ef4444;font-style:italic;'>Greška pri pristupu dozvolama</span>"
        }

        [PSCustomObject]@{
            Folder      = $FolderPath
            Permissions = $permissionsString
        }
    }

    # Poèetak HTML-a
    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Detaljni izvještaj prava pristupa - $TargetDrive</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.drive-indicator {
    background: rgba(255,255,255,0.2);
    color: white;
    padding: 5px 15px;
    border-radius: 20px;
    font-size: 18px;
    font-weight: 600;
    display: inline-block;
    margin-top: -15px;
    margin-bottom: 25px;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
    width: 50px;
}

th:nth-child(2) {
    width: 30%;
}

th:last-child {
    border-top-right-radius: 10px;
    width: 70%;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Stilovi za folder path */
td:nth-child(2) {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 13px;
    color: #1e40af;
    word-break: break-all;
}

/* Stilovi za permissions */
td:nth-child(3) {
    padding: 8px;
}

/* Expandable/Collapsible stilovi */
.permission-header {
    transition: background-color 0.2s ease;
}

.permission-header:hover {
    background-color: rgba(0,0,0,0.05);
    border-radius: 4px;
}

.toggle-icon {
    transition: transform 0.3s ease;
    display: inline-block;
}

.toggle-icon.expanded {
    transform: rotate(90deg);
}

.permission-details {
    animation: slideDown 0.3s ease;
}

@keyframes slideDown {
    from {
        opacity: 0;
        max-height: 0;
        transform: translateY(-10px);
    }
    to {
        opacity: 1;
        max-height: 500px;
        transform: translateY(0);
    }
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1>Detaljni izvještaj prava pristupa</h1>
    <div class='drive-indicator'>$TargetDrive</div>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>R.br</th>
                        <th>Folder</th>
                        <th>Osnovna prava i Advanced Security</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Dohvaæanje foldera
    $rootFolder = Get-Item -Path $TargetDrive
    $subFolders = Get-ChildItem -Path $TargetDrive -Directory -Recurse -ErrorAction SilentlyContinue
    $folders = @($rootFolder) + $subFolders

    $index = 1
    foreach ($folder in $folders) {
        $perm = Get-FolderPermissions -FolderPath $folder.FullName
        $html += "<tr><td>$index</td><td>$($perm.Folder)</td><td>$($perm.Permissions)</td></tr>`n"
        $index++
    }

    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $folderCount = $folders.Count
    
    $html += @"
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $timestamp</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno foldera:</b> $folderCount</p>
                <p><b>Tip izvještaja:</b> $(if ($IncludeAdvancedSettings) { "S Advanced Settings" } else { "Standardni" })</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>

<script type="text/javascript">
function togglePermissions(header) {
    try {
        const details = header.nextElementSibling;
        const icon = header.querySelector('.toggle-icon');
        
        if (!details || !icon) return;
        
        if (details.style.display === 'none' || details.style.display === '') {
            details.style.display = 'block';
            icon.innerHTML = '[-]';
            icon.classList.add('expanded');
        } else {
            details.style.display = 'none';
            icon.innerHTML = '[+]';
            icon.classList.remove('expanded');
        }
    } catch (e) {
        console.error('Error toggling permissions:', e);
    }
}

function toggleAllPermissions() {
    try {
        const allDetails = document.querySelectorAll('.permission-details');
        const allIcons = document.querySelectorAll('.toggle-icon');
        const isAnyExpanded = Array.from(allDetails).some(detail => detail.style.display === 'block');
        
        allDetails.forEach((detail, index) => {
            if (isAnyExpanded) {
                detail.style.display = 'none';
                if (allIcons[index]) {
                    allIcons[index].innerHTML = '[+]';
                    allIcons[index].classList.remove('expanded');
                }
            } else {
                detail.style.display = 'block';
                if (allIcons[index]) {
                    allIcons[index].innerHTML = '[-]';
                    allIcons[index].classList.add('expanded');
                }
            }
        });
    } catch (e) {
        console.error('Error toggling all permissions:', e);
    }
}

// Dodaj kontrole nakon što se DOM uèita
window.addEventListener('DOMContentLoaded', function() {
    try {
        const tableWrapper = document.querySelector('.table-wrapper');
        if (tableWrapper) {
            const controlsDiv = document.createElement('div');
            controlsDiv.style.marginBottom = '15px';
            controlsDiv.style.textAlign = 'center';
            controlsDiv.innerHTML = `
                <button onclick="toggleAllPermissions()" style="
                    padding: 10px 20px;
                    background: linear-gradient(135deg, #667eea, #764ba2);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 13px;
                    font-weight: 600;
                    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
                    transition: transform 0.2s ease;
                " onmouseover="this.style.transform='translateY(-2px)'" onmouseout="this.style.transform='translateY(0)'">
                    ?? Proširi / Skupi sve dozvole
                </button>
            `;
            tableWrapper.insertBefore(controlsDiv, tableWrapper.firstChild);
        }
    } catch (e) {
        console.error('Error adding controls:', e);
    }
});
</script>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}




# === TAB: Korištenje USB s korisnikom ===
function Generate-USBUserReport {
    param (
        [string]$OutputPath,
        [int]$DaysBack = 40
    )

    $DetectionTime = Get-Date -Format "dd.MM.yyyy. HH:mm:ss"
    $startDate = (Get-Date).AddDays(-$DaysBack)
    $fallbackUser = ((Get-WmiObject -Class Win32_ComputerSystem).UserName -split '\\')[-1]

    # === FUNKCIJA ZA PARSIRANJE USB VENDOR ID ===
    function Get-USBVendorInfo {
        param([string]$VendorId)
        
        $vendors = @{
            "2717" = @{ Name = "Xiaomi Corporation"; Type = "Mobilni ureðaj" }
            "04E8" = @{ Name = "Samsung Electronics"; Type = "Mobilni ureðaj" }
            "18D1" = @{ Name = "Google Inc."; Type = "Android ureðaj" }
            "0BB4" = @{ Name = "HTC Corporation"; Type = "Mobilni ureðaj" }
            "22B8" = @{ Name = "Motorola"; Type = "Mobilni ureðaj" }
            "05AC" = @{ Name = "Apple Inc."; Type = "iPhone/iPad" }
            "045E" = @{ Name = "Microsoft"; Type = "Microsoft ureðaj" }
            "1058" = @{ Name = "Western Digital"; Type = "Vanjski disk" }
            "0781" = @{ Name = "SanDisk Corporation"; Type = "USB memorija" }
            "090C" = @{ Name = "Silicon Motion"; Type = "USB kontroler" }
            "0951" = @{ Name = "Kingston Technology"; Type = "USB memorija" }
            "8087" = @{ Name = "Intel"; Type = "Intel ureðaj" }
            "1D6B" = @{ Name = "Linux Foundation"; Type = "USB Hub" }
            "0424" = @{ Name = "Standard Microsystems"; Type = "USB Hub" }
        }
        
        if ($vendors.ContainsKey($VendorId.ToUpper())) {
            return $vendors[$VendorId.ToUpper()]
        }
        
        return @{ Name = "Nepoznat proizvoðaè ($VendorId)"; Type = "Nepoznato" }
    }

    # === FUNKCIJA ZA PARSIRANJE USB PORUKA ===
    function Parse-USBMessage {
        param([string]$Message)
        
        $result = @{
            DeviceType = "USB ureðaj"
            DeviceName = "Nepoznat ureðaj"
            Vendor = "Nepoznato"
            SerialNumber = ""
            DetailedInfo = ""
            IconType = "USB"
        }
        
        try {
            # Parsiranje VID i PID iz razlièitih formata
            if ($Message -match 'USB\.VID_([0-9A-F]{4})&PID_([0-9A-F]{4})\.([^\\s]+)') {
                $vendorId = $matches[1]
                $productId = $matches[2] 
                $deviceSerial = $matches[3]
                
                $vendorInfo = Get-USBVendorInfo -VendorId $vendorId
                $result.Vendor = $vendorInfo.Name
                $result.DeviceType = $vendorInfo.Type
                $result.SerialNumber = $deviceSerial
                
                # Posebni sluèajevi
                switch ($vendorId.ToUpper()) {
                    "2717" {
                        $result.DeviceName = "Xiaomi Android telefon"
                        $result.DeviceType = "Android telefon"
                        $result.DetailedInfo = "Android ureðaj u MTP/ADB naèinu rada"
                        $result.IconType = "ANDROID"
                    }
                    "04E8" {
                        $result.DeviceName = "Samsung Android telefon"
                        $result.DeviceType = "Android telefon"
                        $result.DetailedInfo = "Samsung Galaxy ureðaj"
                        $result.IconType = "ANDROID"
                    }
                    "18D1" {
                        $result.DeviceName = "Google Android ureðaj"
                        $result.DeviceType = "Android telefon"
                        $result.DetailedInfo = "Google Pixel ili Nexus ureðaj"
                        $result.IconType = "ANDROID"
                    }
                    "05AC" {
                        $result.DeviceName = "Apple iPhone/iPad"
                        $result.DeviceType = "iOS ureðaj"
                        $result.DetailedInfo = "Apple mobilni ureðaj"
                        $result.IconType = "APPLE"
                    }
                    "1058" {
                        $result.DeviceName = "Western Digital disk"
                        $result.DeviceType = "Vanjski hard disk"
                        $result.DetailedInfo = "WD vanjski disk za pohranu"
                        $result.IconType = "DISK"
                    }
                    "0781" {
                        $result.DeviceName = "SanDisk USB memorija"
                        $result.DeviceType = "USB flash memorija"
                        $result.DetailedInfo = "SanDisk prijenosna memorija"
                        $result.IconType = "USB"
                    }
                }
            }
            # Parsiranje USBSTOR formata
            elseif ($Message -match 'USBSTOR#(DISK)&VEN_([^&]+)&PROD_([^&]+)&REV_([^#]+).*?#(\w+)&0#') {
                $deviceTypeRaw = $matches[1]
                $vendor = $matches[2] -replace '_', ' '
                $product = $matches[3] -replace '_', ' '
                $revision = $matches[4]
                $serial = $matches[5]
                
                $result.DeviceType = "USB disk"
                $result.DeviceName = "$vendor $product"
                $result.Vendor = $vendor
                $result.SerialNumber = $serial
                $result.DetailedInfo = "Disk: $vendor $product (Rev: $revision, S/N: $serial)"
                $result.IconType = "DISK"
            }
            # Opæeniti USB format
            elseif ($Message -match 'device\s+([^\s.]+)') {
                $deviceString = $matches[1]
                
                if ($deviceString -match 'VID_([0-9A-F]{4})') {
                    $vendorId = $matches[1]
                    $vendorInfo = Get-USBVendorInfo -VendorId $vendorId
                    $result.Vendor = $vendorInfo.Name
                    $result.DeviceType = $vendorInfo.Type
                    $result.DeviceName = "$($vendorInfo.Name) ureðaj"
                }
            }
            
        } catch {
            # Nastavi s default vrijednostima
        }
        
        return $result
    }

    # === SKUPLJANJE LOGONA KORISNIKA ===
    try {
        $logons = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; ID = 4624; StartTime = $startDate } -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $xml = [xml]$_.ToXml()
                $logonType = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "LogonType" } | Select-Object -ExpandProperty '#text'
                $user = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "TargetUserName" } | Select-Object -ExpandProperty '#text'

                if ($logonType -eq "2" -or $logonType -eq "10") {
                    [PSCustomObject]@{
                        TimeCreated = $_.TimeCreated
                        User = "$user"
                    }
                }
            } catch {
                # Preskoèi problematiène eventi
            }
        } | Sort-Object TimeCreated
    } catch {
        $logons = @()
    }

    # === SKUPLJANJE USB DOGAÐAJA ===
    $usbEvents = @()
    
    # Osnovni logovi koji uvijek postoje
    $logSources = @(
        @{ Log = "System"; Ids = @(2003, 2100, 2102, 400, 410, 20001, 20002, 6416) },
        @{ Log = "Microsoft-Windows-DriverFrameworks-UserMode/Operational"; Ids = @(20001, 20003, 1003, 2003, 2100) }
    )

    # Dodatni logovi koji èesto nisu dostupni - provjeri prije dodavanja
    $optionalLogs = @(
        "Microsoft-Windows-USB-USBHUB3/Operational",
        "Microsoft-Windows-USB-USBPORT/Analytic"
    )

    foreach ($logName in $optionalLogs) {
        try {
            $logInfo = Get-WinEvent -ListLog $logName -ErrorAction Stop
            if ($logInfo.IsEnabled) {
                $logSources += @{ Log = $logName; Ids = @(20001, 20003, 1, 2, 3) }
            }
        } catch {
            # Log nije dostupan - nastavi
        }
    }

    $keywords = @("usb", "storage", "removable", "mass", "device", "volume", "vid_", "pid_", "usbstor")

    foreach ($src in $logSources) {
        try {
            # Provjeri postoji li log prije pokušaja èitanja
            $logExists = $false
            try {
                $logInfo = Get-WinEvent -ListLog $src.Log -ErrorAction Stop
                $logExists = $true
            } catch {
                continue
            }
            
            if ($logExists) {
                $events = Get-WinEvent -FilterHashtable @{
                    LogName = $src.Log
                    ID = $src.Ids
                    StartTime = $startDate
                } -ErrorAction SilentlyContinue

                if ($events) {
                    $filtered = $events | Where-Object {
                        $msg = $_.Message.ToLower()
                        $keywords | Where-Object { $msg -like "*$_*" }
                    }

                    if ($filtered) {
                        $usbEvents += $filtered
                    }
                }
            }
        } catch {
            # Nastavi s drugim logovima
        }
    }

    # === RAÈUNANJE STATISTIKA ===
    $androidDevices = 0
    $storageDevices = 0
    $uniqueDeviceNames = @()
    
    foreach ($event in $usbEvents) {
        $usbInfo = Parse-USBMessage -Message $event.Message
        
        if ($usbInfo.IconType -eq "ANDROID") {
            $androidDevices++
        }
        
        if ($usbInfo.IconType -eq "DISK" -or $usbInfo.DeviceType -like "*disk*" -or $usbInfo.DeviceType -like "*storage*") {
            $storageDevices++
        }
        
        if ($uniqueDeviceNames -notcontains $usbInfo.DeviceName) {
            $uniqueDeviceNames += $usbInfo.DeviceName
        }
    }
    
    $uniqueDevices = $uniqueDeviceNames.Count

    # === POÈETAK HTML-A ===
    $htmlContent = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>USB Povijest - s korisnicima (Poboljšano)</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 10px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.period-info {
    color: white;
    font-size: 16px;
    margin-bottom: 30px;
    opacity: 0.9;
    text-align: center;
}

.container {
    width: 100%;
    max-width: 1500px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 700px; 
    overflow-y: auto; 
    overflow-x: auto;
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 12px;
    height: 12px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 12px;
    min-width: 900px;
    table-layout: fixed;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 12px 6px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.3px;
    z-index: 10;
    font-size: 12px;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 8px 4px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
    font-size: 12px;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 50px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Stilovi za specifiène stupce - optimizirane širine */
td:nth-child(1) { /* # - sadržaj æelije */
    width: 50px;
    min-width: 50px;
    max-width: 50px;
}

th:nth-child(1) { /* # - header */
    width: 50px;
    min-width: 50px;
    max-width: 50px;
    color: white !important;
}

td:nth-child(2) { /* Vrijeme - sadržaj æelije */
    font-weight: 500;
    white-space: nowrap;
    width: 130px;
    min-width: 130px;
    max-width: 130px;
}

th:nth-child(2) { /* Vrijeme - header */
    width: 130px;
    min-width: 130px;
    max-width: 130px;
    color: white !important;
}

td:nth-child(3) { /* Tip ureðaja - sadržaj æelije */
    font-weight: 600;
    color: #059669;
    width: 120px;
    min-width: 120px;
    max-width: 120px;
}

th:nth-child(3) { /* Tip ureðaja - header */
    width: 120px;
    min-width: 120px;
    max-width: 120px;
    color: white !important;
}

td:nth-child(4) { /* Naziv ureðaja - sadržaj æelije */
    font-weight: 500;
    color: #1f2937;
    width: 180px;
    min-width: 180px;
    max-width: 180px;
    word-wrap: break-word;
    overflow-wrap: break-word;
}

th:nth-child(4) { /* Naziv ureðaja - header */
    width: 180px;
    min-width: 180px;
    max-width: 180px;
    color: white !important;
}

td:nth-child(5) { /* Opis dogaðaja - sadržaj æelije */
    line-height: 1.4;
    word-wrap: break-word;
    overflow-wrap: break-word;
    width: auto; /* Uzima ostatak prostora */
    min-width: 300px;
}

th:nth-child(5) { /* Opis dogaðaja - header */
    width: auto;
    min-width: 300px;
    color: white !important;
}

td:nth-child(6) { /* Korisnik - sadržaj æelije */
    font-weight: 600;
    color: #dc2626;
    width: 100px;
    min-width: 100px;
    max-width: 100px;
    word-wrap: break-word;
}

th:nth-child(6) { /* Korisnik - header */
    width: 100px;
    min-width: 100px;
    max-width: 100px;
    color: white !important;
}

.device-badge {
    display: inline-block;
    padding: 2px 6px;
    border-radius: 10px;
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
}

.badge-android {
    background-color: #dcfce7;
    color: #166534;
    border: 1px solid #22c55e;
}

.badge-apple {
    background-color: #f3f4f6;
    color: #374151;
    border: 1px solid #6b7280;
}

.badge-disk {
    background-color: #dbeafe;
    color: #1e40af;
    border: 1px solid #3b82f6;
}

.badge-usb {
    background-color: #fef3c7;
    color: #92400e;
    border: 1px solid #f59e0b;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
    margin-bottom: 20px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px;
    border-radius: 20px;
    box-shadow: 0 15px 40px rgba(0,0,0,0.25);
    text-align: left;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
    width: 350px;
    max-width: 350px;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 20px 50px rgba(0,0,0,0.35);
}

.stats-grid {
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-content: flex-start;
    margin-bottom: 25px;
    height: 140px;
}

.stat-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    text-align: center;
    width: 48%;
    height: 65px;
    margin-bottom: 10px;
}

.stat-number {
    font-size: 28px;
    font-weight: 700;
    color: #fff;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    margin-bottom: 2px;
    line-height: 1;
}

.stat-label {
    font-size: 9px;
    color: rgba(255,255,255,0.95);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    line-height: 1.1;
    text-align: center;
}

.stats-info {
    border-top: 1px solid rgba(255,255,255,0.2);
    padding-top: 18px;
    text-align: center;
}

.stats-info p {
    margin: 4px 0;
    font-size: 14px;
    color: rgba(255,255,255,0.9);
    font-weight: 500;
}

/* USB info stilovi */
.usb-info {
    background: #f3f4f6;
    padding: 6px;
    margin-top: 4px;
    border-radius: 6px;
    border-left: 3px solid #667eea;
    font-size: 11px;
    max-width: 100%;
    word-wrap: break-word;
    white-space: pre-line;
}

.usb-info i {
    color: #6b7280;
    font-style: normal;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}

.no-events {
    text-align: center;
    padding: 40px;
    color: #6b7280;
    font-style: italic;
}
</style>
</head>
<body>
    <h1>Korištenje USB ureðaja - detaljni pregled</h1>
    <div class='period-info'>
        Vrijeme generiranja: $DetectionTime<br>
        Period: Zadnjih $DaysBack dana
    </div>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Vrijeme</th>
                        <th>Tip ureðaja</th>
                        <th>Naziv ureðaja</th>
                        <th>Opis dogaðaja</th>
                        <th>Korisnik</th>
                    </tr>
                </thead>
                <tbody>
"@

    if ($usbEvents.Count -eq 0) {
        $htmlContent += "<tr><td colspan='6' class='no-events'>Nema pronaðenih USB dogaðaja u zadnjih $DaysBack dana</td></tr>"
    } else {
        $row = 1
        foreach ($event in $usbEvents | Sort-Object TimeCreated -Descending) {
            $time = $event.TimeCreated
            $formattedTime = $time.ToString("dd.MM.yyyy HH:mm:ss")
            $provider = $event.ProviderName -replace "Microsoft-Windows-", "MS-"
            $id = $event.Id
            $message = ($event.Message -replace '\r?\n', ' ').Trim()

            # Poboljšano parsiranje USB informacija
            $usbInfo = Parse-USBMessage -Message $message
            
            # Badge za tip ureðaja
            $badgeClass = switch ($usbInfo.IconType) {
                "ANDROID" { "badge-android" }
                "APPLE" { "badge-apple" }
                "DISK" { "badge-disk" }
                default { "badge-usb" }
            }
            $deviceTypeBadge = "<span class='device-badge $badgeClass'>$($usbInfo.DeviceType)</span>"

            # Dodatne informacije iz poruke (kao u originalnom kodu)
            $extraInfo = ""
            if ($message -match 'USBSTOR#(DISK)&VEN_([^&]+)&PROD_([^&]+)&REV_([^#]+).*?#(\w+)&0#') {
                $deviceType = $matches[1]
                $vendor = $matches[2] -replace '_', ' '
                $product = $matches[3] -replace '_', ' '
                $revision = $matches[4]
                $serial = $matches[5]
                $extraInfo = "<div class='usb-info'><i>Vrsta ureðaja:</i> $deviceType<br><i>Naziv ureðaja:</i> $vendor $product<br><i>Revizija firmvera:</i> $revision<br><i>Serijski broj:</i> $serial</div>"
            }
            # Dodatni format za VEN__USB (SanDisk i slièni)
            elseif ($message -match 'USBSTOR#(DISK)&VEN_([^&]+)&PROD_([^&]+)&REV_([^#]+).*?#([A-F0-9]{10,})') {
                $deviceType = $matches[1]
                $vendor = $matches[2] -replace '_', ' '
                $product = $matches[3] -replace '_', ' '
                $revision = $matches[4]
                $serial = $matches[5]
                
                # Prepoznaj SanDisk
                if ($product -like "*SANDISK*") {
                    $vendor = "SanDisk"
                    $product = ($product -replace 'SANDISK', '').Trim()
                }
                
                $extraInfo = "<div class='usb-info'><i>Vrsta ureðaja:</i> $deviceType<br><i>Naziv ureðaja:</i> $vendor $product<br><i>Revizija firmvera:</i> $revision<br><i>Serijski broj:</i> $serial</div>"
            }
            elseif ($usbInfo.DetailedInfo -ne "") {
                $extraInfo = "<div class='usb-info'><i>Detalji:</i> $($usbInfo.DetailedInfo)"
                if ($usbInfo.SerialNumber -ne "") {
                    $extraInfo += "<br><i>Serijski broj:</i> $($usbInfo.SerialNumber)"
                }
                $extraInfo += "</div>"
            }

            $combinedMessage = "$message$extraInfo"

            # Pronaði korisnika koji je bio prijavljen u to vrijeme
            $lastLogon = $logons | Where-Object { $_.TimeCreated -le $time } | Select-Object -Last 1
            $user = if ($lastLogon) { $lastLogon.User } else { if ($fallbackUser) { $fallbackUser } else { "Nepoznato" } }

            $htmlContent += "<tr><td>$row</td><td>$formattedTime</td><td>$deviceTypeBadge</td><td>$($usbInfo.DeviceName)</td><td>$combinedMessage</td><td>$user</td></tr>`n"
            $row++
        }
    }

    $htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <div class='stats-container'>
            <div class='stats-box'>
                <div class='stats-grid'>
                    <div class='stat-item'>
                        <div class='stat-number'>$($usbEvents.Count)</div>
                        <div class='stat-label'>USB dogaðaji</div>
                    </div>
                    <div class='stat-item'>
                        <div class='stat-number'>$uniqueDevices</div>
                        <div class='stat-label'>Jedinstveni ureðaji</div>
                    </div>
                    <div class='stat-item'>
                        <div class='stat-number'>$androidDevices</div>
                        <div class='stat-label'>Android ureðaji</div>
                    </div>
                    <div class='stat-item'>
                        <div class='stat-number'>$storageDevices</div>
                        <div class='stat-label'>Ureðaji za pohranu</div>
                    </div>
                </div>
                <div class='stats-info'>
                    <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                    <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                    <p><b>Analizirani period:</b> $DaysBack dana</p>
                    <p><b>© Ivica Rašan 2025</b></p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
"@

    # Spremi HTML datoteku
    try {
        $htmlContent | Set-Content -Path $OutputPath -Encoding UTF8
        return $OutputPath
    } catch {
        throw "Greška pri kreiranju HTML datoteke: $($_.Exception.Message)"
    }
}






# === Funkcija: Usporedba datuma kreiranja (Users folder + Registry + EventLog) ===
function Generate-KreiranjeKRComparisonReport {
    param([string]$OutputPath)

    # Poboljšano dohvaæanje EventLog podataka
    $eventLogData = @{}
    try {
        Write-Host "Dohvaæam EventLog podatke o kreiranju korisnika..." -ForegroundColor Cyan
        
        # Pokušaj s Get-WinEvent (moderniji pristup)
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName = 'Security'
                ID = 4720
                StartTime = (Get-Date).AddDays(-365)
            } -ErrorAction SilentlyContinue
            
            foreach ($event in $events) {
                $xml = [xml]$event.ToXml()
                $targetUserName = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "TargetUserName" } | Select-Object -ExpandProperty '#text'
                
                if ($targetUserName -and -not $eventLogData.ContainsKey($targetUserName)) {
                    $eventLogData[$targetUserName] = $event.TimeCreated.ToString("dd.MM.yyyy HH:mm:ss")
                }
            }
        } catch {
            # Fallback na Get-EventLog
            $events = Get-EventLog -LogName Security -InstanceId 4720 -Newest 1000 -ErrorAction SilentlyContinue |
                      Where-Object { $_.TimeGenerated -gt (Get-Date).AddDays(-365) }

            foreach ($event in $events) {
                $lines = $event.Message -split "`r`n"
                $username = $null
                for ($i = 0; $i -lt $lines.Length; $i++) {
                    if ($lines[$i] -match "^\s*New Account:") {
                        for ($j = $i + 1; $j -lt $i + 6 -and $j -lt $lines.Length; $j++) {
                            if ($lines[$j] -match "Account Name:\s+(\S+)") {
                                $username = $matches[1]
                                break
                            }
                        }
                        break
                    }
                }
                if ($username -and -not $eventLogData.ContainsKey($username)) {
                    $eventLogData[$username] = $event.TimeGenerated.ToString("dd.MM.yyyy HH:mm:ss")
                }
            }
        }
        Write-Host "  ? Pronaðeno $($eventLogData.Count) EventLog zapisa" -ForegroundColor Green
    } catch {
        Write-Warning "Greška pri dohvaæanju EventLog-a: $_"
    }

    # Poboljšano dohvaæanje WMI podataka
    Write-Host "Dohvaæam WMI podatke..." -ForegroundColor Cyan
    $wmiUsers = @{}
    try {
        $wmiData = Get-CimInstance Win32_UserAccount -Filter "LocalAccount=True" -ErrorAction SilentlyContinue
        foreach ($wmiUser in $wmiData) {
            $wmiUsers[$wmiUser.Name] = $wmiUser
        }
        Write-Host "  ? Pronaðeno $($wmiUsers.Count) WMI korisnika" -ForegroundColor Green
    } catch {
        Write-Warning "Greška pri dohvaæanju WMI podataka: $_"
    }

    # Dohvaæanje lokalnih korisnika i administratora
    Write-Host "Analiziram lokalne korisnike i administratore..." -ForegroundColor Cyan
    $localUsers = Get-LocalUser
    
    # Dohvati administratore
    $adminUsers = @{}
    try {
        $adminGroupMembers = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue
        foreach ($admin in $adminGroupMembers) {
            $adminName = ($admin.Name -split '\\')[-1]
            $adminUsers[$adminName] = $true
        }
        Write-Host "  ? Pronaðeno $($adminUsers.Count) administratora" -ForegroundColor Green
    } catch {
        Write-Warning "Greška pri dohvaæanju administratora: $_"
    }

    # Poboljšano dohvaæanje net user podataka
    Write-Host "Dohvaæam net user podatke..." -ForegroundColor Cyan
    $netUserDates = @{}
    foreach ($user in $localUsers) {
        $netUserDates[$user.Name] = "Nema podataka"
        try {
            $output = net user $user.Name 2>&1
            if ($LASTEXITCODE -eq 0) {
                foreach ($line in $output) {
                    # Pokušaj razlièite regex patterns
                    if ($line -match "(?i)Password last set\s+(.+)$") {
                        $netUserDates[$user.Name] = $matches[1].Trim()
                        break
                    } elseif ($line -match "(?i)Lozinka zadnji put postavljena\s+(.+)$") {
                        $netUserDates[$user.Name] = $matches[1].Trim()
                        break
                    } elseif ($line -match "(?i)Account active\s+(.+)$") {
                        # Ako nema password data, pokušaj s account active
                        if ($netUserDates[$user.Name] -eq "Nema podataka") {
                            $netUserDates[$user.Name] = "Account info: " + $matches[1].Trim()
                        }
                    }
                }
            }
        } catch {
            Write-Verbose "Greška za korisnika $($user.Name): $_"
        }
    }

    # Poboljšano dohvaæanje registry podataka
    Write-Host "Analiziram registry i profile podatke..." -ForegroundColor Cyan
    $userData = @()
    $counter = 1

    foreach ($user in $localUsers) {
        $username = $user.Name
        $sid = $user.SID
        $isAdmin = $adminUsers.ContainsKey($username)
        
        # Registry paths
        $regPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
        $profilePath = "C:\Users\$username"
        
        # Dodatni registry putovi za bolju analizu
        $samPath = "Registry::HKEY_LOCAL_MACHINE\SAM\SAM\Domains\Account\Users"
        
        $hasFolder = Test-Path $profilePath
        $hasRegistry = Test-Path $regPath
        
        # Poboljšan status profila
        $statusProfila = if ($hasFolder -and $hasRegistry) {
            "Profil potpuno kreiran"
        } elseif ($hasFolder -or $hasRegistry) {
            if ($hasFolder) { "Samo folder postoji" } else { "Samo registry postoji" }
        } else {
            "Korisnik nikad nije logiran"
        }

        # Poboljšano dohvaæanje datuma
        $folderDate = if ($hasFolder) {
            try { 
                $folderItem = Get-Item $profilePath
                $folderItem.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") 
            } catch { "Greška pri èitanju" }
        } else { "Nema korisnièkog foldera" }

        $regDate = if ($hasRegistry) {
            try { 
                $regItem = Get-Item $regPath
                # Pokušaj dohvatiti ProfileImagePath za dodatnu provjeru
                try {
                    $profileImagePath = Get-ItemProperty -Path $regPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue
                    if ($profileImagePath) {
                        $regDate = $regItem.LastWriteTime.ToString("dd.MM.yyyy HH:mm:ss")
                    } else {
                        $regDate = $regItem.LastWriteTime.ToString("dd.MM.yyyy HH:mm:ss") + " (nepotpun)"
                    }
                } catch {
                    $regDate = $regItem.LastWriteTime.ToString("dd.MM.yyyy HH:mm:ss")
                }
                $regDate
            } catch { "Greška pri èitanju" }
        } else { "Nema u registryju" }

        # EventLog datum
        $matchKey = $eventLogData.Keys | Where-Object { $_.Trim().ToLowerInvariant() -eq $username.Trim().ToLowerInvariant() }
        $eventDate = if ($matchKey) { $eventLogData[$matchKey] } else { "Nema podataka" }

        # WMI datum
        $wmiDate = if ($wmiUsers.ContainsKey($username)) {
            $obj = $wmiUsers[$username]
            if ($obj.InstallDate) {
                try {
                    [System.Management.ManagementDateTimeConverter]::ToDateTime($obj.InstallDate).ToString("dd.MM.yyyy HH:mm:ss")
                } catch { "Greška pri èitanju" }
            } else { "Nema InstallDate" }
        } else { "Nema u WMI" }

        # Net user datum
        $netDate = $netUserDates[$username]

        # Poboljšan algoritam za odreðivanje najbolje source
        $sources = @()
        if ($eventDate -ne "Nema podataka") { $sources += @{Source="EventLog"; Date=$eventDate; Priority=1} }
        if ($wmiDate -notmatch "Nema|Greška") { $sources += @{Source="WMI"; Date=$wmiDate; Priority=2} }
        if ($netDate -ne "Nema podataka" -and $netDate -notmatch "Account info") { $sources += @{Source="Net User"; Date=$netDate; Priority=3} }
        if ($folderDate -notmatch "Nema|Greška") { $sources += @{Source="Users folder"; Date=$folderDate; Priority=4} }
        if ($regDate -notmatch "Nema|Greška") { $sources += @{Source="Registry"; Date=$regDate; Priority=5} }
        
        $izvor = if ($sources.Count -gt 0) {
            ($sources | Sort-Object Priority | Select-Object -First 1).Source
        } else {
            "Nepoznato"
        }

        # Dodatne informacije za admina
        $adminInfo = if ($isAdmin) { " (ADMIN)" } else { "" }

        $userData += [PSCustomObject]@{
            'Redni broj'        = $counter
            'Korisnièko ime'    = $username + $adminInfo
            'Status'            = if ($user.Enabled) { "Aktivan" } else { "Neaktivan" }
            'Admin'             = $isAdmin
            'Status profila'    = $statusProfila
            'Datum (Users)'     = $folderDate
            'Datum (Registry)'  = $regDate
            'Datum (EventLog)'  = $eventDate
            'Datum (WMI)'       = $wmiDate
            'Datum (Net User)'  = $netDate
            'Izvor datuma'      = $izvor
        }
        $counter++
    }

    Write-Host "? Analiza završena. Generiram HTML izvještaj..." -ForegroundColor Green

    # Generiraj HTML tablicu
    $tableRows = ""
    foreach ($u in $userData) {
        $classes = @()

        if ($u.Status -eq "Aktivan") {
            $classes += "active-user"
        } elseif ($u.Status -eq "Neaktivan") {
            $classes += "inactive"
        }
        
        if ($u.Admin) {
            $classes += "admin-user"
        }

        $classAttr = if ($classes.Count) { " class='" + ($classes -join " ") + "'" } else { "" }

        # Dodaj admin badge ako je admin
        $displayName = if ($u.Admin) { 
            ($u.'Korisnièko ime' -replace ' \(ADMIN\)', '') + " <span class='admin-badge'>ADMIN</span>"
        } else { 
            $u.'Korisnièko ime' -replace ' \(ADMIN\)', '' 
        }

        $tableRows += @"
            <tr$classAttr>
                <td>$($u.'Redni broj')</td>
                <td>$displayName</td>
                <td>$($u.Status)</td>
                <td>$($u.'Status profila')</td>
                <td>$($u.'Datum (Users)')</td>
                <td>$($u.'Datum (Registry)')</td>
                <td>$($u.'Datum (EventLog)')</td>
                <td>$($u.'Datum (WMI)')</td>
                <td>$($u.'Datum (Net User)')</td>
                <td><span class='source-badge'>$($u.'Izvor datuma')</span></td>
            </tr>
"@
    }

    $adminCount = ($userData | Where-Object { $_.Admin }).Count
    $activeCount = ($userData | Where-Object { $_.Status -eq "Aktivan" }).Count

    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Usporedba datuma kreiranja korisnièkih raèuna</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.container {
    width: 100%;
    max-width: 1400px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    overflow-x: auto;
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
    height: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 13px;
    min-width: 1200px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 10px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
    white-space: nowrap;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px 10px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    white-space: nowrap;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 80px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Status stilovi */
tr.active-user td {
    background-color: #d1fae5 !important;
}

tr.active-user:nth-child(even) td {
    background-color: #ecfdf5 !important;
}

tr.inactive td {
    background-color: #fee2e2 !important;
    color: #991b1b !important;
}

tr.inactive:nth-child(even) td {
    background-color: #fef2f2 !important;
}

/* Admin korisnici */
tr.admin-user td {
    background-color: #fef3c7 !important;
}

tr.admin-user:nth-child(even) td {
    background-color: #fde68a !important;
}

tr.admin-user:hover td {
    background-color: #fcd34d !important;
}

tr.admin-user td:first-child {
    background-color: #fde68a !important;
}

/* Badge stilovi */
.admin-badge {
    background: linear-gradient(135deg, #f59e0b, #d97706);
    color: white;
    padding: 2px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 8px;
}

.source-badge {
    background: linear-gradient(135deg, #8b5cf6, #7c3aed);
    color: white;
    padding: 3px 8px;
    border-radius: 8px;
    font-size: 11px;
    font-weight: 600;
}

/* Posebni stilovi za stupce */
td:nth-child(3) { /* Status */
    font-weight: 600;
}

td:nth-child(10) { /* Izvor datuma */
    text-align: center;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1>Korisnièki raèuni s datumom nastanka i stanjem</h1>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Korisnièko ime</th>
                        <th>Status</th>
                        <th>Status profila</th>
                        <th>Datum (Users)</th>
                        <th>Datum (Registry)</th>
                        <th>Datum (EventLog)</th>
                        <th>Datum (WMI)</th>
                        <th>Datum (Net User)</th>
                        <th>Izvor datuma</th>
                    </tr>
                </thead>
                <tbody>
                    $tableRows
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno korisnika:</b> $($userData.Count)</p>
                <p><b>Aktivnih:</b> $activeCount</p>
                <p><b>Administratora:</b> $adminCount</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
    Write-Host "? Izvještaj uspješno generiran: $OutputPath" -ForegroundColor Green
}







# === Funkcija: Izvještaj o obrisanim korisnicima (Event ID 4726) ===
function Generate-DeletedUsersReport {
    param([string]$OutputPath)

    $logName = "Security"
    $eventId = 4726
    $events = @()

    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            Id = 4726
            StartTime = (Get-Date).AddDays(-90)
        } -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Greška pri dohvaæanju EventLog-a (4726) korištenjem Get-WinEvent."
    }

    # Priprema podataka za tablicu
    $tableRows = ""
    $rowCount = 0
    
    if ($events.Count -eq 0) {
        $tableRows = "<tr><td colspan='4' style='text-align:center; color:#ef4444; font-style:italic;'>Nema pronaðenih dogaðaja brisanja korisnika u zadnjih 90 dana.</td></tr>"
    } else {
        foreach ($event in $events) {
            $rowCount++
            $eventTime = $event.TimeCreated.ToString("dd.MM.yyyy HH:mm:ss")
            $xml = [xml]$event.ToXml()
            
            # Izvuci podatke iz XML-a
            $subjectUser = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "SubjectUserName" } | Select-Object -ExpandProperty '#text'
            $targetUser = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "TargetUserName" } | Select-Object -ExpandProperty '#text'
            $subjectDomain = $xml.Event.EventData.Data | Where-Object { $_.Name -eq "SubjectDomainName" } | Select-Object -ExpandProperty '#text'
            
            # Formatiranje korisnièkih imena
            $deletedBy = if ($subjectUser -and $subjectDomain) { "$subjectDomain\$subjectUser" } else { $subjectUser }
            $deletedUser = if ($targetUser) { $targetUser } else { "Nepoznat" }
            
            $tableRows += @"
            <tr>
                <td>$rowCount</td>
                <td>$eventTime</td>
                <td><span class='deleted-user'> $deletedUser</span></td>
                <td><span class='admin-user'> $deletedBy</span></td>
            </tr>
"@
        }
    }

    $htmlContent = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Izbrisani korisnièki raèuni</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 60px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Stilovi za korisnike */
.deleted-user {
    color: #ef4444;
    font-weight: 600;
}

.admin-user {
    color: #10b981;
    font-weight: 500;
}

/* Event info box */
.event-info {
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    padding: 15px 20px;
    margin: 20px;
    border-radius: 15px;
    text-align: center;
    font-weight: 600;
    box-shadow: 0 5px 15px rgba(0,0,0,0.2);
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Ikona u naslovu */
.icon-header {
    display: inline-block;
    margin-right: 10px;
    font-size: 36px;
    vertical-align: middle;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Pulse animacija za event info */
@keyframes pulse {
    0% { transform: scale(1); }
    50% { transform: scale(1.02); }
    100% { transform: scale(1); }
}

.event-info {
    animation: pulse 2s infinite;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1><span class="icon-header"> </span>Izbrisani korisnièki raèuni</h1>
    <div class='container'>
        <div class='event-info'>
            Event ID 4726 - Brisanje korisnièkih raèuna (zadnjih 90 dana)
        </div>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>Rbr</th>
                        <th>Vrijeme dogaðaja</th>
                        <th>Izbrisani korisnik</th>
                        <th>Izbrisao</th>
                    </tr>
                </thead>
                <tbody>
                    $tableRows
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno dogaðaja:</b> $($events.Count)</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $htmlContent -Encoding UTF8
}










# === Funkcija: Izvještaj o printerima i povijesti ispisa ===
function Generate-PrinterReport {
    param([string]$OutputPath)

    Write-Host "Poèinje skupljanje podataka o printerima..." -ForegroundColor Yellow

    function Get-PrinterList {
        Write-Host "Dohvaæam listu printera..." -ForegroundColor Cyan
        try {
            $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
            $result = @()
            foreach ($printer in $printers) {
                $printerInfo = @{
                    Name = $printer.Name
                    Status = $printer.PrinterStatus
                    Type = if ($printer.Local) { "Lokalni" } elseif ($printer.Network) { "Mrežni" } else { "Nepoznato" }
                    DriverName = if ($printer.DriverName) { $printer.DriverName } else { "N/A" }
                    PortName = if ($printer.PortName) { $printer.PortName } else { "N/A" }
                    Shared = if ($printer.Shared) { "Da" } else { "Ne" }
                }
                $result += $printerInfo
                Write-Host "  ? Printer: $($printer.Name)" -ForegroundColor Green
            }
            return $result
        } catch {
            Write-Host "  ? Greška pri dohvaæanju printera: $_" -ForegroundColor Red
            return @()
        }
    }

    function Get-PrintHistory {
        Write-Host "Dohvaæam povijest printanja..." -ForegroundColor Cyan
        try {
            $log = Get-WinEvent -ListLog 'Microsoft-Windows-PrintService/Operational'
            if ($log.IsEnabled -eq $false) {
                wevtutil sl "Microsoft-Windows-PrintService/Operational" /e:true
                Start-Sleep -Seconds 2
            }

            $events = Get-WinEvent -FilterHashtable @{ 
                LogName = 'Microsoft-Windows-PrintService/Operational'; 
                ID = 307; 
                StartTime = (Get-Date).AddDays(-90)
            } -ErrorAction SilentlyContinue

            if (-not $events -or $events.Count -eq 0) {
                Write-Host "Nema zapisa o printanju u zadnjih 90 dana" -ForegroundColor Yellow
                return @()
            }

            $result = $events | ForEach-Object {
                [PSCustomObject]@{
                    Vrijeme_printanja = $_.TimeCreated.ToString("dd.MM.yyyy HH:mm:ss")
                    Korisnicko_ime    = $_.Properties[2].Value
                    Naziv_dokumenta   = if ($_.Properties[5].Value) { $_.Properties[5].Value } else { "N/A" }
                    Naziv_printera    = $_.Properties[4].Value
                    Velicina_ispisa   = if ($_.Properties[6].Value) { "$($_.Properties[6].Value) bytes" } else { "N/A" }
                    Stranica          = if ($_.Properties[7].Value) { $_.Properties[7].Value } else { "N/A" }
                }
            }

            Write-Host "  ? Pronaðeno $($result.Count) zapisa o printanju" -ForegroundColor Green
            return $result
        } catch {
            Write-Host "  ? Greška pri dohvaæanju povijesti: $_" -ForegroundColor Red
            return @()
        }
    }

    function Get-PrinterStatus {
        param([array]$PrintHistory)
        
        Write-Host "Dohvaæam status printera..." -ForegroundColor Cyan
        try {
            $status = @()
            
            # Pokušaj Performance Counters samo ako su dostupni
            try {
                $perfData = Get-CimInstance -ClassName Win32_PerfFormattedData_Spooler_PrintQueue -ErrorAction Stop |
                Where-Object { $_.Name -ne '_Total' -and $_.Name -ne '' }
                
                if ($perfData -and $perfData.Count -gt 0) {
                    $status = $perfData | ForEach-Object {
                        [PSCustomObject]@{
                            Naziv             = $_.Name
                            Trenutni_poslovi  = $_.Jobs
                            Ukupno_ispisano   = $_.TotalJobsPrinted
                            Greške            = $_.JobErrors
                            Izvor             = "Performance Counter"
                        }
                    }
                    Write-Host "  ? Pronaðeno $($status.Count) printera iz Performance Counters" -ForegroundColor Green
                }
            } catch {
                Write-Host "  ? Performance Counters nisu dostupni" -ForegroundColor Yellow
            }
            
            # Ako Performance Counters ne rade ili nema podataka, koristi kombinaciju osnovnih podataka i povijesti
            if ($status.Count -eq 0) {
                try {
                    $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
                    $spoolJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue
                    
                    # Kreiraj statistike na osnovu povijesti printanja
                    $printStats = @{}
                    if ($PrintHistory -and $PrintHistory.Count -gt 0) {
                        $printGroups = $PrintHistory | Group-Object Naziv_printera
                        foreach ($group in $printGroups) {
                            $printStats[$group.Name] = $group.Count
                        }
                    }
                    
                    $status = $printers | ForEach-Object {
                        $printerName = $_.Name
                        $currentJobs = @($spoolJobs | Where-Object { $_.Name -match [regex]::Escape($printerName) }).Count
                        $totalPrinted = if ($printStats.ContainsKey($printerName)) { $printStats[$printerName] } else { 0 }
                        
                        [PSCustomObject]@{
                            Naziv             = $printerName
                            Trenutni_poslovi  = $currentJobs
                            Ukupno_ispisano   = $totalPrinted
                            Greške            = if ($_.PrinterStatus -eq 3) { 0 } else { 0 }
                            Izvor             = "Printer Info + Print History"
                        }
                    }
                    
                    Write-Host "  ? Kreiran status za $($status.Count) printera (osnovni podaci + povijest)" -ForegroundColor Green
                } catch {
                    Write-Host "  ? Greška kod osnovnih printer podataka: $_" -ForegroundColor Red
                }
            }
            
            # Ako i dalje nema podataka, kreiraj samo iz povijesti printanja
            if ($status.Count -eq 0 -and $PrintHistory -and $PrintHistory.Count -gt 0) {
                Write-Host "  ? Kreiram status samo iz povijesti printanja..." -ForegroundColor Cyan
                
                $uniquePrinters = $PrintHistory | Group-Object Naziv_printera
                $status = $uniquePrinters | ForEach-Object {
                    [PSCustomObject]@{
                        Naziv             = $_.Name
                        Trenutni_poslovi  = 0
                        Ukupno_ispisano   = $_.Count
                        Greške            = 0
                        Izvor             = "Print History Only"
                    }
                }
                
                Write-Host "  ? Kreiran status za $($status.Count) printera iz povijesti" -ForegroundColor Green
            }
            
            return $status
            
        } catch {
            Write-Host "  ? Greška pri dohvaæanju statusa printera: $_" -ForegroundColor Red
            return @()
        }
    }

    # Dohvati sve podatke
    $printerList = Get-PrinterList
    $printHistoryData = Get-PrintHistory
    $printerStatusData = Get-PrinterStatus -PrintHistory $printHistoryData

    Write-Host "? Svi podaci o printerima skupljeni, generiram HTML..." -ForegroundColor Green

    # Generiraj HTML s modernim dizajnom
    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Izvještaj o printerima i printanju</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.main-container {
    width: 100%;
    max-width: 1400px;
}

.stats-header {
    background: rgba(255, 255, 255, 0.9);
    color: #1f2937;
    padding: 20px;
    border-radius: 20px;
    margin-bottom: 30px;
    text-align: center;
    box-shadow: 0 20px 60px rgba(0,0,0,0.2);
    backdrop-filter: blur(10px);
}

.stats-header h2 {
    margin: 0 0 15px 0;
    font-size: 24px;
    color: #667eea;
}

.stats-header p {
    margin: 10px 0;
    font-size: 16px;
    color: #6b7280;
}

.section {
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
    margin-bottom: 30px;
}

.section h2 {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 30px;
    margin: 0;
    font-size: 20px;
    font-weight: 600;
}

.table-wrapper { 
    max-height: 500px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: #f8f9fa;
    color: #1f2937;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
    border-bottom: 2px solid #e9ecef;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

.highlight { 
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    font-weight: 600;
    padding: 4px 8px; 
    border-radius: 6px;
    font-size: 12px;
}

.status-good {
    background: linear-gradient(135deg, #10b981, #059669);
    color: white;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 12px;
}

.status-warning {
    background: linear-gradient(135deg, #f59e0b, #d97706);
    color: white;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 12px;
}

.status-error {
    background: linear-gradient(135deg, #ef4444, #dc2626);
    color: white;
    padding: 4px 8px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 12px;
}

.no-data {
    text-align: center;
    padding: 40px;
    color: #6b7280;
    font-style: italic;
    background: #f9fafb;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.section {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1>Izvještaj o printerima i printanju</h1>
    
    <div class="main-container">
        <div class="stats-header">
            <h2>Informacije o sustavu</h2>
            <p>
                <strong>Raèunalo:</strong> $env:COMPUTERNAME | 
                <strong>Korisnik:</strong> $env:USERNAME | 
                <strong>Datum i vrijeme:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
            </p>
        </div>

        <div class="section">
            <h2>Instalirani printeri</h2>
"@

    # Dodaj printera ili poruku da nema printera
    if ($printerList.Count -eq 0) {
        $html += "<div class='no-data'>Nema instaliranih printera na ovom sustavu.</div>"
    } else {
        $html += @"
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Naziv printera</th>
                            <th>Status</th>
                            <th>Tip printera</th>
                            <th>Driver</th>
                            <th>Port</th>
                            <th>Dijeljeno</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        
        foreach ($printer in $printerList) {
            $statusClass = switch ($printer.Status) {
                3 { "status-good" }     # Idle
                4 { "status-good" }     # Printing
                default { "status-warning" }
            }
            
            $statusText = switch ($printer.Status) {
                1 { "Other" }
                2 { "Unknown" }
                3 { "Idle" }
                4 { "Printing" }
                5 { "Warmup" }
                default { "Status $($printer.Status)" }
            }
            
            $sharedClass = if ($printer.Shared -eq "Da") { "highlight" } else { "" }
            
            $html += @"
                        <tr>
                            <td><span class='highlight'>$($printer.Name)</span></td>
                            <td><span class='$statusClass'>$statusText</span></td>
                            <td>$($printer.Type)</td>
                            <td>$($printer.DriverName)</td>
                            <td>$($printer.PortName)</td>
                            <td><span class='$sharedClass'>$($printer.Shared)</span></td>
                        </tr>
"@
        }
        
        $html += "</tbody></table></div>"
    }

    $html += "</div><div class='section'><h2>Povijest printanja (zadnjih 90 dana)</h2>"

    # Dodaj povijest printanja
    if ($printHistoryData.Count -eq 0) {
        $html += "<div class='no-data'>Nema zapisa o printanju u zadnjih 90 dana.</div>"
    } else {
        $html += @"
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Vrijeme printanja</th>
                            <th>Korisnik</th>
                            <th>Naziv dokumenta</th>
                            <th>Printer</th>
                            <th>Velièina</th>
                            <th>Stranice</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        
        foreach ($job in $printHistoryData | Select-Object -First 20) {  # Ogranièi na 20 najnovijih
            # Oèisti korisnièko ime
            $cleanUser = $job.Korisnicko_ime -replace '.*\\', ''
            
            $html += @"
                        <tr>
                            <td><span class='highlight'>$($job.Vrijeme_printanja)</span></td>
                            <td>$cleanUser</td>
                            <td>$($job.Naziv_dokumenta)</td>
                            <td>$($job.Naziv_printera)</td>
                            <td>$($job.Velicina_ispisa)</td>
                            <td><span class='highlight'>$($job.Stranica)</span></td>
                        </tr>
"@
        }
        
        $html += "</tbody></table></div>"
        
        if ($printHistoryData.Count -gt 20) {
            $html += "<p style='text-align: center; color: #6b7280; font-style: italic; padding: 20px;'>Prikazano je 20 najnovijih zapisa od ukupno $($printHistoryData.Count) pronaðenih.</p>"
        }
    }

    $html += "</div><div class='section'><h2>Status printera</h2>"

    # Dodaj status printera
    if ($printerStatusData.Count -eq 0) {
        $html += "<div class='no-data'>Nema dostupnih podataka o statusu printera.</div>"
    } else {
        $html += @"
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Naziv</th>
                            <th>Trenutni poslovi</th>
                            <th>Ukupno ispisano</th>
                            <th>Greške</th>
                            <th>Izvor podataka</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        
        foreach ($status in $printerStatusData) {
            $errorClass = if ($status.Greške -gt 0) { "status-error" } else { "status-good" }
            $jobsClass = if ($status.Trenutni_poslovi -gt 0) { "status-warning" } else { "status-good" }
            
            $html += @"
                        <tr>
                            <td><span class='highlight'>$($status.Naziv)</span></td>
                            <td><span class='$jobsClass'>$($status.Trenutni_poslovi)</span></td>
                            <td>$($status.Ukupno_ispisano)</td>
                            <td><span class='$errorClass'>$($status.Greške)</span></td>
                            <td>$($status.Izvor)</td>
                        </tr>
"@
        }
        
        $html += "</tbody></table></div>"
    }

    # Završni dio HTML-a
    $html += @"
        </div>

        <div class='stats-container'>
            <div class='stats-box'>
                <p><strong>Generirano:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><strong>Sustav:</strong> $env:COMPUTERNAME</p>
                <p><strong>Instalirani printeri:</strong> $($printerList.Count)</p>
                <p><strong>Zapisi printanja:</strong> $($printHistoryData.Count)</p>
                <p><strong>Status zapisi:</strong> $($printerStatusData.Count)</p>
                <p><strong>Izrada:</strong> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
    
    Write-Host "Moderan izvještaj o printerima uspješno generiran!" -ForegroundColor Green
    Write-Host "Sadrži:" -ForegroundColor Cyan
    Write-Host "   • $($printerList.Count) instaliranih printera" -ForegroundColor Gray
    Write-Host "   • $($printHistoryData.Count) zapisa o printanju" -ForegroundColor Gray
    Write-Host "   • $($printerStatusData.Count) status zapisa" -ForegroundColor Gray
}







# Funkcija za dohvat èlanova lokalne grupe 'Administrators' i generiranje HTML izvještaja
function Generate-AdminGroupReport {
    param([string]$OutputPath)
    
    # Funkcija za dohvaæanje podataka o istjeku lozinke
    function Get-PasswordExpiry {
        param([string]$Username)
        try {
            $netUserOutput = & net user $Username 2>$null
            if ($LASTEXITCODE -eq 0) {
                $passwordExpiresLine = $netUserOutput | Where-Object { $_ -match "Password expires" }
                if ($passwordExpiresLine) {
                    $passwordExpires = ($passwordExpiresLine -split "\s{2,}")[1]
                    return $passwordExpires
                }
            }
            return "N/A"
        }
        catch {
            return "N/A"
        }
    }
    
    # Dohvat èlanova grupe 'Administrators'
    $groupName = "Administrators"
    $group = Get-WmiObject -Class Win32_Group -Filter "Name='$groupName'"
    $members = $group.GetRelated("Win32_Account")
    
    # Priprema podataka za HTML tablicu
    $i = 1
    $tableRows = ""
    foreach ($member in $members) {
        $status = if ($member.Disabled -eq $true) { "Neaktivan" } else { "Aktivan" }
        $statusClass = if ($member.Disabled -eq $true) { "status-inactive" } else { "status-active" }
        
        # Dohvati podatke o istjeku lozinke
        $passwordExpiry = Get-PasswordExpiry -Username $member.Name
        
        # Definiraj CSS klasu za istijek lozinke
        $expiryClass = ""
        if ($passwordExpiry -match "Never") {
            $expiryClass = "expiry-never"
        } elseif ($passwordExpiry -ne "N/A") {
            try {
                $expiryDate = [DateTime]::Parse($passwordExpiry)
                $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
                if ($daysUntilExpiry -le 7) {
                    $expiryClass = "expiry-critical"
                } elseif ($daysUntilExpiry -le 30) {
                    $expiryClass = "expiry-warning"
                } else {
                    $expiryClass = "expiry-ok"
                }
            } catch {
                $expiryClass = "expiry-unknown"
            }
        } else {
            $expiryClass = "expiry-unknown"
        }
        
        $tableRows += @"
            <tr>
                <td>$i</td>
                <td>$($member.Name)</td>
                <td>$($member.FullName)</td>
                <td><code>$($member.SID)</code></td>
                <td>$($member.Description)</td>
                <td class="$statusClass">$status</td>
                <td class="$expiryClass">$passwordExpiry</td>
            </tr>
"@
        $i++
    }
    
    # HTML
    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Popis èlanova lokalne grupe Administrators</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.container {
    width: 100%;
    max-width: 1400px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 60px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

/* Status stilovi */
.status-active {
    color: #10b981;
    font-weight: 600;
}

.status-inactive {
    color: #ef4444;
    font-weight: 600;
}

/* Stilovi za istijek lozinke */
.expiry-critical {
    color: #dc2626;
    font-weight: 600;
    background-color: #fee2e2;
    padding: 4px 8px;
    border-radius: 6px;
}

.expiry-warning {
    color: #d97706;
    font-weight: 600;
    background-color: #fef3c7;
    padding: 4px 8px;
    border-radius: 6px;
}

.expiry-ok {
    color: #059669;
    font-weight: 600;
}

.expiry-never {
    color: #6b7280;
    font-weight: 600;
    font-style: italic;
}

.expiry-unknown {
    color: #9ca3af;
    font-style: italic;
}

code {
    background: #f1f5f9;
    padding: 2px 6px;
    border-radius: 4px;
    font-family: 'Courier New', monospace;
    font-size: 11px;
    color: #475569;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Ikona u naslovu */
.icon-header {
    display: inline-block;
    margin-right: 10px;
    font-size: 36px;
    vertical-align: middle;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}

/* Admin badge */
.admin-badge {
    display: inline-block;
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    padding: 2px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    margin-left: 5px;
}
</style>
</head>
<body>
    <h1><span class="icon-header"></span>Korisnici s administratorskim ovlastima</h1>
    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Ime</th>
                        <th>Puno ime</th>
                        <th>SID</th>
                        <th>Opis</th>
                        <th>Status</th>
                        <th>Istijek lozinke</th>
                    </tr>
                </thead>
                <tbody>
                    $tableRows
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno èlanova:</b> $($members.Count)</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@
    
    # Spremi HTML
    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}


# Funkcija instalirani programi
function Generate-InstalledProgramsReport {
    param([string]$OutputPath)

    function Get-InstalledPrograms {
        $registryPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        $programs = foreach ($path in $registryPaths) {
            Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -and
                $_.UninstallString -and
                $_.SystemComponent -ne 1 -and
                $_.ReleaseType -ne "Update" -and
                $_.ReleaseType -ne "Hotfix" -and
                $_.ParentKeyName -eq $null
            } | Select-Object @{Name="Name";Expression={$_.DisplayName}},
                              @{Name="Publisher";Expression={$_.Publisher}},
                              @{Name="InstalledOn";Expression={
    try {
        if ($_.InstallDate -match '^\d{8}$') {
            [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null).ToString('dd.MM.yyyy')
        }
        elseif ($_.InstallLocation -and (Test-Path $_.InstallLocation)) {
            (Get-Item $_.InstallLocation).CreationTime.ToString('dd.MM.yyyy')
        }
        elseif ($_.DisplayIcon) {
            $iconPath = $_.DisplayIcon -split ',' | Select-Object -First 1
            if (Test-Path $iconPath) {
                (Get-Item $iconPath).CreationTime.ToString('dd.MM.yyyy')
            }
            else {
                $folder = Split-Path $iconPath -Parent
                if (Test-Path $folder) {
                    (Get-Item $folder).CreationTime.ToString('dd.MM.yyyy')
                } else { '' }
            }
        }
        else {
            ''
        }
    } catch {
        ''
    }
}},
                              @{Name="Size";Expression={
                                  if ($_.EstimatedSize) {
                                      "{0:N2} MB" -f ($_.EstimatedSize / 1024)
                                  } else {
                                      "n/a"
                                  }
                              }},
                              @{Name="Version";Expression={$_.DisplayVersion}}
        }

        return $programs | Sort-Object Name
    }

    $programs = Get-InstalledPrograms

    if ($programs.Count -eq 0) {
        $programTable = @"
            <div style='padding: 40px; text-align: center; color: #6b7280; font-style: italic;'>
                Nema pronaðenih instaliranih programa.
            </div>
"@
    } else {
        $i = 1
        $tableRows = ""
        foreach ($program in $programs) {
            $tableRows += @"
                <tr>
                    <td>$i</td>
                    <td>$($program.Name)</td>
                    <td>$($program.Publisher)</td>
                    <td>$($program.InstalledOn)</td>
                    <td>$($program.Size)</td>
                    <td>$($program.Version)</td>
                </tr>
"@
            $i++
        }
        
        $programTable = @"
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Naziv</th>
                        <th>Izdavaè</th>
                        <th>Instalirano</th>
                        <th>Velièina</th>
                        <th>Verzija</th>
                    </tr>
                </thead>
                <tbody>
                    $tableRows
                </tbody>
            </table>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Instalirani Programi</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 600px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

td:first-child, th:first-child { 
    position: sticky; 
    left: 0; 
    background-color: #f5f5f5;
    font-weight: 600;
    color: #667eea;
    width: 60px;
    text-align: center;
    z-index: 5;
}

tr:hover td:first-child {
    background-color: #e8eaf6;
}

thead th:first-child {
    background: linear-gradient(135deg, #8e94f2, #9a9ff9);
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1>Popis Instaliranih Programa</h1>
    <div class='container'>
        <div class='table-wrapper'>
            $programTable
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Ukupno programa:</b> $($programs.Count)</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    Set-Content -Path $OutputPath -Value $html -Encoding UTF8
}










# Generiranje HTML tablice
function Generate-SerialNumberReport {
    param([string]$OutputPath)

    Write-Host "Poèinje detaljno skupljanje serijskih brojeva..." -ForegroundColor Yellow

    # === HARD DISKOVI S POBOLJŠANIM USB SERIAL DETECTION ===
    Write-Host "Dohvaæam diskove s poboljšanim USB detekcijom..." -ForegroundColor Cyan
    try {
        $drives = Get-WmiObject Win32_DiskDrive
        $hardDrives = @()
        
        foreach ($drive in $drives) {
            # Odredi tip diska
            $diskType = "Tvrdi disk"
            if ($drive.InterfaceType -eq 'USB') {
                $diskType = 'USB/Vanjski disk'
            } elseif ($drive.Model -match 'SSD|Solid') {
                $diskType = 'SSD'
            }
            
            # POBOLJŠANA DETEKCIJA USB SERIJSKIH BROJEVA
            $serialNumber = "N/A"
            if ($drive.SerialNumber) {
                $rawSerial = $drive.SerialNumber.Trim()
                
                # Provjeri ima li èudne znakove (kao g€)
                if ($rawSerial -match '[^\x20-\x7E]' -or $rawSerial.Length -lt 3) {
                    # Pokušaj alternativnu metodu za USB
                    if ($drive.InterfaceType -eq 'USB' -and $drive.PNPDeviceID) {
                        # Izvuci serijski iz PNP Device ID
                        if ($drive.PNPDeviceID -match '\\([^\\]+)$') {
                            $pnpSerial = $matches[1]
                            if ($pnpSerial -and $pnpSerial.Length -gt 3) {
                                $serialNumber = $pnpSerial
                            }
                        }
                    }
                    
                    # Ako i dalje nema, pokušaj preko Win32_Volume
                    if ($serialNumber -eq "N/A" -or $serialNumber -match '[^\x20-\x7E]') {
                        try {
                            $volumes = Get-WmiObject Win32_Volume | Where-Object { $_.DriveLetter -and $_.DriveType -eq 2 }
                            foreach ($vol in $volumes) {
                                if ($vol.SerialNumber) {
                                    $volSerial = [Convert]::ToString($vol.SerialNumber, 16).ToUpper()
                                    if ($volSerial -and $volSerial.Length -gt 2) {
                                        $serialNumber = $volSerial
                                        break
                                    }
                                }
                            }
                        } catch { }
                    }
                } else {
                    $serialNumber = $rawSerial
                }
            }
            
            # Dohvati slova particija
            $letters = @()
            try {
                $partitions = Get-WmiObject Win32_DiskPartition | Where-Object { $_.DiskIndex -eq $drive.Index }
                foreach ($partition in $partitions) {
                    $logicalDisks = Get-WmiObject Win32_LogicalDiskToPartition | Where-Object { $_.Antecedent -match $partition.DeviceID.Replace('\', '\\') }
                    foreach ($logical in $logicalDisks) {
                        $drive_letter = $logical.Dependent -replace '.*="(.*?)".*', '$1'
                        $letters += $drive_letter
                    }
                }
            } catch { }
            
            $diskInfo = @{
                Tip = $diskType
                Model = if ($drive.Model) { $drive.Model.Trim() } else { "N/A" }
                SerijskiBroj = $serialNumber
                Slovo = if ($letters.Count -gt 0) { $letters -join ', ' } else { "N/A" }
                Velièina = if ($drive.Size) { "$([math]::round($drive.Size / 1GB, 2)) GB" } else { "N/A" }
                Firmware = if ($drive.FirmwareRevision) { $drive.FirmwareRevision } else { "N/A" }
            }
            $hardDrives += $diskInfo
            Write-Host "  ? Disk: $($diskInfo.Model) SN: $($diskInfo.SerijskiBroj)" -ForegroundColor Green
        }
    } catch {
        $hardDrives = @()
        Write-Host "  ? Greška pri dohvaæanju diskova: $_" -ForegroundColor Red
    }

    # === RAÈUNALO SERIJSKI BROJEVI ===
    Write-Host "Dohvaæam informacije o raèunalu..." -ForegroundColor Cyan
    try {
        $bios = Get-WmiObject Win32_BIOS
        $computer = Get-WmiObject Win32_ComputerSystem
        $computerModel = if ($computer.Model) { $computer.Model.Trim() } else { "N/A" }
        $computerManufacturer = if ($computer.Manufacturer) { $computer.Manufacturer.Trim() } else { "N/A" }
        $computerSerial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "N/A" }
        
        # Kombiniraj proizvoðaè i model
        $computerFullModel = "$computerManufacturer $computerModel".Trim()
        if ($computerFullModel -eq "N/A N/A") { $computerFullModel = "N/A" }
    } catch {
        $computerFullModel = "N/A"
        $computerSerial = "Greška pri dohvaæanju"
    }

    # === MONITORI ===
    Write-Host "Dohvaæam informacije o monitorima..." -ForegroundColor Cyan
    $monitorInfo = @()
    try {
        $wmiMonitors = Get-WmiObject -Namespace root/wmi -Class WmiMonitorID -ErrorAction SilentlyContinue
        foreach ($monitor in $wmiMonitors) {
            $manufacturerName = "N/A"
            $productName = "N/A"
            $serialNumber = "N/A"
            
            if ($monitor.ManufacturerName) {
                $manufacturerName = [System.Text.Encoding]::ASCII.GetString($monitor.ManufacturerName).Trim([char]0)
            }
            if ($monitor.UserFriendlyName) {
                $productName = [System.Text.Encoding]::ASCII.GetString($monitor.UserFriendlyName).Trim([char]0)
            }
            if ($monitor.SerialNumberID) {
                $serialNumber = [System.Text.Encoding]::ASCII.GetString($monitor.SerialNumberID).Trim([char]0)
            }
            
            $monitorModel = "$manufacturerName $productName".Trim()
            if ($monitorModel -eq "N/A N/A") { $monitorModel = "Monitor" }
            
            $monitorInfo += @{
                Model = $monitorModel
                SerialNumber = $serialNumber
            }
        }
        
        if ($monitorInfo.Count -eq 0) {
            $monitorInfo = @(@{ Model = "Monitor"; SerialNumber = "N/A" })
        }
    } catch {
        $monitorInfo = @(@{ Model = "Monitor"; SerialNumber = "Greška pri dohvaæanju" })
    }

    # === GRAFIÈKE KARTICE ===
    Write-Host "Dohvaæam informacije o grafièkim karticama..." -ForegroundColor Cyan
    try {
        $gpus = Get-WmiObject Win32_VideoController | Where-Object { $_.Name -notmatch "Basic|Generic" }
        $gpuInfo = @()
        foreach ($gpu in $gpus) {
            $gpuModel = if ($gpu.Name) { $gpu.Name.Trim() } else { "Grafièka kartica" }
            $gpuSerial = if ($gpu.PNPDeviceID) { $gpu.PNPDeviceID } else { "N/A" }
            
            $gpuInfo += @{
                Model = $gpuModel
                SerialNumber = $gpuSerial
            }
        }
        if ($gpuInfo.Count -eq 0) { 
            $gpuInfo = @(@{ Model = "Grafièka kartica"; SerialNumber = "N/A" })
        }
    } catch {
        $gpuInfo = @(@{ Model = "Grafièka kartica"; SerialNumber = "Greška pri dohvaæanju" })
    }

    # === MATIÈNA PLOÈA ===
    Write-Host "Dohvaæam informacije o matiènoj ploèi..." -ForegroundColor Cyan
    try {
        $baseboard = Get-WmiObject Win32_BaseBoard
        $motherboardModel = "$($baseboard.Manufacturer) $($baseboard.Product)".Trim()
        if ($motherboardModel -eq " ") { $motherboardModel = "Matièna ploèa" }
        $motherboardSerial = if ($baseboard.SerialNumber) { $baseboard.SerialNumber.Trim() } else { "N/A" }
    } catch {
        $motherboardModel = "Matièna ploèa"
        $motherboardSerial = "Greška pri dohvaæanju"
    }

    # === PROCESOR ===
    Write-Host "Dohvaæam informacije o procesoru..." -ForegroundColor Cyan
    try {
        $cpus = Get-WmiObject Win32_Processor
        $cpuInfo = @()
        foreach ($processor in $cpus) {
            $cpuModel = if ($processor.Name) { $processor.Name.Trim() } else { "Procesor" }
            $cpuSerial = if ($processor.ProcessorId) { $processor.ProcessorId } else { "N/A" }
            
            $cpuInfo += @{
                Model = $cpuModel
                SerialNumber = $cpuSerial
            }
        }
    } catch {
        $cpuInfo = @(@{ Model = "Procesor"; SerialNumber = "Greška pri dohvaæanju" })
    }

    # === RAM ===
    Write-Host "Dohvaæam informacije o RAM memoriji..." -ForegroundColor Cyan
    try {
        $ram = Get-WmiObject Win32_PhysicalMemory
        $ramInfo = @()
        foreach ($module in $ram) {
            $ramModel = "RAM"
            if ($module.Manufacturer -and $module.Manufacturer.Trim() -ne "") {
                $ramModel = $module.Manufacturer.Trim()
            }
            if ($module.PartNumber -and $module.PartNumber.Trim() -ne "") {
                $ramModel += " " + $module.PartNumber.Trim()
            }
            
            $ramDetails = @{
                Model = $ramModel
                Kapacitet = if ($module.Capacity) { [math]::round($module.Capacity / 1GB, 2) } else { "N/A" }
                Tip = if ($module.MemoryType) { 
                    switch ($module.MemoryType) {
                        20 { "DDR" }
                        21 { "DDR2" }
                        24 { "DDR3" }
                        26 { "DDR4" }
                        34 { "DDR5" }
                        default { "Unknown" }
                    }
                } else { "N/A" }
                Brzina = if ($module.Speed) { $module.Speed } else { "N/A" }
                SerijskiBroj = if ($module.SerialNumber) { $module.SerialNumber.Trim() } else { "N/A" }
            }
            $ramInfo += $ramDetails
        }
    } catch {
        $ramInfo = @()
    }

    # === MREŽNE KARTICE ===
    Write-Host "Dohvaæam informacije o mrežnim karticama..." -ForegroundColor Cyan
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapter | Where-Object { 
            $_.PhysicalAdapter -eq $true -and 
            $_.Name -notmatch "Virtual|Loopback|Teredo|isatap"
        }
        
        $nicInfo = @()
        foreach ($adapter in $adapters) {
            $nicModel = if ($adapter.Name) { $adapter.Name.Trim() } else { "Mrežna kartica" }
            $nicSerial = if ($adapter.MACAddress) { $adapter.MACAddress } else { "N/A" }
            
            $nicInfo += @{
                Model = $nicModel
                SerialNumber = $nicSerial
            }
        }
    } catch {
        $nicInfo = @(@{ Model = "Mrežna kartica"; SerialNumber = "Greška pri dohvaæanju" })
    }

    # === OPERATIVNI SUSTAV ===
    Write-Host "Dohvaæam informacije o operativnom sustavu..." -ForegroundColor Cyan
    try {
        $os = Get-WmiObject Win32_OperatingSystem
        $osModel = if ($os.Caption) { $os.Caption.Trim() } else { "Windows" }
        if ($os.Version) { $osModel += " " + $os.Version }
        
        $osInstallDate = "N/A"
        if ($os.InstallDate) {
            try {
                $osInstallDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate).ToString("dd.MM.yyyy HH:mm:ss")
            } catch { }
        }
    } catch {
        $osModel = "Operativni sustav"
        $osInstallDate = "Greška pri dohvaæanju"
    }

# === ANTIVIRUSNI SOFTVER S POBOLJŠANIM WINDOWS DEFENDER DETEKCIJOM ===
    Write-Host "Dohvaæam informacije o antivirusnom softveru..." -ForegroundColor Cyan
    try {
        $antivirus = Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct -ErrorAction SilentlyContinue
        $antivirusInfo = @()
        
        if ($antivirus) {
            foreach ($av in $antivirus) {
                $avModel = if ($av.displayName) { $av.displayName.Trim() } else { "Antivirusni softver" }
                
                # Posebno rukovanje za Windows Defender
                if ($avModel -match "Windows Defender|Microsoft Defender") {
                    try {
                        # Pokušaj dohvatiti verziju i datum preko Get-MpComputerStatus
                        $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
                        if ($defenderStatus) {
                            $avVersion = $defenderStatus.AMProductVersion
                            if (-not $avVersion) {
                                $avVersion = $defenderStatus.AntimalwareVersion
                            }
                            
                            # Dohvati datum izdavanja
                            $avReleased = "N/A"
                            if ($defenderStatus.AntivirusSignatureLastUpdated) {
                                $avReleased = $defenderStatus.AntivirusSignatureLastUpdated.ToString("M/d/yyyy h:mm:ss tt")
                            } elseif ($defenderStatus.AntispywareSignatureLastUpdated) {
                                $avReleased = $defenderStatus.AntispywareSignatureLastUpdated.ToString("M/d/yyyy h:mm:ss tt")
                            }
                        }
                        
                        # Alternativni naèin ako Get-MpComputerStatus ne radi
                        if (-not $avVersion) {
                            try {
                                $defenderPath = "${env:ProgramFiles}\Windows Defender\MpCmdRun.exe"
                                if (Test-Path $defenderPath) {
                                    $versionInfo = (Get-ItemProperty $defenderPath).VersionInfo
                                    $avVersion = $versionInfo.ProductVersion
                                    
                                    # Pokušaj dohvatiti datum datoteke
                                    if (-not $avReleased -or $avReleased -eq "N/A") {
                                        $fileInfo = Get-ItemProperty $defenderPath
                                        if ($fileInfo.LastWriteTime) {
                                            $avReleased = $fileInfo.LastWriteTime.ToString("M/d/yyyy h:mm:ss tt")
                                        }
                                    }
                                }
                            } catch { }
                        }
                        
                        # Još jedan pokušaj preko registra
                        if (-not $avVersion) {
                            try {
                                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows Defender"
                                if (Test-Path $regPath) {
                                    $defenderReg = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                                    if ($defenderReg.EngineVersion) {
                                        $avVersion = $defenderReg.EngineVersion
                                    }
                                }
                            } catch { }
                        }
                        
                        # Finalni pokušaj preko WMI
                        if (-not $avVersion) {
                            try {
                                $defenderWMI = Get-WmiObject -Namespace "root\Microsoft\SecurityClient" -Class AntimalwareInfectionStatus -ErrorAction SilentlyContinue
                                if ($defenderWMI) {
                                    $avVersion = $defenderWMI.ProductVersion
                                }
                            } catch { }
                        }
                        
                    } catch {
                        $avVersion = "N/A"
                        $avReleased = "N/A"
                    }
                }
                # Za ostale antiviruse
                else {
                    $avVersion = if ($av.pathToSignedProductExe) { 
                        try {
                            $version = (Get-ItemProperty $av.pathToSignedProductExe -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
                            if ($version) { $version } else { "N/A" }
                        } catch { "N/A" }
                    } else { "N/A" }
                    
                    # Za ostale antiviruse, pokušaj dohvatiti datum iz exe datoteke
                    $avReleased = "N/A"
                    if ($av.pathToSignedProductExe) {
                        try {
                            $fileInfo = Get-ItemProperty $av.pathToSignedProductExe -ErrorAction SilentlyContinue
                            if ($fileInfo.LastWriteTime) {
                                $avReleased = $fileInfo.LastWriteTime.ToString("M/d/yyyy h:mm:ss tt")
                            }
                        } catch { }
                    }
                }
                
                $antivirusInfo += @{
                    Model = $avModel
                    Version = if ($avVersion) { $avVersion } else { "N/A" }
                    Released = if ($avReleased) { $avReleased } else { "N/A" }
                }
            }
        } else {
            # Ako nema detektiranih antivirusa, provjeri postoji li Windows Defender
            try {
                $defenderService = Get-Service -Name "WinDefend" -ErrorAction SilentlyContinue
                if ($defenderService) {
                    $defenderVersion = "N/A"
                    $defenderReleased = "N/A"
                    
                    # Pokušaj dohvatiti verziju i datum
                    try {
                        $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
                        if ($defenderStatus) {
                            $defenderVersion = $defenderStatus.AMProductVersion
                            if (-not $defenderVersion) {
                                $defenderVersion = $defenderStatus.AntimalwareVersion
                            }
                            
                            # Dohvati datum izdavanja
                            if ($defenderStatus.AntivirusSignatureLastUpdated) {
                                $defenderReleased = $defenderStatus.AntivirusSignatureLastUpdated.ToString("M/d/yyyy h:mm:ss tt")
                            } elseif ($defenderStatus.AntispywareSignatureLastUpdated) {
                                $defenderReleased = $defenderStatus.AntispywareSignatureLastUpdated.ToString("M/d/yyyy h:mm:ss tt")
                            }
                        }
                    } catch { }
                    
                    $antivirusInfo = @(@{ 
                        Model = "Windows Defender"; 
                        Version = $defenderVersion;
                        Released = $defenderReleased
                    })
                } else {
                    $antivirusInfo = @(@{ Model = "Nema antivirusa"; Version = "N/A"; Released = "N/A" })
                }
            } catch {
                $antivirusInfo = @(@{ Model = "Windows Defender"; Version = "N/A"; Released = "N/A" })
            }
        }
    } catch {
        $antivirusInfo = @(@{ Model = "Antivirusni softver"; Version = "Greška pri dohvaæanju"; Released = "N/A" })
    }

    # === BIOS ===
    try {
        $bios = Get-WmiObject Win32_BIOS
        $biosModel = "BIOS"
        if ($bios.Manufacturer) { $biosModel = $bios.Manufacturer.Trim() + " BIOS" }
        $biosVersion = if ($bios.SMBIOSBIOSVersion) { $bios.SMBIOSBIOSVersion } else { "N/A" }
        if ($bios.ReleaseDate) {
            try {
                $biosDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate).ToString("dd.MM.yyyy")
                $biosVersion += ", Datum: $biosDate"
            } catch { }
        }
    } catch {
        $biosModel = "BIOS"
        $biosVersion = "Greška pri dohvaæanju"
    }

    Write-Host "? Svi podaci skupljeni, generiram HTML s modernim dizajnom..." -ForegroundColor Green

    # === HTML TEMPLATE ===
    $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Serijski brojevi komponenti raèunala</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

h1 { 
    color: white;
    font-size: 32px;
    margin-bottom: 30px;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    letter-spacing: 1px;
    text-align: center;
}

.stats-header {
    background: rgba(255, 255, 255, 0.9);
    color: #1f2937;
    padding: 20px;
    border-radius: 20px;
    margin-bottom: 30px;
    text-align: center;
    box-shadow: 0 20px 60px rgba(0,0,0,0.2);
    backdrop-filter: blur(10px);
    max-width: 1200px;
    width: 100%;
}

.stats-header h2 {
    margin: 0 0 15px 0;
    font-size: 24px;
    color: #667eea;
}

.stats-header p {
    margin: 10px 0;
    font-size: 16px;
    color: #6b7280;
}

.container {
    width: 100%;
    max-width: 1200px;
    background: rgba(255, 255, 255, 0.95);
    border-radius: 20px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    overflow: hidden;
    backdrop-filter: blur(10px);
}

.table-wrapper { 
    max-height: 700px; 
    overflow-y: auto; 
    margin: 20px;
}

.table-wrapper::-webkit-scrollbar {
    width: 10px;
}

.table-wrapper::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

.table-wrapper::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea, #764ba2);
    border-radius: 10px;
}

table { 
    width: 100%; 
    border-collapse: collapse;
    font-size: 14px;
}

thead th { 
    position: sticky; 
    top: 0; 
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 15px 12px;
    text-align: left;
    font-weight: 600;
    letter-spacing: 0.5px;
    z-index: 10;
}

th:first-child {
    border-top-left-radius: 10px;
}

th:last-child {
    border-top-right-radius: 10px;
}

td { 
    padding: 12px;
    border-bottom: 1px solid #e0e0e0;
    color: #333;
    vertical-align: top;
}

tr:nth-child(even) td { 
    background-color: #f8f9fa; 
}

tr:hover td { 
    background-color: #e8eaf6;
    transition: background-color 0.3s ease;
}

.model-name {
    font-weight: 600;
    color: #1f2937;
    margin-right: 10px;
}

.serial-highlight { 
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    font-weight: 600;
    padding: 4px 8px; 
    border-radius: 6px;
    font-size: 12px;
    margin-left: 5px;
}

.component-type {
    background: linear-gradient(135deg, #8b5cf6, #7c3aed);
    color: white;
    font-weight: 600;
    padding: 6px 10px;
    border-radius: 8px;
    display: inline-block;
    margin-bottom: 5px;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.stats-container {
    display: flex;
    justify-content: center;
    padding: 30px;
}

.stats-box {
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    padding: 20px 40px;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    text-align: center;
    transform: translateY(0);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.stats-box:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 40px rgba(0,0,0,0.3);
}

.stats-box p {
    margin: 5px 0;
    font-size: 14px;
}

.stats-box b {
    color: #fff;
    font-weight: 600;
}

/* Animacija uèitavanja */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.container {
    animation: fadeIn 0.6s ease-out;
}

/* Responsive dizajn */
@media (max-width: 768px) {
    h1 { font-size: 24px; }
    .table-wrapper { margin: 10px; }
    td, th { padding: 8px; font-size: 12px; }
    .stats-box { padding: 15px 25px; }
}
</style>
</head>
<body>
    <h1>Serijski brojevi komponenti raèunala</h1>
    
    <div class="stats-header">
        <h2>Informacije o sustavu</h2>
        <p>
            <strong>Raèunalo:</strong> $env:COMPUTERNAME | 
            <strong>Korisnik:</strong> $env:USERNAME | 
            <strong>Datum i vrijeme:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
        </p>
    </div>

    <div class='container'>
        <div class='table-wrapper'>
            <table>
                <thead>
                    <tr><th>Komponenta</th><th>Model i serijski broj</th></tr>
                </thead>
                <tbody>
"@

    # === DODAJ PODATKE U NOVOM FORMATU ===
    $htmlContent = ""

    # Raèunalo
    $htmlContent += "<tr><td><span class='component-type'>Raèunalo</span></td><td><span class='model-name'>$computerFullModel</span><span class='serial-highlight'>$computerSerial</span></td></tr>"

    # Monitori
    foreach ($monitor in $monitorInfo) {
        $htmlContent += "<tr><td><span class='component-type'>Monitor</span></td><td><span class='model-name'>$($monitor.Model)</span><span class='serial-highlight'>$($monitor.SerialNumber)</span></td></tr>"
    }

    # Grafièke kartice
    foreach ($gpu in $gpuInfo) {
        $htmlContent += "<tr><td><span class='component-type'>Grafièka kartica</span></td><td><span class='model-name'>$($gpu.Model)</span><span class='serial-highlight'>$($gpu.SerialNumber)</span></td></tr>"
    }

    # Matièna ploèa
    $htmlContent += "<tr><td><span class='component-type'>Matièna ploèa</span></td><td><span class='model-name'>$motherboardModel</span><span class='serial-highlight'>$motherboardSerial</span></td></tr>"

    # Procesori
    foreach ($processor in $cpuInfo) {
        $htmlContent += "<tr><td><span class='component-type'>Procesor</span></td><td><span class='model-name'>$($processor.Model)</span><span class='serial-highlight'>$($processor.SerialNumber)</span></td></tr>"
    }

    # RAM moduli
    foreach ($ram in $ramInfo) {
        $ramDetail = "<span class='model-name'>$($ram.Model) ($($ram.Kapacitet) GB $($ram.Tip) $($ram.Brzina) MHz)</span><span class='serial-highlight'>$($ram.SerijskiBroj)</span>"
        $htmlContent += "<tr><td><span class='component-type'>RAM</span></td><td>$ramDetail</td></tr>"
    }

    # Hard diskovi
    foreach ($disk in $hardDrives) {
        $diskDetail = "<span class='model-name'>$($disk.Model) [$($disk.Velièina)]</span><span class='serial-highlight'>$($disk.SerijskiBroj)</span>"
        $htmlContent += "<tr><td><span class='component-type'>$($disk.Tip)</span></td><td>$diskDetail</td></tr>"
    }

    # Mrežne kartice
    foreach ($nic in $nicInfo) {
        $htmlContent += "<tr><td><span class='component-type'>Mrežna kartica</span></td><td><span class='model-name'>$($nic.Model)</span><span class='serial-highlight'>$($nic.SerialNumber)</span></td></tr>"
    }

    # Operativni sustav
    $htmlContent += "<tr><td><span class='component-type'>Operativni sustav</span></td><td><span class='model-name'>$osModel</span><span style='margin-left: 10px; color: #6b7280;'>Instaliran: $osInstallDate</span></td></tr>"

    # Antivirusni softver
    foreach ($av in $antivirusInfo) {
        $htmlContent += "<tr><td><span class='component-type'>Antivirusni softver</span></td><td><span class='model-name'>$($av.Model)</span><span style='margin-left: 10px; color: #6b7280;'>Verzija: $($av.Version)</span></td></tr>"
    }

    # BIOS
    $htmlContent += "<tr><td><span class='component-type'>BIOS</span></td><td><span class='model-name'>$biosModel</span><span style='margin-left: 10px; color: #6b7280;'>$biosVersion</span></td></tr>"

    # === ZAVRŠNI DIO ===
    $htmlFooter = @"
                </tbody>
            </table>
        </div>
        <div class='stats-container'>
            <div class='stats-box'>
                <p><b>Generirano:</b> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><b>Sustav:</b> $env:COMPUTERNAME</p>
                <p><b>Diskova:</b> $($hardDrives.Count)</p>
                <p><b>RAM modula:</b> $($ramInfo.Count)</p>
                <p><b>Izrada:</b> $((Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName -replace '^$', $env:USERNAME), $(Get-Date -Format 'yyyy')</p>
            </div>
        </div>
    </div>
</body>
</html>
"@

    # Kreiraj finalni HTML
    $finalHTML = $html + $htmlContent + $htmlFooter
    Set-Content -Path $OutputPath -Value $finalHTML -Encoding UTF8
}


# === Funkcija za generiranje Info taba - GUI kompatibilna verzija ===
function Generate-InfoTab {
    param([string]$OutputPath)

    # Provjeri da li je putanja zadana
    if ([string]::IsNullOrEmpty($OutputPath)) {
        Write-Error "OutputPath parametar je obavezan!"
        return
    }

    try {
        # Prikupi informacije potrebne za template
        $currentDate = Get-Date -Format 'dd.MM.yyyy'
        $currentDateTime = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
        $buildInfo = Get-Date -Format 'yyyyMMdd-HHmm'
        $computerName = $env:COMPUTERNAME
        $userName = $env:USERNAME
        
        # Pokušaj dobiti puno ime korisnika
        try {
            $fullName = (Get-WmiObject -Class Win32_UserAccount -Filter "Name='$env:USERNAME' AND LocalAccount='True'").FullName
            if ([string]::IsNullOrEmpty($fullName)) {
                $fullName = $env:USERNAME
            }
        }
        catch {
            $fullName = $env:USERNAME
        }
        
        # PowerShell verzija
        $psVersion = $PSVersionTable.PSVersion.ToString()
        $netVersion = [System.Environment]::Version.ToString()

        $html = @"
<!DOCTYPE html>
<html lang='hr'>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <meta charset='UTF-8'>
    <title>Get-PCInfo - Informacije</title>
<style>
* { 
    margin: 0; 
    padding: 0; 
    box-sizing: border-box; 
}

body { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
    background: #667eea;
    background: -webkit-linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    background: -moz-linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    background: -o-linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    color: #ffffff;
    padding: 40px 20px;
    line-height: 1.6;
}

.container {
    max-width: 1200px;
    margin: 0 auto;
    background: rgba(255, 255, 255, 0.15);
    border-radius: 25px;
    padding: 50px;
    border: 1px solid rgba(255, 255, 255, 0.3);
    -webkit-box-shadow: 0 20px 60px rgba(0, 0, 0, 0.4);
    -moz-box-shadow: 0 20px 60px rgba(0, 0, 0, 0.4);
    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.4);
}

h1 { 
    color: #ffffff; 
    text-align: center; 
    font-size: 42px; 
    margin-bottom: 15px;
    text-shadow: 3px 3px 6px rgba(0, 0, 0, 0.4);
    font-weight: 700;
    letter-spacing: 1px;
}

.subtitle {
    text-align: center;
    font-size: 18px;
    color: rgba(255, 255, 255, 0.9);
    margin-bottom: 40px;
    font-style: italic;
}

.version-info {
    background: #10b981;
    background: -webkit-linear-gradient(135deg, #10b981, #059669);
    background: -moz-linear-gradient(135deg, #10b981, #059669);
    background: -o-linear-gradient(135deg, #10b981, #059669);
    background: linear-gradient(135deg, #10b981, #059669);
    padding: 30px;
    border-radius: 20px;
    margin: 30px 0;
    text-align: center;
    -webkit-box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    -moz-box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    border: 1px solid rgba(255, 255, 255, 0.2);
}

.version-info h2 {
    font-size: 28px;
    margin-bottom: 15px;
    font-weight: 600;
}

.version-info p {
    font-size: 16px;
    margin: 8px 0;
    opacity: 0.95;
}

h2 { 
    color: #ffffff; 
    font-size: 28px; 
    margin-top: 40px; 
    margin-bottom: 20px;
    border-left: 5px solid #10b981;
    padding-left: 20px;
    font-weight: 600;
}

h3 { 
    color: #f1f5f9; 
    font-size: 20px; 
    margin-top: 25px; 
    margin-bottom: 15px;
    font-weight: 500;
}

p { 
    font-size: 16px; 
    margin-bottom: 18px; 
    text-align: justify;
    color: rgba(255, 255, 255, 0.95);
    line-height: 1.7;
}

.feature {
    background: rgba(255, 255, 255, 0.15);
    padding: 25px;
    margin: 20px 0;
    border-radius: 15px;
    border-left: 5px solid #10b981;
    -webkit-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    -webkit-transition: all 0.3s ease;
    -moz-transition: all 0.3s ease;
    -o-transition: all 0.3s ease;
    transition: all 0.3s ease;
}

.feature:hover {
    -webkit-transform: translateY(-3px);
    -moz-transform: translateY(-3px);
    -ms-transform: translateY(-3px);
    -o-transform: translateY(-3px);
    transform: translateY(-3px);
    -webkit-box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
    -moz-box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.3);
}

.feature h3 {
    background: #fbbf24;
    background: -webkit-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: -moz-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: -o-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    font-size: 22px;
    margin-bottom: 15px;
    font-weight: 600;
    padding: 12px 20px;
    border-radius: 12px;
    text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
    -webkit-box-shadow: 0 4px 12px rgba(251, 191, 36, 0.3);
    -moz-box-shadow: 0 4px 12px rgba(251, 191, 36, 0.3);
    box-shadow: 0 4px 12px rgba(251, 191, 36, 0.3);
    display: block;
    text-align: center;
}

.stats-grid {
    width: 100%;
    display: table;
    margin: 30px 0;
}

.stats-row {
    display: table-row;
}

.stat-card {
    display: table-cell;
    background: rgba(255, 255, 255, 0.15);
    padding: 25px;
    border-radius: 15px;
    text-align: center;
    -webkit-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.1);
    margin: 10px;
    width: 16.66%;
    vertical-align: top;
    -webkit-transition: all 0.3s ease;
    -moz-transition: all 0.3s ease;
    -o-transition: all 0.3s ease;
    transition: all 0.3s ease;
}

.stat-card:hover {
    -webkit-transform: translateY(-5px);
    -moz-transform: translateY(-5px);
    -ms-transform: translateY(-5px);
    -o-transform: translateY(-5px);
    transform: translateY(-5px);
    -webkit-box-shadow: 0 15px 30px rgba(0, 0, 0, 0.3);
    -moz-box-shadow: 0 15px 30px rgba(0, 0, 0, 0.3);
    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.3);
}

.stat-card h3 {
    background: #667eea;
    background: -webkit-linear-gradient(135deg, #667eea, #764ba2);
    background: -moz-linear-gradient(135deg, #667eea, #764ba2);
    background: -o-linear-gradient(135deg, #667eea, #764ba2);
    background: linear-gradient(135deg, #667eea, #764ba2);
    color: white;
    font-size: 18px;
    margin-bottom: 15px;
    font-weight: 600;
    padding: 10px 15px;
    border-radius: 10px;
    text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
    -webkit-box-shadow: 0 3px 8px rgba(102, 126, 234, 0.3);
    -moz-box-shadow: 0 3px 8px rgba(102, 126, 234, 0.3);
    box-shadow: 0 3px 8px rgba(102, 126, 234, 0.3);
}

.stat-card p {
    font-size: 14px;
    line-height: 1.6;
    text-align: center;
}

.highlight {
    background: #fbbf24;
    background: -webkit-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: -moz-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: -o-linear-gradient(135deg, #fbbf24, #f59e0b);
    background: linear-gradient(135deg, #fbbf24, #f59e0b);
    color: white;
    padding: 3px 8px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 14px;
    -webkit-box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
}

ul {
    list-style: none;
    padding-left: 0;
}

li {
    margin: 12px 0;
    padding-left: 30px;
    position: relative;
    font-size: 15px;
    line-height: 1.6;
}

li:before {
    content: "?";
    position: absolute;
    left: 0;
    color: #ffffff;
    font-weight: bold;
    font-size: 12px;
    background: #10b981;
    background: -webkit-linear-gradient(135deg, #10b981, #059669);
    background: -moz-linear-gradient(135deg, #10b981, #059669);
    background: -o-linear-gradient(135deg, #10b981, #059669);
    background: linear-gradient(135deg, #10b981, #059669);
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: block;
    text-align: center;
    line-height: 20px;
    top: 2px;
    -webkit-box-shadow: 0 2px 8px rgba(16, 185, 129, 0.4);
    -moz-box-shadow: 0 2px 8px rgba(16, 185, 129, 0.4);
    box-shadow: 0 2px 8px rgba(16, 185, 129, 0.4);
    border: 2px solid rgba(255, 255, 255, 0.3);
}

.functions-grid {
    width: 100%;
    margin: 30px 0;
}

.functions-row {
    width: 100%;
    display: table;
    margin-bottom: 25px;
}

.function-card {
    display: table-cell;
    background: rgba(255, 255, 255, 0.1);
    padding: 20px;
    border-radius: 15px;
    border: 1px solid rgba(255, 255, 255, 0.1);
    margin: 0 12px;
    width: 48%;
    vertical-align: top;
    -webkit-transition: all 0.3s ease;
    -moz-transition: all 0.3s ease;
    -o-transition: all 0.3s ease;
    transition: all 0.3s ease;
}

.function-card:hover {
    background: rgba(255, 255, 255, 0.15);
    -webkit-transform: translateY(-2px);
    -moz-transform: translateY(-2px);
    -ms-transform: translateY(-2px);
    -o-transform: translateY(-2px);
    transform: translateY(-2px);
    -webkit-box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2);
    box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2);
}

.function-card h4 {
    color: #fbbf24;
    font-size: 16px;
    margin-bottom: 8px;
    font-weight: 600;
    text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
}

.function-card p {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.9);
    line-height: 1.5;
    text-align: left;
}

.tech-specs {
    width: 100%;
    display: table;
    margin: 30px 0;
}

.tech-row {
    display: table-row;
}

.tech-card {
    display: table-cell;
    background: rgba(16, 185, 129, 0.2);
    padding: 25px;
    border-radius: 15px;
    text-align: center;
    -webkit-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    -moz-box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    border: 1px solid rgba(16, 185, 129, 0.3);
    margin: 10px;
    width: 25%;
    vertical-align: top;
    -webkit-transition: all 0.3s ease;
    -moz-transition: all 0.3s ease;
    -o-transition: all 0.3s ease;
    transition: all 0.3s ease;
}

.tech-card:hover {
    -webkit-transform: translateY(-3px);
    -moz-transform: translateY(-3px);
    -ms-transform: translateY(-3px);
    -o-transform: translateY(-3px);
    transform: translateY(-3px);
}

.tech-card h3 {
    color: #ffffff;
    font-size: 18px;
    margin-bottom: 15px;
    font-weight: 600;
}

.tech-card p {
    font-size: 14px;
    line-height: 1.6;
    text-align: center;
}

.footer {
    text-align: center;
    margin-top: 50px;
    padding: 30px;
    background: rgba(0, 0, 0, 0.3);
    border-radius: 15px;
    border: 1px solid rgba(255, 255, 255, 0.1);
}

.footer p {
    font-size: 14px;
    margin: 8px 0;
    color: rgba(255, 255, 255, 0.8);
}

.footer strong {
    color: #10b981;
    font-weight: 600;
}

/* Responzivni dizajn za male ekrane */
@media (max-width: 768px) {
    .container {
        padding: 30px 20px;
    }
    
    h1 {
        font-size: 32px;
    }
    
    .stat-card, .function-card, .tech-card {
        display: block;
        width: 100% !important;
        margin: 10px 0;
    }
    
    .stats-grid, .tech-specs {
        display: block;
    }
    
    .functions-row, .stats-row, .tech-row {
        display: block;
    }
}
</style>
</head>
<body>
<div class='container'>
    <h1>Get-PCInfo</h1>
    <p class='subtitle'>Napredni sustav za upravljanje IT infrastrukturom</p>

    <div class='version-info'>
        <h2>Verzija: 2025.6.3 Professional</h2>
        <p><strong>Datum izdanja:</strong> 13.6.2025</p>
        <p><strong>Autor:</strong> Ivica Rašan</p>
        <p><strong>Platforma:</strong> PowerShell 5.1+ / Windows 10/11</p>
    </div>

    <h2>Opis aplikacije</h2>
    <p>
        <span class='highlight'>Get-PCInfo</span> predstavlja najnapredniji PowerShell-based sustav za upravljanje IT infrastrukturom, 
        dizajniran posebno za potrebe modernih IT odjela. Aplikacija kombinira moguænost PowerShell-a s intuitivnim grafièkim suèeljem, 
        omoguæujuæi sistemskim administratorima potpunu kontrolu nad lokalnim korisnièkim raèunima, hardverskim resursima, 
        sigurnosnim komponentama i sustavskim dogaðajima.
    </p>
    
    <p>
        Kroz elegantno dizajnirano suèelje s tabovima, aplikacija pruža centralizirani pristup svim kritiènim informacijama o raèunalu, 
        automatski generirajuæi profesionalne HTML izvještaje s modernim glassmorphism dizajnom. Svaki izvještaj je personaliziran 
        s podacima o korisniku koji ga je kreirao, što osigurava potpunu transparentnost i odgovornost.
    </p>

    <h2>Kompletna lista funkcionalnosti</h2>
    <div class='functions-grid'>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-UserReport</h4>
                <p>Detaljan pregled svih lokalnih korisnièkih raèuna s informacijama o statusu, lozinkama, grupama i SID-ovima</p>
            </div>
            <div class='function-card'>
                <h4>Generate-GroupUserReport</h4>
                <p>Analiza lokalnih grupa i njihovih èlanova s detaljnim informacijama o korisnièkim ulogama</p>
            </div>
        </div>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-LogonHTMLReport</h4>
                <p>Praæenje korisnièkih prijava i odjava kroz EventLog s analizom sesija i trajanja</p>
            </div>
            <div class='function-card'>
                <h4>Generate-InactiveUsersReport</h4>
                <p>Identifikacija neaktivnih korisnièkih raèuna s detaljnim statusom lozinki i administratorskim pravima</p>
            </div>
        </div>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-FolderPermissionsReport</h4>
                <p>Analiza prava pristupa folderima s detaljnim prikazom dozvola za organizacijske strukture</p>
            </div>
            <div class='function-card'>
                <h4>Generate-USBUserReport</h4>
                <p>Praæenje USB aktivnosti s identifikacijom korisnika i vremenskim okvirom korištenja</p>
            </div>
        </div>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-KreiranjeKRComparisonReport</h4>
                <p>Usporedba datuma kreiranja korisnièkih raèuna iz razlièitih izvora (EventLog, WMI, Registry)</p>
            </div>
            <div class='function-card'>
                <h4>Generate-DeletedUsersReport</h4>
                <p>Praæenje izbrisanih korisnièkih raèuna kroz EventLog s informacijama o administratoru</p>
            </div>
        </div>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-PrinterReport</h4>
                <p>Analiza instaliranih printera, povijesti printanja i statusu redova za ispis</p>
            </div>
            <div class='function-card'>
                <h4>Generate-AdminGroupReport</h4>
                <p>Pregled èlanova lokalne Administrators grupe s detaljnim informacijama o privilegijama</p>
            </div>
        </div>
        <div class='functions-row'>
            <div class='function-card'>
                <h4>Generate-InstalledProgramsReport</h4>
                <p>Popis instaliranih programa s verzijama, proizvoðaèima i datumima instalacije</p>
            </div>
            <div class='function-card'>
                <h4>Generate-SerialNumberReport</h4>
                <p>Prikupljanje serijskih brojeva svih hardverskih komponenti s poboljšanom USB detekcijom</p>
            </div>
        </div>
    </div>

    <h2>Kljuène znaèajke</h2>
    <div class='feature'>
        <h3>Upravljanje korisnièkim raèunima</h3>
        <ul>
            <li>Pregled svih lokalnih korisnika s detaljnim informacijama o statusu</li>
            <li>Analiza lozinki: datum postavljanja, istek i sigurnosni status</li>
            <li>Praæenje posljednje prijave i aktivnosti korisnika</li>
            <li>Identifikacija administratorskih raèuna i privilegija</li>
            <li>Analiza èlanstva u grupama i korisnièkih uloga</li>
            <li>Detekcija neaktivnih i onemoguæenih raèuna</li>
        </ul>
    </div>

    <div class='feature'>
        <h3>Sigurnost i nadzor</h3>
        <ul>
            <li>Kontinuirano praæenje korisnièkih prijava kroz Windows EventLog</li>
            <li>Analiza USB aktivnosti s identifikacijom korisnika i ureðaja</li>
            <li>Praæenje administratorskih aktivnosti i promjena</li>
            <li>Nadzor brisanja korisnièkih raèuna s vremenom i izvorima</li>
            <li>Analiza prava pristupa folderima i datotekama</li>
            <li>Detekcija sigurnosnih anomalija i neobiènih aktivnosti</li>
        </ul>
    </div>

    <div class='feature'>
        <h3>Sistemski resursi i hardver</h3>
        <ul>
            <li>Detaljan popis hardverskih komponenti (CPU, RAM, GPU, storage)</li>
            <li>Prikupljanje serijskih brojeva s naprednom USB detekcijom</li>
            <li>BIOS informacije i datum instalacije operativnog sustava</li>
            <li>Analiza instaliranih programa i softverskih komponenti</li>
            <li>Pregled mrežnih adaptera i konfiguracija</li>
            <li>Status i konfiguracija instaliranih printera</li>
        </ul>
    </div>

    <div class='feature'>
        <h3>Napredni izvještaji</h3>
        <ul>
            <li>Automatsko generiranje HTML izvještaja s modernim dizajnom</li>
            <li>Personalizirani izvještaji s podacima o kreatoru</li>
            <li>Glassmorphism dizajn s smooth animacijama</li>
            <li>Responzivni dizajn prilagoðen svim ureðajima</li>
            <li>Detaljne statistike i metrike za svaki aspekt sustava</li>
            <li>Moguænost izvoza u razlièite formate</li>
        </ul>
    </div>

    <h2>Tehnièke specifikacije</h2>
    <div class='tech-specs'>
        <div class='tech-row'>
            <div class='tech-card'>
                <h3>Platforme</h3>
                <p>Windows 10 (1909+)<br>Windows 11<br>Windows Server 2016+<br>Windows Server 2019/2022</p>
            </div>
            <div class='tech-card'>
                <h3>Sistemski zahtjevi</h3>
                <p>PowerShell 5.1+<br>.NET Framework 4.5+<br>WMI/CIM podrška<br>EventLog pristup</p>
            </div>
            <div class='tech-card'>
                <h3>Ovisnosti</h3>
                <p>System.Windows.Forms<br>System.Drawing<br>Microsoft.PowerShell.Utility<br>CimCmdlets</p>
            </div>
            <div class='tech-card'>
                <h3>Privilegije</h3>
                <p>Administratorski pristup<br>SeSecurityPrivilege<br>SeBackupPrivilege<br>Local Group Policy</p>
            </div>
        </div>
    </div>

    <h2>Statistike sustava</h2>
    <div class='stats-grid'>
        <div class='stats-row'>
            <div class='stat-card'>
                <h3>Funkcije</h3>
                <p><strong>12</strong><br>Unificiranih funkcija</p>
            </div>
            <div class='stat-card'>
                <h3>Izvještaji</h3>
                <p><strong>12</strong><br>Tipova izvještaja</p>
            </div>
            <div class='stat-card'>
                <h3>Podaci</h3>
                <p><strong>100+</strong><br>Tipova podataka</p>
            </div>
            <div class='stat-card'>
                <h3>Dizajn</h3>
                <p><strong>Glassmorphism</strong><br>Moderni CSS3</p>
            </div>
            <div class='stat-card'>
                <h3>Responzivnost</h3>
                <p><strong>100%</strong><br>Svi ureðaji</p>
            </div>
            <div class='stat-card'>
                <h3>Performanse</h3>
                <p><strong>Optimizirano</strong><br>Brzo uèitavanje</p>
            </div>
        </div>
    </div>

    <h2>Prednosti korištenja</h2>
    <div class='feature'>
        <ul>
            <li><strong>Vremenska efikasnost:</strong> Automatizacija složenih administrativnih zadataka</li>
            <li><strong>Centralizirani pristup:</strong> Sve informacije dostupne iz jednog suèelja</li>
            <li><strong>Sigurnosni nadzor:</strong> Kontinuirano praæenje korisnièkih aktivnosti</li>
            <li><strong>Profesionalni izvještaji:</strong> Vizualno atraktivni dokumenti za management</li>
            <li><strong>Skalabilnost:</strong> Prilagodljiv razlièitim velièinama organizacija</li>
            <li><strong>Pouzdanost:</strong> Stabilno rješenje temeljeno na PowerShell-u</li>
            <li><strong>Transparentnost:</strong> Svi izvještaji su personalizirani s podacima o kreatoru</li>
            <li><strong>Modernizacija:</strong> Suvremeni dizajn i tehnologije</li>
        </ul>
    </div>

    <h2>Naèin korištenja</h2>
    <p>
        Aplikacija se pokreæe s <span class='highlight'>administratorskim privilegijama</span> putem PowerShell ISE ili 
        PowerShell konzole. Sustav automatski detektira dostupne komponente i generirajte izvještaje za sve module. 
        Svaki izvještaj se otvara u novom prozoru preglednika s moguænostima pretraživanja, sortiranja i izvoza podataka.
    </p>
    
    <p>
        Korisnici mogu koristiti sustav za redovite sigurnosne provjere, inventarizaciju hardvera, nadzor korisnièkih 
        aktivnosti ili generiranje izvještaja za compliance zahtjeve. Svi podaci se prikupljaju u stvarnom vremenu 
        direktno iz Windows sustava.
    </p>

    <h2>Buduæe nadogradnje</h2>
    <div class='feature'>
        <ul>
            <li>Integracija s Active Directory službama</li>
            <li>Automatsko slanje izvještaja putem e-maila</li>
            <li>Moguænost scheduliranja izvještaja</li>
            <li>Integracija s SIEM sustavima</li>
            <li>Mobilna aplikacija za pregled izvještaja</li>
            <li>Napredna analitika i machine learning</li>
        </ul>
    </div>

    <h2>Podrška i kontakt</h2>
    <p>
        Za tehnièku podršku, prijedloge poboljšanja ili prijave grešaka, molimo kontaktirajte razvojni tim. 
        Aplikacija se kontinuirano razvija s novim funkcionalnostima i poboljšanjima na temelju korisnièkih 
        povratnih informacija i tehnoloških trendova.
    </p>

    <div class='footer'>
        <p>
            <strong>© 2025 Ivica Rašan</strong><br>
            Sva prava pridržana. Ova aplikacija je razvijena za potrebe IT administracije.<br>
            <em>Build: $buildInfo | PowerShell $psVersion | .NET $netVersion</em><br>
            <strong>Generirano na:</strong> $computerName | $currentDateTime
        </p>
    </div>
</div>
</body>
</html>
"@


        # Zapišemo HTML u datoteku
        Set-Content -Path $OutputPath -Value $html -Encoding UTF8
        Write-Host "Info tab uspješno generiran: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Greška prilikom generiranja Info taba: $($_.Exception.Message)"
        Write-Error "Detalji: $($_.Exception.StackTrace)"
    }
}

# Test funkcija - ukloni komentar za testiranje
# Generate-InfoTab -OutputPath "C:\temp\test-info-gui.html"



Update-LoadingStatus "Priprema tabova..." 30


# === Generiraj izvješæe i popuni tabove ===
Update-LoadingStatus "Kreiranje tab-a 'Status lokalnih korisnièkih raèuna'..." 40
$outputFile = "$env:TEMP\LocalUsersReport.html"
Write-PCInfoLog "Generating Users Report..."
Generate-UserReport -OutputPath $outputFile
Write-PCInfoLog "Users Report completed"
Add-TabWeb "Korisnici" $outputFile

# === Generiraj HTML za Grupe i korisnike i dodaj tab ===
Update-LoadingStatus "Kreiranje tab-a 'Korisnièke grupe s èlanstvom'..." 50
$groupReportFile = "$env:TEMP\LocalGroupsAndUsers.html"
Generate-GroupUserReport -OutputPath $groupReportFile
Add-TabWeb "Grupe" $groupReportFile

# === Logiranja i odjave ===
Update-LoadingStatus "Kreiranje tab-a 'Zapis aktivnosti korisnika'..." 60
$logonData = Get-LogonHistory -Days 30
$logonPath = "$env:TEMP\logon_report.html"
Generate-LogonHTMLReport -LogonData $logonData -ReportPath $logonPath
Add-TabWeb "Logiranja i odjave" $logonPath

# === Neaktivni korisnici ===
Update-LoadingStatus "Kreiranje tab-a 'Popis neaktivnih korisnièkih raèuna'..." 60
$inactivePath = "$env:TEMP\inactive_users.html"
Generate-InactiveUsersReport -OutputPath $inactivePath
Add-TabWeb "Neaktivni korisnici" $inactivePath

# === Folder Permissions ===
Update-LoadingStatus "Kreiranje tab-a 'Folder Permissions / Prava pristupa'..." 60
$folderReportPath = "$env:TEMP\folder_permissions.html"
Generate-FolderPermissionsReport -TargetDrive "D:\" -OutputPath $folderReportPath
Add-TabWeb "Folder Permissions" $folderReportPath

# === Korištenje USB ===
Update-LoadingStatus "Kreiranje tab-a 'Korištenje USB ureðaja'..." 60
$tempUSBUserReport = "$env:TEMP\usb_report.html"
Generate-USBUserReport -OutputPath $tempUSBUserReport
Add-TabWeb "Korištenje USB" $tempUSBUserReport

# === Kreiranje KR  ===
Update-LoadingStatus "Kreiranje tab-a 'Korisnièki raèuni s datumom nastanka i stanjem'..." 60
$usporedbaPath = "$env:TEMP\kreiranje_kr_usporedba.html"
Generate-KreiranjeKRComparisonReport -OutputPath $usporedbaPath
Add-TabWeb "Kreiranje KR" $usporedbaPath

# === Obrisani KR ===
Update-LoadingStatus "Kreiranje tab-a 'Izbrisani korisnièki raèuni'..." 60
$deletedReportPath = "$env:TEMP\deleted_users_report.html"
Generate-DeletedUsersReport -OutputPath $deletedReportPath
Add-TabWeb "Obrisani KR" $deletedReportPath

# === Print ===
Update-LoadingStatus "Kreiranje tab-a 'Izvještaj o printerima i printanju'..." 60
$printerReportPath = "$env:TEMP\printer_report.html"
Generate-PrinterReport -OutputPath $printerReportPath
Add-TabWeb "Print" $printerReportPath

# === Administratori ===
Update-LoadingStatus "Kreiranje tab-a 'Korisnici s administratorskim ovlastima'..." 60
$adminReportPath = "$env:TEMP\admin_group_report.html"
Generate-AdminGroupReport -OutputPath $adminReportPath
Add-TabWeb "Administratori" $adminReportPath

# === Serijski brojevi ===
Update-LoadingStatus "Kreiranje tab-a 'Serijski brojevi komponenti raèunala'..." 60
$adminReportPath = "$env:TEMP\Serijski_brojevi.html"
Generate-SerialNumberReport -OutputPath $adminReportPath
Add-TabWeb "Serijski brojevi" $adminReportPath

# === Instalirani programi ===
Update-LoadingStatus "Kreiranje tab-a 'Popis instaliranih programa'..." 60
$adminReportPath = "$env:TEMP\InstalledPrograms.html"
Generate-InstalledProgramsReport -OutputPath $adminReportPath
Add-TabWeb "Instalirani programi" $adminReportPath

# === Generiraj Info tab ===
Update-LoadingStatus "Kreiranje tab-a 'Info Get-PCInfo'..." 60
$infoTabPath = "$env:TEMP\info_tab.html"
Generate-InfoTab -OutputPath $infoTabPath
Add-TabWeb "Info" $infoTabPath


# PRVO - definiraj originalWidth PRIJE Add_Shown
$originalWidth = $form.Width

# === Reakcija na poèetno uèitavanje forme ===
Update-LoadingStatus "Finaliziranje korisnièkog suèelja..." 90
$form.Add_Shown({
    try {
        if ($tabControl.SelectedTab -ne $null) {
            $tabName = $tabControl.SelectedTab.Text
            
            Write-Host "DEBUG: Poèetni tab je: '$tabName'" -ForegroundColor Yellow
            
            # Zapamti staru širinu
            $oldWidth = $form.Width
            
            switch ($tabName) {
                "Korisnici"                 { $form.Width = 1620 }
                "Grupe"                     { $form.Width = 1300 }
                "Logiranja i odjave"        { $form.Width = 1300 }
                "Neaktivni korisnici"       { $form.Width = 1400 }
                "Folder Permissions"        { $form.Width = 1300 }
                "Korištenje USB"            { $form.Width = 1400 }
                "Kreiranje KR"              { $form.Width = 1400 }
                "Obrisani KR"               { $form.Width = 1400 }
                "Print"                     { $form.Width = 1400 }
                "Administratori"            { $form.Width = 1400 }
                "Serijski brojevi"          { $form.Width = 1400 }
                "Popis direktorija"         { $form.Width = 1400 }
                "Hardver"                   { $form.Width = 1400 }
                "Kreiranje"                 { $form.Width = 1400 }
                "Instalirani programi"      { $form.Width = 1400 }
                "Info"                      { $form.Width = 1400 } 
                default                     { $form.Width = $originalWidth }
            }
            
            # RE-CENTRIRAJ FORMU NAKON MIJENJANJA ŠIRINE
            if ($form.Width -ne $oldWidth -and $form.Visible) {
                Start-Sleep -Milliseconds 50
                
                $screen = [System.Windows.Forms.Screen]::PrimaryScreen
                if ($screen -and $screen.WorkingArea) {
                    $newLeft = [Math]::Max(0, ($screen.WorkingArea.Width - $form.Width) / 2)
                    if ($newLeft -ge 0 -and $newLeft -lt $screen.WorkingArea.Width) {
                        $form.Left = $newLeft
                    }
                }
            }
            
            Write-Host "DEBUG: Širina postavljena na: $($form.Width)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "DEBUG: Greška pri centriranju: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Dodaj transparentnost fix
    $form.Opacity = 1.0
    $form.BringToFront()
})

# === Reakcija na promjenu taba ===
$tabControl.Add_SelectedIndexChanged({
    try {
        if ($tabControl.SelectedTab -ne $null) {
            $tabName = $tabControl.SelectedTab.Text
            
            Write-Host "DEBUG: Prebacio na tab: '$tabName'" -ForegroundColor Cyan
            
            # Zapamti staru širinu
            $oldWidth = $form.Width
            
            switch ($tabName) {
                "Korisnici"                 { $form.Width = 1620 }
                "Grupe"                     { $form.Width = 1300 }
                "Logiranja i odjave"        { $form.Width = 1300 }
                "Neaktivni korisnici"       { $form.Width = 1400 }
                "Folder Permissions"        { $form.Width = 1300 }
                "Korištenje USB"            { $form.Width = 1400 }
                "Kreiranje KR"              { $form.Width = 1400 }
                "Obrisani KR"               { $form.Width = 1400 }
                "Print"                     { $form.Width = 1400 }
                "Administratori"            { $form.Width = 1400 }
                "Serijski brojevi"          { $form.Width = 1400 }
                "Popis direktorija"         { $form.Width = 1400 }
                "Hardver"                   { $form.Width = 1400 }
                "Kreiranje"                 { $form.Width = 1400 }
                "Instalirani programi"      { $form.Width = 1400 }
                "Info"                      { $form.Width = 1400 } 
                default                     { $form.Width = $originalWidth }
            }
            
            # RE-CENTRIRAJ FORMU NAKON MIJENJANJA ŠIRINE
            if ($form.Width -ne $oldWidth -and $form.Visible) {
                $screen = [System.Windows.Forms.Screen]::PrimaryScreen
                if ($screen -and $screen.WorkingArea) {
                    $newLeft = [Math]::Max(0, ($screen.WorkingArea.Width - $form.Width) / 2)
                    if ($newLeft -ge 0 -and $newLeft -lt $screen.WorkingArea.Width) {
                        $form.Left = $newLeft
                    }
                }
            }
            
            Write-Host "DEBUG: Nova širina: $($form.Width)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "DEBUG: Greška pri centriranju na tab promjeni: $($_.Exception.Message)" -ForegroundColor Red
    }
})



# === Panel s gumbima (FlowLayoutPanel) ===
$panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$panelButtons.Height = 50
$panelButtons.Dock = "Bottom"
$panelButtons.FlowDirection = "LeftToRight"
$panelButtons.WrapContents = $false
$panelButtons.AutoScroll = $true
$panelButtons.Padding = '10,10,10,10'
$panelButtons.BackColor = [System.Drawing.Color]::Gainsboro

# === Gumbi ===
$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Text = "Spremi kao TXT"
$btnExportTxt.Size = New-Object System.Drawing.Size(120, 30)
$btnExportTxt.Add_Click({ Export-ActiveTabEnhanced -format "txt" })

$btnExportHTML = New-Object System.Windows.Forms.Button
$btnExportHTML.Text = "Spremi kao HTML"
$btnExportHTML.Size = New-Object System.Drawing.Size(140, 30)
$btnExportHTML.Add_Click({ Export-ActiveTabEnhanced -format "html" })

$btnExportDoc = New-Object System.Windows.Forms.Button
$btnExportDoc.Text = "Spremi kao DOC"
$btnExportDoc.Size = New-Object System.Drawing.Size(120, 30)
$btnExportDoc.Add_Click({ Export-ActiveTabEnhanced -format "doc" })

$btnExportExcel = New-Object System.Windows.Forms.Button
$btnExportExcel.Text = "Spremi kao Excel"
$btnExportExcel.Size = New-Object System.Drawing.Size(120, 30)
$btnExportExcel.Add_Click({ Export-ActiveTabEnhanced -format "xls" })

$btnExportCSV = New-Object System.Windows.Forms.Button
$btnExportCSV.Text = "Spremi kao CSV"
$btnExportCSV.Size = New-Object System.Drawing.Size(120, 30)
$btnExportCSV.Add_Click({ Export-ActiveTabEnhanced -format "csv" })

$btnExportPDF = New-Object System.Windows.Forms.Button
$btnExportPDF.Text = "Spremi kao PDF"
$btnExportPDF.Size = New-Object System.Drawing.Size(120, 30)
$btnExportPDF.Add_Click({ Export-ActiveTabEnhanced -format "pdf" })

$btnPrint = New-Object System.Windows.Forms.Button
$btnPrint.Text = "Print"
$btnPrint.Size = New-Object System.Drawing.Size(100, 30)
$btnPrint.Add_Click({ Print-ActiveTab })

Write-PCInfoLog "Enhanced export functions loaded successfully"

# === Info gumb ===
$btnInfo = New-Object System.Windows.Forms.Button
$btnInfo.Text = "Info"
$btnInfo.Size = New-Object System.Drawing.Size(100, 30)
$btnInfo.BackColor = [System.Drawing.Color]::LightBlue
$btnInfo.Add_Click({ 
    # Prebaci na Info tab
    for ($i = 0; $i -lt $tabControl.TabPages.Count; $i++) {
        if ($tabControl.TabPages[$i].Text -eq "Info") {
            $tabControl.SelectedIndex = $i
            break
        }
    }
})


# === Dodavanje gumbi u panel ===
$panelButtons.Controls.AddRange(@(
    $btnExportTxt,
    $btnExportCSV,
    $btnExportDoc,
    $btnExportExcel,
    $btnExportPDF,
    $btnExportHTML,
    $btnPrint,
    $btnInfo              # INFO GUMB
))
$form.Controls.Add($panelButtons)

function Generate-ReportWithLogging {
    param(
        [string]$ReportName,
        [scriptblock]$GenerationFunction,
        [string]$OutputPath
    )
    
    try {
        Write-PCInfoLog "Starting generation: $ReportName"
        Show-ExportProgress "Generating $ReportName..."
        
        $startTime = Get-Date
        & $GenerationFunction -OutputPath $OutputPath
        $endTime = Get-Date
        
        $duration = ($endTime - $startTime).TotalSeconds
        Write-PCInfoLog "$ReportName completed in $([math]::Round($duration, 2)) seconds"
        
        Hide-ExportProgress
        return $true
    }
    catch {
        Write-PCInfoLog "Failed to generate $ReportName`: $($_.Exception.Message)" -Level "ERROR"
        Hide-ExportProgress
        throw
    }
}


$btns = @($btnExportTxt, $btnExportDoc, $btnExportExcel, $btnExportPDF, $btnExportCSV, $btnExportHTML)
$x = 10
foreach ($btn in $btns) {
    $btn.Location = New-Object System.Drawing.Point($x, 10)
    $panelButtons.Controls.Add($btn)
    $x += $btn.Width + 10
}

# === Dodavanje kontrola na formu ===
$form.Controls.Add($tabControl)
$form.Controls.Add($panelButtons)

# === STAVITI OVAJ KOD NEPOSREDNO PRIJE $form.ShowDialog() ===

# === PROFESIONALNE BOJE ===
$Colors = @{
    Primary = [System.Drawing.Color]::FromArgb(41, 128, 185)     # Plava
    Secondary = [System.Drawing.Color]::FromArgb(52, 73, 94)     # Tamno siva
    Accent = [System.Drawing.Color]::FromArgb(231, 76, 60)       # Crvena
    Success = [System.Drawing.Color]::FromArgb(46, 204, 113)     # Zelena
    Warning = [System.Drawing.Color]::FromArgb(241, 196, 15)     # Žuta
    Light = [System.Drawing.Color]::FromArgb(236, 240, 241)      # Svijetlo siva
    Dark = [System.Drawing.Color]::FromArgb(44, 62, 80)          # Tamna
    White = [System.Drawing.Color]::White
    Purple = [System.Drawing.Color]::FromArgb(155, 89, 182)      # Ljubièasta
}

Write-Host "Primjenjujem profesionalni styling..." -ForegroundColor Yellow

# === STILIZIRANJE GLAVNE FORME ===
$form.BackColor = $Colors.Light
$form.Text = "Get-PCInfo - Profesionalna verzija | © Ivica Rašan 2025"

# === STILIZIRANJE TAB CONTROL-a ===
$tabControl.BackColor = $Colors.Light
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# === STILIZIRANJE BUTTON PANEL-a ===
$panelButtons.BackColor = $Colors.Secondary
$panelButtons.Height = 60

# === STILIZIRANJE POSTOJEÆIH GUMBOVA ===
# TXT gumb
$btnExportTxt.Text = "Spremi kao TXT"
$btnExportTxt.BackColor = $Colors.Primary
$btnExportTxt.ForeColor = $Colors.White
$btnExportTxt.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportTxt.FlatAppearance.BorderSize = 0
$btnExportTxt.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportTxt.Size = New-Object System.Drawing.Size(150, 35)
$btnExportTxt.Cursor = [System.Windows.Forms.Cursors]::Hand

# CSV gumb
$btnExportCSV.Text = "Spremi kao CSV"
$btnExportCSV.BackColor = $Colors.Success
$btnExportCSV.ForeColor = $Colors.White
$btnExportCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportCSV.FlatAppearance.BorderSize = 0
$btnExportCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportCSV.Size = New-Object System.Drawing.Size(150, 35)
$btnExportCSV.Cursor = [System.Windows.Forms.Cursors]::Hand

# Word gumb
$btnExportDoc.Text = "Spremi kao DOC"
$btnExportDoc.BackColor = $Colors.Primary
$btnExportDoc.ForeColor = $Colors.White
$btnExportDoc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportDoc.FlatAppearance.BorderSize = 0
$btnExportDoc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportDoc.Size = New-Object System.Drawing.Size(150, 35)
$btnExportDoc.Cursor = [System.Windows.Forms.Cursors]::Hand

# Excel gumb
$btnExportExcel.Text = "Spremi kao Excel"
$btnExportExcel.BackColor = $Colors.Success
$btnExportExcel.ForeColor = $Colors.White
$btnExportExcel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportExcel.FlatAppearance.BorderSize = 0
$btnExportExcel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportExcel.Size = New-Object System.Drawing.Size(150, 35)
$btnExportExcel.Cursor = [System.Windows.Forms.Cursors]::Hand

# PDF gumb
$btnExportPDF.Text = "Spremi kao PDF"
$btnExportPDF.BackColor = $Colors.Accent
$btnExportPDF.ForeColor = $Colors.White
$btnExportPDF.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportPDF.FlatAppearance.BorderSize = 0
$btnExportPDF.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportPDF.Size = New-Object System.Drawing.Size(150, 35)
$btnExportPDF.Cursor = [System.Windows.Forms.Cursors]::Hand

# HTML gumb
$btnExportHTML.Text = "Spremi kao HTML"
$btnExportHTML.BackColor = $Colors.Warning
$btnExportHTML.ForeColor = $Colors.White
$btnExportHTML.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExportHTML.FlatAppearance.BorderSize = 0
$btnExportHTML.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnExportHTML.Size = New-Object System.Drawing.Size(150, 35)
$btnExportHTML.Cursor = [System.Windows.Forms.Cursors]::Hand

# Print gumb
$btnPrint.Text = "Print"
$btnPrint.BackColor = $Colors.Dark
$btnPrint.ForeColor = $Colors.White
$btnPrint.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnPrint.FlatAppearance.BorderSize = 0
$btnPrint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnPrint.Size = New-Object System.Drawing.Size(120, 35)
$btnPrint.Cursor = [System.Windows.Forms.Cursors]::Hand

# Info gumb
$btnInfo.Text = "Info"
$btnInfo.BackColor = $Colors.Purple
$btnInfo.ForeColor = $Colors.White
$btnInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnInfo.FlatAppearance.BorderSize = 0
$btnInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnInfo.Size = New-Object System.Drawing.Size(120, 35)
$btnInfo.Cursor = [System.Windows.Forms.Cursors]::Hand

# === DODAJ HOVER EFEKTE ===
function Add-HoverEffect($Button, $BaseColor) {
    # Izraèunaj hover boje unaprijed
    $HoverColor = [System.Drawing.Color]::FromArgb(
        [Math]::Min(255, $BaseColor.R + 30),
        [Math]::Min(255, $BaseColor.G + 30),
        [Math]::Min(255, $BaseColor.B + 30)
    )
    
    $ClickColor = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BaseColor.R - 30),
        [Math]::Max(0, $BaseColor.G - 30),
        [Math]::Max(0, $BaseColor.B - 30)
    )
    
    # Mouse Enter - svjetlija boja
    $Button.Add_MouseEnter({
        $this.BackColor = $HoverColor
    })
    
    # Mouse Leave - vrati originalnu boju
    $Button.Add_MouseLeave({
        $this.BackColor = $BaseColor
    })
    
    # Mouse Down - tamnija boja
    $Button.Add_MouseDown({
        $this.BackColor = $ClickColor
    })
    
    # Mouse Up - vrati hover boju
    $Button.Add_MouseUp({
        $this.BackColor = $HoverColor
    })
}

# Primijeni hover efekte
Add-HoverEffect -Button $btnExportTxt -BaseColor $Colors.Primary
Add-HoverEffect -Button $btnExportCSV -BaseColor $Colors.Success
Add-HoverEffect -Button $btnExportDoc -BaseColor $Colors.Primary
Add-HoverEffect -Button $btnExportExcel -BaseColor $Colors.Success
Add-HoverEffect -Button $btnExportPDF -BaseColor $Colors.Accent
Add-HoverEffect -Button $btnExportHTML -BaseColor $Colors.Warning
Add-HoverEffect -Button $btnPrint -BaseColor $Colors.Dark
Add-HoverEffect -Button $btnInfo -BaseColor $Colors.Purple

# === REPOZICIONIRAJ GUMBOVE ===
$x = 10
$y = 12
$buttons = @($btnExportTxt, $btnExportCSV, $btnExportDoc, $btnExportExcel, $btnExportPDF, $btnExportHTML, $btnPrint, $btnInfo)

foreach ($btn in $buttons) {
    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $x += $btn.Width + 8
}

# === DODAJ STATUS BAR ===
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $Colors.Dark
$statusBar.ForeColor = $Colors.White

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Spremno | $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss') | Naziv raèunala: $env:COMPUTERNAME | Prijavljeni ste kao: $env:USERNAME"
$statusLabel.ForeColor = $Colors.Light
$statusBar.Items.Add($statusLabel)

# Dodaj status bar prije ShowDialog
$form.Controls.Add($statusBar)

# === DODAJ GRADIJENT NA BUTTON PANEL ===
$panelButtons.Add_Paint({
    $graphics = $_.Graphics
    $rect = New-Object System.Drawing.Rectangle(0, 0, $panelButtons.Width, $panelButtons.Height)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect, 
        $Colors.Secondary, 
        $Colors.Dark, 
        [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
    )
    $graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
})

# === DODAJ FADE-IN ANIMACIJU ===
$form.Add_Shown({
    $form.Opacity = 1.0
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.TopMost = $false  # Uklonite TopMost jer može stvarati probleme
    $form.ShowInTaskbar = $true
    Write-Host "Transparentnost ispravljena - Opacity postavljena na: $($form.Opacity)" -ForegroundColor Cyan
})

# DODAJ EVENT HANDLER ZA ZATVARANJE POWERSHELL-A
$form.Add_FormClosed({
    Write-Host "Zatvaranje PowerShell procesa..." -ForegroundColor Yellow
    [Environment]::Exit(0)
})

# ZATVORITI LOADING SCREEN PRIJE POKRETANJA GUI
Start-Sleep -Seconds 1
$timer.Stop()
$timer.Dispose()
$loadingForm.Close()
$loadingForm.Dispose()
Write-Host "Loading screen zatvoren!" -ForegroundColor Green
Write-Host "Vaša glavna aplikacija æe se sada pokrenuti..." -ForegroundColor Green

# SADA POKRENI GLAVNI GUI
$form.Topmost = $true
[void]$form.ShowDialog()