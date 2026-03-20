Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction Stop
}

$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:ConfigPath = Join-Path $Script:AppRoot "XGD.config.json"
$Script:Config = $null
$Script:MainForm = $null
$Script:TxtSalida = $null
$Script:TxtResultados = $null
$Script:TreeResultados = $null
$Script:LblContadorEquipos = $null
$Script:StatusLabel = $null
$Script:ComboGrupoAD = $null
$Script:ListBoxOUs = $null
$Script:SelectedWorkOUs = [System.Collections.Generic.List[string]]::new()
$Script:BtnFiltroEspecial = $null
$Script:LastComputerResults = @()
$Script:LastTextResult = ""
$Script:ResultSearchIndex = -1
$Script:ResultSearchNodes = @()
$Script:ResultSearchPositions = @()

function Convert-ToStringArray {
    param([object]$Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        return @(
            $Value |
            ForEach-Object { [string]$_ } |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return @()
    }

    return @(
        ($text -split "(`r`n|`n|;)") |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Join-Lines {
    param([object]$Value)
    return (Convert-ToStringArray $Value) -join [Environment]::NewLine
}

function Escape-ADFilterValue {
    param([string]$Value)
    if ($null -eq $Value) {
        return ""
    }
    return $Value.Replace("'", "''")
}

function Get-ADCommonParameters {
    $parameters = @{}
    if ($Script:Config -and -not [string]::IsNullOrWhiteSpace($Script:Config.Server)) {
        $parameters.Server = $Script:Config.Server
    }
    return $parameters
}

function Get-DefaultNamingContext {
    try {
        $params = Get-ADCommonParameters
        return (Get-ADRootDSE @params -ErrorAction Stop).defaultNamingContext
    }
    catch {
        return ""
    }
}

function Resolve-ConfiguredBases {
    param([object]$Bases)

    $resolved = Convert-ToStringArray $Bases
    if ($resolved.Count -gt 0) {
        return @($resolved | Select-Object -Unique)
    }

    $defaultNamingContext = Get-DefaultNamingContext
    if ([string]::IsNullOrWhiteSpace($defaultNamingContext)) {
        return @()
    }

    return @($defaultNamingContext)
}

function Get-DefaultConfig {
    $defaultNamingContext = Get-DefaultNamingContext
    $defaultBases = if ($defaultNamingContext) { @($defaultNamingContext) } else { @() }

    return [ordered]@{
        UiTitle                            = "XGD - Utilidad general AD"
        Server                             = ""
        GroupSearchBases                   = $defaultBases
        GroupNameFilter                    = "*"
        ExploreSearchBases                 = $defaultBases
        BrowseRoots                        = $defaultBases
        ComputerContainerName              = "Equipos"
        HiddenOUSegments                   = @("Equipos")
        ExcludedOUPatterns                 = @("_Cuentas deshabilitadas", "Transito", "Pre-Windows 10")
        ExcludedComputerNameRegex          = ""
        FilteredComputerLabel              = "Anadir Equipos Filtrados"
        FilteredComputerIncludeRegex       = ""
        ApplyDelegationOnCreate            = $false
        DelegationGroupDN                  = ""
        CreatedComputerDescriptionTemplate = "Equipo creado {date}"
        FixedTargetGroups                  = @()
        SavedWorkOUs                       = @()
    }
}

function Normalize-Config {
    param([object]$ConfigObject)

    $defaults = Get-DefaultConfig
    $normalized = [ordered]@{}

    foreach ($key in $defaults.Keys) {
        if ($null -eq $ConfigObject -or -not ($ConfigObject.PSObject.Properties.Name -contains $key)) {
            $normalized[$key] = $defaults[$key]
            continue
        }

        switch ($key) {
            { $_ -in @("GroupSearchBases", "ExploreSearchBases", "BrowseRoots", "HiddenOUSegments", "ExcludedOUPatterns", "FixedTargetGroups", "SavedWorkOUs") } {
                $normalized[$key] = Convert-ToStringArray $ConfigObject.$key
            }
            { $_ -in @("ApplyDelegationOnCreate") } {
                $normalized[$key] = [bool]$ConfigObject.$key
            }
            default {
                $normalized[$key] = [string]$ConfigObject.$key
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($normalized.UiTitle)) {
        $normalized.UiTitle = $defaults.UiTitle
    }

    return [PSCustomObject]$normalized
}

function Save-Config {
    param([object]$ConfigObject)

    $normalized = Normalize-Config $ConfigObject
    $json = $normalized | ConvertTo-Json -Depth 6
    Set-Content -Path $Script:ConfigPath -Value $json -Encoding UTF8
    return $normalized
}

function Load-Config {
    if (-not (Test-Path $Script:ConfigPath)) {
        $Script:Config = Save-Config (Get-DefaultConfig)
        return
    }

    try {
        $rawConfig = Get-Content -Path $Script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $Script:Config = Normalize-Config $rawConfig
        Save-Config $Script:Config | Out-Null
    }
    catch {
        $Script:Config = Save-Config (Get-DefaultConfig)
    }
}

function Set-Status {
    param([string]$Text)

    if ($Script:StatusLabel) {
        $Script:StatusLabel.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Mostrar-Mensaje {
    param([string]$Mensaje)

    $timestamp = Get-Date -Format "HH:mm:ss"
    if ($Script:TxtSalida) {
        $Script:TxtSalida.AppendText("$timestamp - $Mensaje`r`n")
        $Script:TxtSalida.SelectionStart = $Script:TxtSalida.TextLength
        $Script:TxtSalida.ScrollToCaret()
    }
    else {
        Write-Host "$timestamp - $Mensaje"
    }

    Set-Status $Mensaje
}

function Show-ErrorDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Clear-Results {
    $Script:TxtResultados.Clear()
    $Script:TxtResultados.Visible = $true
    $Script:TreeResultados.Nodes.Clear()
    $Script:TreeResultados.Visible = $false
    $Script:LblContadorEquipos.Visible = $false
    $Script:LastComputerResults = @()
    $Script:LastTextResult = ""
    $Script:ResultSearchIndex = -1
    $Script:ResultSearchNodes = @()
    $Script:ResultSearchPositions = @()
}

function Split-DistinguishedName {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
        return @()
    }

    return [regex]::Split($DistinguishedName, "(?<!\\),")
}

function Get-DnFriendlyName {
    param([string]$DistinguishedName)

    $parts = Split-DistinguishedName $DistinguishedName
    if ($parts.Count -eq 0) {
        return $DistinguishedName
    }

    if ($parts[0] -like "DC=*") {
        $domainParts = @(
            $parts |
            Where-Object { $_ -like "DC=*" } |
            ForEach-Object { $_.Substring(3) }
        )
        if ($domainParts.Count -gt 0) {
            return ($domainParts -join ".")
        }
    }

    $first = $parts[0]
    $separatorIndex = $first.IndexOf("=")
    if ($separatorIndex -gt 0) {
        return $first.Substring($separatorIndex + 1)
    }

    return $DistinguishedName
}

function Get-OUPathSegments {
    param([string]$DistinguishedName)

    $ouSegments = @(
        (Split-DistinguishedName $DistinguishedName) |
        Where-Object { $_ -like "OU=*" } |
        ForEach-Object { $_.Substring(3) }
    )

    if ($ouSegments.Count -eq 0) {
        return @()
    }

    [array]::Reverse($ouSegments)

    $hiddenSegments = Convert-ToStringArray $Script:Config.HiddenOUSegments
    if ($hiddenSegments.Count -gt 0) {
        $ouSegments = @($ouSegments | Where-Object { $_ -notin $hiddenSegments })
    }

    return $ouSegments
}

function Get-DisplayPathSegments {
    param([string]$DistinguishedName)

    $segments = @(Get-OUPathSegments $DistinguishedName)
    if ($segments.Count -eq 0) {
        return @("Sin clasificar")
    }

    return $segments
}

function Get-DisplayPathText {
    param([string]$DistinguishedName)
    return ((Get-DisplayPathSegments $DistinguishedName) -join " / ")
}

function Test-ExcludedOU {
    param([string]$DistinguishedName)

    foreach ($pattern in (Convert-ToStringArray $Script:Config.ExcludedOUPatterns)) {
        if (-not [string]::IsNullOrWhiteSpace($pattern) -and $DistinguishedName -match $pattern) {
            return $true
        }
    }

    return $false
}

function Test-ComputerNameAllowed {
    param(
        [string]$Name,
        [string]$IncludeRegex
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $false
    }

    if ($Name -match "^CNF:") {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($Script:Config.ExcludedComputerNameRegex) -and $Name -match $Script:Config.ExcludedComputerNameRegex) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($IncludeRegex) -and $Name -notmatch $IncludeRegex) {
        return $false
    }

    return $true
}

function Get-ChildOUs {
    param([string]$SearchBase)

    $params = Get-ADCommonParameters
    try {
        return @(
            Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -SearchScope OneLevel @params -ErrorAction Stop |
            Where-Object { -not (Test-ExcludedOU $_.DistinguishedName) } |
            Sort-Object Name
        )
    }
    catch {
        return @()
    }
}

function Get-ImmediateComputersForTree {
    param(
        [string]$SearchBase,
        [string]$IncludeRegex
    )

    $containerName = $Script:Config.ComputerContainerName
    $shouldLoadDirectComputers = [string]::IsNullOrWhiteSpace($containerName) -or ($SearchBase -imatch "^OU=$([regex]::Escape($containerName))(,|$)")
    if (-not $shouldLoadDirectComputers) {
        return @()
    }

    $params = Get-ADCommonParameters
    try {
        return @(
            Get-ADComputer -Filter * -SearchBase $SearchBase -SearchScope OneLevel @params -ErrorAction Stop |
            Where-Object { Test-ComputerNameAllowed -Name $_.Name -IncludeRegex $IncludeRegex } |
            Sort-Object Name
        )
    }
    catch {
        return @()
    }
}

function Get-ComputersFromOU {
    param(
        [string]$BaseOU,
        [string]$IncludeRegex
    )

    $params = Get-ADCommonParameters
    $computerContainerName = $Script:Config.ComputerContainerName
    $foundComputers = @()

    try {
        if ([string]::IsNullOrWhiteSpace($computerContainerName)) {
            $foundComputers = @(
                Get-ADComputer -Filter * -SearchBase $BaseOU -SearchScope Subtree @params -ErrorAction Stop
            )
        }
        else {
            $isContainer = $BaseOU -imatch "^OU=$([regex]::Escape($computerContainerName))(,|$)"
            if ($isContainer) {
                $containerOUs = @([PSCustomObject]@{ DistinguishedName = $BaseOU })
            }
            else {
                $filterText = Escape-ADFilterValue $computerContainerName
                $containerOUs = @(
                    Get-ADOrganizationalUnit -Filter "Name -eq '$filterText'" -SearchBase $BaseOU -SearchScope Subtree @params -ErrorAction Stop |
                    Where-Object { -not (Test-ExcludedOU $_.DistinguishedName) }
                )
            }

            foreach ($ou in $containerOUs) {
                try {
                    $foundComputers += Get-ADComputer -Filter * -SearchBase $ou.DistinguishedName -SearchScope OneLevel @params -ErrorAction Stop
                }
                catch {
                }
            }
        }
    }
    catch {
        Mostrar-Mensaje "[!] Error al leer equipos en '$BaseOU': $($_.Exception.Message)"
    }

    return @(
        $foundComputers |
        Where-Object { Test-ComputerNameAllowed -Name $_.Name -IncludeRegex $IncludeRegex } |
        Sort-Object DistinguishedName -Unique
    )
}

function Get-ComputersFromSearchBases {
    param(
        [object]$SearchBases,
        [string]$IncludeRegex
    )

    $allComputers = @()
    foreach ($searchBase in (Resolve-ConfiguredBases $SearchBases)) {
        if (Test-ExcludedOU $searchBase) {
            continue
        }

        Mostrar-Mensaje "[...] Explorando equipos en: $searchBase"
        $allComputers += Get-ComputersFromOU -BaseOU $searchBase -IncludeRegex $IncludeRegex
    }

    return @($allComputers | Sort-Object DistinguishedName -Unique)
}

function Get-ADGroupDirectMembersRanged {
    param([string]$Identity)

    $params = Get-ADCommonParameters
    $directMembers = @()

    $group = Get-ADGroup -Identity $Identity -Properties DistinguishedName @params -ErrorAction Stop

    $start = 0
    $step = 1500
    while ($true) {
        $end = $start + $step - 1
        $attributeName = "member;range=$start-$end"

        try {
            $groupWithRange = Get-ADGroup -Identity $group.DistinguishedName -Properties $attributeName @params -ErrorAction Stop
        }
        catch {
            break
        }

        $rangeProperty = $groupWithRange.PSObject.Properties | Where-Object { $_.Name -like "member;range=*" } | Select-Object -First 1
        if (-not $rangeProperty) {
            break
        }

        foreach ($memberDN in @($rangeProperty.Value)) {
            if ([string]::IsNullOrWhiteSpace($memberDN)) {
                continue
            }

            try {
                $directMembers += Get-ADObject -Identity $memberDN -Properties objectClass, distinguishedName, name, samAccountName @params -ErrorAction Stop
            }
            catch {
            }
        }

        if ($rangeProperty.Name -match "\*$") {
            break
        }

        $start = $end + 1
    }

    if ($directMembers.Count -gt 0) {
        return @($directMembers | Sort-Object DistinguishedName -Unique)
    }

    try {
        return @(
            Get-ADGroupMember -Identity $group.DistinguishedName @params -ErrorAction Stop |
            Sort-Object DistinguishedName -Unique
        )
    }
    catch {
        return @()
    }
}

function Get-ADGroupMemberSafe {
    param(
        [string]$Identity,
        [switch]$Recursive
    )

    $allMembers = New-Object System.Collections.Generic.List[object]
    $visitedGroups = New-Object "System.Collections.Generic.HashSet[string]"

    function Visit-GroupMembers {
        param([string]$GroupIdentity)

        $params = Get-ADCommonParameters
        $groupObject = Get-ADGroup -Identity $GroupIdentity -Properties DistinguishedName @params -ErrorAction Stop
        $groupKey = $groupObject.DistinguishedName.ToLowerInvariant()

        if (-not $visitedGroups.Add($groupKey)) {
            return
        }

        $directMembers = @(Get-ADGroupDirectMembersRanged -Identity $groupObject.DistinguishedName)
        foreach ($member in $directMembers) {
            [void]$allMembers.Add($member)
            if ($Recursive -and $member.objectClass -eq "group") {
                Visit-GroupMembers -GroupIdentity $member.DistinguishedName
            }
        }
    }

    Visit-GroupMembers -GroupIdentity $Identity
    return @($allMembers | Sort-Object DistinguishedName -Unique)
}

function Resolve-ComputerRecord {
    param([string]$Identity)

    $value = $Identity.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Debe indicar un nombre de equipo."
    }

    $params = Get-ADCommonParameters

    try {
        return Get-ADComputer -Identity $value -Properties Name, DistinguishedName, SamAccountName @params -ErrorAction Stop
    }
    catch {
    }

    $safeValue = Escape-ADFilterValue $value
    $found = @()
    foreach ($base in (Resolve-ConfiguredBases @((Convert-ToStringArray $Script:Config.BrowseRoots) + (Convert-ToStringArray $Script:Config.ExploreSearchBases)))) {
        try {
            $found += Get-ADComputer -Filter "Name -eq '$safeValue'" -SearchBase $base -SearchScope Subtree @params -ErrorAction SilentlyContinue |
                Select-Object Name, DistinguishedName, SamAccountName
        }
        catch {
        }
    }

    $unique = @($found | Sort-Object DistinguishedName -Unique)
    if ($unique.Count -eq 1) {
        return $unique[0]
    }

    if ($unique.Count -gt 1) {
        throw "Se encontraron varios equipos llamados '$value'. Seleccione uno por OU o use un nombre mas especifico."
    }

    throw "No se encontro el equipo '$value'."
}

function Resolve-GroupRecord {
    param([string]$Identity)

    $value = $Identity.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Debe indicar un grupo."
    }

    $params = Get-ADCommonParameters

    try {
        return Get-ADGroup -Identity $value -Properties Name, DistinguishedName, SamAccountName @params -ErrorAction Stop
    }
    catch {
    }

    $safeValue = Escape-ADFilterValue $value
    $found = @()
    $searchBases = @($Script:SelectedWorkOUs) + @(Resolve-ConfiguredBases $Script:Config.BrowseRoots)
    foreach ($base in ($searchBases | Select-Object -Unique)) {
        try {
            $found += Get-ADGroup -Filter "Name -eq '$safeValue' -or SamAccountName -eq '$safeValue'" -SearchBase $base -SearchScope Subtree @params -ErrorAction SilentlyContinue |
                Select-Object Name, DistinguishedName, SamAccountName
        }
        catch {
        }
    }

    $unique = @($found | Sort-Object DistinguishedName -Unique)
    if ($unique.Count -eq 1) {
        return $unique[0]
    }

    if ($unique.Count -gt 1) {
        throw "Se encontraron varios grupos llamados '$value'. Seleccione el grupo desde la lista cargada."
    }

    throw "No se encontro el grupo '$value'."
}

function Get-SelectedGroupRecord {
    if ($Script:ComboGrupoAD.SelectedItem -and $Script:ComboGrupoAD.SelectedItem.PSObject.Properties.Name -contains "DistinguishedName") {
        return $Script:ComboGrupoAD.SelectedItem
    }

    $groupText = $Script:ComboGrupoAD.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($groupText)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Debe seleccionar o escribir un grupo de AD.",
            "Grupo no seleccionado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $null
    }

    try {
        return Resolve-GroupRecord $groupText
    }
    catch {
        Show-ErrorDialog -Title "Grupo no valido" -Message $_.Exception.Message
        return $null
    }
}

function Get-ComputerSummaryText {
    param(
        [string]$Title,
        [object[]]$Computers
    )

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("============================================================")
    [void]$lines.Add($Title)
    [void]$lines.Add("============================================================")
    [void]$lines.Add("Total de equipos: $($Computers.Count)")
    [void]$lines.Add("")

    foreach ($computer in ($Computers | Sort-Object Path, Name)) {
        [void]$lines.Add(("{0} | {1}" -f $computer.Name, $computer.Path))
    }

    [void]$lines.Add("")
    [void]$lines.Add("Generado: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')")
    return ($lines -join "`r`n")
}

function Show-ResultsText {
    param(
        [string]$Title,
        [string[]]$Lines
    )

    $Script:TreeResultados.Visible = $false
    $Script:TxtResultados.Visible = $true
    $Script:LblContadorEquipos.Visible = $false
    $Script:TxtResultados.Clear()
    $Script:TxtResultados.AppendText("============================================================`r`n")
    $Script:TxtResultados.AppendText("$Title`r`n")
    $Script:TxtResultados.AppendText("============================================================`r`n`r`n")

    foreach ($line in $Lines) {
        $Script:TxtResultados.AppendText("$line`r`n")
    }

    $Script:LastComputerResults = @()
    $Script:LastTextResult = $Script:TxtResultados.Text
}

function Find-OrCreateTreeChild {
    param(
        [System.Windows.Forms.TreeNodeCollection]$Nodes,
        [string]$Text,
        [string]$NameKey
    )

    foreach ($node in $Nodes) {
        if ($node.Name -eq $NameKey) {
            return $node
        }
    }

    $newNode = New-Object System.Windows.Forms.TreeNode
    $newNode.Text = $Text
    $newNode.Name = $NameKey
    [void]$Nodes.Add($newNode)
    return $newNode
}

function Show-ResultsComputerTree {
    param(
        [string]$Title,
        [object[]]$Computers
    )

    $prepared = @(
        $Computers |
        ForEach-Object {
            [PSCustomObject]@{
                Name              = $_.Name
                DistinguishedName = $_.DistinguishedName
                Path              = Get-DisplayPathText $_.DistinguishedName
            }
        }
    )

    $Script:TxtResultados.Visible = $false
    $Script:TreeResultados.Visible = $true
    $Script:TreeResultados.Nodes.Clear()
    $Script:LblContadorEquipos.Text = "$Title - $($prepared.Count) equipos"
    $Script:LblContadorEquipos.Visible = $true

    $Script:TreeResultados.BeginUpdate()
    foreach ($computer in ($prepared | Sort-Object Path, Name)) {
        $segments = @(Get-DisplayPathSegments $computer.DistinguishedName)
        $pathKey = ""
        $currentNodes = $Script:TreeResultados.Nodes

        foreach ($segment in $segments) {
            $pathKey = if ($pathKey) { "$pathKey|$segment" } else { $segment }
            $parentNode = Find-OrCreateTreeChild -Nodes $currentNodes -Text $segment -NameKey $pathKey
            $parentNode.ForeColor = [System.Drawing.Color]::DarkBlue
            $currentNodes = $parentNode.Nodes
        }

        $computerNode = New-Object System.Windows.Forms.TreeNode
        $computerNode.Text = $computer.Name
        $computerNode.Name = "$pathKey|$($computer.Name)"
        $computerNode.ToolTipText = $computer.DistinguishedName
        [void]$currentNodes.Add($computerNode)
    }
    $Script:TreeResultados.EndUpdate()

    $Script:LastComputerResults = $prepared
    $Script:LastTextResult = Get-ComputerSummaryText -Title $Title -Computers $prepared
}

function Focus-TreeNode {
    param(
        [System.Windows.Forms.TreeView]$TreeView,
        [System.Windows.Forms.TreeNode]$Node
    )

    $current = $Node
    while ($current.Parent -ne $null) {
        $current.Parent.Expand()
        $current = $current.Parent
    }

    $TreeView.SelectedNode = $Node
    $Node.EnsureVisible()
    $TreeView.Focus()
}

function Search-LoadedTreeNodes {
    param(
        [System.Windows.Forms.TreeNode]$Node,
        [string]$Text
    )

    $found = @()
    if ($Node.Text -match [regex]::Escape($Text)) {
        $found += $Node
    }

    foreach ($child in $Node.Nodes) {
        $found += Search-LoadedTreeNodes -Node $child -Text $Text
    }

    return $found
}

function Refresh-OUListBox {
    $Script:ListBoxOUs.Items.Clear()
    foreach ($ou in $Script:SelectedWorkOUs) {
        [void]$Script:ListBoxOUs.Items.Add((Get-DnFriendlyName $ou))
    }
}

function Add-WorkOU {
    $selected = Show-OUSelectionDialog -Title "Seleccionar OU de trabajo"
    if (-not $selected) { return }

    $lowerSelected = $selected.ToLowerInvariant()
    foreach ($existing in $Script:SelectedWorkOUs) {
        if ($existing.ToLowerInvariant() -eq $lowerSelected) {
            Mostrar-Mensaje "[!] La OU ya esta en la lista: $(Get-DnFriendlyName $selected)"
            return
        }
    }

    $Script:SelectedWorkOUs.Add($selected)
    Refresh-OUListBox
    Mostrar-Mensaje "[OK] OU anadida: $(Get-DnFriendlyName $selected)"
}

function Remove-WorkOU {
    $index = $Script:ListBoxOUs.SelectedIndex
    if ($index -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Seleccione una OU de la lista para quitarla.",
            "Sin seleccion",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $removed = $Script:SelectedWorkOUs[$index]
    $Script:SelectedWorkOUs.RemoveAt($index)
    Refresh-OUListBox
    Mostrar-Mensaje "[OK] OU quitada: $(Get-DnFriendlyName $removed)"
}

function Load-GroupList {
    if ($Script:SelectedWorkOUs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Debe seleccionar al menos una OU de trabajo antes de cargar grupos.",
            "Sin OUs seleccionadas",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $pattern = $Script:Config.GroupNameFilter
    if ([string]::IsNullOrWhiteSpace($pattern)) {
        $pattern = "*"
    }
    if ($pattern -notmatch "[\*\?]") {
        $pattern = "*$pattern*"
    }

    $safePattern = Escape-ADFilterValue $pattern
    $params = Get-ADCommonParameters
    $groupMap = @{}

    $Script:ComboGrupoAD.Items.Clear()
    $Script:ComboGrupoAD.Text = ""

    Mostrar-Mensaje "[...] Cargando grupos desde $($Script:SelectedWorkOUs.Count) OU(s)"

    foreach ($searchBase in $Script:SelectedWorkOUs) {
        try {
            $groups = Get-ADGroup -Filter "Name -like '$safePattern'" -SearchBase $searchBase -SearchScope Subtree @params -ErrorAction Stop |
                Select-Object Name, DistinguishedName, SamAccountName

            foreach ($group in $groups) {
                $groupMap[$group.DistinguishedName.ToLowerInvariant()] = $group
            }
        }
        catch {
            Mostrar-Mensaje "[!] Error al cargar grupos en '$searchBase': $($_.Exception.Message)"
        }
    }

    $uniqueGroups = @($groupMap.Values | Sort-Object Name, DistinguishedName)
    $duplicateNames = @(
        $uniqueGroups |
        Group-Object Name |
        Where-Object { $_.Count -gt 1 } |
        ForEach-Object { $_.Name }
    )

    $autocomplete = New-Object System.Windows.Forms.AutoCompleteStringCollection

    foreach ($group in $uniqueGroups) {
        $displayName = if ($duplicateNames -contains $group.Name) {
            "$($group.Name) | $($group.DistinguishedName)"
        }
        else {
            $group.Name
        }

        $item = [PSCustomObject]@{
            DisplayName       = $displayName
            Name              = $group.Name
            DistinguishedName = $group.DistinguishedName
            SamAccountName    = $group.SamAccountName
        }

        [void]$Script:ComboGrupoAD.Items.Add($item)
        [void]$autocomplete.Add($displayName)
    }

    $Script:ComboGrupoAD.AutoCompleteCustomSource = $autocomplete
    Mostrar-Mensaje "[OK] Cargados $($uniqueGroups.Count) grupos"
}

function Add-PlaceholderNode {
    param([System.Windows.Forms.TreeNode]$Node)

    $placeholder = New-Object System.Windows.Forms.TreeNode
    $placeholder.Text = "..."
    [void]$Node.Nodes.Add($placeholder)
}

function New-OUNode {
    param([string]$DistinguishedName)

    $node = New-Object System.Windows.Forms.TreeNode
    $node.Text = Get-DnFriendlyName $DistinguishedName
    $node.ToolTipText = $DistinguishedName
    $node.Tag = [PSCustomObject]@{
        Type              = "ou"
        DistinguishedName = $DistinguishedName
    }
    Add-PlaceholderNode -Node $node
    return $node
}

function Load-OUNodeChildren {
    param(
        [System.Windows.Forms.TreeNode]$Node,
        [string]$IncludeRegex
    )

    $tag = $Node.Tag
    if (-not $tag -or $tag.Type -ne "ou") {
        return
    }

    $Node.Nodes.Clear()

    foreach ($ou in (Get-ChildOUs $tag.DistinguishedName)) {
        [void]$Node.Nodes.Add((New-OUNode -DistinguishedName $ou.DistinguishedName))
    }

    foreach ($computer in (Get-ImmediateComputersForTree -SearchBase $tag.DistinguishedName -IncludeRegex $IncludeRegex)) {
        $computerNode = New-Object System.Windows.Forms.TreeNode
        $computerNode.Text = $computer.Name
        $computerNode.ToolTipText = $computer.DistinguishedName
        $computerNode.Tag = [PSCustomObject]@{
            Type              = "computer"
            DistinguishedName = $computer.DistinguishedName
            Name              = $computer.Name
        }
        [void]$Node.Nodes.Add($computerNode)
    }

    if ($Node.Nodes.Count -eq 0) {
        $emptyNode = New-Object System.Windows.Forms.TreeNode
        $emptyNode.Text = "(Vacio)"
        $emptyNode.ForeColor = [System.Drawing.Color]::Gray
        [void]$Node.Nodes.Add($emptyNode)
    }
}

function New-OUTree {
    param(
        [object]$Roots,
        [string]$IncludeRegex,
        [bool]$UseCheckboxes
    )

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock = "Fill"
    $tree.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tree.CheckBoxes = $UseCheckboxes
    $tree.HideSelection = $false
    $tree.ShowNodeToolTips = $true

    foreach ($root in (Resolve-ConfiguredBases $Roots)) {
        [void]$tree.Nodes.Add((New-OUNode -DistinguishedName $root))
    }

    $tree.Add_BeforeExpand({
        param($sender, $eventArgs)

        if ($eventArgs.Node.Nodes.Count -eq 1 -and $eventArgs.Node.Nodes[0].Text -eq "...") {
            Load-OUNodeChildren -Node $eventArgs.Node -IncludeRegex $IncludeRegex
        }
    })

    return $tree
}

function Resolve-ComputerEntriesFromCsv {
    param(
        [string]$Path,
        [string]$IncludeRegex
    )

    $csvRows = Import-Csv -Path $Path -ErrorAction Stop
    if (-not $csvRows -or $csvRows.Count -eq 0) {
        return @()
    }

    $propertyNames = @($csvRows[0].PSObject.Properties.Name)
    $preferredColumns = @("Name", "NombreEquipo", "ComputerName", "Equipo")
    $selectedColumn = $preferredColumns | Where-Object { $_ -in $propertyNames } | Select-Object -First 1
    if (-not $selectedColumn) {
        $selectedColumn = $propertyNames[0]
    }

    $entries = @()
    foreach ($row in $csvRows) {
        $name = [string]$row.$selectedColumn
        if (-not (Test-ComputerNameAllowed -Name $name -IncludeRegex $IncludeRegex)) {
            continue
        }

        try {
            $computer = Resolve-ComputerRecord $name
            $entries += [PSCustomObject]@{
                Name              = $computer.Name
                DistinguishedName = $computer.DistinguishedName
                Path              = Get-DisplayPathText $computer.DistinguishedName
            }
        }
        catch {
            $entries += [PSCustomObject]@{
                Name              = $name
                DistinguishedName = $null
                Path              = "No resuelto"
            }
        }
    }

    return @($entries | Sort-Object Name -Unique)
}

function Show-ComputerSelectionDialog {
    param(
        [string]$Title,
        [string]$IncludeRegex
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(1120, 700)
    $dialog.StartPosition = "CenterParent"
    $dialog.MinimumSize = New-Object System.Drawing.Size(980, 620)
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Location = New-Object System.Drawing.Point(10, 10)
    $leftPanel.Size = New-Object System.Drawing.Size(470, 600)
    $leftPanel.BorderStyle = "FixedSingle"
    $dialog.Controls.Add($leftPanel)

    $leftLabel = New-Object System.Windows.Forms.Label
    $leftLabel.Text = "Seleccion de OUs y equipos"
    $leftLabel.Location = New-Object System.Drawing.Point(8, 8)
    $leftLabel.AutoSize = $true
    $leftLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $leftPanel.Controls.Add($leftLabel)

    $treePanel = New-Object System.Windows.Forms.Panel
    $treePanel.Location = New-Object System.Drawing.Point(8, 34)
    $treePanel.Size = New-Object System.Drawing.Size(452, 548)
    $leftPanel.Controls.Add($treePanel)

    $tree = New-OUTree -Roots $Script:Config.BrowseRoots -IncludeRegex $IncludeRegex -UseCheckboxes $true
    $treePanel.Controls.Add($tree)

    $middlePanel = New-Object System.Windows.Forms.Panel
    $middlePanel.Location = New-Object System.Drawing.Point(490, 10)
    $middlePanel.Size = New-Object System.Drawing.Size(110, 600)
    $dialog.Controls.Add($middlePanel)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = ">>"
    $btnAdd.Location = New-Object System.Drawing.Point(10, 200)
    $btnAdd.Size = New-Object System.Drawing.Size(90, 50)
    $btnAdd.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnAdd.ForeColor = [System.Drawing.Color]::White
    $btnAdd.FlatStyle = "Flat"
    $middlePanel.Controls.Add($btnAdd)

    $btnImportCsv = New-Object System.Windows.Forms.Button
    $btnImportCsv.Text = "CSV"
    $btnImportCsv.Location = New-Object System.Drawing.Point(10, 260)
    $btnImportCsv.Size = New-Object System.Drawing.Size(90, 45)
    $btnImportCsv.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnImportCsv.ForeColor = [System.Drawing.Color]::White
    $btnImportCsv.FlatStyle = "Flat"
    $middlePanel.Controls.Add($btnImportCsv)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "<<"
    $btnRemove.Location = New-Object System.Drawing.Point(10, 315)
    $btnRemove.Size = New-Object System.Drawing.Size(90, 50)
    $btnRemove.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnRemove.ForeColor = [System.Drawing.Color]::White
    $btnRemove.FlatStyle = "Flat"
    $middlePanel.Controls.Add($btnRemove)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Limpiar"
    $btnClear.Location = New-Object System.Drawing.Point(10, 375)
    $btnClear.Size = New-Object System.Drawing.Size(90, 40)
    $btnClear.BackColor = [System.Drawing.Color]::Gray
    $btnClear.ForeColor = [System.Drawing.Color]::White
    $btnClear.FlatStyle = "Flat"
    $middlePanel.Controls.Add($btnClear)

    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Location = New-Object System.Drawing.Point(610, 10)
    $rightPanel.Size = New-Object System.Drawing.Size(485, 600)
    $rightPanel.BorderStyle = "FixedSingle"
    $dialog.Controls.Add($rightPanel)

    $rightLabel = New-Object System.Windows.Forms.Label
    $rightLabel.Text = "Equipos preparados para incluir"
    $rightLabel.Location = New-Object System.Drawing.Point(8, 8)
    $rightLabel.AutoSize = $true
    $rightLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $rightPanel.Controls.Add($rightLabel)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = if ($IncludeRegex) { "Filtro activo: $IncludeRegex" } else { "Sin filtro especial de nombre" }
    $infoLabel.Location = New-Object System.Drawing.Point(8, 30)
    $infoLabel.Size = New-Object System.Drawing.Size(460, 20)
    $infoLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $rightPanel.Controls.Add($infoLabel)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = "Total: 0"
    $countLabel.Location = New-Object System.Drawing.Point(390, 8)
    $countLabel.Size = New-Object System.Drawing.Size(80, 20)
    $countLabel.TextAlign = "TopRight"
    $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $rightPanel.Controls.Add($countLabel)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(8, 55)
    $listView.Size = New-Object System.Drawing.Size(466, 525)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.CheckBoxes = $true
    [void]$listView.Columns.Add("Equipo", 150)
    [void]$listView.Columns.Add("Ruta", 290)
    $rightPanel.Controls.Add($listView)

    $btnAccept = New-Object System.Windows.Forms.Button
    $btnAccept.Text = "Aceptar"
    $btnAccept.Location = New-Object System.Drawing.Point(790, 620)
    $btnAccept.Size = New-Object System.Drawing.Size(140, 40)
    $btnAccept.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnAccept.ForeColor = [System.Drawing.Color]::White
    $btnAccept.FlatStyle = "Flat"
    $dialog.Controls.Add($btnAccept)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(945, 620)
    $btnCancel.Size = New-Object System.Drawing.Size(140, 40)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $selectedEntries = @{}

    $refreshList = {
        $listView.BeginUpdate()
        $listView.Items.Clear()
        foreach ($entry in ($selectedEntries.Values | Sort-Object Name, Path)) {
            $item = New-Object System.Windows.Forms.ListViewItem($entry.Name)
            [void]$item.SubItems.Add($entry.Path)
            $item.Tag = $entry
            [void]$listView.Items.Add($item)
        }
        $listView.EndUpdate()
        $countLabel.Text = "Total: $($selectedEntries.Count)"
    }

    $addEntries = {
        param([object[]]$Entries)

        foreach ($entry in $Entries) {
            $key = if (-not [string]::IsNullOrWhiteSpace($entry.DistinguishedName)) {
                $entry.DistinguishedName.ToLowerInvariant()
            }
            else {
                $entry.Name.ToLowerInvariant()
            }

            $selectedEntries[$key] = [PSCustomObject]@{
                Name              = $entry.Name
                DistinguishedName = $entry.DistinguishedName
                Path              = $entry.Path
            }
        }

        & $refreshList
    }

    $collectEntriesFromTree = {
        $map = @{}

        function Visit-CheckedNode {
            param([System.Windows.Forms.TreeNode]$CurrentNode)

            $nodeTag = $CurrentNode.Tag
            if ($CurrentNode.Checked -and $nodeTag) {
                if ($nodeTag.Type -eq "computer") {
                    $map[$nodeTag.DistinguishedName.ToLowerInvariant()] = [PSCustomObject]@{
                        Name              = $nodeTag.Name
                        DistinguishedName = $nodeTag.DistinguishedName
                        Path              = Get-DisplayPathText $nodeTag.DistinguishedName
                    }
                    return
                }

                if ($nodeTag.Type -eq "ou") {
                    foreach ($computer in (Get-ComputersFromOU -BaseOU $nodeTag.DistinguishedName -IncludeRegex $IncludeRegex)) {
                        $map[$computer.DistinguishedName.ToLowerInvariant()] = [PSCustomObject]@{
                            Name              = $computer.Name
                            DistinguishedName = $computer.DistinguishedName
                            Path              = Get-DisplayPathText $computer.DistinguishedName
                        }
                    }
                    return
                }
            }

            foreach ($childNode in $CurrentNode.Nodes) {
                Visit-CheckedNode -CurrentNode $childNode
            }
        }

        foreach ($rootNode in $tree.Nodes) {
            Visit-CheckedNode -CurrentNode $rootNode
        }

        return @($map.Values | Sort-Object Name)
    }

    $btnAdd.Add_Click({
        $entries = & $collectEntriesFromTree
        if ($entries.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No se han seleccionado equipos ni OUs con equipos validos.",
                "Sin seleccion",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        & $addEntries $entries
    })

    $btnImportCsv.Add_Click({
        $dialogCsv = New-Object System.Windows.Forms.OpenFileDialog
        $dialogCsv.Filter = "Archivos CSV (*.csv)|*.csv"
        $dialogCsv.InitialDirectory = "C:\Users\$env:USERNAME\Desktop"

        if ($dialogCsv.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        try {
            $entries = Resolve-ComputerEntriesFromCsv -Path $dialogCsv.FileName -IncludeRegex $IncludeRegex
            & $addEntries $entries
        }
        catch {
            Show-ErrorDialog -Title "Error de CSV" -Message $_.Exception.Message
        }
    })

    $btnRemove.Add_Click({
        $checkedItems = @($listView.CheckedItems)
        foreach ($item in $checkedItems) {
            $entry = $item.Tag
            $key = if (-not [string]::IsNullOrWhiteSpace($entry.DistinguishedName)) {
                $entry.DistinguishedName.ToLowerInvariant()
            }
            else {
                $entry.Name.ToLowerInvariant()
            }
            $selectedEntries.Remove($key) | Out-Null
        }
        & $refreshList
    })

    $btnClear.Add_Click({
        $selectedEntries.Clear()
        & $refreshList
    })

    $btnAccept.Add_Click({
        if ($selectedEntries.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No hay equipos preparados para incluir.",
                "Lista vacia",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $dialog.Tag = @($selectedEntries.Values | Sort-Object Name, Path)
        $dialog.Close()
    })

    $btnCancel.Add_Click({
        $dialog.Tag = $null
        $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
    return $dialog.Tag
}

function Show-OUSelectionDialog {
    param([string]$Title)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(760, 720)
    $dialog.StartPosition = "CenterParent"
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $dialog.MinimumSize = New-Object System.Drawing.Size(600, 560)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Seleccione una OU"
    $label.Location = New-Object System.Drawing.Point(15, 12)
    $label.AutoSize = $true
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($label)

    $treeHost = New-Object System.Windows.Forms.Panel
    $treeHost.Location = New-Object System.Drawing.Point(15, 40)
    $treeHost.Size = New-Object System.Drawing.Size(710, 540)
    $treeHost.BorderStyle = "FixedSingle"
    $dialog.Controls.Add($treeHost)

    $tree = New-OUTree -Roots $Script:Config.BrowseRoots -IncludeRegex "" -UseCheckboxes $false
    $treeHost.Controls.Add($tree)

    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Text = "Ruta seleccionada: (ninguna)"
    $pathLabel.Location = New-Object System.Drawing.Point(15, 590)
    $pathLabel.Size = New-Object System.Drawing.Size(710, 40)
    $pathLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $dialog.Controls.Add($pathLabel)

    $btnNewOU = New-Object System.Windows.Forms.Button
    $btnNewOU.Text = "Crear OU hija"
    $btnNewOU.Location = New-Object System.Drawing.Point(15, 638)
    $btnNewOU.Size = New-Object System.Drawing.Size(140, 35)
    $btnNewOU.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnNewOU.ForeColor = [System.Drawing.Color]::White
    $btnNewOU.FlatStyle = "Flat"
    $dialog.Controls.Add($btnNewOU)

    $btnAccept = New-Object System.Windows.Forms.Button
    $btnAccept.Text = "Aceptar"
    $btnAccept.Location = New-Object System.Drawing.Point(505, 638)
    $btnAccept.Size = New-Object System.Drawing.Size(105, 35)
    $btnAccept.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnAccept.ForeColor = [System.Drawing.Color]::White
    $btnAccept.FlatStyle = "Flat"
    $dialog.Controls.Add($btnAccept)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(620, 638)
    $btnCancel.Size = New-Object System.Drawing.Size(105, 35)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $selectedDN = ""
    $pathLabel.Tag = ""

    $tree.Add_AfterSelect({
        param($sender, $eventArgs)
        if ($eventArgs.Node.Tag -and $eventArgs.Node.Tag.Type -eq "ou") {
            $pathLabel.Tag = $eventArgs.Node.Tag.DistinguishedName
            $pathLabel.Text = "Ruta seleccionada: $($pathLabel.Tag)"
        }
    })

    $btnNewOU.Add_Click({
        if ([string]::IsNullOrWhiteSpace($pathLabel.Tag)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Seleccione primero una OU padre.",
                "Sin OU seleccionada",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $newOuName = [Microsoft.VisualBasic.Interaction]::InputBox("Nombre de la nueva OU:", "Crear OU", "")
        if ([string]::IsNullOrWhiteSpace($newOuName)) {
            return
        }

        try {
            $params = Get-ADCommonParameters
            New-ADOrganizationalUnit -Name $newOuName.Trim() -Path $pathLabel.Tag @params -ErrorAction Stop | Out-Null
            Mostrar-Mensaje "[OK] OU creada: $newOuName"

            if ($tree.SelectedNode) {
                $tree.SelectedNode.Nodes.Clear()
                Add-PlaceholderNode -Node $tree.SelectedNode
                $tree.SelectedNode.Expand()
            }
        }
        catch {
            Show-ErrorDialog -Title "Error al crear OU" -Message $_.Exception.Message
        }
    })

    $btnAccept.Add_Click({
        if ([string]::IsNullOrWhiteSpace($pathLabel.Tag)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Seleccione una OU antes de continuar.",
                "Sin seleccion",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $dialog.Tag = $pathLabel.Tag
        $dialog.Close()
    })

    $btnCancel.Add_Click({
        $dialog.Tag = $null
        $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
    return $dialog.Tag
}

function Add-ComputerEntriesToGroup {
    param(
        [object]$GroupRecord,
        [object[]]$Entries,
        [string]$OperationLabel
    )

    $successCount = 0
    $errorCount = 0
    $params = Get-ADCommonParameters

    Mostrar-Mensaje "[...] $OperationLabel - grupo destino: $($GroupRecord.Name)"

    foreach ($entry in $Entries) {
        try {
            $memberIdentity = if (-not [string]::IsNullOrWhiteSpace($entry.DistinguishedName)) {
                $entry.DistinguishedName
            }
            else {
                (Resolve-ComputerRecord $entry.Name).DistinguishedName
            }

            Add-ADGroupMember -Identity $GroupRecord.DistinguishedName -Members $memberIdentity @params -ErrorAction Stop
            Mostrar-Mensaje "[OK] Equipo anadido: $($entry.Name)"
            $successCount++
        }
        catch {
            Mostrar-Mensaje "[X] Error con $($entry.Name): $($_.Exception.Message)"
            $errorCount++
        }
    }

    [System.Windows.Forms.MessageBox]::Show(
        "Operacion finalizada.`r`n`r`nExitosos: $successCount`r`nErrores: $errorCount",
        "Resultado",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Show-ModifyGroupDialog {
    param([object]$GroupRecord)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Modificar grupo: $($GroupRecord.Name)"
    $dialog.Size = New-Object System.Drawing.Size(1100, 700)
    $dialog.StartPosition = "CenterParent"
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Miembros directos del grupo"
    $title.Location = New-Object System.Drawing.Point(15, 10)
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = $GroupRecord.DistinguishedName
    $subtitle.Location = New-Object System.Drawing.Point(15, 35)
    $subtitle.Size = New-Object System.Drawing.Size(1040, 20)
    $subtitle.ForeColor = [System.Drawing.Color]::DarkBlue
    $dialog.Controls.Add($subtitle)

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtro:"
    $lblFilter.Location = New-Object System.Drawing.Point(15, 65)
    $lblFilter.AutoSize = $true
    $dialog.Controls.Add($lblFilter)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(65, 62)
    $txtFilter.Size = New-Object System.Drawing.Size(300, 25)
    $dialog.Controls.Add($txtFilter)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refrescar"
    $btnRefresh.Location = New-Object System.Drawing.Point(380, 60)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"
    $dialog.Controls.Add($btnRefresh)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Eliminar marcados"
    $btnRemove.Location = New-Object System.Drawing.Point(490, 60)
    $btnRemove.Size = New-Object System.Drawing.Size(160, 28)
    $btnRemove.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnRemove.ForeColor = [System.Drawing.Color]::White
    $btnRemove.FlatStyle = "Flat"
    $dialog.Controls.Add($btnRemove)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Cerrar"
    $btnClose.Location = New-Object System.Drawing.Point(955, 620)
    $btnClose.Size = New-Object System.Drawing.Size(120, 35)
    $btnClose.BackColor = [System.Drawing.Color]::Gray
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = "Flat"
    $dialog.Controls.Add($btnClose)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = ""
    $status.Location = New-Object System.Drawing.Point(15, 620)
    $status.Size = New-Object System.Drawing.Size(910, 30)
    $status.ForeColor = [System.Drawing.Color]::DarkBlue
    $dialog.Controls.Add($status)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(15, 100)
    $listView.Size = New-Object System.Drawing.Size(1060, 500)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.CheckBoxes = $true
    [void]$listView.Columns.Add("Equipo", 170)
    [void]$listView.Columns.Add("Ruta", 320)
    [void]$listView.Columns.Add("DistinguishedName", 540)
    $dialog.Controls.Add($listView)

    $allEntries = @()
    $params = Get-ADCommonParameters

    $renderEntries = {
        param([object[]]$Entries)

        $listView.BeginUpdate()
        $listView.Items.Clear()
        foreach ($entry in ($Entries | Sort-Object Name, Path)) {
            $item = New-Object System.Windows.Forms.ListViewItem($entry.Name)
            [void]$item.SubItems.Add($entry.Path)
            [void]$item.SubItems.Add($entry.DistinguishedName)
            $item.Tag = $entry
            [void]$listView.Items.Add($item)
        }
        $listView.EndUpdate()
        $status.Text = "Total visibles: $($Entries.Count)"
    }

    $loadEntries = {
        try {
            Mostrar-Mensaje "[...] Obteniendo miembros directos de '$($GroupRecord.Name)'"
            $members = @(Get-ADGroupDirectMembersRanged -Identity $GroupRecord.DistinguishedName | Where-Object { $_.objectClass -eq "computer" })
            $allEntries = @(
                $members |
                ForEach-Object {
                    [PSCustomObject]@{
                        Name              = $_.Name
                        DistinguishedName = $_.DistinguishedName
                        Path              = Get-DisplayPathText $_.DistinguishedName
                    }
                }
            )
            & $renderEntries $allEntries
            Mostrar-Mensaje "[OK] Cargados $($allEntries.Count) miembros directos"
        }
        catch {
            Show-ErrorDialog -Title "Error al cargar el grupo" -Message $_.Exception.Message
        }
    }

    $applyFilter = {
        $text = $txtFilter.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            & $renderEntries $allEntries
            return
        }

        $filtered = @(
            $allEntries |
            Where-Object {
                $_.Name -match [regex]::Escape($text) -or
                $_.Path -match [regex]::Escape($text) -or
                $_.DistinguishedName -match [regex]::Escape($text)
            }
        )
        & $renderEntries $filtered
    }

    $btnRefresh.Add_Click({ & $loadEntries })
    $txtFilter.Add_TextChanged({ & $applyFilter })

    $btnRemove.Add_Click({
        $checkedItems = @($listView.CheckedItems)
        if ($checkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Marque al menos un equipo para eliminar.",
                "Sin seleccion",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Se eliminaran $($checkedItems.Count) equipos del grupo.`r`n`r`nDesea continuar?",
            "Confirmar eliminacion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        $successCount = 0
        $errorCount = 0
        foreach ($item in $checkedItems) {
            $entry = $item.Tag
            try {
                Remove-ADGroupMember -Identity $GroupRecord.DistinguishedName -Members $entry.DistinguishedName -Confirm:$false @params -ErrorAction Stop
                Mostrar-Mensaje "[OK] Eliminado del grupo: $($entry.Name)"
                $successCount++
            }
            catch {
                Mostrar-Mensaje "[X] Error eliminando $($entry.Name): $($_.Exception.Message)"
                $errorCount++
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Operacion finalizada.`r`n`r`nExitosos: $successCount`r`nErrores: $errorCount",
            "Eliminacion completada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        & $loadEntries
    })

    $btnClose.Add_Click({ $dialog.Close() })

    & $loadEntries
    $dialog.ShowDialog() | Out-Null
}

function Set-ComputerDelegation {
    param(
        [string]$ComputerDistinguishedName,
        [string]$DelegationGroupDN
    )

    if ([string]::IsNullOrWhiteSpace($DelegationGroupDN)) {
        return
    }

    $params = Get-ADCommonParameters
    $groupSid = (Get-ADGroup -Identity $DelegationGroupDN @params -ErrorAction Stop).SID
    $ldapPath = if ([string]::IsNullOrWhiteSpace($Script:Config.Server)) {
        "LDAP://$ComputerDistinguishedName"
    }
    else {
        "LDAP://$($Script:Config.Server)/$ComputerDistinguishedName"
    }

    $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
    try {
        $acl = $directoryEntry.PsBase.ObjectSecurity

        $acl.AddAccessRule([System.DirectoryServices.ActiveDirectoryAccessRule]::new(
            $groupSid,
            [System.DirectoryServices.ActiveDirectoryRights]::GenericRead,
            [System.Security.AccessControl.AccessControlType]::Allow
        ))

        foreach ($guid in @(
            [Guid]"00299570-246d-11d0-a768-00aa006e0529",
            [Guid]"ab721a53-1e2f-11d0-9819-00aa0040529b",
            [Guid]"68b1d179-0d15-4d4f-ab71-46152e79a7bc",
            [Guid]"ab721a54-1e2f-11d0-9819-00aa0040529b",
            [Guid]"ab721a56-1e2f-11d0-9819-00aa0040529b"
        )) {
            $acl.AddAccessRule([System.DirectoryServices.ActiveDirectoryAccessRule]::new(
                $groupSid,
                [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight,
                [System.Security.AccessControl.AccessControlType]::Allow,
                $guid
            ))
        }

        foreach ($guid in @(
            [Guid]"77b5b886-944a-11d1-aebd-0000f80367c1",
            [Guid]"e48d0154-bcf8-11d1-8702-00c04fb96050",
            [Guid]"4c164200-20c0-11d0-a768-00aa006e0529"
        )) {
            $acl.AddAccessRule([System.DirectoryServices.ActiveDirectoryAccessRule]::new(
                $groupSid,
                [System.DirectoryServices.ActiveDirectoryRights]::ReadProperty,
                [System.Security.AccessControl.AccessControlType]::Allow,
                $guid
            ))
        }

        $acl.AddAccessRule([System.DirectoryServices.ActiveDirectoryAccessRule]::new(
            $groupSid,
            [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty,
            [System.Security.AccessControl.AccessControlType]::Allow,
            [Guid]"4c164200-20c0-11d0-a768-00aa006e0529"
        ))

        foreach ($guid in @(
            [Guid]"72e39547-7b18-11d1-adef-00c04fd8d5cd",
            [Guid]"f3a64788-5306-11d1-a9c5-0000f80367c1"
        )) {
            $acl.AddAccessRule([System.DirectoryServices.ActiveDirectoryAccessRule]::new(
                $groupSid,
                [System.DirectoryServices.ActiveDirectoryRights]::Self,
                [System.Security.AccessControl.AccessControlType]::Allow,
                $guid
            ))
        }

        $directoryEntry.PsBase.ObjectSecurity = $acl
        $directoryEntry.PsBase.CommitChanges()
    }
    finally {
        $directoryEntry.Dispose()
    }
}

function Show-CreateComputersDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Crear equipos"
    $dialog.Size = New-Object System.Drawing.Size(720, 520)
    $dialog.StartPosition = "CenterParent"
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $selectedOU = ""

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = "OU destino:"
    $lblTarget.Location = New-Object System.Drawing.Point(20, 20)
    $lblTarget.AutoSize = $true
    $lblTarget.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($lblTarget)

    $txtTarget = New-Object System.Windows.Forms.TextBox
    $txtTarget.Location = New-Object System.Drawing.Point(20, 45)
    $txtTarget.Size = New-Object System.Drawing.Size(520, 25)
    $txtTarget.ReadOnly = $true
    $txtTarget.Tag = ""
    $dialog.Controls.Add($txtTarget)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Seleccionar"
    $btnBrowse.Location = New-Object System.Drawing.Point(555, 42)
    $btnBrowse.Size = New-Object System.Drawing.Size(120, 30)
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnBrowse.ForeColor = [System.Drawing.Color]::White
    $btnBrowse.FlatStyle = "Flat"
    $dialog.Controls.Add($btnBrowse)

    $rowY = 95
    $fieldSpecs = @(
        @{ Label = "Prefijo"; X = 20; Text = "" ; Width = 220; Name = "Prefijo" },
        @{ Label = "Sufijo"; X = 255; Text = "" ; Width = 120; Name = "Sufijo" },
        @{ Label = "Numero inicial"; X = 390; Text = "1" ; Width = 120; Name = "Inicio" },
        @{ Label = "Numero final"; X = 525; Text = "10" ; Width = 120; Name = "Fin" }
    )

    $fields = @{}
    foreach ($spec in $fieldSpecs) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $spec.Label
        $label.Location = New-Object System.Drawing.Point($spec.X, $rowY)
        $label.AutoSize = $true
        $dialog.Controls.Add($label)

        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point($spec.X, ($rowY + 22))
        $textbox.Size = New-Object System.Drawing.Size($spec.Width, 25)
        $textbox.Text = $spec.Text
        $dialog.Controls.Add($textbox)
        $fields[$spec.Name] = $textbox
    }

    $useContainerCheckbox = New-Object System.Windows.Forms.CheckBox
    $useContainerCheckbox.Text = if ($Script:Config.ComputerContainerName) {
        "Usar o crear la OU '$($Script:Config.ComputerContainerName)' dentro de la ruta seleccionada"
    }
    else {
        "Crear directamente en la OU seleccionada"
    }
    $useContainerCheckbox.Location = New-Object System.Drawing.Point(20, 155)
    $useContainerCheckbox.Size = New-Object System.Drawing.Size(640, 20)
    $useContainerCheckbox.Checked = -not [string]::IsNullOrWhiteSpace($Script:Config.ComputerContainerName)
    $dialog.Controls.Add($useContainerCheckbox)

    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Text = "Descripcion (use {date} para fecha/hora):"
    $lblDescription.Location = New-Object System.Drawing.Point(20, 190)
    $lblDescription.AutoSize = $true
    $dialog.Controls.Add($lblDescription)

    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Location = New-Object System.Drawing.Point(20, 215)
    $txtDescription.Size = New-Object System.Drawing.Size(655, 25)
    $txtDescription.Text = $Script:Config.CreatedComputerDescriptionTemplate
    $dialog.Controls.Add($txtDescription)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Formato generado: Prefijo + numero con 3 digitos + Sufijo"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 250)
    $lblInfo.Size = New-Object System.Drawing.Size(655, 20)
    $lblInfo.ForeColor = [System.Drawing.Color]::DarkBlue
    $dialog.Controls.Add($lblInfo)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Crear equipos"
    $btnCreate.Location = New-Object System.Drawing.Point(420, 400)
    $btnCreate.Size = New-Object System.Drawing.Size(120, 40)
    $btnCreate.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnCreate.ForeColor = [System.Drawing.Color]::White
    $btnCreate.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCreate)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(555, 400)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $btnBrowse.Add_Click({
        $selected = Show-OUSelectionDialog -Title "Seleccionar OU destino"
        if ($selected) {
            $txtTarget.Tag = $selected
            $txtTarget.Text = $selected
        }
    })

    $btnCreate.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtTarget.Tag)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Seleccione una OU destino antes de continuar.",
                "Sin OU destino",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        foreach ($fieldName in @("Prefijo", "Inicio", "Fin")) {
            if ([string]::IsNullOrWhiteSpace($fields[$fieldName].Text)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Revise los datos de nombre y rango.",
                    "Datos incompletos",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }
        }

        try {
            $startNumber = [int]$fields.Inicio.Text
            $endNumber = [int]$fields.Fin.Text
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Los valores de inicio y fin deben ser numericos.",
                "Rango no valido",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        if ($startNumber -lt 1 -or $endNumber -lt $startNumber) {
            [System.Windows.Forms.MessageBox]::Show(
                "El rango indicado no es valido.",
                "Rango no valido",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $targetOU = $txtTarget.Tag
        $params = Get-ADCommonParameters

        if ($useContainerCheckbox.Checked -and -not [string]::IsNullOrWhiteSpace($Script:Config.ComputerContainerName)) {
            $containerName = $Script:Config.ComputerContainerName
            if (-not ($txtTarget.Tag -imatch "^OU=$([regex]::Escape($containerName))(,|$)")) {
                $targetOU = "OU=$containerName,$($txtTarget.Tag)"
                try {
                    $existing = Get-ADOrganizationalUnit -Identity $targetOU @params -ErrorAction SilentlyContinue
                    if (-not $existing) {
                        New-ADOrganizationalUnit -Name $containerName -Path $txtTarget.Tag @params -ErrorAction Stop | Out-Null
                        Mostrar-Mensaje "[OK] OU '$containerName' creada en $($txtTarget.Tag)"
                    }
                }
                catch {
                    Show-ErrorDialog -Title "Error al preparar la OU" -Message $_.Exception.Message
                    return
                }
            }
        }

        $computerCount = $endNumber - $startNumber + 1
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Se crearan $computerCount equipos en:`r`n`r`n$targetOU`r`n`r`nDesea continuar?",
            "Confirmar creacion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        $results = @()
        $created = 0
        $skipped = 0
        $errors = 0
        $prefix = $fields.Prefijo.Text.Trim().ToUpper()
        $suffix = $fields.Sufijo.Text.Trim().ToUpper()
        $descriptionTemplate = $txtDescription.Text

        foreach ($number in $startNumber..$endNumber) {
            $computerName = "{0}{1:D3}{2}" -f $prefix, $number, $suffix
            $safeName = Escape-ADFilterValue $computerName

            try {
                $exists = Get-ADComputer -Filter "Name -eq '$safeName'" @params -ErrorAction SilentlyContinue
                if ($exists) {
                    $results += "$computerName - YA EXISTE"
                    $skipped++
                    continue
                }

                $description = $descriptionTemplate.Replace("{date}", (Get-Date -Format "yyyy-MM-dd HH:mm"))
                $newComputer = New-ADComputer -Name $computerName -SamAccountName "$computerName`$" -Path $targetOU -Enabled $true -Description $description @params -PassThru -ErrorAction Stop

                if ($Script:Config.ApplyDelegationOnCreate -and -not [string]::IsNullOrWhiteSpace($Script:Config.DelegationGroupDN)) {
                    Set-ComputerDelegation -ComputerDistinguishedName $newComputer.DistinguishedName -DelegationGroupDN $Script:Config.DelegationGroupDN
                }

                $results += "$computerName - CREADO"
                Mostrar-Mensaje "[OK] Equipo creado: $computerName"
                $created++
            }
            catch {
                $results += "$computerName - ERROR: $($_.Exception.Message)"
                Mostrar-Mensaje ("[X] Error creando {0}: {1}" -f $computerName, $_.Exception.Message)
                $errors++
            }
        }

        Show-ResultsText -Title "CREACION DE EQUIPOS" -Lines $results

        [System.Windows.Forms.MessageBox]::Show(
            "Proceso finalizado.`r`n`r`nCreados: $created`r`nOmitidos: $skipped`r`nErrores: $errors",
            "Creacion completada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        $dialog.Close()
    })

    $btnCancel.Add_Click({ $dialog.Close() })
    $dialog.ShowDialog() | Out-Null
}

function Show-CompareComputersDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Comparar grupos de equipos"
    $dialog.Size = New-Object System.Drawing.Size(500, 220)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $fields = @(
        @{ Label = "Equipo 1"; Y = 20; Key = "Equipo1" },
        @{ Label = "Equipo 2"; Y = 80; Key = "Equipo2" }
    )
    $controls = @{}

    foreach ($spec in $fields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $spec.Label
        $label.Location = New-Object System.Drawing.Point(20, $spec.Y)
        $label.AutoSize = $true
        $dialog.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(20, ($spec.Y + 22))
        $textBox.Size = New-Object System.Drawing.Size(445, 25)
        $dialog.Controls.Add($textBox)
        $controls[$spec.Key] = $textBox
    }

    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text = "Comparar"
    $btnCompare.Location = New-Object System.Drawing.Point(250, 145)
    $btnCompare.Size = New-Object System.Drawing.Size(100, 35)
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnCompare.ForeColor = [System.Drawing.Color]::White
    $btnCompare.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCompare)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(365, 145)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.BackColor = [System.Drawing.Color]::Gray
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $btnCompare.Add_Click({
        $firstName = $controls.Equipo1.Text.Trim()
        $secondName = $controls.Equipo2.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($secondName)) {
            return
        }

        $dialog.Close()

        try {
            $firstComputer = Resolve-ComputerRecord $firstName
            $secondComputer = Resolve-ComputerRecord $secondName
            $params = Get-ADCommonParameters

            $groups1 = @(Get-ADPrincipalGroupMembership -Identity $firstComputer.DistinguishedName @params -ErrorAction Stop | Sort-Object Name | Select-Object -ExpandProperty Name)
            $groups2 = @(Get-ADPrincipalGroupMembership -Identity $secondComputer.DistinguishedName @params -ErrorAction Stop | Sort-Object Name | Select-Object -ExpandProperty Name)

            $common = @($groups1 | Where-Object { $_ -in $groups2 })
            $only1 = @($groups1 | Where-Object { $_ -notin $groups2 })
            $only2 = @($groups2 | Where-Object { $_ -notin $groups1 })

            $lines = @(
                "Equipo 1: $($firstComputer.Name)",
                "Equipo 2: $($secondComputer.Name)",
                "",
                "Total grupos equipo 1: $($groups1.Count)",
                "Total grupos equipo 2: $($groups2.Count)",
                "Grupos comunes: $($common.Count)",
                "",
                "Grupos comunes:"
            ) + ($common | ForEach-Object { "  = $_" }) + @(
                "",
                "Solo en $($firstComputer.Name):"
            ) + ($only1 | ForEach-Object { "  - $_" }) + @(
                "",
                "Solo en $($secondComputer.Name):"
            ) + ($only2 | ForEach-Object { "  - $_" })

            Show-ResultsText -Title "COMPARACION DE EQUIPOS" -Lines $lines
            Mostrar-Mensaje "[OK] Comparacion completada entre $($firstComputer.Name) y $($secondComputer.Name)"
        }
        catch {
            Show-ErrorDialog -Title "Error en comparacion" -Message $_.Exception.Message
        }
    })

    $btnCancel.Add_Click({ $dialog.Close() })
    $dialog.ShowDialog() | Out-Null
}

function Show-ExtractComputerGroupsDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Extraer grupos de un equipo"
    $dialog.Size = New-Object System.Drawing.Size(450, 170)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Nombre del equipo:"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.AutoSize = $true
    $dialog.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(20, 45)
    $textBox.Size = New-Object System.Drawing.Size(395, 25)
    $dialog.Controls.Add($textBox)

    $btnExtract = New-Object System.Windows.Forms.Button
    $btnExtract.Text = "Extraer"
    $btnExtract.Location = New-Object System.Drawing.Point(215, 90)
    $btnExtract.Size = New-Object System.Drawing.Size(95, 32)
    $btnExtract.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnExtract.ForeColor = [System.Drawing.Color]::White
    $btnExtract.FlatStyle = "Flat"
    $dialog.Controls.Add($btnExtract)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(320, 90)
    $btnCancel.Size = New-Object System.Drawing.Size(95, 32)
    $btnCancel.BackColor = [System.Drawing.Color]::Gray
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $btnExtract.Add_Click({
        $computerName = $textBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            return
        }

        $dialog.Close()

        try {
            $computer = Resolve-ComputerRecord $computerName
            $params = Get-ADCommonParameters
            $groups = @(Get-ADPrincipalGroupMembership -Identity $computer.DistinguishedName @params -ErrorAction Stop | Sort-Object Name | Select-Object -ExpandProperty Name)
            if ($groups.Count -eq 0) {
                Show-ResultsText -Title "GRUPOS DEL EQUIPO" -Lines @("Equipo: $($computer.Name)", "", "No se encontraron grupos.")
            }
            else {
                Show-ResultsText -Title "GRUPOS DEL EQUIPO" -Lines (@("Equipo: $($computer.Name)", "Total de grupos: $($groups.Count)", "") + $groups)
            }
            Mostrar-Mensaje "[OK] Extraidos grupos para $($computer.Name)"
        }
        catch {
            Show-ErrorDialog -Title "Error al extraer grupos" -Message $_.Exception.Message
        }
    })

    $btnCancel.Add_Click({ $dialog.Close() })
    $dialog.ShowDialog() | Out-Null
}

function Show-CsvToGroupsDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "CSV a grupos"
    $dialog.Size = New-Object System.Drawing.Size(760, 470)
    $dialog.StartPosition = "CenterParent"
    $dialog.BackColor = [System.Drawing.Color]::WhiteSmoke

    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Text = "Archivo CSV con equipos:"
    $lblFile.Location = New-Object System.Drawing.Point(20, 20)
    $lblFile.AutoSize = $true
    $dialog.Controls.Add($lblFile)

    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = New-Object System.Drawing.Point(20, 45)
    $txtFile.Size = New-Object System.Drawing.Size(570, 25)
    $txtFile.ReadOnly = $true
    $dialog.Controls.Add($txtFile)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Buscar"
    $btnBrowse.Location = New-Object System.Drawing.Point(605, 42)
    $btnBrowse.Size = New-Object System.Drawing.Size(120, 30)
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $btnBrowse.ForeColor = [System.Drawing.Color]::White
    $btnBrowse.FlatStyle = "Flat"
    $dialog.Controls.Add($btnBrowse)

    $lblGroups = New-Object System.Windows.Forms.Label
    $lblGroups.Text = "Grupos destino (uno por linea):"
    $lblGroups.Location = New-Object System.Drawing.Point(20, 95)
    $lblGroups.AutoSize = $true
    $dialog.Controls.Add($lblGroups)

    $txtGroups = New-Object System.Windows.Forms.TextBox
    $txtGroups.Location = New-Object System.Drawing.Point(20, 120)
    $txtGroups.Size = New-Object System.Drawing.Size(705, 200)
    $txtGroups.Multiline = $true
    $txtGroups.ScrollBars = "Vertical"
    $dialog.Controls.Add($txtGroups)

    $selectedGroup = $null
    if ($Script:ComboGrupoAD.SelectedItem -and $Script:ComboGrupoAD.SelectedItem.PSObject.Properties.Name -contains "Name") {
        $selectedGroup = $Script:ComboGrupoAD.SelectedItem.Name
    }
    elseif (-not [string]::IsNullOrWhiteSpace($Script:ComboGrupoAD.Text)) {
        $selectedGroup = $Script:ComboGrupoAD.Text.Trim()
    }

    $prefillGroups = @()
    if ($selectedGroup) {
        $prefillGroups += $selectedGroup
    }
    $prefillGroups += Convert-ToStringArray $Script:Config.FixedTargetGroups
    $txtGroups.Text = ($prefillGroups | Select-Object -Unique) -join [Environment]::NewLine

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Aplicar"
    $btnApply.Location = New-Object System.Drawing.Point(470, 370)
    $btnApply.Size = New-Object System.Drawing.Size(120, 40)
    $btnApply.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnApply.ForeColor = [System.Drawing.Color]::White
    $btnApply.FlatStyle = "Flat"
    $dialog.Controls.Add($btnApply)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(605, 370)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 40)
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatStyle = "Flat"
    $dialog.Controls.Add($btnCancel)

    $btnBrowse.Add_Click({
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "Archivos CSV (*.csv)|*.csv"
        $fileDialog.InitialDirectory = "C:\Users\$env:USERNAME\Desktop"
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtFile.Text = $fileDialog.FileName
        }
    })

    $btnApply.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtFile.Text)) {
            return
        }

        $groupNames = Convert-ToStringArray $txtGroups.Text | Select-Object -Unique
        if ($groupNames.Count -eq 0) {
            return
        }

        try {
            $entries = Resolve-ComputerEntriesFromCsv -Path $txtFile.Text -IncludeRegex ""
            $targetGroups = @($groupNames | ForEach-Object { Resolve-GroupRecord $_ })
            $summary = @()

            foreach ($group in $targetGroups) {
                Add-ComputerEntriesToGroup -GroupRecord $group -Entries $entries -OperationLabel "CSV a grupos"
                $summary += "Grupo procesado: $($group.Name)"
            }

            Show-ResultsText -Title "CSV A GRUPOS" -Lines $summary
            $dialog.Close()
        }
        catch {
            Show-ErrorDialog -Title "Error en CSV a grupos" -Message $_.Exception.Message
        }
    })

    $btnCancel.Add_Click({ $dialog.Close() })
    $dialog.ShowDialog() | Out-Null
}

function Save-WorkOUs {
    $Script:Config = Normalize-Config ([ordered]@{
        UiTitle                            = $Script:Config.UiTitle
        Server                             = $Script:Config.Server
        GroupSearchBases                   = $Script:Config.GroupSearchBases
        GroupNameFilter                    = $Script:Config.GroupNameFilter
        ExploreSearchBases                 = $Script:Config.ExploreSearchBases
        BrowseRoots                        = $Script:Config.BrowseRoots
        ComputerContainerName              = $Script:Config.ComputerContainerName
        HiddenOUSegments                   = $Script:Config.HiddenOUSegments
        ExcludedOUPatterns                 = $Script:Config.ExcludedOUPatterns
        ExcludedComputerNameRegex          = $Script:Config.ExcludedComputerNameRegex
        FilteredComputerLabel              = $Script:Config.FilteredComputerLabel
        FilteredComputerIncludeRegex       = $Script:Config.FilteredComputerIncludeRegex
        ApplyDelegationOnCreate            = $Script:Config.ApplyDelegationOnCreate
        DelegationGroupDN                  = $Script:Config.DelegationGroupDN
        CreatedComputerDescriptionTemplate = $Script:Config.CreatedComputerDescriptionTemplate
        FixedTargetGroups                  = $Script:Config.FixedTargetGroups
        SavedWorkOUs                       = @($Script:SelectedWorkOUs)
    })
    Save-Config $Script:Config | Out-Null
    Mostrar-Mensaje "[OK] OUs de trabajo guardadas en configuracion ($($Script:SelectedWorkOUs.Count) OUs)"
}

function Load-SavedWorkOUs {
    $saved = $Script:Config.SavedWorkOUs
    if ($saved -and $saved.Count -gt 0) {
        $Script:SelectedWorkOUs.Clear()
        foreach ($ou in (Convert-ToStringArray $saved)) {
            $Script:SelectedWorkOUs.Add($ou)
        }
        Refresh-OUListBox
        Mostrar-Mensaje "[OK] $($Script:SelectedWorkOUs.Count) OU(s) de trabajo restauradas desde configuracion"
    }
}

function Invoke-ExtractSelectedGroupComputers {
    $group = Get-SelectedGroupRecord
    if (-not $group) {
        return
    }

    try {
        Mostrar-Mensaje "[...] Extrayendo equipos del grupo '$($group.Name)'"
        $computers = @(
            Get-ADGroupMemberSafe -Identity $group.DistinguishedName -Recursive |
            Where-Object { $_.objectClass -eq "computer" } |
            Select-Object Name, DistinguishedName
        )

        if ($computers.Count -eq 0) {
            Show-ResultsText -Title "EQUIPOS DEL GRUPO" -Lines @("No se encontraron equipos en '$($group.Name)'.")
            return
        }

        Show-ResultsComputerTree -Title "EQUIPOS DEL GRUPO $($group.Name)" -Computers $computers
        Mostrar-Mensaje "[OK] Extraidos $($computers.Count) equipos"
    }
    catch {
        Show-ErrorDialog -Title "Error al extraer equipos" -Message $_.Exception.Message
    }
}

function Invoke-ExploreEnvironmentComputers {
    try {
        Mostrar-Mensaje "[...] Explorando equipos en el entorno configurado"
        $computers = Get-ComputersFromSearchBases -SearchBases $Script:Config.ExploreSearchBases -IncludeRegex ""
        if ($computers.Count -eq 0) {
            Show-ResultsText -Title "EXPLORACION DE EQUIPOS" -Lines @("No se encontraron equipos en las bases configuradas.")
            return
        }

        Show-ResultsComputerTree -Title "EXPLORACION DE EQUIPOS" -Computers $computers
        Mostrar-Mensaje "[OK] Exploracion completada: $($computers.Count) equipos"
    }
    catch {
        Show-ErrorDialog -Title "Error al explorar equipos" -Message $_.Exception.Message
    }
}

function Invoke-AddComputersToGroup {
    param([string]$IncludeRegex)

    $group = Get-SelectedGroupRecord
    if (-not $group) {
        return
    }

    $selection = Show-ComputerSelectionDialog -Title "Seleccionar equipos para $($group.Name)" -IncludeRegex $IncludeRegex
    if (-not $selection) {
        return
    }

    Add-ComputerEntriesToGroup -GroupRecord $group -Entries $selection -OperationLabel "Inclusion de equipos"
}

Load-Config

$Script:MainForm = New-Object System.Windows.Forms.Form
$Script:MainForm.Text = $Script:Config.UiTitle
$Script:MainForm.StartPosition = "CenterScreen"
$Script:MainForm.Size = New-Object System.Drawing.Size(1500, 820)
$Script:MainForm.MinimumSize = New-Object System.Drawing.Size(1280, 760)
$Script:MainForm.BackColor = [System.Drawing.Color]::WhiteSmoke
$Script:MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Gestion general de grupos y equipos AD"
$titleLabel.Location = New-Object System.Drawing.Point(15, 15)
$titleLabel.AutoSize = $true
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 60, 120)
$Script:MainForm.Controls.Add($titleLabel)



$lblOUsTrabajo = New-Object System.Windows.Forms.Label
$lblOUsTrabajo.Text = "OUs de trabajo:"
$lblOUsTrabajo.Location = New-Object System.Drawing.Point(18, 78)
$lblOUsTrabajo.AutoSize = $true
$lblOUsTrabajo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$Script:MainForm.Controls.Add($lblOUsTrabajo)

$Script:ListBoxOUs = New-Object System.Windows.Forms.ListBox
$Script:ListBoxOUs.Location = New-Object System.Drawing.Point(130, 75)
$Script:ListBoxOUs.Size = New-Object System.Drawing.Size(400, 56)
$Script:ListBoxOUs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Script:ListBoxOUs.BorderStyle = "FixedSingle"
$Script:ListBoxOUs.HorizontalScrollbar = $true
$Script:MainForm.Controls.Add($Script:ListBoxOUs)

$ouToolTip = New-Object System.Windows.Forms.ToolTip
$ouToolTipIndex = -1
$Script:ListBoxOUs.Add_MouseMove({
    param($sender, $eventArgs)
    $idx = $sender.IndexFromPoint($eventArgs.Location)
    if ($idx -ge 0 -and $idx -lt $Script:SelectedWorkOUs.Count -and $idx -ne $ouToolTipIndex) {
        $ouToolTipIndex = $idx
        $ouToolTip.SetToolTip($sender, $Script:SelectedWorkOUs[$idx])
    }
    elseif ($idx -lt 0) {
        $ouToolTipIndex = -1
        $ouToolTip.SetToolTip($sender, "")
    }
})

$btnAddOU = New-Object System.Windows.Forms.Button
$btnAddOU.Text = "+ OU"
$btnAddOU.Location = New-Object System.Drawing.Point(538, 75)
$btnAddOU.Size = New-Object System.Drawing.Size(55, 27)
$btnAddOU.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$btnAddOU.ForeColor = [System.Drawing.Color]::White
$btnAddOU.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnAddOU)

$btnRemoveOU = New-Object System.Windows.Forms.Button
$btnRemoveOU.Text = "- OU"
$btnRemoveOU.Location = New-Object System.Drawing.Point(538, 105)
$btnRemoveOU.Size = New-Object System.Drawing.Size(55, 27)
$btnRemoveOU.BackColor = [System.Drawing.Color]::FromArgb(211, 47, 47)
$btnRemoveOU.ForeColor = [System.Drawing.Color]::White
$btnRemoveOU.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnRemoveOU)

$btnSaveOUs = New-Object System.Windows.Forms.Button
$btnSaveOUs.Text = "Guardar OUs"
$btnSaveOUs.Location = New-Object System.Drawing.Point(600, 75)
$btnSaveOUs.Size = New-Object System.Drawing.Size(90, 56)
$btnSaveOUs.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$btnSaveOUs.ForeColor = [System.Drawing.Color]::White
$btnSaveOUs.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnSaveOUs)

$lblGrupo = New-Object System.Windows.Forms.Label
$lblGrupo.Text = "Grupo AD:"
$lblGrupo.Location = New-Object System.Drawing.Point(18, 142)
$lblGrupo.AutoSize = $true
$lblGrupo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$Script:MainForm.Controls.Add($lblGrupo)

$Script:ComboGrupoAD = New-Object System.Windows.Forms.ComboBox
$Script:ComboGrupoAD.Location = New-Object System.Drawing.Point(95, 139)
$Script:ComboGrupoAD.Size = New-Object System.Drawing.Size(435, 25)
$Script:ComboGrupoAD.DropDownStyle = "DropDown"
$Script:ComboGrupoAD.AutoCompleteMode = "SuggestAppend"
$Script:ComboGrupoAD.AutoCompleteSource = "CustomSource"
$Script:ComboGrupoAD.DisplayMember = "DisplayName"
$Script:MainForm.Controls.Add($Script:ComboGrupoAD)

$btnCargarGrupos = New-Object System.Windows.Forms.Button
$btnCargarGrupos.Text = "Cargar Grupos"
$btnCargarGrupos.Location = New-Object System.Drawing.Point(538, 137)
$btnCargarGrupos.Size = New-Object System.Drawing.Size(110, 28)
$btnCargarGrupos.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
$btnCargarGrupos.ForeColor = [System.Drawing.Color]::White
$btnCargarGrupos.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnCargarGrupos)

$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(10, 175)
$actionPanel.Size = New-Object System.Drawing.Size(650, 250)
$Script:MainForm.Controls.Add($actionPanel)

function New-ActionButton {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size(200, 60)
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.FlatStyle = "Flat"
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    return $button
}

$btnExtraer = New-ActionButton -Text "Extraer Equipos" -X 0 -Y 0
$btnExplorar = New-ActionButton -Text "Explorar Equipos" -X 220 -Y 0
$btnComparar = New-ActionButton -Text "Comparar Equipos" -X 440 -Y 0
$btnIncluir = New-ActionButton -Text "Incluir Equipos" -X 0 -Y 80
$btnModificar = New-ActionButton -Text "Modificar Grupo" -X 220 -Y 80
$btnExtraerGrupos = New-ActionButton -Text "Extraer Grupos" -X 440 -Y 80
$btnCrearEquipos = New-ActionButton -Text "Crear Equipos" -X 0 -Y 160
$Script:BtnFiltroEspecial = New-ActionButton -Text "Anadir Equipos Filtrados" -X 220 -Y 160
$btnCsvAGrupos = New-ActionButton -Text "CSV a Grupos" -X 440 -Y 160

$actionPanel.Controls.AddRange(@(
    $btnExtraer,
    $btnExplorar,
    $btnComparar,
    $btnIncluir,
    $btnModificar,
    $btnExtraerGrupos,
    $btnCrearEquipos,
    $Script:BtnFiltroEspecial,
    $btnCsvAGrupos
))

$logsLabel = New-Object System.Windows.Forms.Label
$logsLabel.Text = "Registro de actividad"
$logsLabel.Location = New-Object System.Drawing.Point(15, 440)
$logsLabel.AutoSize = $true
$logsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$Script:MainForm.Controls.Add($logsLabel)

$Script:TxtSalida = New-Object System.Windows.Forms.TextBox
$Script:TxtSalida.Location = New-Object System.Drawing.Point(15, 463)
$Script:TxtSalida.Size = New-Object System.Drawing.Size(645, 270)
$Script:TxtSalida.Multiline = $true
$Script:TxtSalida.ScrollBars = "Vertical"
$Script:TxtSalida.ReadOnly = $true
$Script:TxtSalida.Font = New-Object System.Drawing.Font("Consolas", 9)
$Script:MainForm.Controls.Add($Script:TxtSalida)

$resultLabel = New-Object System.Windows.Forms.Label
$resultLabel.Text = "Resultados"
$resultLabel.Location = New-Object System.Drawing.Point(710, 48)
$resultLabel.AutoSize = $true
$resultLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$Script:MainForm.Controls.Add($resultLabel)

$Script:LblContadorEquipos = New-Object System.Windows.Forms.Label
$Script:LblContadorEquipos.Location = New-Object System.Drawing.Point(830, 51)
$Script:LblContadorEquipos.Size = New-Object System.Drawing.Size(430, 20)
$Script:LblContadorEquipos.ForeColor = [System.Drawing.Color]::DarkBlue
$Script:LblContadorEquipos.Visible = $false
$Script:MainForm.Controls.Add($Script:LblContadorEquipos)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Buscar:"
$lblSearch.Location = New-Object System.Drawing.Point(1040, 51)
$lblSearch.AutoSize = $true
$Script:MainForm.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(1090, 48)
$txtSearch.Size = New-Object System.Drawing.Size(220, 25)
$Script:MainForm.Controls.Add($txtSearch)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Buscar"
$btnSearch.Location = New-Object System.Drawing.Point(1318, 47)
$btnSearch.Size = New-Object System.Drawing.Size(70, 27)
$btnSearch.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
$btnSearch.ForeColor = [System.Drawing.Color]::White
$btnSearch.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnSearch)

$btnNextSearch = New-Object System.Windows.Forms.Button
$btnNextSearch.Text = "Siguiente"
$btnNextSearch.Location = New-Object System.Drawing.Point(1396, 47)
$btnNextSearch.Size = New-Object System.Drawing.Size(80, 27)
$btnNextSearch.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
$btnNextSearch.ForeColor = [System.Drawing.Color]::White
$btnNextSearch.FlatStyle = "Flat"
$Script:MainForm.Controls.Add($btnNextSearch)

$Script:TxtResultados = New-Object System.Windows.Forms.TextBox
$Script:TxtResultados.Location = New-Object System.Drawing.Point(710, 78)
$Script:TxtResultados.Size = New-Object System.Drawing.Size(765, 630)
$Script:TxtResultados.Multiline = $true
$Script:TxtResultados.ScrollBars = "Vertical"
$Script:TxtResultados.ReadOnly = $true
$Script:TxtResultados.Font = New-Object System.Drawing.Font("Consolas", 10)
$Script:MainForm.Controls.Add($Script:TxtResultados)

$Script:TreeResultados = New-Object System.Windows.Forms.TreeView
$Script:TreeResultados.Location = New-Object System.Drawing.Point(710, 78)
$Script:TreeResultados.Size = New-Object System.Drawing.Size(765, 630)
$Script:TreeResultados.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Script:TreeResultados.Visible = $false
$Script:MainForm.Controls.Add($Script:TreeResultados)

$resultsButtonPanel = New-Object System.Windows.Forms.Panel
$resultsButtonPanel.Location = New-Object System.Drawing.Point(710, 712)
$resultsButtonPanel.Size = New-Object System.Drawing.Size(765, 32)
$Script:MainForm.Controls.Add($resultsButtonPanel)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Copiar"
$btnCopy.Location = New-Object System.Drawing.Point(0, 0)
$btnCopy.Size = New-Object System.Drawing.Size(110, 30)
$btnCopy.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
$btnCopy.ForeColor = [System.Drawing.Color]::White
$btnCopy.FlatStyle = "Flat"
$resultsButtonPanel.Controls.Add($btnCopy)

$btnClearResults = New-Object System.Windows.Forms.Button
$btnClearResults.Text = "Limpiar"
$btnClearResults.Location = New-Object System.Drawing.Point(120, 0)
$btnClearResults.Size = New-Object System.Drawing.Size(110, 30)
$btnClearResults.BackColor = [System.Drawing.Color]::Gray
$btnClearResults.ForeColor = [System.Drawing.Color]::White
$btnClearResults.FlatStyle = "Flat"
$resultsButtonPanel.Controls.Add($btnClearResults)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Exportar CSV"
$btnExportCsv.Location = New-Object System.Drawing.Point(240, 0)
$btnExportCsv.Size = New-Object System.Drawing.Size(120, 30)
$btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$btnExportCsv.ForeColor = [System.Drawing.Color]::White
$btnExportCsv.FlatStyle = "Flat"
$resultsButtonPanel.Controls.Add($btnExportCsv)

$btnSaveTxt = New-Object System.Windows.Forms.Button
$btnSaveTxt.Text = "Guardar TXT"
$btnSaveTxt.Location = New-Object System.Drawing.Point(370, 0)
$btnSaveTxt.Size = New-Object System.Drawing.Size(120, 30)
$btnSaveTxt.BackColor = [System.Drawing.Color]::FromArgb(0, 106, 170)
$btnSaveTxt.ForeColor = [System.Drawing.Color]::White
$btnSaveTxt.FlatStyle = "Flat"
$resultsButtonPanel.Controls.Add($btnSaveTxt)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$Script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$Script:StatusLabel.Spring = $true
$Script:StatusLabel.TextAlign = "MiddleLeft"
$Script:StatusLabel.Text = "Listo"
$statusStrip.Items.Add($Script:StatusLabel) | Out-Null
$Script:MainForm.Controls.Add($statusStrip)

$btnAddOU.Add_Click({ Add-WorkOU })
$btnRemoveOU.Add_Click({ Remove-WorkOU })
$btnSaveOUs.Add_Click({ Save-WorkOUs })
$btnCargarGrupos.Add_Click({ Load-GroupList })
$Script:ComboGrupoAD.Add_SelectedIndexChanged({
    if ($Script:ComboGrupoAD.SelectedItem -and $Script:ComboGrupoAD.SelectedItem.Name) {
        Mostrar-Mensaje "[OK] Grupo seleccionado: $($Script:ComboGrupoAD.SelectedItem.Name)"
    }
})

$btnExtraer.Add_Click({ Invoke-ExtractSelectedGroupComputers })
$btnExplorar.Add_Click({ Invoke-ExploreEnvironmentComputers })
$btnComparar.Add_Click({ Show-CompareComputersDialog })
$btnIncluir.Add_Click({ Invoke-AddComputersToGroup -IncludeRegex "" })
$Script:BtnFiltroEspecial.Add_Click({ Invoke-AddComputersToGroup -IncludeRegex $Script:Config.FilteredComputerIncludeRegex })
$btnModificar.Add_Click({
    $group = Get-SelectedGroupRecord
    if ($group) {
        Show-ModifyGroupDialog -GroupRecord $group
    }
})
$btnExtraerGrupos.Add_Click({ Show-ExtractComputerGroupsDialog })
$btnCrearEquipos.Add_Click({ Show-CreateComputersDialog })
$btnCsvAGrupos.Add_Click({ Show-CsvToGroupsDialog })

$btnCopy.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($Script:LastTextResult)) {
        [System.Windows.Forms.Clipboard]::SetText($Script:LastTextResult)
        Mostrar-Mensaje "[OK] Resultados copiados al portapapeles"
    }
})

$btnClearResults.Add_Click({ Clear-Results })

$btnExportCsv.Add_Click({
    if ($Script:LastComputerResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hay resultados de equipos para exportar a CSV.",
            "Sin datos",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Archivo CSV (*.csv)|*.csv"
    $saveDialog.InitialDirectory = "C:\Users\$env:USERNAME\Desktop"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:LastComputerResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        Mostrar-Mensaje "[OK] CSV exportado a $($saveDialog.FileName)"
    }
})

$btnSaveTxt.Add_Click({
    if ([string]::IsNullOrWhiteSpace($Script:LastTextResult)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hay texto de resultados para guardar.",
            "Sin datos",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Archivo de texto (*.txt)|*.txt"
    $saveDialog.InitialDirectory = "C:\Users\$env:USERNAME\Desktop"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-Content -Path $saveDialog.FileName -Value $Script:LastTextResult -Encoding UTF8
        Mostrar-Mensaje "[OK] TXT guardado en $($saveDialog.FileName)"
    }
})

$btnSearch.Add_Click({
    $searchText = $txtSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        return
    }

    if ($Script:TreeResultados.Visible) {
        $Script:ResultSearchNodes = @()
        foreach ($rootNode in $Script:TreeResultados.Nodes) {
            $Script:ResultSearchNodes += Search-LoadedTreeNodes -Node $rootNode -Text $searchText
        }

        if ($Script:ResultSearchNodes.Count -eq 0) {
            Mostrar-Mensaje "[!] No se encontraron coincidencias en el arbol"
            return
        }

        $Script:ResultSearchIndex = 0
        Focus-TreeNode -TreeView $Script:TreeResultados -Node $Script:ResultSearchNodes[0]
        Mostrar-Mensaje "[OK] Coincidencias encontradas: $($Script:ResultSearchNodes.Count)"
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($Script:TxtResultados.Text)) {
        $Script:ResultSearchPositions = @()
        $text = $Script:TxtResultados.Text
        $current = 0
        while ($current -lt $text.Length) {
            $position = $text.IndexOf($searchText, $current, [System.StringComparison]::OrdinalIgnoreCase)
            if ($position -lt 0) {
                break
            }
            $Script:ResultSearchPositions += $position
            $current = $position + 1
        }

        if ($Script:ResultSearchPositions.Count -eq 0) {
            Mostrar-Mensaje "[!] No se encontraron coincidencias en el texto"
            return
        }

        $Script:ResultSearchIndex = 0
        $Script:TxtResultados.Select($Script:ResultSearchPositions[0], $searchText.Length)
        $Script:TxtResultados.ScrollToCaret()
        $Script:TxtResultados.Focus()
        Mostrar-Mensaje "[OK] Coincidencias encontradas: $($Script:ResultSearchPositions.Count)"
    }
})

$btnNextSearch.Add_Click({
    if ($Script:TreeResultados.Visible -and $Script:ResultSearchNodes.Count -gt 0) {
        $Script:ResultSearchIndex = ($Script:ResultSearchIndex + 1) % $Script:ResultSearchNodes.Count
        Focus-TreeNode -TreeView $Script:TreeResultados -Node $Script:ResultSearchNodes[$Script:ResultSearchIndex]
        return
    }

    if ($Script:ResultSearchPositions.Count -gt 0) {
        $searchText = $txtSearch.Text.Trim()
        $Script:ResultSearchIndex = ($Script:ResultSearchIndex + 1) % $Script:ResultSearchPositions.Count
        $Script:TxtResultados.Select($Script:ResultSearchPositions[$Script:ResultSearchIndex], $searchText.Length)
        $Script:TxtResultados.ScrollToCaret()
        $Script:TxtResultados.Focus()
    }
})

$txtSearch.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $btnSearch.PerformClick()
        $_.SuppressKeyPress = $true
    }
})

$Script:BtnFiltroEspecial.Text = if ([string]::IsNullOrWhiteSpace($Script:Config.FilteredComputerLabel)) { "Anadir Equipos Filtrados" } else { $Script:Config.FilteredComputerLabel }
Clear-Results
Load-SavedWorkOUs
Mostrar-Mensaje "Aplicacion inicializada. Seleccione OUs de trabajo y cargue grupos para empezar."

$Script:MainForm.ShowDialog() | Out-Null
