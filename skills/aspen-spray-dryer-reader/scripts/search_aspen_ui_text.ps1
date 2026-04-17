[CmdletBinding()]
param(
    [int]$ProcessId,
    [string[]]$Terms = @(),
    [switch]$Exact,
    [switch]$Json
)

Add-Type -AssemblyName UIAutomationClient

function Join-CodePoints {
    param(
        [int[]]$CodePoints
    )

    -join ($CodePoints | ForEach-Object { [char]$_ })
}

if (-not $Terms -or $Terms.Count -eq 0) {
    $Terms = @(
        'DRYER',
        'Results',
        'Summary',
        'Outlet',
        'Temperature',
        (Join-CodePoints @(0x7ED3, 0x679C)),
        (Join-CodePoints @(0x6D41, 0x80A1, 0x7ED3, 0x679C)),
        (Join-CodePoints @(0x6458, 0x8981)),
        (Join-CodePoints @(0x51FA, 0x98CE, 0x6E29, 0x5EA6)),
        (Join-CodePoints @(0x6C7D, 0x76F8, 0x6E29, 0x5EA6))
    )
}

if (-not $ProcessId) {
    $ProcessId = (
        Get-Process |
        Where-Object {
            ($_.ProcessName -match 'Aspen|Apwn|Hysys' -or $_.MainWindowTitle -match 'Aspen') -and
            $_.MainWindowTitle
        } |
        Select-Object -First 1 -ExpandProperty Id
    )
}

if (-not $ProcessId) {
    throw 'No visible Aspen window was found.'
}

$process = Get-Process -Id $ProcessId -ErrorAction Stop
$root = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
$all = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

$matches = New-Object System.Collections.Generic.List[object]

for ($i = 0; $i -lt $all.Count; $i++) {
    $element = $all.Item($i)
    $name = $element.Current.Name

    if ([string]::IsNullOrWhiteSpace($name)) {
        continue
    }

    foreach ($term in $Terms) {
        $isMatch = if ($Exact) {
            $name -eq $term
        } else {
            $name -like "*$term*"
        }

        if ($isMatch) {
            $rect = $element.Current.BoundingRectangle
            $matches.Add([pscustomobject]@{
                Term        = $term
                Name        = $name
                ClassName   = $element.Current.ClassName
                ControlType = $element.Current.ControlType.ProgrammaticName
                Left        = $rect.Left
                Top         = $rect.Top
                Width       = $rect.Width
                Height      = $rect.Height
            })
            break
        }
    }
}

if ($Json) {
    $matches | ConvertTo-Json -Depth 3
    exit 0
}

$matches | Sort-Object Term, Name | Format-Table -AutoSize
