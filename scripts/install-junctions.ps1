param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$CodexSkills = (Join-Path $env:USERPROFILE ".codex\skills")
)

$skillNames = @(
    "business-travel-workflow",
    "travel-application",
    "travel-daily-report",
    "travel-closure",
    "company-expense-reimbursement"
)

if (-not (Test-Path -LiteralPath $CodexSkills)) {
    New-Item -ItemType Directory -Path $CodexSkills -Force | Out-Null
}

foreach ($name in $skillNames) {
    $source = Join-Path $RepoRoot (Join-Path "skills" $name)
    $target = Join-Path $CodexSkills $name

    if (-not (Test-Path -LiteralPath $source -PathType Container)) {
        throw "Skill source does not exist: $source"
    }

    if (Test-Path -LiteralPath $target) {
        $item = Get-Item -LiteralPath $target -Force
        $resolved = $item.Target
        if ($item.LinkType -eq "Junction" -and $resolved -and ((Resolve-Path -LiteralPath $resolved).Path -eq (Resolve-Path -LiteralPath $source).Path)) {
            Write-Output "Already linked: $name"
            continue
        }
        throw "Target already exists and is not the expected Junction: $target"
    }

    New-Item -ItemType Junction -Path $target -Target (Resolve-Path -LiteralPath $source).Path | Out-Null
    Write-Output "Linked: $target -> $source"
}
