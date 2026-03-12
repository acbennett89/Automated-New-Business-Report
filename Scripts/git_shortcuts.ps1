param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("status", "branch", "add", "commit", "push", "new-branch", "to-main", "log", "last")]
    [string]$Action,

    [string]$Message,
    [string]$Name
)

$ErrorActionPreference = "Stop"

switch ($Action) {
    "status" {
        git status
    }
    "branch" {
        git branch --show-current
    }
    "add" {
        git add .
        Write-Host "Staged all changes."
    }
    "commit" {
        if (-not $Message) {
            throw "Provide -Message for commit."
        }
        git commit -m $Message
    }
    "push" {
        git push
    }
    "new-branch" {
        if (-not $Name) {
            throw "Provide -Name for new branch. Example: -Name 'feat/sales-velocity-layout'"
        }
        git checkout -b $Name
        git push -u origin $Name
    }
    "to-main" {
        git checkout main
        git pull
    }
    "log" {
        git log --oneline -n 10
    }
    "last" {
        git log --oneline -n 1
    }
}
