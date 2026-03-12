# Git Cheatsheet (This Project)

## 1. Check where you are
```powershell
git branch --show-current
git status
```

## 2. Save work
```powershell
git add .
git commit -m "your message"
git push
```

## 3. Start a new feature branch
```powershell
git checkout -b feat/your-feature-name
git push -u origin feat/your-feature-name
```

## 4. Switch back to main
```powershell
git checkout main
git pull
```

## 5. See recent commits
```powershell
git log --oneline -n 10
```

## PowerShell "autofill" tips (no memorizing)

### A) Use command history search
- Press `UpArrow` to cycle prior commands.
- Press `Ctrl+R` to search command history.

### B) Use tab completion
- Type part of a command/branch/file and press `Tab`.
- Example:
  - `git checkout f` + `Tab` completes `feat/sales-velocity-layout`.

### C) Use project shortcuts script
You can run:
```powershell
.\Scripts\git_shortcuts.ps1 status
```

Other actions:
```powershell
.\Scripts\git_shortcuts.ps1 branch
.\Scripts\git_shortcuts.ps1 add
.\Scripts\git_shortcuts.ps1 commit -Message "adjust sales velocity layout"
.\Scripts\git_shortcuts.ps1 push
.\Scripts\git_shortcuts.ps1 new-branch -Name "feat/my-change"
.\Scripts\git_shortcuts.ps1 to-main
.\Scripts\git_shortcuts.ps1 log
.\Scripts\git_shortcuts.ps1 last
```
