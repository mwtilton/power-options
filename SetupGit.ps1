Function Invoke-SetupGit {
    Write-host "Clearing pre-exisitng config --global settings"
    #git config --global --unset-all alias
    #git config --global --unset-all core
    git config --global --unset-all credential.helper
    git config --global --unset-all user.name

    $user = Read-Host "User Name"
    git config --global user.name "$user"
    $email = Read-Host "User Email"
    git config --global user.email "$email"

    git config --global core.autocrlf true
    git config --global alias.last "log -1 HEAD"
    git config --global alias.psu "push --set-upstream"
    git config --global alias.NUKE "reset --hard HEAD"
    git config --global alias.stashes "stash list"
    git config --global alias.s "status -s"
    git config --global core.editor "code --wait --new-window"
    git config --global credential.helper store

    git config --get-regexp alias
    git config --get-regexp user
    git config --get-regexp core.editor
    git config --get-regexp credential
    git config --get-regexp diff.tool
    git config --get-regexp difftool
}
