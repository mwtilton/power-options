$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Invoke-SetupGit" {
    Context "Sets the user information"{
        Mock Read-Host {}
        
        Invoke-SetupGit

        It "calls read-host" {
            Assert-MockCalled -CommandName read-host -Exactly 2 -Scope Context
        }
        It "sets the user" {
            git config --get-regexp user.name | Should BeLike "*mwtilton*"
        }
        It "sets the email" {
            git config --get-regexp user.email | Should BeLike "*sandamomivo@gmail.com*"
        }
    }
    Context "Creates the aliases" {
        $gitAliases = "last","psu","NUKE","stashes","s"
        $gitAliases | ForEach-Object{
            It "$_ alias has been set" {
                git config --get-regexp alias.$_ | Should BeLike "*alias.$_*"
            }
        }
        
    }
    Context "Sets up the editor" {
        It "has vscode as the default editor" {
            git config --get-regexp core.editor | Should BeLike "*code*"
        }
    }
    Context "Displays the new setup" {
        Mock Read-Host {}
        Mock git {}
        
        Invoke-SetupGit

        It "shows end results" {
            Assert-MockCalled -CommandName git #-Exactly 2 -Scope Context
        }
    }
}
