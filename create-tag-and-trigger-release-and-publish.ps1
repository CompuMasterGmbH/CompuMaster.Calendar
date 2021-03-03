param(
    [switch]$NonInteractive,
    [string]$TagAsVersion,
    $TagMessage = $null
)

## Constants of project
$assemblyVersionSourceFilePath = ".\CompuMaster.Calendar\AssemblyInfo.vb"
$assemblyVersionSourceVersionPattern = '^.*AssemblyVersion\(\"(.*)\"\).*$' #works for AssemblyInfo.cs as well as AssemblyInfo.vb

## Check for master branch
$branch = git branch
if ($branch -ne '* master')
{
    Throw "Branch must be set to master before"
}

## Check for missing commits
[string]$changes = git status -s
if($LASTEXITCODE -ne 0)
{
    Throw "Some git commits are missing: git status failed"
}
if($changes -ne "" -and $changes -ne $null)
{
    "Missing commits for:"
    $changes
    Throw "Some git commits are missing: there are file modifications present"
}

## Show existing list of tags
"Existing tags:"
git tag -n5 #max 5 lines per message
""

## Read version default from source file
[string]$versionInput = $TagAsVersion
if ($versionInput -eq "" -and $assemblyVersionSourceFilePath -ne $null -and $assemblyVersionSourceFilePath -ne "" -and (Test-Path -Path $assemblyVersionSourceFilePath))
{
    Write-Host "Probing $assemblyVersionSourceFilePath for version no. with $assemblyVersionSourceVersionPattern"
    [string]$assemblyVersionSource = Get-Content -Path $assemblyVersionSourceFilePath
    $found = $assemblyVersionSource -match $assemblyVersionSourceVersionPattern
    if ($found) 
    {
        $versionFromSourceFile = $matches[1]
    }
}

## Read version from user input
if ($NonInteractive -eq $false)
{
    if ($versionInput -eq $null -or $versionInput -eq "")
    {
        if ($versionFromSourceFile -ne "")
        {
            $versionInput = Read-Host "Please enter a version no. for the new git tag [$versionFromSourceFile]"
            if ($versionInput -eq '')
            {
                $versionInput = $versionFromSourceFile
            }
        }
        else
        {
            $versionInput = Read-Host "Please enter a version no. for the new git tag"
        }
    }
}
if ($versionInput -eq '')
{
    Throw "Version no. for tag must be defined"
}
elseif ([Version]::TryParse($versionInput, [ref]$null) -eq $false)
{
    Throw "Invald version no."
}
[string]$tagName = "v" + $versionInput

## Read tag message from user input
[string]$tagBody = $TagMessage
if ($NonInteractive -eq $false)
{
    if ($null -eq $TagMessage)
    {
        $tagBody = Read-Host "Please enter a description/message/body for the new git tag (empty = auto from commits, line breaks with <Shift>+<Enter>)"
    }
}

## Start creation of tag
"Creating tag ""$tagName"" . . ."
$tagBody
""
if ($tagBody -eq "")
{
    #$lastCommit = git log -1 --format=%H
    #if($LASTEXITCODE -ne 0)
    #{
    #    Throw "Git last commit detection failed"
    #}
    #git tag -f -a "$tagName" $lastCommit
    $lastCommitMessage = git log -1 --format=%s
    if($LASTEXITCODE -ne 0)
    {
        Throw "Git last commit detection failed"
    }
    git tag -f -a "$tagName" -m "$lastCommitMessage"
}
else
{
    git tag -f -a "$tagName" -m "$tagBody"
}
if($LASTEXITCODE -ne 0)
{
    Throw "Git tagging failed"
}
""

## Pushing tag to remote
"Pushing tag ""$tagName"" to remote . . ."
git push origin "$tagName"
if($LASTEXITCODE -ne 0)
{
    Throw "Git push failed"
}
""
