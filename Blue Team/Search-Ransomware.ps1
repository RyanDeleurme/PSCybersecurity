# file search to find all files with the name and return them to the user
#install-module PSSearchEverything is needed
Function Search-Ransomware
{
    param(
    [Parameter(Mandatory=$true)]
    [string]
    $ext,

    [Parameter(Mandatory=$true)]
    [string]
    $processName
    )
    #global variables
    $VerbosePreference = 'continue'
    $ErrorActionPreference = 'silentlycontinue'

    #Searching files in all directories 
    $ransomware = New-Object System.Collections.Generic.List[System.Object] #storage we will use to display all the ransomware found on the computer
    $files = Search-Everything -Extension $ext -Global
    foreach($f in $files)
    {
        Write-Verbose -Message "Adding $f to the array..."
    
        $ransomware.Add([pscustomobject] @{
            Name = $f.Split('\')[-1] #split each value after the slash, get the last value (which is the filename) 
            TimeInfo = $null
            FoundIn = "File System"
            Path = $f
            })
    }
    
## Searching for running processes ##

    $processes = Get-Process

    foreach($p in $processes)
    {
        if($p.Name -like "*$processName*")
        {
            Write-Verbose -Message "Adding $($p.Name) to the array..."
    
            $ransomware.Add([pscustomobject] @{
                Name = $p.Name
                TimeInfo = $p.StartTime
                FoundIn = "Running Processes"
                Path = $p.Path
                })
        }
    }

    # Starting to search for scheduled tasks #

    $tasks = Get-ScheduledTask 
    foreach($t in $tasks)
    {
        if($t.TaskName -like "*$processName*")
        {
            Write-Verbose -Message "Adding $($t.TaskName) to the array..."
    
            $ransomware.Add([pscustomobject] @{
                Name = $t.TaskName
                TimeInfo = Get-ScheduledTaskInfo -TaskName "$($t.TaskName)" | Select-Object LastRunTime
                FoundIn = "Scheduled Tasks"
                Path = $t.TaskPath
                })
        }
    }

    # Registry Keys Start #

    $keys = Get-ChildItem -Path Registry::HKCR #we will get all keys using this, we will loop through later
    $check = (Get-ItemProperty -Path "Registry::HKCR\.$ext" | Select-Object '(default)').'(default)' #this will give us the default value which is commonly used wiht ransowmare
    if($check.Length -gt 1)
    {
        Write-Host "Ransomware entry in registry found" 
        foreach($k in $keys)
        {
            $current = (Get-ItemProperty -Path Registry::HKCR\$($k.PSChildName) | Select-Object "(default)").'(default)'
        
            if($current -eq "$check") #if ransomware value is equal to another key in the registry
            {
                Write-Verbose -Message "Adding $($k.Name) to the array..."

                $ransomware.Add([pscustomobject] @{
                    Name = $k.Name
                    TimeInfo = $null
                    FoundIn = "Registry"
                    Path = $k.PSPath
                })
            }
        }
    }
    else
    {
    Write-Host "No Ransomware entries were found in the regsitry, please double check manually using regedit."
    }

    return $ransomware | Out-GridView -Title "Ransomware Found" #display back to user
}