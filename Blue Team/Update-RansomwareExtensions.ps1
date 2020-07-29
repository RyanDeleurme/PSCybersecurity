

$fileGroup = "RansomwareExtensions"

remove-FsrmFileGroup -Name $fileGroup
Write-Verbose -Message "$fileGroup has been removed."

New-FsrmFileGroup -Name $fileGroup -includepattern @((Invoke-WebRequest -Uri "https://fsrm.experiant.ca/api/v1/get").content | convertfrom-json | ForEach-Object {$_.filters})

$count = get-FsrmFileGroup -name $fileGroup | Select-Object -expandproperty includepattern | Measure-Object | Select-Object -ExpandProperty count

Write-Host -ForegroundColor Green "$count ransomware extensions have been added to the filegroup $fileGroup"