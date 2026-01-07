$currentVersion = (Get-Content '.\windexter-cli.csproj' | ? {$_ -match "<Version>" }).TrimStart().Replace("<Version>","").Replace("</Version>","")
$contained = $(ls .\bin\Release\publish\linux-x64\windexter-cli-v$currentVersion).Name
if ($null -ne $contained)
{
    if ($contained -match $currentVersion){
        ((Get-FileHash .\bin\Release\publish\linux-x64\$contained -Algorithm SHA256).Hash).ToLower() | Out-File ..\binaries\$contained.sha256
        Move-Item -Path .\bin\Release\publish\linux-x64\$contained -Destination ..\binaries\$contained -Force
        $pdb = $contained + '.pdb'
		Remove-Item -Path .\bin\Release\publish\linux-x64\$pdb -Force
    }
}

