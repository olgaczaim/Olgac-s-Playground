$user = Get-SPUser -Web [sitename] -Identity "i:0#.w|domain\username"
Move-SPUser -Identity $user -NewAlias "i:0#.w|domain\username" -IgnoreSID
