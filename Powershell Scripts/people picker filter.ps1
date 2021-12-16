Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
$wa = Get-SPWebApplication http://..........
$sites = $wa.Sites
foreach($s in $sites)
{
  $paths = stsadm -o getsiteuseraccountdirectorypath -url $s.Url
  $exist = $false;
  foreach($p in $paths)
  {
    if($p.contains("CN=Users,DC=tnb,DC=org"))
    {
        $adr = $s.Url;
        Write-Host "$adr - Already exist!";
        $exist = $true;
        break;
    }
  }
  if($exist -eq $false)
  {
    stsadm -o setsiteuseraccountdirectorypath -path "CN=Users,DC=tnb,DC=org" -url $s.Url
  }
}