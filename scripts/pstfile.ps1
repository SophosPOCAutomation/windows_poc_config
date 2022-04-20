$PSTPath = "C:\Users\Sophos\Documents\Sophos.pst"
Add-type -assembly "Microsoft.Office.Interop.Outlook"
$outlook = new-object -comobject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
Get-ChildItem $PSTPath -Filter *.pst |
ForEach-Object {
  $namespace.AddStore($_.FullName)
}
