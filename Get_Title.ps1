$folder = 'C:\users\misue\videos\'
$shell = New-Object -COMObject Shell.Application
$shellfolder = $shell.Namespace($folder)
$h1 = $shellfolder.GetDetailsOf("",0)
$h2 = $shellfolder.GetDetailsOf("",21)
write-host $h1.PadRight(20," "),$h2.PadRight(20," ")
write-host "--------------------------------------------------"
foreach ($file in $shellfolder.items())
{
   $name = $shellfolder.GetDetailsOf($file, 0)
   $title = $shellfolder.GetDetailsOf($file, 21)
   write-host $name.PadRight(20," "),$title.PadRight(20," ")
}
