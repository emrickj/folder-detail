write-host "<!DOCTYPE html>"
write-host "<html lang='en'>"
write-host "<head>"
write-host "  <title>Folder Detail</title>"
write-host "  <meta charset="utf-8">"
write-host "  <meta name='viewport' content='width=device-width, initial-scale=1'>"
write-host "  <link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css'>"
write-host "  <script src='https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js'></script>"
write-host "  <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
write-host "  <script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js'></script>"
write-host "</head>"
write-host "<body>"
$folder = 'C:\users\misue\videos\'
write-host "<h2>Folder: $folder</h2>"
$shell = New-Object -COMObject Shell.Application
$shellfolder = $shell.Namespace($folder)
write-host "<table class='table table-striped'>"
write-host "  <tr>"
For ($i=0; $i -lt 22; $i++) {
  $h = $shellfolder.GetDetailsOf("",$i)
  write-host "    <th>$h</th>"
}
write-host "  </tr>"
foreach ($file in $shellfolder.items())
{
   write-host "  <tr>"
   For ($i=0; $i -lt 22; $i++) {
     $di = $shellfolder.GetDetailsOf($file, $i)
     write-host "    <td>$di</td>"
   }
   write-host "  </tr>"
}
write-host "</table>"
write-host "</body>"
write-host "</html>"
