write-host "<!DOCTYPE html>"
write-host "<html lang='en'>"
write-host "<head>"
write-host "  <title>Folder Detail</title>"
write-host "  <meta charset="utf-8">"
write-host "  <meta name='viewport' content='width=device-width, initial-scale=1'>"
write-host "  <link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css'>"
write-host "  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>"
write-host "  <script src='https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js'></script>"
write-host "  <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
write-host "  <script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js'></script>"
write-host "</head>"
write-host "<body>"
write-host "<div class='container'>"
$folder = 'C:\users\misue\downloads\'
write-host "<h2><i class='fa fa-folder-open-o'></i> $folder</h2>"
$shell = New-Object -COMObject Shell.Application
$shellfolder = $shell.Namespace($folder)
write-host "<div class='table-responsive'>"
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
   $name = $shellfolder.GetDetailsOf($file, 0)
   $type = $shellfolder.GetDetailsOf($file, 2)
   $icon = ""
   switch($type){
      "File folder"               {$icon = "fa fa-folder-o"}
	  "MP4 File"                  {$icon = "fa-file-video-o"}
	  "M4V File"                  {$icon = "fa-file-video-o"}
	  "JPG File"                  {$icon = "fa-file-image-o"}
	  "PNG File"                  {$icon = "fa-file-image-o"}
	  "PDF File"                  {$icon = "fa-file-pdf-o"}
	  "PHP File"                  {$icon = "fa-file-code-o"}
	  "Windows Powershell Script" {$icon = "fa-file-code-o"}
	  "Microsoft Word Document"   {$icon = "fa-file-word-o"}
	  "File"                      {$icon = "fa-file-o"}
   }
   write-host "    <td><i class='fa $icon'></i> $name</td>"
   For ($i=1; $i -lt 22; $i++) {
     $di = $shellfolder.GetDetailsOf($file, $i)
     write-host "    <td>$di</td>"
   }
   write-host "  </tr>"
}
write-host "</table>"
write-host "</div>"
write-host "</div>"
write-host "</body>"
write-host "</html>"
