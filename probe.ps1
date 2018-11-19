# Author : Paul MURITH
# Last update 24/09/18


function SendMail(){
$encodingMail = [System.Text.Encoding]::UTF8
$From = "LocalProbe@domain.tld"
$To = "your.email@domain.tld"
$Subject = "$pc errors"
$Body = "Insert your custom body test : ($nmbApplis):
$nomApplis"
$SMTPServer = "{ServerIP}"
$SMTPPort = "25"
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Encoding $encodingMail
}

function Probe() {
  $ImportData = import-csv "c:\Path\to\csv.csv" | Select PC
  Foreach ($data in $ImportData)
  {
    $pc = $data.PC
    $Filter = 'Microsoft|Intel|Nvidia|Java'
    $Applis = Get-WmiObject -Class Win32_Product -ComputerName $poste -ErrorAction 'SilentlyContinue' | Where {$_.Name -notmatch $Filter} | Select Name
    $nmbApplis = $Applis.count
    $nomApplis = $Applis.Name -join ', '
    if ($nmbApplis -gt 0){SemdMail}
  }
}
Probe
