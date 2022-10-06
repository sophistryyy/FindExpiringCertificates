#findexpcerts.ps1
#Looks at all certs and sends an alert when a cert is a certain amount of days away from expiry
#Checks within a 7 day window so must run weekly to detect all milestones

.$PSScriptRoot\CertParser.ps1 #import helper script

#how many days away to send an alert
$m1 = 0 #milestone 1
$m2 = 14 #milestone 2
$m3 = 60 #milestone 3

[string[]]$emailReceivers = "<email address>", "<email address>" #string array of email recipients

#variable initializaitons
$certs = @()
$num = 0
$array1 = New-Object string[] 100
$array2 = New-Object string[] 100
$array3 = New-Object string[] 100
$i1 = 0
$i2 = 0
$i3 = 0

#format alert tables if there are any certs expiring in that timeframe
#then call send email
function make-alert{ 
    if($i1 -ne 0){
        $formated1 = "
            <h4>Expiring this Week:</h4>
            <style>
                table,
                th,
                td {
                padding: 10px;
                border: 1px solid black;
                border-collapse: collapse;
                }
            </style>
            <table>
                <tr>
                    <th>Expiry Date</th>
                    <th>Request ID</th>
                    <th>Common Name</th>
                    <th>Requester</th>
                    <th>Template</th>
                </tr>" + ($array1 -join "") + "</table></br>"
    }
    if($i2 -ne 0){
        $formated2 = "
            <h4>Expiring in next $($m2) Days:</h4>
            <table>
                <tr>
                    <th>Expiry Date</th>
                    <th>Request ID</th>
                    <th>Common Name</th>
                    <th>Requester</th>
                    <th>Template</th>
                </tr>" + ($array2 -join "") + "</table></br>"
    }
    if($i3 -ne 0){
        $formated3 = "
            <h4>Expiring in next $($m3) Days:</h4>
            <table>
                <tr>
                    <th>Expiry Date</th>
                    <th>Request ID</th>
                    <th>Common Name</th>
                    <th>Requester</th>
                    <th>Template</th>
                </tr>" + ($array3 -join "") + "</table></br>"
    }
    send-email
}


#send email listing all certs that are about to expire
function send-email{
    $MailMessage = @{
        To = $emailReceivers
        From = "Certificate Expiry <certificates@domain.ca>"
        Subject = "Certificates Expiring Soon"
        Body = "<h1>Certificates Expiring</h1>
        $($formated1)
        $($formated2)
        $($formated3)"
        Smtpserver = "<mailserver>"
        Port = 25
        BodyAsHtml = $true
    }
    Send-MailMessage @MailMessage
}

###########MAIN###########

$certs = make-certs $m3 #calling function from helper script to make certificate array

#iterate through each cert, checking if cert is at a checkpoint (75, 200)
foreach($c in $certs){
    if($c.notafter -ne $null){
        #group expiring the week of $m1 days
        #check if alert date less than/equal to current date+7, AND check if alert date is after/equal to current date
        if(($c.notafter.adddays(-$m1) -le (get-date).adddays(7)) -and ($c.notafter.adddays(-$m1) -ge (get-date)) ){ 
            $array1[$i1] = "
                <tr>
                    <th>$($c.notafter)</th>
                    <th>$($c.requestid)</th>
                    <th>$($c.commonname)</th>
                    <th>$($c.requestername)</th>
                    <th>$($c.certificatetemplate)</th>
                </tr>"
            $i1 = $i1 + 1
            $num++
        }

        #group expiring in $m2 days
        #check if expiry date is happens within a week of 60 days from now (in other words 60-day alert date is within next 7 days)
        if(($c.notafter.adddays(-$m2) -le (get-date).adddays(7)) -and ($c.notafter.adddays(-$m2) -ge (get-date)) ){ 
            $array2[$i2] = "
                <tr>
                    <th>$($c.notafter)</th>
                    <th>$($c.requestid)</th>
                    <th>$($c.commonname)</th>
                    <th>$($c.requestername)</th>
                    <th>$($c.certificatetemplate)</th>
                </tr>"
            $i2 = $i2 + 1
            $num++
        }

        #group expiring in max $m3 days
        if(($c.notafter.adddays(-$m3) -le (get-date).adddays(7)) -and ($c.notafter.adddays(-$m3) -ge (get-date)) ){
            $array3[$i3] = "
                <tr>
                    <th>$($c.notafter)</th>
                    <th>$($c.requestid)</th>
                    <th>$($c.commonname)</th>
                    <th>$($c.requestername)</th>
                    <th>$($c.certificatetemplate)</th>
                </tr>"
            $i3 = $i3 + 1
            $num++
        }
    }
}
#if any cert matches above criteria then send alert
if($num -ne 0){
    make-alert
}