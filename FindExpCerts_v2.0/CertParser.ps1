#CertParser.ps1
#parses string output from certutil

#takes int representing how many days ahead to look and returns array of cert objects
function make-certs([int]$m3){
    #variable initializations
    $templates = @() #list of templates to query
    $certs = @() 

    #custom certificate object
    $certificate = @{
        CommonName = ""
        RequestID = ""
        RequesterName = ""
        RequestType = ""
        CertificateTemplate = ""
        NotAfter = [DateTime]"1/1/1900" #default
    }

    #template object ID (OID) parsing
    $templateDump = certutil -v -template #get all templates existing on the CA
    $i = 0
    $rawtemplates = @(
        foreach($line in $templateDump){
            if($line -like "*TemplatePropOID =*"){
                (($templateDump[$i + 1]) -split " ")[4]
            }
        $i++
        }
    )

    #filter templates by OIDs we want to exclude in Exclusions.ps1
    foreach($t in $rawtemplates){
        $isExclusion = select-string -path $PSScriptRoot\Exclusions.txt -pattern $t
        if($isExclusion -eq $null){
            $templates += $t
        }
    }

    #querying certificates
    foreach($template in $templates){
        $rows = @()
        #get certificate output
        $output = certutil -view -restrict "Certificate Template = $template, Certificate Expiration Date <= $((get-date).addDays($m3+7)), Certificate Expiration Date >= $(get-date)" -out "CommonName, NotAfter, RequestID, CertificateTemplate, RequesterName, RequestType"
        
        #split output into an array of rows
        $output = [string]$output
        $sep = @("Row [1-9]*:")
        $rows = $output -split $sep

        #process each row separately
        foreach ($r in $rows){
            $name = "" #common name
            $type = "" #request type
            $id = "" #request id
            $requester = "" #requester
            $date = "" #expiration date
            $temp = "" #template

            #parse output
            #common name
            $name = $r | select-string "Issued Common Name:"
            $name -match 'Issued Common Name: (.*)?Certificate Expiration'
            if($name){
                $nameparsed = $Matches[1]
                write-host $nameparsed
            }
            #request id
            $id = $r | select-string "Issued Request ID:"
            $id -match 'Issued Request ID: (.*)?Certificate Template:'
            if($id){
                $idparsed = $Matches[1]
                write-host $idparsed
            }
            #requester name
            $requester = $r | select-string "Requester Name:"
            $requester -match 'Requester Name: (.*)?Request Type:'
            if($requester){
                $requesterparsed = $Matches[1]
                write-host $requesterparsed
            }
            #request type
            #$type = $r | select-string "Request Type:"
            #$type -match 'Request Type: (.*)?Maximum Row'
            #if($type){
            #    $typeparsed = $Matches[1]
            #    write-host $typeparsed
            #}
            #certificate template
            $temp = $r | select-string "Certificate Template:"
            $temp -match 'Certificate Template: (.*)?Requester Name:'
            if($temp){
                $templateparsed = $Matches[1]
                write-host $templateparsed
            }
            #expiration date
            $date = $r | select-string "Certificate Expiration Date:"
            $date -match 'Certificate Expiration Date: (.* .*)?Issued Request ID'
            if($date){
                $dateparsed = $Matches[1]
                write-host $dateparsed
            }
            $newcert = new-object -typename psobject -property $certificate
            $newcert.CommonName = $nameparsed
            $newcert.RequestID = $idparsed
            $newcert.RequesterName = $requesterparsed
            #$newcert.RequestType = $typeparsed
            $newcert.CertificateTemplate = $templateparsed
            
            #add certificate to cert array if it has a valid expiration date
            if($date){
                write-host "************************"
                $newcert.NotAfter =[DateTime]$dateparsed #turn date string into datetime object
                $certs += $newcert
            }
            
        }
    }
    return $certs
}
