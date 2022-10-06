# FindExpCerts_v2.0
Sept. 26, 2022

## About

*FindExpCerts runs on the certificate authority server weekly and sends an email detailling which certificates are about to expire. 
*It uses certutil.exe (a Windows command line tool that is installed as part of Certificate Services) to query the certificate database.

## Usage

*FindExpCerts.ps1 is scheduled to run with Windows Task Scheduler and calls CertParser.ps1 to parse the output from certutil before it formats and sends an email.
*The receiviers of the email are configured in the $emailReceivers string array in FindExpCerts.ps1. They are formatted as strings separated by commas.
*The script can exclude querying certain certificate templates by putting the template number in Exclusions.txt, each on a new line.
*For example, to exclude the Workstation Authentication template, the template number 1.3.6.1.4.1.311.21.8.13499496.12876343.14502321.11943615.13984618.174.6199101.11143143 is entered into Exlusions.txt.

## Content

*FindExpCerts.ps1
*CertParser.ps1
*Exclusions.txt

## Requirements

*This script requires no PowerShell modules
*This script requires certutil
*This script must run on a certificate authority server.

## Troubleshooting

*To view all available templates use "certutil -catemplates -v | select-string displayname". To see the corresponding template numbers use "certutil -catemplates -v | select-string displayname, msPKI-Cert-Template-OID"

