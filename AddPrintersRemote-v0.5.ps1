# ************************************************************************************
#
# Script Name   : AddPrintersRemote.ps1
# Version       : 0.4
# Author        : A.R. Winters
# Date          : 02 november 2009
# Final Version : 
#
# Description   : Script kan remote printers installeren, de printers 
#                 worden uitgelezen middels een csv bestand. Het script
#                 is vrij makkelijk aan te passen.
#
# Comments      : De Win32_Print WMI provider code om printers mee aan
#                 te maken middels PowerShell is in dit script opgenomen.
#                 de provider werkt nog niet wegens de provider in Powershell
#                 de functie niet ondersteund (misschien werkt het wel in W2k8. 
#                 Er is geen error handling in het script gebouwd, in het geval 
#                 van geen csv file en -of onjuiste gegevens wordt het script afgebroken.
#
# ************************************************************************************


# Application Variables
$Domain="coevorden.intern"
$DNSServer = "dom001.w2k3dom.lan"
$PrintServer="fes"
$CSVFile="ImportPrinters.csv"


# Functions

# Function CreatePort
function CreatePort
{
  Write-Host "Add new printer ports on server: $PrintServer... Please Wait" -BackgroundColor Black -ForegroundColor Yellow
  Write-Host ""
    Import-Csv $CSVFile | %{foreach($PrinterPort in $_.PrinterPort)
	            {
				   Write-Host "[action]      - Adding Printer Port: $PrinterPort" -BackgroundColor DarkGreen
			       $port = ([WMICLASS]"\\$PrintServer\ROOT\cimv2:Win32_TCPIPPrinterPort").createInstance() 
                   $port.Name="$PrinterPort.$Domain" 
                   $port.SNMPEnabled=$false 
                   $port.Protocol=1 
                   $port.HostAddress="$PrinterPort.$Domain"  
                   $port.Put() 
                   $port 
				   Write-Host "[information] - Printer Port: $PrinterPort, added successfull" -BackgroundColor DarkCyan
			    }} > NULL
  Write-Host ""
  Write-Host "Creating Printer Ports. Done!" -BackgroundColor Black -ForegroundColor Yellow
}

# Function AddPrinters
function AddPrinters
{
  Write-Host ""
  Write-Host "Creating printers and link the ports on printserver: $PrintServer, Please Wait" -BackgroundColor Black -ForegroundColor Yellow
  Write-Host ""
  
  $csv=@()
  $csv += Import-csv $CSVFile
     foreach($line in $csv)
       { 
 		  
					 
		  #$printdriver = '"PCL6 Driver for Universal Print"'	# Mogelijk om dit in de CSV file op te nemen als bijvoorbeeld: $printdriver = $line.Driver.
		  $printdriver = '"Generic / Text Only"'
		  $Printer = $line.Printer
		  $Port = $line.PrinterPort
		  		  
		  # Wegens de Win32_Printer API nog niet via 
		  # Powershell kan worden aangesproken, is de
		  # wordt dit nu met een omweg uitgevoerd in
		  # een dos box.
		  Write-Host "[action]      - Adding Printer: $Printer to $Port" -BackgroundColor DarkGreen 		  
		  $VBSCRIPT="cscript prnmgr.vbs -a -b "+$Printer+" -c "+""+"\\"+$PrintServer+" -m "+""+$Printdriver+" -r "+"$Port.$Domain"		  
		  $Addprinters = "$VBSCRIPT"		  
          cmd /c $Addprinters 		  
			
		  # Code onderaan kan gebruikt worden als de Win32_Printer API 
		  # volledig is aan te spreken met PowerShell. Voor nu maakt de 
		  # code de printer niet aan maar toont de gevulde properties 
		  # alleen op het scherm!
		  
		  #$print = ([WMICLASS]"\\$PrintServer\ROOT\cimv2:Win32_Printer").createInstance() 
		  ##$print = ([WMICLASS]"\\$PrintServer\root\cimv2:Win32_Printer").AddPrinterConnection()
		  #$print.Servername=$PrintServer	
          #$print.PrinterName=$line.Printer
          #$print.Drivername=$printdriver 
          #$print.Portname=$line.PrinterPort
          #$print.Shared=$true 
          #$print.Sharename=$line.Printer
          #$print.Published=$true
          #$print.Location="test locatie" 
          #$print.Comment="No comment"    				  
 		  #$print.Put		 
		  #$print
       }	   
           Write-Host "Creating printers and linking the ports on Print Server: $PrintServer.... done!" -BackgroundColor Black -ForegroundColor Yellow	   
		   Write-Host "Starting creation of A \ PTR DNS records on DNS Server: $DNSServer. Please Wait." -BackgroundColor Black -ForegroundColor Yellow	   
}

function CreateDNS
{
  Write-Host ""
  Write-Host "Creating printer DNS records server: $DNSServer, Please Wait" -BackgroundColor Black -ForegroundColor Yellow
  Write-Host ""
  
  $csv=@()
  $csv += Import-csv $CSVFile
  
  # Variable needed for the WMI DNS Provider
  $DNSClassA = [wmiclass]"\\$DNSserver\root\MicrosoftDNS:MicrosoftDNS_AType"
  $DNSClassPTR = [wmiclass]"\\$DNSserver\root\MicrosoftDNS:MicrosoftDNS_PTRType"
  
	 foreach($line in $csv)
       {  
		  $Printer = $line.Printer
		  $PrinterIP = $line.IP
		  $PrinterFQDN = $Printer+".$domain"
		  $Port = $line.PrinterPort+".$domain"
          #$Zone = "$domain"
		  $Zone = "192.168.200.x Subnet"
          $class = 1
          $TTL = 3600
		  $IPAddress = "$PrinterIP"	
		  
		     # Maak DNS A Record
             #$DNSClassA.CreateInstanceFromPropertyData($DNSserver, $zone, $Port, $class, $ttl, $IPAddress)
			 $DNSClassA.CreateInstanceFromPropertyData($DNSserver, $zone, $Port, $class, $ttl, $IPAddress)
			 
			 # Naar overleg met Raimond kan dit worden vergeten.
			 # In de toekomst kan dit alsnog worden gebruikt, waarschijnlijk moet de code iets worden aangepast.
			 
			 # Maak DNS PTR Record
			 #$DNSClassPTR.CreateInstanceFromPropertyData($DNSserver, $zone, $Printer, $class, $ttl,$PrinterFQDN)
        }
	   
   Write-Host ""	   
   Write-Host "Creating DNS A \ PTR records for printer ports on DNS server: $DNSServer.... done!" -BackgroundColor Black -ForegroundColor Yellow
   Write-Host "Assumption is the mother of all screw up's, so plese check the dns and printserver te be sure all went ok!" -BackgroundColor Black -ForegroundColor Yellow
}

function CreatePTRDNS
{
  Write-Host ""
  Write-Host "Creating printer DNS records server: $DNSServer, Please Wait" -BackgroundColor Black -ForegroundColor Yellow
  Write-Host ""
  
  $csv=@()
  $csv += Import-csv $CSVFile
  
  # Variable needed for the WMI DNS Provider
  $DNSClassPTR = [wmiclass]"\\$DNSserver\root\MicrosoftDNS:MicrosoftDNS_ResourceRecord"
  
	 foreach($line in $csv)
       {  
		  $Printer = $line.Printer
		  $PrinterIP = $line.IP
		  $PrinterFQDN = $Printer+".$domain"
		  $Port = $line.PrinterPort+".$domain"
          #$Zone = "$domain"
		  $Zone = "192.168.200.x Subnet"
          $class = 1
          $TTL = 3600
		  $IPAddress = "$PrinterIP"		
		  
			#create our ip address variable
			$raddress = $line.IP

			#Get the name record
			$rname = $line.Printer

			#break the address into octets
			$breakaddress = $raddress.split(‘.’)

			#create octets
			$rFirst = $breakaddress[0];  
			$rSecond = $breakaddress[1] ;
			$rThird = $breakaddress[2]  ;
            $rFourth = $breakaddress[3]

			#create the Reverse lookup String
			$strReverseRR = “$rFourth”+”.”+”$rThird”+”.”+”$rSecond”+” IN PTR $rname.$Domain”
			$strReverseDomain = “Anthony.$rFirst”+”.in-addr.arpa.”

			#Call Create Method
			#$objRR.CreateInstanceFromTextRepresentation($DNSserver,$strReverseDomain,$strReverseRR)
			
			Write-Host $strReverseRR
			#Write-Host $strReverseDomain
			 			 
			 # Maak DNS PTR Record			 
		    # $DNSClassPTR.CreateInstanceFromTextRepresentation($DNSserver,$strReverseDomain,$strReverseRR)	 
        }
	   
   Write-Host ""	   
   Write-Host "Creating DNS A \ PTR records for printer ports on DNS server: $DNSServer.... done!" -BackgroundColor Black -ForegroundColor Yellow
   Write-Host "Assumption is the mother of all screw up's, so plese check the dns and printserver te be sure all went ok!" -BackgroundColor Black -ForegroundColor Yellow
}
	
# MainLoop
CreatePort
AddPrinters
#CreateDNS
#CreatePTRDNS