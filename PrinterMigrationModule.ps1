#System.Printing doesn't remember its own connections without this =P
$host.Runspace.ThreadOptions = "ReuseThread"

#Loads the system.printing namespace
[Reflection.Assembly]::LoadWithPartialName("System.Printing")

#We pass these as access level requestors when we connect to servers and queues.	
$pserveadmin = [System.Printing.PrintSystemDesiredAccess]::AdministrateServer
$pqueueadmin = [System.Printing.PrintSystemDesiredAccess]::AdministratePrinter

#------------------------------------------------------------------------------

function Get-PrintQueue
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourceQueueName) 

	$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
	$SourceQueue = new-object -com "Printer.Printer.1"

	$PrintMaster.PrinterGet("\\" + $SourceServerName, $SourceQueueName, $SourceQueue)

	return $SourceQueue
}

#------------------------------------------------------------------------------

function Add-PrintQueue
{
	param ([string]$DestinationServerName=$DEFAULT_DESTINATIONSERVERNAME, [string]$DestinationQueueName, [string]$DestinationDriverName=$DEFAULT_DESTINATIONDRIVERNAME, [string]$DestinationPortName, [string]$DestinationProcessorName=$DEFAULT_DESTINATIONPROCESSORNAME, [bool]$DestinationDirectoryPublish=$DEFAULT_DESTINATIONDIRECTORYPUBLISH) 
	
	$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
	$DestinationQueue = new-object -com "Printer.Printer.1"

	$DestinationQueue.PrinterName = $DestinationQueueName
	$DestinationQueue.ServerName = "\\" + $DestinationServerName
	$DestinationQueue.PortName = $DestinationPortName
	$DestinationQueue.DriverName = $DestinationDriverName
	$DestinationQueue.PrintProcessor = $DestinationProcessorName
	$DestinationQueue.Published = $DestinationDirectoryPublish
	$PrintMaster.PrinterAdd($DestinationQueue)

	$NewQueue = new-object -com "Printer.Printer.1"
	$PrintMaster.PrinterGet("\\" + $DestinationServerName, $DestinationQueueName, $NewQueue)
	$NewQueue.Shared = $True
	$NewQueue.ShareName = $DestinationQueueName

	$PrintMaster.PrinterSet($NewQueue)
}

#------------------------------------------------------------------------------

function Clone-PrintQueue
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourceQueueName, [string]$DestinationServerName=$DEFAULT_DESTINATIONSERVERNAME, [string]$DestinationDriverName=$DEFAULT_DESTINATIONDRIVERNAME, [string]$DestinationProcessorName=$DEFAULT_DESTINATIONPROCESSORNAME) 

	#Create objects
		$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
		$DestinationQueue = new-object -com "Printer.Printer.1"
		$DestinationPort = new-object -com "Port.Port.1"

	#Get source print queue
		$PrintMaster.PrinterGet("\\" + $SourceServerName, $SourceQueueName, $DestinationQueue)

	#Clone source print queue IP printer port to destination server
		$PrintMaster.PortGet("\\" + $SourceServerName, $DestinationQueue.PortName, $DestinationPort)
		$DestinationPort.ServerName = "\\" + $DestinationServerName
		$PrintMaster.PortAdd($DestinationPort)

	#Clone source print queue to destination server, optionally changing the driver and/or print processor
		$DestinationQueue.PrinterName = $SourceQueueName
		$DestinationQueue.ServerName = "\\" + $DestinationServerName
		if ( $DestinationDriverName.Length -gt 0 ){
			$DestinationQueue.DriverName = $DestinationDriverName}
		if ( $DestinationProcessorName.Length -gt 0 ){
			$DestinationQueue.PrintProcessor = $DestinationProcessorName}
		$DestinationQueue.Published = $DestinationDirectoryPublish
		$PrintMaster.PrinterAdd($DestinationQueue)

	#Share destination print queue
		$PrintMaster.PrinterGet("\\" + $DestinationServerName, $SourceQueueName, $DestinationQueue)
		$DestinationQueue.Shared = $True
		$DestinationQueue.ShareName = $SourceQueueName
		$PrintMaster.PrinterSet($DestinationQueue)
}

#------------------------------------------------------------------------------

function Get-PrintPort
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourcePortName)

    	$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
	$SourcePort = new-object -com "Port.Port.1"
	
	$PrintMaster.PortGet("\\" + $SourceServerName, $SourcePortName, $SourcePort)
	
	return $SourcePort
}

#------------------------------------------------------------------------------

function Add-PrintPort
{
	param ([string]$DestinationServerName=$DEFAULT_DESTINATIONSERVERNAME, [string]$DestinationPortName, [string]$DestinationPortIPAddress)

	$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
	$DestinationPort = new-object -com "Port.Port.1"

	$DestinationPort.ServerName = "\\" + $DestinationServerName
	$DestinationPort.PortName = $DestinationPortName
	$DestinationPort.HostAddress = $DestinationPortIPAddress

	$PrintMaster.PortAdd($DestinationPort)
}

#------------------------------------------------------------------------------

function Clone-PrintPort
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourcePortName, [string]$DestinationServerName=$DEFAULT_DESTINATIONSERVERNAME) 

	$PrintMaster = new-object -com "PrintMaster.PrintMaster.1"
	$DestinationPort = new-object -com "Port.Port.1"

	$PrintMaster.PortGet("\\" + $SourceServerName, $SourcePortName, $DestinationPort)
	$DestinationPort.ServerName = "\\" + $DestinationServerName
	
	$PrintMaster.PortAdd($DestinationPort)
}

#------------------------------------------------------------------------------

function Export-PrintDriverCapabilities
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourceQueueName, [string]$ExportPath=$DEFAULT_EXPORTPATH)

	$SourceServer = new-object system.printing.printserver("\\" + $SourceServerName)
	$SourceQueue = $SourceServer.GetPrintQueue($SourceQueueName)

	$SourceCapabilities = $SourceQueue.GetPrintCapabilitiesAsXml()
	$ExportPathFull = $ExportPath + "\PrintDriverCapabilities_" + (Convert-DriverNameToValidFilename($SourceQueue.QueueDriver.Name)) + ".xml"
	$SourceCapabilitiesExport = new-object System.IO.FileStream("$ExportPathFull", [IO.FileMode]::Create)
	$SourceCapabilities.WriteTo($SourceCapabilitiesExport)
	$SourceCapabilitiesExport.flush()
	$SourceCapabilitiesExport.close()
}

#------------------------------------------------------------------------------

function Export-PrintQueueDefaultTicket
{
	param ([string]$SourceServerName=$DEFAULT_SOURCESERVERNAME, [string]$SourceQueueName, [string]$ExportPath=$DEFAULT_EXPORTPATH)

	$SourceServer = new-object system.printing.printserver("\\" + $SourceServerName)
	$SourceQueue = $SourceServer.GetPrintQueue($SourceQueueName)
	$SourceDefaultPrintTicket = $SourceQueue.DefaultPrintTicket.GetXMLStream()

	$ExportPathFull = $ExportPath + "\PrintQueueDefaultTicket_" + $SourceServerName + "_" + (Convert-DriverNameToValidFilename($SourceQueueName)) + ".xml"

	$SourceDefaultPrintTicketExport = new-object System.IO.FileStream("$ExportPathFull", [IO.FileMode]::Create)
	$SourceDefaultPrintTicket.WriteTo($SourceDefaultPrintTicketExport)
	$SourceDefaultPrintTicketExport.flush()
	$SourceDefaultPrintTicketExport.close()
}

#------------------------------------------------------------------------------

function Convert-DriverNameToValidFilename
{
	param ([string]$DriverName)

	$ValidFilename = $DriverName
	$ValidFilename = $ValidFilename.Replace("/","[SLS]")
	$ValidFilename = $ValidFilename.Replace("\","[BSL]")
	$ValidFilename = $ValidFilename.Replace(":","[COL]")
	$ValidFilename = $ValidFilename.Replace("*","[AST]")
	$ValidFilename = $ValidFilename.Replace("?","[QST]")
	$ValidFilename = $ValidFilename.Replace("`"","[QOT]")
	$ValidFilename = $ValidFilename.Replace("<","[LST]")
	$ValidFilename = $ValidFilename.Replace(">","[GRT]")
	Return $ValidFilename
}

function Convert-ValidFilenameToDriverName
{
	param ([string]$ValidFilename)

	$DriverName = $ValidFilename
	$DriverName = $DriverName.Replace("[SLS]","/")
	$DriverName = $DriverName.Replace("[BSL]","\")
	$DriverName = $DriverName.Replace("[COL]",":")
	$DriverName = $DriverName.Replace("[AST]","*")
	$DriverName = $DriverName.Replace("[QST]","?")
	$DriverName = $DriverName.Replace("[QOT]","`"")
	$DriverName = $DriverName.Replace("[LST]","<")
	$DriverName = $DriverName.Replace("[GRT]",">")
	Return $DriverName
}