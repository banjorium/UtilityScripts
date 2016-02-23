#
# Purge-MailboxFolder.ps1
#
# By David Barrett, Microsoft Ltd. 2013. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

# Parameters

[CmdletBinding()]
param (
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,
	
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Specifies the folder to be purged")]
	[ValidateNotNullOrEmpty()]
	[string]$Folder,

	[Parameter(Position=2,Mandatory=$False,HelpMessage="Specifies the date before which items will be deleted")]
	[DateTime]$PurgeBeforeDate,
	
	[switch]$SearchForFolder,
	[string]$Username,
	[string]$Password,
	[string]$Domain,
	[switch]$Impersonate,
	[string]$EwsUrl,
	[switch]$IgnoreSSLCertificate,
	[string]$EWSManagedApiPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll",
    [switch]$ReportOnly
)

Function ShowParams()
{
	Write-Host "Purge-MailboxFolder -Mailbox <string>"
	Write-Host "                    -Folder <string>"
	Write-Host "                    -PurgeBeforeDate <DateTime>"
	Write-Host "                    [-SearchForFolder]"
	Write-Host "                    [-Username <string> -Password <string> [-Domain <string>]]"
	Write-Host "                    [-Impersonate]"
	Write-Host "                    [-EwsUrl <string>]"
	Write-Host "                    [-IgnoreSSLCertificate]"
	Write-Host "                    [-EWSManagedApiPath <string>]"
	Write-Host "";
	Write-Host "Required:"
	Write-Host " -Mailbox : Mailbox SMTP email address (OR filename for source list of mailboxes; text file, one mailbox per line)"
	Write-Host " -Folder : Folder to be purged (items within it deleted).  Full path must be specified if SearchForFolder parameter is missing."
	Write-Host " -PurgeBeforeDate : The date before which items will be deleted from the folder"
	Write-Host ""
	Write-Host "Optional:"
	Write-Host " -SearchForFolder : If present, the mailbox will be searched for until the named folder is found"
	Write-Host " -Username : Username for the account being used to connect to EWS (if not specified, current user is assumed)"
	Write-Host " -Password : Password for the specified user (required if username specified)"
	Write-Host " -Domain : If specified, used for authentication (not required even if username specified)"
	Write-Host " -Impersonate : If present, impersonation will be used"
	Write-Host " -EwsUrl : Forces a particular EWS URl (otherwise autodiscover is used, which is recommended)"
	Write-Host " -IgnoreSSLCertificate : If present, any SSL errors will be ignored"
	Write-Host " -EWSManagedApiPath : Filename and path to the DLL for EWS Managed API (if not specified, default path for v2.0 is used)"
	Write-Host ""
}



# Define our functions

Function FindFolder()
{
	# Search for a folder of the given name
	
	$RootFolder, $FolderName = $args[0]
	return RecurseFolder($RootFolder, $FolderName)
}

Function RecurseFolder()
{
	$folder, $FolderName = $args[0]

	$view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(500)
	ForEach ($subFolder in $folder.FindFolders($view))
	{
		if ( $subFolder.DisplayName -eq $FolderName )
		{
			return $subFolder
			break
		}
		RecurseFolder($subFolder, $FolderName)
	}	
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath = $args[0]
	
	$Folder = $RootFolder
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View)
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null
					Write-Host "Failed to find" $PathElements[$i] -ForegroundColor Red
					Write-Host "Requested folder path:" $FolderPath -ForegroundColor Yellow
					break
				}
				
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id)
			}
		}
	}
	
	return $Folder
}


Function PurgeFolder()
{
	# Purge the folder
	$Folder = $args[0]
	Write-Host "Purging folder" $Folder.DisplayName
	
	$itemsToDelete = @{}
	
	$pageSize = 500 # We will get details for up to 500 items at a time
	$offset = 0
	$moreItems = $true
	
	if ( $PurgeBeforeDate)
	{
		$filter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $PurgeBeforeDate)
		$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
		$searchFilter.Add( $filter1 )
	}
			
	# Process the items in the folder and determine which should be deleted
	$i=0
	while ($moreItems)
	{
		$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
		$view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
		$view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
		
		if ( $searchFilter )
		{
			$results = $service.FindItems( $Folder.Id, $searchFilter, $view )
		}
		else
		{
			$results = $service.FindItems( $Folder.Id, $view )
		}
		
		ForEach ($item in $results.Items)
		{
			$itemsToDelete.Add($i++, $item)
		}
		Write-Verbose ($results.Items.Count.ToString() + " items added to delete list")
		
		$moreItems = $results.MoreAvailable
		$offset += $pageSize
	}
	
	# Now delete the items (we will do this in batches of 500
	$itemId = New-Object Microsoft.Exchange.WebServices.Data.ItemId("xx")
	$itemIdType = [Type] $itemId.GetType()
	$baseList = [System.Collections.Generic.List``1]
	$genericItemIdList = $baseList.MakeGenericType(@($itemIdType))
	$deleteIds = [Activator]::CreateInstance($genericItemIdList)
	ForEach ($item in $itemsToDelete.Values)
	{
		$deleteIds.Add($item.Id)
		if ($deleteIds.Count -ge 500)
		{
			# Send the delete request
			[void]$service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
			$deleteIds = [Activator]::CreateInstance($genericItemIdList)
		}
	}
	if ($deleteIds.Count -gt 0)
	{
		[void]$service.DeleteItems( $deleteIds, [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone, $null )
	}
	Write-Host $deleteIds.Count "items deleted"
}


Function ProcessMailbox()
{
    # This function does all the work for a particular mailbox
    
    # Set EWS URL if specified, or use autodiscover if no URL specified.
    if ($EwsUrl)
    {
    	$service.URL = New-Object Uri($EwsUrl)
    }
    else
    {
    	try
    	{
    		Write-Host "Performing autodiscover for $Mailbox"
    		$service.AutodiscoverUrl($Mailbox, {$True})
            Write-Host "EWS Url is:", $service.Url
    	}
    	catch
    	{
    		throw
    	}
    }
 
    # Set impersonation if specified
    if ($Impersonate)
    {
    	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox)
    }
    
	$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
	if ( $RootFolder -eq $null )
	{
		Write-Host "FAILED to bind to root folder for mailbox $Mailbox" -foregroundcolor Red
		return
	}
	
	if ($SearchForFolder)
	{
		$Folder = FindFolder( $RootFolder, $Folder )
	}
	Else
	{
		$Folder = GetFolder( $RootFolder, $Folder )
	}
	if ( $Folder -eq $null ) { return }

    # Purge this folder
	PurgeFolder( $Folder )
}

Function SearchDll()
{
	# Search for a program/library within Program Files (x64 and x86)
	$path = $args[0]
	$programDir = $Env:ProgramFiles
	if (Get-Item -Path ($programDir + $path) -ErrorAction SilentlyContinue)
	{
		return $programDir + $path
	}
	
	$programDir = [environment]::GetEnvironmentVariable("ProgramFiles(x86)")
	if ( [string]::IsNullOrEmpty($programDir) ) { return "" }
	
	if (Get-Item -Path ($programDir + $path) -ErrorAction SilentlyContinue)
	{
		return $programDir + $path
	}
}

Function LoadEWSManagedAPI()
{
	# Check EWS Managed API available
	
	if ( !(Get-Item -Path $EWSManagedApiPath -ErrorAction SilentlyContinue) )
	{
		$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll")
		if ( [string]::IsNullOrEmpty($EWSManagedApiPath) )
		{
			$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll")
			if ( [string]::IsNullOrEmpty($EWSManagedApiPath) )
			{
				$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll")
			}
		}
	}
	
	If ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		# Load EWS Managed API
		Write-Host "Using managed API found at:" $EWSManagedApiPath -ForegroundColor Gray
		Add-Type -Path $EWSManagedApiPath
		return $true
	}
	return $false
}


# The following is the main script


if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}
 
# If we are ignoring any SSL errors, set up a callback
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
}

# Create Service Object.  We use Exchange 2010 schema
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)

# Set credentials if specified, or use logged on user.
 if ($Username -and $Password)
 {
     if ($Domain)
     {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)
     } else {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password)
     }
     
} else {
     $service.UseDefaultCredentials = $true
}

if ( !(Get-Item -Path $Mailbox -ErrorAction SilentlyContinue) )
{
    # No file referenced, so mailbox must be single mailbox
    ProcessMailbox;
}
else
{
    # $Mailbox references a file, so read it and process all the mailboxes specified
    $Mailboxes = Get-Content $Mailbox
    foreach ( $SMTP in $Mailboxes )
    {
        if ( $SMTP -ne $null )
        {
            if ( $SMTP -ne "" )
            {
                $Mailbox = $SMTP
                ProcessMailbox
            }
        }
    }
}