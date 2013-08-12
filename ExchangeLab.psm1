Add-Type -AssemblyName System.Web
$users = import-csv $psScriptRoot\users.csv

Function New-LabMailbox {
<#
	.SYNOPSIS
		This function creates mailboxes intended for lab use.

	.DESCRIPTION
		This function generates random user names for mailboxes using a CSV file.
		The CSV file used to provide the names contains 1000 unique, commonly used,
		first and last names provided by the US Census Bureau.

	.PARAMETER  Count
		Specifies the number of mailboxes to create. The default value is 1.

	.PARAMETER  Password
		The password for the Active Directory account. If not provided, a random
		strong password will be generated automatically.

	.PARAMETER  Database
		Specifies the database used for the mailbox. This parameter is required 
		when used with Exchange 2007 SP2. This paramter is optional on Exchange 2010.
		
	.PARAMETER  UPNSuffix
		The UPN suffix for the Active Directory user. If not provided, the Active
		Directory root domain name will be used.
		
	.PARAMETER  OrganizationalUnit
		Specifies the OU for the Active Directory account. If no value is provided,
		the account will be created in the default users container.

	.EXAMPLE
		New-LabMailbox -Count 5
		
		Description
		-----------
		Creates 5 mailboxes using random user names.		

	.EXAMPLE
		Get-MailboxDatabase | %{ New-LabMailbox -Count 10 -Database $_.Name }
		
		Description
		-----------
		Creates 10 mailboxes in every Exchange database.		
		
	.NOTES
		Author: Mike Pfeiffer
		Blog  : http://www.mikepfeiffer.net/		

#>
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$false)]
		[System.Int32]
		$count = 1,

		[Parameter(Position=1, Mandatory=$false)]
		[System.String]
		$password = [System.Web.Security.Membership]::GeneratePassword(10,2),
		
		[Parameter(Position=2, Mandatory=$false)]
		[System.String]
		$database,

		[Parameter(Position=3, Mandatory=$false)]
		[System.String]
		$UpnSuffix = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().RootDomain.Name,

		[Parameter(Position=4, Mandatory=$false)]
		[System.String]
		$OrganizationalUnit = ("CN=users," + ([ADSI]"LDAP://RootDSE").defaultNamingContext)
	)
	
	$userpwd = ConvertTo-SecureString -AsPlainText $password -Force

    1..$count | %{
		$r1 = Get-Random -Min 1 -Maximum 1000
		$r2 = Get-Random -Min 1 -Maximum 1000
		
		$firstname = $users[$r1].firstname
		$lastname = $users[$r2].lastname
		
		$upn = "$($firstname[0])$lastname@$UpnSuffix"
		$name = "$firstname $lastname"
		$alias = "$($firstname[0])$lastname".ToLower()
		
		if(!(Get-Mailbox $alias -ErrorAction SilentlyContinue)){
			if(!$database) {
				New-Mailbox -Alias $alias -userprincipalname $upn -firstname `
				$firstname -lastname $lastname -password $userpwd -name `
				$name -organizationalunit $OrganizationalUnit -DisplayName $name
			}
			else {
				New-Mailbox -Alias $alias -userprincipalname $upn -firstname `
				$firstname -lastname $lastname -password $userpwd -database $database -name $name `
				-organizationalunit $OrganizationalUnit -DisplayName $name
			}			
		}    
    }	
}

Function Send-LabMailMessage {
<#
	.SYNOPSIS
		The function generates test email data and is inteded only for lab use.

	.DESCRIPTION
		This function sends an email message using the EWS Managed API with a message
		size of at least 100kb.
		
	.PARAMETER  PrimarySmtpAddress
		Specifies the recipients email address. This value can be bound automatically
		from the pipeline when used with Get-Mailbox or Get-DistributionGroup.
		
	.PARAMETER  MessageSize
		This parameter allows you to specify the size of the message. 
		If no value is provided the default value of 100kb will be used.

	.PARAMETER  Count
		The number of email messages to be sent, the default value is 1.
		
	.PARAMETER  Url
		The number of email messages to be sent, the default value is 1.
		
	.PARAMETER  Version
		Use the version parameter to specify the Exchange version. By default, the
		version is set to Exchange2010. If you are working with Exchange 2007, set 
		the version parameter value to Exchange2007_SP1.

	.EXAMPLE
		Send-LabMailMessage administrator@contoso.com -Count 25
		
		Description
		-----------
		Sends 25 email messages to admnistrator@contoso.com	

	.EXAMPLE
		Get-Mailbox -resultsize unlimited | Send-LabMailMessage -MessageSize 1mb
		
		Description
		-----------
		Sends an email to every mailbox in the organization with an attachment
		1mb in size.
		
	.EXAMPLE
		Send-LabMailMessage -To administrator@contoso.com -Url https://mail.contoso.com/ews/exchange.asmx
		
		Description
		-----------
		Sends a message to the administrator mailbox and manually specifies the EWS
		Url using the -Url parameter. Autodiscover will not be attempted when using the
		-Url parameter.
		
	.EXAMPLE
		Send-LabMailMessage -To administrator@contoso.com -Version Exchange2007_SP1
		
		Description
		-----------
		Sends an email to the administrator mailbox on an Exchange 2007 server.		
		
	.NOTES
		Author: Mike Pfeiffer
		Blog  : http://www.mikepfeiffer.net/		
#>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
		[alias("To")]
        [Object]
        $PrimarySmtpAddress,

        [Parameter(Position=1, Mandatory=$false)]
        [System.Int64]
        $MessageSize = 100kb,
        		
        [Parameter(Position=2, Mandatory=$false)]
        [System.Int32]
        $Count = 1,

        [Parameter(Position=3, Mandatory=$false)]
        [System.String]
        $Url,

        [Parameter(Position=4, Mandatory=$false)]
        [System.String]
        $Version = "Exchange2010_SP2"		
        )

	begin {
		Add-Type -Path "$psScriptRoot\bin\Microsoft.Exchange.WebServices.dll"
		$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
		$user = [ADSI]"LDAP://<SID=$sid>"
		
		$path = [System.IO.Path]::GetTempFileName()
		$file = [io.file]::Create($path)
		$file.SetLength($MessageSize)
		$file.Close()
	}
	
	process {
		1..$Count | %{
			$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $Version
			
			if($Url) {
				$uri = New-Object System.Uri -ArgumentList $Url
				$service.url = $uri
			}
			else {
				$service.AutodiscoverUrl($user.Properties.mail)
			}
		
			if($PrimarySmtpAddress.GetType().fullname -eq "Microsoft.Exchange.Data.SmtpAddress") {
				$Recipient = $PrimarySmtpAddress.ToString()
			}
			else {
				$Recipient = $PrimarySmtpAddress
			}
			$message = "Test Message from ExchangeLab Module $((Get-Date).ToString())"
			$mail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
			$mail.Subject = $message
			$mail.Body = $message
			[Void]$mail.Attachments.AddFileAttachment($path)
			[Void] $mail.ToRecipients.Add($Recipient)
			$mail.Send()
		}
	}
	
	end {
		Remove-Item -Path $path -Force -Confirm:$false
	}
}