PARAM
(
  $Username,
  $Password
)

BEGIN
{
	# CONFIG **********************************************************************
	$Server = "10.1.0.40:8090"
	$Log = "log-export.txt"
	#******************************************************************************
	# ABOUT: Version: 1.0, Author: kimberg88@gmail.com
	# REQUIREMENT: Webservice: "Visma Enterprise-WS (BrukerAdministrasjon)"
	
	
	$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
	$Credentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)
	
	Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory
	
	"$(Get-Date) :: Export Start" | Out-File $Log -Append
}
PROCESS
{
	$User = $_
	"$(Get-Date) :: Process changes for $($User.'[Anchor]') ..." | Out-File $Log -Append
	
	foreach ($ChangeAttribute in $User.'[ChangedAttributeNames]') {
		if ($User.'[ObjectModificationType]' -match '(Add|Replace)') {
			switch ($ChangeAttribute) {
				"HRM_Mail" {
					"$(Get-Date) :: Change VismaHRM-attribute 'email' to: $($User.HRM_Mail)" | Out-File $Log -Append
					$URI = "http://$($Server)/enterprise_ws/secure/user/$($User.HRM_UserID)/email/WORK/$($User.HRM_Mail)"
					#Invoke-RestMethod -Uri $URI -Method PUT -Credential $Credentials
				}
				"HRM_TelephoneWork" {
					"$(Get-Date) :: Change VismaHRM-attribute 'Telephone Work' to: $($User.HRM_TelephoneWork)" | Out-File $Log -Append
					$URI = "http://$($Server)/enterprise_ws/secure/user/$($User.HRM_UserID)/phone/WORK/$($User.HRM_TelephoneWork)"
					#Invoke-RestMethod -Uri $URI -Method PUT -Credential $Credentials
				}
				"HRM_TelephonePrivate" {
					"$(Get-Date) :: Change VismaHRM-attribute 'Cellphone' to: $($User.HRM_TelephonePrivate)" | Out-File $Log -Append
					$URI = "http://$($Server)/enterprise_ws/secure/user/$($User.HRM_UserID)/phone/PRIVATE/$($User.HRM_TelephonePrivate)"
					#Invoke-RestMethod -Uri $URI -Method PUT -Credential $Credentials
				}
				"HRM_Alias" {
					"$(Get-Date) :: Change VismaHRM-attribute 'Alias' to: $($User.HRM_Alias)" | Out-File $Log -Append
					try {
						$URI = "http://$($Server):8090/enterprise_ws/secure/user/$($User.HRM_UserID)/username"
						$CurrentAliases = Invoke-RestMethod -Uri $URI -Method GET -Credential $Credentials
						$CurrentAliases = $CurrentAliases.usernames.alias.username
						"                          CURRENT: $($CurrentAliases)"  | Out-File $Log -Append
						
						# Add the alias only if it is not already present
						if (!($CurrentAliases) -or ($CurrentAliases -and !$CurrentAliases.ToLower().Contains($User.HRM_Alias.toLower()))) {
							"                          ADD: $($User.HRM_Alias)" | Out-File $Log -Append
							$URI = "http://$($Server)/enterprise_ws/secure/user/$($User.HRM_UserID)/username"
							$postParams = @{user="$($User.HRM_Alias)";}
							Invoke-RestMethod -Uri $URI -Method POST -Credential $Credentials -Body $postParams
						}
						
						# Delete "leftover" aliases, keeping it simple
						# This way we can flow AccountName directly from the Metaverse, else we'd have to use an advanced flow rule
						# * If the alias is the same as the initials it cannot be deleted, for some reason
						Foreach ($Alias in $CurrentAliases) {
							if ($Alias -eq $User.HRM_Alias) { continue }
							"                          DELETE: $($Alias)" | Out-File $Log -Append
							
								$URI = "http://$($Server):8090/enterprise_ws/secure/user/$($User.HRM_UserID)/username/$($Alias)"
								Invoke-RestMethod -Uri $URI -Method DELETE -Credential $Credentials
							
						}
					} catch {
							$_ | Out-File $Log -Append
					}
				}
				default {
					"$(Get-Date) :: No Add/Replace switch implemented for attribute: $($ChangeAttribute)" | Out-File $Log -Append
				}
			}
		} 
	}
}
END
{
	"$(Get-Date) :: Export End" | Out-File $Log -Append
}
