PARAM
(
  $Username,
  $Password
)

BEGIN
{
	# CONFIG **********************************************************************
	$Server = "zimbra.mgk.no"	# Remember to whitelist the FIM-server from the zimbra-ddos filter
	$URL_ZimbraSoap = "https://zimbra.mgk.no:7071"
	
	$Log = "log-export.txt"
	$DebugPreference="Continue"
	
	#******************************************************************************
	# ABOUT: Version: 0.1, Author: kimberg88@gmail.com
	
	
	Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory
	. .\ZimbraSOAP-Utils.ps1
	
	"$(Get-Date) :: Export Start" | Out-File $Log -Append
	"$(Get-Date) :: Authenticating ..." | Out-File $Log -Append
	$AuthResponse = Create-ZimbraSession $URL_ZimbraSoap $Username $Password -IsAdmin $true
}
PROCESS
{
	$User = $_
	"$(Get-Date) :: Process changes for $($User.Zimbra_Id) ..." | Out-File $Log -Append
	
	foreach ($ChangeAttribute in $User.'[ChangedAttributeNames]') {
		switch ($ChangeAttribute) {
			"Zimbra_Firstname" {
				"$(Get-Date) :: Change zimbra-attribute 'givenName' to: $($User.Zimbra_Firstname)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "givenName" $User.Zimbra_Firstname
			}
			"Zimbra_Lastname" {
				"$(Get-Date) :: Change zimbra-attribute 'sn' to: $($User.Zimbra_Lastname)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "sn" $User.Zimbra_Lastname
			} 
			"Zimbra_Displayname" {
				"$(Get-Date) :: Change zimbra-attribute 'displayName' to: $($User.Zimbra_displayName)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "displayName" $User.Zimbra_displayName
			} 
			"Zimbra_JobTitle" {
				"$(Get-Date) :: Change zimbra-attribute 'title' to: $($User.Zimbra_JobTitle)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "title" $User.Zimbra_JobTitle
			}
			"Zimbra_Telephone" {
				"$(Get-Date) :: Change zimbra-attribute 'telephoneNumber' to: $($User.Zimbra_Telephone)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "telephoneNumber" $User.Zimbra_Telephone
			}
			"Zimbra_Company" {
				"$(Get-Date) :: Change zimbra-attribute 'company' to: $($User.Zimbra_Company)" | Out-File $Log -Append
				Modify-ZimbraAccount $User.Zimbra_Id "company" $User.Zimbra_Company
			}
			default {
				"$(Get-Date) :: No switch implemented for zimbra-attribute: $($ChangeAttribute)" | Out-File $Log -Append
			}
		}
	}
}
END
{
	"$(Get-Date) :: Export End" | Out-File $Log -Append
}
