param (
	$Username,
	$Password,
	$OperationType = "Full"
)

# CONFIG **********************************************************************

$Server = "zimbra.mgk.no"
$URL_ZimbraSoap = "https://zimbra.mgk.no:7071"

$Log = "log-import.txt"
$DebugPreference = "Continue"

#******************************************************************************
# ABOUT: Version: 0.1, Author: kimberg88@gmail.com


Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory
. .\ZimbraSOAP-Utils.ps1

"$(Get-Date) :: Import Start" | Out-File $Log -Append
"$(Get-Date) :: Authenticating ..." | Out-File $Log -Append
$AuthResponse = Create-ZimbraSession $URL_ZimbraSoap $Username $Password -IsAdmin $true

"$(Get-Date) :: Send request for all users" | Out-File $Log -Append
$response = Get-ZimbraUser
$UserXML = (Select-Xml -Xml $response -Namespace $zimbraNameSpaces -XPath "//soap:Envelope/soap:Body/zimbraAdmin:SearchDirectoryResponse/zimbraAdmin:account").node

Foreach ($User in $UserXML) {
	$zimbraId					= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='zimbraId']").node.InnerText
	$uid						= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='uid']").node.InnerText
	$mail						= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='mail']").node.InnerText
	$zimbraAccountStatus		= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='zimbraAccountStatus']").node.InnerText
	$zimbraLastLogonTimestamp	= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='zimbraLastLogonTimestamp']").node.InnerText
	$zimbraMailTransport		= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='zimbraMailTransport']").node.InnerText
	$zimbraMailHost				= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='zimbraMailHost']").node.InnerText
	$sn							= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='sn']").node.InnerText
    $GivenName					= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='givenName']").node.InnerText
	$displayName				= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='displayName']").node.InnerText
	$title						= (Select-Xml -Xml $User -Namespace $zimbraNameSpaces -XPath "zimbraAdmin:a[@n='title']").node.InnerText
		
	"$(Get-Date) :: Project user: $($uid) ($($zimbraId))" | Out-File $Log -Append
	
	$obj = @{}
	$obj.add("Id", $zimbraId)
	$obj.add("objectClass", "user")
	$obj.add("Zimbra_AccountName", $uid)
	$obj.add("Zimbra_Id", $zimbraId)
	$obj.add("Zimbra_Firstname", $GivenName)
	$obj.add("Zimbra_Lastname", $sn) 
	$obj.add("Zimbra_DisplayName", $displayName) 
    $obj.add("Zimbra_AccountStatus", $zimbraAccountStatus) 
	$obj.add("Zimbra_JobTitle", $title) 
	$obj
}




    
