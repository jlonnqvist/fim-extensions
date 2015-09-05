$zimbraNameSpaces = @{
					'soap'='http://www.w3.org/2003/05/soap-envelope';
					'zimbraAccount'='urn:zimbraAccount';
					'zimbraAdmin'='urn:zimbraAdmin';
					'zimbraMail'='urn:zimbraMail';
				  }
				  
function Get-ZimbraUser {
	param ([string]$username)
	
	if ($username) {
		$ldapfilter = 	"uid=$username"
	} else {
		$ldapfilter = 	"&amp;" +
						"(!(zimbraIsSystemAccount=TRUE))" +
						"(!(zimbraIsExternalVirtualAccount=TRUE))" +
						"(!(zimbraIsAdminAccount=TRUE))" 
	}

	$attributes = "givenName,displayName,title,sn,zimbraId,zimbraAccountStatus,uid,zimbraMailTransport,mail,zimbraMailHost"
	$request =  "<SearchDirectoryRequest xmlns='urn:zimbraAdmin' offset='0' limit='0' sortBy='name' sortAscending='1' applyCos='false' applyConfig='false' attrs='$($attributes)' types='accounts'>" +
					"<query>($($ldapfilter))</query>" +
				"</SearchDirectoryRequest>"

	send-request $request
}
				  
function Modify-ZimbraAccount {
	param ([string]$zid, [string]$field, [string]$newValue)
	
	$request  = 	"<ModifyAccountRequest xmlns='urn:zimbraAdmin' id='$($zid)'>" +
						"<a n='$($field)'>$($newValue)</a>" +
					"</ModifyAccountRequest>"
	
	send-request $request
}
				  
function Create-ZimbraSession {
    param ([string]$server, [string]$account, [string]$password, [switch]$isAdmin)
        
	begin {
		function NewZimbraSession([string]$server, [string]$account) {
			$zSession = new-object System.Object
			$zSession | add-member -membertype noteproperty -name Server -value $server
			$zSession | add-member -membertype noteproperty -name Account -value $account
			$zSession | add-member -membertype noteproperty -name AuthToken -value $null
			$zSession | add-member -memberType noteproperty -name Request -value $null
			$zSession | add-member -membertype noteproperty -name Response -value $null
			$zSession
		}
		
        function Usage() {
			""
			"USAGE"
			"    Create-ZimbraSession -Server <soap-service-url> -Account <zimbra-account> -Password <password> [-IsAdmin <`$true|`$false>]"
			""
			"SYNOPSIS"
			"    Creates a new session to the zimbra server for the given user."
			"    Sets `$global:zimbraSession to the created ZimbraSession object"
			"    which contains the server, account, authToken and AuthResponse"
			""
			"EXAMPLES"
			"    Create-ZimbraSession https://zimbra.example.com user@example.com test123"
			"    Create-ZimbraSession https://zimbra.example.com admin@example.com test123 -isAdmin"
			"    Create-ZimbraSession -Server http://zimbra.example.com:7443 -Account user@example.com -Password test123"
			""
        }

        if (($args[0] -eq "-?") -or ($args[0] -eq "-help")) {
			Usage
        }
    }
    
    process {
    }
      
    end {
		if ($server -and $account -and $password) {
			if ($isAdmin) {
				$path = "/service/admin/soap"
				$xmlns = "urn:zimbraAdmin"
			} else {
				$path = "/service/soap"
				$xmlns = "urn:zimbraAccount"
			}
			$url = $server + $path
		
			$global:zimbraSession = NewZimbraSession $url $account
			
			$authRequest = "<AuthRequest xmlns='$($xmlns)' name='$($account)' password='$($password)'></AuthRequest>"
			send-request $authRequest
		}
    }
}


function Send-Request {
    param ([string]$request)
    begin { 
    	function Build-Request([xml]$body){
				if ($global:zimbraSession.authToken -ne "") {
					$authTokenElement = "<authToken>$($global:zimbraSession.authToken)</authToken>"
				} 
				
				"<?xml version='1.0' encoding='utf-8' ?>"
				"<soap:Envelope xmlns:soap='http://www.w3.org/2003/05/soap-envelope'>"
					"<soap:Header>"
						"<context xmlns='urn:zimbra'>$($authTokenElement)</context>"
					"</soap:Header>"
					"<soap:Body>$($body.get_innerxml())</soap:Body>"
				"</soap:Envelope>"
		}

		function Send-Request([xml]$request) {
			try {
				$server = $global:zimbraSession.server
				
				# Request
				$request = Build-Request $request
				$global:zimbraSession.Request = [xml]$request
				 
				write-debug (" ----Server--------------------------------- ") 
				write-debug ($server)
				write-debug (" ----Request-------------------------------- ")
				write-debug ($request.innerXml)
				write-debug (" ******************************************* ") 
				 
				# Response:
				$response = Invoke-RestMethod -Method Post -Body $request -Uri $server -ErrorAction Stop
				$global:zimbraSession.Response = $null
				$authToken = $response.Envelope.Body.AuthResponse.AuthToken
				
				if ($authToken -ne $null -and $authToken -ne "") {
					$global:zimbraSession.AuthToken = $authToken
				} 

				$xmlResponseStr = $response.InnerXml
					
				write-debug (" ----Response------------------------------- ")
				write-debug ($xmlResponseStr)
				write-debug (" ******************************************* ") 
				
				if ($response.Envelope.Body.Fault.Reason.Text -ne $null) {
					throw $xmlResponseStr
				}
				
				# Return response
				$global:zimbraSession.Response = $response
				$response
			} catch {
				throw
			}
    	}
	}
    
    process {
    }
	    
    end {
		if ($request) {
			Send-Request $request
		}
    }
}
