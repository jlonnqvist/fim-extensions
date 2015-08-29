param (
	$Username,
	$Password,
	$OperationType = "Full"
)

# CONFIG **********************************************************************

$FailSafe = 700  # Throw an error if the number of returned users is less then this

$ADDomain = "mgk.no" 
$ADPathEmployees = "OU=[FIM] Ansatte,OU=MGK,DC=mgk,DC=no" 
$ADPathDisabledEmployees = "OU=Deaktivert,OU=[FIM] Ansatte,OU=MGK,DC=mgk,DC=no" 
$ADPathGroups    = "OU=[FIM] Grupper,OU=MGK,DC=mgk,DC=no" 

$Server = "10.1.0.40:8090"
$WebRequestTimeout = 999

$CompanyID = 1
$EmployeeStartID = 1
$EmployeeEndID = 9999
$DaysUntillStart = 30 # Get employees that have not started yet as well

$Log = "log.txt"

#******************************************************************************
# ABOUT: Version: 0.4, Author: kimberg88@gmail.com


Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory
$Groups = @{}
$global:ReturnedUsers = 0

Function ProjectGroups($Groups) {
    Foreach ($Group in $Groups.GetEnumerator()) {
        $GroupName = "FIM-VISMA." + $Group.Key

        $GroupName = $GroupName.Replace("/", " ").Replace("\", " ").Replace(":", " ").Replace(","," ") # Trim unwanted characters

        if ($GroupName.Length -gt 60) { # Max CN limit in Active Directory is 64
            $GroupName = $GroupName.Substring(0,60) 
        }

        $obj = @{}
        $obj.add("Id", $GroupName)
        $obj.add("objectClass", "group")
        $obj.add("displayName", $GroupName)
        $obj.add("HRM_ADPath", $ADPathGroups)
        $obj.add("Member", $Group.Value) 
        $obj
        
    }
}

Function ProjectUsers($URI, $OnlyFutureEmployees) {
    try {
	    "[$(Get-Date)] Request URI: $($URI)" | Out-File $Log -Append
	
        $Culture = (Get-Culture).TextInfo
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

        $WebPage = Invoke-RestMethod $URI -Credential $Credentials -TimeoutSec $WebRequestTimeout
	    $EmployeeXML =  $WebPage.personsXML.person
	
        Foreach ($Employee in $EmployeeXML) {
            $EmployeeID = $Employee.employments.employment.employeeId.ToString()
            $EmployeeStart = $Employee.employments.employment.startDate
            "[$(Get-Date)] Process employee: $($EmployeeID)" | Out-File $Log -Append

            if ($OnlyFutureEmployees) { 
                if ($(Get-Date) -gt [DateTime]$EmployeeStart) {
                    # To avoid returning duplicates
                    Continue
                }
            }

            $obj = @{}

		    $obj.add("HRM_employeeID", $EmployeeID)
		    $obj.add("objectClass", "user")

            $obj.add("HRM_ssn", $Employee.ssn.toString())
            $obj.add("HRM_firstname", $Culture.ToTitleCase($Employee.givenName.toLower()).Trim())
            $obj.add("HRM_lastname", $Culture.ToTitleCase($Employee.familyName.toLower()).Trim())
            $obj.add("HRM_fullname" , $Culture.ToTitleCase(($Employee.givenName + " " + $Employee.familyName).toLower()).Trim())
            $obj.add("HRM_type", "employee")
            $obj.add("HRM_ADPath", $ADPathEmployees)
			$obj.add("HRM_ADPathDisabled", $ADPathDisabledEmployees)
			$obj.add("HRM_ADDomain", $ADDomain)
			$obj.add("HRM_status", "Active") # Inactive users are not returned by the web service

            Foreach ($Job in $Employee.employments.employment.positions.position) {
                if ([DateTime]$Job.validFromDate -gt $(Get-Date)) { 
                    # Skip jobs that have not started
                    continue 
                } 

                $Department = $Job.costCentres.dimension2.name

                if ($Department) {
                    If($Groups.Count -gt 0) { # Pretty, pretty Lame
                        If($Groups.Get_Item($Department)) {
                            if(-not $Groups.Get_Item($Department).Contains($EmployeeID)) {
                                $Members = $Groups.Get_Item($Department) + $EmployeeID
                                $Groups.Set_Item($Department, $Members)
                            }
                        } else { 
                            $Groups.Add($Department, @($EmployeeID))
                        }
                    } else {
                        $Groups.Add($Department, @($EmployeeID))
                    }

                    if ($Job.isPrimaryPosition -eq "true") {
                        $obj.add("HRM_mainDepartment", $Department)
						$obj.add("HRM_jobtitle", $Job.positionStatistics.workClassification.name)
                    }
                }
            }
			$obj.add("HRM_comment", "FIM-VISMA : $Department : Ansatt $($EmployeeStart)")
			$obj | Out-File $Log -Append
            $obj
			$global:ReturnedUsers++
        }
	
	    "[$(Get-Date)] Completed Successfully" | Out-File $Log -Append
    } catch {
        $_ | Out-File $Log -Append
    }
}


$URI = "http://$($Server)/hrm_ws/secure/persons/company/$($CompanyID)/start-id/$($EmployeeStartID)/end-id/$($EmployeeEndID)"
ProjectUsers $URI $False

$URI = "http://$($Server)/hrm_ws/secure/persons/company/$($CompanyID)/not-started/date-interval/$((Get-Date).AddDays(1).ToString('yyyy-MM-dd'))/$((Get-Date).AddDays($DaysUntillStart).ToString('yyyy-MM-dd'))"
ProjectUsers $URI $True

ProjectGroups $Groups


if ($ReturnedUsers -lt $FailSafe) {
	$Error = "The script returned less users then the specified failsafe-threshold. Something likely went wrong... Threshold: $($FailSafe). Returned users: $($ReturnedUsers)"
	Throw $Error
	$Error | Out-File $Log -Append
}