param (
	$Username,
	$Password,
	$OperationType = "Full"
)

# CONFIG **********************************************************************

$FailSafe = 1  # Throw an error if the number of returned users is less then this

$ADDomain = "mgk.no" 
$ADPathEmployees = "OU=[FIM] Ansatte,OU=MGK,DC=mgk,DC=no" 
$ADPathDisabledEmployees = "OU=Deaktivert,OU=[FIM] Ansatte,OU=MGK,DC=mgk,DC=no" 
$ADPathGroups    = "OU=[FIM] Grupper,OU=MGK,DC=mgk,DC=no" 

$Server = "10.1.0.40:8090"
$WebRequestTimeout = 999

$EmployeeStartID = 1
$EmployeeEndID = 9999
$DaysUntillStart = 30 # Get employees that have not started yet as well

$Log = "log-import.txt"

#******************************************************************************
# ABOUT: Version: 1.9, Author: kimberg88@gmail.com
# REQUIREMENT: Webservice: "Visma HRM-WS"


Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory
$Culture = (Get-Culture).TextInfo

$global:GroupsOC = @()
$global:ReturnedUsers = 0


Function ProjectUsers($URI, $OnlyFutureEmployees) {
    try {
	    "[$(Get-Date)] Request URI: $($URI)" | Out-File $Log -Append
		$Culture = (Get-Culture).TextInfo
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

        $WebPage = Invoke-RestMethod $URI -Credential $Credentials -TimeoutSec $WebRequestTimeout
	    $XML =  $WebPage.personsXML.person
	
        Foreach ($Employee in $XML) {
			$VismaID = $Employee.personIdHRM
            $EmployeeID = $Employee.employments.employment.employeeId.ToString()
            $EmployeeStart = $Employee.employments.employment.startDate
			$Department = "Ingen Avdeling"
			$PrimaryDepartment = $NULL
			$JobTitle = $NULL
			$ManagerID = $NULL
			
            "[$(Get-Date)] Process employee: $($EmployeeID)" | Out-File $Log -Append

            if ($OnlyFutureEmployees) { 
                if ($(Get-Date) -gt [DateTime]$EmployeeStart) {
                    # To avoid returning duplicates
                    Continue
                }
            }
			
            Foreach ($Job in $Employee.employments.employment.positions.position) {
				# Skip jobs that have not yet started
                if ([DateTime]$Job.validFromDate -gt $(Get-Date)) { continue } 

                $Department = $Job.chart.unit.id
				try {
					if ($Department) {	
						# Store department membership information in an array that we can parse later...
						$StoredGroup = $global:GroupsOC | Where { $_.ID -eq $Job.chart.unit.id }
						if ($StoredGroup) {
							if(-not ($StoredGroup.Members -contains $VismaID)) { 
								$StoredGroup.Members += $VismaID
							}
						} else {
							if($Job.chart.unit.manager.id) {
								$Members = @($Job.chart.unit.manager.id, $VismaID) # Add the manager as well
							} else {
								$Members = @($VismaID)
							}
							$global:GroupsOC += @{
								ID = $Job.chart.unit.id
								Name = $Job.chart.unit.name
								Members = $Members
							}
						}

						if ($Job.isPrimaryPosition -eq "true") {
							$PrimaryDepartment = $Job.chart.unit.name
							
							$JobTitle = $Job.positionStatistics.workClassification.name
							
							if ($JobTitle.Contains("Rektor")) { $JobTitle = $JobTitle.substring(0, $JobTitle.indexOf("(")) }
							if ($JobTitle.Contains("Uten godkjent")) { $JobTitle = $JobTitle.substring(0, $JobTitle.indexOf("(")) }
							if ($JobTitle.Contains("vikar (grunnskole)")) { $JobTitle = "Laerer" } # Encoding issue with Æ
							if ($JobTitle.Contains("Foster")) { $JobTitle = " " }

							if ($JobTitle) {
								$JobTitle = $JobTitle.Replace("&","og").Replace("/", " ").Replace("\", " ")
							}
						}

					}
				} catch {
					"Ingen info om avdeling" | Out-File $Log -Append
				}
            }
			
			
			$CurrentAliases = $Employee.authentication.alias
			if ($CurrentAliases) { $CurrentAliases = $CurrentAliases.toLower() }
			
			$obj = @{}
		    $obj.add("HRM_EmployeeID", $EmployeeID)
			$obj.add("HRM_VismaHRMID", $VismaID)
		    $obj.add("objectClass", "user")
            $obj.add("HRM_ssn", $Employee.ssn.toString())
			$obj.add("HRM_UserID", $Employee.authentication.userId.toString())
			$obj.add("HRM_Alias", $CurrentAliases)
            $obj.add("HRM_firstname", $Culture.ToTitleCase($Employee.givenName.toLower()).Trim())
            $obj.add("HRM_lastname", $Culture.ToTitleCase($Employee.familyName.toLower()).Trim())
            $obj.add("HRM_fullname" , $Culture.ToTitleCase(($Employee.givenName + " " + $Employee.familyName).toLower()).Trim())
            $obj.add("HRM_type", "employee")
			$obj.add("HRM_mainDepartment", $PrimaryDepartment)
			$obj.add("HRM_jobtitle", $JobTitle)
			$obj.add("HRM_ManagerID", $ManagerID)
            $obj.add("HRM_ADPath", $ADPathEmployees)
			$obj.add("HRM_ADPathDisabled", $ADPathDisabledEmployees)
			$obj.add("HRM_ADDomain", $ADDomain)
			$obj.add("HRM_ADLDSPath", $ADLDSPathEmployees)
			$obj.add("HRM_ADLDSDomain", $ADLDSDomain)
			$obj.add("HRM_Affilliation", ("member", "employee", "staff"))
			$obj.add("HRM_Entitlement", $NULL)
			$obj.add("HRM_MemberOfOrganization", $OrgNr)
			$obj.add("HRM_comment", "FIM-VISMA : $($PrimaryDepartment)")
			if ($Employee.contactInfo.mobilePhone) {
				$obj.add("HRM_CellphoneWork", $Employee.contactInfo.mobilePhone.toString())
			}
			if ($Employee.contactInfo.privateMobilePhone) {
				$obj.add("HRM_CellphonePrivate", $Employee.contactInfo.privateMobilePhone.toString())
			}
            $obj
			
			$global:ReturnedUsers++
        }
	
	    "[$(Get-Date)] Completed Successfully" | Out-File $Log -Append
    } catch {
        $_ | Out-File $Log -Append
    }
}


Function ProjectGroups($GroupsOC) {
    Foreach ($Group in $GroupsOC) {
        $GroupName = "FIM-VISMA." + $Group.Name
		$GroupName = $GroupName.Replace("/", " ").Replace("\", " ").Replace(":", " ").Replace(","," ")

        if ($GroupName.Length -gt 60) { # Max CN limit in Active Directory is 64
            $GroupName = $GroupName.Substring(0,60) 
        }

        $obj = @{}
        $obj.add("Id", $Group.ID)
        $obj.add("objectClass", "group")
        $obj.add("displayName", $GroupName)
        $obj.add("HRM_ADPath", $ADPathGroups)
        $obj.add("Member", $Group.Members) 
        $obj 
    }
}



$URI = "http://$($Server)/hrm_ws/secure/persons/company/1/start-id/$($EmployeeStartID)/end-id/$($EmployeeEndID)"
ProjectUsers $URI $False

$URI = "http://$($Server)/hrm_ws/secure/persons/company/1/not-started/date-interval/$((Get-Date).AddDays(1).ToString('yyyy-MM-dd'))/$((Get-Date).AddDays($DaysUntillStart).ToString('yyyy-MM-dd'))"
ProjectUsers $URI $True

ProjectGroups $GroupsOC


if ($ReturnedUsers -lt $FailSafe) {
	$Error = "The script returned less users then the specified failsafe-threshold. Something likely went wrong... Threshold: $($FailSafe). Returned users: $($ReturnedUsers)"
	Throw $Error
	$Error | Out-File $Log -Append
}
