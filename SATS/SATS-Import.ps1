param (
	$Username,
	$Password,
	$OperationType = "Full"
)

# CONFIG **********************************************************************

$FailSafe = 700  # Throw an error if the number of returned users is less then this

$XML = "feideData.xml"

$ADDomain = "mgk.lab"
$ADPathEmployees = "OU=[FIM] Ansatte,OU=MGK,DC=mgk,DC=lab"  # OBS! LAB
$ADPathStudents  = "OU=[FIM] Elever,OU=MGK,DC=mgk,DC=lab"   # OBS! LAB
$ADPathGroups  = "OU=[FIM] Grupper,OU=MGK,DC=mgk,DC=lab"  # OBS! LAB

$ADLDSDomain = "@midtre-gauldal.kommune.no"
$ADLDSPathEmployees = "OU=Ansatte,OU=FeideKatalog,DC=midtregauldal,DC=kommune,DC=no"
$ADLDSPathStudents  = "OU=Elever,OU=FeideKatalog,DC=midtregauldal,DC=kommune,DC=no"
$ADLDSPathOrganizations  = "CN=Organization,DC=midtregauldal,DC=kommune,DC=no" #Trap

$Log = "log.txt"

$OrgNr = "NO970187715"
$OrgName = "Sør-Trøndelag"
$OrgMail = "postmottak@midtre-gauldal.kommune.no"
$OrgTlf = "72403000"
$OrgAddress = "Rørosveien 11"
$norEduOrgSchemaVersion = "1.5"

#**************************************************************************


Set-Location $(split-path -parent $MyInvocation.MyCommand.Definition) # Set working directory to script directory

$Groups = @{}
$global:ReturnedUsers = 0

Function AddToGroup($GroupName, $MemberSSN) {
    If($Groups.Count -gt 0) { # Pretty, pretty Lame
        If($Groups.Get_Item($GroupName)) {
            if(-not $Groups.Get_Item($GroupName).Contains($SSN)) {
                $Members = $Groups.Get_Item($GroupName) + $SSN
                $Groups.Set_Item($GroupName, $Members)
            }
        } else { 
            $Groups.Add($GroupName, @($SSN))
        }
    } else {
        $Groups.Add($GroupName, @($SSN))
    }
}

try {
    $Culture = (Get-Culture).TextInfo
	[xml]$PersonXML =  Get-Content $XML
    "$(Get-Date) :: Import Start" | Out-File $Log -Append
 

    "$(Get-Date) :: Process Group Relations" | Out-File $Log -Append
    Foreach ($Relation in $PersonXML.document.relation) {
        if($Relation.subject.groupid.groupidtype -eq "kl-ID") {
            $GroupName = $Relation.subject.groupid.'#text'
            $SSN = $Relation.object.personid[0].'#text'        
            
            AddToGroup $GroupName $SSN
        }

        if ($Relation.subject.org) {
            $UserType = $Relation.relationtype
            $GroupName = $Relation.subject.org.ouid.'#text'

            switch ($UserType) { 
                "has-pupil"   { $GroupName_All = $GroupName + "_ALLE-ELEVER" }
                "has-teacher" { $GroupName_All = $GroupName + "_ALLE-LAERERE" }
                "has-staff"   { $GroupName_All = $GroupName + "_ALLE-ANSATTE" }
            }

            Foreach ($P in $Relation.object.personid) {
                if ($P.personidtype -eq "Fnr") {
                    $SSN = $P.'#text'
                    AddToGroup ($GroupName + "_ALLE") $SSN
                    AddToGroup $GroupName_All $SSN
                }
            }
        }

    }

    Foreach ($Group in $Groups.GetEnumerator()) {
        $GroupName = $Group.key
        $GroupName = "FIM-SATS." + $GroupName.Replace("/", " ").Replace("\", " ").Replace(":","_") # Trim unwanted characters
       "$(Get-Date) :: Project Group: $($GroupName)" | Out-File $Log -Append

        if ($GroupName.Length -gt 60) { # Max CN limit in Active Directory is 64
            $GroupName = $GroupName.Substring(0,60) 
        }
            
        $obj = @{}
        $obj.add("SATS_name", $GroupName)
		$obj.add("objectClass", "group")
        $obj.add("SATS_ADPath", $ADPathGroups)
        $obj.add("Member", $Group.Value) 
        $obj
    } 

    Foreach ($Person in $PersonXML.document.person) {
        $SSN = $Person.personid[0].'#text'

        "$(Get-Date) :: Project Person: $SSN " | Out-File $Log -Append

        $MemberOfGroups = $Groups.GetEnumerator() | Where-Object { $_.Value.Contains($SSN) }
        if($MemberOfGroups | Where-Object { $_.Name.contains('ELEVER') } ) {
            $Type = "student"
            $ADPath = $ADPathStudents
            $ADLDSPath = $ADLDSPathStudents
        } else {
            $Type = "employee"
            $ADPath = $ADPathEmployees
            $ADLDSPath = $ADLDSPathEmployees
        }

        $obj = @{}
		$obj.add("SATS_ssn", $SSN)
		$obj.add("objectClass", "user")
        $obj.add("SATS_Fullname", $Person.name.fn)
        $obj.add("SATS_Firstname", $Person.name.n.given)
        $obj.add("SATS_Lastname", $Person.name.n.family)
        $obj.add("SATS_Status", "Active")
        $obj.add("SATS_Comment", "FIM-SATS: " + $Type)
        $obj.add("SATS_Type", $Type)
        $obj.add("SATS_ADPath", $ADPath)
        $obj.add("SATS_ADLDSPath", $ADLDSPath)
        $obj.add("SATS_ADDomain", $ADDomain)
        $obj.add("SATS_ADLDSDomain", $ADLDSDomain)
        $obj
		$global:ReturnedUsers++
    }


    Foreach ($Unit in $PersonXML.document.organization.ou) {
        $OrgNR = "NO" + $Unit.ouid[1].'#text'
        "$(Get-Date) :: Project Unit: $OrgNR" | Out-File $Log -Append

        $obj = @{}
        $obj.add("SATS_OrgNr", $OrgNR)
	    $obj.add("objectClass", "unit")
        $obj.add("SATS_OrgName", $Unit.ouname[0].'#text')
        $obj.add("SATS_OrgMail", $OrgMail) #$Unit.contactinfo[2].'#text'
        $obj.add("SATS_telephone", $Unit.contactinfo[0].'#text') 
        $obj.add("MemberOf", $OrgNr) 
        $obj.add("SATS_ADLDSPath", $ADLDSPathOrganizations)
        $obj.add("SATS_ADLDSDomain", $ADLDSDomain)
        $obj
    }

	
    "$(Get-Date) :: Project Main Organization" | Out-File $Log -Append
    $obj = @{}
    $obj.add("SATS_OrgNr", $OrgNr)
	$obj.add("objectClass", "organization")
    $obj.add("SATS_OrgName", $OrgName)
    $obj.add("SATS_OrgMail", $OrgMail ) 
    $obj.add("SATS_telephone", $OrgTlf ) 
    $obj.add("SATS_postalAddress", $OrgAddress) 
	$obj.add("SATS_norEduOrgSchemaVersion", $norEduOrgSchemaVersion) 
    $obj.add("SATS_ADLDSPath", $ADLDSPathOrganizations)
    $obj.add("SATS_ADLDSDomain", $ADLDSDomain)
    $obj

    "$(Get-Date) :: Import End" | Out-File $Log -Append
	
	if ($ReturnedUsers -lt $FailSafe) {
		$Error = "The script returned less users then the specified failsafe-threshold. Something likely went wrong... Threshold: $($FailSafe). Returned users: $($ReturnedUsers)"
		Throw $Error
		$Error | Out-File $Log -Append
	}

} catch {
    $_ | Out-File $Log -Append
}

