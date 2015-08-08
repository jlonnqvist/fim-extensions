$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-SATS_SSN|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "user"
$obj | Add-Member -Type NoteProperty -Name "SATS_Firstname|String" -Value "Soren"
$obj | Add-Member -Type NoteProperty -Name "SATS_Lastname|String" -Value "Granfeldt"
$obj | Add-Member -Type NoteProperty -Name "SATS_Fullname|String" -Value "Soren Granfeldt"
$obj | Add-Member -Type NoteProperty -Name "SATS_Status|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_Comment|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_Type|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADDomain|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSDomain|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "MemberOf|Reference[]" -Value (2,3)
$obj

$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-SATS_Name|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "group"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "Member|Reference[]" -Value (2,3)
$obj

$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-SATS_OrgNr|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "organization"
$obj | Add-Member -Type NoteProperty -Name "SATS_OrgName|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_OrgMail|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_norEduOrgSchemaVersion|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_telephone|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_postalAddress|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSDomain|String" -Value "Active"
$obj

$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-SATS_OrgNr|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "unit"
$obj | Add-Member -Type NoteProperty -Name "SATS_OrgName|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_OrgMail|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_telephone|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_postalAddress|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "MemberOf|Reference[]" -Value (2,3)
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "SATS_ADLDSDomain|String" -Value "Active"
$obj
