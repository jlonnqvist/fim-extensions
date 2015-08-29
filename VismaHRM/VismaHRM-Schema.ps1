$obj = New-Object -Type PSCustomObject

$obj | Add-Member -Type NoteProperty -Name "Anchor-HRM_EmployeeID|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "user"
$obj | Add-Member -Type NoteProperty -Name "HRM_ssn|String" -Value "99"
$obj | Add-Member -Type NoteProperty -Name "HRM_firstname|String" -Value "Soren"
$obj | Add-Member -Type NoteProperty -Name "HRM_lastname|String" -Value "Granfeldt"
$obj | Add-Member -Type NoteProperty -Name "HRM_fullname|String" -Value "Soren Granfeldt"
$obj | Add-Member -Type NoteProperty -Name "HRM_status|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_mainDepartment|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_jobtitle|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_type|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_comment|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_ADPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_ADPathDisabled|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "HRM_ADDomain|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "MemberOf|Reference[]" -Value (2,3)

$obj


$obj = New-Object -Type PSCustomObject

$obj | Add-Member -Type NoteProperty -Name "Anchor-Id|String" -Value 1
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "group"
$obj | Add-Member -Type NoteProperty -Name "displayName|String" -Value "TestGTemplate"
$obj | Add-Member -Type NoteProperty -Name "HRM_ADPath|String" -Value "Active"
$obj | Add-Member -Type NoteProperty -Name "Member|Reference[]" -Value (2,3)

$obj