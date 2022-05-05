<###############################
Title: GET ClaimsXten Active Directory Roles and how they are configured
Author: TW
Original: 2022_05_05
Last Updated: 2022_05_05
	

Overview:
- This script gets the member and memberof each application role for claimsXten. 
- The reason why this is needed is because we are deleting all of the application roles per new active directory structure and if we need to revert back we have how the roles were configure in AD
###############################>


$OUpaths = @()
$dev_domain = 'DEV-Domain' 
$us_domain = 'PROD-Domain' 

#create excel application
$excel = New-Object -ComObject excel.application

#Make Excel Visable 
$excel.application.Visible = $true
$excel.DisplayAlerts = $false

#Create WorkBook
$workBook = $excel.Workbooks.Add()

$row = 1
$column = 1

foreach($OUpath in $OUpaths){
    #Update Default Sheet Name
    $workSheets = $workBook.worksheets.add()
    $OUName = ($OUpath -split ',')[0].Substring(3)
    $workSheets.name = $OUName

    $workSheets.Activate() | Out-Null
    
    #Headers
    $workSheets.Cells.Item($row,$column) = "Application Role"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "MemberOf"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Members"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Member"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15

    #Resets column back to 1 for each worksheet
    $column = 1
    if(($OUName -eq 'CXtenDEV') -or ($OUName -eq 'CXtenQA') -or ($OUName -eq 'CXten7D')){
        $info = @(Get-ADGroup -Filter 'Name -like "Can*" -or Name -like "Menu*"' -SearchBase $OUpath -Server $dev_domain | foreach{ Get-ADGroup $_ -Properties *} | Select SamAccountName,  MemberOf, Members, member)
            $row++
            for($i = 0; $i -lt $info.Length; $i++){
                $workSheets.Cells.Item($row,$column) = $info.SamAccountName[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.MemberOf[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.Members[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.member[$i]
                #Resets column back to 1 for each worksheet
                $column = 1
                $row++
            }

        $row = 1
    }

    if($OUName -eq 'AGP'){
        $info = @(Get-ADGroup -Filter 'Name -like "Can*" -or Name -like "Menu*"' -SearchBase $OUpath -Server $us_domain  | foreach{ Get-ADGroup $_ -Properties *} | Select SamAccountName,  MemberOf, Members, member)
            $row++
            for($i = 0; $i -lt $info.Length; $i++){
                $workSheets.Cells.Item($row,$column) = $info.SamAccountName[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.MemberOf[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.Members[$i]
                $column++
                $workSheets.Cells.Item($row,$column) = $info.member[$i]
                #Resets column back to 1 for each worksheet
                $column = 1
                $row++
            }

        $row = 1
    }
    #Auto fit everything so it looks better
    $usedRange = $workSheets.UsedRange
    $usedRange.EntireColumn.AutoFit() | Out-Null
}



#Delete Default Sheet
$workbook.worksheets.item("Sheet1").Delete()

#Save the file
$workbook.SaveAs("\\va01pstodfs003.corp.agp.ads\apps\Local\EMT\COTS\McKesson\ClaimsXten\Active Directory\CXT_AD_Roles.xlsx")

#close workbook
#$workbook.Close

#Quit the application
$excel.Quit()

#Release COM Object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null