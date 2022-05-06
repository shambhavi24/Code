$url = "https://nttds.service-now.com/api/now/table/incident?assignment_group=6ccc176fdbdf6b00729372e9af961943&incident_state=2&incident_state=-2&sysparm_fields=number,assigned_to,short_description,sys_updated_on,company,incident_state,sys_created_on,priority&sysparm_display_value=true&sysparm_exclude_reference_link=true"


Get-Date -UFormat "%d/%m/%Y %R"
Write-Host "$date" 

$username = "automation.service"
$password = "automation2016"
$authInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$username`:$password"))
$headers = @{"X-Requested-With"="powershell";"Authorization"="Basic $authInfo"}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$result = Invoke-RestMethod -Uri $url -Headers $headers -method get | ConvertTo-Json

$resobj = ConvertFrom-Json -InputObject $result

$body = $resobj.result

$header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
#Alert {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }

</style>


"@

$bodyarry = @()
for($i=0; $i -lt $body.length; $i++){
$bodyarry += $body[$i].sys_created_on
}

#$strArray = $body | Foreach {"$($_.sys_created_on)"}
$currentdate = Get-Date -UFormat %R
$currentdate = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentdate, [System.TimeZoneInfo]::Local.Id, 'India Standard Time')
$array1=@()
#for($i=0; $i -lt $strArray.length; $i++){
#  $converteddate = [Datetime]$strArray[$i]
#  $array1 += $converteddate
#  }
for($i=0; $i -lt $bodyarry.length; $i++){
  $converteddate = [Datetime]$bodyarry[$i]
  $array1 += $converteddate
  }

#convert object to string:

$array = @()
$z= for($i=0; $i -lt $array1.length; $i++){
  $ts = New-TimeSpan -Start $array1[$i] -End $currentdate 
  $array += [math]::Round($ts.TotalHours)
}

for($i=0; $i -lt $array.length; $i++)
{
    if ($array.length -eq $body.length){
    $body[$i] | Add-Member -NotePropertyName AgeInHours -NotePropertyValue $array[$i] -Force
    }
}

#SLA checks:

for($k=0; $k -lt $body.Length; $k++){
    if(($body[$k].priority.Substring(0,1) -contains "3") -and ($body[$k].AgeInHours -ge "24")) {
        $body[$k] | Add-Member -NotePropertyName Notes -NotePropertyValue Alert -Force
       }
     }

#####Test these checks:
for($k=0; $k -lt $body.Length; $k++){
    if(($body[$k].priority.Substring(0,1) -contains "2") -and ($body[$k].AgeInHours -eq "8")) {
        $body[$k] | Add-Member -NotePropertyName Notes -NotePropertyValue Alert -Force
       }
     }
for($k=0; $k -lt $body.Length; $k++){
    if(($body[$k].priority.Substring(0,1) -contains "1") -and ($body[$k].AgeInHours -ge "4")) {
        $body[$k] | Add-Member -NotePropertyName Notes -NotePropertyValue Alert -Force
       }
     }
##########

$obj = New-Object -TypeName psobject

for($k=0; $k -lt $body.Length; $k++){
    for($m=0; $m -lt $body.Length; $m++){
        
        $obj[$m] | Add-Member -NotePropertyName Number -NotePropertyValue $body.number -Force
        $obj[$m] | Add-Member -NotePropertyName CreatedOn -NotePropertyValue $body[$k].sys_created_on -Force
        $obj[$m] | Add-Member -NotePropertyName State -NotePropertyValue $body[$k].incident_state -Force
        $obj[$m] | Add-Member -NotePropertyName AssignedTo -NotePropertyValue $body[$k].assigned_to -Force
        $obj[$m] | Add-Member -NotePropertyName Description -NotePropertyValue $body[$k].short_description -Force
        $obj[$m] | Add-Member -NotePropertyName Company -NotePropertyValue $body[$k].company -Force
        $obj[$m] | Add-Member -NotePropertyName LastUpdatedOn -NotePropertyValue $body[$k].sys_updated_on -Force
        $obj[$m] | Add-Member -NotePropertyName AgeInHours -NotePropertyValue $body[$k].AgeInHours -Force
        $obj[$m] | Add-Member -NotePropertyName Priority -NotePropertyValue $body[$k].priority -Force
        $obj[$m] | Add-Member -NotePropertyName Notes -NotePropertyValue $body[$k].Notes -Force
    }
}

$Report = $obj | ConvertTo-Html -Property Number,CreatedOn,State,AssignedTo,Description,Company,LastUpdatedOn,AgeInHours,Priority,Notes  -PreContent "<p>Hi Team, Please find the Incidents Report for: $currentdate <p>"

$final = $Report = ConvertTo-HTML -Body "$Report" -Head $header -PostContent "<p>Regards, DSO Team<p>" -
$a = $final | Out-File .\Incident-Report.html | Out-String
     
   