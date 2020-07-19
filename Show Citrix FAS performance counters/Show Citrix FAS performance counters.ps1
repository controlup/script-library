

#Check for FAS Counters Availability

function Test-FASNamespaceExists {

$servicename_search_string = "CitrixFederatedAuthenticationService"

if(!(get-service -name $servicename_search_string -ErrorAction SilentlyContinue)){

    write-warning "FAS service unavailable, FAS not installed?"

    break

}

}

 

Function Get-FASPerformanceData{

   #confirm FAS installed first

   Test-FASNamespaceExists

   $paths=@()

   $paths+="\Citrix Federated Authentication Service\Active Sessions"

   $paths+="\Citrix Federated Authentication Service\Average Concurrent Certificate Signing Requests"

   $paths+="\Citrix Federated Authentication Service\Average Private Key Operations per Minute"

   $paths+="\Citrix Federated Authentication Service\Average Request Time Milliseconds"

   $paths+="\Citrix Federated Authentication Service\Certificate Count"

   $paths+="\Citrix Federated Authentication Service\Certificate Signing Requests per Minute"

   $paths+="\Citrix Federated Authentication Service\High Load Level"

   $paths+="\Citrix Federated Authentication Service\Medium Load Level"

   $paths+="\Citrix Federated Authentication Service\Low Load Level"

   $counters=get-counter -Counter $paths

   $counters.countersamples | select @{Expression={$_.path.split("\")[-1]};Label="Counter Name"},@{Expression={$_.cookedvalue};Label="Value"}

}

 

Get-FASPerformanceData
