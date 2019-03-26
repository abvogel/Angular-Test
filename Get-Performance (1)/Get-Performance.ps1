param
    (
        [String]                                   $OrgName,
        [Object]                                   $Cred,
        [String]                                   $LogFilePath,
        [Int]                                      $IncidentTotal,
        [Int]                                      $PhoneCallTotal,
        [Int]                                      $LeadTotal,
        [Int]                                      $OpportunityTotal,
        [Int]                                      $WorkOrderTotal,
        [Int]                                      $ChatTotal,
        [Switch]                                   $AutoConfigure = $false
    )

Function Get-Performance {
    [CmdletBinding()]
    param
    (
        [String]                                   $OrgName,
        [Object]                                   $Cred,
        [String]                                   $LogFilePath,
        [Int]                                      $IncidentTotal,
        [Int]                                      $PhoneCallTotal,
        [Int]                                      $LeadTotal,
        [Int]                                      $OpportunityTotal,
        [Int]                                      $WorkOrderTotal,
        [Int]                                      $ChatTotal
    )

    Write-Host "Initializing RunspacePool";
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1,8);
	$RunspacePool.CleanupInterval = [timespan]::TicksPerMinute * 2;
	$RunspacePool.ApartmentState = [System.Threading.ApartmentState]::STA;
	$RunspacePool.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::UseNewThread;
	$RunspacePool.Open();
	$firstrun = $true;
    $Runspaces = @();
    
    Write-Host "Configuring Parameters";
    $Parameters = New-Object 'System.Collections.Generic.Dictionary[String,Object]';
        $Parameters.Add("OrgName", $OrgName);
        $Parameters.Add("Cred", $Cred);
        $Parameters.Add("LogFilePath", "");
        $Parameters.Add("IncidentTotal", $IncidentTotal);
        $Parameters.Add("PhoneCallTotal", $PhoneCallTotal);
        $Parameters.Add("LeadTotal", $LeadTotal);
        $Parameters.Add("OpportunityTotal", $OpportunityTotal);
        $Parameters.Add("WorkOrderTotal", $WorkOrderTotal);
        $Parameters.Add("ChatTotal", $ChatTotal);

    Write-Host ($Parameters | ConvertTo-JSON);
    
    Write-Host "Setting up PowerShell Instance #1";
    $Script = Get-Script;
    $Parameters.LogFilePath = $("{0}\Desktop\PerformanceLog1.csv" -f $env:USERPROFILE);
    $PSInstance1 = ([powershell]::Create().AddScript($Script).AddParameters($Parameters));
    $PSInstance1.RunspacePool = $RunspacePool;
    $Runspaces += New-Object PSObject -Property @{
        Instance = $PSInstance1;
        IAResult = $PSInstance1.BeginInvoke();
        Arguments = $Parameters;
    }

    Write-Host "Setting up PowerShell Instance #2";
    $Parameters.LogFilePath = $("{0}\Desktop\PerformanceLog2.csv" -f $env:USERPROFILE);
    $PSInstance2 = ([powershell]::Create().AddScript($Script).AddParameters($Parameters));
    $PSInstance2.RunspacePool = $RunspacePool;
    $Runspaces += New-Object PSObject -Property @{
        Instance = $PSInstance2;
        IAResult = $PSInstance2.BeginInvoke();
        Arguments = $Parameters;
    }

    Start-Sleep -Milliseconds 10;

    Write-Host ("{0} threads currently running." -f @($Runspaces | Where-Object { -not $_.IAResult.IsCompleted}).Count);

    While (($Runspaces | Where-Object { -not $_.IAResult.IsCompleted}).Count -gt 1)
    {
        Start-Sleep -Milliseconds 10;
    }
}

Function Get-Script {
    $Script = {
            param
            (
                [String]                                   $OrgName,
                [Object]                                   $Cred,
                [String]                                   $LogFilePath,
                [Int]                                      $IncidentTotal,
                [Int]                                      $PhoneCallTotal,
                [Int]                                      $LeadTotal,
                [Int]                                      $OpportunityTotal,
                [Int]                                      $WorkOrderTotal,
                [Int]                                      $ChatTotal
            )
    
            $OrgName = "bbyb2cdryrun07";
            $IncidentTotal = 210;
            $PhoneCallTotal = 239;
            $LeadTotal = 43;
            $OpportunityTotal = 36;
            $WorkOrderTotal = 116;
            $ChatTotal = 354;

            Function Set-CreateRequest {
                param
                (
                    [String]    $EntityLogicalName,
                    [Hashtable] $Fields,
                    [Object]    $Conn
                )
    
                $CreateRequest = [Microsoft.Xrm.Sdk.Messages.CreateRequest]::new();
    
                $ent = [Microsoft.Xrm.Sdk.Entity]::new()
                $ent.LogicalName = $EntityLogicalName;
    
                ForEach ($key in $Fields.Keys) {
                    $ent[$key] = $Fields[$key];
                }
    
                $CreateRequest.Target = $ent;
    
                if ($EntityLogicalName -eq "incident") {
                    $Global:IncidentCreateRequests.Add($CreateRequest);
                } elseif ($EntityLogicalName -eq "bby_chat") {
                    $Global:ChatCreateRequests.Add($CreateRequest);
                } else {
                    $Global:OtherCreateRequests.Add($CreateRequest);
                }
            }
    
            Function Get-ChatUpdateRequest {
                param
                (
                    [String]    $Id,
                    [Int]       $Iteration
                )
    
                $UpdateRequest = [Microsoft.Xrm.Sdk.Messages.UpdateRequest]::new();
                $RandValue = $RandomGenerator.Next(10);
                
                $Fields = @{
                    actualstart = $(Get-Date("2018-03-28T09:37:00-05:00")).AddHours($Iteration/(1000 - $RandValue));
                    bby_chatlog = ("[{&quot;Content&quot;:&quot;Chat session transfered to agent&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522247824329+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;Client Connected&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522247824256+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;Client Disconnected&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522248067619+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;LMI Session: 605291183 has been connected to the Case&quot;,&quot;Direction&quot;:1,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522251444640+0800)\/&quot;,&quot;Type&quot;:1}]" + ($RandomGenerator.Next(1000)) + ($RandValue.toString()));
                    bby_customeremail = ("{0}@bestbuy-fake-{1}-{2}.com" -f $env:USERNAME, $Iteration, $RandValue);
                    actualend = $(Get-Date("2018-03-28T10:37:50-05:00")).AddHours($Iteration/(1000 + $RandValue));
                    bby_customerjabberuserid = ("{0}@bestbuy-fake.com/{1}/{2}" -f $env:USERNAME, $Iteration, $RandValue);
                    subject = ("Fulfillment Chat for {0}@bestbuy-fake-{1}-{2}.com" -f $env:USERNAME, $Iteration, $Id.toString());
                };
    
                $ent = [Microsoft.Xrm.Sdk.Entity]::new()
                $ent.LogicalName = "bby_chat";
                $ent.Id = $Id;
    
                ForEach ($key in $Fields.Keys) {
                    $ent[$key] = $Fields[$key];
                }
    
                $UpdateRequest.ConcurrencyBehavior = [Microsoft.Xrm.Sdk.ConcurrencyBehavior]::AlwaysOverwrite;
                $UpdateRequest.Target = $ent;
    
                Write-Output $UpdateRequest;
            }
    
            Function Add-Incident {
                param
                (
                    [Int]                                      $Iteration
                )
                
                $NewGuid = Set-CreateRequest -EntityLogicalName incident -Fields @{
                    bby_callbackphonenumber = "1231231234";
                    bby_entitlementcustomerid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
                    bby_callbacktimezone = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630001));
                    bby_partnerid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bby_partner", "{F90CB08A-9E61-E711-8114-5065F38A7BC1}"));
                    title=$("Support Case for - {0} - Seed {1}" -f $env:USERNAME, $Iteration);
                    bby_breezecasenumber = $Iteration.toString();
                    bby_customeremail = $("{0}@bestbuy-fake.com" -f $env:USERNAME);
                    bby_lineofbusinessid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bby_lineofbusiness", "{471384A8-7ACE-E711-8125-5065F38A0A21}"));
                    ticketnumber = ("CAS-FAKE-{0}" -f $Iteration);
                    customerid=$([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
                } -conn $Conn;
    
                Write-Output 1;
            }
    
            Function Add-PhoneCall {
                param
                (
                    [Int]                                      $Iteration
                )
    
                Set-CreateRequest -EntityLogicalName phonecall -Fields @{
                    phonenumber = "(123) 123-1234";
                    description = $("Confirmed 12-545 on ETA 729 that Iteration {0} meets expected specs." -f $Iteration);
                    subject=$("{0} hour call - it's a callathon!" -f $Iteration)
                    regardingobjectid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
                    bby_calltype = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630010));
                } -conn $Conn | Out-Null;
    
                Write-Output 1;
            }
    
            Function Add-Lead {
                param
                (
                    [Int]                                      $Iteration
                )
    
                Set-CreateRequest -EntityLogicalName lead -Fields @{
                    telephone1 = "(202) 230-7555";
                    bby_categoryofsolutions = "Home Theater|Smart Home / Connected Home";
                    address1_city = "Upper Marlboro";
                    address1_postalcode = "20774";
                    bby_refempfirstname = $env:USERNAME;
                    bby_scopeofwork = "N/A";
                    emailaddress1 = ("{0}@bestbuy-fake.com" -f $env:USERNAME);
                    bby_channel = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630010));
                    msdyn_ordertype = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(192350000));
                    parentaccountid = $([Microsoft.Xrm.Sdk.EntityReference]::new("account", "{6F472761-1394-E711-80F8-3863BB35AC20}"));
                    bby_primaryphonetype = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630030));
                    bby_refempid = $env:USERNAME;
                    bby_resourcerequirementid = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_resourcerequirement", "{7A842E5B-1394-E711-80F8-3863BB35AC20}"));
                    lastname = "SomeLastName";
                    bby_refemplastname = "SomeotherLastName";
                    bby_budgetrange = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630003));
                    qualifyingopportunityid = $([Microsoft.Xrm.Sdk.EntityReference]::new("opportunity", "{BB0E9839-AC94-E711-80F8-3863BB35AC20}"));
                    bby_bookableresourcebookingid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bookableresourcebooking", "{CD625027-AC94-E711-80FD-3863BB36EDB8}"));
                    firstname = $env:USERNAME;
                    description = ("It was like that because that's how it was done. Iteration {0}" -f $Iteration);
                    subject = $("Lead for - {0} - Seed {1}" -f $env:USERNAME, $Iteration);
                    parentcontactid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{98842E5B-1394-E711-80F8-3863BB35AC20}"));
                    address1_name = "Check";
                } -conn $Conn | Out-Null;
    
                Write-Output 1;
            }
    
            Function Add-Opportunity {
                param
                (
                    [Int]                                      $Iteration
                )
    
                Set-CreateRequest -EntityLogicalName opportunity -Fields @{
                    msdyn_ordertype = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(192350000));
                    bby_referringemployeefirstname = "Ibe";
                    parentcontactid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
                    totalamount = $Iteration;
                    description = $("It's the end of the world as we know it - Iteration {0}" -f $Iteration);
                    bby_scopeofwork = "N/A";
                    name = $("{0} Indomitable" -f $env:USERNAME);
                    customerid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
                    originatingleadid = $([Microsoft.Xrm.Sdk.EntityReference]::new("lead", "{5C842E5B-1394-E711-80F8-3863BB35AC20}"));
                    bby_referringemployeeid = $env:USERNAME;
                    bby_categoryofsolutions = "Home Theater|Smart Home / Connected Home";
                    bby_resourcerequirementid = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_resourcerequirement", "{7A842E5B-1394-E711-80F8-3863BB35AC20}"));
                    bby_bookableresourcebookingid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bookableresourcebooking", "{CD625027-AC94-E711-80FD-3863BB36EDB8}"));
                    bby_lineofbusiness = $([Microsoft.Xrm.Sdk.EntityReference]::new("bby_lineofbusiness", "{7C150582-BC94-E711-8116-5065F38A2B61}"));
                } -conn $Conn | Out-Null;
    
                Write-Output 1;
            }
    
            Function Add-Chat {
                param
                (
                    [Int]                                      $Iteration
                )
                $IncidentList = $Global:CreatedIncidents.ToArray();
                $IncidentGuid = $IncidentList[$RandomGenerator.Next($IncidentList.Count)];
    
                Set-CreateRequest -EntityLogicalName bby_chat -Fields @{
                    isbilled = $false;
                    bby_waittimeseconds = 29.51;
                    bby_partnerid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bby_partner", "{F90CB08A-9E61-E711-8114-5065F38A7BC1}"));
                    isworkflowcreated = $false;
                    bby_isinitialchat = $false;
                    prioritycode = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(1));
                    actualstart = $(Get-Date("2018-03-28T09:37:00-05:00")).AddHours($Iteration/1000);
                    regardingobjectid = $([Microsoft.Xrm.Sdk.EntityReference]::new("incident", $IncidentGuid)); # link to incidents made above
                    statecode = 1;
                    bby_chatlog = "[{&quot;Content&quot;:&quot;Chat session transfered to agent&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522247824329+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;Client Connected&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522247824256+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;Client Disconnected&quot;,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522248067619+0800)\/&quot;,&quot;Type&quot;:1},{&quot;Content&quot;:&quot;LMI Session: 605291183 has been connected to the Case&quot;,&quot;Direction&quot;:1,&quot;From&quot;:&quot;System&quot;,&quot;Time&quot;:&quot;\/Date(1522251444640+0800)\/&quot;,&quot;Type&quot;:1}]";
                    bby_acknowledgementfromdotcom = $false;
                    activitytypecode = 10384;
                    bby_customeremail = ("{0}@bestbuy-fake-{1}.com" -f $env:USERNAME, $Iteration);
                    deliveryprioritycode = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(1));
                    instancetypecode = 0;
                    ismapiprivate = $false;
                    leftvoicemail = $false;
                    bby_mobilecontext = $false;
                    actualend = $(Get-Date("2018-03-28T10:37:50-05:00")).AddHours($Iteration/1000);
                    bby_customerjabberuserid = ("{0}@bestbuy-fake.com/{1}" -f $env:USERNAME, $Iteration);
                    isregularactivity = 1;
                    statuscode = 2;
                    subject = ("Fulfillment Chat for {0}@bestbuy-fake-{1}.com" -f $env:USERNAME, $Iteration);
                } -conn $Conn | Out-Null;
    
                Write-Output 1;
            }
    
            Function Add-WorkOrder {
                param
                (
                    [Int]                                      $Iteration
                )
    
                $IncidentList = $Global:CreatedIncidents.ToArray();
                $IncidentGuid = $IncidentList[$RandomGenerator.Next($IncidentList.Count)];
    
                Set-CreateRequest -EntityLogicalName msdyn_workorder -Fields @{
                    msdyn_closedby = $([Microsoft.Xrm.Sdk.EntityReference]::new("systemuser", "{2f1a425f-f024-e711-80fc-e0071b6a1041}"));                   #TODO RANDOM
                    statecode = 0;
                    msdyn_longitude = -93.30542;
                    msdyn_name = ("000{0}-{1}-{2}" -f $Iteration, $env:USERNAME, $((Get-Date).ToShortDateString()));
                    msdyn_estimatesubtotalamount = 0.0000;
                    msdyn_postalcode = "55423-8500";
                    msdyn_workordertype = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_workordertype", "9d54fe0a-1281-e711-8113-5065f38a2b61"));
                    msdyn_latitude = 44.86392;
                    msdyn_timeclosed = $(Get-Date("5/3/2018 5:12:21 PM")).AddHours($Iteration/1000);
                    msdyn_worklocation = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(690970000));
                    msdyn_primaryincidenttype = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_incidenttype", "b18f33bf-35f2-e711-812d-e0071b6af0b1"));
                    msdyn_taxable = $false;
                    msdyn_pricelist = $([Microsoft.Xrm.Sdk.EntityReference]::new("pricelevel", "84eb4950-2e49-e611-80d7-c4346bac6974"));
                    msdyn_stateorprovince = "MN";
                    msdyn_address1 = "7601 Penn Ave S";
                    transactioncurrencyid = $([Microsoft.Xrm.Sdk.EntityReference]::new("transactioncurrency", "3b0d0508-cf08-e711-8104-5065f38b5281"));
                    msdyn_billingaccount = $([Microsoft.Xrm.Sdk.EntityReference]::new("account", "0dafcb29-2431-e811-8113-5065f38b42e1"));                  #TODO RANDOM
                    msdyn_customerasset = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_customerasset", "69afccd3-2431-e811-8113-5065f38b42e1"));
                    bby_lineofbusinessid = $([Microsoft.Xrm.Sdk.EntityReference]::new("bby_lineofbusiness", "471384a8-7ace-e711-8125-5065f38a0a21"));
                    msdyn_substatus = $([Microsoft.Xrm.Sdk.EntityReference]::new("msdyn_workordersubstatus", "d985a792-048c-e711-811a-5065f38a4ba1"));
                    msdyn_serviceaccount = $([Microsoft.Xrm.Sdk.EntityReference]::new("account", "0dafcb29-2431-e811-8113-5065f38b42e1"));                  #TODO RANDOM
                    statuscode  = 1;
                    msdyn_ismobile = $false;
                    msdyn_totalsalestax = 0.0000;
                    bby_paymentsettled = $false;
                    msdyn_city = "Richfield";
                    msdyn_subtotalamount = 39.9900;
                    msdyn_servicerequest = $([Microsoft.Xrm.Sdk.EntityReference]::new("incident", $IncidentGuid));
                    msdyn_systemstatus = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(690970004));
                    msdyn_totalamount = 39.9900;
                } -conn $Conn | Out-Null;
    
                Write-Output 1;
            }
    
            $Global:RandomGenerator = [Random]::new();
            $Iteration = $RandomGenerator.Next(1000000000);
            $Global:CreatedIncidents = [System.Collections.Generic.Queue[string]]::new();
            $Global:CreatedIncidents.Enqueue("{49A7B6A8-9132-E811-8113-5065F38B5221}");
            $Global:CreatedChats = [System.Collections.Generic.Queue[string]]::new();
            $Metrics = [PSCustomOBject]@{
                "TimeStamp" = (Get-Date)
                "SecondsToComplete" = 0
                "IncidentTotal" = 0
                "PhoneCallTotal" = 0
                "LeadTotal" = 0
                "OpportunityTotal" = 0
                "WorkOrderTotal" = 0
                "ChatTotal" = 0
                "TotalCreated" = 0
                "TotalUpdated" = 0;
                "Created" = ""
            }
            Write-Host ("{0} {1} {2} {3} {4} {5}" -f $IncidentTotal, $PhoneCallTotal, $LeadTotal, $OpportunityTotal, $WorkOrderTotal, $ChatTotal);
            $TotalRatioCount = $IncidentTotal + $PhoneCallTotal + $LeadTotal + $OpportunityTotal + $WorkOrderTotal + $ChatTotal;
            $Conn = Connect-CrmOnline -Credential $Cred -ServerUrl ("https://{0}.crm.dynamics.com" -f $OrgName);
    
            $Global:request = [Microsoft.Xrm.Sdk.Messages.ExecuteMultipleRequest]::new()
                $settings = [Microsoft.Xrm.Sdk.ExecuteMultipleSettings]::new()
                $settings.ContinueOnError = $true;
                $settings.ReturnResponses = $true;
                $Global:request.Settings = $settings;
                $Global:IncidentCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
                $Global:ChatCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
                $Global:OtherCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
                $Global:ChatUpdateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
    
            While ($true) {
                if ($Global:IncidentCreateRequests.Count -gt 49) {
                    $Metrics.TotalCreated += ($Global:IncidentCreateRequests.Count);
                    $Global:request.Requests = $Global:IncidentCreateRequests;
                    $settings.ReturnResponses = $true;
                    $Global:request.Settings = $settings;
                    $response = $Conn.Execute($request);
                    $Global:IncidentCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
    
                    $Global:CreatedIncidents.Clear();
                    $response.Responses.response.id | ForEach { $Global:CreatedIncidents.Enqueue($_); }
    
                    #Write-Host ($Metrics | Select-Object TimeStamp, IncidentTotal, PhoneCallTotal, LeadTotal, OpportunityTotal, WorkOrderTotal, ChatTotal, TotalCreated, TotalUpdated);
                }
    
                if ($Global:ChatCreateRequests.Count -gt 49) {
                    $Metrics.TotalCreated += ($Global:ChatCreateRequests.Count);
                    $Global:request.Requests = $Global:ChatCreateRequests;
                    $settings.ReturnResponses = $true;
                    $Global:request.Settings = $settings;
                    $response = $Conn.Execute($request);
                    $Global:ChatCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
    
                    # $fetchxml = "<fetch count='10' >  <entity name='bby_chat' >    <attribute name='bby_partnerid' />    <attribute name='activityid' />    <attribute name='bby_channelname' />    <attribute name='bby_isinitialchat' />    <attribute name='bby_partneridname' />    <attribute name='bby_origination' />    <attribute name='bby_joinpin' />    <attribute name='bby_partyid' />    <attribute name='bby_chatlog' />    <attribute name='bby_solutionid' />    <attribute name='bby_mobilecontextname' />    <attribute name='bby_assignedworkforce' />    <attribute name='bby_assignedworkforcename' />    <attribute name='bby_customeremail' />    <attribute name='bby_customerjabberuserid' />    <attribute name='bby_assignedworkforceyominame' />    <attribute name='bby_isinitialchatname' />    <attribute name='bby_origintype' />    <attribute name='bby_solutionname' />    <attribute name='bby_legacypin' />    <attribute name='bby_channel' />    <attribute name='bby_acknowledgementfromdotcomname' />    <attribute name='bby_mobilecontext' />    <attribute name='bby_acknowledgementfromdotcom' />    <attribute name='bby_triagesku' />    <attribute name='bby_legacychatid' />    <filter type='and' >      <condition attribute='bby_chatlog' operator='not-null' />    </filter>    <order attribute='createdon' descending='true' />  </entity></fetch>";
                    # $records = Get-CrmRecordsByFetch -Fetch $fetchxml;
    
                    $Global:CreatedChats.Clear();
                    $response.Responses.response.id | ForEach { $Global:CreatedChats.Enqueue($_); }
    
                    #Write-Host ($Metrics | Select-Object TimeStamp, IncidentTotal, PhoneCallTotal, LeadTotal, OpportunityTotal, WorkOrderTotal, ChatTotal, TotalCreated, TotalUpdated);
                }
    
                if ($Global:CreatedChats.Count -gt 24) {
                    While ($Global:CreatedChats.Count -gt 0) {
                        $ChatGuid = $Global:CreatedChats.Dequeue();
                        For ($i=0; $i -le 10; $i++) {
                            $Global:ChatUpdateRequests.Add($(Get-ChatUpdateRequest -Id $ChatGuid -Iteration $Iteration));
                        }
                    }
    
                    $Metrics.TotalUpdated += ($Global:ChatUpdateRequests.Count);
                    $Global:request.Requests = $Global:ChatUpdateRequests;
                    $settings.ReturnResponses = $false;
                    $Global:request.Settings = $settings;
                    $response = $Conn.Execute($request);
                    $Global:ChatUpdateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
    
                    #Write-Host ($Metrics | Select-Object TimeStamp, IncidentTotal, PhoneCallTotal, LeadTotal, OpportunityTotal, WorkOrderTotal, ChatTotal, TotalCreated, TotalUpdated);
                }
    
                if ($Global:OtherCreateRequests.Count -gt 49) {
                    $Metrics.TotalCreated += ($Global:OtherCreateRequests.Count);
                    $Global:request.Requests = $Global:OtherCreateRequests;
                    $settings.ReturnResponses = $false;
                    $Global:request.Settings = $settings;
                    $response = $Conn.Execute($request);
                    $Global:OtherCreateRequests = [Microsoft.Xrm.Sdk.OrganizationRequestCollection]::new();
    
                    #Write-Host ($Metrics | Select-Object TimeStamp, IncidentTotal, PhoneCallTotal, LeadTotal, OpportunityTotal, WorkOrderTotal, ChatTotal, TotalCreated, TotalUpdated);
                }
    
                $Metrics.TimeStamp = (Get-Date);
                $CurrentRunIndex = $RandomGenerator.Next($TotalRatioCount);
                switch ($CurrentRunIndex) {
                    { $CurrentRunIndex -lt $IncidentTotal } { 
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.IncidentTotal += Add-Incident -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "Incident";
                        break;
                    }
                    { $CurrentRunIndex -lt ($IncidentTotal + $PhoneCallTotal) } {
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.PhoneCallTotal += Add-PhoneCall -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "PhoneCall";
                        break;
                    }
                    { $CurrentRunIndex -lt ($IncidentTotal + $PhoneCallTotal + $LeadTotal) } {
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.LeadTotal += Add-Lead -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "Lead";
                        break;
                    }
                    { $CurrentRunIndex -lt ($IncidentTotal + $PhoneCallTotal + $LeadTotal + $OpportunityTotal) } {
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.OpportunityTotal += Add-Opportunity -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "Opportunity";
                        break;
                    }
                    { $CurrentRunIndex -lt ($IncidentTotal + $PhoneCallTotal + $LeadTotal + $OpportunityTotal + $WorkOrderTotal) } {
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.WorkOrderTotal += Add-WorkOrder -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "WorkOrder";
                        break;
                    }
                    Default {
                        $Metrics.SecondsToComplete = Measure-Command { $Metrics.ChatTotal += Add-Chat -Iteration $Iteration; } | Select -ExpandProperty TotalSeconds
                        $Metrics.Created = "Chat";
                    }
                }
    
                $Iteration += 1;
                $Metrics | Select TimeStamp, SecondsToComplete, TotalCreated, TotalUpdated, Created | Export-Csv -NoTypeInformation -Append -Path $LogFilePath;
            }
        }

    Write-Output $Script;
}

if ($AutoConfigure) {
    if (-not $Cred) {
        Write-Host "Requesting Credentials...";
        $Cred = Get-Credential;
    }

    if (-not $LogFilePath) {
        Write-Host "Setting Log Path...";
        $LogFilePath = $("{0}\Desktop\PerformanceLog.csv" -f $env:USERPROFILE);
    }

    if (-not $OrgName) {
        Write-Host "Setting Org Name...";
        $OrgName = "bbyb2cdryrun07";
    }

    if (-not $IncidentTotal) {
        Write-Host "Setting ratios based on sample month growth..."
        $IncidentTotal = 210;
        $PhoneCallTotal = 239;
        $LeadTotal = 43;
        $OpportunityTotal = 36;
        $WorkOrderTotal = 116;
        $ChatTotal = 354;
    }

    Write-Host "Initializing Primary Script...";
    Get-Performance `
        -OrgName $OrgName `
        -Cred $Cred `
        -LogFilePath $LogFilePath `
            -IncidentTotal $IncidentTotal `
            -PhoneCallTotal $PhoneCallTotal `
            -LeadTotal $LeadTotal `
            -OpportunityTotal $OpportunityTotal `
            -WorkOrderTotal $WorkOrderTotal `
            -ChatTotal $ChatTotal;

}

