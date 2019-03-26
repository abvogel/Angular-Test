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

    $RandomGenerator = [Random]::new();
    $Iteration = $RandomGenerator.Next(1000000000);
    $Global:RecentIncidents = [System.Collections.Generic.Queue[string]]::new();
    $Global:RecentIncidents.Enqueue("{49A7B6A8-9132-E811-8113-5065F38B5221}");
    $Global:Metrics = [PSCustomOBject]@{
        "TimeStamp" = (Get-Date)
        "SecondsToComplete" = 0
        "IncidentTotal" = 0
        "PhoneCallTotal" = 0
        "LeadTotal" = 0
        "OpportunityTotal" = 0
        "WorkOrderTotal" = 0
        "ChatTotal" = 0
	"LastGuid" = "0"
        "Created" = ""
    }
    Write-Host ("{0} {1} {2} {3} {4} {5}" -f $IncidentTotal, $PhoneCallTotal, $LeadTotal, $OpportunityTotal, $WorkOrderTotal, $ChatTotal);
    $TotalRatioCount = $IncidentTotal + $PhoneCallTotal + $LeadTotal + $OpportunityTotal + $WorkOrderTotal + $ChatTotal;
    $Conn = Connect-CrmOnline -Credential $Cred -ServerUrl ("https://{0}.crm.dynamics.com" -f $OrgName);
    
    While ($true) {
	Start-Sleep -Milliseconds 1000;
        $Metrics.TimeStamp = (Get-Date);
        $CurrentRunIndex = $RandomGenerator.Next($TotalRatioCount);
        #Write-Host ("Next Run Index: {0} / {1}" -f $CurrentRunIndex, $TotalRatioCount);
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
        $Metrics | Export-Csv -NoTypeInformation -Append -Path $LogFilePath;
        Write-Host $Metrics;
    }
}

Function Add-Incident {
    param
    (
        [Int]                                      $Iteration
    )
    
    $NewGuid = New-CrmRecord -EntityLogicalName incident -Fields @{
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

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    if ($Global:RecentIncidents.Count -gt 10) {
        $Global:RecentIncidents.Dequeue | Out-Null;
    }

    $Global:RecentIncidents.Enqueue($NewGuid.Guid);

    Write-Output 1;
}

Function Add-PhoneCall {
    param
    (
        [Int]                                      $Iteration
    )

    $NewGuid = New-CrmRecord -EntityLogicalName phonecall -Fields @{
        phonenumber = "(123) 123-1234";
        description = $("Confirmed 12-545 on ETA 729 that Iteration {0} meets expected specs." -f $Iteration);
        subject=$("{0} hour call - it's a callathon!" -f $Iteration)
        regardingobjectid = $([Microsoft.Xrm.Sdk.EntityReference]::new("contact", "{8F8AA533-99B2-E811-8128-5065F38B91F1}"));
        bby_calltype = $([Microsoft.Xrm.Sdk.OptionSetValue]::new(864630010));
    } -conn $Conn;

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    Write-Output 1;
}

Function Add-Lead {
    param
    (
        [Int]                                      $Iteration
    )

    $NewGuid = New-CrmRecord -EntityLogicalName lead -Fields @{
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
    } -conn $Conn;

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    Write-Output 1;
}

Function Add-Opportunity {
    param
    (
        [Int]                                      $Iteration
    )

    $NewGuid = New-CrmRecord -EntityLogicalName opportunity -Fields @{
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
    } -conn $Conn;

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    Write-Output 1;
}

Function Add-Chat {
    param
    (
        [Int]                                      $Iteration
    )

    $IncidentList = $Global:RecentIncidents.ToArray();
    $IncidentGuid = $IncidentList[$RandomGenerator.Next($IncidentList.Count)];

    $NewGuid = New-CrmRecord -EntityLogicalName bby_chat -Fields @{
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
    } -conn $Conn;

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    Write-Output 1;
}

Function Add-WorkOrder {
    param
    (
        [Int]                                      $Iteration
    )

    $IncidentList = $Global:RecentIncidents.ToArray();
    $IncidentGuid = $IncidentList[$RandomGenerator.Next($IncidentList.Count)];

    $NewGuid = New-CrmRecord -EntityLogicalName msdyn_workorder -Fields @{
        msdyn_closedby = $([Microsoft.Xrm.Sdk.EntityReference]::new("systemuser", "{2f1a425f-f024-e711-80fc-e0071b6a1041}"));                   #TODO RANDOM
        statecode = 0;
        msdyn_longitude = -93.30542;
        msdyn_name = ("000{0}" -f $Iteration);
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

    $Global:Metrics.LastGuid = ($NewGuid.Guid);

    Write-Output 1;
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
        $WorkOrderTotal = 0; #116;
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