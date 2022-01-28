function Get-SPListEntries{
<#
.SYNOPSIS
    Retrieve all entries in a Sharepoint list
.DESCRIPTION
    Specify a Sharepoint site and sharepoint list and return all entires in the list.
.PARAMETER WebURL
    Specify the Sharepoint site's URL
.PARAMETER ListName
    Specify the Sharepoint list name to pull data from
.NOTES
    Version:        1.0
    Author:         disposablecat
    Purpose/Change: Initial script development
.EXAMPLE
    Get-SPListEntries -WebURL https://sharepoint.company.com/subsite -ListName ServerInventory
    Retrieves a the Sharepoint list "ServerInventory"
#>
    [CmdletBinding()]
    [OutputType('System.Collections.Generic.List[System.Object]')]
    
    #Define parameters
    Param
    (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$WebURL = "https://collaboration.ucf.edu/sites/itr/cst/itinfrastructure",

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$ListName = "ServerInventory"
    )

    Begin
    {   
        #construct complete URL based on sute and list name   
        $endpointUrl = "$WebUrl/_vti_bin/listdata.svc/$ListName"
        $Xml = New-Object System.Xml.XmlDocument
        $Results = New-Object System.Collections.Generic.List[System.Object]
        $BaseEntryObject = New-Object PSCustomObject
        $i = 0    
    }
    Process
    {
        Try
        {
            [xml]$Xml = Invoke-WebRequest -Uri $endpointUrl -Method Get -UseDefaultCredentials -ContentType "application/xml" -ErrorAction Stop
            #Loop to iterate through first entry and map properties to custom object
            ForEach($ChildNode in $Xml.feed.entry[0].content.properties.ChildNodes)
            {
                $BaseEntryObject | Add-Member -type NoteProperty -Name $ChildNode.LocalName -Value $null
            }
            #Loop to catch lists over 1000 (limite of Sharepoint)
            While ($i -lt 1)
            {
            $i++
            #Loop Through all list items in this batch
                Foreach($Entries in $Xml.feed.entry.content.properties)
                {
                    #Set temporary EntryObject
                    $EntryObject = $BaseEntryObject | Select *
                    #Loop through all properties for specific list item           
                    ForEach($Entry in $Entries.ChildNodes)
                    {
                        $TempEntryLocalName = $Entry.LocalName
                        #Assign Value to dynamically created property
                        $EntryObject."$TempEntryLocalName" = $Entry.InnerText
                    }
                    #Add all properties from list item to collection
                    $Results.Add($EntryObject)
                }
                #Test to see if there is a link to next batch of entries (Sharepoint lists over 1000)
                if($xml.feed.link | Where rel -eq "next")
                {
                    #Reduce number to continue while loop
                    $i--
                    #Grab next XML to process
                    $endpointUrl = $xml.feed.link | Where rel -eq "next" | Select -ExpandProperty href
                    #Grab xml
                    [xml]$Xml = Invoke-WebRequest -Uri $endpointUrl -Method Get -UseDefaultCredentials -ContentType "application/xml"
                }
            }
        }
        Catch
        {
            #Catch any error.
            Write-Verbose “Exception Caught”
            Write-Verbose “Exception Type: $($_.Exception.GetType().FullName)”
            Write-Verbose “Exception Message: $($_.Exception.Message)”
        }
        #return results
        return $Results

    }
    End
    {
        #Will execute last. Will execute once. Good for cleanup. 
    }
}

function Get-SPListEntry{
<#
.SYNOPSIS
    Retrieve a single entry in a Sharepoint list
.DESCRIPTION
    Specify a Sharepoint site and sharepoint list and return a single entry from the list.
.PARAMETER WebURL
    Specify the Sharepoint site's URL
.PARAMETER ListName
    Specify the Sharepoint list name to pull data from
.PARAMETER ItemID
    Specify the list items's ID 
.NOTES
    Version:        1.0
    Author:         disposablecat
    Purpose/Change: Initial script development
.EXAMPLE
    Get-SPListEntries -WebURL https://sharepoint.company.com/subsite -ListName ServerInventory -ItemID 22
    Retrieves the Sharepoint list entry with ID 22 from the list "ServerInventory"
#>
    [CmdletBinding()]
    [OutputType('System.Collections.Generic.List[System.Object]')]
    
    #Define parameters
    Param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$WebURL,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$ListName,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [int]$ItemID
    )

    Begin
    {   
        #construct complete URL based on sute and list name   
        $endpointUrl = "$WebUrl/_vti_bin/listdata.svc/$ListName($ItemID)"
        $Xml = New-Object System.Xml.XmlDocument
        $Results = New-Object System.Collections.Generic.List[System.Object]
        $BaseEntryObject = New-Object PSCustomObject
        $i = 0    
    }
    Process
    {
        Try
        {
            [xml]$Xml = Invoke-WebRequest -Uri $endpointUrl -Method Get -UseDefaultCredentials -ContentType "application/xml" -ErrorAction Stop
            #Loop to iterate through first entry and map properties to custom object
            ForEach($ChildNode in $Xml.entry.content.properties.ChildNodes)
            {
                $BaseEntryObject | Add-Member -type NoteProperty -Name $ChildNode.LocalName -Value $null
            }
           
            Foreach($Entries in $Xml.entry.content.properties)
            {
                #Set temporary EntryObject
                $EntryObject = $BaseEntryObject | Select *
                #Loop through all properties for specific list item           
                ForEach($Entry in $Entries.ChildNodes)
                {
                    $TempEntryLocalName = $Entry.LocalName
                    #Assign Value to dynamically created property
                    $EntryObject."$TempEntryLocalName" = $Entry.InnerText
                }
                #Add all properties from list item to collection
                $Results.Add($EntryObject)
            }
                
        }
        Catch
        {
            #Catch any error.
            Write-Verbose “Exception Caught”
            Write-Verbose “Exception Type: $($_.Exception.GetType().FullName)”
            Write-Verbose “Exception Message: $($_.Exception.Message)”
        }
        #return results
        return $Results

    }
    End
    {
        #Will execute last. Will execute once. Good for cleanup. 
    }
}

function Get-SPLists{
<#
.SYNOPSIS
    Retrieve all lists from a sharepoint site
.DESCRIPTION
    Specify a Sharepoint site and return all lists on that site. Usefull for discovering lists to use with other funcations/cmdlets.
.PARAMETER WebURL
    Specify the Sharepoint site's URL
.NOTES
    Version:        1.0
    Author:         disposablecat
    Purpose/Change: Initial script development
.EXAMPLE
    Get-SPLists -WebURL https://sharepoint.company.com/subsite
    Retrieves all lists from the sharepoint site https://sharepoint.company.com/subsite
#>
    [CmdletBinding()]
    [OutputType('System.Collections.Generic.List[System.Object]')]
    
    #Define parameters
    Param
    (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$WebURL = "https://collaboration.ucf.edu/sites/itr/cst/itinfrastructure"

    )

    Begin
    {   
        #construct complete URL based on sute and list name   
        $endpointUrl = "$WebUrl/_vti_bin/listdata.svc"
        $Xml = New-Object System.Xml.XmlDocument
        $Results = New-Object System.Collections.Generic.List[System.Object]
        $BaseEntryObject = New-Object PSCustomObject
        $BaseEntryObject | Add-Member -Type NoteProperty -Name ListName -Value $null   
    }
    Process
    {
        Try
        {
            [xml]$Xml = Invoke-WebRequest -Uri $endpointUrl -Method Get -UseDefaultCredentials -ContentType "application/xml" -ErrorAction Stop
            #Loop Through all list names
            Foreach($Entry in $Xml.service.workspace.collection)
            {
                #Set temporary EntryObject
                $EntryObject = $BaseEntryObject | Select *
                #Assign value to list name
                $EntryObject.ListName = $Entry.InnerText
                #Add list enty into list
                $Results.Add($EntryObject)
            }     
        }
        Catch
        {
            #Catch any error.
            Write-Verbose “Exception Caught”
            Write-Verbose “Exception Type: $($_.Exception.GetType().FullName)”
            Write-Verbose “Exception Message: $($_.Exception.Message)”
        }
        #return results
        return $Results

    }
    End
    {
        #Will execute last. Will execute once. Good for cleanup. 
    }
}

function Set-SPListEntries{
<#
.SYNOPSIS
    Change or set a list entry value
.DESCRIPTION
    Specify a Sharepoint site and list name and set/change a value in it
.PARAMETER WebURL
    Specify the Sharepoint site's URL
.PARAMETER ListName
    Specify the Sharepoint list name to set data in
.PARAMETER ItemID
    Specify the list items's ID
.PARAMETER Properties
    Specify the properties to set/change in a hashtable
.NOTES
    Version:        1.0
    Author:         disposablecat
    Purpose/Change: Initial script development
.EXAMPLE
    Set-SPListEntries -WebURL https://sharepoint.company.com/subsite -ListName List1 -ItemID 22 -PropertyName Name -PropertyValue "Joe Smith"
    Sets the "Name" field to "Joe Smith" for list item "22" in the list "List1" on the site "https://sharepoint.company.com/subsite"
#>
    [CmdletBinding()]
    [OutputType('System.Collections.Generic.List[System.Object]')]
    
    #Define parameters
    Param
    (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$WebURL = "https://collaboration.ucf.edu/sites/itr/cst/itinfrastructure",

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$ListName = "ServerInventory",

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [int]$ItemID,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [hashtable]$Properties,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]$Credential

    )

    Begin
    {   
        #construct complete URL based on sute and list name   
        $endpointUrl = "$WebUrl/_vti_bin/listdata.svc/$ListName($ItemID)"
        $header = @{
            "X-HTTP-Method" = "MERGE";
            "If-Match" = "*"
        }
        $Payload = $Properties | ConvertTo-Json
        $EntryObject = New-Object PSCustomObject
        $EntryObject | Add-Member -Type NoteProperty -Name WebURL -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name ListName -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name Properties -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name Status -Value $null 
    }
    Process
    {
        Try
        {
            $EntryObject.WebURL = $WebURL
            $EntryObject.ListName = $ListName
            $EntryObject.Properties = $Properties
            if($Credential -eq $false)
            {
                $Response = Invoke-WebRequest -Uri $endpointUrl -Method POST -UseDefaultCredentials -Headers $header -ContentType "application/json;odata=verbose" -Body $Payload -ErrorAction Stop
            }
            else
            {
                $Response = Invoke-WebRequest -Uri $endpointUrl -Method POST -Headers $header -ContentType "application/json;odata=verbose" -Body $Payload -Credential (Get-Credential) -ErrorAction Stop
            }
            If ($Response.StatusCode -like "2*")
            {
                $EntryObject.Status = "Success"
                Write-Verbose "Success"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
            elseif (($Response.StatusCode -like "4*") -or ($Response.StatusCode -like "5*") -or ($Response.StatusCode -like "3*"))
            {
                $EntryObject.Status = "Failed"
                Write-Verbose "Failed"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
            else
            {
                $EntryObject.Status = "Unknown"
                Write-Verbose "Unknown"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
        }
        Catch
        {
            #Catch any error.
            Write-Verbose “Exception Caught”
            Write-Verbose “Exception Type: $($_.Exception.GetType().FullName)”
            Write-Verbose “Exception Message: $($_.Exception.Message)”
        }
        #return results
        return $EntryObject

    }
    End
    {
        #Will execute last. Will execute once. Good for cleanup. 
    }
}

function New-SPListEntry{
<#
.SYNOPSIS
    Create a new Sharepoint List Entry
.DESCRIPTION
    Create a new Sharepoint list entry in the specified Sharepoint list.
.PARAMETER WebURL
    Specify the Sharepoint site's URL
.PARAMETER ListName
    Specify the Sharepoint list name to set data in
.PARAMETER Properties
    Specify the properties to set/change in hashtable form
.PARAMETER Credential
    Specify an alternative set of credentails, otherwise the currently logged in user will be used
.NOTES
    Version:        1.0
    Author:         disposablecat
    Purpose/Change: Initial script development
.EXAMPLE
    New-SPListEntry -WebURL https://sharepoint.company.com/subsite -ListName List1 -Properties @{Name="Joe Smith";Department="HR";Status="Active"}
    Creates a new list entry in the "List1" list on the Sharepoint site "https://sharepoint.company.com/subsite"
#>
    [CmdletBinding()]
    [OutputType('System.Collections.Generic.List[System.Object]')]
    
    #Define parameters
    Param
    (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$WebURL = "https://collaboration.ucf.edu/sites/itr/cst/itinfrastructure",

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]$ListName = "ServerInventory",

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [hashtable]$Properties,

        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]$Credential

    )

    Begin
    {   
        #construct complete URL based on sute and list name   
        $endpointUrl = "$WebUrl/_vti_bin/listdata.svc/$ListName"
        $header = @{
            "Accept" = "application/json;odata=verbose";
        }
        $Payload = $Properties | ConvertTo-Json
        $EntryObject = New-Object PSCustomObject
        $EntryObject | Add-Member -Type NoteProperty -Name WebURL -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name ListName -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name Properties -Value $null
        $EntryObject | Add-Member -Type NoteProperty -Name Status -Value $null 
    }
    Process
    {
        Try
        {
            $EntryObject.WebURL = $WebURL
            $EntryObject.ListName = $ListName
            $EntryObject.Properties = $Properties
            if($Credential -eq $false)
            {
                $Response = Invoke-WebRequest -Uri $endpointUrl -Method POST -UseDefaultCredentials -Headers $header -ContentType "application/json;odata=verbose" -Body $Payload -ErrorAction Stop
            }
            else
            {
                $Response = Invoke-WebRequest -Uri $endpointUrl -Method POST -Headers $header -ContentType "application/json;odata=verbose" -Body $Payload -Credential (Get-Credential) -ErrorAction Stop
            }
            If ($Response.StatusCode -like "2*")
            {
                $EntryObject.Status = "Success"
                Write-Verbose "Success"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
            elseif (($Response.StatusCode -like "4*") -or ($Response.StatusCode -like "5*") -or ($Response.StatusCode -like "3*"))
            {
                $EntryObject.Status = "Failed"
                Write-Verbose "Failed"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
            else
            {
                $EntryObject.Status = "Unknown"
                Write-Verbose "Unknown"
                Write-Verbose "Status Code: $($Response.StatusCode)"
                Write-Verbose "Status Message: $($Response.StatusDescription)"
            }
            
        }
        Catch
        {
            #Catch any error.
            Write-Verbose “Exception Caught”
            Write-Verbose “Exception Type: $($_.Exception.GetType().FullName)”
            Write-Verbose “Exception Message: $($_.Exception.Message)”
        }
        #return results
        return $EntryObject

    }
    End
    {
        #Will execute last. Will execute once. Good for cleanup. 
    }
}