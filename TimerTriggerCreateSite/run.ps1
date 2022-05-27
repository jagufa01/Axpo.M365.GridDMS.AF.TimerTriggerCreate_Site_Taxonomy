###########################################################
#  Aveniq AG, Bruggerstrasse 68, 5400 Baden
#  james.gurtner@aveniq.ch
#
#  17.05.2022
#
#  Das Script erstellt bestellte Räume, welche in der List 
#  "Site Orders" aufgeführt sind.
#
###########################################################
using namespace System.Net

# Input bindings are passed in via param block.
param($Timer)

$systemDMS = $env:AxpoGridDMSystem
Write-Host $systemDMS

################# CONNECTION SETTINGS ####################
#Connection Settings
$systemDMS= "LAB2" #initial setting; just for test, delete/remark after tests
$tenantId = "8619c67c-945a-48ae-8e77-35b1b71c9b98"
$appClientId = "ebc8f19d-604f-4303-8bb2-de62aba9119d" # Reg App -> cl-axsa-az-appl-griddms-nonprod-scheduler
$certPath = "C:\home\site\wwwroot\axsa4shrd4np4griddms4kv-gridDMS-scheduler-ad-app-20220308.pfx"
$adminSPUrl = "https://axpogrp-admin.sharepoint.com/"
################# END CONNECTION SETTINGS ################

################# FUNCTIONS  ####################
function checkEnviromentDMSSystem()
{
    if($systemDMS)
    {
        Write-Host "Umgebungsvariable ist vorhanden."
    }
    else
    {
        Write-Host "Umgebungsvariable ist nicht vorhanden."
    }
}

function checkDownloadFolder()
{
    #if not exist create download folder
    if(Test-Path -Path $download) {
        Write-Host "Folder download exist.."
    }
    else{
        New-Item -Type "Directory" -Name "download" -Path $homedrive
        Write-Host "Create folder download.."
    }
}

function CreateNewSite($adminUrl, $siteTitle, $siteAlias, $siteHubSiteId)
{
        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Write-Host "Connect to " $adminUrl

        $newSite = New-PnPSite -Type TeamSite -Title $siteTitle -Alias $siteAlias -HubSiteId $siteHubSiteId -Lcid 1033
        #$newSite = "https://axpogrp.sharepoint.com/sites/appllabaxedmptr02529065"
        Write-Host "New Sharepoint Site is  $newSite  created."
        Get-PnPContext
        Exit

        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPSite -Identity $newSite  -DenyAndAddCustomizePages $false

        Connect-PnPOnline -Url $newSite -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

        #Wait until TaxonomyHiddenListe is created
        $maxWaitingCycle = 0
        $hiddenList = $null

        DO
        {
            Write-Host "Waiting"
            start-sleep 10
            $maxWaitingCycl += 1

            $hiddenListlist = Get-PnPList -Identity "TaxonomyHiddenList"
        } Until ($maxWaitingCycl -gt 50 -or ($hiddenListlist -ne $null))
        #end wait

        #Check if is Unterwerk, Trasse, Projekt 
        if($templateName -eq "Unterwerk"){
            Get-PnPFile -url $tempListFolderUnterwerk -Path $download -FileName "PnP-Provisioning-Site-Template-Unterwerk.xml" -AsFile -Force
            #Invoke-PnPSiteTemplate -Path "C:\__Axpo\Temp\PnP-Provisioning-Site-Template-Unterwerk.xml"

            #Set PropertyBagValue
            Set-PnPPropertyBagValue -Key "axeSiteType" -Value $templateName
        }
        elseif($templateName -eq "Trassee"){
            Get-PnPFile -url $tempListFolderTrassee -Path $download -FileName "PnP-Provisioning-Site-Template-Trassee.xml" -AsFile -Force
            #Invoke-PnPSiteTemplate -Path "C:\__Axpo\Temp\PnP-Provisioning-Site-Template-Trassee.xml"
            
            #Set PropertyBagValue
            Set-PnPPropertyBagValue -Key "axeSiteType" -Value $templateName
        }
        elseif($templateName -eq "Projekt"){
            Get-PnPFile -url $tempListFolderProjekt -Path $download -FileName "PnP-Provisioning-Site-Template-Projekt.xml" -AsFile -Force
            #Invoke-PnPSiteTemplate -Path "C:\__Axpo\Temp\PnP-Provisioning-Site-Template-Projekt.xml"

            #Set PropertyBagValue
            Set-PnPPropertyBagValue -Key "axeSiteType" -Value $templateName
        }
        else{
            #do nothing
            Write-Host "do nothing PROD"
        }

        #Invoke to new site
        Connect-PnpOnline -Url $adminUrl -Interactive
        Set-PnPSite -Identity $newSite  -DenyAndAddCustomizePages $true

        #Disconnect-PnPOnline
        return $newSite
}

#Set propertybag values for DocIdPrefix
function setPropertyBagDocIdPrefix($adminUrl, $newUrl)
{
        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPSite -Identity $newUrl  -DenyAndAddCustomizePages $false

        Connect-PnPOnline -Url $newUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPPropertyBagValue -key "docid_msft_hier_siteprefix" -value "AXE01" 

        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPSite -Identity $newUrl -DenyAndAddCustomizePages $true

        Connect-PnPOnline -Url $newUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
}

#Set propertybag values axeTPLevels
function setPropertyBag($adminUrl, $newUrl)
{
        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPSite -Identity $newUrl  -DenyAndAddCustomizePages $false

        Connect-PnPOnline -Url $newUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
     
        Set-PnPPropertyBagValue -key "axeTPLevel1" -value $tpLevel1Value -Indexed
        Set-PnPPropertyBagValue -key "axeTPLevel1Des" -value $tpLevel1DesValue -Indexed
        Set-PnPPropertyBagValue -key "axeTPLevel1ID " -value $tpLevel1IDValue -Indexed
        Set-PnPPropertyBagValue -key "axeSiteType" -value $templateName -Indexed

        Connect-PnpOnline -Url $adminUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        Set-PnPSite -Identity $newUrl -DenyAndAddCustomizePages $true

        Connect-PnPOnline -Url $newUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        
        #disable social things
        Set-PnPSite -SocialBarOnSitePagesDisabled $true
}

#Create additional groups
function createGroup($url, $groupName) 
{
        Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

        $group = $groupName + " Members ohne Löschen"
        $groupadmin = $groupName + " Admin"

        New-PnPSiteGroup -Group $group -PermissionLevels "Mitwirken ohne Löschen"
        New-PnPSiteGroup -Group $groupAdmin -PermissionLevels "Full Control"
}

function setPermissionLevel()
{

}

#Set default values to the document metadata fields
function SetListDefaultValue($url, $tpLevel1, $tpLevel1ID, $tpLevel1Des, $siteType)
{
        Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

        $lists = Get-PnPList | where {$_.BaseTemplate -eq "101"}

        foreach($list in $lists)
        {
            $contentTypes = Get-PnPContentType -List $list

            foreach($contentType in $contentTypes)
            {
                if($contentType.Name -eq "AXE DMS Dokument")# -or $contentType.Name -eq "AXE DMS Dokument Set")
                {
                    $list.Title  + "  " + $contentType.Name

                    #Set default Value to  text field and set to hidden
                    Set-PnPField -List $list -Identity "axeTPLevel1ID" -Values @{DefaultValue = $tpLevel1ID; Hidden=$True }
                    Set-PnPField -List $list -Identity "axeTPLevel1Des" -Values @{DefaultValue = $tpLevel1Des; Hidden=$True }
                
                    #Set default value to taxonomie fields
                    Set-PnPDefaultColumnValues -List $list -Field "axeTP" -Value ("AXE DMS|Technischer Platz|" + $tpLevel1)
                    Set-PnPDefaultColumnValues -List $list -Field "axeTPLevel1" -Value ("AXE DMS|Technischer Platz|" + $tpLevel1)
                    Set-PnPDefaultColumnValues -List $list -Field "axeSiteType" -Value ("AXE DMS|Seitentyp|" + $siteType)
               
                    #Set taxonomie fields to hidden
                    Set-PnPField -List $list -Identity "axeSiteType" -Values @{ Hidden=$True }
                    Set-PnPField -List $list -Identity "axeTPLevel1" -Values @{ Hidden=$True }
                    
                    #Set other settings to fields 
                    Set-PnPField -List $list -Identity "axeTPaltern" -Values @{ AllowMultipleValues = $true }

                    #Set Display Format on TP and Dok-Nr Field
                    Set-PnPField -List $list -Identity "axeDocType" -Values @{IsPathRendered = $true }
                    #Set-PnPField -List $list -Identity "axeDocNu" -Values @{IsPathRendered = $true }
                     
                    Remove-PnPContentTypeFromList -List $list -ContentType "Document"
                    Remove-PnPContentTypeFromList -List $list -ContentType "Dokument"

                    Remove-PnPView -List $list -Identity "All Documents" -Force
                }
            }
        }
}

#Call example CreateFolderFromCSV $siteConfig "https://axpogrp.sharepoint.com/sites/appllabaxedmuUnterwerk900177" 
function CreateFolderFromCSV($siteConfigUrl, $url)
{
    Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
    $siteType = Get-PnPPropertyBag -Key "axeSiteType"

    Connect-PnPOnline -Url $siteConfigUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

    if($siteType -eq "Unterwerk")
    {
        $file = Get-PnPFile -Url $tempListFolderUnterwerk -AsString
    }
    elseif($siteType -eq "Trassee")
    {
        $file = Get-PnPFile -Url $tempListFolderTrassee -AsString
    }
    elseif($siteType -eq "Projekt")
    {
        $file = Get-PnPFile -Url $tempListFolderProjekt -AsString
    }
    else
    {
        $file = $null
    }
    
    #Add to a array
    $dataFile = $file.Split("`n")

    if($file -ne $null)
    {
        Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

        $site = Get-PnPSite
        $web = Get-PnPWeb

        #"C:\__Axpo\Temp\_PnP-Provisioning-ListFolderSchema-Template-Unterwerk.csv"
        foreach($line in $dataFile) 
        {
            $data = $null
            $data = $line.Split(";")
            #$data[0]
            #$data[1]
            #$site.Url

            if($data[1] -ne $null)
            {
                $newFolder = $data[1].Replace("[SITE_REL_URL]", $web.ServerRelativeUrl)
                $data[0]
                $data[1]
                $newFolder         
                "---------"
                $newFolder = $newFolder.TrimEnd()

                Add-PnPFolder -Name $data[0] -Folder $newFolder
            }
        }
    }
    else
    {
        #No data
    }
}

#Copy search file from template library
function copySearchPage($scrUrl, $url)
{
        Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
        $web = Get-PnPWeb
        $tmpRelUrl = $web.ServerRelativeUrl +"/sitepages/" +$tempListSearchPageName
        
        Copy-PnPFile -SourceUrl $scrUrl -TargetUrl ($web.ServerRelativeUrl +"/sitepages") -Force -Overwrite
        start-sleep 3 #to fast for next command
        Rename-PnPFile -ServerRelativeUrl $tmpRelUrl -TargetFileName search.aspx -Force -OverwriteIfAlreadyExists
}

#Publish the copied search aspx to full version
function publishSearchFile($url){
    #Connect to PnP Online
    Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId
     
    #Get all files from the document library
    $ListItems = Get-PnPListItem -List "SitePages" -PageSize 20| Where { $_.FileSystemObjectType -eq "File" }
 
    #Iterate through each file
    ForEach ($Item in $ListItems)
    {
        #Get the File from List Item
        $File = Get-PnPProperty -ClientObject $Item -Property File
 
        #Check if file draft (Minor version)
        If($File.CheckOutType -eq "None" -and $File.MinorVersion)
        {
            $File.Publish("Major version Published by Script")
            $File.Update()
            Invoke-PnPQuery
            Write-host -f Green "Published file at '$($File.ServerRelativeUrl)'"
        }
    }
}

#Remove dublicated nodes  
function DeleteNavNodes($url)
{
    Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

    foreach($item in Get-PnPNavigationNode)
    {
       $nodes = Get-PnPNavigationNode | Where-Object {$_.Title -eq $item.Title}
   
       if($nodes.count -gt 1)
       {
          write-host "Delete dublicate navigation nodes" -BackgroundColor Green

          Remove-PnPNavigationNode -id $nodes[1].Id -Force
       } 
    }

    #Other Element to remove
    Remove-PnPNavigationNode -id 1034 -Force
    Remove-PnPNavigationNode -id 2002 -Force
    Remove-PnPNavigationNode -id 2004 -Force
}

#Disable/hide comment and social features on search page (site)
function additionalSiteSettings($url){
    Connect-PnPOnline -Url $url -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

    Set-PnPSite -Identity $url -CommentsOnSitePagesDisabled $true
    Set-PnPSite -Identity $url  -SocialBarOnSitePagesDisabled $true 
}

function getDefaultTPValueTax($nameTax)
{
    $idTax = Get-PnPTerm -TermGroup "AXE DMS" -TermSet "Technischer Platz" | Where-Object {$_.Name -eq $nameTax} | select id
    $nameTaxNew = "-1;#" + $nameTax + "|" + $idTax.id

    $nameTaxNew
    #Set-PnPField -List "00 Anlagendokumente" -Identity "axeTPLevel1" -Values @{DefaultValue = ($nameTaxNew) }
}

function getDefaultSiteValueTax($nameTax)
{
    $idTax = Get-PnPTerm -TermGroup "AXE DMS" -TermSet "SeitenTyp" | Where-Object {$_.Name -eq $nameTax} | select id
    $nameTaxNew = "-1;#" + $nameTax + "|" + $idTax.id

    $nameTaxNew
    #Set-PnPField -List "00 Anlagendokumente" -Identity "axeTPLevel1" -Values @{DefaultValue = ($nameTaxNew) }
}

################# END FUNCTIONS  ################

################# MAIN  ####################
# varibalen
############################################
Write-Host "Start main program"
$homedrive = "C:\home"
$download = $homedrive+"\download"
$strRootSite = ""
$adminDMUrl =  $null

#Check enviroment vaiable, if not exist create it
checkEnviromentDMSSystem

#Check folder, if not exist create it
checkDownloadFolder

if($systemDMS -eq "PROD") #define logic IF PROD or LAB   $systemDMS
{
    #Prod
    $adminDMUrl ="https://axpogrp.sharepoint.com/sites/applaxedm"
    Write-Host "*************** PROD *************"
}
else
{
    #LAB
    $adminDMUrl ="https://axpogrp.sharepoint.com/sites/appllabaxedm"
    Write-Host "*************** LAB *************"
}

Connect-PnpOnline -Url $adminDMUrl -CertificatePath $certPath -Tenant $tenantId -ClientId $appClientId

#Get all items with status neu
$items = Get-PnPListItem -List "Site Orders" | Where-Object { $_["Status"] -eq "Neu" }

foreach($item in $items)
{
    #Neue Seiten erstellen
    Write-Host $item["Title"]
    Write-Host $item["SeitenTyp"]

    $name = $item["Title"]
    $arrNames = $name.Split(',') #BK, Name
    $templateName = $item["SeitenTyp"]
    $siteTitleName = $name.Replace(", "," | ")
    $titleGroup = $name.Replace(", "," _ ")
    $urlName = $arrNames[0].ToLower().Replace(" ","") + "9099" #für Tests immer die Nummer verwenden

    Write-Host $siteTitleName
    Write-Host $titleGroup
    Write-Host $urlName

    #Change main variable PROD LAB ME10, Schaltanlage Zürich
    if($systemDMS  -eq "PROD")
    {
        #Prod
        $siteConfig = "https://axpogrp.sharepoint.com/sites/applaxedm"
        $hubsiteid = "XXXXXXXXXXXXXXXXXXX"
        $title = 'PROD | ' + $name 

        #make url
        if($templateName -eq "Unterwerk"){
            $alias = 'applaxedm' +'u'+ $urlName  # u Unterwek  t Trassee!!
        }
        elseif($templateName -eq "Trassee"){
            $alias = 'applaxedm' +'t'+ $urlName 
        }
        elseif($templateName -eq "Projekt"){
            $alias = 'applaxedm' +'p'+ $urlName 
        }
        else{
            #do nothing
            Write-Host "do nothing PROD"
        }

        $tempListFolderUnterwerk = "/sites/applaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-Unterwerk.csv"
        $tempListFolderTrassee = "/sites/applaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-Trassee.csv"
        $tempListFolderProjekt = "/sites/applaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-Projekt.csv"
        $tempListSearchPageUrl = "/sites/applaxedm/SiteTemplate/PnP-Provisioning-Search-Page-" + $templateName + ".aspx"
        $tempListSearchPageName = "PnP-Provisioning-Search-Page-" + $templateName + ".aspx"
    }
    else
    {
        #LAB
        $siteConfig = "https://axpogrp.sharepoint.com/sites/appllabaxedm"
        $hubsiteid = "388cea30-6275-4709-bde7-9bbee5012244"
        #$title = 'LAB | ' + $name 
        #$title = $name.Replace(",", " | ") 



        #make url
        if($templateName -eq "Unterwerk"){
            $alias = 'appllabaxedm' +'u'+ $urlName  # u Unterwek  t Trassee!!
        }
        elseif($templateName -eq "Trassee"){
            $alias = 'appllabaxedm' +'t'+ $urlName 
        }
        elseif($templateName -eq "Projekt"){
            $alias = 'appllabaxedm' +'p'+ $urlName 
        }
        else{
            #do nothing
            Write-Host "do nothing LAB"
        }

        #dodo: Check if we need the variable globaly
        $tempListFolderUnterwerk = "/sites/appllabaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-" + $templateName + ".csv"
        $tempListFolderTrassee = "/sites/appllabaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-" + $templateName + ".csv"
        $tempListFolderProjekt = "/sites/appllabaxedm/SiteTemplate/PnP-Provisioning-ListFolder-Template-" + $templateName + ".csv"
        $tempListSearchPageUrl = "/sites/appllabaxedm/SiteTemplate/PnP-Provisioning-Search-Page-" + $templateName + ".aspx"
        $tempListSearchPageName = "PnP-Provisioning-Search-Page-" + $templateName + ".aspx"

        Write-Host $alias
        Write-Host $tempListFolderUnterwerk
        Write-Host $tempListFolderTrassee
        Write-Host $tempListFolderProjekt 
        Write-Host $tempListSearchPageUrl
        Write-Host $tempListSearchPageName 
    }

    #Create Site
    $alias
    $adminSPUrl 
    $siteTitleName
    $alias
    $hubsiteid

    $siteUrl = CreateNewSite $adminSPUrl $siteTitleName $alias $hubsiteid

    #Create Groups
    createGroup $siteUrl $titleGroup

    #setPermissionLevel muss noch gemacht werden
    setPermissionLevel

    #Set SystemType
    setPropertyBag $adminSPUrl $siteUrl

    #Set Prefix DocumetnID
    setPropertyBagDocIdPrefix $adminSPUrl $siteUrl

    #Set list field default value
    #SetListDefaultValue $siteUrl $tpLevel1Value $tpLevel1IDValue $tpLevel1DesValue $templateName
    SetListDefaultValue $siteUrl $tpLevel1Value $tpLevel1IDValue  $tpLevel1DesValue $templateName
    #SetListDefaultValue $siteUrl "TR0146, Bütschwil - Bazenheid" "TR0146" "Bütschwil - Bazenheid" "Trassee"

    #Create folder on lists
    CreateFolderFromCSV $siteConfig $siteUrl

    #Copy search page  Copy-PnPFile -SourceUrl "/sites/project/Shared Documents/company.docx" -TargetUrl "/sites/otherproject/Shared Documents"
    copySearchPage $tempListSearchPageUrl $siteUrl 

    #Publis Search Site
    publishSearchFile $siteUrl

    #Delete dublicate NavNodes
    DeleteNavNodes $siteUrl

    #additionalSiteSettings
    additionalSiteSettings $siteUrl

    #add user to groups
    #todo

    #if site is created update site orders to Erstellt
    #todo

    #remove 365group from sitecolladmin 
    #todo
    
    ################
}
################# END MAIN  ####################