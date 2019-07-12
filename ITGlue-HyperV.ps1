Param (
       [string]$key = ""
       )

if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
{
$certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
    Add-Type $certCallback
 }
[ServerCertificateValidationCallback]::Ignore()

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$assettypeID = 134744

$ITGbaseURI = "https://api.itglue.com"
$ErrorActionPreference = "SilentlyContinue"
 
$headers = @{
    "x-api-key" = $key
}

Import-Module C:\temp\itglue\modules\itgluepowershell\ITGlueAPI.psd1 -Force
Add-ITGlueAPIKey -Api_Key $key
Add-ITGlueBaseURI -base_uri $ITGbaseURI

#
# Functions
#

function Get-ITGlueID($ServerName){

(Get-ITGlueConfigurations -filter_name $ServerName).data.id 

}

function Get-Customer($ServerName){

(Get-ITGlueConfigurations -filter_name $ServerName).data.attributes.'organization-id'

}

function GetAllITGItems($Resource) {
    $array = @()
    
    $body = Invoke-RestMethod -Method get -Uri "$ITGbaseURI/$Resource" -Headers $headers -ContentType application/vnd.api+json
    $array += $body.data
    Write-Host "Retrieved $($array.Count) items"
        
    if ($body.links.next) {
        do {
            $body = Invoke-RestMethod -Method get -Uri $body.links.next -Headers $headers -ContentType application/vnd.api+json
            $array += $body.data
            Write-Host "Retrieved $($array.Count) items"
        } while ($body.links.next)
    }
    return $array
}

function CreateITGItem ($resource, $body) {
    $item = Invoke-RestMethod -Method POST -ContentType application/vnd.api+json -Uri $ITGbaseURI/$resource -Body $body -Headers $headers
    #return $item
}

function UpdateITGItem ($resource, $existingItem, $newBody) {
    $updatedItem = Invoke-RestMethod -Method Patch -Uri "$ITGbaseUri/$Resource/$($existingItem.id)" -Headers $headers -ContentType application/vnd.api+json -Body $newBody
    return $updatedItem
}

function BuildHyperVAsset ($HyperVInfo) {
    
    $body = @{
        data = @{
            type       = "flexible-assets"
            attributes = @{
                "organization-id"        = $ITGlueOrganisation
                "flexible-asset-type-id" = $assettypeID
            traits                   = @{
                # Manual sync
            'vm-host-name' = $HyperVInfo.ServerID
            # Host platform
            'virtualisation-platform' = 'Hyper-V'
            # Host CPU data
            'cpu' = $HyperVInfo.cpu
            # Host RAM data
            'ram-gb' = $HyperVInfo.ram
            # Host disk data
            'disk-information' = $HyperVInfo.DiskInfo
            # Virutal network cards (vNIC)
            'virtual-switches' = $HyperVInfo.VSwitches
            # Number of VMs on host
            'current-number-of-vm-guests-on-this-vm-host' = $HyperVInfo.NumberofGuests
            # General VM data (start type, cpu, ram...)
            'vm-guest-names-and-information' = $HyperVInfo.GuestInfo
            # VMs' name and VHD paths
            'vm-guest-virtual-disk-paths' = $HyperVInfo.VMPaths
            # Snapshop data
            'vm-guests-snapshot-information' = $HyperVInfo.VMSnapshot
            # VMs' bios settings
            'vm-guests-bios-settings' = $HyperVInfo.BIOS
            # NIC and IP assigned to each VM
            'assigned-virtual-switches-and-ip-information' = $HyperVInfo.NICIPs

                 }
            }
        }
    }
    $HyperVAsset = $body | ConvertTo-Json -Depth 10
    return $HyperVAsset
}


#
# Data Collection
#
Write-Output "Retrieving configurations from IT Glue (org id: $organization_id)..."
$configurations = @{}
$MACs = @{}
$page_number = 1
do{
    Write-Output "Calling the IT Glue api for configurations (page $page_number, page size 1000)..."
    $api_call = Get-ITGlueConfigurations -organization_id $organization_id -page_size 1000 -page_number ($page_number++)
    foreach($_ in $api_call.data) {
        $configurations[$_.attributes.name] = $_
        if($_.attributes.'mac-address') {
            $MACs[$_.attributes.'mac-address'.replace(':','')] = $_
        }
    }
} while($api_call.links.next)
Write-Output "Done."


# All VMs on the host (with some data)
Write-Output "Trying to match VMs against IT Glue configurations and building VM data object..."
$VMs = @{}
foreach($vm in Get-VM) {
    $htmlname = $vm.name
    $conf_id = -1

    if($configurations[$vm.Name]) {
        $htmlname = '<a href="{0}/{1}/configurations/{2}">{3}</a>' -f $url,  $configurations[$vm.Name].attributes.'organization-id',  $configurations[$vm.Name].id, $vm.name
        $conf_id = $configurations[$vm.Name].id
        Write-Output "Matched $($vm.Name) on name to $($configurations[$vm.Name].id)."
    } elseif($MACs[($vm.Name | Get-VMNetworkAdapter).MacAddress]) {
        $config = $MACs[($vm.Name | Get-VMNetworkAdapter).MacAddress]
        $conf_id = $config.id
        $htmlname = '<a href="{0}/{1}/configurations/{2}">{3}</a>' -f $url,  $config.attributes.'organization-id',  $config.id, $config.attributes.name
        Write-Output "Matched $($vm.Name) on MAC address to $($config.id)."
    } else {
        $configurations.GetEnumerator() | Where {$_.Name -like "*$($vm.name)*"} | ForEach-Object {
            $htmlname = '<a href="{0}/{1}/configurations/{2}">{3}</a>' -f $url,  $_.value.attributes.'organization-id',  $_.value.id, $vm.name
            $conf_id = $_.value.id
            Write-Output "Matched $($vm.Name) on wildcard to $($_.value.id)."
        }
    }

    Write-Output "name = $($vm.name), vm = $($vm), conf_id = $($conf_id), htmlname = $($htmlname)"

    $VMs[$vm.name] = [PSCustomObject]@{
        name = $vm.name
        vm = $vm
        htmlname = $htmlname
        conf_id = $conf_id
    }
}
Write-Output "[1/9] VM data object done."

# Hyper-V host's disk information / "Disk information"
Write-Output "Getting host's disk data..."
$diskDataHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>Disk name</td>
                <td>Total(GB)</td>
                <td>Used(GB)</td>
                <td>Free(GB)</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ((Get-PSDrive -PSProvider FileSystem).foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
    </tr>' -f $_.Root, [math]::round(($_.free+$_.used)/1GB), [math]::round($_.used/1GB), [math]::round($_.free/1GB)} | Out-String)
Write-Output "[2/9] Host's disk data done."

# Virtual swtiches / "Virtual switches"
Write-Output "Getting virtual swtiches..."
$virtualSwitchsHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>Name</td>
                <td>Switch type</td>
                <td>Interface description</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ((Get-VMSwitch).foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
    </tr>' -f $_.Name, $_.SwitchType, $_.NetAdapterInterfaceDescription} | Out-String)
Write-Output "[3/9] Virtual swtiches done."

# General information about virtual machines / "VM guest names and information"
Write-Output "Getting general guest information..."
$guestInformationHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>VM guest name</td>
                <td>Start action</td>
                <td>RAM (GB)</td>
                <td>vCPU</td>
                <td>Size (GB)</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ($VMs.GetEnumerator().foreach{
    $diskSize = 0
    ($_.value.vm.HardDrives | Get-VHD).FileSize.foreach{$diskSize += $_}
    $diskSize = [Math]::Round($diskSize/1GB)
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>{4}</td>
    </tr>' -f $_.value.htmlname, $_.value.vm.AutomaticStartAction, [Math]::Round($_.value.vm.MemoryStartup/1GB), $_.value.vm.ProcessorCount, $diskSize} | Out-String)
Write-Output "[4/9] General guest information done."

# Virutal machines' disk file locations / "VM guest virtual disk paths"
Write-Output "Getting VM machine paths..."
$virtualMachinePathsHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>VM guest name</td>
                <td>Path</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ($VMs.GetEnumerator().foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
    </tr>' -f $_.value.htmlname, ((Get-VHD -id $_.value.vm.id).path | Out-String).Replace([Environment]::NewLine, '<br>').TrimEnd('<br>')} | Out-String)
Write-Output "[5/9] VM machine paths done."

# Snapshot data / "VM guests snapshot information"
Write-Output "Getting snapshot data..."
$vmSnapshotHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>VMName</td>
                <td>Name</td>
                <td>Snapshot type</td>
                <td>Creation time</td>
                <td>Parent snapshot name</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ((Get-VMSnapshot -VMName * | Sort VMName, CreationTime).foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>{4}</td>
    </tr>' -f $VMs[$_.VMName].htmlname, $_.Name, $_.SnapshotType, $_.CreationTime, $_.ParentSnapshotName} | Out-String)
Write-Output "[6/9] Snapshot data done."

# Virutal machines' bios settings / "VM guests BIOS settings"
Write-Output "Getting VM BIOS settings..."
# Generation 1
$vmBiosSettingsTableData = (Get-VMBios * -ErrorAction SilentlyContinue).foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>Gen 1</td>
    </tr>' -f $VMs[$_.VMName].htmlname, ($_.StartupOrder | Out-String).Replace([Environment]::NewLine, ', ').TrimEnd(', '), 'N/A', 'N/A'}
Write-Output "Generation 1 done..."

# Generation 2
Try {$vmBiosSettingsTableData += ( Get-VMFirmware * -ErrorAction SilentlyContinue).foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>Gen 2</td>
    </tr>' -f $VMs[$_.VMName].htmlname, ($_.BootOrder.BootType | Out-String).Replace([Environment]::NewLine, ', ').TrimEnd(', '), $_.PauseAfterBootFailure, $_.SecureBoot}}
    catch {
    'Get-VMFirmware Failed, may not be Gen2'
    }

Write-Output "Generation 2 done..."

$vmBIOSSettingsHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>VM guest name</td>
                <td>Startup order</td>
                <td>Pause After Boot Failure</td>
                <td>Secure Boot</td>
                <td>Generation</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ($vmBiosSettingsTableData | Out-String)
Write-Output "[7/9] VM BIOS settings done."

# Guest NICs and IPs
Write-Output "Getting VM NICs..."
$guestNICsIPsHTML = '<div>
    <table>
        <tbody>
            <tr>
                <td>VM guest name</td>
                <td>Swtich name</td>
                <td>IPv4</td>
                <td>IPv6</td>
                <td>MAC address</td>
            </tr>
            {0}
        </tbody>
    </table>
</div>' -f ((Get-VMNetworkAdapter * | Sort 'VMName').foreach{
    '<tr>
        <td>{0}</td>
        <td>{1}</td>
        <td>{2}</td>
        <td>{3}</td>
        <td>{4}</td>
    </tr>' -f $VMs[$_.VMName].htmlname, $_.switchname, $_.ipaddresses[0], $_.ipaddresses[1], $($_.MacAddress -replace '(..(?!$))','$1:') } | Out-String)
Write-Output "[8/9] VM NICs done."


$CPU = Get-VMHost | Select -ExpandProperty LogicalProcessorCount

$RAM = ((Get-CimInstance CIM_PhysicalMemory).capacity | Measure -Sum).Sum/1GB

$NumberofGuests = ($VMs.GetEnumerator() | measure).Count

$hostname = $env:computername

$ServerID = Get-ITGlueID -ServerName $hostname
$ITGlueOrganisation = Get-Customer -ServerName $hostname

$PSObject = @()
$object = New-Object psobject
$object | Add-Member -MemberType NoteProperty -Name ServerID -Value $ServerID
$object | Add-Member -MemberType NoteProperty -Name ITGlueOrg -Value $ITGlueOrganisation
$object | Add-Member -MemberType NoteProperty -Name CPU -Value $CPU
$object | Add-Member -MemberType NoteProperty -Name RAM -Value $RAM
$object | Add-Member -MemberType NoteProperty -Name NumberofGuests -Value $NumberofGuests
$object | Add-Member -MemberType NoteProperty -Name DiskInfo -Value $diskDataHTML
$object | Add-Member -MemberType NoteProperty -Name VSwitches -Value $virtualSwitchsHTML
$object | Add-Member -MemberType NoteProperty -Name GuestInfo -Value $guestInformationHTML
$object | Add-Member -MemberType NoteProperty -Name VMPaths -Value $virtualMachinePathsHTML
$object | Add-Member -MemberType NoteProperty -Name VmSnapshot -Value $vmSnapshotHTML
$object | Add-Member -MemberType NoteProperty -Name BIOS -Value $vmBiosSettingsTableData
$object | Add-Member -MemberType NoteProperty -Name NICIPs -Value $guestNICsIPsHTML
$PSObject += $object

$existingAssets = @()
$existingAssets += GetAllITGItems -Resource "flexible_assets?filter[organization_id]=$ITGlueOrganisation&filter[flexible_asset_type_id]=$assetTypeID"
$matchingAsset = $existingAssets | Where-Object {$_.attributes.traits.'vm-host-name'.values.id -contains $PSObject.serverID}

if ($matchingAsset) {
        Write-Output "Updating Hyper-V Flexible Asset"
        $UpdatedBody = BuildHyperVAsset -HyperVInfo $PSObject
        $updatedItem = UpdateITGItem -resource flexible_assets -existingItem $matchingAsset -newBody $UpdatedBody
        Start-Sleep -Seconds 3
    }
    else {
        Write-Output "Creating Hyper-V Flexible Asset"
        $body = BuildHyperVAsset -HyperVInfo $PSObject
        CreateITGItem -resource flexible_assets -body $body
        Start-Sleep -Seconds 3
        
    }