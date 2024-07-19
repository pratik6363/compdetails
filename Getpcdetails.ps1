# Function to get RAM details in GB
function Get-RAMInfo {
    $ram = Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
    return [math]::Round($ram.Sum / 1GB, 2)  # Convert to GB and round to 2 decimal places
}

# Function to get total hard disk size in GB
function Get-TotalHardDiskSize {
    $disks = Get-CimInstance -ClassName Win32_DiskDrive | Measure-Object -Property Size -Sum
    return [math]::Round($disks.Sum / 1GB, 2)  # Convert to GB and round to 2 decimal places
}

# Function to get graphics card details (if available)
function Get-GraphicsCardInfo {
    $graphics = Get-CimInstance -ClassName Win32_VideoController
    if ($graphics) {
        return $graphics.Name
    } else {
        return "N/A"
    }
}

# Function to prompt for owner name and department
function Get-OwnerAndDepartment {
    $owner = Read-Host "Enter owner name"
    $department = Read-Host "Enter user department"
    return @{
        Owner = $owner
        Department = $department
    }
}

# Function to get laptop model and serial number
function Get-LaptopInfo {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios = Get-CimInstance -ClassName Win32_BIOS
    $model = $cs.Model
    $serialNumber = $bios.SerialNumber
    return @{
        Model = $model
        SerialNumber = $serialNumber
    }
}

# Function to get disk details
function Get-DiskInfo {
    $disks = Get-CimInstance -ClassName Win32_DiskDrive
    $diskInfo = @()
    foreach ($disk in $disks) {
        $diskInfo += [PSCustomObject]@{
            Name = $disk.Model
            SerialNumber = $disk.SerialNumber
        }
    }
    return $diskInfo
}

# Function to get Microsoft Store account ID
function Get-MicrosoftStoreAccountID {
    $result = Get-AppxPackage -AllUsers | Where-Object { $_.Name -eq "Microsoft.WindowsStore" }
    if ($result) {
        return $result.PackageFamilyName
    } else {
        return "N/A"
    }
}

# Get owner name and department
$ownerAndDept = Get-OwnerAndDepartment
$owner = $ownerAndDept.Owner
$department = $ownerAndDept.Department

# Get computer details
$computerName = $env:COMPUTERNAME
$os = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
$installDate = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty InstallDate
$ram = Get-RAMInfo
$totalHardDiskSize = Get-TotalHardDiskSize
$drive = Get-CimInstance Win32_LogicalDisk | Where-Object {$_.DeviceID -eq 'C:'} | Select-Object -ExpandProperty FreeSpace
$user = $env:USERNAME

# Get laptop details
$laptopInfo = Get-LaptopInfo
$laptopModel = $laptopInfo.Model
$laptopSerial = $laptopInfo.SerialNumber

# Get graphics card info
$graphicsCard = Get-GraphicsCardInfo

# Get disk details
$diskInfo = Get-DiskInfo

# Get Microsoft Store account ID
$microsoftStoreID = Get-MicrosoftStoreAccountID

# Determine desktop path for the current user
$desktopPath = [System.Environment]::GetFolderPath("Desktop")

# Construct file name with owner name and department
$fileName = Join-Path -Path $desktopPath -ChildPath "${owner}_${department}_ComputerDetails.csv"

# Create an object to store all details
$computerDetails = [PSCustomObject]@{
    Owner = $owner
    Department = $department
    ComputerName = $computerName
    OperatingSystem = $os
    WindowsInstallDate = $installDate
    RAM_GB = $ram
    TotalHardDiskSize_GB = $totalHardDiskSize
    FreeDriveSpace_GB = "{0:N2}" -f ($drive / 1GB)
    UserName = $user
    LaptopModel = $laptopModel
    LaptopSerialNumber = $laptopSerial
    GraphicsCard = $graphicsCard
    MicrosoftStoreAccountID = $microsoftStoreID
    Disks = $diskInfo
}

# Export details to CSV file
$computerDetails | Export-Csv -Path $fileName -NoTypeInformation

Write-Host "Computer details have been saved to $fileName"
