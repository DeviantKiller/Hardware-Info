  <##
##############################################################################################################
#                                                                                                            #
#       SCRIPT CREATE BY DANIEL K.                                                                           #
#       SCRIPT WILL GATHER ALL HARDWARE INFORMATION ON COMPUTERS OBTAINED VIA A CSV OR ACTIVE DIRECTORY.     #
#       TO USE A CSV FILE, UNCOMMENT OUT THE LOCATION PICKER AND COMMENT OUT THE AD FILTER                   #
#       SAVE LOCATOIN CAN BE MODIFIED USING THE $path VARIABLE                                               #
#       TO CHANGE THE OUTPUT FILENAME, THIS CAN BE CHANGED AT $exportLocation                                #
#       SCRIPT WILL DELETE PREVIOUS livePC list to make sure no entries are duplicated.                      #
#                                                                                                            #
##############################################################################################################
1.0 - 	First Draft created
1.1 -   Added Last_ReBoot
	    Added LastUser
	    Added Build_Date
	    Added ability to import computers from csv file
1.2	    Added Clean-ups to run at start of script
	    Added Graphics card adapters
1.2.1	Added Eco965 Graphics
1.3	    Added Location Information
	    Added Rename to end of script to date csv
	    Improved method on getting last logged on user
	    Changed naming convention for foreach loops.
	    Made Variables easier to modify
	    Re-arranged Script to make it more logical and grouped
1.3.1	Improved method on getting last logon time
1.3.2   Cleaned up un-needed Code (Old methods that are now unused.)
        Removed un-needed blank lines
1.3.3	Added "Completion time" line
1.4     Removed old test methods of last login, last reboot
        Fixed old variable in write-host that no longer existed
        Added [0] for ecoquiets as 2 drivers appear in device ID
1.4.1   Changed [0] method to Select-Object -first 1
1.5     Added array for computer counting and numbering for info
1.6     New method of obtaining CPU info - Less resource and similtanious
1.7     Re-write of all info gathering. now using  "Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance ...} "
1.8     Added Out-Gridview
        Ping once on hardware gather to avoid hangs if computers are shutdown
1.9     Created Functions to deal with bulk of script
1.9.1   Added filter to remove USB devices showing up as a HDD/SSD
2.0     Added UEFI/Legacy check + SecureBoot check. - Requires Testing
2.0.1   Added Memory bank information - Slots used/Total and if spare
        Fixed Gridview "Video card" to use '$global:CardName' instead of '$CardName'
        Changed Gridview "HDD Name" to "HDD Model"
2.0.2   Updated BIOS/UEFI mode
        Added Network speed pulled from network adapter.
##>

Function Pingit
    {
        $ping = Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16
    }
 
Function FindOS
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:OS = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_OperatingSystem} | Select InstallDate, OSArchitecture, Caption, LastBootUpTime
            $global:WindowsOS = $OS.Caption
            $global:WindowsInstallDate = $global:OS.InstallDate
            $global:WindowsLastBootUpTime = $global:OS.LastBootUpTime
            $global:WindowsOSArchitecture = $global:OS.OSArchitecture
        }
        Else {
            $Global:WindowsOS = "Station Shutdown"
            $global:WindowsLastBootUpTime = "Station Shutdown"
            $global:WindowsOSArchitecture = "Station Shutdown"
        }
    }


Function FindBIOS
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:bios = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance win32_bios} | Select SerialNumber, SMBIOSBIOSVersion
            $global:systemBios = $global:Bios.serialnumber
            $global:BiosVer = $global:bios.SMBIOSBIOSVersion
        }
        Else {
            $global:systemBios = "Station Shutdown"
            $global:BiosVer = "Station Shutdown"
        }
    }

Function FindcompSys
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:compSys = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_computerSystem} | Select SerialNumber, Manufacturer, Model, TotalPhysicalMemory, SystemType
            $global:Manufacturer = $global:compSys.Manufacturer
            $global:PCModel = $($global:compSys.Model)
            $global:OSType = $global:compSys.SystemType
            $global:totalMemory = [math]::round($global:compSys.TotalPhysicalMemory/1024/1024/1024, 2)
        }
        Else {
            $global:Manufacturer = "Station Shutdown"
            $global:PCModel = "Station Shutdown"
            $global:OSType = "Station Shutdown"
            $global:totalMemory = "Station Shutdown"
        }
    }

Function FindProcessor
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:cpu = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_Processor} | Select Name, NumberOfCores, NumberOfLogicalProcessors
            $global:CpuName = $global:cpu.Name
            $global:CpuCores = $global:cpu.NumberOfCores
            $global:CpuLogical = $global:cpu.NumberOfLogicalProcessors
        }
        Else {
            $global:CpuName = "Station Shutdown"
            $global:CpuCores = "Station Shutdown"
            $global:CpuLogical = "Station Shutdown"
        }
    }

Function RAMInfo
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:memory = Invoke-Command -ComputerName $livePC -Scriptblock {Get-WmiObject -class "win32_physicalmemory" -namespace "root\CIMV2"}
            #Write-Host "Memore Modules:" -ForegroundColor Green
            #$global:memory | Format-Table Tag,BankLabel,@{n="Capacity(GB)";e={$_.Capacity/1GB}},Manufacturer,PartNumber,Speed -AutoSize
            $global:TotalSlots = Invoke-Command -ComputerName $livePC -Scriptblock {((Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2").MemoryDevices | Measure-Object -Sum).Sum}
            $global:roundedMem = $((($global:memory).Capacity | Measure-Object -Sum).Sum/1GB)
            $global:UsedSlots = (($global:memory) | Measure-Object).Count
            $global:FreeBankCount = $global:TotalSlots-$global:UsedSlots
        }
        Else {
            $global:UsedSlots = "Station Shutdown"
            $global:TotalSlots = "Station Shutdown"
            $global:roundedMems = "Station Shutdown"
        }
        If ($global:UsedSlots -eq $global:TotalSlots) {
            $global:spare = "No"
        }
        Else {
            $global:spare = "Yes"
        }
    }

Function FindWMI
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:wmi = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_WmiSetting} | Select BuildVersion
            $global:BuildVersion = $global:wmi.BuildVersion
        }
        Else {
            $global:BuildVersion = "Station Shutdown"
        }
    }

Function FindWin10Ver
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:newSysBuild = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseId | select ReleaseId}
            $global:Win10Ver = $global:newSysBuild.ReleaseId
        }
        Else {
            $global:Win10Ver = "Station Shutdown"
        }
    }

Function FindWin10VerNew
    {
        $regkeypath= "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" 
        $value1 = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path $regkeypath}.ReleaseId -eq $null
        if ($value1 -eq $False) {
            $global:newSysBuild = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseId | select ReleaseId}
            $global:Win10Ver = $global:newSysBuild.ReleaseId
        }
        Elseif ($value1 -eq $False) {
            $global:Win10Ver = "Not Win10"
        }
        Else {
            $global:Win10Ver = "Station Offline"
        }
    }

Function FindNetwork
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:networks = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_NetworkAdapterConfiguration | ? {$_.IPEnabled}} | Select IPAddress, MACAddress, Description
            $global:IPAddress = $global:Networks.IPAddress[0]
            $global:MACAddress = $global:Networks.MACAddress
            $global:NetworkDescription = $global:Networks.Description
        }
        Else {
            $global:IPAddress = "Station Shutdown"
            $global:MACAddress = "Station Shutdown"
        }
    }

Function FindNetworkSpeed
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:networkSpeed = Invoke-Command -ComputerName $livePC -Scriptblock {Get-NetAdapter | ? {$_.InterfaceOperationalStatus}} | Select-Object -Property Name, InterfaceDescription, MacAddress, FullDuplex, LinkSpeed  
            $global:InterfaceDescription = $global:networkSpeed.InterfaceDescription
            $global:duplexspeed = $global:networkSpeed.FullDuplex
            $global:networkLinkSpeed = $global:networkSpeed.LinkSpeed
        }
        Else {
            $global:duplexspeed = "Station Shutdown"
            $global:networkLinkSpeed = "Station Shutdown"
        }
    }

Function FindVol
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:vol = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_volume} | Select Caption, driveletter, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)}  }, @{LABEL='HDDCapacity';EXPRESSION={"{0:N2}" -f($_.Capacity/1GB)}  }| Where-Object { $_.driveletter -match "C:" } | where {$_.Caption -like "C:\*"} 
            $global:HDD = $global:vol.HDDCapacity
            $global:HDDFree = $global:vol.GBfreespace
        }
        Else {
            $global:HDD = "Station Shutdown"
            $global:HDDFree = "Station Shutdown"
        }
    }

Function FindDiskType
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:Disk = Invoke-Command -ComputerName $livePC -Scriptblock {Get-PhysicalDisk | Select FriendlyName, MediaType, Model, SerialNumber, Bustype} | Where-Object -FilterScript {$_.Bustype -NotLike "USB"}
            $global:Name = $global:Disk.Model
            $global:DiskType = $global:Disk.MediaType
            $global:DiskSerial = $global:Disk.SerialNumber
        }
        Else {
            $global:Name = "Station Shutdown"
            $global:DiskType = "Station Shutdown"
        }
    }
    
Function FindVideoController
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
                $global:gpuDriver = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_VideoController} | Select Name, AdapterCompatibility, DriverVersion, CurrentHorizontalResolution, CurrentVerticalResolution
                $global:CardName = $global:gpuDriver.Name
                $global:CardAdapter = $global:gpuDriver.AdapterCompatibility
                $global:CardDrivers = $global:gpuDriver.DriverVersion
                $global:CardHoriz = $global:gpuDriver.CurrentHorizontalResolution
                $global:CardVert = $global:gpuDriver.CurrentVerticalResolution
                $global:monRes = "$global:CardHoriz x $global:CardVert"
        }
        Else {

        }
    }

Function FindNvidia
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:nvidiaID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*nvidia*"} | Select-Object -first 1
            $global:Nvidia = $global:nvidia.HardWareID 
        }
        Else {
                $global:Nvidia = "Station Shutdown"
        }
    }

Function FindAMD
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:AMDID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*AMD*"} | Select-Object -first 1
            $global:AMD = $global:AMDID.HardWareID 
        }
        Else {
                $global:AMD = "Station Shutdown"
        }
    }

Function FindIntel
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:intelID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Graphics*"} | Select-Object -first 1
            $global:intel = $global:intelID.HardWareID
        }
        Else {
            $global:intel = "Station Shutdown"
        }
    }
    
Function FindBasic
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:MicroID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Basic Display Ad*"} | Select-Object -first 1
            $global:MSBasic = $global:MicroID.HardWareID 
        }
        Else {
            $global:MSBasic = "Station Shutdown"
        }
    }
    
Function FindHyperV
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:MicroHyperID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Hyper-V Video*"} | Select-Object -first 1
            $global:HyperV = $global:MicroHyperID.HardWareID 
        }
        Else {
            $global:MSBasic = "Station Shutdown"
        }
    }

Function FindEco965
    {
       if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
            $global:EcoID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Intel(R) 965 Express*"} | Select-Object -first 1 
            $global:Eco965 = $global:EcoID.HardWareID
        }
        Else {
            $global:Eco965 = "Station Shutdown"
        }
    }
    
Function FindLastUser
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
                $global:username = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\Software\Microsoft\windows\currentVersion\Authentication\LogonUI\' -Name LastLoggedOnUser | select LastLoggedOnUser}
                $global:User = $global:username.LastLoggedOnUser
        }
        Else {
                $global:User = "Station Shutdown"
        }
    }
   
Function FindCurrentLocation
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
                $global:location = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\RM\Connect\' -Name CurrentLocation | select CurrentLocation} 
                $global:StationLocation = $global:location.CurrentLocation
        }
        Else {
                $global:StationLocation = "Station Shutdown"
        }
    }
    
Function FindLastLoggedOn
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16) {
                $global:lastlogtime = Get-ADComputer -identity $livePC -Properties * | select LastLogonDate
                $global:Lastlogontime = $global:lastlogtime.LastLogonDate
        }
        Else {
                $global:Lastlogontime = "Station Shutdown"
        }
    }
    
Function CheckExisting
    {
        If(!(Test-path $path)) {
                New-Item -ItemType Directory -Force -Path $path
                Write-Host -foregroundcolor Green "$path has been created."
            }
        Else {
                Write-Host -foregroundcolor Cyan "$path has ALREADY been created."
            }

#Deletes livePC's text 
        if (Test-Path $fileName) {
                Remove-Item $fileName
                Write-Host -foregroundcolor Cyan "The old $fileName has been deleted."
            }
        Else {
                Write-Host -foregroundcolor Green "$fileName does not exist..."
            }
        #Removes incomplete $csv 
        if (Test-Path $exportLocation) {
                Remove-Item $exportLocation
                Write-Host -foregroundcolor Red "The old $exportLocation has been deleted."
            }
        Else {
                #Write-Host -foregroundcolor Darkyellow "$exportLocation does not exist..."
            }
    }

Function BiosUEFI
    {
        if(!(Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16)) {
            $global:BIOSUEFIstatus = "Shutdown" 
            }
        else {
            $biosUEFI = Invoke-Command -ComputerName $livePC {Get-SecureBootUEFI -Name setupmode}
            }
        if ($biosUEFI.Bytes) {
            $global:BIOSUEFIstatus = "UEFI"
        }
        else {
            $global:BIOSUEFIstatus = "BIOS"
        }
    }

Function SecureBoot
    {
        if(!(Test-Connection -ComputerName $livePC  -Quiet -count 1 -BufferSize 16)) {
            $global:SecureBootstatus = "Shutdown" 
            }
        else {
            $SecureBoot = Invoke-Command -ComputerName $livePC {Get-SecureBootUEFI -Name SecureBoot} 
            }
        if ($SecureBoot.Bytes) {
            $global:SecureBootstatus = "SecureBoot On"
        }
        else {
            $global:SecureBootstatus = "SecureBoot Off"
        }        
    }

Function WriteHardwareCSV
    {
        $OutputObj  = New-Object -Type PSObject
        $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $livePC.ToUpper()
        $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $global:Manufacturer
        $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $global:PCModel
        $OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $global:systemBios
        $OutputObj | Add-Member -MemberType NoteProperty -Name CPU -Value $global:CpuName
        $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Cores -Value $global:CpuCores
        $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_LogicalCores -Value $global:CpuLogical
        $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Physical_Memory -Value $global:totalMemory
        $OutputObj | Add-Member -MemberType NoteProperty -Name Used/total -Value "$global:UsedSlots/$global:TotalSlots"
        $OutputObj | Add-Member -MemberType NoteProperty -Name C:Size -Value $global:HDD
        $OutputObj | Add-Member -MemberType NoteProperty -Name C:GBfreeSpace -Value $global:HDDFree
        $OutputObj | Add-Member -MemberType NoteProperty -Name DiskName -Value $global:Name
        $OutputObj | Add-Member -MemberType NoteProperty -Name DiskType -Value $global:DiskType
        $OutputObj | Add-Member -MemberType NoteProperty -Name DiskSerial -Value $global:DiskSerial
        $OutputObj | Add-Member -MemberType NoteProperty -Name Microsoft_ID -Value $global:MSBasic
        $OutputObj | Add-Member -MemberType NoteProperty -Name Intel_ID -Value $global:intel
        $OutputObj | Add-Member -MemberType NoteProperty -Name Eco965_ID -Value $global:Eco965
        $OutputObj | Add-Member -MemberType NoteProperty -Name nVidia_ID -Value $global:Nvidia
        $OutputObj | Add-Member -MemberType NoteProperty -Name HyperV_ID -Value $global:HyperV
        $OutputObj | Add-Member -MemberType NoteProperty -Name HyperV_ID -Value $global:AMDID
            Foreach ($Card in $gpuDriver) {
                $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Name" $Card.Name
                #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Description" $Card.Description #Probably not needed. Seems to just echo the name. Left here in case I'm wrong!
                $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Vendor" $Card.AdapterCompatibility
                #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_PNPDeviceID" $Card.PNPDeviceID
                $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_DriverVersion" $Card.DriverVersion
                $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_VideoMode" $Card.VideoModeDescription
            }
        $OutputObj | Add-Member -MemberType NoteProperty -Name OS -Value $global:WindowsOS
        $OutputObj | Add-Member -MemberType NoteProperty -Name SystemType -Value $global:WindowsOSArchitecture
        $OutputObj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $global:BuildVersion
        $OutputObj | Add-Member -MemberType NoteProperty -Name BIOS/UEFI -Value $global:BIOSUEFIstatus
        $OutputObj | Add-Member -MemberType NoteProperty -Name SecureBoot -Value $global:SecureBootstatus

        #$OutputObj | Add-Member -MemberType NoteProperty -Name Windows10Ver -Value $global:Win10Ver
        #$OutputObj | Add-Member -MemberType NoteProperty -Name SPVersion -Value $OS.csdversion
        $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $global:IPAddress
        $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $global:MACAddress
        $OutputObj | Add-Member -MemberType NoteProperty -Name LastUser -Value $global:User 
        $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Logon -Value $global:Lastlogontime
        $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Reboot -Value $global:WindowsLastBootUpTime
        $OutputObj | Add-Member -MemberType NoteProperty -Name Build_Date -Value $global:WindowsInstallDate
        $OutputObj | Add-Member -MemberType NoteProperty -Name Location -Value $global:StationLocation #OU information for CSV
        $OutputObj | Export-Csv $exportLocation -Append -NoTypeInformation
    }

Function HardwareGrid
    {
        ForEach-Object{
            $newobj = [PSCustomObject]@{
                'ComputerName' = $livePC.ToUpper()
                'Location' = $global:StationLocation
                'Manufacturer' = $global:Manufacturer
                'Model' = $global:PCModel
                'Serial Number' = $global:systemBios
                'CPU' = $global:CpuName
                'Cores' = $global:CpuCores
                'Logical' = $global:CpuLogical
                'RAM' = $global:roundedMem
                'Used/Total' = "$global:UsedSlots/$global:TotalSlots"
                'OS' = $global:WindowsOS
                'BIOS/UEFI' = $global:BIOSUEFIstatus
                'SecureBoot' = $global:SecureBootstatus                             
                'HDD Size' = $global:HDD
                'HDD Free' = $global:HDDFree
                'HDD Model' = $global:Name
                'HDD Type' = $global:DiskType
                'HDD Serial' = $global:DiskSerial
                'Video Card' = $global:CardName
                'Resolution' = $global:monRes
                'IP Address' = $global:IPAddress
                'MAC' = $global:MACAddress
                'Last User' = $global:User 
                'Login Date' = $global:Lastlogontime
                'Last Restart' = $global:WindowsLastBootUpTime
                'Last Build' = $global:WindowsInstallDate
            }
        $ResultObject.Add($newobj)
        }
    }

Function DevShowAllInfo
    {
        Write-Host "Station: " -f White -nonewline; Write-Host "$livePC" -ForegroundColor Yellow -nonewline; Write-Host " Location: " -f White -nonewline; Write-Host "$global:StationLocation" -ForegroundColor Yellow;
        #Computer
        Write-Host -ForegroundColor White "Manufacturer:$global:Manufacturer"
        Write-Host -ForegroundColor White "PCModel: $global:PCModel"
        Write-Host -ForegroundColor White "Serial Number: $global:systemBios"
        #Processor
        Write-Host -ForegroundColor Yellow "Processor"
        Write-Host -ForegroundColor White "Processor: $global:CpuName"
        Write-Host -ForegroundColor White "Cores: $global:CpuCores"
        Write-Host -ForegroundColor White "Logical Cores: $global:CpuLogical"
        #RAM
        Write-Host -ForegroundColor Yellow "RAM"
        Write-Host -ForegroundColor White "RAM: $global:roundedMem"
       # Write-Host -ForegroundColor White "RAM: $global:totalMemory"
       # Write-Host -ForegroundColor White "RAM Spare Slot: $global:spare"
        Write-Host -ForegroundColor White "Used/Total: $global:UsedSlots/$global:TotalSlots"
        #Storage
        Write-Host -ForegroundColor Yellow "Storage"
        Write-Host -ForegroundColor White "Capacity: $global:HDD"
        Write-Host -ForegroundColor White "Free: $global:HDDFree"
        Write-Host -ForegroundColor White "Disk Name: $global:Name"
        Write-Host -ForegroundColor White "Disk Type: $global:DiskType"
        Write-Host -ForegroundColor White "Disk Serial: $global:DiskSerial"
        #Graphics
        Write-Host -ForegroundColor Yellow "Graphics"
            If($global:MSBasic) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "Hardware ID :$global:MSBasic"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Elseif($global:intel) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "Intel Graphics Hardware ID :$global:intel"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Elseif($global:Eco965) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "Intel/Eco965Hardware ID :$global:Eco965"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Elseif($global:Nvidia) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "Nvidia Hardware ID :$global:Nvidia"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Elseif($global:AMD) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "AMD Hardware ID : $global:AMD"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Elseif($global:HyperV) {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Green "Graphics Driver in use: $global:CardName"
                Write-Host -ForegroundColor White "Nvidia Hardware ID : $global:HyperV"
                Write-Host -ForegroundColor White "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
            Else {
                Write-Host -ForegroundColor White "Manufacturer: $global:CardAdapter"
                Write-Host -ForegroundColor Yellow "Graphics Driver in use: $global:CardName"
                Write-Warning "Please report station name to Dan"
                Write-Host -ForegroundColor Red "Not Sure What This Is?"
                Write-Host -ForegroundColor Green "Version: $global:CardDrivers"
                Write-Host -ForegroundColor White "Resolution: $global:monRes" 
            }
        #Windows OS Info
        Write-Host -ForegroundColor Yellow "Windows"
        Write-Host "$global:WindowsOS" -f white -nonewline; Write-Host " - $global:Win10Ver" -ForegroundColor white
        Write-Host -ForegroundColor White "Architecture: $global:WindowsOSArchitecture"
        Write-Host -ForegroundColor White "BuildVersion: $global:BuildVersion"
        Write-Host -ForegroundColor White "BIOS/UEFI: $global:BIOSUEFIstatus"
        Write-Host -ForegroundColor White "SecureBoot Status: $global:SecureBootstatus"
        Write-Host -ForegroundColor White "IP Address: $global:IPAddress"
        Write-Host -ForegroundColor White "MAC Address: $global:MACAddress"
        Write-Host -ForegroundColor White "Last User: $global:User"
        Write-Host -ForegroundColor White "Last Login: $global:Lastlogontime"
        Write-Host -ForegroundColor White "Last Restart: $global:WindowsLastBootUpTime"
        Write-Host -ForegroundColor White "Windows Install: $global:WindowsInstallDate"
        write-host ""
    }

# On error the script will continue silently without 
$erroractionpreference = "SilentlyContinue"
$startTime = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")  
<# Using a text file with a location picker
$fileLocation = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
InitialDirectory = [Environment]::GetFolderPath('Desktop') 
Filter = 'Documents (*.csv)|*.csv|txt files (*.txt)|*.txt'
}
$null = $fileLocation.ShowDialog()
$testcomputers = $fileLocation.FileName
#>

# Use AD to get Computers and filter out all that aren't Windows Server OS
#$testcomputers = Get-ADComputer -Filter {(OperatingSystem -notlike "*windows*server*")} 
$testcomputers = Get-ADComputer -Filter { OperatingSystem -notlike "*Windows*Server*" } | where-object {[string]$_.Name -notlike "*NETMAN*" }
# Looking through the txt file above and counting computer names. then working out % based on how many have been pinged
$test_computer_count = $testcomputers.Length;
#Percentage Count
$x = 0;
#Dead Computer Count
$xDead = 0;
#Online Computer Count
$xLive = 0;
#Total Computer Count
$xTotal = 0;
#Resets values if re-ran in ISE
$ResultObject = $null
$newobj = $null
$ping = $null
#Can Change These 
$path = "D:\temp\Hardware"
$csvName = "pcInventory.csv"
$deadPCs = "deadPCs.txt"
#Hard-coded
$fileName = "$path\$deadPCs"
$exportLocation = "$path\$csvName"
$global:online = @()
$global:offline = @()
$ResultObject = New-Object System.Collections.Generic.List[object]

CheckExisting

write-host `r
write-host -foregroundcolor Yellow "Checking online status of $test_computer_count computers, this may take a while."
# PING COMPUTERS TO FIND OUT ONLINE/OFFLINE LIST
foreach ($computer in $testcomputers) {
    $livecomp = $($computer.Name)
    $ping = Test-Connection -ComputerName $livecomp  -Quiet -count 1 -BufferSize 16
    if ($ping){
        $hostname = $livecomp
        $status = 'ONLINE'
        $online += $hostname
        $xLive++;
        $xTotal++;
        #Write-Host -foregroundcolor Green "$xTotal. $livecomp has been added to the Live List. $xLive of $test_computer_count ONLINE"
        Write-Host "$xTotal." -f white -nonewline; Write-Host " $livecomp has been added to the Live List." -ForegroundColor Cyan -nonewline; Write-Host " $xLive of $test_computer_count ONLINE" -f Green;
    }
    else {
        $hostname = $livecomp
        $status = 'OFFLINE'
        $offline += $hostname
        $xDead++;
        $xTotal++;
        Add-Content -value $computer.Name -path $fileName
        #Write-Host -foregroundcolor Red "$xTotal. $livecomp cannot be contacted. $xDead of $test_computer_count OFFLINE"
        Write-Host "$xTotal." -f white -nonewline; Write-Host " $livecomp cannot be contacted." -ForegroundColor Cyan -nonewline; Write-Host " $xDead of $test_computer_count OFFLINE" -f Red;
    }
    $testcomputer_progress = [int][Math]::Ceiling((($x / $test_computer_count) * 100))
	# Progress bar
    Write-Progress  "Testing Connections" -PercentComplete $testcomputer_progress -Status "Percent Complete - $testcomputer_progress%" -Id 1;
	Sleep(1);
    $x++;
} 

write-host `r
write-host -foregroundcolor cyan "Testing Connection complete. $xLive of $test_computer_count Computers are online."
write-host -foregroundcolor Red "$xDead of $test_computer_count Computers could not be contacted."
write-host -foregroundcolor cyan "Writing Data to CSV."
write-host `r
 
$ComputerName = $online
$computer_count = $ComputerName.Length;
$i = 0;
$p = 0;
#$livePC = "BRKNETMAN"
foreach ($livePC in $ComputerName){
    if (Test-Connection -ComputerName $livePC -Quiet -count 1 -BufferSize 16){
        $p++;
        write-host -foregroundcolor Green "Probing $livePC - $p of $($ComputerName.Count)"
        $computer_progress = [int][Math]::Ceiling((($i / $computer_count) * 100))
	    # Progress bar
        Write-Progress  "Gathering Hardware Info" -PercentComplete $computer_progress -Status "Percent Complete - $computer_progress%" -Id 1;
	    Sleep(1);
        $i++;

        FindOS
        BiosUEFI
        SecureBoot
        FindBIOS
        FindcompSys
        FindProcessor
        RAMinfo
        FindWMI
        FindWin10Ver
        FindNetwork
        FindNetworkSpeed
        FindVol
        FindDiskType
        #Graphics Adapters
        FindVideoController
        #Individual Graphics Devices
        FindNvidia
        FindIntel
        FindBasic
        FindHyperV
        FindEco965
        FindAMD

        #Registry Commands
        FindLastUser
        FindCurrentLocation
    
        #Active Directory
        FindLastLoggedOn
        
        #Methods of viewing data #Comment out which are not required.
        DevShowAllInfo
        WriteHardwareCSV
        HardwareGrid
    } 
}
write-host ""
write-host -foregroundcolor cyan "Renaming $csvName" 
$newCSV = Get-ChildItem $exportLocation 
$time = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")    
$RenameCSV = "$($newCSV.DirectoryName)\$($newCSV.BaseName)[$($time)]$($newCSV.Extension)"
Rename-Item $exportlocation -NewName $RenameCSV    

Write-Host -ForegroundColor Yellow
#write-host -foregroundcolor cyan "$csvName has been renamed to $RenameCSV"
#write-host -foregroundcolor cyan "Script is complete, the results are here: $path"
#Write-host -foregroundcolor green "Script started at $startTime and finished at $time"
Write-Host "$csvName " -f yellow -nonewline; Write-Host "has been renamed to" -ForegroundColor white -nonewline; Write-Host " $RenameCSV." -f yellow;
Write-Host "Script has complete." -f green -nonewline; Write-Host " The Results have been saved to" -ForegroundColor white -nonewline; Write-Host "  $path." -f yellow;
Write-Host "The Script started at" -f white -nonewline; Write-Host " $startTime" -ForegroundColor Yellow -nonewline; Write-Host " and finished at" -f white; Write-Host "$time" -ForegroundColor Yellow -nonewline;
Invoke-Item $path
$ResultObject  | Out-Gridview -Title "Hardware Information"
