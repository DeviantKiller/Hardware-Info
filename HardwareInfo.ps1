  <##
SCRIPT CREATE BY DANIEL K.
SCRIPT WILL GATHER ALL HARDWARE INFORMATION ON COMPUTERS OBTAINED VIA A CSV OR ACTIVE DIRECTORY.
TO USE A CSV FILE, UNCOMMENT OUT THE LOCATION PICKER AND COMMENT OUT THE AD FILTER
    SAVE LOCATOIN CAN BE MODIFIED USING THE $path VARIABLE
    TO CHANGE THE OUTPUT FILENAME, THIS CAN BE CHANGED AT $exportLocation
SCRIPT WILL DELETE PREVIOUS livePC list to make sure no entries are duplicated.
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

	
##>
 
Function FindOS
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:OS = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_OperatingSystem} | Select InstallDate, OSArchitecture, Caption, LastBootUpTime
                $global:WindowsOS = $OS.Caption
                $global:WindowsInstallDate = $global:OS.InstallDate
                $global:WindowsLastBootUpTime = $global:OS.LastBootUpTime
                $global:WindowsOSArchitecture = $global:OS.OSArchitecture
            }
        Else
            {
                $Global:WindowsOS = "Station Shutdown"
                $global:WindowsLastBootUpTime = "Station Shutdown"
                $global:WindowsOSArchitecture = "Station Shutdown"
            }
    }


Function FindBIOS
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:bios = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance win32_bios} | Select SerialNumber, SMBIOSBIOSVersion
                $global:systemBios = $global:Bios.serialnumber
                $global:BiosVer = $global:bios.SMBIOSBIOSVersion
            }
        Else
            {
                $global:systemBios = "Station Shutdown"
                $global:BiosVer = "Station Shutdown"
            }
    }

Function FindcompSys
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:compSys = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_computerSystem} | Select SerialNumber, Manufacturer, Model, TotalPhysicalMemory, SystemType
                $global:Manufacturer = $global:compSys.Manufacturer
                $global:PCModel = $($global:compSys.Model)
                $global:OSType = $global:compSys.SystemType
                $global:totalMemory = [math]::round($global:compSys.TotalPhysicalMemory/1024/1024/1024, 2)
            }
        Else
            {
                $global:Manufacturer = "Station Shutdown"
                $global:PCModel = "Station Shutdown"
                $global:OSType = "Station Shutdown"
                $global:totalMemory = "Station Shutdown"
            }
    }

Function FindProcessor
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:cpu = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_Processor} | Select Name, NumberOfCores, NumberOfLogicalProcessors
                $global:CpuName = $global:cpu.Name
                $global:CpuCores = $global:cpu.NumberOfCores
                $global:CpuLogical = $global:cpu.NumberOfLogicalProcessors
            }
        Else
            {
                $global:CpuName = "Station Shutdown"
                $global:CpuCores = "Station Shutdown"
                $global:CpuLogical = "Station Shutdown"
            }
    }

Function FindWMI
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:wmi = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_WmiSetting} | Select BuildVersion
                $global:BuildVersion = $global:wmi.BuildVersion
            }
        Else
            {
                $global:BuildVersion = "Station Shutdown"
            }
    }

Function FindWin10Ver
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:newSysBuild = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseId | select ReleaseId}
                $global:Win10Ver = $global:newSysBuild.ReleaseId
            }
        Else
            {
                $global:Win10Ver = "Station Shutdown"
            }
    }


Function FindWin10VerNew
    {
        $regkeypath= "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\" 
        $value1 = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path $regkeypath}.ReleaseId -eq $null
        if ($value1 -eq $False)
            {
                $global:newSysBuild = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseId | select ReleaseId}
                $global:Win10Ver = $global:newSysBuild.ReleaseId
            }
        Elseif ($value1 -eq $False) 
            {
                $global:Win10Ver = "Not Win10"
            }
        Else
            {
                 $global:Win10Ver = "Station Offline"
            }
    }


Function FindNetwork
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:networks = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_NetworkAdapterConfiguration | ? {$_.IPEnabled}} | Select IPAddress, MACAddress
                $global:IPAddress = $global:Networks.IPAddress[0]
                $global:MACAddress = $global:Networks.MACAddress
            }
        Else
            {
                $global:IPAddress = "Station Shutdown"
                $global:MACAddress = "Station Shutdown"
            }
    }

Function FindVol
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:vol = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_volume} | Select Caption, driveletter, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)}  }, @{LABEL='HDDCapacity';EXPRESSION={"{0:N2}" -f($_.Capacity/1GB)}  }| Where-Object { $_.driveletter -match "C:" } | where {$_.Caption -like "C:\*"} 
                $global:HDD = $global:vol.HDDCapacity
                $global:HDDFree = $global:vol.GBfreespace
            }
        Else
            {
                $global:HDD = "Station Shutdown"
                $global:HDDFree = "Station Shutdown"
            }
    }

Function FindVideoController($OP)
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:gpuDriver = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_VideoController} 
            }
        Else
            {

            }
    }

Function FindNvidia
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:nvidiaID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*nvidia*"} | Select-Object -first 1
                $global:Nvidia = $global:nvidia.HardWareID 
            }
        Else
            {
                $global:Nvidia = "Station Shutdown"
            }
    }

Function FindIntel
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:intelID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Graphics*"} | Select-Object -first 1
                $global:intel = $global:intelID.HardWareID
            }
        Else
            {
                $global:intel = "Station Shutdown"
            }
    }
    
Function FindBasic
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:MicroID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Basic Display Ad*"} | Select-Object -first 1
                $global:MSBasic = $global:MicroID.HardWareID 
            }
        Else
            {
                $global:MSBasic = "Station Shutdown"
            }
    }
    
Function FindEco965
    {
       if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:EcoID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Intel(R) 965 Express*"} | Select-Object -first 1 
                $global:Eco965 = $global:EcoID.HardWareID
            }
        Else
            {
                $global:Eco965 = "Station Shutdown"
            }
    }
    
Function FindLastUser
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:username = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\Software\Microsoft\windows\currentVersion\Authentication\LogonUI\' -Name LastLoggedOnUser | select LastLoggedOnUser}
                $global:User = $global:username.LastLoggedOnUser
            }
        Else
            {
                $global:User = "Station Shutdown"
            }
    }
   
Function FindCurrentLocation
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:location = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\RM\Connect\' -Name CurrentLocation | select CurrentLocation} 
                $global:StationLocation = $global:location.CurrentLocation
            }
        Else
            {
                $global:StationLocation = "Station Shutdown"
            }
    }
    
Function FindLastLoggedOn
    {
        if (Test-Connection -ComputerName $livePC  -Quiet -count 1)
            {
                $global:lastlogtime = Get-ADComputer -identity $livePC -Properties * | select LastLogonDate
                $global:Lastlogontime = $global:lastlogtime.LastLogonDate
            }
        Else
            {
                $global:Lastlogontime = "Station Shutdown"
            }
    }
    
Function CheckExisting
    {
        If(!(test-path $path))
            {
                New-Item -ItemType Directory -Force -Path $path
                Write-Host -foregroundcolor Green "$path has been created."
            }
        Else
            {
                Write-Host -foregroundcolor DarkYellow "$path has ALREADY been created."
            }

#Deletes livePC's text 
        if (Test-Path $fileName) 
            {
                Remove-Item $fileName
                Write-Host -foregroundcolor DarkYellow "The old $fileName has been deleted."
            }
        Else
            {
                Write-Host -foregroundcolor Darkyellow "$fileName does not exist..."
            }
        #Removes incomplete $csv 
        if (Test-Path $exportLocation) 
            {
                Remove-Item $exportLocation
                Write-Host -foregroundcolor Red "The old $exportLocation has been deleted."
            }
        Else
            {
                Write-Host -foregroundcolor Darkyellow "$exportLocation does not exist..."
            }
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
$testcomputers = Get-ADComputer -Filter {(OperatingSystem -notlike "*windows*server*")}
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
#Can Change These 
$path = "D:\temp\Hardware"
$csvName = "pcInventory.csv"
$livePCs = "livePCs.txt"
#Hard-coded
$fileName = "$path\$livePCs"
$exportLocation = "$path\$csvName"
$ResultObject = New-Object System.Collections.Generic.List[object]

CheckExisting

write-host `r
write-host -foregroundcolor Yellow "Checking online status of $test_computer_count computers, this may take a while."
# PING COMPUTERS TO FIND OUT ONLINE/OFFLINE LIST
foreach ($computer in $testcomputers) {
        # I only send 2 echo requests to speed things up, if you want the defaut 4 
        # delete the -count 2 portion
   if (Test-Connection -ComputerName $computer.Name -Quiet -count 1){
        # The path to the livePCs.txt file, change to meet your needs
        Add-Content -value $computer.Name -path $fileName
        $xLive++;
        $xTotal++;
        Write-Host -foregroundcolor Green "$xTotal. $($computer.Name) has been added to the Live List. $xLive of $test_computer_count"
        }else{
        # The path to the deadPCs.txt file, change to meet your needs
        Add-Content -value $computer.Name -path $path\deadPCs.txt
        $xDead++;
        $xTotal++;
        Write-Host -foregroundcolor Red "$xTotal. $($computer.Name) cannot be contacted. $xDead of $test_computer_count"
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
 
$ComputerName = gc -Path "$fileName"
 
$computer_count = $ComputerName.Length;
# The results of the script are here
 
$i = 0;
$p = 0;
#$livePC = "BRKAD002"
foreach ($livePC in $ComputerName){
    if (Test-Connection -ComputerName $livePC -Quiet -count 1){
    $p++;
    write-host -foregroundcolor Green "Probing $livePC - $p of $($ComputerName.Count)"
    $computer_progress = [int][Math]::Ceiling((($i / $computer_count) * 100))
	# Progress bar
    Write-Progress  "Gathering Hardware Info" -PercentComplete $computer_progress -Status "Percent Complete - $computer_progress%" -Id 1;
	Sleep(1);
    $i++;

    FindOS
    FindBIOS
    FindcompSys
    FindProcessor
    FindWMI
    FindWin10Ver
    FindNetwork
    FindVol
    
    #Graphics Adapters
    FindVideoController
    #Individual Graphics Devices
    FindNvidia
    FindIntel
    FindBasic
    FindEco965

    #Registry Commands
    FindLastUser
    FindCurrentLocation
    
    #Active Directory
    FindLastLoggedOn
    
    $OutputObj  = New-Object -Type PSObject
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $livePC.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $global:Manufacturer
    $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $global:PCModel
    $OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $global:systemBios
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU -Value $global:CpuName
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Cores -Value $global:CpuCores
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_LogicalCores -Value $global:CpuLogical
    #$OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Socket -Value $cpu.SocketDesignation
    #Doesn't always work. Some Show as CPU0/CPU1
    $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Physical_Memory -Value $global:totalMemory
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:GBfreeSpace -Value $global:HDDFree
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:Size -Value $global:HDD
    $OutputObj | Add-Member -MemberType NoteProperty -Name Microsoft_ID -Value $global:MSBasic
    $OutputObj | Add-Member -MemberType NoteProperty -Name Intel_ID -Value $global:intel
    $OutputObj | Add-Member -MemberType NoteProperty -Name Eco965_ID -Value $global:Eco965
    $OutputObj | Add-Member -MemberType NoteProperty -Name nVidia_ID -Value $global:Nvidia
    Foreach ($Card in $gpuDriver)
        {
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Name" $Card.Name
        #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Description" $Card.Description #Probably not needed. Seems to just echo the name. Left here in case I'm wrong!
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Vendor" $Card.AdapterCompatibility
        #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_PNPDeviceID" $Card.PNPDeviceID
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_DriverVersion" $Card.DriverVersion
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_VideoMode" $Card.VideoModeDescription
        }
    #$OutputObj | Add-Member -MemberType NoteProperty -Name Installed_Video -Value ""
    $OutputObj | Add-Member -MemberType NoteProperty -Name OS -Value $global:WindowsOS
    $OutputObj | Add-Member -MemberType NoteProperty -Name SystemType -Value $global:WindowsOSArchitecture
    $OutputObj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $global:BuildVersion
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
                'RAM' = $global:totalMemory
                'OS' = $global:WindowsOS                
                'HDD Size' = $global:HDD
                'HDD Free' = $global:HDDFree
                'Video Card' = $Card.Name
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
    }

    $ResultObject  | Out-Gridview -Title "Hardware Information"
    <##
    Write-Host -ForegroundColor Yellow "$computer has been accounted for"
    Write-Host -ForegroundColor Green "$IPAddress"
    Write-Host -ForegroundColor Green "$Computer"
    Write-Host -ForegroundColor Green "$buildDate"
    Write-Host -ForegroundColor Green "$MACAddress"
    Write-Host -ForegroundColor Green "$buildDate"
    Write-Host -ForegroundColor Green "$lastboot"
    Write-Host -ForegroundColor Green "$OS.Caption"
    Write-Host -ForegroundColor Green "$cpu.Name"
    Write-Host -ForegroundColor Green "$totalMemory"
   # Write-Host -ForegroundColor Green "$GPU"
   Write-Host -ForegroundColor Green "$($location.CurrentLocation)"
   ##>
   #Write-Host -ForegroundColor Green "$($buildDate.InstallDate)"
   
   
 
  
write-host -foregroundcolor cyan "Renaming $csvName" 
 
$newCSV = Get-ChildItem $exportLocation 
$time = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")    
$RenameCSV = "$($newCSV.DirectoryName)\$($newCSV.BaseName)[$($time)]$($newCSV.Extension)"
Rename-Item $exportlocation -NewName $RenameCSV    


write-host -foregroundcolor cyan "$csvName has been renamed to $RenameCSV)"
write-host -foregroundcolor cyan "Script is complete, the results are here: $path"
Write-host -foregroundcolor green "Script started at $startTime and finished at $time"

Invoke-Item $path
