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
1.8     Attempt of usage of Out-Gridview
	
##>
 
# On error the script will continue silently without 
$erroractionpreference = "SilentlyContinue"
$startTime = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")  
<# Using a text file with a location picker
#
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
$x = 0;
$xDead = 0;
$xLive = 0;
$xTotal = 0;
 
$path = "D:\temp\Hardware"
$csvName = "pcInventory.csv"
$exportLocation = "$path\$csvName"
$fileName = "$path\livePCs.txt"


#Creates Hardware Folder if non existing
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
      Write-Host -foregroundcolor Green "$path has been created."
}
#Deletes livePC's text 
if (Test-Path $fileName) 
{
    Remove-Item $fileName
    Write-Host -foregroundcolor Green "The old $fileName has been deleted."
}
 
write-host `r
write-host -foregroundcolor Yellow "Checking online status of $test_computer_count computers, this may take a while."
 
foreach ($computer in $testcomputers) {
        # I only send 2 echo requests to speed things up, if you want the defaut 4 
        # delete the -count 2 portion
   if (Test-Connection -ComputerName $computer.Name -Quiet -count 1){
        # The path to the livePCs.txt file, change to meet your needs
        Add-Content -value $computer.Name -path $path\livePCs.txt
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
 
$ComputerName = gc -Path "$path\livePCs.txt"
 
$computer_count = $ComputerName.Length;
# The results of the script are here
 
$i = 0;
$p = 0;
#$livePC = "RIGDT143"
foreach ($livePC in $ComputerName){
    $p++;
    write-host -foregroundcolor Green "Probing $livePC - $p of $($ComputerName.Count)"
    $computer_progress = [int][Math]::Ceiling((($i / $computer_count) * 100))
	# Progress bar
    Write-Progress  "Gathering Hardware Info" -PercentComplete $computer_progress -Status "Percent Complete - $computer_progress%" -Id 1;
	Sleep(1);
    $i++;

    $OS = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_OperatingSystem} | Select InstallDate, OSArchitecture, Caption, LastBootUpTime
        #$OS.InstallDate
        #$OS.OSArchitecture
        #$OS.Caption
        #$OS.LastBootUpTime
    $bios = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance win32_bios} | Select SerialNumber
        #$bios.SerialNumber
    $compSys = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_computerSystem} | Select SerialNumber, Manufacturer, Model, TotalPhysicalMemory, SystemType
        #$compSys.Manufacturer
        #$compSys.Model
        #$compSys.SystemType
        #$totalMemory = [math]::round($compSys.TotalPhysicalMemory/1024/1024/1024, 2)
    $cpu = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_Processor} | Select Name, NumberOfCores, NumberOfLogicalProcessors
        #$cpu.Name
        #$cpu.NumberOfCores
        #$cpu.NumberOfLogicalProcessors
    $wmi = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_WmiSetting} | Select BuildVersion
        #$wmi.BuildVersion

    $newSysBuild = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseId | select ReleaseId}
        #$newSysBuild.ReleaseId
    $buildDate = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_OperatingSystem} | Select InstallDate
    $networks = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_NetworkAdapterConfiguration | ? {$_.IPEnabled}} | Select IPAddress, MACAddress
        #$network.IPAddress
        #$network.MACAddress
    $vol = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_volume} | Select Caption, driveletter, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)}  }, @{LABEL='HDDCapacity';EXPRESSION={"{0:N2}" -f($_.Capacity/1GB)}  }| Where-Object { $_.driveletter -match "C:" } | where {$_.Caption -like "C:\*"} 
        #$vol.freespace
        #$vol.HDDCapacity
    
    #Graphics Adapters
    $gpuDriver = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_VideoController} 
    #Individual Graphics Devices
    $nvidiaID = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*nvidia*"} | Select-Object -first 1
    $intelID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Graphics*"} | Select-Object -first 1
    $MicroID  = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Basic Display Ad*"} | Select-Object -first 1
    $EcoID    = Invoke-Command -ComputerName $livePC -Scriptblock {Get-CIMInstance Win32_PnPSignedDriver} | select devicename, HardWareID | where {$_.devicename -like "*Intel(R) 965 Express*"} | Select-Object -first 1

    #Registry Commands
    $username = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\Software\Microsoft\windows\currentVersion\Authentication\LogonUI\' -Name LastLoggedOnUser | select LastLoggedOnUser}
    $location = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\RM\Connect\' -Name CurrentLocation | select CurrentLocation}
    
    #Active Directory
    $lastlogtime = Get-ADComputer -identity $livePC -Properties * | select LastLogonDate

    $totalMemory = [math]::round($compSys.TotalPhysicalMemory/1024/1024/1024, 2)
    #$lastBoot = $OS.ConvertToDateTime($OS.LastBootUpTime)

    <##
    OLD METHOD below to be removed and updated to Invoke / Get-CimInstances
    

    #$Bios = get-wmiobject win32_bios -Computername $livePC
    #$Hardware = get-wmiobject Win32_computerSystem -Computername $livePC
    #$Sysbuild = get-wmiobject Win32_WmiSetting -Computername $livePC
    #$OS = gwmi Win32_OperatingSystem -Computername $livePC
    #$Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $livePC | ? {$_.IPEnabled}
    
    $driveSpace = gwmi win32_volume -computername $livePC -Filter 'drivetype = 3' | 
    select PScomputerName, driveletter, label, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)} } |
    Where-Object { $_.driveletter -match "C:" }
   
    #$totalMemory = [math]::round($compSys.TotalPhysicalMemory/1024/1024/1024, 2)
    #$lastBoot = $OS.ConvertToDateTime($OS.LastBootUpTime)
    #Graphics Adapters
    #$gpuDriver = Get-WmiObject Win32_VideoController -ComputerName $livePC
    #Individual Graphics Devices
    #$nvidiaID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*nvidia*"} | Select-Object -first 1
    #$intelID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Graphics*"} | Select-Object -first 1
    #$MicroID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Basic Display Ad*"} | Select-Object -first 1
    #$EcoID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Intel(R) 965 Express*"} | Select-Object -first 1
    ##>

    foreach ($Network in $Networks) {
    $IPAddress  = $Network.IpAddress[0]
    $MACAddress  = $Network.MACAddress
    $systemBios = $Bios.serialnumber
    $OutputObj  = New-Object -Type PSObject
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $livePC.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $compSys.Manufacturer
    $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $compSys.Model
    $OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $($bios.SerialNumber)
    $OutputObj | Add-Member -MemberType NoteProperty -Name New_CPU_Info -Value $($cpu.Name)
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Cores -Value $($cpu.NumberOfCores)
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_LogicalCores -Value $($cpu.NumberOfLogicalProcessors)
    #$OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Socket -Value $cpu.SocketDesignation
    #Doesn't always work. Some Show as CPU0/CPU1
    $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Physical_Memory -Value $totalMemory
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:GBfreeSpace -Value $vol.GBfreespace
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:Size -Value $vol.HDDCapacity
    $OutputObj | Add-Member -MemberType NoteProperty -Name Microsoft_ID -Value $MicroID.HardWareID
    $OutputObj | Add-Member -MemberType NoteProperty -Name Intel_ID -Value $intelID.HardWareID
    $OutputObj | Add-Member -MemberType NoteProperty -Name Intel965_ID -Value $EcoID.HardWareID
    $OutputObj | Add-Member -MemberType NoteProperty -Name nVidia_ID -Value $nvidiaID.HardWareID
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
    $OutputObj | Add-Member -MemberType NoteProperty -Name OS -Value $OS.Caption
    $OutputObj | Add-Member -MemberType NoteProperty -Name SystemType -Value $compSys.SystemType
    $OutputObj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $wmi.BuildVersion
    #$OutputObj | Add-Member -MemberType NoteProperty -Name Windows10Ver -Value $newSysBuild
    #$OutputObj | Add-Member -MemberType NoteProperty -Name SPVersion -Value $OS.csdversion
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name LastUser -Value $($username.LastLoggedOnUser)
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Logon -Value $($lastlogtime.LastLogonDate)
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Reboot -Value $OS.LastBootUpTime
    $OutputObj | Add-Member -MemberType NoteProperty -Name Build_Date -Value $($buildDate.InstallDate)
    $OutputObj | Add-Member -MemberType NoteProperty -Name Location -Value $($location.CurrentLocation) #OU information for CSV
    $OutputObj | Export-Csv $exportLocation -Append -NoTypeInformation

    ForEach-Object{
            $newobj = [PSCustomObject]@{
                'ComputerName' = $livePC.ToUpper()
                'IP Address' = $IPAddress
                'Manufacturer' = $compSys.Manufacturer
                'Model' = $compSys.Model
                'CPU' = $($cpu.Name)
            }
            $ResultObject.Add($newobj)
            }
        }
    } 

    $ResultObject  | Out-Gridview
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
   
   }
 
  }
write-host -foregroundcolor cyan "Renaming $csvName" 
 
$newCSV = Get-ChildItem $exportLocation 
$time = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")    
$RenameCSV = "$($newCSV.DirectoryName)\$($newCSV.BaseName)[$($time)]$($newCSV.Extension)"
Rename-Item $exportlocation -NewName $RenameCSV    


write-host -foregroundcolor cyan "$csvName has been renamed to $RenameCSV)"
write-host -foregroundcolor cyan "Script is complete, the results are here: $path"
Write-host -foregroundcolor green "Script started at $startTime and finished at $time"

