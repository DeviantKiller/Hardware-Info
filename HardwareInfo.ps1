  <##
SCRIPT CREATE BY DANIEL KILVERT.
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
	
##>
 
# On error the script will continue silently without 
$erroractionpreference = "SilentlyContinue"
 
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
 
$path = "D:\temp\Hardware"
$csvName = "pcInventory.csv"
$exportLocation = "$path\$csvName"
$fileName = "$path\livePCs.txt"

#Creates Hardware Folder if non existing
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}
#Deletes livePC's text 
if (Test-Path $fileName) 
{
  Remove-Item $fileName
}
 
write-host `r
write-host -foregroundcolor Yellow "Checking online status of $test_computer_count computers, this may take a while."
 
foreach ($computer in $testcomputers) {
        # I only send 2 echo requests to speed things up, if you want the defaut 4 
        # delete the -count 2 portion
   if (Test-Connection -ComputerName $computer.Name -Quiet -count 1){
        # The path to the livePCs.txt file, change to meet your needs
        Add-Content -value $computer.Name -path $path\livePCs.txt
        Write-Host -foregroundcolor Green "$($computer.Name) has been added to the Live List."
        }else{
        # The path to the deadPCs.txt file, change to meet your needs
        Add-Content -value $computer.Name -path $path\deadPCs.txt
        Write-Host -foregroundcolor Red "$($computer.Name) cannot be contacted."
        }
    $testcomputer_progress = [int][Math]::Ceiling((($x / $test_computer_count) * 100))
	# Progress bar
    Write-Progress  "Testing Connections" -PercentComplete $testcomputer_progress -Status "Percent Complete - $testcomputer_progress%" -Id 1;
	Sleep(1);
    $x++;
 
} 
 
write-host `r
write-host -foregroundcolor cyan "Testing Connection complete"
write-host -foregroundcolor cyan "Writing Data to CSV."
write-host `r
 
$ComputerName = gc -Path "$path\livePCs.txt"
 
$computer_count = $ComputerName.Length;
# The results of the script are here
 
$i = 0;
foreach ($livePC in $ComputerName){
    $Bios = get-wmiobject win32_bios -Computername $livePC
    $Hardware = get-wmiobject Win32_computerSystem -Computername $livePC
    $Sysbuild = get-wmiobject Win32_WmiSetting -Computername $livePC
    $OS = gwmi Win32_OperatingSystem -Computername $livePC
    $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $livePC | ? {$_.IPEnabled}
    $driveSpace = gwmi win32_volume -computername $livePC -Filter 'drivetype = 3' | 
    select PScomputerName, driveletter, label, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)} } |
    Where-Object { $_.driveletter -match "C:" }
    $cpu = Get-WmiObject Win32_Processor  -computername $livePC
    $username = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\Software\Microsoft\windows\currentVersion\Authentication\LogonUI\' -Name LastLoggedOnUser | select LastLoggedOnUser}
    $location = Invoke-Command -ComputerName $livePC {Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\RM\Connect\' -Name CurrentLocation | select CurrentLocation}
    $lastlogtime = Get-ADComputer -identity $livePC -Properties * | select LastLogonDate
    $totalMemory = [math]::round($Hardware.TotalPhysicalMemory/1024/1024/1024, 2)
    $lastBoot = $OS.ConvertToDateTime($OS.LastBootUpTime)
    $buildDate = ([WMI]'').ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate)
    #Graphics Adapters
    $gpuDriver = Get-WmiObject Win32_VideoController -ComputerName $livePC
    #Individual Graphics Devices
    $nvidiaID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*nvidia*"}
    $intelID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Graphics*"}
    $MicroID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Basic Display Ad*"}
    $EcoID = Get-WmiObject Win32_PnPSignedDriver -ComputerName $livePC| select devicename, HardWareID | where {$_.devicename -like "*Intel(R) 965 Express*"}

    $computer_progress = [int][Math]::Ceiling((($i / $computer_count) * 100))
	# Progress bar
    Write-Progress  "Gathering Hardware Info" -PercentComplete $computer_progress -Status "Percent Complete - $computer_progress%" -Id 1;
	Sleep(1);
    $i++;
    foreach ($Network in $Networks) {
    $IPAddress  = $Network.IpAddress[0]
    $MACAddress  = $Network.MACAddress
    $systemBios = $Bios.serialnumber
    $OutputObj  = New-Object -Type PSObject
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $livePC.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $Hardware.Manufacturer
    $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $Hardware.Model
    $OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $systemBios
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Info -Value $cpu.Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Physical_Memory -Value $totalMemory
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:_GBfreeSpace -Value $driveSpace.GBfreespace
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
    $OutputObj | Add-Member -MemberType NoteProperty -Name SystemType -Value $Hardware.SystemType
    $OutputObj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $SysBuild.BuildVersion
    #$OutputObj | Add-Member -MemberType NoteProperty -Name SPVersion -Value $OS.csdversion
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
    #$OutputObj | Add-Member -MemberType NoteProperty -Name Last-Login -Value $username.LastWriteTime
    $OutputObj | Add-Member -MemberType NoteProperty -Name LastUser -Value $($username.LastLoggedOnUser)
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Logon -Value $($lastlogtime.LastLogonDate)
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Reboot -Value $lastboot
    $OutputObj | Add-Member -MemberType NoteProperty -Name Build_Date -Value $buildDate
    $OutputObj | Add-Member -MemberType NoteProperty -Name Location -Value $($location.CurrentLocation) #OU information for CSV
    $OutputObj | Export-Csv $exportLocation -Append -NoTypeInformation

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
   }
 
  }
write-host -foregroundcolor cyan "Renaming $csvName" 
 
$newCSV = Get-ChildItem $exportLocation 
$time = (Get-Date).tostring("dd-MM-yyyy-HH.mm.ss")    
$RenameCSV = "$($newCSV.DirectoryName)\$($newCSV.BaseName)[$($time)]$($newCSV.Extension)"
Rename-Item $exportlocation -NewName $RenameCSV    

write-host -foregroundcolor cyan "$csvfile has been renamed to $RenameCSV)"
write-host -foregroundcolor cyan "Script is complete, the results are here: $path"

