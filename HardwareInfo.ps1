<##
SCRIPT CREATE BY DANIEL KILVERT.
SCRIPT WILL GATHER ALL HARDWARE INFORMATION ON COMPUTERS OBTAINED VIA A CSV OR ACTIVE DIRECTORY.
TO USE A CSV FILE, UNCOMMENT OUT THE LOCATION PICKER AND COMMENT OUT THE AD FILTER
    SAVE LOCATOIN CAN BE MODIFIED USING THE $path VARIABLE
    TO CHANGE THE OUTPUT FILENAME, THIS CAN BE CHANGED AT $exportLocation
SCRIPT WILL DELETE PREVIOUS livePC list to make sure no entries are duplicated.
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
 
#Check Path to Hardware folder and create if it dows not exist.
$path = "D:\temp\Hardware"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}
 
<#
CAN I ADD AND ELSE HERE? TEST ME
Else 
{
    Remove-Item $path\*.txt
}
#>
 
#Check if livePC's txt file already exists and delete if true
$fileName = "$path\livePCs.txt"
if (Test-Path $fileName) 
{
  Remove-Item $fileName
}
 
$exportLocation = "$path\pcInventory.csv"
 
write-host `r
write-host -foregroundcolor cyan "Testing $test_computer_count computers, this may take a while."
 
foreach ($computer in $testcomputers) {
        # I only send 2 echo requests to speed things up, if you want the defaut 4 
        # delete the -count 2 portion
   if (Test-Connection -ComputerName $computer.DNSHostName -Quiet -count 2){
        # The path to the livePCs.txt file, change to meet your needs
        Add-Content -value $computer.DNSHostName -path $path\livePCs.txt
        Write-Host "$computer has been added to the Live List."
        }else{
        # The path to the deadPCs.txt file, change to meet your needs
        Add-Content -value $computer.DNSHostName -path $path\deadPCs.txt
        Write-Host "$computer cannot be contacted."
        }
    $testcomputer_progress = [int][Math]::Ceiling((($x / $test_computer_count) * 100))
	# Progress bar
    Write-Progress  "Testing Connections" -PercentComplete $testcomputer_progress -Status "Percent Complete - $testcomputer_progress%" -Id 1;
	Sleep(1);
    $x++;
 
} 
 
write-host `r
write-host -foregroundcolor cyan "Testing Connection complete"
write-host `r
 
$ComputerName = gc -Path "$path\livePCs.txt"
 
$computer_count = $ComputerName.Length;
# The results of the script are here
 
$i = 0;
 foreach ($Computer in $ComputerName){
   $Bios =get-wmiobject win32_bios -Computername $Computer
   $Hardware = get-wmiobject Win32_computerSystem -Computername $Computer
   $Sysbuild = get-wmiobject Win32_WmiSetting -Computername $Computer
   $OS = gwmi Win32_OperatingSystem -Computername $Computer
   $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | ? {$_.IPEnabled}
   $driveSpace = gwmi win32_volume -computername $Computer -Filter 'drivetype = 3' | 
   select PScomputerName, driveletter, label, @{LABEL='GBfreespace';EXPRESSION={"{0:N2}" -f($_.freespace/1GB)} } |
   Where-Object { $_.driveletter -match "C:" }
   $cpu = Get-WmiObject Win32_Processor  -computername $computer
   $username = Get-ChildItem "\\$computer\c$\Users" | Sort-Object LastWriteTime -Descending | Select Name, LastWriteTime -first 1
   $totalMemory = [math]::round($Hardware.TotalPhysicalMemory/1024/1024/1024, 2)
   $lastBoot = $OS.ConvertToDateTime($OS.LastBootUpTime)
   $buildDate = ([WMI]'').ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate)
   $gpuDriver = Get-WmiObject Win32_VideoController -ComputerName $Computer
    
   #write-host -foregroundcolor yellow "Found $computer"
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
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $Hardware.Manufacturer
    $OutputObj | Add-Member -MemberType NoteProperty -Name Model -Value $Hardware.Model
    $OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $systemBios
    $OutputObj | Add-Member -MemberType NoteProperty -Name CPU_Info -Value $cpu.Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name Total_Physical_Memory -Value $totalMemory
    $OutputObj | Add-Member -MemberType NoteProperty -Name C:_GBfreeSpace -Value $driveSpace.GBfreespace
    $OutputObj | Add-Member -MemberType NoteProperty -Name Installed_Video -Value ""
    $OutputObj | Add-Member -MemberType NoteProperty -Name OS -Value $OS.Caption
    $OutputObj | Add-Member -MemberType NoteProperty -Name SystemType -Value $Hardware.SystemType
    $OutputObj | Add-Member -MemberType NoteProperty -Name BuildVersion -Value $SysBuild.BuildVersion
    #$OutputObj | Add-Member -MemberType NoteProperty -Name SPVersion -Value $OS.csdversion
    Foreach ($Card in $gpuDriver)
        {
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Name" $Card.Name
        #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Description" $Card.Description #Probably not needed. Seems to just echo the name. Left here in case I'm wrong!
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_Vendor" $Card.AdapterCompatibility
        #$OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_PNPDeviceID" $Card.PNPDeviceID
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_DriverVersion" $Card.DriverVersion
        $OutputObj | Add-Member NoteProperty "$($Card.DeviceID)_VideoMode" $Card.VideoModeDescription
        }
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last-Login -Value $username.LastWriteTime
    $OutputObj | Add-Member -MemberType NoteProperty -Name LastUser -Value $username.Name
    $OutputObj | Add-Member -MemberType NoteProperty -Name Last_Reboot -Value $lastboot
    $OutputObj | Add-Member -MemberType NoteProperty -Name Build_Date -Value $buildDate
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
   ##>
   }
} 
 
 write-host -foregroundcolor cyan "Script is complete, the results are here: $exportLocation" 
