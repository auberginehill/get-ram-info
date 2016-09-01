<#
Get-RAMInfo.ps1
#>


$path = $env:temp
$computer = $env:COMPUTERNAME
$timestamp = Get-Date -UFormat "%Y%m%d"
$empty_line = ""


# Function used to convert bytes to MB or GB or TB                                            # Credit: clayman2: "Disk Space"
function ConvertBytes {
    param($size)
    If ($size -lt 1MB) {
        $drive_size = $size / 1KB
        $drive_size = [Math]::Round($drive_size, 2)
        [string]$drive_size + ' KB'
    } ElseIf ($size -lt 1GB) {
        $drive_size = $size / 1MB
        $drive_size = [Math]::Round($drive_size, 2)
        [string]$drive_size + ' MB'
    } ElseIf ($size -lt 1TB) {
        $drive_size = $size / 1GB
        $drive_size = [Math]::Round($drive_size, 2)
        [string]$drive_size + ' GB'
    } Else {
        $drive_size = $size / 1TB
        $drive_size = [Math]::Round($drive_size, 2)
        [string]$drive_size + ' TB'
    } # else
} # function




$obj_memory = @()
$memory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $computer

    ForEach ($memblock in $memory) {


        # Memory type
        Switch ($memblock.FormFactor) {
            { $_ -lt 0 } { $memory_type = "" }              
            { $_ -eq 0 } { $memory_type = "Unknown" }            
            { $_ -eq 1 } { $memory_type = "Other" }
            { $_ -eq 2 } { $memory_type = "SIP" }
            { $_ -eq 3 } { $memory_type = "DIP " }
            { $_ -eq 4 } { $memory_type = "ZIP" }
            { $_ -eq 5 } { $memory_type = "SOJ" }
            { $_ -eq 6 } { $memory_type = "Proprietary" }
            { $_ -eq 7 } { $memory_type = "SIMM" }
            { $_ -eq 8 } { $memory_type = "DIMM" }
            { $_ -eq 9 } { $memory_type = "TSOP" }
            { $_ -eq 10 } { $memory_type = "PGA" }
            { $_ -eq 11 } { $memory_type = "RIMM" }
            { $_ -eq 12 } { $memory_type = "SODIMM" }
            { $_ -eq 13 } { $memory_type = "SRIMM" }
            { $_ -eq 14 } { $memory_type = "SMD" }
            { $_ -eq 15 } { $memory_type = "SSMP" }
            { $_ -eq 16 } { $memory_type = "QFP" }
            { $_ -eq 17 } { $memory_type = "TQFP" }
            { $_ -eq 18 } { $memory_type = "SOIC" }
            { $_ -eq 19 } { $memory_type = "LCC" }
            { $_ -eq 20 } { $memory_type = "PLCC" }
            { $_ -eq 21 } { $memory_type = "BGA" }
            { $_ -eq 22 } { $memory_type = "FPBGA" }
            { $_ -eq 23 } { $memory_type = "LGA" }
            { $_ -gt 23 } { $memory_type = "" }            
        } # switch formfactor




        # Type Detail
        Switch ($memblock.TypeDetail) {
            { $_ -lt 1 } { $type_detail = "" }              
            { $_ -eq 1 } { $type_detail = "Reserved" }            
            { $_ -eq 2 } { $type_detail = "Other" }
            { $_ -eq 4 } { $type_detail = "Unknown" }
            { $_ -eq 8 } { $type_detail = "Fast-paged" }
            { $_ -eq 16 } { $type_detail = "Static column" }
            { $_ -eq 32 } { $type_detail = "Pseudo-static" }
            { $_ -eq 64 } { $type_detail = "RAMBUS" }
            { $_ -eq 128 } { $type_detail = "Synchronous" }
            { $_ -eq 256 } { $type_detail = "CMOS" }
            { $_ -eq 512 } { $type_detail = "EDO" }
            { $_ -eq 1024 } { $type_detail = "Window DRAM" }
            { $_ -eq 2048 } { $type_detail = "Cache DRAM" }
            { $_ -eq 4096 } { $type_detail = "Non-volatile" }
            { $_ -gt 4096 } { $type_detail = "" }              
        } # switch typedetail




        $obj_memory += New-Object -TypeName PSCustomObject -Property @{


                        'Capacity (GB)'     = $memblock.Capacity / 1GB
                        'Class'             = $memblock.Name
                        'Computer'          = $memblock.__SERVER
                        'Location'          = $memblock.DeviceLocator
                        'Manufacturer'      = $memblock.Manufacturer
                        'Memory Type'       = $memory_type                                
                        'Part Number'       = $memblock.PartNumber
                        'RAM Type'          = $memory_type + ' ' + (ConvertBytes($memblock.Capacity)) + ' (' + ($memblock.Speed) + ' MHz)'
                        'Serial Number'     = $memblock.SerialNumber
                        'Speed'             = [string]($memblock.Speed) + ' MHz'
                        'Type Detail'       = $type_detail


                    } # New-Object
                $obj_memory.PSObject.TypeNames.Insert(0,"Memory")
                $obj_memory_selection = $obj_memory | Select-Object 'Computer','Location','RAM Type','Manufacturer','Part Number','Class','Serial Number'


    } # foreach


# Write the memory block results in console
Write-Output $empty_line
Write-Output $empty_line
Write-Output $empty_line
Write-Output $obj_memory_selection | Format-Table -auto




# Gather some data for a first summary table
$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer 
$used_memory_perc = $os | Select-Object @{Label='UsedMemoryPerc'; 
                                Expression={"{0:N1}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) / $_.TotalVisibleMemorySize) * 100) }}

$used_memory = $os | Select-Object @{Label='UsedMemory'; 
                                Expression={(($_.TotalVisibleMemorySize * 1kb) - ($_.FreePhysicalMemory) * 1kb) }}

$available_memory_perc = $os | Select-Object @{Label='AvailableMemoryPerc'; 
                                Expression={"{0:N1}" -f ((($os.FreePhysicalMemory) / ($os.TotalVisibleMemorySize)) * 100) }}

$arrays = Get-WmiObject -Class Win32_PhysicalMemoryArray -computerName $computer
$number_of_arrays = ($arrays | Measure-Object).Count
$slots = 0

    ForEach ($array In $arrays) { $slots += $array.MemoryDevices }




$summary_table = $obj_memory | Measure-Object -Property 'Capacity (GB)' -Sum | Select-Object -Property @{Label='Computer'; Expression={$computer}},@{Label='Total Slots'; Expression={


                                If ($slots -ge 2) {
                                    [string]$slots + ' Slots'
                                } ElseIf ($slots -eq 1) {
                                    [string]'1 Slot'
                                } ElseIf ($slots -eq 0) {
                                    [string]'Did not detected any memory slots.'
                                } Else {
                                    [string]''
                                } # else


        }},@{Label='Slots in Use'; Expression={


                                If ($_.count -ge 2) {
                                    [string]$_.count + ' Slots'
                                } ElseIf ($_.count -eq 1) {
                                    [string]'1 Slot'
                                } ElseIf ($_.count -eq 0) {
                                    [string]'All memory slots seem to be empty.'
                                } Else {
                                    [string]''
                                } # else


        }},@{Label='Free Slots'; Expression={
            
            
                                If ($slots - $_.count -ge 2) {
                                    [string]$slots - $_.count + ' Slots'
                                } ElseIf ($slots - $_.count -eq 1) {
                                    [string]'1 Slot'
                                } ElseIf ($slots - $_.count -eq 0) {
                                    [string]'None'
                                } Else {
                                    [string]''
                                } # else     
            
            
            
        }},@{Label='Total Memory'; 
            Expression={[string]$_.sum + ' GB'}},
        @{Label='Memory in Use'; 
            Expression={"$(ConvertBytes($($used_memory.UsedMemory)))" + " ($($used_memory_perc.UsedMemoryPerc) %)"}}, 
        @{Label='Available Memory'; 
            Expression={"$(ConvertBytes($os.FreePhysicalMemory * 1kb))" + " ($($available_memory_perc.AvailableMemoryPerc) %)"}}


# Write the first summary table in console
Write-Output $empty_line
Write-Output $summary_table | Format-Table -auto




# Gather some data for a second summary table
$gps = Get-Process | Measure-Object -Property ProcessName
$average_load = Get-WmiObject -Class Win32_Processor -ComputerName $computer | Measure-Object -property LoadPercentage -Average
$used_perc = Get-WmiObject -Class Win32_Volume -ComputerName $computer -Filter "DriveLetter = 'C:'" | Select-Object @{Label='C_Drive'; 
                                Expression={"{0:N1}" -f ((($_.Capacity - $_.FreeSpace) / $_.Capacity) * 100) }}


# Write the second summary table in console
Write-Output $empty_line
Write-Output $empty_line
Write-Output "Processes: $($gps.Count)         Average CPU Load: $($average_load.Average) %         C:-Drive Usage: $($used_perc.C_Drive) %        Physical Memory in Use: $($used_memory_perc.UsedMemoryPerc) %"
Write-Output $empty_line
Write-Output $empty_line
Write-Output $empty_line
Write-Output $empty_line



# [End of Line]


<#

   ____        _   _
  / __ \      | | (_)
 | |  | |_ __ | |_ _  ___  _ __  ___
 | |  | | '_ \| __| |/ _ \| '_ \/ __|
 | |__| | |_) | |_| | (_) | | | \__ \
  \____/| .__/ \__|_|\___/|_| |_|___/
        | |
        |_|



# Write the memory info to a CSV-file
$obj_memory_selection | Export-Csv $path\memory_info.csv -Delimiter ';' -NoTypeInformation -Encoding UTF8


# Open the memory info CSV-file
Invoke-Item -Path $path\memory_info.csv


memory_info_$timestamp.csv                                                                    # an alternative filename format
$time = Get-Date -Format g                                                                    # a "general short" time-format (short date and short time)



   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


http://powershell.com/cs/media/p/7476.aspx                                                    # clayman2: "Disk Space"



  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#

.SYNOPSIS
Retrieves basic memory information.

.DESCRIPTION
Get-RAMInfo uses Windows Management Instrumentation (WMI) to retrieve basic
memory information and displays the results in console.

.OUTPUTS
Displays general memory information, such as used Memory Slots (Location) and RAM Type, Capacity, 
Speed, Manufacturer, Part Number, Memory Class and Serial Number of individual Memory Modules and 
also Total number of Memory Slots, Total number of Memory Slots in Use, Total Amount of Free Memory 
Slots, Total Physical Memory and both Memory in Use and Available Memory as Size and as Percentage, 
Number of Processes running, Average CPU Load, Physical Memory in Use and C:-Drive Usage in console.

.NOTES
Please note that the optional file listed under Options-header will(, if the option is enabled by
copy-pasting the relevant code above the [End of Line] -marker) be created in a directory, which is
specified with the $path variable (at line 6). The $env:temp variable points to the current temp
folder. The default value of the $env:temp variable is C:\Users\<username>\AppData\Local\Temp
(i.e. each user account has their own separate temp folder at path %USERPROFILE%\AppData\Local\Temp).
To see the current temp path, for instance a command
    [System.IO.Path]::GetTempPath()
may be used at the PowerShell prompt window [PS>]. To change the temp folder for instance 
to C:\Temp, please, for example, follow the instructions at
http://www.eightforums.com/tutorials/23500-temporary-files-folder-change-location-windows.html

    Homepage:           https://github.com/auberginehill/get-ram-info
    Version:            1.0

.EXAMPLE
./Get-RAMInfo
Run the script. Please notice to insert ./ or .\ before the script name.

.EXAMPLE
help ./Get-RAMInfo -Full
Display the help file.

.EXAMPLE
Set-ExecutionPolicy remotesigned
This command is altering the Windows PowerShell rights to enable script execution. Windows PowerShell
has to be run with elevated rights (run as an administrator) to actually be able to change the script
execution properties. The default value is "Set-ExecutionPolicy restricted".


    Parameters:

    Restricted      Does not load configuration files or run scripts. Restricted is the default
                    execution policy.

    AllSigned       Requires that all scripts and configuration files be signed by a trusted
                    publisher, including scripts that you write on the local computer.

    RemoteSigned    Requires that all scripts and configuration files downloaded from the Internet
                    be signed by a trusted publisher.

    Unrestricted    Loads all configuration files and runs all scripts. If you run an unsigned
                    script that was downloaded from the Internet, you are prompted for permission
                    before it runs.

    Bypass          Nothing is blocked and there are no warnings or prompts.

    Undefined       Removes the currently assigned execution policy from the current scope.
                    This parameter will not remove an execution policy that is set in a Group
                    Policy scope.


For more information,
type "help Set-ExecutionPolicy -Full" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-RAMInfo.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent -NoClobber mode
built into it, so that the procedure will halt, if overwriting (replacing the contents) of an existing
file is about to happen. Overwriting a file with the New-Item cmdlet requires using the Force.
For more information, please type "help New-Item -Full".

.LINK
http://powershell.com/cs/media/p/7476.aspx
http://stackoverflow.com/questions/37756770/getting-ram-info-powershell
https://msdn.microsoft.com/en-us/library/aa394347(v=vs.85).aspx

#>
