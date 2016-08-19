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

        $obj_memory += New-Object -TypeName PSCustomObject -Property @{


                                'Capacity'          = (ConvertBytes($memblock.Capacity))
                                'Capacity (GB)'     = $memblock.Capacity / 1GB
                                'Computer'          = $memblock.__SERVER
                                'Location'          = $memblock.DeviceLocator
                                'Manufacturer'      = $memblock.Manufacturer
                                'Part Number'       = $memblock.PartNumber
                                'Serial Number'     = $memblock.SerialNumber
                                'Speed'             = [string]($memblock.Speed) + ' MHz'
                                'Type'              = $memblock.Name


                            } # New-Object
                        $obj_memory.PSObject.TypeNames.Insert(0,"Memory")
                        $obj_memory_selection = $obj_memory | Select-Object 'Computer','Location','Capacity','Speed','Manufacturer','Part Number','Type','Serial Number'


    } # foreach


# Write the memory block results in console
Write-Output $empty_line
Write-Output $empty_line
Write-Output $empty_line
Write-Output $obj_memory_selection | Format-Table -auto


# Gather some data for a first summary table
$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer 
$used_memory_perc = $os | Select-Object @{Label='UsedMemoryPerc'; Expression={"{0:N1}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) / $_.TotalVisibleMemorySize) * 100) }}
$used_memory = $os | Select-Object @{Label='UsedMemory'; Expression={(($_.TotalVisibleMemorySize * 1kb) - ($_.FreePhysicalMemory) * 1kb) }}
$available_memory_perc = $os | Select-Object @{Label='AvailableMemoryPerc'; Expression={"{0:N1}" -f ((($os.FreePhysicalMemory) / ($os.TotalVisibleMemorySize)) * 100) }}


$summary_table = $obj_memory | Measure-Object -Property 'Capacity (GB)' -Sum | Select-Object -Property @{Label='Computer';Expression={$computer}},@{Label='Slots in Use';Expression={


                                If ($_.count -ge 2) {
                                    [string]$_.count + ' Slots'
                                } ElseIf ($_.count -eq 1) {
                                    [string]$_.count + ' Slot'
                                } ElseIf ($_.count -eq 0) {
                                    [string]'All memory slots seem to be empty.'
                                } Else {
                                    [string]''
                                } # else


        }},@{Label='Total Physical Memory';Expression={[string]$_.sum + ' GB'}},@{Label='Memory in Use';Expression={(ConvertBytes($($used_memory.UsedMemory)))}},@{Label='Memory in Use (%)';Expression={"$($used_memory_perc.UsedMemoryPerc) %"}},@{Label='Available Memory';Expression={(ConvertBytes($os.FreePhysicalMemory * 1kb))}},@{Label='Available Memory (%)';Expression={"$($available_memory_perc.AvailableMemoryPerc) %"}}


# Write the first summary table in console
Write-Output $empty_line
Write-Output $summary_table | Format-Table -auto


# Gather some data for a second summary table
$gps = Get-Process | Measure-Object -Property ProcessName
$average_load = Get-WmiObject -Class Win32_Processor -ComputerName $computer | Measure-Object -property LoadPercentage -Average
$used_perc = Get-WmiObject -Class Win32_Volume -ComputerName $computer -Filter "DriveLetter = 'C:'" | Select-Object @{Label='C_Drive'; Expression={"{0:N1}" -f  ((($_.Capacity - $_.FreeSpace) / $_.Capacity) * 100) }}


# Write the second summary table in console
Write-Output $empty_line
Write-Output $empty_line
Write-Output "Processes: $($gps.Count)         Average CPU Load: $($average_load.Average) %         Physical Memory in Use: $($used_memory_perc.UsedMemoryPerc) %        C:-Drive Usage: $($used_perc.C_Drive) %"
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
Displays general memory information in console.

.NOTES
Please note that the optional file listed under Options-header will(, if the option is enabled by
copy-pasting the relevant code above the [End of Line] -marker) be created in a directory, which is
specified with the $path variable (at line 6). The $env:temp variable points to the current temp
folder. The default value of the $env:temp variable is C:\Users\<username>\AppData\Local\Temp
(i.e. each user account has their own separate temp folder at path %USERPROFILE%\AppData\Local\Temp).
To change the temp folder for instance to C:\Temp, please, for example, follow the instructions at
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

#>
