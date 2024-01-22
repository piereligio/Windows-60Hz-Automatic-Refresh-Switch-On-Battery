function Set-PowerRefreshRate {
        [CmdletBinding()]
    param (
    )

    # Check if QRes.exe is accessible from the PATH
    $QResPath = Get-Command "QRes.exe" -ErrorAction SilentlyContinue
    if ($null -eq $QResPath) {
        Write-Host "This script requires QRes.exe to be in the PATH or the same folder."
        Write-Host "You can download QRes.exe from https://www.majorgeeks.com/files/details/qres.html"
        return -1
    }

    # Check the number of connected and active displays
    $displayCount = (Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | select Active).Count
    if ($displayCount -gt 1) {
        Write-Host "More than one active display detected. QRes will not be called."
        return
    }


	# Select the active monitor
	$activeMonitor = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | Where-Object { $_.Active -eq $true }
	
	# Get the instance name of the active monitor
	$instanceName = $activeMonitor.InstanceName
	
	# Get the supported source modes for the active monitor
	$supportedModes = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorListedSupportedSourceModes | Where-Object { $_.InstanceName -eq $instanceName }
	
	# Variable to hold the highest refresh rate
	$maxRefreshRate = 0
	
	# Iterate through the supported modes to find the highest refresh rate
	foreach ($mode in $supportedModes.MonitorSourceModes) {
		# Calculate refresh rate
		$refreshRate = [math]::Round($mode.VerticalRefreshRateNumerator / $mode.VerticalRefreshRateDenominator)
		
		# Compare with maxRefreshRate and update if higher
		if ($refreshRate -gt $maxRefreshRate) {
			$maxRefreshRate = $refreshRate
		}
	}
	
	# Output the highest refresh rate
	Write-Host "Highest Refresh Rate for Current Resolution: $maxRefreshRate Hz"


    # Get battery status
    $batteryStatus = (Get-WmiObject -Class Win32_Battery).BatteryStatus

    # Set refresh rate based on power status
    if ($batteryStatus -eq 1) {
        # On battery
        "$([DateTime]::Now) Switched Monitor to 60 Hz"
        Write-Host "Switched Monitor to 60 Hz"
        Start-Process $QResPath.Source -ArgumentList "/r:60"
    } else {
        # Plugged in
        $refreshRateToSet = $maxRefreshRate
        "$([DateTime]::Now) Switched Monitor to $refreshRateToSet Hz"
        Write-Host "Switched Monitor to $refreshRateToSet Hz"
        Start-Process $QResPath.Source -ArgumentList "/r:$refreshRateToSet"
    }
}