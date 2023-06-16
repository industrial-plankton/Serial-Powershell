
if ($null -eq $port) {
    Write-Host "Port needs to be configured, ex:"
    Write-Host '$port= new-Object System.IO.Ports.SerialPort COM3,9600,None,8,one'
    return
}

do {
    try {
        $port.Open()
        $line = $port.ReadExisting()
        if ($line) {
            Write-Host -NoNewline $line
        }
    }
    catch {
        <#Do this if a terminating exception happens#>
    }
    finally {
        $port.Close()
    }
} while ($true)