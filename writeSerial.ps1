param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $false)]
    [System.String]
    $ToWrite
)

if ($null -eq $port) {
    Write-Host "Port needs to be configured, ex:"
    Write-Host '$port= new-Object System.IO.Ports.SerialPort COM3,9600,None,8,one'
    return
}

$timout = 0
do {
    Start-Sleep -Milliseconds -10
    $timout = $timout + 1
} while ($port.IsOpen -and ($timout -le 100))

try {
    $res = $port.Open()
    if ($res -eq $null) {
        $port.Write($ToWrite)
    }
    else {
        Write-Host $res
    }
}
finally {
    $port.Close()
}

