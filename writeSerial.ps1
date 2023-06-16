param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $false)]
    [System.String]
    $ToWrite
)


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
}
finally {
    $port.Close()
}

else {
    Write-Host $res
}
