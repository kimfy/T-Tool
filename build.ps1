# Build command
# "C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" /in "t_tool.au3" /out ".\T-Tool.exe" /icon ".\sys\gui\icon.ico" /gui /x86

Function Pre {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [Alias("Path")]
        [string[]] $Paths = @("C:\Program Files (x86)\AutoIt3\Aut2Exe\", "C:\Program Files\AutoIt3\Aut2Exe\"),
        [Parameter(Mandatory = $true)]
        [string] $Version
    )

    begin {
        # Set new version
        try {
            ((Get-Content .\t_tool.au3 -Encoding UTF8 | Out-String) -replace '(Dim \$SWVersion = ")(.+)(")', "`${1}$Version`$3" | Out-File .\t_tool.au3 -Encoding utf8)
        }
        catch {
            Write-Host "Something went wrong when setting new version in source file"
            Write-Error $Error[0]
        }
    }

    process {
        if (IsConverterInPath) {
            Write-Verbose "Aut2Exe is in PATH"
        }
        else {
            $Added = $false
            foreach ($Path in $Paths) {
                if (ValidatePath -Path $Path) {
                    Write-Verbose "Added $Path to PATH variable for this session"
                    $env:Path += ";$Path"
                    $Added = $true
                    break
                }
            }
            if (-not $Added) {
                Write-Error "Unable to proceed; Aut2exe missing from PATH" -ErrorAction Stop
            }
        }
    }
}

Function Build {
    Write-Host "Building application..."
}

Function IsConverterInPath {
    try {
        Get-Command "Aut2exe.exe" -ErrorAction Stop
        return $true
    }
    catch {
        Write-Verbose "Aut2Exe not found in PATH variable"
        return $false
    }
}

Function ValidatePath {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true)
        ]
        [string] $Path
    )
    return Test-Path -Path ($Path + "\Aut2Exe.exe")
}

Pre -Verbose
Build -VerboseFunction ValidatePath