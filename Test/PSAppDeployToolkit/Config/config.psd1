@{
    Assets = @{
        # Specify filename of the logo.
        Logo = '..\Assets\AppIcon.png'

        # Specify filename of the banner (Classic-only).
        Banner = '..\Assets\Banner.Classic.png'
    }

    MSI = @{
        # Installation parameters used for non-silent MSI actions.
        InstallParams = 'REBOOT=ReallySuppress /QB-!'

        # Logging level used for MSI logging.
        LoggingOptions = '/L*V'

        # Log path used for MSI logging.
        LogPath = '$envWinDir\Logs\Software'

        # Log path used for MSI logging when RequireAdmin is False.
        LogPathNoAdminRights = '$envProgramData\Logs\Software'

        # The length of time in seconds to wait for the MSI installer service to become available. Default is 600 seconds (10 minutes).
        MutexWaitTime = 600

        # Installation parameters used for silent MSI actions.
        SilentParams = 'REBOOT=ReallySuppress /QN'

        # Installation parameters used for MSI uninstall actions.
        UninstallParams = 'REBOOT=ReallySuppress /QN'
    }

    Toolkit = @{
        # Specify the path for the cache folder.
        CachePath = '$envProgramData\SoftwareCache'

        # Specify if the log files should be bundled together in a compressed zip file.
        CompressLogs = $false

        # Choose from either 'Native' for native PowerShell file copy via Copy-ADTItem, or 'Robocopy' to use robocopy.exe.
        FileCopyMode = 'Native'

        # Specify if an existing log file should be appended to.
        LogAppend = $true

        # Specify if debug messages such as bound parameters passed to a function should be logged.
        LogDebugMessage = $false

        # Specify maximum number of previous log files to retain.
        LogMaxHistory = 10

        # Specify maximum file size limit for log file in megabytes (MB).
        LogMaxSize = 10

        # Log path used for Toolkit logging.
        LogPath = '$envWinDir\Logs\Software'

        # Same as LogPath but used when RequireAdmin is False.
        LogPathNoAdminRights = '$envProgramData\Logs\Software'

        # Specifies that a subfolder based on InstallName should be used for all log capturing.
        LogToSubfolder = $false

        # Specify if log file should be a CMTrace compatible log file or a Legacy text log file.
        LogStyle = 'CMTrace'

        # Specify if log messages should be written to the console.
        LogWriteToHost = $true

        # Specify if console log messages should bypass PowerShell's subsystems and be sent direct to stdout/stderr.
        # This only applies if "LogWriteToHost" is true, and the script is being ran in a ConsoleHost (not the ISE, or another host).
        LogHostOutputToStdStreams = $false

        # Automatically changes DeployMode to Silent during the OOBE.
        OobeDetection = $true

        # Registry key used to store toolkit information (with PSAppDeployToolkit as child registry key), e.g. deferral history.
        RegPath = 'HKLM:\SOFTWARE'

        # Same as RegPath but used when RequireAdmin is False. Bear in mind that since this Registry Key should be writable without admin permission, regular users can modify it also.
        RegPathNoAdminRights = 'HKCU:\SOFTWARE'

        # Specify if Administrator Rights are required. Note: Some functions won't work if this is set to false, such as deferral, block execution, file & registry RW access and potentially logging.
        RequireAdmin = $true

        # Automatically changes DeployMode for session zero (SYSTEM) operations.
        SessionDetection = $true

        # Path used to store temporary Toolkit files (with PSAppDeployToolkit as subdirectory), e.g. cache toolkit for cleaning up blocked apps. Normally you don't want this set to a path that is writable by regular users, this might lead to a security vulnerability. The default Temp variable for the LocalSystem account is C:\Windows\Temp.
        TempPath = '$envTemp'

        # Same as TempPath but used when RequireAdmin is False.
        TempPathNoAdminRights = '$envTemp'
    }

    UI = @{
        # Used to turn automatic balloon notifications on or off.
        BalloonNotifications = $true

        # The name to show by default for all balloon notifications.
        BalloonTitle = 'PSAppDeployToolkit'

        # Choose from either 'Fluent' for contemporary dialogs, or 'Classic' for PSAppDeployToolkit 3.x WinForms dialogs.
        DialogStyle = 'Fluent'

        # Exit code used when a UI prompt times out.
        DefaultExitCode = 1618

        # Time in seconds after which the prompt should be repositioned centre screen when the -PersistPrompt parameter is used. Default is 60 seconds.
        DefaultPromptPersistInterval = 60

        # Time in seconds to automatically timeout installation dialogs. Default is 55 minutes so that dialogs timeout before Intune times out.
        DefaultTimeout = 3300

        # Exit code used when a user opts to defer.
        DeferExitCode = 60012

        # Specify whether to re-enumerate running processes dynamically while displaying Show-ADTInstallationWelcome.
        # If the CloseProcesses items were not running when the prompt was displayed, and are subsequently detected to be running, the prompt will be updated with the apps to close.
        # If the CloseProcesses items were running when the prompt was displayed and are subsequently detected not to be running then the installation will automatically continue if deferral is not available.
        # If the running applications change (new CloseProcesses launched or running processes closed), the list box will dynamically update to reflect the currently running applications.
        DynamicProcessEvaluation = $true

        # Time in seconds after which to re-enumerate running processes while displaying the Show-ADTInstallationWelcome prompt. Default is 2 seconds.
        DynamicProcessEvaluationInterval = 2

        <# Specify a static UI language using the one of the Language Codes listed below to override the language culture detected on the system.
            Language Code    Language       |       Language Code    Language
            =============    ========       |       =============    ========
            AR               Arabic         |       KO               Korean
            CZ               Czech          |       NL               Dutch
            DA               Danish         |       NB               Norwegian (Bokmål)
            DE               German         |       PL               Polish
            EN               English        |       PT               Portuguese (Portugal)
            EL               Greek          |       PT-BR            Portuguese (Brazil)
            ES               Spanish        |       RU               Russian
            FI               Finnish        |       SK               Slovak
            FR               French         |       SV               Swedish
            HE               Hebrew         |       TR               Turkish
            HU               Hungarian      |       ZH-Hans          Chinese (Simplified)
            IT               Italian        |       ZH-Hant          Chinese (Traditional)
            JA               Japanese       |
        #>
        LanguageOverride = $null

        # Time in seconds after which to re-prompt the user to close applications in case they ignore the prompt or they cancel the application's save prompt.
        PromptToSaveTimeout = 120

        # Time in seconds after which the restart prompt should be re-displayed/repositioned when the -NoCountdown parameter is specified. Default is 600 seconds.
        RestartPromptPersistInterval = 600
    }
}

# SIG # Begin signature block
# MIIuKwYJKoZIhvcNAQcCoIIuHDCCLhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDwDj8+4CuYCZXD
# Te8j83I6+zljIMQv4xnE1sDvMTx+g6CCE5UwggWQMIIDeKADAgECAhAFmxtXno4h
# MuI5B72nd3VcMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0xMzA4MDExMjAwMDBaFw0z
# ODAxMTUxMjAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjQjBAMA8GA1Ud
# EwEB/wQFMAMBAf8wDgYDVR0PAQH/BAQDAgGGMB0GA1UdDgQWBBTs1+OC0nFdZEzf
# Lmc/57qYrhwPTzANBgkqhkiG9w0BAQwFAAOCAgEAu2HZfalsvhfEkRvDoaIAjeNk
# aA9Wz3eucPn9mkqZucl4XAwMX+TmFClWCzZJXURj4K2clhhmGyMNPXnpbWvWVPjS
# PMFDQK4dUPVS/JA7u5iZaWvHwaeoaKQn3J35J64whbn2Z006Po9ZOSJTROvIXQPK
# 7VB6fWIhCoDIc2bRoAVgX+iltKevqPdtNZx8WorWojiZ83iL9E3SIAveBO6Mm0eB
# cg3AFDLvMFkuruBx8lbkapdvklBtlo1oepqyNhR6BvIkuQkRUNcIsbiJeoQjYUIp
# 5aPNoiBB19GcZNnqJqGLFNdMGbJQQXE9P01wI4YMStyB0swylIQNCAmXHE/A7msg
# dDDS4Dk0EIUhFQEI6FUy3nFJ2SgXUE3mvk3RdazQyvtBuEOlqtPDBURPLDab4vri
# RbgjU2wGb2dVf0a1TD9uKFp5JtKkqGKX0h7i7UqLvBv9R0oN32dmfrJbQdA75PQ7
# 9ARj6e/CVABRoIoqyc54zNXqhwQYs86vSYiv85KZtrPmYQ/ShQDnUBrkG5WdGaG5
# nLGbsQAe79APT0JsyQq87kP6OnGlyE0mpTX9iV28hWIdMtKgK1TtmlfB2/oQzxm3
# i0objwG2J5VT6LaJbVu8aNQj6ItRolb58KaAoNYes7wPD1N1KarqE3fk3oyBIa0H
# EEcRrYc9B9F1vM/zZn4wggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0G
# CSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTla
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C
# 0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce
# 2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0da
# E6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6T
# SXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoA
# FdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7Oh
# D26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM
# 1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z
# 8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05
# huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNY
# mtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP
# /2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0T
# AQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYD
# VR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNV
# HR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATAN
# BgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95Ry
# sQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HL
# IvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5Btf
# Q/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnh
# OE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIh
# dXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV
# 9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/j
# wVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYH
# Ki8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmC
# XBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l
# /aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZW
# eE4wggdJMIIFMaADAgECAhAK+Vu2vqIMhQ6YxvuOrAj5MA0GCSqGSIb3DQEBCwUA
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwHhcNMjQwOTA1MDAwMDAwWhcNMjcwOTA3MjM1OTU5WjCB0TET
# MBEGCysGAQQBgjc8AgEDEwJVUzEZMBcGCysGAQQBgjc8AgECEwhDb2xvcmFkbzEd
# MBsGA1UEDwwUUHJpdmF0ZSBPcmdhbml6YXRpb24xFDASBgNVBAUTCzIwMTMxNjM4
# MzI3MQswCQYDVQQGEwJVUzERMA8GA1UECBMIQ29sb3JhZG8xFDASBgNVBAcTC0Nh
# c3RsZSBSb2NrMRkwFwYDVQQKExBQYXRjaCBNeSBQQywgTExDMRkwFwYDVQQDExBQ
# YXRjaCBNeSBQQywgTExDMIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEA
# uydxko2Hrl6sANJUjfdypKP60qBH5EkhfaRQAnn+e3vg2eVcbiEWIjlrMYzvK2sg
# OMBbwGebqAURkFmUCKDdGxcxKeuXdaXPHWPKwc2WjYCFajrX6HofiiwNzOCdL6VE
# 4PDQhPRR7SIdNNFSrx5C4ZDN1T6OH+ydX7EQF8+NBUNHRbEVdl+h9H5Aexx63afa
# 8zu3g/GXluyXKbb+JHtgNJaUgFuFORTxw1TO6qH+S6Hrppf9QcAFmu4xGtkc2FSh
# gv0NgWMNGDZqJr/o9sqJ2tdaZHDyr6H8PvY8egoUshF7ccgEYtEEdB9SRR8mVQik
# 1w5oGTjDWjHj+8jgTpzletRywptk/m8PehVBN8ntqoSdvLLcuQVzmuPLzN/iuKh5
# sZeWvqPONApcEnZcONpXebyiUPnEePr5rZAU7hMjMw2ZPnQlMcbGvtgP2qi7m2f3
# mXFYxWjlKCxaApYHeqSFeWC8zM7OYL2HlZ+GuK4XG8jKVE6sWSW9Wk/dm0vJbasv
# AgMBAAGjggICMIIB/jAfBgNVHSMEGDAWgBRoN+Drtjv4XxGG+/5hewiIZfROQjAd
# BgNVHQ4EFgQU5GCU3SEqeIbhhY9eyU0LcTI75X8wPQYDVR0gBDYwNDAyBgVngQwB
# AzApMCcGCCsGAQUFBwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwDgYD
# VR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMIG1BgNVHR8Ega0wgaow
# U6BRoE+GTWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMFOgUaBPhk1odHRw
# Oi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmlu
# Z1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDCBlAYIKwYBBQUHAQEEgYcwgYQwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBcBggrBgEFBQcwAoZQ
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29k
# ZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcnQwCQYDVR0TBAIwADANBgkq
# hkiG9w0BAQsFAAOCAgEAqA0ub/ilMgdIvMiBeWBoiMxe5OIblObGI7lemcP2WEqa
# EASW11/wVwJU63ZwhtkQaNU4rXjf6fqy5pOUzpQXgYjSaO4D/AOMJKHlypxslFqZ
# /dYpcue2xE3H7lmO4KPf8VxXuFIUqjLetU+kkh7o/Q52RabVAuOrPFKnObixy1HI
# x0/5F+RuP9xhqmDbfM7l5zUAcuOCCkY7buuInEsip9BZXUiVb8K5bPR9Rk7Doat4
# FQmN72xjakcEZOMU/vg0ZgVa8nxkBXtVsjxbsr+bODn0cddHK1QHWil/PmpANkxN
# 7H8tdCAZ8bTzIvvudxSLnt7ssbbQDkAyNw0btDH+MKv/l+VcYyQH51Z5xT9DvHCm
# Ed774boZkP2GfTFvn7/gISEjTdOuUGstdrgSwg1zJPqgK7zWxK48xC7awpa3gwOs
# 9pnyiqHG3rx84/SHUiAL2lkljsD3epmRxsWeZhZNY93xEpQHe9LBvo/t4VRjZzqU
# z+pfEMPqeX/g5+mpb4ap6ZmNJuAYJFmU0LIkCLQN9mKXi1Il9WU6ifn3vYutGMSL
# /BdeWP+7fM7MZLiO+1BIsBdSmV6pZVS3LRBAy3wIlbWL69mvyLCPIQ7z4dtfuzwC
# 36E9k2vhzeiDQ+k1dFJDSdxTDetsck0FuD1ovhiu2caL4BdFsCWsXPLMyvu6OlYx
# ghnsMIIZ6AIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAr5W7a+ogyFDpjG+46sCPkwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgQ1a7AJE6ncffXOVZpTXpD3O5oZxa0Nn/XEal9rVxKBAw
# DQYJKoZIhvcNAQEBBQAEggGAKCjfE8EVh9BHUJqh8GTaL1s2kBC9dZKyNyMwUdM0
# qwXjhSeqmIXzaW4Atz2KtCoyYG5g3xMPQkKVZ0hBF7iY2QKnOZVNEEavKJAr/RuX
# kIFmxOOD4Zzvl5at8TKhlD3VJE1BngYO0qJTezT/vD1Htjvi4fm0ZAf5n4NDCZsF
# OZJZNKCWCHIPlOEP2SUiW2mE/ttrxV1EArujNOaptJKqbyWXfcSPvObBXPsm6x8f
# VTvb9/jz69q9IzM7O8QeKFnsajfws9DcUec5FrqOhCRZXOiI+OiguOtgpY348QfV
# zNvKT26n7gkPZdztKRev1j3g4o2OYq4xQzdCam7r9+XITRXFgETRaU9Mx0lZ/fVz
# OMRUZF7nAPuVOxdnDTo9C6XoFqycikdkz1Vyu60OUmiBNTUT9URL/DYbWq0ITYFH
# YVp1Y7c3e3JZhRMwv24gwcIgRf56HZACcxHAS/onYckTb6OkFYLMJChBqD9PS39X
# eTWlurdMeltw3AFuN2CX57s6oYIXOTCCFzUGCisGAQQBgjcDAwExghclMIIXIQYJ
# KoZIhvcNAQcCoIIXEjCCFw4CAQMxDzANBglghkgBZQMEAgEFADB3BgsqhkiG9w0B
# CRABBKBoBGYwZAIBAQYJYIZIAYb9bAcBMDEwDQYJYIZIAWUDBAIBBQAEIPoTA8/d
# T0daKmUs8BX7u5dKR9tchSaxS08fZbs2ny13AhAHW5jjYoWoJe09cgGTgJCLGA8y
# MDI0MTIxOTIyNDQxNFqgghMDMIIGvDCCBKSgAwIBAgIQC65mvFq6f5WHxvnpBOMz
# BDANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBT
# SEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTI0MDkyNjAwMDAwMFoXDTM1MTEyNTIz
# NTk1OVowQjELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSAwHgYDVQQD
# ExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyNDCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAL5qc5/2lSGrljC6W23mWaO16P2RHxjEiDtqmeOlwf0KMCBDEr4I
# xHRGd7+L660x5XltSVhhK64zi9CeC9B6lUdXM0s71EOcRe8+CEJp+3R2O8oo76EO
# 7o5tLuslxdr9Qq82aKcpA9O//X6QE+AcaU/byaCagLD/GLoUb35SfWHh43rOH3bp
# LEx7pZ7avVnpUVmPvkxT8c2a2yC0WMp8hMu60tZR0ChaV76Nhnj37DEYTX9ReNZ8
# hIOYe4jl7/r419CvEYVIrH6sN00yx49boUuumF9i2T8UuKGn9966fR5X6kgXj3o5
# WHhHVO+NBikDO0mlUh902wS/Eeh8F/UFaRp1z5SnROHwSJ+QQRZ1fisD8UTVDSup
# WJNstVkiqLq+ISTdEjJKGjVfIcsgA4l9cbk8Smlzddh4EfvFrpVNnes4c16Jidj5
# XiPVdsn5n10jxmGpxoMc6iPkoaDhi6JjHd5ibfdp5uzIXp4P0wXkgNs+CO/CacBq
# U0R4k+8h6gYldp4FCMgrXdKWfM4N0u25OEAuEa3JyidxW48jwBqIJqImd93NRxvd
# 1aepSeNeREXAu2xUDEW8aqzFQDYmr9ZONuc2MhTMizchNULpUEoA6Vva7b1XCB+1
# rxvbKmLqfY/M/SdV6mwWTyeVy5Z/JkvMFpnQy5wR14GJcv6dQ4aEKOX5AgMBAAGj
# ggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEw
# HwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFJ9XLAN3
# DigVkGalY17uT5IfdqBbMFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3Rh
# bXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1l
# U3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggIBAD2tHh92mVvjOIQSR9lD
# kfYR25tOCB3RKE/P09x7gUsmXqt40ouRl3lj+8QioVYq3igpwrPvBmZdrlWBb0Hv
# qT00nFSXgmUrDKNSQqGTdpjHsPy+LaalTW0qVjvUBhcHzBMutB6HzeledbDCzFzU
# y34VarPnvIWrqVogK0qM8gJhh/+qDEAIdO/KkYesLyTVOoJ4eTq7gj9UFAL1UruJ
# KlTnCVaM2UeUUW/8z3fvjxhN6hdT98Vr2FYlCS7Mbb4Hv5swO+aAXxWUm3WpByXt
# gVQxiBlTVYzqfLDbe9PpBKDBfk+rabTFDZXoUke7zPgtd7/fvWTlCs30VAGEsshJ
# mLbJ6ZbQ/xll/HjO9JbNVekBv2Tgem+mLptR7yIrpaidRJXrI+UzB6vAlk/8a1u7
# cIqV0yef4uaZFORNekUgQHTqddmsPCEIYQP7xGxZBIhdmm4bhYsVA6G2WgNFYagL
# DBzpmk9104WQzYuVNsxyoVLObhx3RugaEGru+SojW4dHPoWrUhftNpFC5H7QEY7M
# hKRyrBe7ucykW7eaCuWBsBb4HOKRFVDcrZgdwaSIqMDiCLg4D+TPVgKx2EgEdeoH
# NHT9l3ZDBD+XgbF+23/zBjeCtxz+dL/9NWR6P2eZRi7zcEO1xwcdcqJsyz/JceEN
# c2Sg8h3KeFUCS7tpFk7CrDqkMIIGrjCCBJagAwIBAgIQBzY3tyRUfNhHrP0oZipe
# WzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdp
# Q2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcwMzIyMjM1
# OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5
# BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0
# YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAxoY1Bkmz
# wT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo+n3znIkL
# f50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY3+vaPcQXf6sZKz5C
# 3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8ald68Dd5
# n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71ORJn7w6lY2zkpsUd
# zTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW8LAxE6lXKZYnLvWH
# po9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU2YIqx5K/
# oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2RrOdOqPV
# A+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR8zZJTYsg
# 0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI/rOP015LdhJRk8mM
# DDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNYCQEoAA6E
# VO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0wggFZMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WMaiCPnshvMB8GA1Ud
# IwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0f
# BDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1
# c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBguEE0TzzBT
# zr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFnzbYSlm/E
# UExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/tYLaqT5Fm
# niye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/KFSJ8NWKcXZl2szw
# cqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7pp1yr8TH
# wcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkPlM05et3/
# JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRxktcma+Q4c6umAU+9
# Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kaddSweJywm
# 228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrdVcA6KYawmKAr7ZVB
# tzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL5HYCJtnw
# ZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QVSucTDh3bNzgaoSv2
# 7dZ8/DCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEM
# BQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJ
# RCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zC
# pyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf
# 1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x
# 4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEio
# ZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7ax
# xLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZ
# OjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJ
# l2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz
# 2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH
# 4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb
# 5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ
# 9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYD
# VR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuC
# MS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0g
# ADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs
# 7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq
# 3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/
# Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9
# /HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWoj
# ayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDGCA3YwggNyAgEB
# MHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYD
# VQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFt
# cGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEFAKCB0TAaBgkq
# hkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI0MTIxOTIy
# NDQxNFowKwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU29OF7mLb0j575PZxSFCHJNWG
# W0UwLwYJKoZIhvcNAQkEMSIEIKOLcbtEd1DaLX1esY6NWiEp0D8CkyqOpU4M5POO
# lW0ZMDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEIHZ2n6jyYy8fQws6IzCu1lZ1/tdz
# 2wXWZbkFk5hDj5rbMA0GCSqGSIb3DQEBAQUABIICAH1EJX+OKgcWaJDUG/uH/xGu
# Iw2jNiMjhmQgzu5fI+ZAxLgegK+/9QJdcDrFcKXAZNHqTll5KEGesbs1OofLar/p
# wnUvFl5+NNjsJ3rYPLkL5m/16iHQGdxu6bend/L4a7U0iPO5/yowp7hHz2AKxIZH
# yYZZydaBdrFcSQUiyWWQYjQKpSHe8ioUwo3qVAl+H8TrFQ54bn/UiuNW82D1YDva
# tQuqAo+I6Mwya4Zl4W6CMEhRlB2J7VknatY6pFwIguwGtd6CYF1kTHaUMfwVGSsN
# ngh/1/326eNc4gC9amGj8G0MhA45q75tQP6qOqZXbpcGFFlUuaVv5y1mAb/O+dIP
# 0ExUusN2c55y6zDQ67w1EVA5f5WPkUpOujbzRVodl47HeHA7eHj+GyuguCcNRH8g
# 3CfEx5BUmEmlDZEMVLC/phSV3o8SyS0G6E49lWQUEN8bhHB4plYySRO5j54C3GUU
# MY+MM6DjYe1w4gJAV+Sp6eKV5myKDWytrY9lfWlvcbT971NqwQuHy/Eh35Jom23H
# 4BEmpSCA0JoafSdXUDG0MWyCjpd76YfMLbWWhD1TCuEYSi03gi+zcNe8dAP4PFJA
# wsGifcqsdVvhdAahY0Nlo9sdJl3LkWxiOHEvmJ/Z3+xFON6hz7vKboO0B38GZqhg
# e3gPoPgRZnZr7OABLR8B
# SIG # End signature block
