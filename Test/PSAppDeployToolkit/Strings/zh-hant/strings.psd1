@{
    BalloonTip = @{
        Start = @{
            Install = '安裝開始。'
            Repair = '修復開始。'
            Uninstall = '卸載開始。'
        }
        Complete = @{
            Install = '安裝完成。'
            Repair = '修復完成。'
            Uninstall = '解除安裝完成。'
        }
        RestartRequired = @{
            Install = '安裝完成。需要重新啟動。'
            Repair = '修復完成。需要重新啟動。'
            Uninstall = '卸載完成。需要重新啟動。'
        }
        FastRetry = @{
            Install = '安裝未完成。'
            Repair = '修復未完成。'
            Uninstall = '解除安裝未完成。'
        }
        Error = @{
            Install = '安裝失敗。'
            Repair = '維修失敗。'
            Uninstall = '解除安裝失敗。'
        }
    }
    BlockExecutionText = @{
        Message = @{
            Install = '啟動此應用程式已被暫時阻止，以便完成安裝作業。'
            Repair = '啟動此應用程式已被暫時阻止，以便完成修復操作。'
            Uninstall = '啟動此應用程式已暫時受阻，以便完成卸載作業。'
        }
        Subtitle = @{
            Install = '{Toolkit\CompanyName} - 應用程式安裝'
            Repair = '{Toolkit\CompanyName} - 應用程式維修'
            Uninstall = '{Toolkit\CompanyName} - 應用程式卸載'
        }
    }
    DiskSpaceText = @{
        Message = @{
            Install = "您沒有足夠的磁碟空間來完成安裝：`n{0}`n`n所需空間： {1}MB `n 可用空間： {2}MB`n`n請釋放足夠的磁碟空間，以便繼續安裝。"
            Repair = "您沒有足夠的磁碟空間來完成修復:`n{0}`n`n所需空間： {1}MB`n可用空間： {2}MB`n`n請釋放足夠的磁碟空間，以便繼續進行修復。"
            Uninstall = "您沒有足夠的磁碟空間來完成卸載:`n{0}`n`n所需空間： {1}MB`n可用空間： {2}MB`n`n請釋放足夠的磁碟空間，以便繼續卸載。"
        }
    }
    InstallationPrompt = @{
        Subtitle = @{
            Install = '{Toolkit\CompanyName} - 應用程式安裝'
            Repair = '{Toolkit\CompanyName} - 應用程式維修'
            Uninstall = '{Toolkit\CompanyName} - 應用程式卸載'
        }
    }
    ProgressPrompt = @{
        Message = @{
            Install = '安裝中。 請稍候…'
            Repair = '維修中。 請稍候…'
            Uninstall = '正在卸載。 請稍候…'
        }
        MessageDetail = @{
            Install = '安裝完成後，此視窗將自動關閉。'
            Repair = '維修完成後，此視窗將自動關閉。'
            Uninstall = '卸載完成後，此視窗將自動關閉。'
        }
        Subtitle = @{
            Install = '{Toolkit\CompanyName} - 應用程式安裝'
            Repair = '{Toolkit\CompanyName} - 應用程式維修'
            Uninstall = '{Toolkit\CompanyName} - 應用程式卸載'
        }
    }
    RestartPrompt = @{
        ButtonRestartLater = '最小化'
        ButtonRestartNow = '現在重新啟動'
        Message = @{
            Install = '為了完成安裝，您必須重新啟動電腦。'
            Repair = '為了完成維修，您必須重新啟動電腦。'
            Uninstall = '為了讓卸載完成，您必須重新啟動電腦。'
        }
        CustomMessage = ''
        MessageRestart = '您的電腦會在倒數計時結束時自動重新啟動。'
        MessageTime = '請儲存您的工作，並在指定時間內重新啟動。'
        TimeRemaining = '剩餘時間：'
        Title = '需要重新啟動'
        Subtitle = @{
            Install = '{Toolkit\CompanyName} - 應用程式安裝'
            Repair = '{Toolkit\CompanyName} - 應用程式維修'
            Uninstall = '{Toolkit\CompanyName} - 應用程式卸載'
        }
    }
    CloseAppsPrompt = @{
        Classic = @{
            WelcomeMessage = @{
                Install = '下列應用程式即將安裝：'
                Repair = '下列應用程式即將被修復：'
                Uninstall = '下列應用程式即將被解除安裝：'
            }
            CloseAppsMessage = @{
                Install = "在繼續安裝之前，必須關閉下列程式。`n`n請儲存您的工作，關閉程式，然後繼續。或者，保存您的工作，然後按一下`「關閉程式`」。"
                Repair = "在進行修復之前，必須關閉下列程式。`n`n請儲存您的工作，關閉程式，然後繼續。或者，保存您的工作并单击 「关闭程序」。"
                Uninstall = "卸載程式前，必須先關閉下列程式。或者，保存您的工作並點擊 「關閉程式」。"
            }
            ExpiryMessage = @{
                Install = '您可以選擇延遲安裝，直到延遲到期:'
                Repair = '您可以選擇延遲修復，直到延遲到期：'
                Uninstall = '您可以選擇延遲卸載，直到延遲期限到期：'
            }
            DeferralsRemaining = '剩餘的延遲：'
            DeferralDeadline = '截止日期:'
            ExpiryWarning = '一旦延遲到期，您將無法再選擇延遲。'
            CountdownDefer = @{
                Install = '安裝會自動繼續：'
                Repair = '維修會自動繼續進行：'
                Uninstall = '卸載將自動繼續中：'
            }
            CountdownClose = @{
                Install = '注意：程式會自動關閉：'
                Repair = '注意：程式會自動關閉：'
                Uninstall = '注意：程式會自動關閉：'
            }
            ButtonClose = '關閉程式'
            ButtonDefer = '延遲'
            ButtonContinue = '繼續'
            ButtonContinueTooltip = '僅在關閉上述列出的應用程式後選擇 「繼續」。'
        }
        Fluent = @{
            DialogMessage = @{
                Install = '請先保存您的工作再繼續，因為下列應用程式會自動關閉。'
                Repair = '請先保存您的工作再繼續，因為下列應用程式會自動關閉。'
                Uninstall = '請先保存您的工作再繼續，因為下列應用程式會自動關閉。'
            }
            DialogMessageNoProcesses = @{
                Install = '請選擇「安裝」繼續安裝。'
                Repair = '請選擇「維修」繼續進行維修。'
                Uninstall = '請選擇「卸載」繼續進行卸載。'
            }
            AutomaticStartCountdown = '自動啟動倒數計時'
            DeferralsRemaining = '剩餘延期'
            DeferralDeadline = '延期截止日期'
            ButtonLeftText = @{
                Install = '關閉應用程式並安裝'
                Repair = '關閉應用程式與修復'
                Uninstall = '關閉應用程式並卸載'
            }
            ButtonLeftNoProcessesText = @{
                Install = '安裝'
                Repair = '修復'
                Uninstall = '解除安裝'
            }
            ButtonRightText = '延遲'
            Subtitle = @{
                Install = '{Toolkit\CompanyName} - 應用程式安裝'
                Repair = '{Toolkit\CompanyName} - 應用程式維修'
                Uninstall = '{Toolkit\CompanyName} - 應用程式卸載'
            }
        }
        CustomMessage = ''
    }
}

# SIG # Begin signature block
# MIIuaAYJKoZIhvcNAQcCoIIuWTCCLlUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAWqXgdDK0QNkXp
# aj7BHllXzCsz4HaPr7xrlMEETxgPRqCCE5UwggWQMIIDeKADAgECAhAFmxtXno4h
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
# ghopMIIaJQIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAr5W7a+ogyFDpjG+46sCPkwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgtjKaqlmVPzh3YRKin223Nc8OW7AwDfvfv27mIARWwXQw
# DQYJKoZIhvcNAQEBBQAEggGAMnyM2VG9vX2M0ku/XbOSBOXZDNlAU6/ZwcnGaNyo
# WQexpae6HvN8evPTkq1+UpnCRdVvVD5k9e/SLaDUQYpPB6BBJw+zeiMxtoOG4DCL
# hvTgOK2pK+v3Ymlz2cMZOwu8YfbZmDE0+0xZJImmLade8hvAULZDFRGjjx8CuxQB
# QaOdgS/0O5KYPUsXOMsgdWt3awQlX9I+NhM18c6ms26/0A02KK/Q0UW59zT3WeGk
# bjfPrUpx0bxE8KvUOhvz/BKL50fayK5Gc0nqUxq63dVrGnUwl5j5ww6ysWwDGqhV
# eucYXiHyaXiRBn/AK0gVqOc8pSy9H/mQeLRNWgcRAvVFUAvNXLJxrl1VYpXTLN3s
# a5POaGGmZAjEhcDL7Yvd3dwE5GXN8IkKg4PU4jVFtW03FMrUnddxNUFe7dL3QJYX
# N0k5zuvEhR5GdUvhRS9sJExKECCpMkqpGSdzvj2ubVuiXYhVMRQNb24yQymvWoe2
# 6zc7Y29zIV+pH8Fj6vtTIJFQoYIXdjCCF3IGCisGAQQBgjcDAwExghdiMIIXXgYJ
# KoZIhvcNAQcCoIIXTzCCF0sCAQMxDzANBglghkgBZQMEAgEFADB3BgsqhkiG9w0B
# CRABBKBoBGYwZAIBAQYJYIZIAYb9bAcBMDEwDQYJYIZIAWUDBAIBBQAEICYSdmgg
# wTPLoNO7S9BSDh7uGKWf/KoWgOOTpMkFaF9XAhATf8AbaA7+bFyYN0pUu0rhGA8y
# MDI1MDgwNzA3MTMwM1qgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC0cR2p5V0
# aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1w
# aW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAwMFoXDTM2
# MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAg
# UmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# ANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA69HFTBdw
# bHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6wW2R6kSu9
# RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00Cll8pjrU
# cCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOMA3CoB/iU
# SROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmotuQhcg9tw
# 2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1OpbybpMe4
# 6YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeHVZlc4seA
# O+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1roSrgHjSH
# lq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSUROwnu7zER6
# EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW00aHzrDch
# Ic2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGVMIIBkTAM
# BgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM6DAfBgNV
# HSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYD
# VR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFt
# cGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5o
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3Rh
# bXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwB
# BAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcEua5gQezR
# CESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/YmRDfxT7C0
# k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8AQ/UdKFO
# tj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/EABgfZXLW
# U0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQVTeLni2n
# HkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gVutDojBIF
# eRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85EE8LUkqR
# hoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hggt8l2Yv7
# roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJgKf47Cdx
# VRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLvUxxVZE/r
# ptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7POGT75qaL
# 6vdCvHlshtjdNXOCIUjsarfNZzCCBrQwggScoAMCAQICEA3HrFcF/yGZLkBDIgw6
# SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGln
# aUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4MDExNDIz
# NTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEw
# PwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2
# IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# ALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81FzJsQqr5
# G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUiljNOqnIVD
# /gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9+0F2w4QJ
# LVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8NF5vp7eaZ2CV
# NxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUUFREmDrMxSNlr
# /NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8SdLnEut/
# GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCyQEEGcbLe
# 1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespYMQmUiote8lad
# jS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrPV6/7umw052Ak
# yiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+zoFpp4Ra+MlK
# M2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFdMIIBWTAS
# BgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMC
# AYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0
# MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCG
# SAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZxML2+C8i1NKZ/
# zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97frPBtIj+ZLzdp+
# yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+NMvEuBd/2vmd
# YxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYAgwPyWLKu6Rna
# ID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSjjf4J2a7j
# LzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+BFo+z7bKSBwZ
# XTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06VXxyKkOirv6o
# 02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284NHNboDGcmWXf
# wXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDezooIs8CVnrpHM
# iD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7WP/UwgOkw
# /HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpSM9LHJmyrxaFt
# oza2zNaQ9k+5t1wwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOIIQBhaMA0GCSqG
# SIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFz
# c3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTExMDkyMzU5NTla
# MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9v
# dCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo3rvkXUo8
# MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutWxpdtHauy
# efLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQRBAu34Lz
# B4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nPzaw0QF+x
# embud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6PgNq2kZhA
# kHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEBfCFM1Lyu
# GwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3wWmIdph2
# PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2mkQZK37A
# lLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2Yn72gLD7
# 6GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF13nqsX40/
# ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+dIPyhHsXA
# j6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNVHRMBAf8EBTAD
# AQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYIKwYBBQUHAQEE
# bTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYB
# BQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDARBgNVHSAECjAI
# MAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6CjdBs9thbX979X
# B72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/gLn4offyct4k
# vFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcudT6cGAxN3J0TU
# 53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3osdz9HNj0d1pc
# VIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1VWvPJ6tsds5v
# Iy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eqXijiuZQxggN8
# MIIDeAIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5j
# LjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNB
# NDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUD
# BAIBBQCggdEwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqGSIb3DQEJ
# BTEPFw0yNTA4MDcwNzEzMDNaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYEFN1iMKyG
# Ci0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCANF123aZRr8ymipDhUXJtT
# cGOQmZiHysqjw0jrqbkI5TA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBKoD+iLNdc
# hMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASCAgB5RpjN
# 4e/tN7BVVcnVWwxq0xQWJFR2JuPreiX7BxqPAOpjW2Au9rkDVeLBMTvnNzpctEAp
# /e7k74BsB6GUhPXIwpQnQLdpcxgAoP4MMqx+OdUil2SWtjjKNE85SSNAUKiPpH+Z
# pc4Y3MZyDTcXj3VJvp5QvkCU/joGuUusdbuJS6NUF+clFTnsKrVeykIWSaIQkQVq
# 1MqaVXcGa9zinLREEAstAJC9PSR/lkp9ng+bB7vgSsoT/zPNgRaZ1GvOCM9dtpUe
# CkefiNuNlEU6Yf0ZMIfoT4rN7lTZUngELFimqkrc+wh5x20LkjNkbudpn5qCp5dl
# 5uLuKGEP6T+OiJvjlmlz4c82E+muczIZkKamV7QCMRGkt6Xvzei0kr2yjp4nhgs7
# vNW3px3z1ipL1BYnbyonloIaT4g8B/E6iYmxle2yrZtzaB+Oly/Mk0PMeuN0vhvT
# QaIrO2YqUGNZ7uGeJwXG9ktQHmd89Xz3t2KC9tvZHN0S7nB60mntY94gF5tZfaqD
# /AVShMvLqf4SM+VwnTgnDsRes/0gH1ObswiRfeZN4k9hIBWG81WsYgLgga/repzM
# Ys2h1iDPhcyUAuiJuSBeVGw/DyZlgLI2+2R8fbnsMzAGySj8eWuln+6S162nE6Ir
# 7fpKWbal7o63anWDsmBADBtNLPKYrfIe78DI1Q==
# SIG # End signature block
