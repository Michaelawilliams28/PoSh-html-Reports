
function Get-Config {
    $script:loggingDate = (get-date -Format MM-dd-yyyy-hh:mm:ss)

    $script:logDate = (Get-Date -Format MM-dd-yyyy)

    $script:currentdir = (Get-Location)

    $script:steps = 25

    $script:CompanyLogo = ""

    $script:RightLogo = ""

    $script:ReportTitle = "Active Directory Report"

    $script:ReportSavePath = "$script:currentdir\report\"

    $script:Days = 30

    $script:UserCreatedDays = 7

    $script:DaysUntilPWExpireINT = 7

    $script:ADModNumber = 3
}