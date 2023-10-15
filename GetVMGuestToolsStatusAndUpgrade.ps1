Write-Host "環境變數配置中`n"
[string]$currentPath = Get-Location
$wsh = New-Object -ComObject WScript.Shell
# 檔案選擇視窗函數
Function Get-FileName($initialDirectory) {
  # 使用 .NET 框架建立視窗
  [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
  $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
  # 設定初始目錄
  $OpenFileDialog.initialDirectory = $initialDirectory
  # 設定篩選器為 CSV 檔、所有檔案
  $OpenFileDialog.filter = "CSV Files(*.csv)|*.csv|All Files|*.*"
  # 顯示視窗
  $OpenFileDialog.ShowDialog() | Out-Null
  # 傳回選擇的檔案名稱
  $OpenFileDialog.filename
}

Write-Host "設定PowerCLI忽略SSL證書檢查`n"
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
Write-Host "`n"
Write-Host "`n"
$vCenter = Read-Host "請輸入vCenter IP"
Write-Host "`n"
Write-Host "注意!! 請等待視窗彈出後再輸入帳號密碼`n(帳號格式範例 administrator@vsphere.local)。`n連線 vCenter IP為" $vCenter
$conn = Connect-VIServer -server $vCenter
Write-Host "`n"
if ($conn.IsConnected -ne "True") {
  Write-Host "連線失敗請重新執行`n"
  return
}

Write-Host "以下為所有 ESXi 主機名稱"
ForEach ($VMHost in Get-VMHost | select name) {
  Write-Host $VMHost.Name
}
Write-Host "`n"

$VMHostName = Read-Host "請輸入 ESXi 主機名稱"
Write-Host "`n"

Write-Host "正在取得VMware Tools 狀態"
$TotalVMs = (Get-VMHost -name $VMHostName | Get-VM).Count
$CurrentVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsCurrent" } | select name).Count
$NeedUpgradeVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNeedUpgrade" } | select name).Count
$NotInstalledVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNotInstalled" } | select name).Count

$VMsToolsStatus = New-Object PSObject -Property ([ordered]@{
    "VMware Tools is installed, and the version is current. (已安裝最新版總數)"      = $CurrentVMs
    "VMware Tools is installed, but the version is not current. (需升級總數)"  = $NeedUpgradeVMs
    "VMware Tools has never been installed. (尚未安裝VMware Tools)" = $NotInstalledVMs
    "Total VMs (VM 總數量)" = $TotalVMs
  })

$VMsToolsStatus

Write-Host "即將跳出視窗請選擇保存路徑`n"
$csvPath = Get-FileName $currentPath

if ($csvPath -ne "") {
  Write-Host "輸出為CSV到" $csvPath "`n"
  Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | select name, @{Name = "ToolsVersion"; Expression = { $_.config.tools.toolsversion } }, @{ Name = "ToolStatus"; Expression = { $_.Guest.ToolsVersionStatus } } | Sort-Object ToolStatus | Export-Csv -Path $csvPath -NoTypeInformation
}
else {
  Write-Host "路徑錯誤! 不保存CSV"
}

$OutofDateVMs = Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNeedUpgrade" } | select name

if ($NeedUpgradeVMs -eq 0) {
  Write-Host "沒有 VM 需要升級 Vmware Guest Tools"
}
else {
  Write-Host "以下為需要更新 Vmware Guest Tools 的 VM`n"
  Write-Host "VM Name"
  ForEach ($VMs in $OutOfDateVMs) {
    Write-Host $VMs.Name
  }

  $wsh = New-Object -ComObject WScript.Shell
  $answer = $wsh.Popup("開始逐一升級VM上的Vmware Guest Tools", 300, "是否升級Vmware Guest Tools", 4 + 32)
  if ($answer -eq 7) {
    Write-Host "`n您選擇不升級 Vmware Guest Tools。"
  }
  if ($answer -eq 6) {
    Write-Host "`n您選擇升級 Vmware Guest Tools，開始逐一升級。`n"
    ForEach ($VMs in $OutOfDateVMs) {
      Write-Host "Start upgrading Vmware Guest Tools on VM:" $VMs.Name
      Update-Tools $VMs.Name -Verbose -NoReboot
    }
  }
}

Write-Host "`n中斷vCenter的連線中"
Disconnect-VIServer -Server $vCenter -Confirm:$false
Write-Host "`n執行結束"