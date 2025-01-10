Write-Host "環境變數配置中`n"
[string]$currentPath = Get-Location
$wsh = New-Object -ComObject WScript.Shell

# 檔案選擇視窗函數
Function Get-FileName($initialDirectory) {
  [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
  $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.filter = "CSV Files(*.csv)|*.csv|All Files|*.*"
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.filename
}

Write-Host "設定PowerCLI忽略SSL證書檢查`n"
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
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

Write-Host "正在取得所有 VM 資訊..."

# 抓取所有 VM 的資料
$allVMs = Get-VM | Select-Object Name, NumCpu, MemoryGB, @{Name="Storage (GB)";Expression={(Get-HardDisk -VM $_ | Measure-Object -Property CapacityGB -Sum).Sum}}, @{Name="UsedStorage (GB)";Expression={(Get-HardDisk -VM $_ | Measure-Object -Property CapacityGB -Sum).Sum - (Get-HardDisk -VM $_ | Measure-Object -Property FreeSpaceGB -Sum).Sum}}, PowerState

# 統計資料
$totalVMs = $allVMs.Count
$poweredOnVMs = ($allVMs | Where-Object { $_.PowerState -eq "PoweredOn" }).Count
$poweredOffVMs = ($allVMs | Where-Object { $_.PowerState -eq "PoweredOff" }).Count

Write-Host "VM 總數量: $totalVMs"
Write-Host "已開機 VM 數量: $poweredOnVMs"
Write-Host "已關機 VM 數量: $poweredOffVMs"

Write-Host "即將跳出視窗請選擇保存路徑`n"
$csvPath = Get-FileName $currentPath

if ($csvPath -ne "") {
  Write-Host "輸出為CSV到" $csvPath "`n"
  $allVMs | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}
else {
  Write-Host "路徑錯誤! 不保存CSV"
}

Write-Host "中斷vCenter的連線中"
Disconnect-VIServer -Server $vCenter -Confirm:$false
Write-Host "執行結束"
