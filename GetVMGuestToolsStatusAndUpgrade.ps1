Write-Host "�����ܼưt�m��`n"
[string]$currentPath = Get-Location
$wsh = New-Object -ComObject WScript.Shell
# �ɮ׿�ܵ������
Function Get-FileName($initialDirectory) {
  # �ϥ� .NET �ج[�إߵ���
  [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
  $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
  # �]�w��l�ؿ�
  $OpenFileDialog.initialDirectory = $initialDirectory
  # �]�w�z�ﾹ�� CSV �ɡB�Ҧ��ɮ�
  $OpenFileDialog.filter = "CSV Files(*.csv)|*.csv|All Files|*.*"
  # ��ܵ���
  $OpenFileDialog.ShowDialog() | Out-Null
  # �Ǧ^��ܪ��ɮצW��
  $OpenFileDialog.filename
}

Write-Host "�]�wPowerCLI����SSL�Ү��ˬd`n"
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
Write-Host "`n"
Write-Host "`n"
$vCenter = Read-Host "�п�JvCenter IP"
Write-Host "`n"
Write-Host "�`�N!! �е��ݵ����u�X��A��J�b���K�X`n(�b���榡�d�� administrator@vsphere.local)�C`n�s�u vCenter IP��" $vCenter
$conn = Connect-VIServer -server $vCenter
Write-Host "`n"
if ($conn.IsConnected -ne "True") {
  Write-Host "�s�u���ѽЭ��s����`n"
  return
}

Write-Host "�H�U���Ҧ� ESXi �D���W��"
ForEach ($VMHost in Get-VMHost | select name) {
  Write-Host $VMHost.Name
}
Write-Host "`n"

$VMHostName = Read-Host "�п�J ESXi �D���W��"
Write-Host "`n"

Write-Host "���b���oVMware Tools ���A"
$TotalVMs = (Get-VMHost -name $VMHostName | Get-VM).Count
$CurrentVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsCurrent" } | select name).Count
$NeedUpgradeVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNeedUpgrade" } | select name).Count
$NotInstalledVMs = (Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNotInstalled" } | select name).Count

$VMsToolsStatus = New-Object PSObject -Property ([ordered]@{
    "VMware Tools is installed, and the version is current. (�w�w�˳̷s���`��)"      = $CurrentVMs
    "VMware Tools is installed, but the version is not current. (�ݤɯ��`��)"  = $NeedUpgradeVMs
    "VMware Tools has never been installed. (�|���w��VMware Tools)" = $NotInstalledVMs
    "Total VMs (VM �`�ƶq)" = $TotalVMs
  })

$VMsToolsStatus

Write-Host "�Y�N���X�����п�ܫO�s���|`n"
$csvPath = Get-FileName $currentPath

if ($csvPath -ne "") {
  Write-Host "��X��CSV��" $csvPath "`n"
  Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | select name, @{Name = "ToolsVersion"; Expression = { $_.config.tools.toolsversion } }, @{ Name = "ToolStatus"; Expression = { $_.Guest.ToolsVersionStatus } } | Sort-Object ToolStatus | Export-Csv -Path $csvPath -NoTypeInformation
}
else {
  Write-Host "���|���~! ���O�sCSV"
}

$OutofDateVMs = Get-VMHost -name $VMHostName | Get-VM | % { get-view $_.id } | Where-Object { $_.Guest.ToolsVersionStatus -like "guestToolsNeedUpgrade" } | select name

if ($NeedUpgradeVMs -eq 0) {
  Write-Host "�S�� VM �ݭn�ɯ� Vmware Guest Tools"
}
else {
  Write-Host "�H�U���ݭn��s Vmware Guest Tools �� VM`n"
  Write-Host "VM Name"
  ForEach ($VMs in $OutOfDateVMs) {
    Write-Host $VMs.Name
  }

  $wsh = New-Object -ComObject WScript.Shell
  $answer = $wsh.Popup("�}�l�v�@�ɯ�VM�W��Vmware Guest Tools", 300, "�O�_�ɯ�Vmware Guest Tools", 4 + 32)
  if ($answer -eq 7) {
    Write-Host "`n�z��ܤ��ɯ� Vmware Guest Tools�C"
  }
  if ($answer -eq 6) {
    Write-Host "`n�z��ܤɯ� Vmware Guest Tools�A�}�l�v�@�ɯšC`n"
    ForEach ($VMs in $OutOfDateVMs) {
      Write-Host "Start upgrading Vmware Guest Tools on VM:" $VMs.Name
      Update-Tools $VMs.Name -Verbose -NoReboot
    }
  }
}

Write-Host "`n���_vCenter���s�u��"
Disconnect-VIServer -Server $vCenter -Confirm:$false
Write-Host "`n���浲��"