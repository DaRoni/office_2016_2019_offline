 # Не показывать начальный экран при запуске 
IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DisableBootToOfficeStart -Value 1 -Force 

 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Word\Options)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DisableBootToOfficeStart -Value 1 -Force 
 ### OfficeLinkedIn
  IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Name OfficeLinkedIn -Value 0 -Force

 #### автономный режим
 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerDataOptInReason -Value 1 -Force
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerDataOptIn -Value 1 -Force
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerData -Value 1 -Force
 
 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer -Name UseOnlineContent -Value 2 -Force
 
 ### узнать методом тыка значение параметра 
 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\Internet)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Internet -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Internet -Name UseOnlineContent -Value 0 -Force
#on
#[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Internet]
#"UseOnlineContent"=dword:00000001
#off
#[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Internet]
#пустой

####
 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Name UserContentDisabled -Value 2 -Force
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Name DownloadContentDisabled -Value 2 -Force

 IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous -Name DisconnectedState -Value 2 -Force

IF (!(Test-Path HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson)) 
 { 
     New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson -Force 
 } 
 New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson -Name PTWOptIn -Value 3 -Force
 
sleep 5
