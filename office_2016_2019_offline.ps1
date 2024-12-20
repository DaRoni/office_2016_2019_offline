# 2024.12.20
write-host "Перевод в автономный режим MS Office 2016-2019"
# Outlook Social Connector disable
IF (!(Test-Path -Path  HKCU:\software\policies\microsoft\office\outlook\socialconnector))
{
	New-Item -Path HKCU:\software\policies\microsoft\office\outlook\socialconnector -Force }
	New-ItemProperty -Path HKCU:\software\policies\microsoft\office\outlook\socialconnector -Name runosc -PropertyType DWord -Value 1 -Force

# Socialnetworkconnectivity
IF (!(Test-Path -Path HKCU:\software\policies\microsoft\office\outlook\socialconnector))
{
	New-Item -Path HKCU:\software\policies\microsoft\office\outlook\socialconnector -Force }
	New-ItemProperty -Path HKCU:\software\policies\microsoft\office\outlook\socialconnector -Name disablesocialnetworkconnectivity -PropertyType DWord -Value 1 -Force

# Outlook disable attachment prewiew
IF (!(Test-Path -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences))
{
	New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences -Force }
	New-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences -Name DisableAttachmentPreviewing -PropertyType DWord -Value 1 -Force
# Не показывать начальный экран при запуске 
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options -Name DisableBootToOfficeStart -Value 1 -Force 
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Word\Options -Name DisableBootToOfficeStart -Value 1 -Force 
#### Автономный режим
IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\LinkedIn -Name OfficeLinkedIn -Value 0 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer -Name UseOnlineContent -PropertyType DWord -Value 1 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\Internet))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Internet -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Internet -Name UseOnlineContent -PropertyType DWord -Value 1 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\Privacy))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Name UserContentDisabled -PropertyType DWord -Value 2 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy -Name DownloadContentDisabled -PropertyType DWord -Value 2 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous -Name DisconnectedState -PropertyType DWord -Value 2 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous -Name ControllerConnectedServicesState -PropertyType DWord -Value 2 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson -Name PTWOptIn -PropertyType DWord  -Value 0 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options -Name DiscoveryNeedOptIn -PropertyType DWord  -Value 1 -Force

IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Common\Security\FileValidation))
		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Common\Security\FileValidation -Force }
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common\Security\FileValidation -Name DisableReporting -PropertyType DWord  -Value 1 -Force

New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerDataOptInReason -PropertyType DWord -Value 0 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerDataOptIn -PropertyType DWord -Value 0 -Force
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Common -Name SendCustomerData -PropertyType DWord -Value 0 -Force

# Autodiscover off
#IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover))
#		{New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Force }		
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeHttpRedirect -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeHttpsAutoDiscoverDomain -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeHttpsRootDomain -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeScpLookup -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeSrvRecord -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeLastKnownGoodURL -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name ExcludeExplicitO365Endpoint -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name DisableAutoDiscoverv2Service -PropertyType DWord  -Value 1 -Force


#IF (!(Test-Path -Path  HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover))
#		{New-Item -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Force }		
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeHttpRedirect -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeHttpsAutoDiscoverDomain -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeHttpsRootDomain -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeScpLookup -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeSrvRecord -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Autodiscover -Name ExcludeLastKnownGoodURL -PropertyType DWord  -Value 1 -Force
#New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Autodiscover -Name DisableAutoDiscoverv2Service -PropertyType DWord  -Value 1 -Force