###############################################################################################################
#              Version 1.0 | Tested on Windows 10
#              This is a Script to set the User Experience Profiles in PCoIP Agent
#
#              Profile      Exp      Bandwith        Network      User Roles
#              Profile A	Best	 Highest	     LAN	      Review/CAD and Creative
#              Profile B	Great	 Moderate	     LAN/WAN	  CAD and Creative/Knowledge Worker 
#              Profile C	Good	 Optimized	     WAN	      Knowledge Worker
#              Profile D	Good	 Constrained	 WAN	      Knowledge/Task Worker
#              Profile E	Limited	 Lowest	         Shared WAN	  Task Workers with small screens
#
#              In case of concerns, please reach out to kairon.alves@hp.com
#####################
Write-Host "###### Profile Selection Script v1.0 | HP Anyware" -ForegroundColor Green
Write-Host "Checking PCoIP Agent Installation Status"
$Installcheck = (get-service -Name "PCoIPAgent").Status
if ($Installcheck -eq "Running"){
    #Get EXE File, Type and check Version
    $Path = (Get-CimInstance -ClassName win32_service -Filter "Name Like 'PCoIPAgent'").PathName
    $Path = $Path.Replace('"','')
    $PCoIPVersion = ((get-command $Path).FileVersionInfo).ProductVersion
    Write-Host "PCoIP Agent version $PCoIPVersion" -ForegroundColor Green
	$Pathver = $Path.Replace('\bin\pcoip_agent.exe','')
    $AgentType = (get-content -Path "$Pathver\pcoip_agent_release_manifest.txt") | Select -first 1
    $AgentType = ($AgentType.Replace('INSTALLER_NAME = ','')).replace('.exe','')
    Write-host "$AgentType" -ForegroundColor Green
	

    #Checking Windows Tuning Configuration
    #Checking Visual Effects
    
    $checkWinVisualEffects = (Get-Item -Path 'HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects').Property
    $key_VXpath = 'HKCU:Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects'
    $ValueVX = (Get-ItemProperty -Path $key_VXpath)."VisualFxSetting"

    #Checking Processor Schedling
    $checkWinProcScheduling = (Get-Item -Path 'HKLM:SYSTEM\CurrentControlSet\Control\PriorityControl').Property
    $key_PSpath = 'HKLM:SYSTEM\CurrentControlSet\Control\PriorityControl'
    $ValuePS = (Get-ItemProperty -Path $key_PSpath)."Win32PrioritySeparation"

    #Checking Paging File Config
    $checkWinPagingFiles = (Get-Item -Path 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management').Property
    $key_PFpath = 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
    $ValuePF = (Get-ItemProperty -Path $key_PFpath)."PagingFiles"

    if($checkWinVisualEffects -match "VisualFXSetting" -and "$ValueVX" -eq "2" -and "$ValuePS" -eq "38" -and "$ValuePF" -notmatch "pagefile"){
    
    Write-Host "This Agent is Tunned" -ForegroundColor Green
    } else {
    Write-Host "This Agent is not tunned for PCoIP Config" -ForegroundColor Red
    write-host "Would you like to apply the tunning settings on your Agent? Reboot of the system is Required!" -ForegroundColor Red
    $TunningWindowsQ = Read-Host -Prompt "Y or N"
        If ("$TunningWindowsQ" -eq "Y"){
            # Applying Visual Effects
            if("$checkWinVisualEffects" -eq "VisualFXSetting" -and "$ValueVX" -eq "2"){Write-Host "PC Agent has the recommended VisualFXSetting" -ForegroundColor Green}
            else { 
			write-host "Applying VisualEffects" -ForegroundColor Green
			Set-ItemProperty -Path $key_VXpath -Name VisualFXSetting -Value 2 -Type Dword }
            #Applying Processor Priority
            if("$ValuePS" -eq "38"){Write-Host "PC Agent has the recommended Processor Scheduling Priority" -ForegroundColor Green}
            else { 
			write-host "Applying Proc Priority" -ForegroundColor Green
			Set-ItemProperty -Path $key_PSpath -Name Win32PrioritySeparation -Value 38 }
            #Applying Paging File Config
            if("$ValuePF" -notmatch "pagefile"){Write-Host "PC Agent has the recommended Paging File Configuration" -ForegroundColor Green}
			else {
				write-host "Applying Paging File" -ForegroundColor Green
               $sys = Get-WmiObject Win32_Computersystem –EnableAllPrivileges
               $sys.AutomaticManagedPagefile = $false 
               $sys.put()

               $pagefilelist = (Get-WmiObject Win32_PagefileSetting).Name
               foreach($pagefilename in $pagefilelist){
               $pagedelete = Get-WmiObject Win32_PagefileSetting | Where-Object {$_.name -eq $pagefilename}
               $pagedelete.delete()}
                 }

              ######### Rebooting OS to Apply Tunning Settings
              Write-Host "Rebooting Windows to Apply Tunning Settings" -ForegroundColor Red
			  Start-Sleep -Seconds 5
              Restart-Computer -Force
                     
        }
    }
			#Profile Selection and Configuration Step
			  Write-Host "##### Profile Menu Selection
Profile A, Highest Bandwith and Recommended for Review/CAD and Creative Users
Profile B, Moderate Bandwith and Recommended for CAD and Creative/Knowledge Worker Users
Profile C, Optimized Bandwith and Recommended for Knowledge Worker Users
Profile D, Constrained Bandwith and Recommended for Knowledge/Task Worker users
Profile E, Limited Bandwith and Recommended for Task Workers with small screens users" -ForegroundColor Yellow
			$ProfileSelection = Read-Host -Prompt "Please Select The Profile that you want to apply [A,B,C,D,E]:"
			while($ProfileSelection -ne "A" -and $ProfileSelection -ne "B" -and $ProfileSelection -ne "C" -and $ProfileSelection -ne "D" -and $ProfileSelection -ne "E"){
			write-host "PLEASE SELECT A VALID PROFILE [A,B,C,D,E]" -ForegroundColor red
			$ProfileSelection = Read-Host -Prompt "Please Select The Profile that you want to apply [A,B,C,D,E]:"
			}
			
			#Gathering user details to apply on PCoIP Agent Settings
	
			$HKCUGPPath = "HKLM:SOFTWARE\Policies\Teradici\PCoIP\pcoip_admin_defaults"
			
					if($ProfileSelection -eq "A"){
					write-host "You have Selected the Profile [A]" -ForegroundColor Yellow
					
					# Enable Ultra and Verify pcoip_admin_defaults key is created or not
					
					if(Test-Path -Path $HKCUGPPath){
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 1 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
					}else{
					
					New-Item -Path "HKLM:SOFTWARE\Policies" -Name Teradici
					New-Item -Path "HKLM:Software\Policies\Teradici" -Name PCoIP
					New-Item -Path "HKLM:Software\Policies\Teradici\PCoIP" -Name pcoip_admin_defaults

					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 1 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
							 
					}


					#Chroma Subsampling Graph Only
				#	Set-ItemProperty -Path $HKCUGPPath -Name setting -Value 000 -Type Dword
					
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.yuv_chroma_subsampling -Value 0 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.frame_rate_vs_quality_factor -Value 50 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.use_client_img_settings -Value 0 -Type Dword		
					
								
					#Maximum PCoIP Session Bandwith
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.max_link_rate -Value 900000 -Type Dword

					#Enable BTL
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.enable_build_to_lossless -Value 1 -Type Dword

					#Minimum Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.minimum_image_quality -Value 50 -Type Dword
					
					#Maximum Initial Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_initial_image_quality -Value 90 -Type Dword

					#Maximum Frame Rate (fps)
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_frame_rate -Value 60 -Type Dword

					#Session Audio Bandwidth Limit (kbps)
					$AudioCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -erroraction 'silentlycontinue'
					if($AudioCheck -ne $Null){
					Remove-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit"
					}
					#Windows Visual Effects Optimizations
					#Already Applied During Tunning
					
					
					}
					if($ProfileSelection -eq "B"){
					write-host "You have Selected the Profile [B]" -ForegroundColor Yellow
					
					# Enable Ultra and Verify pcoip_admin_defaults key is created or not
					
					if(Test-Path -Path $HKCUGPPath){
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
					}else{
					
					New-Item -Path "HKLM:SOFTWARE\Policies" -Name Teradici
					New-Item -Path "HKLM:Software\Policies\Teradici" -Name PCoIP
					New-Item -Path "HKLM:Software\Policies\Teradici\PCoIP" -Name pcoip_admin_defaults

					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
							 
					}


					#Chroma Subsampling Graph Only
				#	Set-ItemProperty -Path $HKCUGPPath -Name setting -Value 000 -Type Dword
					
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.yuv_chroma_subsampling -Value 0 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.frame_rate_vs_quality_factor -Value 50 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.use_client_img_settings -Value 0 -Type Dword		
					
								
					#Maximum PCoIP Session Bandwith
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.max_link_rate -Value 900000 -Type Dword

					#Enable BTL
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.enable_build_to_lossless -Value 0 -Type Dword

					#Minimum Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.minimum_image_quality -Value 40 -Type Dword
					
					#Maximum Initial Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_initial_image_quality -Value 80 -Type Dword

					#Maximum Frame Rate (fps)
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_frame_rate -Value 60 -Type Dword

					#Session Audio Bandwidth Limit (kbps)
					$AudioCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -erroraction 'silentlycontinue'
					if($AudioCheck -ne $Null){
						Remove-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit"
					}
						
			
					
					
					}
					if($ProfileSelection -eq "C"){
					write-host "You have Selected the Profile [C]" -ForegroundColor Yellow
					
										
					# Enable Ultra and Verify pcoip_admin_defaults key is created or not
					
					if(Test-Path -Path $HKCUGPPath){
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
					}else{
					
					New-Item -Path "HKLM:SOFTWARE\Policies" -Name Teradici
					New-Item -Path "HKLM:Software\Policies\Teradici" -Name PCoIP
					New-Item -Path "HKLM:Software\Policies\Teradici\PCoIP" -Name pcoip_admin_defaults

					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
							 
					}


					#Chroma Subsampling Graph Only
				#	Set-ItemProperty -Path $HKCUGPPath -Name setting -Value 000 -Type Dword
					
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.yuv_chroma_subsampling -Value 1 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.frame_rate_vs_quality_factor -Value 50 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.use_client_img_settings -Value 0 -Type Dword		
					
								
					#Maximum PCoIP Session Bandwith
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.max_link_rate -Value 10000 -Type Dword

					#Enable BTL
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.enable_build_to_lossless -Value 0 -Type Dword

					#Minimum Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.minimum_image_quality -Value 40 -Type Dword
					
					#Maximum Initial Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_initial_image_quality -Value 80 -Type Dword

					#Maximum Frame Rate (fps)
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_frame_rate -Value 30 -Type Dword

					#Session Audio Bandwidth Limit (kbps)
					$AudioCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -erroraction 'silentlycontinue'
					
					if($AudioCheck -ne $Null){
						Remove-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit"
					}
				
					
					}
					if($ProfileSelection -eq "D"){
					write-host "You have Selected the Profile [D]" -ForegroundColor Yellow
										
					# Enable Ultra and Verify pcoip_admin_defaults key is created or not
					
					if(Test-Path -Path $HKCUGPPath){
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
					}else{
					
					New-Item -Path "HKLM:SOFTWARE\Policies" -Name Teradici
					New-Item -Path "HKLM:Software\Policies\Teradici" -Name PCoIP
					New-Item -Path "HKLM:Software\Policies\Teradici\PCoIP" -Name pcoip_admin_defaults

					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
							 
					}


					#Chroma Subsampling Graph Only
				#	Set-ItemProperty -Path $HKCUGPPath -Name setting -Value 000 -Type Dword
					
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.yuv_chroma_subsampling -Value 1 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.frame_rate_vs_quality_factor -Value 50 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.use_client_img_settings -Value 0 -Type Dword		
					
								
					#Maximum PCoIP Session Bandwith
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.max_link_rate -Value 2400 -Type Dword

					#Enable BTL
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.enable_build_to_lossless -Value 0 -Type Dword

					#Minimum Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.minimum_image_quality -Value 30 -Type Dword
					
					#Maximum Initial Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_initial_image_quality -Value 70 -Type Dword

					#Maximum Frame Rate (fps)
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_frame_rate -Value 16 -Type Dword

					#Session Audio Bandwidth Limit (kbps)
					$AudioCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -erroraction 'silentlycontinue'
					
					if($AudioCheck -eq $Null){
						Set-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -Value 48 -Type Dword
					}
					
					
					}
					if($ProfileSelection -eq "E"){
					write-host "You have Selected the Profile [E]" -ForegroundColor Yellow
					
					# Disable Ultra and Verify pcoip_admin_defaults key is created or not
					
					if(Test-Path -Path $HKCUGPPath){
					$UltraCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.ultra" -erroraction 'silentlycontinue'
					if($UltraCheck -ne $null){
					Remove-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra
					Remove-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps}
					}else{
					
					New-Item -Path "HKLM:SOFTWARE\Policies" -Name Teradici
					New-Item -Path "HKLM:Software\Policies\Teradici" -Name PCoIP
					New-Item -Path "HKLM:Software\Policies\Teradici\PCoIP" -Name pcoip_admin_defaults

					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra -Value 3 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.ultra_offload_mpps -Value 10 -Type Dword
							 
					}


					#Chroma Subsampling Graph Only
								
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.yuv_chroma_subsampling -Value 1 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.frame_rate_vs_quality_factor -Value 50 -Type Dword
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.use_client_img_settings -Value 0 -Type Dword		
					
								
					#Maximum PCoIP Session Bandwith
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.max_link_rate -Value 600 -Type Dword

					#Enable BTL
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.enable_build_to_lossless -Value 0 -Type Dword

					#Minimum Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.minimum_image_quality -Value 30 -Type Dword
					
					#Maximum Initial Image Quality
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_initial_image_quality -Value 70 -Type Dword

					#Maximum Frame Rate (fps)
					Set-ItemProperty -Path $HKCUGPPath -Name pcoip.maximum_frame_rate -Value 8 -Type Dword

					#Session Audio Bandwidth Limit (kbps)
					$AudioCheck = Get-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -erroraction 'silentlycontinue'
					
					if($AudioCheck -eq $Null){
						Set-ItemProperty -Path $HKCUGPPath -Name "pcoip.audio_bandwidth_limit" -Value 32 -Type Dword
					}
					
					}
			
			write-host "Please Reconnect to your PCoIP Session to Apply the changes" -ForegroundColor Yellow

} else {
Write-Host "PCoIP Agent is not installed or Service is not running, Please verify" -ForegroundColor Red
Start-Sleep -Seconds 5
exit}
