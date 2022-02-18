; #FUNCTION# ====================================================================================================================
; Name ..........: OpenBS
; Description ...:
; Syntax ........: OpenBS([$bRestart = False])
; Parameters ....: $bRestart            - [optional] a boolean value. Default is False.
; Return values .: None
; Author ........: GkevinOD (2014), Hervidero (2015)
; Modified ......: Cosote (12-2015), KnowJack (08-2015)
; Remarks .......: This file is part of MyBot, previously known as ClashGameBot. Copyright 2015-2019
;                  MyBot is distributed under the terms of the GNU GPL
; Related .......:
; Link ..........: https://github.com/MyBotRun/MyBot/wiki
; Example .......: No
; ===============================================================================================================================

Func OpenBlueStacks5($bRestart = False)
	SetLog("Starting BlueStacks and Clash Of Clans", $COLOR_SUCCESS)

	Local $hTimer, $iCount = 0, $cmdOutput, $process_killed, $i, $connected_to, $PID, $cmdPar

	CloseUnsupportedBlueStacks5()

	; always start ADB first to avoid ADB connection problems
	LaunchConsole($g_sAndroidAdbPath, AddSpace($g_sAndroidAdbGlobalOptions) & "start-server", $process_killed)

	$hTimer = __TimerInit()
	WinGetAndroidHandle()
	Local $bStopIfLaunchFails = False
	While $g_hAndroidControl = 0
		If Not $g_bRunState Then Return False
		; check that HD-Frontend.exe process is really there
		;$cmdPar = " -t " & $g_sAndroidInstance
		$cmdPar = GetAndroidProgramParameter()
		$PID = LaunchAndroid($g_sAndroidProgramPath, $cmdPar, $g_sAndroidPath, Default, $bStopIfLaunchFails)
		If $PID > 0 Then $PID = ProcessExists2($g_sAndroidProgramPath, $cmdPar)
		If $PID <= 0 Then
			CloseAndroid("OpenBlueStacks5")
			SetScreenBlueStacks5()
			$bStopIfLaunchFails = True
			If _Sleep(1000) Then Return False
		EndIf

		_StatusUpdateTime($hTimer)
		If __TimerDiff($hTimer) > $g_iAndroidLaunchWaitSec * 1000 Or ($PID = 0 And $bStopIfLaunchFails = True) Then ; if no BS position returned in 4 minutes, BS/PC has major issue so exit
			SetScreenBlueStacks5()
			SetLog("Serious error has occurred, please restart PC and try again", $COLOR_ERROR)
			SetLog($g_sAndroidEmulator & " refuses to load, waited " & Round(__TimerDiff($hTimer) / 1000, 2) & " seconds", $COLOR_ERROR)
			SetError(1, @extended, False)
			Return False
		EndIf
		If _Sleep(3000) Then Return False
		_StatusUpdateTime($hTimer, $g_sAndroidEmulator & " Starting")
		WinGetAndroidHandle()
	WEnd

	; enable window title so BS2 can be moved again
	WinGetAndroidHandle()
	Local $aWin = WinGetPos($g_hAndroidWindow)
	Local $lCurStyle = _WinAPI_GetWindowLong($g_hAndroidWindow, $GWL_STYLE)
	; Enable Title Bar and Border
	_WinAPI_SetWindowLong($g_hAndroidWindow, $GWL_STYLE, BitOR($lCurStyle, $WS_CAPTION, $WS_SYSMENU))
	Local $iCaptionHeight = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
	If BitAND($lCurStyle, BitOR($WS_CAPTION, $WS_SYSMENU)) <> BitOR($WS_CAPTION, $WS_SYSMENU) And UBound($aWin) > 3 Then
		; adjust window height due to caption
		WinMove2($g_hAndroidWindow, "", $aWin[0], $aWin[1], $aWin[2], $aWin[3] + $iCaptionHeight)
	EndIf
	;_WinAPI_SetWindowPos($g_hAndroidWindow, 0, 0, 0, 0, 0, BitOr($SWP_NOMOVE, $SWP_NOSIZE, $SWP_FRAMECHANGED)) ; redraw

	If $g_hAndroidControl Then
		$connected_to = ConnectAndroidAdb(False, 3000) ; small time-out as ADB connection must be available now

		;If WaitForDeviceBlueStacks5($g_iAndroidLaunchWaitSec - __TimerDiff($hTimer) / 1000, $hTimer) Then Return
		If WaitForAndroidBootCompleted($g_iAndroidLaunchWaitSec - __TimerDiff($hTimer) / 1000, $hTimer) Then Return
		If Not $g_bRunState Then Return False

		SetLog($g_sAndroidEmulator & " Loaded, took " & Round(__TimerDiff($hTimer) / 1000, 2) & " seconds to begin.", $COLOR_SUCCESS)
		AndroidAdbLaunchShellInstance()

		If Not $g_bRunState Then Return False
		ConfigBlueStacks5WindowManager()

		Return True

	EndIf

	Return False

EndFunc   ;==>OpenBlueStacks5

Func GetBlueStacks5AdbPath()
	Local $adbPath = $__BlueStacks_nxt_Path & "HD-Adb.exe"
	If FileExists($adbPath) Then Return $adbPath
	Return ""
EndFunc   ;==>GetBlueStacks5AdbPath


Func InitBlueStacksNxt($bCheckOnly = False, $bAdjustResolution = False, $bLegacyMode = False)

	Local $frontend_exe = ["HD-Player.exe"]

	Local $i, $aFiles = [$frontend_exe, "HD-Adb.exe"] ; first element can be $frontend_exe array!
	Local $Values[4][3] = [ _
			["Screen Width", $g_iAndroidClientWidth, $g_iAndroidClientWidth], _
			["Screen Height", $g_iAndroidClientHeight, $g_iAndroidClientHeight], _
			["Window Width", $g_iAndroidWindowWidth, $g_iAndroidWindowWidth], _
			["Window Height", $g_iAndroidWindowHeight, $g_iAndroidWindowHeight] _
			]
	Local $bChanged = False

	$__BlueStacks_nxt_Version = RegRead($g_sHKLM & "\SOFTWARE\BlueStacks_nxt\", "Version")
	$__BlueStacks_nxt_Path = RegRead($g_sHKLM & "\SOFTWARE\BlueStacks_nxt\", "InstallDir")
	$__BlueStacks_nxt_User_Path = RegRead($g_sHKLM & "\SOFTWARE\BlueStacks_nxt\", "UserDefinedDir")
	If @error <> 0 Then
		$__BlueStacks_nxt_Path = @ProgramFilesDir & "\BlueStacks_nxt\"
		$__BlueStacks_nxt_User_Path = @AppDataCommonDir & "\BlueStacks_nxt\"
		SetError(0, 0, 0)
	EndIf
	If StringRight($__BlueStacks_nxt_Path, 1) <> "\" Then $__BlueStacks_nxt_Path &= "\"
	$__BlueStacks_nxt_Path = StringReplace($__BlueStacks_nxt_Path, "\\", "\")
	If StringRight($__BlueStacks_nxt_User_Path, 1) <> "\" Then $__BlueStacks_nxt_User_Path &= "\"
	$__BlueStacks_nxt_User_Path = StringReplace($__BlueStacks_nxt_User_Path, "\\", "\")

	Local $sPreferredADB = FindPreferredAdbPath()
	If $sPreferredADB Then _ArrayDelete($aFiles, 1)

	For $i = 0 To UBound($aFiles) - 1
		Local $File
		Local $bFileFound = False
		Local $aFiles2 = $aFiles[$i]
		If Not IsArray($aFiles2) Then Local $aFiles2 = [$aFiles[$i]]
		For $j = 0 To UBound($aFiles2) - 1
			$File = $__BlueStacks_nxt_Path & $aFiles2[$j]
			$bFileFound = FileExists($File)
			If $bFileFound Then
				; check if $frontend_exe is array, then convert
				If $i = 0 And IsArray($frontend_exe) Then $frontend_exe = $aFiles2[$j]
				ExitLoop
			EndIf
		Next
		If Not $bFileFound Then
			If Not $bCheckOnly Then
				SetLog("Serious error has occurred: Cannot find " & $g_sAndroidEmulator & ":", $COLOR_ERROR)
				SetLog($File, $COLOR_ERROR)
				SetError(1, @extended, False)
			EndIf
			Return False
		EndIf
	Next

	If Not $bCheckOnly Then

		OpenBlueStacks5Adb()

		$g_iAndroidAdbSuCommand = "/system/xbin/bstk/su"
		Local $BootParameter = RegRead($g_sHKLM & "\SOFTWARE\BlueStacks\Guests\" & $g_sAndroidInstance & "\", "BootParameters")
		Local $OEMFeatures
		Local $aRegExResult = StringRegExp($BootParameter, "OEMFEATURES=(\d+)", $STR_REGEXPARRAYGLOBALMATCH)
		If Not @error Then
			; get last match!
			$OEMFeatures = $aRegExResult[UBound($aRegExResult) - 1]
			$g_bAndroidHasSystemBar = BitAND($OEMFeatures, 0x000001) = 0
		EndIf

		; update global variables
		$g_sAndroidPath = $__BlueStacks_nxt_Path
		$g_sAndroidProgramPath = $__BlueStacks_nxt_Path & $frontend_exe
		$g_sAndroidAdbPath = $sPreferredADB
		If $g_sAndroidAdbPath = "" Then $g_sAndroidAdbPath = $__BlueStacks_nxt_Path & "HD-Adb.exe"
		$g_sAndroidVersion = $__BlueStacks_nxt_Version

		ConfigureSharedFolderBlueStacks5(0) ; something like D:\ProgramData\BlueStacks_nxt\Engine\UserData\SharedFolder\

		SetDebugLog($g_sAndroidEmulator & " OEM Features: " & $OEMFeatures)
		SetDebugLog($g_sAndroidEmulator & " System Bar is " & ($g_bAndroidHasSystemBar ? "" : "not ") & "available")
		#cs as of 2016-01-26 CoC release, system bar is transparent and should be closed when bot is running
			If $bAdjustResolution Then
			If $g_bAndroidHasSystemBar Then
			;$Values[0][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][5]
			$Values[1][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][6] + $__BlueStacks_SystemBar
			;$Values[2][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][7]
			$Values[3][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][8] + $__BlueStacks_SystemBar
			Else
			;$Values[0][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][5]
			$Values[1][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][6]
			;$Values[2][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][7]
			$Values[3][2] = $g_avAndroidAppConfig[$g_iAndroidConfig][8]
			EndIf
			EndIf
			$g_iAndroidClientWidth = $Values[0][2]
			$g_iAndroidClientHeight = $Values[1][2]
			$g_iAndroidWindowWidth =  $Values[2][2]
			$g_iAndroidWindowHeight = $Values[3][2]
		#ce

		For $i = 0 To UBound($Values) - 1
			If $Values[$i][1] <> $Values[$i][2] Then
				$bChanged = True
				SetDebugLog($g_sAndroidEmulator & " " & $Values[$i][0] & " updated from " & $Values[$i][1] & " to " & $Values[$i][2])
			EndIf
		Next

		WinGetAndroidHandle()
	EndIf

	Return True

EndFunc   ;==>InitBlueStacksNxt

Func ConfigureSharedFolderBlueStacks5($iMode = 0, $bSetLog = Default) ; TODO bst.shared_folders="Documents,Pictures,InputMapper,BstSharedFolder"
	If $bSetLog = Default Then $bSetLog = True
	Local $bResult = False

	Switch $iMode
		Case 0 ; check that shared folder is configured in VM
			For $i = 0 To 5
				If StringInStr(ConfigRead("bst.shared_folders"), "BstSharedFolder") > 0 Then
					$bResult = True
					$g_bAndroidSharedFolderAvailable = True
					$g_sAndroidPicturesPath = "/storage/sdcard/windows/BstSharedFolder/"
					$g_sAndroidPicturesHostPath = $__BlueStacks_nxt_User_Path & "Engine\UserData\SharedFolder\"
					ExitLoop
				EndIf
			Next
		Case 1 ; create missing shared folder
		Case 2 ; Configure VM and add missing shared folder
	EndSwitch

	Return SetError(0, 0, $bResult)

EndFunc   ;==>ConfigureSharedFolderBlueStacks5

Func InitBlueStacks5($bCheckOnly = False)
	Local $bInstalled = InitBlueStacksNxt($bCheckOnly, True)
	If $bInstalled And StringInStr($__BlueStacks_nxt_Version, "5.") <> 1  Then
		If Not $bCheckOnly Then
			SetLog("BlueStacks supported version 5.x not found", $COLOR_ERROR)
			SetError(1, @extended, False)
		EndIf
		Return False
	EndIf

	If $bInstalled And Not $bCheckOnly Then
		$__VBoxManage_Path = $__BlueStacks_nxt_Path & "BstkVMMgr.exe"
		Local $bsNow = GetVersionNormalized($__BlueStacks_nxt_Version)
		If $bsNow > GetVersionNormalized("4.0") Then
			; only Version 4 requires new options
			;$g_sAndroidAdbInstanceShellOptions = " -t -t" ; Additional shell options, only used by BlueStacks5 " -t -t"
			$g_sAndroidAdbShellOptions = " /data/anr/../../system/xbin/bstk/su root" ; Additional shell options when launch shell with command, only used by BlueStacks5 " /data/anr/../../system/xbin/bstk/su root"

			; tcp forward not working in BS4
			$g_iAndroidAdbMinitouchMode = 1
		EndIf

		CheckBlueStacksVersionMod()

		; read ADB port
		Local $BstAdbPort = ConfigRead("bst.instance.Nougat64.status.adb_port")
		If $BstAdbPort Then
			$g_sAndroidAdbDevice = "127.0.0.1:" & $BstAdbPort
		Else
			; use default
			$g_sAndroidAdbDevice = $g_avAndroidAppConfig[$__BS2_Idx][10]
		EndIf
	EndIf

	Return $bInstalled
EndFunc   ;==>InitBlueStacks5

; Will Check all the differences between versions
Func CheckBlueStacks5VersionMod()
	Local $bsNow = GetVersionNormalized($__BlueStacks_nxt_Version)
	Local $aOff = [0, 13]
	; < 2.6.105.x - BS2
	; Undocked -> Zoomout [OK] , Mouse[OK]
	; Docked -> Zoomout [OK] , Mouse[OK]
	; $__BlueStacks5Version_2_5_or_later = False

	Local $bs3 = GetVersionNormalized("2.50.0.0")
	; 2.50.53.x - BS3
	; Undocked -> Zoomout [OK] , Mouse[Need compensation]
	; Docked -> Zoomout [OK] , Mouse[OK]
	; $__BlueStacks5Version_2_5_or_later = False

	Local $bs3WithFrame = GetVersionNormalized("2.56.75")
	; 2.56.75 excelent version - 2.56.77 - BS3
	; Undocked -> Zoomout [NO*] , Mouse[OK] ; *$__BlueStacks5Version_2_5_or_later = True
	; Docked -> Zoomout [OK] , Mouse[OK]

	Local $bs3NNoFrame = GetVersionNormalized("4.0.0.0")
	; 4.2.1.9724 - New N version
	; Undocked -> Zoomout [OK] , Mouse[Need compensation]
	; Docked -> Zoomout [OK] , Mouse[OK]

	Local $bs3NWithFrame = GetVersionNormalized("4.3.28.0")
	; 4.3.28.4020 Last version
	; Undocked -> Zoomout [NO*] , Mouse[OK] ; *$__BlueStacks5Version_2_5_or_later = True
	; Docked -> Zoomout [OK] , Mouse[OK]

	If ($bsNow >= $bs3 And $bsNow < $bs3WithFrame) Or ($bsNow > $bs3NNoFrame And $bsNow < $bs3NWithFrame) Then
		; Mouse clicks in Window are off by -13 on Y-axis, so set special value now
		If $g_aiMouseOffsetWindowOnly[0] <> $aOff[0] Or $g_aiMouseOffsetWindowOnly[1] <> $aOff[1] Then
			$g_aiMouseOffsetWindowOnly = $aOff
			SetDebugLog("BlueStacks " & $__BlueStacks_nxt_Version & ": Adjust mouse clicks when running undocked by: " & $aOff[0] & ", " & $aOff[1])
		EndIf
	EndIf

	;Zoomout Function when is not Docked
	If $bsNow >= $bs3NWithFrame Or ($bsNow >= $bs3WithFrame And $bsNow < $bs3NNoFrame) Then
		SetDebugLog("BlueStacks " & $__BlueStacks_nxt_Version & " adjustment on ZoomOut")
		$__BlueStacks5Version_2_5_or_later = True
	EndIf

EndFunc   ;==>CheckBlueStacks5VersionMod

Func GetBlueStacks5BackgroundMode()
	; check if BlueStacks 5 is running in OpenGL mode
	Local $GlRenderMode = ConfigRead("bst.instance.Nougat64.graphics_renderer")
	Switch $GlRenderMode
		Case "dx"
			; DirectX
			Return $g_iAndroidBackgroundModeDirectX
		Case "gl"
			; OpenGL
			Return $g_iAndroidBackgroundModeOpenGL
		Case Else
			SetLog($g_sAndroidEmulator & " unsupported render mode " & $GlRenderMode, $COLOR_WARNING)
			Return 0
	EndSwitch
EndFunc   ;==>GetBlueStacks5BackgroundMode

; Called from checkMainScreen
Func RestartBlueStacks5CoC()
	If Not $g_bRunState Then Return False
	Local $cmdOutput
	If Not InitAndroid() Then Return False
	If WinGetAndroidHandle() = 0 Then Return False
	$cmdOutput = AndroidAdbSendShellCommand("am start -W -n " & $g_sAndroidGamePackage & "/" & $g_sAndroidGameClass, 60000) ; timeout of 1 Minute ; disabled -S due to long wait after 2017 Dec. Update
	SetLog("Please wait for CoC restart......", $COLOR_INFO) ; Let user know we need time...
	Return True
EndFunc   ;==>RestartBlueStacks5CoC

Func CheckScreenBlueStacks5($bSetLog = True)
	Local $aValues[3][2] = [ _
			["dpi", 160], _
			["fb_height", $g_iAndroidClientHeight], _
			["fb_width", $g_iAndroidClientWidth], _
			]
	Local $i, $Value, $iErrCnt = 0
	For $i = 0 To UBound($aValues) - 1
		$Value = ConfigRead("bst.instance.Nougat64." & $aValues[$i][0])
		If $Value <> $aValues[$i][1] Then
			If $iErrCnt = 0 Then
				SetDebugLog("MyBot doesn't work with " & $g_sAndroidEmulator & " screen configuration!", $COLOR_ERROR)
			EndIf
			SetDebugLog("Setting of " & $aValues[$i][0] & " is " & $Value & " and will be changed to " & $aValues[$i][1], $COLOR_ERROR)
			$iErrCnt += 1
		EndIf
	Next
	If $iErrCnt > 0 Then Return False
	Return True
EndFunc   ;==>CheckScreenBlueStacks5

Func SetScreenBlueStacks5()

	If Not InitAndroid() Then Return False

	Local $cmdOutput, $process_killed, $aConfig

	Local $sConfigFile = GetBlueStacks5ConfigFile()
	If FileExists($sConfigFile) Then
		SetDebugLog("Configure BlueStacks5 screen config: " & $sConfigFile)
		_FileReadToArray($sConfigFile, $aConfig, $FRTA_NOCOUNT)
		_ChangeValueForKey($aConfig, "bst.instance.Nougat64.fb_height", $g_iAndroidClientHeight)
		_ChangeValueForKey($aConfig, "bst.instance.Nougat64.fb_width", $g_iAndroidClientWidth)
		_ChangeValueForKey($aConfig, "bst.instance.Nougat64.dpi", "160")
		_FileWriteFromArray($sConfigFile, $aConfig)
	Else
		SetDebugLog("Cannot find BlueStacks5 config to cnfigure screen: " & $sConfig, $COLOR_ERROR)
	EndIf

	Return True

EndFunc   ;==>SetScreenBlueStacks5

Func ConfigBlueStacks5WindowManager()
	If Not $g_bRunState Then Return
	Local $cmdOutput
	; shell wm density 160
	; shell wm size 860x672
	; shell reboot

	; Reset Window Manager size
	$cmdOutput = AndroidAdbSendShellCommand("wm size reset", Default, Default, False)

	; Set expected dpi
	$cmdOutput = AndroidAdbSendShellCommand("wm density 160", Default, Default, False)

	; Set font size to normal
	AndroidSetFontSizeNormal()
EndFunc   ;==>ConfigBlueStacks5WindowManager

Func RebootBlueStacks5SetScreen($bOpenAndroid = True)

	;RebootAndroidSetScreenDefault()

	If Not InitAndroid() Then Return False

	ConfigBlueStacks5WindowManager()

	; Close Android
	CloseAndroid("RebootBlueStacks5SetScreen")
	If _Sleep(1000) Then Return False

	SetScreenAndroid()
	If Not $g_bRunState Then Return False

	If $bOpenAndroid Then
		; Start Android
		OpenAndroid(True)
	EndIf

	Return True

EndFunc   ;==>RebootBlueStacks5SetScreen

Func GetBlueStacks5RunningInstance($bStrictCheck = True)
	WinGetAndroidHandle()
	Local $a[2] = [$g_hAndroidWindow, ""]
	If $g_hAndroidWindow <> 0 Then Return $a
	If $bStrictCheck Then Return False
	Local $WinTitleMatchMode = Opt("WinTitleMatchMode", -3) ; in recent 2.3.x can be also "BlueStacks App Player"
	Local $h = WinGetHandle("Bluestacks", "") ; Need fixing as BS2 Emulator can have that title when configured in registry
	If @error = 0 Then
		$a[0] = $h
	EndIf
	Opt("WinTitleMatchMode", $WinTitleMatchMode)
	Return $a
EndFunc   ;==>GetBlueStacks5RunningInstance

Func GetBlueStacks5ProgramParameter($bAlternative = False)
	Return "--instance " & $g_sAndroidInstance
EndFunc   ;==>GetBlueStacks5ProgramParameter

Func BlueStacks5BotStartEvent()
	If $g_bAndroidEmbedded = False Then
		SetDebugLog("Disable " & $g_sAndroidEmulator & " minimize/maximize Window Buttons")
		DisableBS5($g_hAndroidWindow, $SC_MINIMIZE)
		DisableBS5($g_hAndroidWindow, $SC_MAXIMIZE)
	EndIf
	If $g_bAndroidHasSystemBar Then Return AndroidCloseSystemBar()
	Return False
EndFunc   ;==>BlueStacks5BotStartEvent

Func BlueStacks5BotStopEvent()
	If $g_bAndroidEmbedded = False Then
		SetDebugLog("Enable " & $g_sAndroidEmulator & " minimize/maximize Window Buttons")
		EnableBS5($g_hAndroidWindow, $SC_MINIMIZE)
		EnableBS5($g_hAndroidWindow, $SC_MAXIMIZE)
	EndIf
	If $g_bAndroidHasSystemBar Then Return AndroidOpenSystemBar()
	Return False
EndFunc   ;==>BlueStacks5BotStopEvent

Func BlueStacks5AdjustClickCoordinates(ByRef $x, ByRef $y)
	$x = Round(32767.0 / $g_iAndroidClientWidth * $x)
	$y = Round(32767.0 / $g_iAndroidClientHeight * $y)
	;Local $Num = 32728
	;$x = Int(($Num * $x) / $g_iAndroidClientWidth)
	;$y = Int(($Num * $y) / $g_iAndroidClientHeight)
EndFunc   ;==>BlueStacks5AdjustClickCoordinates

Func DisableBS5($HWnD, $iButton)
	Local $hSysMenu = _GUICtrlMenu_GetSystemMenu($HWnD, 0)
	_GUICtrlMenu_RemoveMenu($hSysMenu, $iButton, False)
	_GUICtrlMenu_DrawMenuBar($HWnD)
EndFunc   ;==>DisableBS5

Func EnableBS5($HWnD, $iButton)
	Local $hSysMenu = _GUICtrlMenu_GetSystemMenu($HWnD, 1)
	_GUICtrlMenu_RemoveMenu($hSysMenu, $iButton, False)
	_GUICtrlMenu_DrawMenuBar($HWnD)
EndFunc   ;==>EnableBS5

Func GetBlueStacksSvcPid5()

	; find process PID
	Local $aFiles = ["HD-Plus-Service.exe", "HD-Service.exe"]
	For $sFile In $aFiles
		Local $PID
		$PID = ProcessExists2($sFile, $g_sAndroidInstance)
		If $PID Then Return $PID
	Next
	Return 0

EndFunc   ;==>GetBlueStacksSvcPid2

Func CloseBlueStacks5()

	Local $bOops = False

	If Not InitAndroid() Then Return

	If Not CloseUnsupportedBlueStacks5(False) And GetVersionNormalized($g_sAndroidVersion) > GetVersionNormalized("2.10") Then
		; BlueStacks 3 supports multiple instance
		Local $aFiles = ["HD-Frontend.exe", "HD-Plus-Service.exe", "HD-Service.exe"]

		Local $bError = False
		For $sFile In $aFiles
			Local $PID
			$PID = ProcessExists2($sFile, $g_sAndroidInstance)
			If $PID Then
				ShellExecute(@WindowsDir & "\System32\taskkill.exe", " -f -t -pid " & $PID, "", Default, @SW_HIDE)
				If _Sleep(1000) Then Return ; Give OS time to work
			EndIf
		Next
		If _Sleep(1000) Then Return ; Give OS time to work
		For $sFile In $aFiles
			Local $PID
			$PID = ProcessExists2($sFile, $g_sAndroidInstance)
			If $PID Then
				SetLog($g_sAndroidEmulator & " failed to kill " & $sFile, $COLOR_ERROR)
			EndIf
		Next

		; also close vm
		CloseVboxAndroidSvc()
	Else
		SetDebugLog("Closing BlueStacks: " & $__BlueStacks_nxt_Path & "HD-Quit.exe")
		RunWait($__BlueStacks_nxt_Path & "HD-Quit.exe")
		If @error <> 0 Then
			SetLog($g_sAndroidEmulator & " failed to quit", $COLOR_ERROR)
			;SetError(1, @extended, -1)
			;Return False
		EndIf
	EndIf

	If _Sleep(2000) Then Return ; wait a bit

	If $bOops Then
		SetError(1, @extended, -1)
	EndIf

EndFunc   ;==>CloseBlueStacks5

Func CloseUnsupportedBlueStacks5($bClose = True) ;TODO
	Local $WinTitleMatchMode = Opt("WinTitleMatchMode", -3) ; in recent 2.3.x can be also "BlueStacks App Player"
	Local $sPartnerExePath = RegRead($g_sHKLM & "\SOFTWARE\BlueStacks\Config\", "PartnerExePath")
	If IsArray(ControlGetPos("Bluestacks App Player", "", "")) Or ($sPartnerExePath And ProcessExists2($sPartnerExePath)) Then ; $g_avAndroidAppConfig[1][4]
		Opt("WinTitleMatchMode", $WinTitleMatchMode)
		; Offical "Bluestacks App Player" v2.0 not supported because it changes the Android Screen!!!
		If $bClose = True Then
			SetLog("MyBot doesn't work with " & $g_sAndroidEmulator & " App Player", $COLOR_ERROR)
			SetLog("Please let MyBot start " & $g_sAndroidEmulator & " automatically", $COLOR_INFO)
			RebootBlueStacks5SetScreen(False)
		EndIf
		Return True
	EndIf
	Opt("WinTitleMatchMode", $WinTitleMatchMode)
	Return False
EndFunc   ;==>CloseUnsupportedBlueStacks5

Func GetBlueStacks5ConfigFile()
	Return $__BlueStacks_nxt_User_Path & "bluestacks.conf"
EndFunc   ;==>GetBlueStacks5ConfigFile

Func _ChangeValueForKey(ByRef $aConfig, $sKey, $sValue)
    For $i = 0 To UBound($aConfig) - 1
        If StringLeft($aConfig[$i], StringLen($sKey) + 1) = $sKey & '=' Then
            $aConfig[$i] = $sKey & '=' & '"' & $sValue & '"'
            Return
        EndIf
    Next
EndFunc   ;==>_ChangeValueForKey

Func ConfigRead($sKey)
	Local $aConfig, $aSplit
	Local $sConfigFile = GetBlueStacks5ConfigFile()
	_FileReadToArray($sConfigFile, $aConfig, $FRTA_NOCOUNT)
    For $i = 0 To UBound($aConfig) - 1
        If StringLeft($aConfig[$i], StringLen($sKey) + 1) = $sKey & '=' Then
			$aSplit = StringSplit($aConfig[$i], '=')
			SetDebugLog("ConfigRead" & $aSplit[1] & "=" & $aSplit[2])
            Return StringMid($aSplit[2], 2, StringLen($aSplit[2]) - 2)
        EndIf
    Next
	Return Null
EndFunc   ;==>_ChangeValueForKey

Func OpenBlueStacks5Adb()
	Local $aConfig
	Local $sConfigFile = GetBlueStacks5ConfigFile()
	If FileExists($sConfigFile) Then
		SetDebugLog("Configure BlueStacks5 adb config: " & $sConfigFile)
		_FileReadToArray($sConfigFile, $aConfig, $FRTA_NOCOUNT)
		_ChangeValueForKey($aConfig, "bst.enable_adb_access", "1")
		_FileWriteFromArray($sConfigFile, $aConfig)
	Else
		SetDebugLog("Cannot find BlueStacks5 config to cnfigure screen: " & $sConfig, $COLOR_ERROR)
	EndIf

	Return True

EndFunc   ;==>SetScreenBlueStacks5
