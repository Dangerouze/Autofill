#NoEnv
#SingleInstance Ignore
#Persistent
SendMode Input

AppVersion = 5.1.0

Gui, Add, CheckBox, x105 y57 vSelectRTN ,
Gui, Add, Text, x23 y57 w80 Left, Selection Return
Gui, Add, Button, x190 y54 h20 w100 gUpdate, Check of Updates

Gui, Show, w300 h130, GHG Autofill
Gui, Add, Text, x8 y12 w80 Right, Excell Window
Gui, Add, Edit, x156 y10 w130 h20 vExcellWnd ReadOnly Left,
Gui, Add, Text, x8 y35 w80 Right, GHG Window
Gui, Add, Edit, x156 y33 w130 h20 vGHGWnd ReadOnly Left,
Gui, Add, Text, x8 y77 w80 Right, Speed
Gui, Add, DropDownList, x93 y75 w193 h90 altSubmit gSpeedvar vSpeed, Instantaneous|Fast|Normal|Slow|Custom
Gui, Add, Button, x10 y100 w279 gApply vApplyVar, Apply
Gui, Add, Button, X93 y10 w60 h20 gExcellC, Choose
Gui, Add, Button, X93 y33 w60 h20 gGHGC, Choose

Gui, Add, Edit, x133 y100 w80 vTabSpeed Hidden
Gui, Add, Edit, x133 y124 w80 vDelaySpeed Hidden
Gui, Add, Edit, x133 y148 w80 vSlowDelaySpeed Hidden
Gui, Add, Edit, x133 y172 w80 vBarangaySelectSpeed Hidden
Gui, Add, Edit, x133 y196 w80 vAltTabSpeed Hidden
Gui, Add, Edit, x133 y220 w80 vBarangayWaitSpeed Hidden

Gui, Add, Text, x8 y102 w120 Left vTabSpeed1 Hidden, Tab Speed
Gui, Add, Text, x8 y126 w120 Left vDelaySpeed1 Hidden, Delay Speed
Gui, Add, Text, x8 y150 w120 Left vSlowDelaySpeed1 Hidden, Slow Delay Speed
Gui, Add, Text, x8 y174 w120 Left vBarangaySelectSpeed1 Hidden, Barangay Select Speed
Gui, Add, Text, x8 y198 w120 Left vAltTabSpeed1 Hidden, Alt Tab Speed
Gui, Add, Text, x8 y222 w120 Left vBarangayWaitSpeed1 Hidden, Barangay Wait Speed

Gui, Add, Text, x216 y102 w120 Left vTabSpeed2 Hidden, (Default : 100)
Gui, Add, Text, x216 y126 w120 Left vDelaySpeed2 Hidden, (Default : 70)
Gui, Add, Text, x216 y150 w120 Left vSlowDelaySpeed2 Hidden, (Default : 200)
Gui, Add, Text, x216 y174 w120 Left vBarangaySelectSpeed2 Hidden, (Default : 300)
Gui, Add, Text, x216 y198 w120 Left vAltTabSpeed2 Hidden, (Default : 800)
Gui, Add, Text, x216 y222 w120 Left vBarangayWaitSpeed2 Hidden, (Default : 500)

if FileExist("C:\Users\Client\AppData\Local\Temp\Autofill.ini")
	{
	IniRead, Excell, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, ExcellWindow
	IniRead, GHG, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, GHGWindow
	IniRead, Speed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, Speed
	IniRead, SelectRTN, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, SelectionReturn
	;MsgBox, Excell = %Excell% GHG = %GHG% Speed = %Speed% SelectRTN = %SelectRTN%
	GuiControl,, ExcellWnd, %Excell%
	GuiControl,, GHGWnd, %GHG%
	ControlGet, checked, checked,, Button1, GHG Autofill
	If SelectRTN = 1
	Control, check,, Button1, GHG Autofill
	else
	Control, uncheck,, Button1, GHG Autofill
	If Speed = 1
	guicontrol, choose, Speed , Instantaneous
	Else If Speed = 2
	guicontrol, choose, Speed , Fast
	Else If Speed = 4
	guicontrol, choose, Speed , Slow
	Else If Speed = 5
	{
	guicontrol, choose, Speed , Custom
	IniRead, TabSpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, TabSpeed
	IniRead, DelaySpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, DelaySpeed
	IniRead, SlowDelaySpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, SlowDelaySpeed
	IniRead, BarangaySelectSpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, BarangaySelectSpeed
	IniRead, AltTabSpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, AltTabSpeed
	IniRead, BarangayWaitSpeed, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, BarangayWaitSpeed
	guicontrol,, TabSpeed, %TabSpeed%
	guicontrol,, DelaySpeed, %DelaySpeed%
	guicontrol,, SlowDelaySpeed, %SlowDelaySpeed%
	guicontrol,, BarangaySelectSpeed, %BarangaySelectSpeed%
	guicontrol,, AltTabSpeed, %AltTabSpeed%
	guicontrol,, BarangayWaitSpeed, %BarangayWaitSpeed%
	Gosub, Speedvar
	}
	Else
	{
	TabSpeed = ""
	DelaySpeed = ""
	SlowDelaySpeed = ""
	BarangaySelectSpeed = ""
	AltTabSpeed = ""
	BarangayWaitSpeed = ""
	guicontrol, choose, Speed , Normal
	}
	}
else
guicontrol, choose, Speed , Normal

Config = 0
;Menu, Tray, NoStandard
menu, tray, Add, Open, opengui
menu, tray, Default, Open
Menu, Tray, Add, Pause Hotkey, HotkeySuspend
Menu, Tray, Add, Exit, GuiClose
Return

HotkeySuspend:
Pause
return

Update:
version = ""
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("GET", "https://raw.githubusercontent.com/Dangerouze/Autofill/Dangerouzes-working-branch/Version.txt", true)
whr.Send()
; Using 'true' above and the call below allows the script to remain responsive.
whr.WaitForResponse()
version := whr.ResponseText
Msgbox, 8260, Update Checker, Would you like to update your software? Your current version: %AppVersion%. Latest Version: %version%.
IfMsgbox Yes
	{
	url = "https://github.com/Dangerouze/Autofill/Dangerouzes-working-branch/Autofill"%version%.ahk
	Filename = "Autofill"%version%.exe
	UrlDownloadToFile, *0 %url%, %A_WorkingDir%\%Filename%
	if ErrorLevel = 1
		MsgBox, There was some error updating the file. Please check you internet connection.
	else if ErrorLevel = 0
		{
		run	%A_WorkingDir%\%Filename%
		ExitApp
		}
	}
return
			
MsgBox % version
UrlDownloadToFile, *0 %url%, %A_WorkingDir%\%Filename%

GuiClose:
FileDelete, C:\Users\Client\AppData\Local\Temp\Autofill.ini
ExitApp

Speedvar:
Gui, Submit, Nohide
If (Speed == 5)
	{
	Gui, Show, w300 h275, GHG Autofill
	GuiControl, MoveDraw, ApplyVar, x10 y245
	GuiControl, Show, TabSpeed
	GuiControl, Show, DelaySpeed
	GuiControl, Show, SlowDelaySpeed
	GuiControl, Show, BarangaySelectSpeed
	GuiControl, Show, AltTabSpeed
	GuiControl, Show, BarangayWaitSpeed
	GuiControl, Show, TabSpeed1
	GuiControl, Show, DelaySpeed1
	GuiControl, Show, SlowDelaySpeed1
	GuiControl, Show, BarangaySelectSpeed1
	GuiControl, Show, AltTabSpeed1
	GuiControl, Show, BarangayWaitSpeed1
	GuiControl, Show, TabSpeed2
	GuiControl, Show, DelaySpeed2
	GuiControl, Show, SlowDelaySpeed2
	GuiControl, Show, BarangaySelectSpeed2
	GuiControl, Show, AltTabSpeed2
	GuiControl, Show, BarangayWaitSpeed2
	return
	}
Else
	{
	Gui, Show, w300 h130, GHG Autofill
	GuiControl, MoveDraw, ApplyVar, x10 y100
	GuiControl, Hide, TabSpeed
	GuiControl, Hide, DelaySpeed
	GuiControl, Hide, SlowDelaySpeed
	GuiControl, Hide, BarangaySelectSpeed
	GuiControl, Hide, AltTabSpeed
	GuiControl, Hide, BarangayWaitSpeed
	GuiControl, Hide, TabSpeed1
	GuiControl, Hide, DelaySpeed1
	GuiControl, Hide, SlowDelaySpeed1
	GuiControl, Hide, BarangaySelectSpeed1
	GuiControl, Hide, AltTabSpeed1
	GuiControl, Hide, BarangayWaitSpeed1
	GuiControl, Hide, TabSpeed2
	GuiControl, Hide, DelaySpeed2
	GuiControl, Hide, SlowDelaySpeed2
	GuiControl, Hide, BarangaySelectSpeed2
	GuiControl, Hide, AltTabSpeed2
	GuiControl, Hide, BarangayWaitSpeed2
	If (Speed == 1)
  {
  Msgbox, 8240, Caution, This option is highy unstable. Not recommended option for pc's with lag issues. If lag spike occurs, please double check immediately.
  }
	return
  }
ExcellC:
{
MsgBox, 8192, Note, Please Select excell windows and press "alt + a", 0
Config = 1
Gosub Activation
Gui, Hide
return
}


GHGC:
{
MsgBox, 8192, Note, Please Select GHG windows and press "alt + a", 0
Config = 2
Gosub Activation
Gui, Hide
return
}



Apply:
{
If (Excell != "") and If (GHG != "")
  {
  Gui, Submit, Nohide
   If (Speed == 1)
    {
    TabSpeed = 5
    DelaySpeed = 5
    SlowDelaySpeed = 100
    BarangaySelectSpeed = 150
    AltTabSpeed = 300
    BarangayWaitSpeed = 350
    }
   Else If (Speed == 2)
    {
    TabSpeed = 70
    DelaySpeed = 40
    SlowDelaySpeed = 140
    BarangaySelectSpeed = 210
    AltTabSpeed = 560
    BarangayWaitSpeed = 350
    }
    Else If (Speed == 3)
    {
    TabSpeed = 100
    DelaySpeed = 70
    SlowDelaySpeed = 200
    BarangaySelectSpeed = 300
    AltTabSpeed = 800
    BarangayWaitSpeed = 500
    }
    Else If (Speed == 4)
    {
    TabSpeed = 200
    DelaySpeed = 140
    SlowDelaySpeed = 400
    BarangaySelectSpeed = 600
    AltTabSpeed = 1600
    BarangayWaitSpeed = 1000
    }
    Else If (Speed == 5)
    {
    If (TabSpeed = "") 
     or (DelaySpeed = "")
     or (SlowDelaySpeed = "")
     or (BarangaySelectSpeed = "")
     or (AltTabSpeed = "")
     or (BarangayWaitSpeed = "")
      {
      MsgBox, 8192, Error, Please Fill All Blank Speeds, 0
      return
      }
     Else If !RegExMatch(TabSpeed,"^\d+$")
      or !RegExMatch(DelaySpeed,"^\d+$")
      or !RegExMatch(SlowDelaySpeed,"^\d+$")
      or !RegExMatch(BarangaySelectSpeed,"^\d+$")
      or !RegExMatch(AltTabSpeed,"^\d+$")
      or !RegExMatch(BarangayWaitSpeed,"^\d+$")
      {
      MsgBox, 8192, Error ,No Letters is Allowed on Speeds, 0
      return
      }
    Gui, Submit, Nohide
    }
    Else If (Speed = "")
      {
      MsgBox, 8192,Error, Please Select Speed, 0
      return
      }
  Gosub Activation
  Gui, Hide
  sleep 100
  Traytip, Configuration Saved. You can open the window by double clicking the icon or by pressing alt + C, 10, 0
  }
Else
  {
  MsgBox, Please Choose Excell and/or GHG Windows
  }
Return
}

Activation:
{
opengui:
!c::
Config = 0
Gui, Show
return

^Escape::
Gui, Submit, Nohide
IniWrite, %Excell%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, ExcellWindow
IniWrite, %GHG%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, GHGWindow
IniWrite, %Speed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, Speed
IniWrite, %SelectRTN%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, Configuration, SelectionReturn
IniWrite, %TabSpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, TabSpeed
IniWrite, %DelaySpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, DelaySpeed
IniWrite, %SlowDelaySpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, SlowDelaySpeed
IniWrite, %BarangaySelectSpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, BarangaySelectSpeed
IniWrite, %AltTabSpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, AltTabSpeed
IniWrite, %BarangayWaitSpeed%, C:\Users\Client\AppData\Local\Temp\Autofill.ini, CustomSpeedSettings, BarangayWaitSpeed
Reload
return

!a::
If (Config == 1)
  {
   WinGetTitle, TempWin , A
   Excell = %TempWin%
   Gui, Show
   GuiControl,, ExcellWnd, %TempWin%
   Config = 0
   return
  }
Else If (Config == 2)
  {
   WinGetTitle, TempWin , A
   GHG = %TempWin%
   Gui, Show
   GuiControl,, GHGWnd, %TempWin%
   Config = 0
   return
  }
If (Excell = "") or If (GHG = "")
  {
  MsgBox, 8192, Error, Please Choose Excell and/or GHG Windows, 0
  Gui, Show
  return
  }
 
WinGetTitle, TempWin , A
If (TempWin != Excell)
  {
  MsgBox, 8192, Error, This is Not Your Current Excell Window (You can open config by pressing "alt + c"), 0
  return
  }
  
  
Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline1:
Barangay = %Clipboard%
If (Barangay = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline1
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline1
  return
  }
Sleep, %DelaySpeed%
Clipboard :=

If InStr(Barangay, "AGCALAGA")
 or InStr(Barangay, "Aglibacao")
 or InStr(Barangay, "AGLONOK")
 or InStr(Barangay, "Alibunan")
 or InStr(Barangay, "Badlan Grande")
 or InStr(Barangay, "Badlan Pequeño")
 or InStr(Barangay, "BADU")
 or InStr(Barangay, "Baje San Julian")
 or InStr(Barangay, "Balaticon")
 or InStr(Barangay, "Banban Grande")
 or InStr(Barangay, "Banban Pequeño")
 or InStr(Barangay, "BINOLOSAN GRANDE")
 or InStr(Barangay, "BINOLOSAN PEQUEÑO")
 or InStr(Barangay, "Bo. Calinog ")
 or InStr(Barangay, "Cabagiao")
 or InStr(Barangay, "Cabugao")
 or InStr(Barangay, "CAHIGON")
 or InStr(Barangay, "Camalongo")
 or InStr(Barangay, "Canabajan")
 or InStr(Barangay, "Caratagan")
 or InStr(Barangay, "CARVASANA")
 or InStr(Barangay, "Dalid")
 or InStr(Barangay, "DATAGAN")
 or InStr(Barangay, "GAMA GRANDE")
 or InStr(Barangay, "Gama Pequeño")
 or InStr(Barangay, "GARANGAN")
 or InStr(Barangay, "Ginbunyugan")
 or InStr(Barangay, "GUISO")
 or InStr(Barangay, "Guiso")
 or InStr(Barangay, "HILWAN")
 or InStr(Barangay, "IMPALIDAN")
 or InStr(Barangay, "IPIL")
 or InStr(Barangay, "Jamin-ay")
 or InStr(Barangay, "LAMPAYA")
 or InStr(Barangay, "Libot")
 or InStr(Barangay, "LONOY")
 or InStr(Barangay, "Malag It")
 or InStr(Barangay, "Malaguinabot")
 or InStr(Barangay, "Malapawe")
 or InStr(Barangay, "Malitbog Centro")
 or InStr(Barangay, "Mambiranan")
 or InStr(Barangay, "Manaripay")
 or InStr(Barangay, "Marandig")
 or InStr(Barangay, "Masaroy")
 or InStr(Barangay, "Maspasan")
 or InStr(Barangay, "Nalbugan")
 or InStr(Barangay, "OWAK")
 or InStr(Barangay, "Pob. Centro")
 or InStr(Barangay, "POBLACION DELGADO")
 or InStr(Barangay, "POBLACION RIZAL ILAWOD")
 or InStr(Barangay, "POBLACION RIZAL ILAYA")
 or InStr(Barangay, "San Nicolas")
 or InStr(Barangay, "Simsiman")
 or InStr(Barangay, "SUPANGA")
 or InStr(Barangay, "Tabucan")
 or InStr(Barangay, "Tahing")
 or InStr(Barangay, "Tibiao")
 or InStr(Barangay, "Tigbayog")
 or InStr(Barangay, "Toyungan")
 or InStr(Barangay, "ULAYAN")
	{

Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline2:
HHN = %Clipboard%
If (HHN = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline2
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline2
  return
  }
Sleep, %DelaySpeed%
Clipboard :=
Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline3:
First = %Clipboard%
If (First = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline3
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline3
  return
  }
Sleep, %DelaySpeed%
Clipboard :=
Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline4:
Middle = %Clipboard%
If (Middle = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline4
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline4
  return
  }
Sleep, %DelaySpeed%
Clipboard :=
Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline5:
Last = %Clipboard%
If (Last = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline5
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline5
  return
  }
Sleep, %DelaySpeed%
Clipboard :=
Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline6:
Extension = %Clipboard%
If (Extension = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline6
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline6
  return
  }
Sleep, %DelaySpeed%
Clipboard :=
Send, {Tab}
Sleep, %TabSpeed%

Send, ^c
ClipWait, 2
Sleep, %DelaySpeed%
Loopline7:
Purok = %Clipboard%
If (Purok = "")
  {
  WinGetTitle, TempWin , A
  If (TempWin = "Microsoft Excel")
    {
    Sleep, %DelaySpeed%
    Send, {Enter}
    Gosub, Loopline7
    return
    }
  Sleep, %DelaySpeed%
  Send, ^c
  ClipWait, 2
  Sleep, %DelaySpeed%
  Gosub, Loopline7
  return
  }
Sleep, %DelaySpeed%
Clipboard :=

If (SelectRTN = 1)
  { 
  Send, {Left}
  Send, {Left}
  Send, {Left}
  Send, {Left}
  Send, {Left}
  Send, {Left}
  Sleep, %SlowDelaySpeed%
  Sleep, %SlowDelaySpeed%
  }
Send, !{Tab}
Sleep, %AltTabSpeed%
WinGetTitle, TempWin , A
If (TempWin != GHG)
	{
	MsgBox, 8192, Error, This is Not Your Current GHG Window (You can open config by pressing "alt + c"), 0
	return
	}

Click, 60, 130
Sleep %SlowDelaySpeed%
Send, {Tab}
Sleep, %SlowDelaySpeed%

If InStr(Barangay, "BINOLOSAN GRANDE")
	{
	Send, BINOLOSAN
	}
else if InStr(Barangay, "BINOLOSAN PEQUEÑO")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, BINOLOSAN
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else If InStr(Barangay, "Badlan Grande")
	{
	Send, BADLAN
	}
else if InStr(Barangay, "Badlan Pequeño")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, BADLAN
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else If InStr(Barangay, "Banban Grande")
	{
	Send, BANBAN
	}
else if InStr(Barangay, "Banban Pequeño")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, BANBAN
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else If InStr(Barangay, "GAMA GRANDE")
	{
	Send, GAMA
	}
else if InStr(Barangay, "Gama Pequeño")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, GAMA
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else If InStr(Barangay, "POBLACION DELGADO")
	{
	Send, POBLACION
	}
else if InStr(Barangay, "POBLACION RIZAL ILAYA")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, POBLACION
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else if InStr(Barangay, "POBLACION RIZAL ILAWOD")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, POBLACION
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else if InStr(Barangay, "Pob. Centro")
	{
	Send, {Enter}
	Sleep, %BarangaySelectSpeed%
	Send, POBLACION
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Down}
	Sleep, %SlowDelaySpeed%
	Send, {Enter}
	Sleep, %SlowDelaySpeed%
	}
else If InStr(Barangay, "Bo. Calinog ")
	{
	Send, BARRIO
	}
else If InStr(Barangay, "Malitbog Centro")
	{
	Send, Malitbog
	}
else
{
Send, %Barangay%
}

Sleep, %BarangayWaitSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %HHN%

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %First%

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %Middle%

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %Last%

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %Extension%

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, {Space}

Sleep, %TabSpeed%
Send, {Tab}
Sleep, %TabSpeed%
Send, %Purok%
Sleep, %TabSpeed%
Send, {Tab}
Gosub Closing
return
}
else
{
Msgbox, Invalid Barangay
Gosub Closing
return
}
}

Closing:
{
; MsgBox, HHN=%HHN% Barangay=%Barangay% First=%First% Middle=%Middle% Last=%Last% Extension=%Extension% Purok=%Purok%
HNN = ""
Barangay = ""
First = ""
Middle = ""
Last = ""
Extension = ""
Purok = ""
}