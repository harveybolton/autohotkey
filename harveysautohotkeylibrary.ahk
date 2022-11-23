;    _____      __            
;   / ___/___  / /___  ______ 
;   \__ \/ _ \/ __/ / / / __ \
;  ___/ /  __/ /_/ /_/ / /_/ /
; /____/\___/\__/\__,_/ .___/ 
;                    /_/      
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;    __  __          ____      __   _____ __               __             __      
;   / / / /_______  / __/_  __/ /  / ___// /_  ____  _____/ /________  __/ /______
;  / / / / ___/ _ \/ /_/ / / / /   \__ \/ __ \/ __ \/ ___/ __/ ___/ / / / __/ ___/
; / /_/ (__  )  __/ __/ /_/ / /   ___/ / / / / /_/ / /  / /_/ /__/ /_/ / /_(__  ) 
; \____/____/\___/_/  \__,_/_/   /____/_/ /_/\____/_/   \__/\___/\__,_/\__/____/                                                                             
																				
; Shortcut to type email addresses
::hmail::harvey@providentitsolutions.co.uk
::tmail::tom@providentitsolutions.co.uk
::bmail::ben@providentitsolutions.co.uk
::dmail::dan@providentitsolutions.co.uk
::smail::support@providentitsolutions.co.uk

; Create new Autotask Ticket with Ctrl + Alt + T
^!t::
InputBox, UserInput, New Ticket, Enter customer name., ,
if ErrorLevel
    sleep 100
else
    Run, https://ww16.autotask.net/Autotask/AutotaskExtend/AutotaskCommand.aspx?Code=NewTicket&AccountName=%UserInput%
Return

; Ping/Resolve Hostname with ALT+R
!R::
sleep 50
Send, ^C
sleep 50
Run, C:\Windows\System32\cmd.exe
sleep 50
Send, ping %clipboard%
Send, {Enter}
Return

; Perform WHOIS lookup with ALT+W
!W::
Send, ^C
sleep 50
Run, https://who.is/whois/%clipboard%
Return

; Search Google for any highlighted text with CTRL+SHIFT+C
^+c::
Send, ^c
Sleep 50
Run, http://www.google.com/search?q=%clipboard%
Return

; Type the clipboard text by pressing CTRL+ALT+V, useful when clipboard syncing isn't working on Splashtop or a field can't be pasted
^!v::
sleep 2000
send,%clipboard%
return

; Set the currently active window to always on top with CTRL+SPACE
 ^SPACE:: Winset, Alwaysontop, , A
 return

; Automate answering "Hope you're well" by pressing WIN+F10
#F10::
send, Not too bad over here thank you, hopefully the same for yourselves!
return

;     ___         __           __                _     
;    /   | __  __/ /_____     / /   ____  ____ _(_)___ 
;   / /| |/ / / / __/ __ \   / /   / __ \/ __ `/ / __ \
;  / ___ / /_/ / /_/ /_/ /  / /___/ /_/ / /_/ / / / / /
; /_/  |_\__,_/\__/\____/  /_____/\____/\__, /_/_/ /_/ 
;                                      /____/          
									 
; Login to everything with CTRL+ALT+L
^!l::
dattologin()
sleep 500
autotasklogin()
sleep 500
itgluelogin()
Return


dattologin() {
SetTitleMatchMode, 2 ; match any part of the title
Send ^{t}
sleep 1000
Send https://merlot.centrastage.net/
Send {Enter}
sleep 3500
If WinActive("Log In") {
	Send !{PgDn}
	sleep 1000
	Send {Enter}
	sleep 2000
	Send {tab}
	Send {Enter}
	sleep 2000
	If WinActive("Two-Factor Login") {
		sleep 100
		Send %Clipboard%
		Send {Enter}
	}
	If WinActive("Choose") {
		Send {tab}
		Send {Enter}
	}
}
Return
}


autotasklogin() {
SetTitleMatchMode, 2 ; match start of the title
Send ^{t}
sleep 1000
Send https://ww16.autotask.net/Mvc/Framework/Navigation.mvc/Landing
Send, {Enter}
sleep 3000
If WinActive("Autotask") {
	Send !{PgDn}
	sleep 500	
	Send {Enter}
	sleep 1500
	Send !{PgDn}
	sleep 500	
	Send {Enter}
	Sleep 1500
	Sleep 150
	Send %Clipboard%
	sleep 150
	Send {Enter}
}
Return
}


itgluelogin() {
SetTitleMatchMode, 1 ; match start of the title
Send ^{t}
sleep 1000
Send https://provident-it.eu.itglue.com/
Send, {Enter}
sleep 4000
If WinActive("Dashboard") {
Return
}
If WinActive("IT Glue") {
	Send !{PgDn}
	sleep 500
	Send {Enter}
	sleep 1000
	Send %Clipboard%
	Sleep 500
	Send, {Enter}
	return
}
}

; Automatic Datto Login with CTRL+ALT+SHIFT+4
^!+4::
dattologin()
Return


; Automatic Autotask Login with CTRL+ALT+SHIFT+5
^!+5::
autotasklogin()
Return

; Automatic IT Glue login with CTRL+ALT+I
^!i::
itgluelogin()
Return

;     __  __                        ___              _      __              __ 
;    / / / /___  ____ ___  ___     /   |  __________(_)____/ /_____ _____  / /_
;   / /_/ / __ \/ __ `__ \/ _ \   / /| | / ___/ ___/ / ___/ __/ __ `/ __ \/ __/
;  / __  / /_/ / / / / / /  __/  / ___ |(__  |__  ) (__  ) /_/ /_/ / / / / /_  
; /_/ /_/\____/_/ /_/ /_/\___/  /_/  |_/____/____/_/____/\__/\__,_/_/ /_/\__/  
																			 
; Open Blinds with CTRL+ALT+UP
^!+Up::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/open-blinds")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Return

; Close Blinds with CTRL+ALT+DOWN
^!+Down::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/close-blinds")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Return

; Stop Blinds with CTRL+ALT+RIGHT
^!+Right::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/stop-blinds")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Run, https://who.is/w
Return

; Lights Red 20% with CTRL+SHIFT+1
^!+1::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/turn_lights_red")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Return

; Lights Off with CTRL+SHIFT+2
^!+2::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/turn_lights_off")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Return

; Lights White 100% with CTRL+SHIFT+3
^!+3::
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("POST", "https://homeassistant.bolton.cloud/api/webhook/turn_lights_white")
Body := "original_url=http://autohotkey.com&response_type=json"
whr.Send(Body)
Return

;     __  ____          
;    /  |/  (_)_________
;   / /|_/ / / ___/ ___/
;  / /  / / (__  ) /__  
; /_/  /_/_/____/\___/  
                      
; Shutdown on Ctrl + F1
^F1::
Shutdown, 5


; Rapid-fire clicking with Ctrl + 8
^8::
Toggle := !Toggle
Loop
{
	If (!Toggle)
		return
	Click
	Sleep 50 ; Make this number higher for slower clicks, lower for faster.
}

; Rapid-fire message typing with Ctrl + 9
^9::
Toggle := !Toggle
Loop
{
	If (!Toggle)
		return
	Send, Hello
	Send, {Enter}
	sleep 30
}

; Get current cursor co-ordinates by pressing CTRL+Q, for when developing AHK scripts.
^q::
MouseGetPos, xpos, ypos 
MsgBox, The cursor is at X%xpos% Y%ypos%.
return

; Reload the AutoHotKey script with Ctrl + Esc to quickly load changes without killing in Task Manager
^Esc::Reload
