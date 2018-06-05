#SingleInstance force

!+t::

send {HOME}
    sleep, 270
send ^+{RIGHT}
    sleep, 270
send +{LEFT}
    sleep, 270
send ^c
clipwait
var1 := clipboard
clipboard := ClipboardRaw
    sleep, 270
send {RIGHT 2}
    sleep, 270
send ^+{RIGHT}
    sleep, 270
; send +{LEFT}
;     sleep, 270
send ^c
clipwait
var2 := clipboard

WinActivate, Oracle Fusion Middleware Forms Services:  Open > SOAPCOL - SHATRNS - SHATAEQ
    sleep, 270

IfInString, var1, Spr
    {
        Send +{TAB 3}
            sleep, 270
        Send {BACKSPACE}
            sleep, 270
        SendRaw % var2
        send 20
            sleep, 270
        SendInput {TAB}
            sleep, 270
        SendInput {TAB}
            sleep, 270
        WinActivate, ahk_exe WINWORD.exe 
            sleep, 270
    }

else IfInString, var1, Summer 
    {
        Send +{TAB 3}
            sleep, 270
        Send {BACKSPACE}
            sleep, 270
        SendRaw % var2
        sendRaw 30
            sleep, 270
        SendInput {TAB}
            sleep, 270
        SendInput {TAB}
            sleep, 270
        WinActivate, ahk_exe WINWORD.exe 
            sleep, 270
    }

else
    {
        var2 += 1
            sleep, 270
        Send +{TAB 3}
            sleep, 270
        Send {BACKSPACE}
            sleep, 270
        SendRaw % var2
        send 10
            sleep, 270
        Send {TAB}
            sleep, 350
        Send {TAB}
            sleep, 350
        WinActivate, ahk_exe WINWORD.exe 
            sleep, 270
    }

return
