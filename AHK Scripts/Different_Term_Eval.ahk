#d::

send {HOME}
    sleep, 100
send ^+{RIGHT}
    sleep, 100
send +{LEFT}
    sleep, 100
send ^c
clipwait
var1 := clipboard
    sleep, 100
send {RIGHT 2}
    sleep, 100
send ^+{RIGHT}
    sleep, 100
send +{LEFT}
    sleep, 100
send ^c
clipwait
var2 := clipboard
    sleep, 100
send {END}
    sleep, 100
send ^+{LEFT}
    sleep, 100
send ^c
clipwait
    sleep, 400

WinActivate, Oracle Fusion Middleware Forms Services:  Open > SOAPCOL - SHATRNS - SHATAEQ 

sendRaw % var1
    sleep, 100
send {TAB}
    sleep, 100
sendRaw % var2
    sleep, 100
send {TAB 2}
    sleep, 100
sendInput ^v
    sleep, 100
 send {DOWN}
     sleep, 100
 send {TAB}
     sleep, 100
 send {BACKSPACE}
     sleep, 100

WinActivate, beta_123 - Word

send {DOWN 2}
    sleep, 100
