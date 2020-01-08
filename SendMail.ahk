;Name:        Sending Mail Macro             
;Author:      Thomas Specq
;Created:     24/08/2017
;Copyright:   (c) Thomas Specq 2017 
;Licence:     BSD
;Website:     http://www.thomas-specq.work
;Link:        <a href="http://www.thomas-specq.work">Freelance Web Design & Développement</a>

global Mail := 1, global NameSurname := 2,

shiftToExcel()
{
MouseMove, 707,  899
Sleep, 500
MouseClick, left, 462,  884
Sleep, 500
}

shiftToChrome()
{
MouseMove, 707,  899
Sleep, 500
MouseClick, left,  717,  880
Sleep, 500
}

recordMail()
{
; click Excel
MouseClick, left,  348,  18
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
Mail = %clipboard%
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
NameSurname = %clipboard%
}

writeMail()
{
Send, %Mail%
Sleep, 500
MouseClick, left,  868,  469
Sleep, 1000
Send, Quick IT question to %NameSurname%
Sleep, 500
MouseClick, left,  800,  472
Sleep, 500
Send, Dear %NameSurname%
Send, May I ask you a simple question
Send, {ENTER}
Send, Does your company have any software development needs? 
Send, {ENTER}
Send, If yes, we have many offers for you such as the services of a Software Engineer Full-Time for a very competitive price.
Send, {ENTER}
Send, {ENTER}
Send, Best regards
Send, {ENTER}
Send, Thomas 
Send, {ENTER}
Send, Sales@LogimineAsia
Sleep, 1000
Send, {CTRLDOWN}{ENTER}
Send, {CTRLUP}
}

writeDone(){
Send, {RIGHT}
Send, done
}

shiftDown1Line(){
MouseClick, left,  1376,  845
Sleep, 100
Send, {DOWN}{LEFT}{LEFT}
}

;main
Loop, 100
{
Sleep, 1000
recordMail()
shiftToChrome()
Send, c
Sleep, 1000
writeMail()
shiftToExcel()
writeDone()
shiftDown1Line()
}

Esc::ExitApp
