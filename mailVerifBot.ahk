;Name:        Verify E-mail adress macro             
;Author:      Thomas Specq
;Created:     24/08/2017
;Copyright:   (c) Thomas Specq 2017 
;Licence:     BSD
;Website:     http://www.thomas-specq.work
;Link:        <a href="http://www.thomas-specq.work">Freelance Web Design & Développement</a>

global URL1 := 1, URL2 := 2, URL3 := 3, URL4 := 4, URL5 := 5, Result1 := 1, Result2 := 2, Result3 := 3, Result4 := 4, Result5 := 5,

WaitTillNetworkPage5IsLoaded()
{ 
Loop {
 ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\Albert\Desktop\ahkImage\networkIcon5.bmp
 if ErrorLevel = 2 
 {
 MsgBox Could not conduct the search.
 } 
 if ErrorLevel = 1 
 {
 Sleep, 500
 
 }
 if ErrorLevel = 0
 {

 break
 
 }
 } 
 } 
 
 WaitTillNetworkPage4IsLoaded()
{ 
Loop {
 ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\Albert\Desktop\ahkImage\networkIcon4.bmp
 if ErrorLevel = 2 
 {
 MsgBox Could not conduct the search.
 } 
 if ErrorLevel = 1 
 {
 Sleep, 500
 
 }
 if ErrorLevel = 0
 {
 
 break
 
 }
 } 
 }
 
WaitTillNetworkPage3IsLoaded()
{ 
Loop {
 ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\Albert\Desktop\ahkImage\networkIcon3.bmp
 if ErrorLevel = 2 
 {
 MsgBox Could not conduct the search.
 } 
 if ErrorLevel = 1 
 {
 Sleep, 500

 }
 if ErrorLevel = 0
 {

 break
 
 }
 }  
 }
 
WaitTillNetworkPage2IsLoaded()
{ 
Loop {
 ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\Albert\Desktop\ahkImage\networkIcon2.bmp
 if ErrorLevel = 2 
 {
 MsgBox Could not conduct the search.
 } 
 if ErrorLevel = 1 
 {
 Sleep, 500
 
 }
 if ErrorLevel = 0
 {
 
 break

 }
 } 
 }
 
 WaitTillNetworkPage1IsLoaded()
{ 
Loop {
 ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, C:\Users\Albert\Desktop\ahkImage\networkIcon1.bmp
 if ErrorLevel = 2 
 {
 MsgBox Could not conduct the search.
 } 
 if ErrorLevel = 1 
 {
 Sleep, 500
 
 }
 if ErrorLevel = 0
 {

 break
 
 }
 } 
 }

shiftToChrome()
{
MouseMove, 707,  899
Sleep, 500
MouseClick, left,  717,  880
Sleep, 500
}

shiftToExcel()
{
MouseMove, 707,  899
Sleep, 500
MouseClick, left, 462,  884
Sleep, 500
}

recordURL(){
; click Excel
MouseClick, left,  348,  18

Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
URL1 = %clipboard%
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
URL2 = %clipboard%
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
URL3 = %clipboard%
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
URL4 = %clipboard%
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
URL5 = %clipboard%
}

openTabAndPasteLink()
{
Send, %URL1%
Sleep, 500
Send, {CTRLDOWN}t{CTRLUP}
Sleep,500
Send, %URL2%
Sleep, 500
Send, {CTRLDOWN}t{CTRLUP}
Sleep,500
Send, %URL3%
Sleep, 500
Send, {CTRLDOWN}t{CTRLUP}
Sleep,500
Send, %URL4%
Sleep, 500
Send, {CTRLDOWN}t{CTRLUP}
Sleep,500
Send, %URL5%
Sleep, 500


}


shiftTo1stTab(){
Send, {CTRLDOWN}{TAB}{CTRLUP}
}

copyResultAndCloseTab(){

MouseClick, left,  450,  252
MouseClick, left,  450,  252
MouseClick, left,  450,  252
Send, {CTRLDOWN}c{CTRLUP}
Result1 = %clipboard%
Sleep,500
Send, {CTRLDOWN}w{CTRLUP}

MouseClick, left,  450,  252
MouseClick, left,  450,  252
MouseClick, left,  450,  252
Send, {CTRLDOWN}c{CTRLUP}
Result2 = %clipboard%
Sleep,500
Send, {CTRLDOWN}w{CTRLUP}

MouseClick, left,  450,  252
MouseClick, left,  450,  252
MouseClick, left,  450,  252
Send, {CTRLDOWN}c{CTRLUP}
Result3 = %clipboard%
Sleep,500
Send, {CTRLDOWN}w{CTRLUP}

MouseClick, left,  450,  252
MouseClick, left,  450,  252
MouseClick, left,  450,  252
Send, {CTRLDOWN}c{CTRLUP}
Result4 = %clipboard%
Sleep,500
Send, {CTRLDOWN}w{CTRLUP}

MouseClick, left,  637,  248
MouseClick, left,  637,  248
MouseClick, left,  637,  248
Send, {CTRLDOWN}c{CTRLUP}
Result5 = %clipboard%
Sleep,500
Send, {CTRLDOWN}t{CTRLUP}
shiftTo1stTab()
Send, {CTRLDOWN}w{CTRLUP}
}

writeToExcel(){
clipboard = %Result1% 
Send, {RIGHT}
Send, {CTRLDOWN}v{CTRLUP}
clipboard = %Result2% 
Send, {RIGHT}
Send, {CTRLDOWN}v{CTRLUP}
clipboard = %Result3%
Send, {RIGHT}
Send, {CTRLDOWN}v{CTRLUP}
clipboard = %Result4%
Send, {RIGHT}
Send, {CTRLDOWN}v{CTRLUP}
clipboard = %Result5%
Send, {RIGHT}
Send, {CTRLDOWN}v{CTRLUP}
}

1lineDown(){
MouseClick, left,  1373,  846
Sleep, 100
}

getBackToOriginalPosition(){
Send, {DOWN}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}{LEFT}
}

;main command
Loop, 50 {
recordURL()
shiftToChrome()
openTabAndPasteLink()

WaitTillNetworkPage1IsLoaded()
WaitTillNetworkPage2IsLoaded()
WaitTillNetworkPage3IsLoaded()
WaitTillNetworkPage4IsLoaded()
WaitTillNetworkPage5IsLoaded()
shiftTo1stTab()
copyResultAndCloseTab()
shiftToExcel()
writeToExcel()
1lineDown()
getBackToOriginalPosition()
}

Esc::ExitApp
