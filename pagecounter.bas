Attribute VB_Name = "Module1"
Sub CGI_Main()
Dim cValue, Pages
SendHeader "ben's script" ' duno what this does, but the script wont work without it
Pages = GetCgiValue("page") ' if you did  cgi-bin/blah.exe?page=hello    then Pages would equal "hello"
If Pages = "" Then  '  if no page is specified  then ...
    Send "Error: No page specified." ' send an error
    End ' then end the program
End If
P$ = Pages ' i put the contents of pages in P$ cause getfromini likes it when its a string
cValue = GetFromINI("pages", P$, "c:\windows\pagescript.ini") ' gets the count for Pages
If cValue = "" Then Let cValue = 0  ' if the pages counter value doesnt exist, then give it a value
cValue = cValue + 1  ' add 1 to the pages value
P$ = Pages ' writetoini likes strings also
c$ = cValue ' read ^
WriteToINI "pages", P$, c$, "c:\windows\pagescript.ini" ' save the new counter value on the hard drive
Send c$ ' send the new counter value back to the server
End ' ends the program
End Sub
