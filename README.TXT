
MAXIMIZE FIRST TO GET THE PROPER VIEW !!

'*******************************
'*                             *
'*     	   Active Line         *
'*       Copyright ©2001       *
'*      Austin K. Hayward      *
'*                             *
'*******************************

Option Explicit


Dim strCompanyInformation As String
Dim strVersionNumber As String
Dim strMiscInformation As String
Dim strEmail As String


Private Sub ReadMe_Load()

    strCompanyInformation = "Copyright ©2001 Austin Hayward"

    strVersionNumber = "1.0.0"

    strEmail = "achayward@frontlinefocus.com"

    strMiscInformation = "This product is supplied as is.  Any changes or deletions due to " & _
			 "yourself seeing room for improvement should be emailed to Austin Hayward " & _
			 "so as to better improve the software for future releases." & vbcrlf & _
			 "" & _
			 "You may use this control freely in your projects, but please do not distribute it " & _
			 "without the expressed written consent of Austin Hayward." & vbcrlf & _
			 "" & _
			 "Any comments or suggestions are always welcome.  Please email me at: " & strEmail

End Sub