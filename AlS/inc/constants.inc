<SCRIPT language="VBScript" RUNAT="SERVER">

  Const NumWeeksPerYear = 52
  Const NumMonthPerYear = 12
  Const NumOfAdminMenu = 11

  Const DAYCOLOR		= "#FF9999"
  Const SATCOLOR		= "#E7EBF5"
  Const SUNCOLOR		= "#C2CCE7"
  Const HOLIDAYCOLOR	= "#FFC6C6"
  Const MAXHOUR		= 24
  Const PageSizeDefault	= 100
  Const preferences	= "<a href='XXtools/preferences.asp' class='c' onMouseOver='self.status=&quot;Preferences&quot; ; return true;' onMouseout='self.status=&quot;&quot;'>Preferences</a>"
	Const logoff = "<a href='XXlogout.asp' class='c' onMouseOver='self.status=&quot;Log out&quot; ; return true;' onMouseout='self.status=&quot;&quot;'>Log out</a>"
	Const help	= "<a href='javascript: void(0);' class='c' onMouseOver='self.status=&quot;Help&quot; ; return true;' onMouseout='self.status=&quot;&quot;'>Help</a>"

	
	Const SMTPsendusing=2
	Const SMTPserver ="smtp.gmail.com"
	Const SMTPserverport = 25
	Const SMTPusessl = 1
	Const SMTPconnectiontimeout = 60
	Const SMTPauthenticate = 1
	'Const SMTPsendusername ="no-reply@atlasindustries.com"	
	'Const SMTPsendpassword ="time7*sheed"
	Const SMTPsendusername ="atlas.ais.noreply@gmail.com"	
	Const SMTPsendpassword ="m!234567"

	Const webname="Atlas Information System"
</SCRIPT>