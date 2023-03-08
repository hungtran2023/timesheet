
function saveCookie(name,value,days) 
{
	if (days) 
	{
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000))
		var expires = "; expires=" + date.toGMTString()
	}
	else expires = ""
	document.cookie = name + "=" + value + expires + ";path=/"
}

function readCookie(name) 
{
	var nameEQ = name + "="
	var ca = document.cookie.split(';')
	for(var i=0;i<ca.length;i++) 
	{
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length)
		if (c.indexOf(nameEQ) == 0) 
		{
			return c.substring(nameEQ.length,c.length)
		}
	}
	return null
}

// change irow and icolumn color
function crosshair(iRow,iCol,sColor)
{
  var r, c;
  var obj;

  if ( iRow && iCol ) // Crosshair
  {
    for ( r = 0; r <= iRow; r++ )
    {
      id = 'RC' + r + '_' + iCol;
      obj = document.all( id );
      if (obj)
        obj.bgColor = sColor;
    }

    for ( c = 0; c <= iCol; c++ )
    {
      id = 'RC' + iRow + '_' + c;
      obj = document.all( id );
      if (obj)
        obj.bgColor = sColor;
    }
  }
  else if ( iRow == 0 ) // Doing the whole column
  {
    r = 0;
    id = 'RC' + r + '_' + iCol;
    obj = document.all( id );
    while (obj)
    {
      obj.bgColor = sColor;
      r++;
      id = 'RC' + r + '_' + iCol;
      obj = document.all( id );
    }
  }
  else if ( iCol == 0 ) // Doing the whole row
  {
    c = 0;
    id = 'RC' + iRow + '_' + c;
    obj = document.all( id );
    while (obj)
    {
      obj.bgColor = sColor;
      c++;
      id = 'RC' + iRow + '_' + c;
      obj = document.all( id );
    }
  }
}
function deleteCookie(name) 
{
	saveCookie(name,'',-1)
}

function isempty(str) 
{
	atmp = str.split(String.fromCharCode(32));
	for (ii=0; ii< atmp.length;ii++ ) 
	{
		if (atmp[ii] != "") 
		return false;
	}	
	return true;
}

function isNumEnter(vobject) 
{
	if (isempty(vobject.value) == true) 
	{
		vobject.focus();				
		return false;
	}
	else
		if (isNaN(vobject.value))	
		{
				alert("Please enter the Number !");	
				vobject.focus();				
			return false;
		}		
	return true;	
}

function replace(expression,find,replacement) 
{	
	if (expression.length>0) 
	{
		atmp = expression.split(find)	
		stmp ="";
		for (ii=0;ii<atmp.length;ii++) 
		{
			stmp += atmp[ii];
			if (ii < atmp.length - 1)
				stmp += replacement;
		}	
		return stmp
	}
	else
		return "";
}

function rs(n,u,w,h,l,t,s) 
{	
  /* format : n = name
			  u : url
			  w : width
			  h : height
			  l : position of window left
  			  t : position of window top 
  			  s : scrollbar  = yes or no
	Example :  rs("wtest","appraisal/180/begintest.asp",500,500,0,0,"yes")
  */  		  
  	
  args="width="+w+",height="+h+",top="+t+",left="+l+",resizable=yes,scrollbars=" + s + ",status=0,toolbar=no,menubar=no,location=no";
  remote=window.open(u,n,args);
  if (remote != null) 
  {
    if (remote.opener == null)
      remote.opener = self;
  }
}

function comparedate(sbeg,send) 
{	
	abeg = sbeg.split("/")
	aend = send.split("/")	
	
	dbeg = abeg[0]; 
	if (abeg[0].substr(0,1) == "0")
		dbeg = abeg[0].substr(1,1);
	dend = aend[0]; 
	if (aend[0].substr(0,1) == "0")
		dend = aend[0].substr(1,1);
	
	mbeg = abeg[1]; 
	if (abeg[1].substr(0,1) == "0")
		mbeg = abeg[1].substr(1,1);
	mend = aend[1]; 
	if (aend[1].substr(0,1) == "0")
		mend = aend[1].substr(1,1);
		
	if (parseInt(abeg[2]) > parseInt(aend[2]))
		return false
	else 
	{
		if (parseInt(abeg[2]) == parseInt(aend[2]))
			if (parseInt(mbeg) > parseInt(mend)) 
					return false	
		else 
		{
			if  (parseInt(mbeg) == parseInt(mend))
					if (parseInt(dbeg) > parseInt(dend) & parseInt(mbeg) >= parseInt(mend))
						return false
		}
	}	
			
	return true;							
}

function isdate(s) 
{
	if (s =="") return false;
	var atmp,ii;
	s = replace(s,' ', '')
	atmp = s.split("/")	
	if (atmp.length != 3) return false;	
	for (ii=0; ii<atmp.length ;ii++) 
	{				
		if (isNaN(atmp[ii])) return false;
		else 
		{			
			switch (ii) 
			{				
				case 0,1: //month					
					if (atmp[ii].length <= 0 || atmp[ii].length > 2) 
					{					
						return false;
					}
					switch (parseInt(atmp[1])) 
					{
						case 4:
						case 6:
						case 9:
						case 11:
							if (parseInt(atmp[0]) > 30) 														
							return false;
							break;							
						case 2:													
							if (parseInt(atmp[2])%4 == 0 ) 
								{if (parseInt(atmp[0]) > 29) 
								
								return false;}
							else
								{if (parseInt(atmp[0]) > 28) 																									
								return false;}
							break;			
						default:
							if (atmp[1] <=0 || atmp[1]>12) 										
							return false;
							if (parseInt(atmp[0]) > 31) 							
							return false;
					}					
					break;				
				case 2: //year
					if (atmp[ii].length != 4) 
					{					
						return false;
					}
			}						
		}		
	}
	return true;
}

function isnull(str) 
{	
	atmp = str.split(String.fromCharCode(32));
	for (ii=0; ii< atmp.length;ii++ ) 
	{
		if (atmp[ii] != "") 
		{
			return false;
		}
	}	
	return true;
}

//function isemail(value) 
//{	
//	var pos1,pos2,pos3
//	pos1=value.indexOf("@");	
//	pos2=value.indexOf(" ");
//	pos3=value.lastIndexOf(".");	
//	if ( (pos1 == -1) || (pos2!= -1) || (pos3 == -1) || ( pos3<pos1) ) 
//	{
//		alert("Invalid value email address \nValid format is: 'NickName@domain.com'");
//		return false;
//	}
//	return true;
//}

function isEmpty_(s){
	return((s==null)||(s.length==0))
}

function isWhitespace(s){
	var i;
	if(isEmpty_(s))
		return true;
	for(i=0;i<s.length;i++)
		{   
		var c=s.charAt(i);
		if(" \t\n\r".indexOf(c)==-1)
			return false;
		}
	return true;
}

function isemail(s){
	if(isEmpty_(s)) 
		if(isemail.arguments.length == 1)
			return false;
		else
			return(isemail.arguments[1]==true);
	if(isWhitespace(s))
		return false;
	var i=1;
	var sLength=s.length;
	while((i<sLength)&&(s.charAt(i)!="@"))
		 i++;
	if((i>=sLength)||(s.charAt(i)!="@"))
		return false;
	else
		i+=2;
	while((i<sLength)&&(s.charAt(i)!="."))
		i++;
	if((i>=sLength-1)||(s.charAt(i)!="."))
		return false;
	else
		return true;
}

function emailCheck (emailStr) {
	var emailPat=/^(.+)@(.+)$/
	var specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]"
	var validChars="\[^\\s" + specialChars + "\]"
	var quotedUser="(\"[^\"]*\")"
	var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/
	var atom=validChars + '+'
	var word="(" + atom + "|" + quotedUser + ")"
	var userPat=new RegExp("^" + word + "(\\." + word + ")*$")
	var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$")
	var matchArray=emailStr.match(emailPat)
	if (matchArray==null) {
		alert("This address seems to be incorrect (they usually have an @ symbol)")
		return false
	}
	var user=matchArray[1]
	var domain=matchArray[2]
	if (user.match(userPat)==null) {
	    alert("The username doesn't seem to be valid.")
	    return false
	}
	var IPArray=domain.match(ipDomainPat)
	if (IPArray!=null) {
	  for (var i=1;i<=4;i++) {
	    if (IPArray[i]>255) {
	        alert("Destination IP address is invalid!")
			return false
	    }
	    }
	    return true
	}
	var domainArray=domain.match(domainPat)
	if (domainArray==null) {
		alert("The domain name doesn't seem to be valid.")
	    return false
	}
	var atomPat=new RegExp(atom,"g")
	var domArr=domain.match(atomPat)
	var len=domArr.length
	if (domArr[domArr.length-1].length<2 ||
	    domArr[domArr.length-1].length>3) {
	   alert("The address must end in a three-letter domain, or two letter country.")
	   return false
	}
	if (len<2) {
	   var errStr="This address is missing a hostname!"
	   alert(errStr)
	   return false
	}
	return true;
}

function iswebsite(strWebsite) {
  if (strWebsite !=  "") {
    if (strWebsite.substring(0, 11) == "http://www." && strWebsite.indexOf(".", 12) != -1) 
      return true;
    else { 
      if (strWebsite.substring(0, 7) == "http://" && strWebsite.indexOf(".", 7) != -1) 
        return true;
      else {
		if (strWebsite.substring(0, 4) == "www." && strWebsite.indexOf(".", 4) != -1) {
			return true;
			}
		else
			return false;
		}
	}
  }
}

function isnumber(value,name) 
{
	if (isNaN(value) ==  true) 
	{
		alert("Invalid value '" + name.toUpperCase() + "' field!"); 
		return false;
	}			
}

function trim(text)
{
	pos1=0;
	pos2=text.length-1;
	for(i=0;i<=text.length-1;i++)
		if(text.substr(i,1)==" ") pos1=i;
		else break;
	for(i=length-1;i>=0;i--)
		if(text.substr(i,1)==" ") pos2=i;
		else break;
	if (pos2<pos1) return ""
	return text.substr(pos1,pos2-pos1)
}

function checknewpassword(pass,conf)
{
	if( (pass=="") || (conf=="") || (trim(pass)=="") || (trim(conf)=="") ) 
	{	
		if (pass=="") 
		{
			//alert("Password is blank");	
			return false;
		} 
		else 
		{
			//alert("Confirm password is blank");	
			return false;			
		}	
	}
	if ( (pass.length<8) || (conf.length<8) )
	{	
		//alert("The password must be 8 or more characters");	
		return false;
	}
	if (pass!=conf)
	{	
		//alert("Invalid password");	
		return false;
	}
	return true;
}

function istroi(text)
{
 if(text.indexOf(" ")!=-1) return false;
 return true;
}

function isdateusa(s) 
{
	if (s =="") return false;
	var atmp,ii;
	atmp = s.split("/")	
	if (atmp.length != 3) return false;
	for (ii=0; ii<atmp.length ;ii++) {				
		if (isNaN(atmp[ii])) return false;
		else {
			switch (ii) {
				case 0,1: //month					
					if (atmp[ii].length <= 0 || atmp[ii].length > 2) {return false;}
					switch (parseInt(atmp[0])) {
						case 4,6,9,11:
							if (parseInt(atmp[1]) > 30) return false;
							break;
						case 2:
							if (parseInt(atmp[2])%4 == 0 ) 
								{if (parseInt(atmp[1]) > 29) return false;}
							else
								{if (parseInt(atmp[1]) > 28) return false;}
							break;			
						default:
							if (atmp[0] <=0 || atmp[0]>12) return false;
							if (parseInt(atmp[1]) > 31) return false;
					}					
					break;				
				case 2: //year
					if (atmp[ii].length != 4) {return false;}
			}						
		}		
	}
	return true;
}

function checkExt(str)
{
x=str.lastIndexOf(".")
	if(x>0)
	{
		str1=str.substring(1,x-1)
		seconddot=str1.lastIndexOf(".")
		if(seconddot!=-1)
		{
			//alert("seconddot");
			return false
		}
		else
		{		
			ext=str.substring(x+1,str.length);
			//if( (ext!="gif")&&(ext!="jpg"))return false;
			if(ext!="mdb")
				return false;
		}
	}
	else return false;
}

function isemptym(str) 
{	
	atmp = str.split(String.fromCharCode(32));
	for (ii=0; ii< atmp.length;ii++ ) 
	{
		if (atmp[ii] != "") 
		{
			return true;
		}
	}	
	return false;
}

function checkusername(busername)
{
	ls_user = busername
	flag = false
		li_len = ls_user.length
		for (i=0;i<li_len;i++) {
			ls_char = ls_user.substr(i,1).toLowerCase()
			if  (((ls_char >="a") && (ls_char <="z")) || ((ls_char<="9") && (ls_char>="0")) || (ls_char=="_")) {
				flag = true 
			} else { 
				flag = false
				break
			}	
		}
		if (flag == true) 
		{
			//alert("true")
			return true 
		} 
		else return false
}

function value_onchange(vobject) 
{
	if (isempty(eval(vobject).value) == true) 
	{
		if (isNaN(eval(vobject).value))
			{
				alert("Enter the Number !");
				eval(vobject).value=0;
				eval(vobject).focus();				
				return false;
			}	
	 }
	else {eval(vobject).value=0};		
}


function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function alltrim(text)
{
	pos1=0;
	pos2=text.length-1;
	while (text.substring(pos1, pos1 + 1)==" ") {
		pos1 = pos1 + 1;
	}
	pos1 = pos1 - 1;

    pos2 = text.length-1;
	while (text.substring(pos2, pos2 + 1)==" ") {
		pos2 = pos2 - 1;
	}
	pos2 = pos2 + 1
	
	if (pos2<pos1) return ""
	return text.substring(pos1 + 1, pos2)
}

function selfsubmit(straction) {
	window.document.forms[0].target = "_self";
	window.document.forms[0].action = straction;
	window.document.forms[0].submit();
}

/**************************************/
/*                                     */
/**************************************/
function watermark(inputBox, text) {
    
    //var inputBox = document.getElementById(inputId);
    //alert(text);
    if (inputBox.value.length > 0) 
    {
        if (inputBox.value == text) inputBox.value = '';
    }
    else 
        inputBox.value = text; 
}
/**************************************/
/*                                     */
/**************************************/
function Left(str, n) {
    if (n <= 0)
        return "";
    else if (n > String(str).length)
        return str;
    else
        return String(str).substring(0, n);
}
/**************************************/
/*                                     */
/**************************************/
function Right(str, n) {
    if (n <= 0)
        return "";
    else if (n > String(str).length)
        return str;
    else {
        var iLen = String(str).length;
        return String(str).substring(iLen, iLen - n);
    }
}
