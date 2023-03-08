<%
Response.Write @@curpage & " " & @@numpage
Response.End
%>
<table width="780" border="0" cellspacing="0" cellpadding="0" height="80%" align="center">
  <tr> 
    <td width="6" background="../../../images/l-03-3b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
    <td width="164" valign="top" height="100%">@@menu</td>
    <td width="1"  background="../../../images/dot-01.gif" bgcolor="#FFE8E8" height="100%"><img src="../images/dot-01.gif" width="1" height="1"></td>
    <td width="610" valign="top">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">          
			@@content
          <tr> 
            <td> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20">
			  <tr> 
			    <td align="right" bgcolor="#E7EBF5"> 
			      <table width="70%" border="0" cellspacing="1" cellpadding="0" height="20">
			        <tr class="black-normal"> 
			          <td align="right" valign="middle" width="37%" class="blue-normal">Page 
			          </td>
			          <td align="center" valign="middle" width="13%" class="blue-normal"> 
			            <input type="text" name="txtpage" class="blue-normal" value="@@curpage" size="2" style="width:50">
			          </td>
			          <td align="left" valign="middle" width="7%" class="blue-normal">&nbsp;<a href="javascript:go();"  onMouseOver="self.status='Go to page'; return true;" onMouseOut="self.status='';"><font color="#990000">Go</font></a> 
			          </td>
			          <td align="right" valign="middle" width="15%" class="blue-normal"><%If CInt(@@numpage) <> 0 Or @@numpage <> "" Then%>Pages @@curpage/@@numpage<%End If%>&nbsp;&nbsp;</td>
			          <td valign="middle" align="right" width="28%" class="blue-normal"><%If CInt(@@curpage) <> 1 Then%><a href="javascript:prev();"  
			          onMouseOver="self.status='Previous page'; return true;" onMouseOut="self.status='';">Previous</a><%End If%><%If CInt(@@curpage) <> 1 And  CInt(@@curpage) <> CInt(@@numpage) Then%>/<%End If%><%If CInt(@@curpage) <> CInt(@@numpage) And (CInt(@@numpage) <> 0 Or @@numpage <> "") Then%><a href="javascript:next();"  onMouseOver="self.status='Next page'; return true;" onMouseOut="self.status='';"> Next</a><%End If%>&nbsp;&nbsp;&nbsp;</td>
			        </tr>
			      </table>
			    </td>
			  </tr>
			</table>
	        </td>
          </tr>
        </table>    
    </td>
    <td width="2" background="../../../images/l-03-2b.gif" bgcolor="#FFE8E8" height="100%">&nbsp;</td>
  </tr>
</table>