<div align="center">

## Creating a ADO connection to SQL Server


</div>

### Description

A article showing how to create a global ADO connection to SQL server from a Visual Basic Client.I have more examples and programming solutions on my web site www.SQLwarehouse.com
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joe Povilaitis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joe-povilaitis.md)
**Level**          |Beginner
**User Rating**    |3.8 (30 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joe-povilaitis-creating-a-ado-connection-to-sql-server__1-12376/archive/master.zip)





### Source Code

<p><font color="#000099"><img border="0" src="bd00386_.gif" width="100%" height="4"></font></p>
<p align="center"><font color="#0000FF" size="4"><b>Creating a ADO Connection To
SQL Server</b></font></p>
<p>Here is a example to create a ado connection. You could create a&nbsp; basic
module and add it to your project and then create a Global ADO connection, so
your program will use one connection instance for the whole program. That way
once you open up your connection it will stay until you close the connection or
exit the program. Make sure in your VB project , you have in your references
menu option,Microsoft Activex Dataobjects selected. And also Dcom installed.</p>
<p>In the General/declarations of your basic module declare your connection ..</p>
<p><font color="#008000" size="2"><b>Global SQLCON As New ADODB.Connection</b></font></p>
<p><font color="#000000">Then , in your project , say under a command button the
code to open your connection</font>, would be ...</p>
<p><font color="#008000"><b><font size="2">Public Sub Command1_Click()<br>
&nbsp;&nbsp;&nbsp; ' Connect to SQL server through SQL Server OLE DB Provider.<br>
</font></b></font></p>
<p><font color="#008000"><b><font size="2">&nbsp;&nbsp;&nbsp; ' Set the ADO connection properties.<br>
&nbsp;&nbsp;&nbsp; SQLCON.ConnectionTimeout = 25&nbsp; ' Time out for the
connection<br>
&nbsp;&nbsp;&nbsp; SQLCON.Provider = "sqloledb"&nbsp;&nbsp; ' OLEDB Provider<br>
&nbsp;&nbsp;&nbsp; SQLCON.Properties("Network Address").Value =
&quot;111.111.111.111&quot;&nbsp; ' set the ip address of your sql server<br>
&nbsp;&nbsp;&nbsp; SQLCON.CommandTimeout = 180 ' set timeout for 3 minutes<br>
<br>
&nbsp;&nbsp;&nbsp; ' Now set your network library to use one of these libraries
.. un-rem only the one you want to use !<br>
&nbsp;&nbsp;&nbsp; 'SQLCON.Properties("Network Library").Value = "dbmssocn" ' set the network library to use win32 winsock
tcp/ip<br>
&nbsp;&nbsp;&nbsp; 'SQLCON.Properties("Network Library").Value = "dbnmpntw" ' set the network library to use win32 named
pipes<br>
&nbsp;&nbsp;&nbsp; 'SQLCON.Properties("Network Library").Value = "dbmsspxn" ' set the network library to use win32
spx/ipx<br>
&nbsp;&nbsp;&nbsp; 'SQLCON.Properties("Network Library").Value = "dbmsrpcn" ' set the network library to use win32
multi-protocol</font></b></font></p>
<p><font size="2" color="#008000"><b>&nbsp;&nbsp;&nbsp; 'Now set the SQL server
name , and the default data base .. change these for your server !</b></font><font size="2"><b><font color="#008000"><br>
&nbsp;&nbsp;&nbsp; SQLCON.Properties("Data Source").Value = &quot;MYSERVERNAME&quot;<br>
&nbsp;&nbsp;&nbsp; SQLCON.Properties("Initial Catalog").Value = &quot;MYSQLDATABASE&quot;<br>
&nbsp;&nbsp;&nbsp; SQLCON.CursorLocation = adUseServer ' For ADO cursor location<br>
<br>
&nbsp;&nbsp;&nbsp; 'Now you need to decide what authorization type you want to
use .. WinNT or SQL Server.<br>
&nbsp;&nbsp;&nbsp; 'un-rem this line for NT authorization.</font></b></font></p>
<p><font size="2"><b><font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
'SQLCON.Properties("Integrated Security").Value = &quot;SSPI&quot;</font></b></font></p>
<p><font color="#008000" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp; ' Or if you want
to use SQL authorization , un-rem these 2 lines and supply SQL server login name
and password</b></font></p>
<p><font color="#008000">&nbsp;&nbsp;&nbsp; '</font><font size="2"><b><font color="#008000">SQLCON.Properties("User ID").Value =&quot;SQLUSERNAME&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp; 'SQLCON.Properties(&quot;Password&quot;).Value =
&quot;SQLPASSWORD&quot;<br>
</font>
<br>
<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp; ' Now we can open&nbsp; the ADO Connection to SQl
server&nbsp; !..<br>
&nbsp;&nbsp;&nbsp;&nbsp; SQLCON.Open<br>
</font>
</b></font></p>
<p>&nbsp;&nbsp;&nbsp;<font size="2" color="#008000"><b> ' Now we can do a simple
test of the new ADO connection<br>
&nbsp;&nbsp;&nbsp;&nbsp; ' Lets return the Time and Date the SQL server thinks
it is ..</b></font></p>
<p><font size="2" color="#008000"><b>&nbsp;&nbsp;&nbsp; Dim RS As ADODB.Recordset<br>
&nbsp;&nbsp;&nbsp; Set RS = New ADODB.Recordset<br>
&nbsp;&nbsp;&nbsp; SQLstatement = "SELECT GETDATE() AS SQLDATE &quot; ' Set a
Simple Sql query to return the servers time<br>
&nbsp;&nbsp;&nbsp; RS.Open SQLstatement, SQLCON&nbsp; ' Lets open a connection
with our new SQLCON connection , and our SQL statement<br>
&nbsp;&nbsp;&nbsp; ' Move to first row.<br>
&nbsp;&nbsp;&nbsp; RS.MoveFirst<br>
&nbsp;&nbsp;&nbsp; junk = MsgBox( &quot;Server Time is &quot; &amp; RS(&quot;SQLDATE&quot;),
vbOKOnly, &quot; SQL SERVER INFO")<br>
</b></font>&nbsp;&nbsp;&nbsp;</p>
<p><font size="2" color="#008000"><b>End Sub</b></font></p>
<p><font color="#008000"><br>
</font><font color="#000000">Of course , you need to add error handling routines
, and more user friendly code, if you want selectable logon options, but this
should at least get you talking to the SQL server.</font></p>
&nbsp;
<p align="center"><img border="0" src="newlogosmall.jpg" width="480" height="120"></p>
<p align="center">

