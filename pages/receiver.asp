<%
	Dim cn
	Dim rs
	Dim name
	Dim address
	Dim phone
	Dim gen
	Dim birth
	Dim A
	Dim ve
	set cn=server.createObject("ADODB.connection")
	set rs=server.createObject("ADODB.recordset")
	cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =D:\WORKSPACE\pages\receiver.mdb"
	rs.open "select * from receiver",cn,2,2
	rs.AddNew
'rs.fields(0)=Request.form("name")
	name=Request.form("name")
	Response.Write("Name :"&name& "<br>")
'rs.fields(1)=Request.form("address")
	address=Request.form("address")
	Response.Write("Address :"&address& "<br>")
'rs.fields(2)=Request.form("phone")
	phone=Request.form("phone")
	Response.Write("Phone :"&phone& "<br>")
'rs.fields(3)=Request.form("gender")
	gen=Request.form("gen")
	Response.Write("Gender :"&gen& "<br>")
'rs.fields(4)=Request.form("birth")
	birth=Request.form("birth")
	Response.Write("Birth :"&birth& "<br>")
'rs.fields(5)=Request.form("A")
	A=Request.form("A")
	Response.Write("Group :"&A& "<br>")
'rs.fields(6)=Request.form("ve")
	ve=Request.form("ve")
	Response.Write("Type :"&ve& "<br>")


rs.update
cn.Execute("INSERT INTO receiverVALUES('"&name&"','"&address&"','"&phone&"','"&gen&"','"&birth&"','"&A&"','"&ve&"')")
%>
<html>
<a href="http://localhost/pages/receiver.html"> 
New Record addded Successfully in Table</a>
</html>