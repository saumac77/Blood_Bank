<%
	Dim cn
	Dim rs
	Dim name
	Dim email
	Dim phone
	Dim gen
	Dim birth
	Dim A
	Dim ve
	Dim fb
	Dim clear
	set cn=server.createObject("ADODB.connection")
	set rs=server.createObject("ADODB.recordset")
	cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\Documents and Settings\Administrator\Desktop\WORKSPACE\pages\donor.accdb"
	rs.open "select * from donor",cn,2,2
	rs.AddNew
'rs.fields(0)=Request.form("txtNAME")
	name=Request.form("txtNAME")
	Response.Write("Name :"&name& "<br>")
'rs.fields(1)=Request.form("txtADDRESS")
	password=Request.form("txtADDRESS")
	Response.Write("Email :"&address& "<br>")
'rs.fields(2)=Request.form("txtPHONE")
	repassword=Request.form("PHONE")
	Response.Write("Gender :"&phone& "<br>")
'rs.fields(3)=Request.form("GENDER")
	address=Request.form("GENDER")
	Response.Write("Address :"&gender& "<br>")
'rs.fields(4)=Request.form("BIRTH")
	gender=Request.form("BIRTH")
	Response.Write("Gender :"&birth& "<br>")
'rs.fields(5)=Request.form("A")
	gender=Request.form("A")
	Response.Write("Type :"&A& "<br>")

rs.update
cn.Execute("INSERT INTO donor VALUES('"&name&"','"&password&"','"&phone&"','"&gender&"','"&birth&"','"&A&"')")
%>
<html>
<a href="http://localhost/pages/rec1.html"> 
New Record added Successfully in Table</a>
<html>
