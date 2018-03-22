<html>
<head><title>feedback</title>
<%
 dim strconn, objconn, objRS
set objconn=server.createobject("ADODB.Connection")
strconn=Provider="Microsoft.Jet.OLEDB.4.0; data source="&_"packages1.accdb"

set objRS="server.createobject("ADODB.Recordset")
objRS.open "Order" , objconn
objRs("Name")=request.form("t1")
objRs("Address")=request.form("t2")
objRs("Packages")=request.form("t3")
objRs("No of adults")=request.form("t4")
objRs("No of Kids")=request.form("t5")
objRS.close
objconn.close
set objRS=nothing
set objconn=nothing
%>
<script language="vbscript">
sub b1_onclick
dim a,b,c
document.form1.t4.value=a
document.form1.t5.value=b
document.form1.t6.value=c

if selected.value="0" then
c=(cint(a)*10000)+(cint(b)*8600)
else if selected.value="1" then
c=(cint(a)*5500)+(cint(b)*3500)
else if selected.value="2" then
c=(cint(a)*8500)+(cint(b)*6500)
else 
c=(cint(a)*7300)+(cint(b)*5800)
end if

</head>
<body bgcolor="green">
<embed src="C:\Users\Owner\Desktop\html\ishq wala love instrumental - [Mp3Truck.Net].mp3"
hidden="true" autostart="true" loop="true">
<form name="form1">

<center><% 
set mycontrot=server.createobject("mswc.contentrotator")
response.write(mycontrot.choosecontent("contentrotator.txt"))
%></center>

<b>Name:</b><input type="text" maxlength="40" name="t1"><br>
<b>Phone:</b><input type="text" maxlength="15" name="t2"><br>
<b>Address:</b><textarea cols="40" rows="5"></textarea><br>
<b>City:</b><input type="text"><br>
<b>Country:</b><input type="text"><br>
<b>E-mail:</b><input type="text" maxlength="60"></br>

<b>Sex:</b><br>
Male :<input type="radio" name=1 value="Male">
Female:<input type="radio" name=1 value="Female"><br>

<b>Please choose the package desired:</b><br><select name="t3">
<option selected>PACKAGE 1</option>
<option>PACKAGE 2</option>
<option>PACKAGE 3</option>
<option>PACKAGE 4</option>
</select><br>

<b>Number of adults:</b><input type="text" maxlength="3" name="t4"><br>
<b>Number of kids   :</b><input type="text" maxlength="3" name="t5"><br>

<b>Any Special Requirements:</b><br><textarea cols="30" rows="5"></textarea><br>
<input type="submit" value="Submit" name="b1">
<input type="submit" value="Reset">
<br>
Calculated Amount:<input type="text" name="t6">
</form>
</body>
</html>

