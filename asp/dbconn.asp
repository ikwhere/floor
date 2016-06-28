<%dim db
set db=server.CreateObject("adodb.connection")
//response.Write server.MapPath("db1.mdb")
//db.open "dbq="&server.MapPath("..\db\db1.mdb")&";driver={microsoft access driver (*.mdb)}"
db.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.MapPath("..\db\db1.mdb")
%>