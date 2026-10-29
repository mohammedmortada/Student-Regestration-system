<%
Dim studentID
studentID = Request.QueryString("id")
If studentID = "" Then Response.Redirect "students.asp"
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("students.mdb") & ";"
conn.Execute("DELETE FROM Students WHERE ID=" & studentID)
conn.Close
Set conn = Nothing
Response.Redirect "students.asp"
%>
