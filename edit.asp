<%
Dim message, msgType, studentID
message = ""
msgType = ""
studentID = Request.QueryString("id")
If studentID = "" Then Response.Redirect "students.asp"

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("students.mdb") & ";"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim studentName, studentNumber, department, email, age
    studentName   = Trim(Request.Form("studentName"))
    studentNumber = Trim(Request.Form("studentNumber"))
    department    = Trim(Request.Form("department"))
    email         = Trim(Request.Form("email"))
    age           = Trim(Request.Form("age"))
    If studentName = "" Or studentNumber = "" Or department = "" Or email = "" Or age = "" Then
        message = "Please fill in all fields."
        msgType = "error"
    Else
        Dim updateSQL
        updateSQL = "UPDATE Students SET StudentName='" & studentName & "', StudentNumber='" & studentNumber & "', Department='" & department & "', Email='" & email & "', Age=" & age & " WHERE ID=" & studentID
        conn.Execute(updateSQL)
        message = "Student updated successfully!"
        msgType = "success"
    End If
End If

Dim rs
Set rs = conn.Execute("SELECT * FROM Students WHERE ID=" & studentID)
If rs.EOF Then Response.Redirect "students.asp"
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Edit Student - Student Registration System</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="navbar">
        <div class="logo">&#127891; Student Registration System</div>
        <div class="nav-links">
            <a href="index.html">Home</a>
            <a href="register.asp">Register Student</a>
            <a href="students.asp" class="active">View Students</a>
            <a href="calculator.asp">Fee Calculator</a>
        </div>
    </div>
    <div class="container">
        <h2 class="page-title">Edit Student Record</h2>
        <% If message <> "" Then %>
            <div class="alert alert-<%= msgType %>"><%= message %></div>
        <% End If %>
        <div class="form-box">
            <form method="POST" action="edit.asp?id=<%= studentID %>">
                <div class="form-row">
                    <div class="form-group">
                        <label>Full Name</label>
                        <input type="text" name="studentName" value="<%= rs("StudentName") %>" required>
                    </div>
                    <div class="form-group">
                        <label>Student Number</label>
                        <input type="text" name="studentNumber" value="<%= rs("StudentNumber") %>" required>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Department</label>
                        <select name="department" required>
                            <option value="">-- Select Department --</option>
                            <% Dim depts(4)
                            depts(0) = "Computer Engineering"
                            depts(1) = "Software Engineering"
                            depts(2) = "Information Technology"
                            depts(3) = "Electrical Engineering"
                            depts(4) = "Business Administration"
                            Dim i
                            For i = 0 To 4 %>
                                <option value="<%= depts(i) %>" <%= If(rs("Department") = depts(i), "selected", "") %>><%= depts(i) %></option>
                            <% Next %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Age</label>
                        <input type="number" name="age" value="<%= rs("Age") %>" min="16" max="60" required>
                    </div>
                </div>
                <div class="form-group">
                    <label>Email Address</label>
                    <input type="email" name="email" value="<%= rs("Email") %>" required>
                </div>
                <div class="form-actions">
                    <input type="submit" value="Save Changes" class="btn btn-success">
                    <a href="students.asp" class="btn btn-secondary">Cancel</a>
                </div>
            </form>
        </div>
    </div>
    <% rs.Close : conn.Close : Set rs = Nothing : Set conn = Nothing %>
    <div class="footer"><p>AII005 - Internet Programming Project &copy; 2026</p></div>
</body>
</html>
