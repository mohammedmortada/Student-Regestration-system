<%
Dim message, msgType
message = ""
msgType = ""

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
        Dim conn, sql
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("students.mdb") & ";"
        sql = "INSERT INTO Students (StudentName, StudentNumber, Department, Email, Age) VALUES ('" & studentName & "', '" & studentNumber & "', '" & department & "', '" & email & "', " & age & ")"
        conn.Execute(sql)
        conn.Close
        Set conn = Nothing
        message = "Student registered successfully!"
        msgType = "success"
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Register Student - Student Registration System</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="navbar">
        <div class="logo">&#127891; Student Registration System</div>
        <div class="nav-links">
            <a href="index.html">Home</a>
            <a href="register.asp" class="active">Register Student</a>
            <a href="students.asp">View Students</a>
            <a href="calculator.asp">Fee Calculator</a>
        </div>
    </div>
    <div class="container">
        <h2 class="page-title">Register New Student</h2>
        <% If message <> "" Then %>
            <div class="alert alert-<%= msgType %>"><%= message %></div>
        <% End If %>
        <div class="form-box">
            <form method="POST" action="register.asp">
                <div class="form-row">
                    <div class="form-group">
                        <label>Full Name</label>
                        <input type="text" name="studentName" placeholder="Enter full name" required>
                    </div>
                    <div class="form-group">
                        <label>Student Number</label>
                        <input type="text" name="studentNumber" placeholder="e.g. 20210001" required>
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>Department</label>
                        <select name="department" required>
                            <option value="">-- Select Department --</option>
                            <option value="Computer Engineering">Computer Engineering</option>
                            <option value="Software Engineering">Software Engineering</option>
                            <option value="Information Technology">Information Technology</option>
                            <option value="Electrical Engineering">Electrical Engineering</option>
                            <option value="Business Administration">Business Administration</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Age</label>
                        <input type="number" name="age" placeholder="e.g. 20" min="16" max="60" required>
                    </div>
                </div>
                <div class="form-group">
                    <label>Email Address</label>
                    <input type="email" name="email" placeholder="e.g. student@university.edu" required>
                </div>
                <div class="form-actions">
                    <input type="submit" value="Register Student" class="btn btn-success">
                    <a href="students.asp" class="btn btn-secondary">View All Students</a>
                </div>
            </form>
        </div>
    </div>
    <div class="footer"><p>AII005 - Internet Programming Project &copy; 2026</p></div>
</body>
</html>
