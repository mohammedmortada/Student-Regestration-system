<%
' ============================================================
' students.asp - View all students with Search, Filter, Sort, Print
' ============================================================

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("students.mdb") & ";"

Dim searchTerm, filterDept, sortBy
searchTerm = Trim(Request.QueryString("search"))
filterDept = Trim(Request.QueryString("dept"))
sortBy     = Trim(Request.QueryString("sort"))
If sortBy = "" Then sortBy = "ID DESC"

Dim sql
sql = "SELECT * FROM Students WHERE 1=1"
If searchTerm <> "" Then
    sql = sql & " AND (StudentName LIKE '%" & searchTerm & "%' OR StudentNumber LIKE '%" & searchTerm & "%')"
End If
If filterDept <> "" Then
    sql = sql & " AND Department='" & filterDept & "'"
End If
sql = sql & " ORDER BY " & sortBy

Dim rs
Set rs = conn.Execute(sql)

Dim rsCount, totalCount
Set rsCount = conn.Execute("SELECT COUNT(*) AS Total FROM Students")
totalCount = rsCount("Total")
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>All Students - Student Registration System</title>
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
        <h2 class="page-title">All Registered Students</h2>

        <div class="top-bar">
            <p>Total Students: <strong><%= totalCount %></strong></p>
            <div style="display:flex; gap:10px; flex-wrap:wrap;">
                <a href="register.asp" class="btn btn-primary">+ Register New Student</a>
                <button onclick="window.print()" class="btn btn-secondary">Print</button>
            </div>
        </div>

        <!-- SEARCH AND FILTER -->
        <div class="search-box">
            <form method="GET" action="students.asp">
                <div class="form-row">
                    <div class="form-group">
                        <label>Search by Name or Student Number</label>
                        <input type="text" name="search" value="<%= searchTerm %>" placeholder="e.g. Mohammed or 20210001">
                    </div>
                    <div class="form-group">
                        <label>Filter by Department</label>
                        <select name="dept">
                            <option value="">-- All Departments --</option>
                            <%
                            Dim depts(4), i
                            depts(0) = "Computer Engineering"
                            depts(1) = "Software Engineering"
                            depts(2) = "Information Technology"
                            depts(3) = "Electrical Engineering"
                            depts(4) = "Business Administration"
                            For i = 0 To 4
                                If filterDept = depts(i) Then
                                    Response.Write "<option value='" & depts(i) & "' selected>" & depts(i) & "</option>"
                                Else
                                    Response.Write "<option value='" & depts(i) & "'>" & depts(i) & "</option>"
                                End If
                            Next
                            %>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Sort By</label>
                        <select name="sort">
                            <option value="ID DESC">Latest First</option>
                            <option value="StudentName ASC">Name A-Z</option>
                            <option value="Age ASC">Age Low to High</option>
                            <option value="Age DESC">Age High to Low</option>
                            <option value="Department ASC">Department A-Z</option>
                        </select>
                    </div>
                </div>
                <div class="form-actions">
                    <input type="submit" value="Search / Filter" class="btn btn-success">
                    <a href="students.asp" class="btn btn-secondary">Clear</a>
                </div>
            </form>
        </div>

        <!-- TABLE -->
        <div class="table-box">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Full Name</th>
                        <th>Student Number</th>
                        <th>Department</th>
                        <th>Email</th>
                        <th>Age</th>
                        <th class="no-print">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    <% If rs.EOF Then %>
                        <tr><td colspan="7" class="no-records">No students found.</td></tr>
                    <% Else %>
                        <%
                        Dim counter
                        counter = 1
                        Do While Not rs.EOF
                        %>
                        <tr>
                            <td><%= counter %></td>
                            <td><%= rs("StudentName") %></td>
                            <td><%= rs("StudentNumber") %></td>
                            <td><%= rs("Department") %></td>
                            <td><%= rs("Email") %></td>
                            <td><%= rs("Age") %></td>
                            <td class="no-print">
                                <div class="actions">
                                    <a href="edit.asp?id=<%= rs("ID") %>" class="btn btn-warning">Edit</a>
                                    <a href="delete.asp?id=<%= rs("ID") %>" class="btn btn-danger" onclick="return confirm('Are you sure?')">Delete</a>
                                </div>
                            </td>
                        </tr>
                        <%
                            counter = counter + 1
                            rs.MoveNext
                        Loop
                        %>
                    <% End If %>
                </tbody>
            </table>
        </div>
    </div>

    <%
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    %>

    <div class="footer"><p>AII005 - Internet Programming Project &copy; 2026</p></div>
</body>
</html>
