<%
' ============================================================
' calculator.asp - Tuition Fee Monthly Installment Calculator
' Subject 3: ASP Calculations (similar to midterm exam)
' Formula: I = (P + (P * IR/100)) / M
' ============================================================

Dim price, interestRate, months, installment
Dim showResult
showResult = False

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    price        = CDbl(Request.Form("price"))
    interestRate = CDbl(Request.Form("interestRate"))
    months       = CInt(Request.Form("months"))

    ' Calculate monthly installment
    ' I = (P + (P * IR/100)) / M
    installment = (price + (price * interestRate / 100)) / months

    showResult = True

End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Fee Calculator - Student Registration System</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>

    <!-- NAVIGATION -->
    <div class="navbar">
        <div class="logo">&#127891; Student Registration System</div>
        <div class="nav-links">
            <a href="index.html">Home</a>
            <a href="register.asp">Register Student</a>
            <a href="students.asp">View Students</a>
            <a href="calculator.asp">Fee Calculator</a>
            <a href="calculator.asp" class="active">Fee Calculator</a>
        </div>
    </div>

    <div class="container">
        <h2 class="page-title">Tuition Fee Calculator</h2>

        <!-- RESULT BOX -->
        <% If showResult Then %>
        <div class="result-box">
            <h3>Calculation Result</h3>
            <table class="calc-table">
                <tr>
                    <td><strong>Tuition Price:</strong></td>
                    <td>$<%= FormatNumber(price, 2) %></td>
                </tr>
                <tr>
                    <td><strong>Interest Rate:</strong></td>
                    <td><%= interestRate %>%</td>
                </tr>
                <tr>
                    <td><strong>Payment Duration:</strong></td>
                    <td><%= months %> Months</td>
                </tr>
                <tr>
                    <td><strong>Total Amount:</strong></td>
                    <td>$<%= FormatNumber(price + (price * interestRate / 100), 2) %></td>
                </tr>
                <tr class="highlight-row">
                    <td><strong>Monthly Installment:</strong></td>
                    <td><strong>$<%= FormatNumber(installment, 2) %></strong></td>
                </tr>
            </table>
            <p class="formula-note">
                Formula used: I = (P + (P &times; IR/100)) / M
            </p>
        </div>
        <% End If %>

        <div class="form-box">
            <form method="POST" action="calculator.asp">

                <div class="form-group">
                    <label for="price">Tuition Fee Price ($)</label>
                    <input type="number" id="price" name="price" placeholder="e.g. 15000" min="0" step="0.01" required>
                </div>

                <div class="form-group">
                    <label>Payment Duration (Months)</label>
                    <div class="radio-group">
                        <label class="radio-label">
                            <input type="radio" name="months" value="6" required> 6 Months
                        </label>
                        <label class="radio-label">
                            <input type="radio" name="months" value="12"> 12 Months
                        </label>
                    </div>
                </div>

                <div class="form-group">
                    <label>Interest Rate</label>
                    <div class="radio-group">
                        <label class="radio-label">
                            <input type="radio" name="interestRate" value="2" required> 2%
                        </label>
                        <label class="radio-label">
                            <input type="radio" name="interestRate" value="4"> 4%
                        </label>
                    </div>
                </div>

                <div class="form-actions">
                    <input type="submit" value="Calculate Installment" class="btn btn-primary">
                    <a href="index.html" class="btn btn-secondary">Back to Home</a>
                </div>

            </form>
        </div>
    </div>

    <div class="footer">
        <p>AII005 - Internet Programming Project &copy; 2026</p>
    </div>

</body>
</html>
