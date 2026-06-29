<%@ Language=VBScript CodePage=65001 %>
<%
' Force UTF-8 response
Response.CharSet = "UTF-8"
Response.CodePage = 65001

' =========================================================================
' --- SMTP CONFIGURATION SECTION ---
Dim SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS, USE_AUTH, USE_SSL
SMTP_SERVER = "mail.incitycarhire.com" ' SMTP Host
SMTP_PORT   = 465                      ' SMTP Port for SSL
USE_AUTH    = True                     ' Authentication enabled
SMTP_USER   = "info@incitycarhire.com" ' SMTP Username
SMTP_PASS   = "Ozkan39300"            ' SMTP Password
USE_SSL     = True                     ' SSL enabled
' =========================================================================

Dim fullname, phone, serviceType, vehicle, startDate, endDate, notes
fullname = Request.Form("fullname")
phone = Request.Form("phone")
serviceType = Request.Form("serviceType")
vehicle = Request.Form("vehicle")
startDate = Request.Form("startDate")
endDate = Request.Form("endDate")
notes = Request.Form("notes")

Dim isSent, errorMsg
isSent = False
errorMsg = ""

Dim honeypot
honeypot = Request.Form("website_url")

If honeypot <> "" Then
    ' Bot detected. Pretend it was sent to avoid giving clues to the spammer.
    isSent = True
ElseIf fullname <> "" And phone <> "" Then
    Dim ObjSendMail
    Set ObjSendMail = CreateObject("CDO.Message")
    
    ' CDO configuration
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
    
    If USE_AUTH Then
        ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USER
        ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASS
    End If
    
    If USE_SSL Then
        ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    End If
    
    ObjSendMail.Configuration.Fields.Update
    
    ObjSendMail.To = "info@incitycarhire.com"
    ObjSendMail.Subject = "New Booking Request from " & fullname
    ObjSendMail.From = SMTP_USER
    
    Dim emailBody
    emailBody = "<html><head><style>body{font-family:Arial,sans-serif;color:#333;}table{border-collapse:collapse;width:100%;max-width:600px;}td{padding:8px;border:1px solid #ddd;}</style></head><body>" & _
                "<h2>New Booking Request - incity Rent a Car</h2>" & _
                "<table>" & _
                "<tr><td><strong>Full Name</strong></td><td>" & Server.HTMLEncode(fullname) & "</td></tr>" & _
                "<tr><td><strong>Phone Number</strong></td><td>" & Server.HTMLEncode(phone) & "</td></tr>" & _
                "<tr><td><strong>Service Type</strong></td><td>" & Server.HTMLEncode(serviceType) & "</td></tr>" & _
                "<tr><td><strong>Selected Vehicle</strong></td><td>" & Server.HTMLEncode(vehicle) & "</td></tr>" & _
                "<tr><td><strong>Start Date</strong></td><td>" & Server.HTMLEncode(startDate) & "</td></tr>" & _
                "<tr><td><strong>End Date</strong></td><td>" & Server.HTMLEncode(endDate) & "</td></tr>" & _
                "<tr><td><strong>Notes / Details</strong></td><td>" & Replace(Server.HTMLEncode(notes), vbCrLf, "<br>") & "</td></tr>" & _
                "</table></body></html>"
                
    ObjSendMail.HTMLBody = emailBody
    
    On Error Resume Next
    ObjSendMail.Send
    If Err.Number = 0 Then
        isSent = True
    Else
        errorMsg = Err.Description
    End If
    On Error GoTo 0
    Set ObjSendMail = Nothing
Else
    errorMsg = "Required fields are missing."
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0">

    <title>Booking Status | İncity Rent a Car Fethiye</title>

    <link href="assets/css/bootstrap.css" rel="stylesheet">    
    <link href="assets/css/fontawesome-all.css" rel="stylesheet">  
    <link href="assets/css/iconfont.css" rel="stylesheet">  
    <link href="assets/css/owl.css" rel="stylesheet">
    <link href="assets/css/global.css" rel="stylesheet">
    <link href="assets/css/jquery.fancybox.min.css" rel="stylesheet">
    <link href="assets/css/elements-css/header.css" rel="stylesheet">
    <link href="assets/css/elements-css/footer.css" rel="stylesheet">
    <link href="assets/css/style.css" rel="stylesheet">
    <link href="assets/css/responsive.css" rel="stylesheet">

    <!-- Fav Icon -->
    <link rel="shortcut icon" href="assets/images/favicon.png" type="image/x-icon">
    <link rel="icon" href="assets/images/favicon.png" type="image/x-icon">

    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@100..900&display=swap" rel="stylesheet">
</head>

<body class="boxed_wrapper dark-bg">

<!-- Main Header -->
<header class="header dark-header">
    <div class="main_header">
        <div class="container"> 
            <div class="main_header_inner">
                <div class="main_header_logo">
                    <figure>
                        <a href="index.html">
                            <img src="incity-logo.svg" alt="İncity Rent a Car Logo" style="height: 50px; width: auto;">
                        </a>
                    </figure>
                </div>
                <div class="main_header_menu menu_area">
                    <div class="mobile-nav-toggler">
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                        <span class="icon-bar"></span>
                    </div>
                    <nav class="main-menu">
                        <div class="collapse navbar-collapse show" id="navbarSupportedContent">
                            <ul class="navigation">
                                <li><a href="index.html">Home</a></li>
                                <li><a href="about-us.html">About Us</a></li>
                                <li><a href="car-fleet.html">Cars</a></li>
                                <li class="active"><a href="booking.html">Booking</a></li>
                                <li><a href="contact.html">Contact</a></li>
                            </ul>
                        </div>
                    </nav>
                </div>
            </div>
        </div>
    </div>
</header>

<section class="booking-status-section pt_120 pb_120 text-center">
    <div class="container">
        <div class="row">
            <div class="col-lg-8 mx-auto" style="background-color: #1a1a1a; padding: 50px; border-radius: 10px; border: 1px solid #333;">
                <% If isSent Then %>
                    <span style="font-size: 80px; color: #25d366;"><i class="far fa-check-circle"></i></span>
                    <h2 class="mt-4" style="color: #fff; font-weight: 700;">Booking Request Sent!</h2>
                    <p class="mt-3" style="color: #e0e0e0; font-size: 18px;">Thank you for choosing İncity Rent a Car. We have received your request and will contact you via email or phone shortly to confirm your booking.</p>
                <% Else %>
                    <span style="font-size: 80px; color: #ff3b30;"><i class="far fa-times-circle"></i></span>
                    <h2 class="mt-4" style="color: #fff; font-weight: 700;">Submission Failed</h2>
                    <p class="mt-3" style="color: #e0e0e0; font-size: 18px;">We encountered an issue while sending your booking via email: <strong><%= errorMsg %></strong></p>
                    <p style="color: #e0e0e0; font-size: 16px;">Please try sending your request via WhatsApp or contact us directly.</p>
                <% End If %>
                <div class="mt-5">
                    <a href="booking.html" class="btn-style-four" style="background-color: #e5ab1c; border-color: #e5ab1c; color: #111; border-radius: 5px; font-size: 16px; padding: 12px 30px;">Go Back</a>
                    <a href="https://wa.me/905342890048" class="btn-style-four" style="background-color: #25d366; border-color: #25d366; color: #fff; border-radius: 5px; font-size: 16px; padding: 12px 30px; margin-left: 15px;"><i class="fab fa-whatsapp"></i> Chat on WhatsApp</a>
                </div>
            </div>
        </div>
    </div>
</section>

<!-- Main Footer -->
<footer class="main_footer">
    <div class="container">
        <div class="footer-top-outer">
            <div class="row">
                <div class="col-xl-4 col-lg-4 col-md-6 col-sm-12">
                    <div class="footer_widget footer_about_widget">
                        <figure class="footer_widget_logo">
                            <a href="index.html">
                                <img src="incity-logo.svg" alt="İncity Rent a Car" style="height: 55px; width: auto; margin-bottom: 20px;">
                            </a>
                        </figure>
                        <p style="color: #e0e0e0;">Providing professional car hire & Dalaman airport transfer services in Fethiye and surrounding regions since 2017.</p>
                    </div>
                </div>
                <div class="col-xl-5 col-lg-5 col-md-6 col-sm-12">
                    <div class="footer_widget footer_contact_widget">
                        <h4 class="footer_widget_title">Contact İncity</h4>
                        <p>Akarca Mah. Şehit Nihat Oran Caddesi, No: 30, Fethiye, Turkey</p>
                        <ul class="footer-info-list">
                            <li>Email: <a href="mailto:info@incitycarhire.com">info@incitycarhire.com</a></li>
                            <li>Phone: <a href="tel:+905343656565">+90 534 365 65 65</a></li>
                            <li>Phone: <a href="tel:+905342890048">+90 534 289 00 48</a></li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    </div>
</footer>

<script src="assets/js/jquery-3.7.1.min.js"></script>
<script src="assets/js/bootstrap.min.js"></script>
<script src="assets/js/main.js"></script>

</body>
</html>
