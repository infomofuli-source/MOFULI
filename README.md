<!DOCTYPE html>
<html lang="sw">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AFMIND DRAUGHTS LEAGUE – Firebase</title>

    <!-- Bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <!-- SheetJS (Excel) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- jsPDF -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>

    <style>
        /* ===== ZOHO STYLES (hazijabadilishwa) ===== */
        html, body, div, span, applet, object, iframe, h1, h2, h3, h4, h5, h6, p, blockquote, pre, a, abbr, acronym, address, big, cite, code, del, dfn, em, img, ins, kbd, q, s, samp, small, strike, strong, sub, sup, tt, var, b, u, i, center, dl, dt, dd, ol, ul, li, fieldset, form, label, legend, caption, article, aside, canvas, details, embed, figure, figcaption, footer, header, hgroup, menu, nav, output, ruby, section, summary, time, mark, audio, video { margin:0; padding:0; border:0; font:inherit; vertical-align:baseline; }
        body { background-attachment:fixed; color:#444444; font:75%/1.3 Arial, Helvetica, sans-serif; margin:0 auto; }
        input, input[type="text"], input[type="search"], isindex, textarea, button { outline:none; margin:0 auto; padding:5px 10px; -webkit-box-sizing:border-box; -moz-box-sizing:border-box; box-sizing:border-box; }
        img, a { border:0px; outline:none; color:#3a6cae; text-decoration:none; }
        img, a:hover { text-decoration:underline; }
        .zf-flLeft { float:left; } .zf-flRight { float:right; } .zf-clearBoth, .zf-eclearBoth { clear:both; }
        ol, ul { list-style:none outside none; }
        .zf-normalText { font-size:13px; line-height:1.5; }
        .zf-smallText { font-size:0.9em; font-weight:normal; }
        .zf-smallHeading { font-size:18px; } .zf-heading { font-size:2em; } .zf-subHeading { font-size:1.5em; }
        .zf-boldText, .zf-boldText a { font-weight:bold; text-decoration:none; }
        .zf-italicText { font-style:italic; } .zf-heading a { text-decoration:underline; } .zf-heading a:hover { text-decoration:none; }
        .zf-blodText { font-weight:bold; } .zf-overAuto { overflow:auto; } .zf-split { color:#8B9193; padding:0 3px; }
        .zf-backgroundBg { background:rgba(237,239,244,1); }
        .zf-templateWidth { margin:0 auto; padding:30px 20px; max-width:800px; width:100%; box-sizing:border-box; }
        .zf-templateWrapper { background:#fff; -webkit-box-shadow:0px 0px 22px 0px #d8dfed; -moz-box-shadow:0px 0px 22px 0px #d8dfed; box-shadow:0px 0px 22px 0px #d8dfed; -webkit-border-radius:2px; -moz-border-radius:2px; border-radius:10px; }
        .zf-tempContDiv input[type="text"], .zf-tempContDiv textarea, .zf-tempContDiv .zf-pdfTextArea { background:#fff; border:1px solid rgba(184,187,211,1); -webkit-border-radius:2px; -moz-border-radius:2px; border-radius:4px; padding:5px; font-size:15px; color:rgb(37,44,62); padding:11px 10px 10px 10px; height:40px; transition:0.3s; }
        .zf-tempContDiv input[type="text"]::placeholder, .zf-tempContDiv textarea::placeholder, .zf-tempContDiv .zf-pdfTextArea::placeholder { color:rgb(37,44,62); opacity:0.5; }
        .zf-tempContDiv input[type="text"]:hover, .zf-tempContDiv textarea:hover { border:1px solid rgba(184,187,211,1); }
        .zf-tempContDiv input[type="text"]:focus, .zf-tempContDiv textarea:focus { border:1px solid #2eb79f; box-shadow:0px 0px 2px 0px #2eb79f; }
        .zf-tempContDiv textarea { min-height:100px; height:100px; font-family:Arial, Helvetica, sans-serif; padding:10px; }
        .zf-errorMessage { font:15px Arial, Helvetica, sans-serif; color:#f41033; padding-top:10px; }
        .zf-important { color:#ff0000 !important; padding:0; font-size:17px !important; margin-left:2px; font-weight:bold; }
        .zf-instruction { color:#465475; font-style:normal; font-size:13px; overflow:visible !important; word-break:break-all; padding:8px 0px 0px 0px; font-weight:500; clear:both; }
        .zf-symbols { padding:0 5px; } .zf-overflow { overflow:hidden; }
        .zf-tempHeadBdr { margin:0; padding:0; overflow:hidden; }
        .zf-tempHeadContBdr { background:#ffffff; border-bottom:1px solid #ced3e0; margin:0; padding:28px 40px; -webkit-border-radius:2px 2px 0 0; -moz-border-radius:2px 2px 0 0; border-radius:10px 10px 0 0; }
        .zf-tempHeadContBdr .zf-frmTitle { color:#252c3e; margin:0; padding:0; font-size:33px; font-weight:500; text-align:center; }
        .zf-tempHeadContBdr .zf-frmDesc { color:#667291; font-size:18px; font-weight:400; margin:0; padding-top:8px; text-align:center; }
        .zf-subContWrap { padding:16px 0 16px 0; }
        .zf-tempFrmWrapper { padding:16px 40px 16px 40px; margin:0; box-sizing:border-box; }
        .zf-tempFrmWrapper .zf-tempContDiv { margin:0; padding:0; }
        .zf-tempFrmWrapper .zf-labelName { font-weight:500; font-size:16px; color:#252c3e; }
        .zf-form-sBox { padding:9px 10px 10px 5px; background:#fff; border:1px solid rgba(184,187,211,1); border-radius:4px; font-size:16px; height:40px; transition:0.3s; vertical-align:middle; position:relative; color:rgba(37,44,62,1); }
        .zf-form-sBox:focus, .zf-form-sBox:focus:hover { border:1px solid #2eb79f; box-shadow:0px 0px 2px 0px #2eb79f; outline:none; }
        .zf-form-sBox::after { content:" "; position:absolute; }
        .zf-name .zf-tempContDiv span, .zf-phone .zf-tempContDiv span, .zf-time .zf-tempContDiv span { float:left; display:block; }
        .zf-name .zf-tempContDiv span { margin-left:4%; } .zf-name .zf-tempContDiv span.last { margin-right:0; }
        .zf-name .zf-tempContDiv span label { display:block; padding-top:3px; }
        .zf-name .zf-tempContDiv input[type="text"] { width:100%; }
        .zf-phone .zf-tempContDiv span label, .zf-date .zf-tempContDiv span label, .zf-time .zf-tempContDiv span label, .zf-address .zf-tempContDiv span label, .zf-geolocation .zf-tempContDiv span label, .zf-name .zf-tempContDiv span label { font-style:normal; font-size:13px; overflow:visible !important; word-break:break-all; padding:8px 0px 0px 0px; font-weight:500; }
        .zf-phone .zf-tempContDiv label, .zf-date .zf-tempContDiv label, .zf-time .zf-tempContDiv label, .zf-address .zf-tempContDiv label, .zf-name .zf-tempContDiv span label { color:#252c3e; opacity:.8; }
        .zf-phone .zf-tempContDiv span label { display:block; padding-top:3px; text-align:left; }
        .zf-phone .zf-tempContDiv .zf-symbols { display:block; margin:9px 1%; width:2%; text-align:center; padding:0; padding-top:3px; }
        .zf-currency .zf-tempContDiv span { display:inline-block; font-size:15px; font-weight:600; color:#252c3e; margin-right:8px; float:left; margin-top:10px; }
        .zf-currency .zf-tempContDiv input[type="text"]~span { margin-left:8px; margin-right:0; float:none; }
        .zf-currency .zf-tempContDiv span label { display:block; padding-top:3px; }
        .zf-currency .zf-tempContDiv .zf-symbol { font-size:14px; margin-left:5px; margin-top:4px; width:auto; font-weight:bold; }
        .zf-decesion .zf-tempContDiv { width:100% !important; margin-top:4px; }
        .zf-decesion input[type="checkbox"] { display:block; height:13px; margin:0; padding:0; width:13px; float:left; margin-top:4px; }
        .zf-decesion label { display:block; margin:0px 0 0 25px !important; padding-bottom:0 !important; width:95% !important; float:none !important; line-height:21px !important; text-align:left !important; }
        .zf-tempContDiv input[type="file"] { outline:none; margin:0 auto; width:50%; border:1px dashed rgba(184,187,211,1); border-radius:4px; display:inline-block; vertical-align:middle; padding:12px 12px; font-size:15px; color:rgb(37,44,62); }
        .zf-address .zf-tempContDiv span, .zf-geolocation .zf-tempContDiv span { display:block; padding-top:15px; }
        .zf-address .address_row_1 .zf-addresCols { padding-top:0; }
        .zf-address .zf-tempContDiv span label, .zf-geolocation .zf-tempContDiv span label { display:block; }
        .zf-address .zf-tempContDiv .zf-addOne, .zf-geolocation .zf-tempContDiv .zf-addOne { float:none; padding-bottom:15px; margin-right:0; padding-right:0; }
        .zf-address .zf-tempContDiv .zf-addOne input, .zf-geolocation .zf-tempContDiv .zf-addOne input { width:100%; }
        .zf-leftAlign .zf-address .zf-tempContDiv span.zf-addtwo, .zf-leftAlign .zf-geolocation .zf-tempContDiv span.zf-addtwo, .zf-rightAlign .zf-address .zf-tempContDiv span.zf-addtwo, .zf-rightAlign .zf-geolocation .zf-tempContDiv span.zf-addtwo { width:100%; }
        .zf-leftAlign .zf-address.zf-addrmedium .zf-tempContDiv span.zf-addtwo, .zf-leftAlign.zf-addrmedium .zf-geolocation .zf-tempContDiv span.zf-addtwo, .zf-rightAlign .zf-address.zf-addrmedium .zf-tempContDiv span.zf-addtwo, .zf-rightAlign.zf-addrmedium .zf-geolocation .zf-tempContDiv span.zf-addtwo { width:47%; float:left; }
        .zf-leftAlign .zf-address.zf-addrlarge .zf-tempContDiv span.zf-addtwo, .zf-leftAlign.zf-addrlarge .zf-geolocation .zf-tempContDiv span.zf-addtwo, .zf-rightAlign .zf-address.zf-addrlarge .zf-tempContDiv span.zf-addtwo, .zf-rightAlign.zf-addrlarge .zf-geolocation .zf-tempContDiv span.zf-addtwo { width:48%; float:left; }
        .zf-leftAlign .zf-address.zf-addrlarge .zf-tempContDiv span.zf-addtwo:nth-last-of-type(2), .zf-leftAlign .zf-address.zf-addrmedium .zf-tempContDiv span.zf-addtwo:nth-last-of-type(2), .zf-rightAlign .zf-address.zf-addrlarge .zf-tempContDiv span.zf-addtwo:nth-last-of-type(2), .zf-rightAlign .zf-address.zf-addrmedium .zf-tempContDiv span.zf-addtwo:nth-last-of-type(2) { padding-bottom:0; }
        .zf-address .zf-tempContDiv span.zf-addtwo:nth-child(even), .zf-geolocation .zf-tempContDiv span.zf-addtwo:nth-child(even) { padding-right:0; }
        .zf-address .zf-tempContDiv span.zf-addtwo input, .zf-geolocation .zf-tempContDiv span.zf-addtwo input { width:100%; }
        .zf-address .zf-tempContDiv span.zf-addtwo .zf-form-sBox { width:100%; }
        .zf-signContainer { margin:0; padding:0; width:100%; }
        .zf-signContainer canvas { cursor:crosshair; border:1px solid rgba(184,187,211,1); background:#fff; border-radius:5px; width:100%; height:130px; box-sizing:border-box; }
        .zf-signContainer a { font-size:14px; text-decoration:underline; display:block; color:#465475; margin-top:8px; }
        .zf-section h2 { border-bottom:1px solid #a7abb2; font-size:22px; color:#000; font-weight:500; padding-bottom:10px; }
        .zf-section p { color:#465475; margin-top:10px; font-size:15px; }
        .zf-note .zf-labelName { padding-top:7px; } .zf-templateWrapper .zf-note { overflow:hidden; }
        .zf-date .zf-tempContDiv span label { display:block; text-align:left; color:#252c3e; padding-top:8px; font-size:14px; opacity:.8; }
        .zf-subDate { margin-right:10px; } .zf-subDate label { text-align:left !important; }
        .zf-time .zf-tempContDiv span label { display:block; padding-top:8px; font-size:13px; }
        .zf-time .zf-tempContDiv .zf-form-sBox { min-width:58px; width:72px; padding:9px 20px 10px 6px; }
        .zf-time .zf-tempContDiv .zf-symbols { padding-top:12px; }
        .zf-tempContDiv input[type="checkbox"], .zf-tempContDiv input[type="radio"] { display:block; height:13px; margin:4px 0 0; padding:0; width:13px; cursor:pointer; }
        .zf-tempContDiv input[type="radio"] { -webkit-appearance:none; border:1.2px solid #47476b; border-radius:50%; width:20px; height:20px; }
        .zf-tempContDiv input[type="radio"]~label { cursor:pointer; }
        .zf-tempContDiv input[type="checkbox"] { -webkit-appearance:none; border-radius:3px; border:1.2px solid #47476b; transition:0.5s ease all; position:relative; width:20px; height:20px; }
        .zf-tempContDiv input[type="checkbox"]~label { cursor:pointer; }
        .zf-tempContDiv .zf-termsAccept input[type="checkbox"]~label { cursor:default; }
        .zf-tempContDiv input[type="radio"]:before { content:''; display:block; width:12px; height:12px; border-radius:50%; margin-top:3px; margin-left:3px; box-sizing:border-box; }
        .zf-tempContDiv input[type="checkbox"]:before { content:''; display:block; width:60%; height:60%; margin:19% auto; }
        .zf-tempContDiv input[type="radio"]:checked:before { background:rgba(46,183,159,1); }
        .zf-tempContDiv input[type="radio"]:checked { border:1.2px solid rgba(46,183,159,1); box-shadow:0px 0px 0px 0.5px rgba(46,183,159,1); outline:none; }
        .zf-tempContDiv input[type="checkbox"]:checked { border:1.2px solid rgba(46,183,159,1); box-shadow:0px 0px 2px 0px rgba(46,183,159,1); outline:none; }
        .zf-tempContDiv input[type="checkbox"]:checked:before { content:" "; display:inline-block; transform:rotate(45deg); height:10px; width:4px; border-bottom:2px solid rgba(46,183,159,1); border-right:2px solid rgba(46,183,159,1); position:absolute; top:-2px; left:6px; }
        .zf-sideBySide .zf-tempContDiv span { margin:0 4% 16px 0; padding:0; width:auto; float:left; display:block; }
        .zf-sideBySide .zf-tempContDiv span input[type="checkbox"] { display:block; min-width:20px; width:20px; height:20px; padding:0; margin-top:1px; float:left; margin-left:2px; }
        .zf-sideBySide .zf-tempContDiv span input[type="radio"] { display:block; width:20px; height:20px; margin-top:1px; padding:0; float:left; margin-left:1px; }
        .zf-sideBySide .zf-tempContDiv span label { line-height:21px; display:block; padding:0 0 0 32px; cursor:pointer; font-size:15px; color:#47476b; }
        .zf-oneColumns .zf-tempContDiv span { margin:0 0 16px 0; padding:0; width:100%; display:block; clear:both; }
        .zf-oneColumns .zf-multiAttType label~input[type="text"], .zf-twoColumns .zf-multiAttType label~input[type="text"], .zf-threeColumns .zf-multiAttType label~input[type="text"], .zf-sideBySide .zf-multiAttType label~input[type="text"] { margin-top:15px; width:100%; }
        .zf-oneColumns .zf-tempContDiv span:last-child { margin-bottom:0; }
        .zf-oneColumns .zf-tempContDiv span input[type="checkbox"] { display:block; min-width:20px; width:20px; height:20px; margin:0; padding:0; margin-top:1px; float:left; margin-left:2px; }
        .zf-oneColumns .zf-tempContDiv span input[type="radio"] { display:block; min-width:20px; width:20px; height:20px; margin-top:1px; padding:0; float:left; margin-left:1px; }
        .zf-oneColumns .zf-tempContDiv span label { line-height:21px; display:block; margin:0 0 0 32px; padding:0; font-size:15px; color:#47476b; }
        .zf-twoColumns .zf-tempContDiv span { margin:0 4% 16px 0; width:48%; float:left; display:block; }
        .zf-twoColumns .zf-tempContDiv span:nth-child(even) { margin-right:0; }
        .zf-twoColumns .zf-tempContDiv span input[type="checkbox"] { display:block; min-width:20px; width:20px; height:20px; margin:0; padding:0; margin-top:1px; float:left; margin-left:2px; }
        .zf-twoColumns .zf-tempContDiv span input[type="radio"] { display:block; min-width:20px; width:20px; height:20px; margin-top:1px; padding:0; float:left; margin-left:1px; }
        .zf-twoColumns .zf-tempContDiv span label { line-height:21px; display:block; margin:0 0 0 32px; padding:0; font-size:15px; color:#47476b; }
        .zf-threeColumns .zf-tempContDiv span { margin:0 4% 15px 0; width:30.6%; float:left; display:block; }
        .zf-threeColumns .zf-tempContDiv span:nth-child(3n) { margin-right:0; }
        .zf-threeColumns .zf-tempContDiv span input[type="checkbox"] { display:block; min-width:20px; width:20px; height:20px; padding:0; margin-top:1px; float:left; margin-left:2px; }
        .zf-threeColumns .zf-tempContDiv span input[type="radio"] { display:block; min-width:20px; width:20px; height:20px; margin-top:1px; padding:0; float:left; margin-left:1px; }
        .zf-threeColumns .zf-tempContDiv span label { line-height:21px; display:block; margin:0 0 0 32px; padding:0; font-size:15px; color:#47476b; }
        .zf-mSelect select { font-size:15px; border:1px solid rgba(184,187,211,1); overflow:auto; border-radius:4px; color:#47476b; outline:none; box-sizing:border-box; height:101px; }
        .zf-mSelect select option { padding:8px 10px; box-sizing:border-box; }
        .zf-fmFooter { margin:0; padding:10px 25px 40px 25px; text-align:center; }
        .zf-fmFooter .zf-submitColor { font-size:16px; padding:13px 38px; }
        .zf-submitColor { color:#fff; border:1px solid transparent; background:rgba(46,183,159,1); border-radius:150px; min-width:100px; transition:transform .25s cubic-bezier(0.33, 0.04, 0.63, 0.93), -webkit-transform .25s cubic-bezier(0.33, 0.04, 0.63, 0.93), -o-transform .25s cubic-bezier(0.33, 0.04, 0.63, 0.93); -webkit-transform:translate3d(0,0,0); transform:translate3d(0,0,0); cursor:pointer; }
        .zf-submitColor:hover { -webkit-transform:scale3d(1.03,1.03,1) translate3d(0,0,0) perspective(500px); transform:scale3d(1.03,1.03,1) translate3d(0,0,0) perspective(500px); }
        .zf-small .zf-tempContDiv input[type="text"], .zf-small .zf-tempContDiv textarea, .zf-small .zf-mSelect select, .zf-small .zf-tempContDiv .zf-sliderCont, .zf-small .zf-tempContDiv .zf-pdfTextArea, .zf-small .zf-signContainer { width:50%; }
        .zf-medium .zf-tempContDiv input[type="text"], .zf-medium .zf-tempContDiv textarea, .zf-medium .zf-mSelect select, .zf-medium .zf-tempContDiv .zf-sliderCont, .zf-medium .zf-tempContDiv .zf-pdfTextArea, .zf-medium .zf-signContainer { width:75%; }
        .zf-large .zf-tempContDiv input[type="text"], .zf-large .zf-tempContDiv textarea, .zf-large .zf-mSelect select, .zf-large .zf-tempContDiv .zf-sliderCont, .zf-large .zf-tempContDiv .zf-pdfTextArea, .zf-large .zf-signContainer { width:100%; }
        .signContainer canvas { width:100%; }
        .zf-small .zf-tempContDiv .zf-form-sBox { width:50%; } .zf-medium .zf-tempContDiv .zf-form-sBox { width:75%; } .zf-large .zf-tempContDiv .zf-form-sBox { width:100%; }
        .zf-name .zf-tempContDiv .zf-form-sBox { width:100%; padding:8px 10px 10px 4px; }
        .zf-namesmall .zf-nameWrapper { width:50%; } .zf-namesmall .zf-tempContDiv span { width:49%; margin-left:2%; }
        .zf-namesmall .zf-oneType .zf-salutationWrapper span { width:63%; } .zf-namesmall .zf-oneType .zf-salutationWrapper .zf-salutation { width:33%; }
        .zf-namesmall .zf-twoType .zf-salutationWrapper span { width:34%; margin-left:3%; } .zf-namesmall .zf-twoType .zf-salutationWrapper .zf-salutation { width:26%; }
        .zf-namesmall .zf-threeType .zf-nameWrapper span { width:32%; margin-left:2%; } .zf-namesmall .zf-threeType .zf-salutationWrapper span { width:25%; float:left; margin-left:2%; margin-bottom:0; } .zf-namesmall .zf-threeType .zf-salutationWrapper .zf-salutation { width:19%; }
        .zf-namesmall .zf-tempContDiv span:first-child { margin-left:0; }
        .zf-leftAlign .zf-namesmall .zf-threeType .zf-salutationWrapper span, .zf-rightAlign .zf-namesmall .zf-threeType .zf-salutationWrapper span { float:left; margin-left:2%; width:25%; }
        .zf-leftAlign .zf-namesmall .zf-threeType .zf-salutationWrapper .zf-salutation, .zf-rightAlign .zf-namesmall .zf-threeType .zf-salutationWrapper .zf-salutation { width:19%; }
        .zf-namemedium .zf-nameWrapper { width:75%; } .zf-namemedium .zf-tempContDiv span { width:49%; margin-left:2%; }
        .zf-namemedium .zf-oneType .zf-salutationWrapper span { width:73%; } .zf-namemedium .zf-oneType .zf-salutationWrapper .zf-salutation { width:25%; }
        .zf-namemedium .zf-twoType .zf-salutationWrapper span { width:38%; margin-left:2%; } .zf-namemedium .zf-twoType .zf-salutationWrapper .zf-salutation { width:20%; }
        .zf-namemedium .zf-threeType .zf-nameWrapper span { width:32%; margin-left:2%; } .zf-namemedium .zf-threeType .zf-salutationWrapper span { width:25%; margin-left:2%; } .zf-namemedium .zf-threeType .zf-salutationWrapper .zf-salutation { width:19%; }
        .zf-namemedium .zf-tempContDiv span:first-child { margin-left:0; }
        .zf-namelarge .zf-tempContDiv span { width:23.5%; margin-left:2%; margin-right:0; margin-bottom:0; }
        .zf-namelarge .zf-twoType .zf-nameWrapper span { width:49%; margin-left:2%; } .zf-namelarge .zf-threeType .zf-nameWrapper span { width:32%; margin-left:2%; }
        .zf-namelarge .zf-twoType .zf-salutationWrapper span { width:40%; margin-left:2%; } .zf-namelarge .zf-twoType .zf-salutationWrapper .zf-salutation { width:16%; }
        .zf-namelarge .zf-threeType .zf-salutationWrapper span { width:26%; margin-left:2%; } .zf-namelarge .zf-threeType .zf-salutationWrapper .zf-salutation { width:19.8%; }
        .zf-namelarge .zf-oneType .zf-salutationWrapper span { width:73%; margin-left:2%; } .zf-namelarge .zf-oneType .zf-salutationWrapper .zf-salutation { width:25%; }
        .zf-namelarge .zf-tempContDiv span:first-child { margin-left:0 !important; }
        .zf-csmall .zf-tempContDiv input[type="text"] { width:100%; } .zf-cmedium .zf-tempContDiv input[type="text"] { width:100%; } .zf-clarge .zf-tempContDiv input[type="text"] { width:100%; }
        .zf-nsmall .zf-tempContDiv input[type="text"] { width:50%; } .zf-nmedium .zf-tempContDiv input[type="text"] { width:75%; } .zf-nlarge .zf-tempContDiv input[type="text"] { width:100%; }
        .zf-signSmall .zf-tempContDiv .zf-signContainer .zf-signArea { width:49%; } .zf-signMedium .zf-tempContDiv .zf-signContainer .zf-signArea { width:60%; } .zf-signLarge .zf-tempContDiv .zf-signContainer .zf-signArea { width:74%; }
        .zf-addrsmall .zf-tempContDiv .zf-addrCont { width:50%; } .zf-addrmedium .zf-tempContDiv .zf-addrCont { width:75%; } .zf-addrlarge .zf-tempContDiv .zf-addrCont { width:100%; }
        .zf-leftAlign .zf-currency .zf-tempContDiv input[type="text"], .zf-rightAlign .zf-currency .zf-tempContDiv input[type="text"] { float:left; }
        .zf-leftAlign .zf-currency.zf-clarge .zf-tempContDiv input[type="text"], .zf-rightAlign .zf-currency.zf-clarge .zf-tempContDiv input[type="text"] { float:left; }
        .zf-leftAlign .zf-currency.zf-clarge .zf-tempContDiv input[type="text"]~span, .zf-rightAlign .zf-currency.zf-clarge .zf-tempContDiv input[type="text"]~span { margin-top:9px; }
        .zf-topAlign .zf-tempFrmWrapper .zf-labelName { padding-bottom:10px; display:block; }
        .zf-topAlign .zf-threeColumns .zf-labelName, .zf-topAlign .zf-twoColumns .zf-labelName, .zf-topAlign .zf-oneColumns .zf-labelName, .zf-topAlign .zf-sideBySide .zf-labelName { padding-bottom:8px; }
        .zf-leftAlign { display:block; } .zf-leftAlign .zf-tempFrmWrapper .zf-labelName { float:left; width:30%; line-height:20px; padding-right:30px; box-sizing:border-box; }
        .zf-leftAlign .zf-tempFrmWrapper .zf-tempContDiv { float:right; width:70%; }
        .zf-leftAlign .zf-slider .zf-tempContDiv { margin-top:6px; }
        .zf-leftAlign .zf-decesion .zf-tempContDiv, .zf-rightAlign .zf-decesion .zf-tempContDiv { margin-left:0 !important; }
        .zf-rightAlign { display:block; } .zf-rightAlign .zf-tempFrmWrapper .zf-labelName { float:left; width:30%; line-height:20px; text-align:right; padding-right:30px; box-sizing:border-box; }
        .zf-rightAlign .zf-tempFrmWrapper .zf-tempContDiv { float:right; width:70%; }
        .zf-matrixTable { font-size:13px; overflow-x:auto; padding-bottom:15px !important; } .zf-matrixTable table th, .zf-matrixTable table td { padding:10px; }
        .zf-matrixTable thead th, .zf-matrixTable table td { text-align:center; } .zf-matrixTable table td input[type="radio"], .zf-matrixTable table td input[type="checkbox"] { display:inline-block; }
        .zf-matrixTable tbody th { font-weight:normal; font-size:16px; text-align:left; padding:12px 10px; color:#252c3e; width:218px; box-sizing:border-box; }
        .zf-matrixTable thead th { word-break:normal; font-weight:normal; font-size:16px; padding:12px 10px; color:#252c3e; text-align:center; }
        .zf-termsContainer { margin:0; padding:0; } .zf-termsContainer .zf-termsMsg { border:1px solid #252c3e; max-height:250px; overflow-y:auto; padding:12px 10px 12px; margin-bottom:12px; border-radius:4px; min-height:70px; font-size:13px; }
        .zf-termsContainer .zf-termScrollRemove { border:1px solid #252c3e; overflow-y:auto; padding:12px 10px 12px; margin-bottom:12px; border-radius:4px; min-height:70px; font-size:13px; }
        .zf-termsAccept { margin-top:0 !important; } .zf-termsAccept input[type="checkbox"] { margin-top:0 !important; float:left; }
        .zf-termsAccept label { margin-left:30px; font-size:15px; float:none; display:block; color:#252c3e; }
        .zf-termsWrapper .zf-tempContDiv { margin-left:0 !important; } .zf-termsWrapper .zf-labelName { width:100% !important; text-align:left !important; padding-bottom:8px !important; }
        .zf-medium .zf-phwrapper { width:75%; } .zf-phwrapper.zf-phNumber span { width:100% !important; }
        .zf-phwrapper span:first-child { margin-left:0; width:22%; } .zf-phwrapper label { display:block; color:#252c3e; font-size:13px; padding-top:8px; opacity:0.8; }
        .zf-medium .zf-phonefld input[type="text"], .zf-small .zf-phonefld input[type="text"] { width:100%; }
        .zf-small .zf-phwrapper { width:50%; } .zf-tempFrmWrapper.zf-phone span { width:30%; margin:inherit; }
        .zf-tempFrmWrapper .zfPhoneUSA span { width:22.3%; position:relative; } .zf-tempFrmWrapper.zf-phone span input[type="text"] { width:100%; }
        .zfMultiColGrid .zf-tempFrmWrapper.zf-phone .zf-symbols { display:none; } .zf-tempFrmWrapper.zf-phone .zfPhoneUSA { display:flex; }
        .zfMultiColGrid .zf-tempFrmWrapper.zf-phone .zfPhoneUSA span, .zfoneColumn .zf-tempFrmWrapper.zf-phone .zfPhoneUSA span { width:30%; flex:1 1 auto; margin-left:8px; }
        .zf-tempFrmWrapper.zf-phone .zfPhoneUSA span:first-of-type { margin-left:0; } .zf-phwrapper span { float:left; width:76%; margin-left:2%; }
        .zf-descFld a { text-decoration:underline; color:#252c3e; } .zf-descFld em { font-style:italic; } .zf-descFld b { font-weight:bold; } .zf-descFld i { font-style:italic; } .zf-descFld u { text-decoration:underline; }
        .zf-descFld ul { margin:auto; } .zf-descFld ul { list-style:disc; } .zf-descFld ol { list-style:decimal; } .zf-descFld ul, .zf-descFld ol { margin:10px 0; padding-left:20px; }
        .zf-descFld ol.code { list-style-position:outside; list-style-type:decimal; padding:0 30px; } .zf-descFld ol.code li { background-color:#F5F5F5; border-left:2px solid #CCCCCC; margin:1px 0; padding:2px; }
        .zf-descFld blockquote.zquote { border-left:3px solid #EFEFEF; padding-left:35px; } .zf-descFld blockquote.zquote span.txt { -moz-user-focus:ignore; -moz-user-input:disabled; -moz-user-select:none; color:#058BC2; float:left; font:bold 50px Arial, Helvetica, sans-serif; margin:-10px 0 0 -30px; }
        .zf-descFld blockquote.block_quote { background:url("../images/newQuote.gif") no-repeat scroll 12px 10px rgba(0,0,0,0); border-left:3px solid #EFEFEF; font:13px/20px georgia, Arial, verdana, Helvetica, sans-serif; margin:15px 3px 15px 15px; padding:10px 10px 10px 40px; }
        .zf-descFld body { font-family:Arial, Helvetica, sans-serif; font-size:13px; margin:8px; }
        .note .noteCont { overflow:hidden; } .note .zf-descFld { overflow:hidden; font-size:13px; } .zf-descFld img { width:auto; }
        .zf-date .zf-tempContDiv input[type="text"] { width:340px; } .zf-leftAlign .zf-date .zf-tempContDiv input[type="text"], .zf-rightAlign .zf-date .zf-tempContDiv input[type="text"] { width:238px; }
        .zf-leftAlign .zf-date.zf-time .zf-tempContDiv input[type="text"], .zf-rightAlign .zf-date.zf-time .zf-tempContDiv input[type="text"] { width:223px; }
        .zf-decesion .zf-tempContDiv input[type="checkbox"] { margin-top:0; } .zf-tempFrmWrapper.zf-decesion .zf-labelName { margin-left:28px !important; }
        .zf-leftAlign .zf-tempFrmWrapper .zf-matrixTable, .zf-rightAlign .zf-tempFrmWrapper .zf-matrixTable { width:100%; clear:both; float:none; padding-top:10px; }
        .zf-leftAlign .zf-tempFrmWrapper.zf-termsandCond .zf-tempContDiv, .zf-rightAlign .zf-tempFrmWrapper.zf-termsandCond .zf-tempContDiv { float:none; width:100%; clear:both; padding-top:10px; }
        .zf-leftAlign .zf-tempFrmWrapper.zf-termsandCond .zf-labelName, .zf-rightAlign .zf-tempFrmWrapper.zf-termsandCond .zf-labelName { width:100%; float:none; text-align:left; display:block; }
        .zf-leftAlign .zf-namesmall .zf-threeType .zf-salutationWrapper span:first-of-type, .zf-rightAlign .zf-namesmall .zf-threeType .zf-salutationWrapper span:first-of-type { margin-left:0; }
        .zf-rightAlign .zf-tempFrmWrapper.zf-matrixTable .zf-labelName { text-align:left; }
        .zfAddressTwoCol { display:flex; justify-content:space-between; flex-wrap:wrap; gap:0 16px; } .zf-address .zf-tempContDiv .zfAddressTwoCol .zf-addOne { width:47%; }
        .zf-addrCont .zfAddressTwoCol~.zfAddressTwoCol span.zf-addtwo:last-of-type { padding-bottom:0; }
        .zf-csmall.zf-currency .zf-tempContDiv>div { width:50%; display:inline-flex; } .zf-cmedium.zf-currency .zf-tempContDiv>div { width:75%; display:inline-flex; } .zf-clarge.zf-currency .zf-tempContDiv>div { width:100%; display:inline-flex; }
        .zf-nameWrapper.zf-salutationWrapper { display:flex; } .zf-leftAlign .zf-tempFrmWrapper .zf-termsandCond { float:none; width:100%; }
        .zf-tempContDiv span input[type="checkbox"]:focus { border:1.2px solid rgb(46,183,159,1); box-shadow:0px 0px 2px 0px rgb(46,183,159,1); }
        .zf-multiAttType.fullColWrap, .zf-twoColumns .zf-tempContDiv .zf-multiAttType.fullColWrap, .zf-threeColumns .zf-tempContDiv .zf-multiAttType.fullColWrap, .zf-sideBySide .zf-tempContDiv .zf-multiAttType.fullColWrap { width:100%; }
        .zf-address.zf-tempContDiv .zf-form-sBox, .zf-address.zf-tempContDiv input[type="text"] { width:100%; }
        .address_row_3, .address_row_4 { display:flex; justify-content:space-between; flex-wrap:wrap; gap:0 16px; }
        .address_row_3 span.zf-addresCols, .address_row_4 span.zf-addresCols { width:47%; flex:1 1 auto; }
        .address_row_1 .zf-addresCols, .address_row_2 .zf-addresCols, .address_row_3 .zf-addresCols, .address_row_4 .zf-addresCols { width:100%; }
        .zf-tempFrmWrapper.gridName .zf-twoType .zf-nameWrapper span { width:100%; margin-left:0; margin-bottom:15px; }
        .zf-tempFrmWrapper.gridName .zf-threeType .zf-salutationWrapper .zf-salutation { width:100%; margin-left:0; }
        .zf-tempFrmWrapper.gridName .zf-threeType .zf-salutationWrapper span { width:100%; margin-left:0; margin-bottom:15px; }
        .zf-tempFrmWrapper.gridName .zf-nameWrapper.zf-salutationWrapper { display:block; }
        .zf-templateWrapper .zf-note table td { word-break:break-all; } .zf-leftAlign .zf-tempFrmWrapper.zf-matrixTable .zf-labelName { width:100%; }

        .zfgrid_Wrapper { padding:0 40px 0 40px; margin:0; box-sizing:border-box; } .zfgridLabelCont { padding:12px 0; } .zfgridLabelCont label { font-size:22px; color:#252c3e; } .zfgridLabelCont span { font-size:14px; color:#465474; padding-top:8px; display:block; }
        .zftwoColumn, .zfthreeColumn { display:flex; flex-wrap:nowrap; gap:32px; } .zfoneColumn .zfCol { max-width:100%; } .zftwoColumn .zfCol { flex:1 1; overflow-x:hidden; max-width:50%; } .zfthreeColumn .zfCol { flex:1 1; overflow-x:hidden; max-width:33%; }
        .zfgrid_Wrapper .zf-tempFrmWrapper { padding-left:0; padding-right:0; }
        .zfMultiColGrid .zf-medium .zf-tempContDiv input[type="text"], .zfMultiColGrid .zf-medium .zf-tempContDiv textarea, .zfMultiColGrid .zf-medium .zf-mSelect select, .zfMultiColGrid .zf-medium .zf-tempContDiv .zf-sliderCont, .zfMultiColGrid .zf-medium .zf-tempContDiv .zf-pdfTextArea, .zfMultiColGrid .zf-medium .zf-signContainer { width:100%; }
        .zfMultiColGrid .zf-small .zf-tempContDiv input[type="text"], .zfMultiColGrid .zf-small .zf-tempContDiv textarea, .zfMultiColGrid .zf-small .zf-mSelect select, .zfMultiColGrid .zf-small .zf-tempContDiv .zf-sliderCont, .zfMultiColGrid .zf-small .zf-tempContDiv .zf-pdfTextArea, .zfMultiColGrid .zf-small .zf-signContainer { width:100%; }
        .zfMultiColGrid .zf-medium .zf-phwrapper { width:100%; } .zfMultiColGrid .zf-small .zf-phwrapper { width:100%; }
        .zfMultiColGrid .zf-tempFrmWrapper.zf-phone span, .zfoneColumn .zf-tempFrmWrapper.zf-phone span { width:30.6%; } .zfoneColumn .zf-tempFrmWrapper.zf-phone span.zf-symbols { display:none; }
        .zfMultiColGrid .zf-phone .zf-tempContDiv .zf-symbols { margin:9px 1%; width:2%; } .zfMultiColGrid .zf-cmedium.zf-currency .zf-tempContDiv>div { width:100%; } .zfMultiColGrid .zf-csmall.zf-currency .zf-tempContDiv>div { width:100%; }
        .zfMultiColGrid .zf-medium .zf-tempContDiv .zf-form-sBox, .zfMultiColGrid .zf-small .zf-tempContDiv .zf-form-sBox { width:100%; }
        .zfMultiColGrid .zf-date .zf-tempContDiv input[type="text"], .zfMultiColGrid .zf-date.zf-time .zf-tempContDiv input[type="text"] { width:100%; }
        .zfMultiColGrid .zf-tempContDiv input[type="file"] { width:100%; } .zfMultiColGrid .zf-namemedium .zf-nameWrapper, .zfMultiColGrid .zf-namesmall .zf-nameWrapper { width:100%; }
        .zfMultiColGrid .zf-addrmedium .zf-tempContDiv .zf-addrCont, .zfMultiColGrid .zf-addrsmall .zf-tempContDiv .zf-addrCont { width:100%; }
        .zftwoColumn .address_row_3, .zftwoColumn .address_row_4, .zfthreeColumn .address_row_3, .zfthreeColumn .address_row_4 { flex-wrap:nowrap; gap:0 8px; }
        .zftwoColumn .gridAddress .address_row_3, .zftwoColumn .gridAddress .address_row_4, .zfthreeColumn .gridAddress .address_row_3, .zfthreeColumn .gridAddress .address_row_4 { flex-wrap:wrap; width:100%; }
        .gridAddress .address_row_3 .zf-addresCols, .gridAddress .address_row_4 .zf-addresCols { width:100%; }
        .zfgrid_Wrapper .zf-time .zf-tempContDiv .zf-symbols { display:none; } .zfMultiColGrid .zf-time .zf-tempContDiv span { width:32%; margin-right:2%; } .zfMultiColGrid .zf-time .zf-tempContDiv span:last-of-type { margin-right:0; }
        .zfMultiColGrid .zf-time .zf-tempContDiv .zf-form-sBox { min-width:auto; width:100%; } .zfMultiColGrid .zf-date.zf-time .zf-tempContDiv .zf-subDate { width:25%; } .zfMultiColGrid .zf-date.zf-time .zf-tempContDiv span { width:23%; }
        .zfgrid_Wrapper .zf-oneColumns .zf-tempContDiv span label, .zfgrid_Wrapper .zf-twoColumns .zf-tempContDiv span label, .zfgrid_Wrapper .zf-threeColumns .zf-tempContDiv span label, .zfgrid_Wrapper .zf-sideBySide .zf-tempContDiv span label { word-break:break-all; }
        .zf-leftAlign .zfMultiColGrid .zf-tempFrmWrapper .zf-labelName, .zf-rightAlign .zfMultiColGrid .zf-tempFrmWrapper .zf-labelName { float:none; display:block; width:100%; padding-bottom:10px; text-align:left; }
        .zf-leftAlign .zfMultiColGrid .zf-tempFrmWrapper .zf-tempContDiv, .zf-rightAlign .zfMultiColGrid .zf-tempFrmWrapper .zf-tempContDiv { float:none; width:100%; }
        .zfInstrucTop .zf-instruction { padding:0px 0px 10px 0px; } .zfInstrucTop .zf-decesion .zf-instruction { padding:10px 0px 0px 0px; }
        .zfInstrucTop.zf-topAlign .zf-tempFrmWrapper .zf-labelName { padding-bottom:6px; display:block; }
        .zfoneColumn .zf-time .zf-tempContDiv { width:285px; } .zf-leftAlign .zfoneColumn .zf-time .zf-tempContDiv, .zf-rightAlign .zfoneColumn .zf-time .zf-tempContDiv, .zf-leftAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv, .zf-rightAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv { width:70%; }
        .zfoneColumn .zf-time .zf-tempContDiv span { width:25%; margin-right:14px; } .zf-leftAlign .zfoneColumn .zf-time .zf-tempContDiv span, .zf-rightAlign .zfoneColumn .zf-time .zf-tempContDiv span { width:72px; }
        .zfoneColumn .zf-time.zf-date .zf-tempContDiv { width:100%; } .zfoneColumn .zf-time.zf-date .zf-tempContDiv span { width:72px; margin-right:14px; } .zfoneColumn .zf-time.zf-date .zf-tempContDiv .zf-subDate { width:340px; }
        .zf-leftAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv .zf-subDate, .zf-rightAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv .zf-subDate { width:223px; }
        .zf-leftAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv span, .zf-rightAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv span { margin-right:12px; } .zf-leftAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv span:last-of-type, .zf-rightAlign .zfoneColumn .zf-time.zf-date .zf-tempContDiv span:last-of-type { margin-right:0; }

        .zf-divider { border-top-color:#b8bbd3 !important; margin:0 auto; }
        .zf-small .zf-divider { width:50%; } .zf-medium .zf-divider { width:75%; } .zf-large .zf-divider { width:100%; }
        .zf-divider.line-One { border-top:1px; } .zf-divider.line-Two { border-top:2px; } .zf-divider.line-Three { border-top:3px; } .zf-divider.line-Four { border-top:4px; } .zf-divider.line-Five { border-top:5px; } .zf-divider.line-Six { border-top:6px; } .zf-divider.line-Seven { border-top:7px; } .zf-divider.line-Eight { border-top:8px; } .zf-divider.line-Nine { border-top:9px; } .zf-divider.line-Ten { border-top:10px; }
        .zf-divider.solidType { border-top-style:solid; } .zf-divider.dashedType { border-top-style:dashed; } .zf-divider.dottedType { border-top-style:dotted; } .zf-divider.doubleType { border-top-style:double; }
        .layout3 .dividerContainer { background:transparent; box-shadow:none; border:none; padding:12px 0; }
        .zf-spacer { display:block; width:100%; margin:0 auto; background:white; }

        .zf-regexFldCont { display:flex; flex-wrap:wrap; align-items:center; gap:12px; } .zf-regexTagCont { display:flex; align-items:center; align-self:stretch; gap:12px; max-width:100%; }
        .zf-regexTag { background:#f5f5f5; border:1px solid #d0d0d0; color:#777; font-size:13px; border-radius:2px; padding:9px 10px; align-self:stretch; }
        .zf-tempFrmWrapper.fieldModel_4 .zf-regexTag { border-radius:4px; } .zf-regexSymbol { color:#555; font-size:13px; flex-shrink:0; } .zf-regexSymbol.regUnderscore { align-self:flex-end; }
        .zf-regexInput { display:flex; align-items:center; gap:12px; max-width:100%; width:20%; flex-grow:1; min-width:120px; }
        .zf-regexInput input[type="text"] { width:100%; box-sizing:border-box; padding:7px 10px; border:1px solid #d0d0d0; border-radius:2px; font-size:13px; } .zf-regexInput input[type="text"]:focus { outline:none; border-color:#4285f4; }

        @media (max-width:700px) {
            .zftwoColumn, .zfthreeColumn { flex-wrap:wrap; } .zftwoColumn .zfCol, .zfthreeColumn .zfCol { max-width:100%; flex:1 1 100%; }
            .zf-tempHeadContBdr { padding:20px 16px; } .zf-tempHeadContBdr .zf-frmTitle { font-size:24px; }
            .zf-tempFrmWrapper { padding:12px 16px; } .zfgrid_Wrapper { padding:0 16px; }
            .zf-leftAlign .zf-tempFrmWrapper .zf-labelName, .zf-rightAlign .zf-tempFrmWrapper .zf-labelName { float:none; width:100%; padding-right:0; text-align:left; }
            .zf-leftAlign .zf-tempFrmWrapper .zf-tempContDiv, .zf-rightAlign .zf-tempFrmWrapper .zf-tempContDiv { float:none; width:100%; }
            .zf-name .zf-tempContDiv span { width:48% !important; margin-left:2% !important; }
            .zf-namelarge .zf-tempContDiv span { width:48% !important; margin-left:2% !important; }
            .zf-namelarge .zf-twoType .zf-salutationWrapper span { width:48% !important; margin-left:2% !important; } .zf-namelarge .zf-twoType .zf-salutationWrapper .zf-salutation { width:25% !important; }
            .zf-tempContDiv input[type="file"] { width:100%; }
        }
        @media (max-width:480px) {
            .zf-templateWidth { padding:16px 8px; } .zf-tempFrmWrapper { padding:10px 12px; } .zfgrid_Wrapper { padding:0 12px; }
            .zf-tempHeadContBdr { padding:16px 12px; } .zf-tempHeadContBdr .zf-frmTitle { font-size:20px; }
            .zf-name .zf-tempContDiv span { width:100% !important; margin-left:0 !important; margin-bottom:8px; }
            .zf-namelarge .zf-tempContDiv span { width:100% !important; margin-left:0 !important; margin-bottom:8px; }
            .zf-namelarge .zf-twoType .zf-salutationWrapper span { width:100% !important; margin-left:0 !important; } .zf-namelarge .zf-twoType .zf-salutationWrapper .zf-salutation { width:100% !important; }
            .zf-phwrapper span { width:100% !important; margin-left:0 !important; } .zf-phwrapper.zf-phNumber span { width:100% !important; }
            .zf-fmFooter .zf-submitColor { padding:12px 24px; font-size:15px; min-width:80px; }
        }
        .zf-hidden { display:none !important; } .zf-visible { display:block !important; }

        /* Custom styles */
        .nav-tabs .nav-link { color: #252c3e; font-weight: 500; }
        .nav-tabs .nav-link.active { color: #0d6efd; border-bottom: 3px solid #0d6efd; }
        .player-id { font-family: monospace; background: #e9ecef; padding: 2px 8px; border-radius: 4px; }
        .status-pending { color: #ffc107; } .status-approved { color: #198754; } .status-rejected { color: #dc3545; }
        .ad-banner img { max-height: 80px; width: auto; }
        .ranking-row:hover { background: #f8f9fa; }
        .top-rank { background: #fff3cd; }
        .card-dashboard { border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: none; }
        .card-dashboard .card-body { padding: 20px; }
        .tab-content { padding-top: 20px; }
        .view-tab { display: none; }
        .view-tab.active { display: block; }
        .rank-view { display: none; }
        .rank-view.active { display: block; }

        /* Loading overlay */
        #loadingOverlay {
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255,255,255,0.7);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
        }

        /* Team form styles (from file 7, without emojis) */
        .form-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px 20px; }
        .form-grid .full-width { grid-column:1 / -1; }
        .form-group { display:flex; flex-direction:column; gap:4px; }
        .form-group label { font-size:0.85rem; font-weight:600; color:#5a5a5a; }
        .form-group label .required-star { color:#d63031; margin-left:2px; }
        .form-group input, .form-group textarea, .form-group select {
            width:100%; padding:10px 13px; border:1.5px solid #dce0e5; border-radius:6px;
            font-size:0.95rem; font-family:inherit; background:#f9fafb; transition:all 0.2s;
        }
        .form-group input:focus, .form-group textarea:focus, .form-group select:focus {
            border-color:#2c5f9e; background:#eef3ff; box-shadow:0 0 0 3px rgba(44,95,158,0.08);
        }
        .form-group input[type="file"] { padding:8px 10px; font-size:0.85rem; background:#fff; border-style:dashed; cursor:pointer; }
        .form-group input[type="file"]:hover { border-color:#2c5f9e; background:#fdfdff; }
        .table-wrapper { border-radius:6px; border:1px solid #dce0e5; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
        .player-table { width:100%; border-collapse:collapse; font-size:0.9rem; background:#fff; }
        .player-table thead th { background:#f1f4f9; color:#1a3c6e; font-weight:700; font-size:0.8rem; text-transform:uppercase; padding:12px 10px; border-bottom:2px solid #dce0e7; text-align:center; }
        .player-table tbody td { padding:8px 6px; border-bottom:1px solid #f0f2f5; text-align:center; vertical-align:middle; }
        .player-table tbody tr.sub-header td { background:#fffef7; font-weight:700; color:#8b6f10; font-size:0.82rem; text-transform:uppercase; padding:10px; border-bottom:2px solid #f0e6c0; }
        .player-table input { width:100%; padding:7px 8px; border:1.5px solid #e8ecf2; border-radius:4px; font-size:0.88rem; background:#fdfdfe; text-align:center; transition:all 0.2s; }
        .player-table input:focus { border-color:#2c5f9e; background:#fff; box-shadow:0 0 0 2px rgba(44,95,158,0.06); outline:none; }
        .player-table .num-cell { font-weight:700; color:#1a3c6e; width:40px; }
        .pledge-box { background:#fdfcf7; border:2px dashed #e0d5a8; border-radius:10px; padding:20px 22px; margin-top:8px; }
        .pledge-box .pledge-text { font-size:0.9rem; color:#5c4f1f; font-weight:500; margin-bottom:16px; line-height:1.7; text-align:center; font-style:italic; }
        .pledge-grid { display:grid; grid-template-columns:2fr 1.5fr; gap:14px 20px; align-items:end; }
        .attachments-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px 20px; }
        .attachment-card { background:#fafbfc; border:1.5px solid #e8ecf2; border-radius:6px; padding:14px 16px; display:flex; flex-direction:column; gap:6px; }
        .attachment-card .att-label { font-weight:700; font-size:0.85rem; color:#1a3c6e; text-align:center; }
        .att-badge { display:inline-block; font-size:0.68rem; font-weight:600; padding:2px 8px; border-radius:12px; text-align:center; align-self:center; }
        .att-badge.required { background:#ffeaea; color:#c0392b; }
        .att-badge.optional { background:#eef6ee; color:#3a7d44; }
        .attachment-card input[type="file"] { width:100%; padding:7px 8px; font-size:0.8rem; border:1.5px dashed #d0d5dc; border-radius:4px; background:#fff; cursor:pointer; }
        .attachment-card input[type="file"]:hover { border-color:#2c5f9e; background:#fdfdff; }
        .signature-section { margin-top:24px; border:2px solid #dce0e5; border-radius:10px; padding:18px; background:#fafbfc; }
        .signature-section .sig-title { font-weight:700; color:#1a3c6e; margin-bottom:10px; font-size:0.95rem; }
        .canvas-container { position:relative; width:100%; max-width:500px; margin:0 auto; border:1.5px solid #ccc; border-radius:6px; background:#fff; touch-action:none; }
        canvas { display:block; width:100%; height:auto; background:white; cursor:crosshair; border-radius:4px; }
        .sig-buttons { display:flex; justify-content:center; gap:10px; margin-top:10px; flex-wrap:wrap; }
        .btn-clear-sig { background:#fff; border:1.5px solid #d0d5dc; padding:8px 20px; border-radius:20px; font-weight:600; font-size:0.85rem; cursor:pointer; transition:all 0.2s; font-family:inherit; color:#555; }
        .btn-clear-sig:hover { background:#f0f0f0; border-color:#b0b8c3; }
        .stamp-area { display:flex; align-items:center; justify-content:flex-end; gap:16px; margin-top:14px; padding-top:10px; border-top:1px dotted #d5d5d5; flex-wrap:wrap; }
        .stamp-placeholder { width:90px; height:90px; border:2px dashed #c5c5c5; border-radius:8px; display:flex; align-items:center; justify-content:center; font-size:0.7rem; color:#aaa; text-align:center; background:#fafafa; flex-shrink:0; }
        .btn-row { display:flex; gap:14px; justify-content:center; flex-wrap:wrap; margin-top:28px; padding-top:16px; border-top:1px solid #eee; }
        .btn { padding:12px 28px; border-radius:30px; font-weight:700; font-size:0.95rem; font-family:inherit; letter-spacing:0.4px; cursor:pointer; border:none; transition:all 0.25s ease; display:flex; align-items:center; gap:8px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
        .btn-submit { background:#1a3c6e; color:#fff; }
        .btn-submit:hover { background:#0f2d54; box-shadow:0 4px 12px rgba(0,0,0,0.15); transform:translateY(-1px); }
        .btn-reset { background:#fff; color:#2c2c2c; border:1.5px solid #d0d5dc; }
        .btn-reset:hover { background:#f5f6f8; border-color:#b0b8c3; }
        .btn-print { background:#fff; color:#1a3c6e; border:1.5px solid #2c5f9e; }
        .btn-print:hover { background:#eef3fc; }
        .form-footer-note { text-align:center; font-size:0.78rem; color:#999; margin-top:18px; }
        @media (max-width:640px) { .form-grid, .form-grid.three-col { grid-template-columns:1fr; } .pledge-grid { grid-template-columns:1fr; } .attachments-grid { grid-template-columns:1fr; } }
        .section-title { display:flex; align-items:center; gap:12px; font-size:1.15rem; font-weight:700; color:#1a3c6e; margin:28px 0 16px; padding-bottom:10px; border-bottom:2px solid #eef1f7; flex-wrap:wrap; }
        .section-title .section-badge { font-size:0.7rem; font-weight:600; background:#e8a817; color:#1a3c6e; padding:3px 10px; border-radius:20px; text-transform:uppercase; margin-left:auto; }
        .section-title .section-icon { display:flex; align-items:center; justify-content:center; width:36px; height:36px; border-radius:8px; background:#eef3fc; color:#1a3c6e; font-size:1.1rem; flex-shrink:0; }
    </style>
</head>
<body class="zf-backgroundBg">

<!-- Loading Overlay -->
<div id="loadingOverlay">
    <div class="spinner-border text-primary" role="status" style="width: 3rem; height: 3rem;">
        <span class="visually-hidden">Loading...</span>
    </div>
</div>

<!-- ============================================= -->
<!-- NAVIGATION + TABS -->
<!-- ============================================= -->
<nav class="navbar navbar-expand-lg navbar-dark bg-primary">
    <div class="container">
        <a class="navbar-brand" href="#"><i class="fas fa-chess-queen"></i> AFMIND Draughts</a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navTabs">
            <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navTabs">
            <ul class="navbar-nav me-auto nav-tabs border-0" id="mainTabs" role="tablist">
                <li class="nav-item"><a class="nav-link active" data-view="register" href="#register">Register</a></li>
                <li class="nav-item"><a class="nav-link" data-view="teams" href="#teams">Teams</a></li>
                <li class="nav-item"><a class="nav-link" data-view="staffreg" href="#staffreg">Staff Reg</a></li>
                <li class="nav-item"><a class="nav-link" data-view="ranking" href="#ranking">Ranking</a></li>
                <li class="nav-item"><a class="nav-link" data-view="staff" href="#staff">Results</a></li>
                <li class="nav-item"><a class="nav-link" data-view="admin" href="#admin">Admin</a></li>
            </ul>
            <div class="d-flex">
                <span id="userDisplay" class="navbar-text me-2 text-white"></span>
                <button id="loginBtn" class="btn btn-outline-light btn-sm" onclick="openPinLogin()">Login</button>
                <button id="logoutBtn" class="btn btn-outline-light btn-sm d-none" onclick="clearAccess()">Logout</button>
            </div>
        </div>
    </div>
</nav>

<!-- ============================================= -->
<!-- MAIN CONTENT -->
<!-- ============================================= -->
<div class="container mt-3">
    <div id="adsBanner" class="row mb-3"></div>

    <div class="tab-content">
        <!-- ============================= -->
        <!-- TAB: REGISTER (PLAYER) -->
        <!-- ============================= -->
        <div class="view-tab active" id="register" role="tabpanel">
            <div class="zf-templateWidth" style="padding-top:0;">
                <form id="afmindForm" name="form" method="POST" enctype="multipart/form-data">
                    <input type="hidden" name="zf_referrer_name" value="">
                    <input type="hidden" name="zf_redirect_url" value="">
                    <input type="hidden" name="zc_gad" value="">

                    <div class="zf-templateWrapper">
                        <ul class="zf-tempHeadBdr">
                            <li class="zf-tempHeadContBdr">
                                <h2 class="zf-frmTitle"><em>AFMIND DRAUGHTS PLAYER REGISTRATION FORM</em></h2>
                                <p class="zf-frmDesc"></p>
                                <div class="zf-clearBoth"></div>
                            </li>
                        </ul>

                        <div class="zf-subContWrap zf-topAlign">
                            <ul>
                                <!-- PICHA YA MSHIRIKI -->
                                <div class="zfgrid_Wrapper">
                                    <div class="zftwoColumn zfMultiColGrid">
                                        <div class="zfCol"></div>
                                        <div class="zfCol">
                                            <div class="zf-tempFrmWrapper">
                                                <label class="zf-labelName">PICHA YA MSHIRIKI <em class="zf-important">*</em></label>
                                                <div class="zf-tempContDiv">
                                                    <input type="file" accept="image/*" name="ImageUpload2" checktype="c1" />
                                                    <p id="ImageUpload2_error" class="zf-errorMessage" style="display:none;">Choose any file for this field.</p>
                                                </div>
                                                <div class="zf-clearBoth"></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Full Name -->
                                <div class="zf-tempFrmWrapper zf-name zf-namelarge">
                                    <label class="zf-labelName">Full name ( Andika Majina matatu) <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv zf-twoType">
                                        <div class="zf-nameWrapper zf-salutationWrapper">
                                            <span class="zf-salutation">
                                                <select class="zf-form-sBox" name="Name_Salutation" fieldType="7">
                                                    <option value="-Select-">-Select-</option>
                                                    <option value="Mr.">Mr.</option>
                                                    <option value="Mrs.">Mrs.</option>
                                                    <option value="Ms.">Ms.</option>
                                                    <option value="Prof">Prof</option>
                                                    <option value="Eng">Eng</option>
                                                    <option value="Dr">Dr</option>
                                                    <option value="CPA">CPA</option>
                                                </select>
                                                <label>Title</label>
                                            </span>
                                            <span>
                                                <input type="text" maxlength="255" name="Name_Last" fieldType="7" placeholder="" />
                                                <label>Full Name / Jina Kamili</label>
                                            </span>
                                            <span>
                                                <input type="text" maxlength="255" name="Name_First" fieldType="7" placeholder="" />
                                                <label>Nickname / Jina maarufu (Utani)</label>
                                            </span>
                                            <div class="zf-clearBoth"></div>
                                        </div>
                                        <p id="Name_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Gender -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Gender / Jinsia</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown8" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Male / Mwanaume">Male / Mwanaume</option>
                                            <option value="Female / Mwanamke">Female / Mwanamke</option>
                                        </select>
                                        <p id="Dropdown8_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Nationality -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Nationality / Utaifa <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown9" id="nationalitySelect" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Tanzania">Tanzania</option>
                                            <option value="Kenya">Kenya</option>
                                            <option value="Uganda">Uganda</option>
                                            <option value="Rwanda">Rwanda</option>
                                            <option value="Burundi">Burundi</option>
                                            <option value="South Sudan">South Sudan</option>
                                        </select>
                                        <p id="Dropdown9_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- ADDRESS -->
                                <div id="Address" class="zf-tempFrmWrapper zf-address zf-addrlarge">
                                    <label class="zf-labelName">ADDRESS / MAKAZI</label>
                                    <div class="zf-tempContDiv zf-address">
                                        <div class="zf-addrCont">
                                            <div class="address_row_1">
                                                <span class="zf-addresCols">
                                                    <input type="text" maxlength="255" name="Address_City" checktype="c1" placeholder="" />
                                                    <label>District / Wilaya</label>
                                                </span>
                                            </div>
                                            <div class="address_row_2">
                                                <span class="zf-addresCols">
                                                    <input type="text" maxlength="255" name="Address_AddressLine1" checktype="c1" placeholder="" />
                                                    <label>Street / Mtaa</label>
                                                </span>
                                            </div>
                                            <div class="zf-clearBoth"></div>
                                            <p id="Address_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                        </div>
                                    </div>
                                    <div class="zf-eclearBoth"></div>
                                </div>

                                <!-- REGION FIELDS -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="Tanzania" style="display:none;">
                                    <label class="zf-labelName">MKOA <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown10" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Arusha">Arusha</option><option value="Dar es Salaam">Dar es Salaam</option><option value="Dodoma">Dodoma</option>
                                            <option value="Geita">Geita</option><option value="Iringa">Iringa</option><option value="Kagera">Kagera</option>
                                            <option value="Katavi">Katavi</option><option value="Kigoma">Kigoma</option><option value="Kilimanjaro">Kilimanjaro</option>
                                            <option value="Lindi">Lindi</option><option value="Manyara">Manyara</option><option value="Mara">Mara</option>
                                            <option value="Mbeya">Mbeya</option><option value="Morogoro">Morogoro</option><option value="Mtwara">Mtwara</option>
                                            <option value="Mwanza">Mwanza</option><option value="Njombe">Njombe</option><option value="Pwani">Pwani</option>
                                            <option value="Rukwa">Rukwa</option><option value="Ruvuma">Ruvuma</option><option value="Shinyanga">Shinyanga</option>
                                            <option value="Simiyu">Simiyu</option><option value="Singida">Singida</option><option value="Songwe">Songwe</option>
                                            <option value="Tabora">Tabora</option><option value="Tanga">Tanga</option>
                                            <option value="Unguja Kaskazini">Unguja Kaskazini</option><option value="Unguja Kusini">Unguja Kusini</option>
                                            <option value="Mjini Magharibi">Mjini Magharibi</option><option value="Pemba Kaskazini">Pemba Kaskazini</option>
                                            <option value="Pemba Kusini">Pemba Kusini</option>
                                        </select>
                                        <p id="Dropdown10_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Kenya Region -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="Kenya" style="display:none;">
                                    <label class="zf-labelName">COUNTY <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown11" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Nairobi City">Nairobi City</option><option value="Mombasa">Mombasa</option><option value="Kisumu">Kisumu</option>
                                            <option value="Nakuru">Nakuru</option><option value="Eldoret">Eldoret</option><option value="Thika">Thika</option>
                                        </select>
                                        <p id="Dropdown11_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Uganda Region -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="Uganda" style="display:none;">
                                    <label class="zf-labelName">DISTRICT <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown12" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Kampala">Kampala</option><option value="Jinja">Jinja</option><option value="Gulu">Gulu</option>
                                        </select>
                                        <p id="Dropdown12_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Rwanda Region -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="Rwanda" style="display:none;">
                                    <label class="zf-labelName">PROVINCE <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown13" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Kigali">Kigali</option><option value="Eastern">Eastern</option><option value="Southern">Southern</option>
                                        </select>
                                        <p id="Dropdown13_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Burundi Region -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="Burundi" style="display:none;">
                                    <label class="zf-labelName">PROVINCE <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown14" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Bujumbura">Bujumbura</option><option value="Gitega">Gitega</option>
                                        </select>
                                        <p id="Dropdown14_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- South Sudan Region -->
                                <div class="zf-tempFrmWrapper zf-large region-field" data-country="South Sudan" style="display:none;">
                                    <label class="zf-labelName">STATE <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown15" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Central Equatoria">Central Equatoria</option><option value="Western Equatoria">Western Equatoria</option>
                                        </select>
                                        <p id="Dropdown15_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Date of Birth -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Date of Birth / Tarehe ya kuzaliwa</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine3" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine3_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Type of identification -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Type of identification / Aina ya Kitambulisho <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="NIN / Namba ya Utaia (NIDA kwa Tanzania)">NIN / Namba ya Utaia (NIDA kwa Tanzania)</option>
                                            <option value="Passport">Passport</option><option value="Voter ID">Voter ID</option><option value="Driving Licence">Driving Licence</option>
                                        </select>
                                        <p id="Dropdown_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- ID Number -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Enter your identification number / Weka namba ya kitambulisho husika <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine2" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine2_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Upload ID -->
                                <div class="zf-tempFrmWrapper">
                                    <label class="zf-labelName">Upload your identification document / Weka nakala ya kitambulisho <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv"><input type="file" accept="image/*" name="ImageUpload" checktype="c1" /><p id="ImageUpload_error" class="zf-errorMessage" style="display:none;">Choose any file for this field.</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Phone Number -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Enter your phone number / Namba ya Simu <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv zf-phonefld">
                                        <div class="zf-phwrapper zf-phNumber">
                                            <span><input type="text" compname="PhoneNumber" name="PhoneNumber_countrycode" maxlength="20" checktype="c7" value="" phoneFormat="1" isCountryCodeEnabled="false" fieldType="11" id="international_PhoneNumber_countrycode" valType="number" phoneFormatType="1" placeholder="" /><label>Number</label></span>
                                            <div class="zf-clearBoth"></div>
                                        </div>
                                        <p id="PhoneNumber_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Email -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Enter your email address / Barua pepe</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="Email" checktype="c5" value="" maxlength="255" fieldType="9" placeholder="" /></span><p id="Email_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Participation Type -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Select your participation type / Aina ya Mshiriki <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown1" id="participationType" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Individual Player / Pekee">Individual Player / Pekee</option>
                                            <option value="Team / Timu">Team / Timu</option>
                                        </select>
                                        <p id="Dropdown1_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Team Name -->
                                <div class="zf-tempFrmWrapper zf-large" id="teamNameWrapper" style="display:none;">
                                    <label class="zf-labelName">Team's Name / Jina la Timu</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Experience -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Experience in playing draughts / Uzoefu wa kucheza drafti</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown2" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Less than 1 year">Less than 1 year</option>
                                            <option value="1-3 years">1-3 years</option>
                                            <option value="4-6 years">4-6 years</option>
                                            <option value="7-10 years">7-10 years</option>
                                            <option value="More than 10 years">More than 10 years</option>
                                        </select>
                                        <p id="Dropdown2_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Draughts Type -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Type of draughts you play / Aina ya drafti unayocheza</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown3" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="International Draughts">International Draughts</option>
                                            <option value="English Draughts">English Draughts</option>
                                            <option value="Russian Draughts">Russian Draughts</option>
                                            <option value="Brazilian Draughts">Brazilian Draughts</option>
                                        </select>
                                        <p id="Dropdown3_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Player Level -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Your current level / Kiwango chako</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown4" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Beginner">Beginner</option>
                                            <option value="Intermediate">Intermediate</option>
                                            <option value="Advanced">Advanced</option>
                                            <option value="Expert">Expert</option>
                                            <option value="Master">Master</option>
                                        </select>
                                        <p id="Dropdown4_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Education -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Education level / Kiwango cha elimu</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown5" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Primary">Primary</option>
                                            <option value="Secondary">Secondary</option>
                                            <option value="Certificate">Certificate</option>
                                            <option value="Diploma">Diploma</option>
                                            <option value="Bachelor">Bachelor</option>
                                            <option value="Master">Master</option>
                                            <option value="PhD">PhD</option>
                                        </select>
                                        <p id="Dropdown5_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Employment Status -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Employment status / Hali ya ajira</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown6" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Student">Student</option>
                                            <option value="Employed">Employed</option>
                                            <option value="Self-Employed">Self-Employed</option>
                                            <option value="Unemployed">Unemployed</option>
                                            <option value="Retired">Retired</option>
                                        </select>
                                        <p id="Dropdown6_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- School/Workplace -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Name of school / place of work / Shule au sehemu ya kazi</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine1" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine1_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Job Title -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Job title / Wadhifa</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine5" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine5_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Disability -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Do you have any disability? / Je, una ulemavu wowote?</label>
                                    <div class="zf-tempContDiv zf-oneColumns">
                                        <span><input type="radio" name="Radio" value="Ndio" class="zf-radioBtn" /><label>Ndio</label></span>
                                        <span><input type="radio" name="Radio" value="Hapana" class="zf-radioBtn" /><label>Hapana</label></span>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Disability Details -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">If Yes, explain / Kama ndio, eleza</label>
                                    <div class="zf-tempContDiv"><span><textarea name="MultiLine" checktype="c1" rows="3" cols="40"></textarea></span><p id="MultiLine_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Emergency Contact -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Emergency contact name / Jina la mtu wa dharura</label>
                                    <div class="zf-tempContDiv"><span><input type="text" name="SingleLine6" checktype="c1" value="" maxlength="255" fieldType="1" placeholder="" /></span><p id="SingleLine6_error" class="zf-errorMessage" style="display:none;">Invalid value</p></div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Emergency Relation -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Relation / Uhusiano</label>
                                    <div class="zf-tempContDiv">
                                        <select class="zf-form-sBox" name="Dropdown7" checktype="c1">
                                            <option selected="true" value="-Select-">-Select-</option>
                                            <option value="Parent">Parent</option>
                                            <option value="Spouse">Spouse</option>
                                            <option value="Sibling">Sibling</option>
                                            <option value="Friend">Friend</option>
                                            <option value="Other">Other</option>
                                        </select>
                                        <p id="Dropdown7_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- Emergency Phone -->
                                <div class="zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Emergency phone number / Namba ya simu ya dharura</label>
                                    <div class="zf-tempContDiv zf-phonefld">
                                        <div class="zf-phwrapper zf-phNumber">
                                            <span><input type="text" compname="PhoneNumber1" name="PhoneNumber1_countrycode" maxlength="20" checktype="c7" value="" phoneFormat="1" isCountryCodeEnabled="false" fieldType="11" id="international_PhoneNumber1_countrycode" valType="number" phoneFormatType="1" placeholder="" /><label>Number</label></span>
                                            <div class="zf-clearBoth"></div>
                                        </div>
                                        <p id="PhoneNumber1_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- ======= ORIGINAL ZOHO SIGNATURE (HAIBADILISHWI) ======= -->
                                <!-- Zoho signature block - tunaiweka ili Zoho validation isivunjike -->
                                <div class="zf-sign zf-tempFrmWrapper zf-large" style="display:none;">
                                    <label class="zf-labelName">Sign to confirm the information provided is accurate <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <div class="zf-signContainer" id="signContainer-Signature">
                                            <object name="Signature_obj" checktype="c8" compname="Signature">
                                                <canvas id="drawingCanvas-Signature-Zoho" tabindex="-1"></canvas>
                                                <input type="hidden" id="hiddenSignInput-Signature-Zoho" name="Signature" value="">
                                                <div class="zf-signArea" id="sign-area-Signature" style="display:none"></div>
                                                <a id="clearSign" href="javascript:void(0);" onClick="zf_ClearSignature('drawingCanvas-Signature-Zoho');">Clear</a>
                                            </object>
                                        </div>
                                        <p id="Signature_error" class="zf-errorMessage" style="display:none;">Invalid value</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                                <!-- ======= CUSTOM SIGNATURE (CHINI) ======= -->
                                <div class="zf-sign zf-tempFrmWrapper zf-large">
                                    <label class="zf-labelName">Sahihi ya kuthibitisha taarifa zote ni sahihi <em class="zf-important">*</em></label>
                                    <div class="zf-tempContDiv">
                                        <div style="border:1px solid #ccc;border-radius:6px;padding:10px;background:#fff;">
                                            <canvas id="playerSignatureCanvas" width="500" height="150" style="width:100%;height:auto;cursor:crosshair;touch-action:none;border:1px solid #ddd;border-radius:4px;"></canvas>
                                            <div style="display:flex;gap:10px;margin-top:8px;flex-wrap:wrap;">
                                                <button type="button" class="btn btn-secondary btn-sm" onclick="clearPlayerSig()">Futa Sahihi</button>
                                                <button type="button" class="btn btn-info btn-sm" onclick="previewPlayerSig()">Preview</button>
                                                <button type="button" class="btn btn-success btn-sm" onclick="printPlayerSig()">Chapisha</button>
                                            </div>
                                            <input type="hidden" id="playerSignatureData" name="playerSignatureData" value="">
                                        </div>
                                        <p id="Signature_custom_error" class="zf-errorMessage" style="display:none;">Tafadhali chora sahihi yako.</p>
                                    </div>
                                    <div class="zf-clearBoth"></div>
                                </div>

                            </ul>
                        </div>

                        <ul>
                            <li class="zf-fmFooter">
                                <button type="submit" class="zf-submitColor">Submit</button>
                            </li>
                        </ul>
                    </div>
                </form>
            </div>
        </div>

        <!-- ============================= -->
        <!-- TAB: TEAMS REGISTRATION -->
        <!-- ============================= -->
        <div class="view-tab" id="teams" role="tabpanel">
            <div class="zf-templateWidth" style="padding-top:0;">
                <div class="zf-templateWrapper">
                    <ul class="zf-tempHeadBdr">
                        <li class="zf-tempHeadContBdr">
                            <h2 class="zf-frmTitle">AFMIND CUP – Fomu ya Usajili wa Timu</h2>
                            <div class="zf-clearBoth"></div>
                        </li>
                    </ul>
                    <div class="zf-subContWrap zf-topAlign">
                        <form id="teamRegistrationForm" enctype="multipart/form-data">
                            <!-- SEHEMU A: TAARIFA ZA TIMU -->
                            <div class="section-title">
                                <span class="section-icon">A</span>
                                <span>SEHEMU A: Taarifa za Timu</span>
                                <span class="section-badge">Sehemu A</span>
                            </div>
                            <div class="form-grid">
                                <div class="form-group">
                                    <label for="teamName">1. Jina la Timu <span class="required-star">*</span></label>
                                    <input type="text" id="teamName" name="teamName" required>
                                </div>
                                <div class="form-group">
                                    <label for="regionTeam">2. Mkoa <span class="required-star">*</span></label>
                                    <input type="text" id="regionTeam" name="region" required>
                                </div>
                                <div class="form-group">
                                    <label for="district">3. Wilaya <span class="required-star">*</span></label>
                                    <input type="text" id="district" name="district" required>
                                </div>
                                <div class="form-group">
                                    <label for="ward">4. Kata/Mtaa/Kijiji <span class="required-star">*</span></label>
                                    <input type="text" id="ward" name="ward" required>
                                </div>
                                <div class="form-group">
                                    <label for="yearFounded">5. Mwaka wa Kuanzishwa <span class="required-star">*</span></label>
                                    <input type="text" id="yearFounded" name="yearFounded" required>
                                </div>
                                <div class="form-group">
                                    <label for="address">6. Anwani ya Timu</label>
                                    <input type="text" id="address" name="address">
                                </div>
                                <div class="form-group">
                                    <label for="teamPhone">7. Simu ya Timu <span class="required-star">*</span></label>
                                    <input type="tel" id="teamPhone" name="teamPhone" required>
                                </div>
                                <div class="form-group">
                                    <label for="email">8. Barua Pepe (kama ipo)</label>
                                    <input type="email" id="email" name="email">
                                </div>
                            </div>

                            <!-- SEHEMU B: VIONGOZI -->
                            <div class="section-title">
                                <span class="section-icon">B</span>
                                <span>SEHEMU B: Viongozi wa Timu</span>
                                <span class="section-badge">Sehemu B</span>
                            </div>
                            <p style="font-weight:700; margin-bottom:6px; color:#1a3c6e; font-size:0.9rem;">1. Mwenyekiti wa Timu</p>
                            <div class="form-grid" style="grid-template-columns: 1fr 1fr;">
                                <div class="form-group">
                                    <label>Jina Kamili <span class="required-star">*</span></label>
                                    <input type="text" name="chairpersonName" required>
                                </div>
                                <div class="form-group">
                                    <label>Simu <span class="required-star">*</span></label>
                                    <input type="tel" name="chairpersonPhone" required>
                                </div>
                            </div>
                            <p style="font-weight:700; margin:16px 0 6px; color:#1a3c6e; font-size:0.9rem;">2. Katibu wa Timu</p>
                            <div class="form-grid" style="grid-template-columns: 1fr 1fr;">
                                <div class="form-group">
                                    <label>Jina Kamili <span class="required-star">*</span></label>
                                    <input type="text" name="secretaryName" required>
                                </div>
                                <div class="form-group">
                                    <label>Simu <span class="required-star">*</span></label>
                                    <input type="tel" name="secretaryPhone" required>
                                </div>
                            </div>
                            <p style="font-weight:700; margin:16px 0 6px; color:#1a3c6e; font-size:0.9rem;">3. Kiongozi wa Timu mmiliki (kama yupo) / Taasisi / Shirika n.k.</p>
                            <div class="form-grid" style="grid-template-columns: 1fr 1fr;">
                                <div class="form-group">
                                    <label>Jina Kamili</label>
                                    <input type="text" name="ownerName">
                                </div>
                                <div class="form-group">
                                    <label>Simu</label>
                                    <input type="tel" name="ownerPhone">
                                </div>
                            </div>

                            <!-- SEHEMU C: WACHEZAJI -->
                            <div class="section-title">
                                <span class="section-icon">C</span>
                                <span>SEHEMU C: Orodha ya Wachezaji</span>
                                <span class="section-badge">Sehemu C</span>
                            </div>
                            <div class="table-wrapper">
                                <table class="player-table" id="playerTable">
                                    <thead><tr><th>Na.</th><th>Jina Kamili</th><th>Namba ya Usajili wa AFMIND</th><th>Simu</th></tr></thead>
                                    <tbody>
                                        <tr class="sub-header"><td colspan="4">WACHEZAJI WAKUU (1 – 5)</td></tr>
                                        <tr><td class="num-cell">1</td><td><input type="text" name="player1Name" required></td><td><input type="text" name="player1Reg" required></td><td><input type="tel" name="player1Phone" required></td></tr>
                                        <tr><td class="num-cell">2</td><td><input type="text" name="player2Name" required></td><td><input type="text" name="player2Reg" required></td><td><input type="tel" name="player2Phone" required></td></tr>
                                        <tr><td class="num-cell">3</td><td><input type="text" name="player3Name" required></td><td><input type="text" name="player3Reg" required></td><td><input type="tel" name="player3Phone" required></td></tr>
                                        <tr><td class="num-cell">4</td><td><input type="text" name="player4Name" required></td><td><input type="text" name="player4Reg" required></td><td><input type="tel" name="player4Phone" required></td></tr>
                                        <tr><td class="num-cell">5</td><td><input type="text" name="player5Name" required></td><td><input type="text" name="player5Reg" required></td><td><input type="tel" name="player5Phone" required></td></tr>
                                        <tr class="sub-header"><td colspan="4">WACHEZAJI WA AKIBA (6 – 8) — Si lazima</td></tr>
                                        <tr><td class="num-cell">6</td><td><input type="text" name="player6Name"></td><td><input type="text" name="player6Reg"></td><td><input type="tel" name="player6Phone"></td></tr>
                                        <tr><td class="num-cell">7</td><td><input type="text" name="player7Name"></td><td><input type="text" name="player7Reg"></td><td><input type="tel" name="player7Phone"></td></tr>
                                        <tr><td class="num-cell">8</td><td><input type="text" name="player8Name"></td><td><input type="text" name="player8Reg"></td><td><input type="tel" name="player8Phone"></td></tr>
                                    </tbody>
                                </table>
                            </div>
                            <p style="font-size:0.75rem; color:#888; margin-top:6px; font-style:italic;">* Wachezaji wa akiba (6–8) si lazima kujazwa. Wachezaji wakuu 1–5 ni lazima.</p>

                            <!-- SEHEMU D: AHADI -->
                            <div class="section-title">
                                <span class="section-icon">D</span>
                                <span>SEHEMU D: Ahadi ya Timu</span>
                                <span class="section-badge">Sehemu D</span>
                            </div>
                            <div class="pledge-box">
                                <p class="pledge-text">
                                    Sisi viongozi na wachezaji wa timu hii tunathibitisha kuwa taarifa zote zilizotolewa ni sahihi.
                                    Tunakubali kufuata Katiba, Kanuni, Sheria na Maamuzi yote ya <strong>AFMIND/Mofuli</strong>.
                                </p>
                                <div class="pledge-grid">
                                    <div class="form-group">
                                        <label>Jina la Kiongozi Mkuu wa Timu <span class="required-star">*</span></label>
                                        <input type="text" name="leaderName" required>
                                    </div>
                                    <div class="form-group">
                                        <label>Tarehe <span class="required-star">*</span></label>
                                        <input type="date" name="pledgeDate" required>
                                    </div>
                                </div>
                            </div>

                            <!-- SAHIHI YA TIMU -->
                            <div class="signature-section">
                                <div class="sig-title">Sahihi ya Kiongozi Mkuu (Chora hapa chini)</div>
                                <div class="canvas-container">
                                    <canvas id="teamSignatureCanvas" width="500" height="180"></canvas>
                                </div>
                                <div class="sig-buttons">
                                    <button type="button" class="btn-clear-sig" onclick="clearTeamSig()">Futa Sahihi</button>
                                    <button type="button" class="btn btn-info btn-sm" onclick="previewTeamSig()">Preview</button>
                                    <button type="button" class="btn btn-success btn-sm" onclick="printTeamSig()">Chapisha</button>
                                </div>
                                <input type="hidden" name="teamSignatureData" id="teamSignatureData" value="">
                                <p style="font-size:0.72rem; color:#888; text-align:center; margin-top:6px;">Tumia kipanya au kidole kuchora sahihi yako.</p>
                            </div>

                            <!-- SEHEMU E: VIAMBATISHO -->
                            <div class="section-title">
                                <span class="section-icon">E</span>
                                <span>SEHEMU E: Viambatisho vya Usajili</span>
                                <span class="section-badge">Sehemu E</span>
                            </div>
                            <div class="attachments-grid">
                                <div class="attachment-card">
                                    <div class="att-label">1. Katiba ya Timu</div>
                                    <span class="att-badge optional">Hiari (kama ipo)</span>
                                    <input type="file" name="attachment_katiba" accept=".pdf,.doc,.docx,image/*">
                                </div>
                                <div class="attachment-card">
                                    <div class="att-label">2. Cheti cha Usajili wa Kikundi/Timu</div>
                                    <span class="att-badge optional">Hiari (kama ipo)</span>
                                    <input type="file" name="attachment_ceti" accept=".pdf,.doc,.docx,image/*">
                                </div>
                                <div class="attachment-card">
                                    <div class="att-label">3. Nembo (Logo) ya Timu <span class="required-star">*</span></div>
                                    <span class="att-badge required">Lazima</span>
                                    <input type="file" name="attachment_logo" accept="image/*" required>
                                </div>
                                <div class="attachment-card">
                                    <div class="att-label">4. Picha ya Pamoja ya Timu <span class="required-star">*</span></div>
                                    <span class="att-badge required">Lazima</span>
                                    <input type="file" name="attachment_picha" accept="image/*" required>
                                </div>
                                <div class="attachment-card">
                                    <div class="att-label">5. NIDA ya Kiongozi wa Timu <span class="required-star">*</span></div>
                                    <span class="att-badge required">Lazima</span>
                                    <input type="file" name="attachment_nida" accept="image/*,.pdf" required>
                                </div>
                                <div class="attachment-card">
                                    <div class="att-label">6. Barua ya Afisa Muajiri</div>
                                    <span class="att-badge optional">Hiari (kwa kampuni/taasisi)</span>
                                    <input type="file" name="attachment_barua" accept=".pdf,.doc,.docx,image/*">
                                </div>
                            </div>

                            <div class="stamp-area">
                                <span style="font-size:0.8rem; color:#777;">Muhuri wa Timu / Kiambatisho cha ziada (kama upo):</span>
                                <div class="stamp-placeholder">Muhuri</div>
                                <input type="file" name="stampAttachment" accept="image/*,.pdf" style="font-size:0.8rem; max-width:200px;">
                            </div>

                            <div class="btn-row">
                                <button type="submit" class="btn btn-submit">Wasilisha Fomu</button>
                                <button type="reset" class="btn btn-reset">Futa Yote</button>
                                <button type="button" class="btn btn-print" onclick="window.print()">Chapisha Fomu</button>
                                <button type="button" class="btn btn-info" onclick="previewTeamForm()">Preview Fomu</button>
                            </div>
                            <p class="form-footer-note">Tafadhali hakikisha taarifa zote ni sahihi kabla ya kuwasilisha.</p>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- ============================= -->
        <!-- TAB: STAFF REGISTRATION -->
        <!-- ============================= -->
        <div class="view-tab" id="staffreg" role="tabpanel">
            <div class="zf-templateWidth" style="padding-top:0;">
                <div class="zf-templateWrapper">
                    <ul class="zf-tempHeadBdr">
                        <li class="zf-tempHeadContBdr">
                            <h2 class="zf-frmTitle">STAFF REGISTRATION (REFEREE / COORDINATOR)</h2>
                            <div class="zf-clearBoth"></div>
                        </li>
                    </ul>
                    <div class="zf-subContWrap zf-topAlign">
                        <form id="staffRegForm">
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Full Name <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv"><input type="text" id="staffFullName" class="zf-form-sBox" required /></div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Email</label>
                                <div class="zf-tempContDiv"><input type="email" id="staffEmail" class="zf-form-sBox" /></div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Phone Number <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv"><input type="text" id="staffPhone" class="zf-form-sBox" required /></div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Role <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv">
                                    <select id="staffRole" class="zf-form-sBox" required>
                                        <option value="referee">Referee</option>
                                        <option value="coordinator">Coordinator</option>
                                    </select>
                                </div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Nationality <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv">
                                    <select id="staffNationality" class="zf-form-sBox">
                                        <option>Tanzania</option><option>Kenya</option><option>Uganda</option>
                                        <option>Rwanda</option><option>Burundi</option><option>South Sudan</option>
                                    </select>
                                </div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Region (Mkoa) <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv"><input type="text" id="staffRegion" class="zf-form-sBox" required /></div>
                            </div>
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Identification Number</label>
                                <div class="zf-tempContDiv"><input type="text" id="staffIdNumber" class="zf-form-sBox" /></div>
                            </div>
                            <div class="zf-tempFrmWrapper">
                                <label class="zf-labelName">Upload ID Image</label>
                                <div class="zf-tempContDiv"><input type="file" id="staffIdImage" accept="image/*" /></div>
                            </div>

                            <!-- Staff Signature -->
                            <div class="zf-tempFrmWrapper zf-large">
                                <label class="zf-labelName">Sahihi (Chora hapa chini) <em class="zf-important">*</em></label>
                                <div class="zf-tempContDiv">
                                    <div style="border:1px solid #ccc;border-radius:6px;padding:10px;background:#fff;">
                                        <canvas id="staffSignatureCanvas" width="500" height="150" style="width:100%;height:auto;cursor:crosshair;touch-action:none;"></canvas>
                                        <div style="display:flex;gap:10px;margin-top:8px;flex-wrap:wrap;">
                                            <button type="button" class="btn btn-secondary btn-sm" onclick="clearStaffSig()">Futa Sahihi</button>
                                            <button type="button" class="btn btn-info btn-sm" onclick="previewStaffSig()">Preview</button>
                                            <button type="button" class="btn btn-success btn-sm" onclick="printStaffSig()">Chapisha</button>
                                        </div>
                                        <input type="hidden" name="staffSignatureData" id="staffSignatureData" value="">
                                    </div>
                                    <p id="staffSignature_error" class="zf-errorMessage" style="display:none;">Tafadhali chora sahihi yako.</p>
                                </div>
                                <div class="zf-clearBoth"></div>
                            </div>

                            <div class="zf-fmFooter">
                                <button type="submit" class="zf-submitColor">Register as Staff</button>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- ============================= -->
        <!-- TAB: RANKING -->
        <!-- ============================= -->
        <div class="view-tab" id="ranking" role="tabpanel">
            <div class="card card-dashboard">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <span><i class="fas fa-trophy text-warning"></i> Ranking</span>
                    <button class="btn btn-sm btn-outline-secondary" onclick="exportRankingExcel()"><i class="fas fa-file-excel"></i> Export Excel</button>
                </div>
                <div class="card-body">
                    <!-- Sub Tabs -->
                    <ul class="nav nav-tabs" id="rankingSubTabs" role="tablist">
                        <li class="nav-item"><a class="nav-link active" data-rankview="individual" href="#rankIndividual">Individual</a></li>
                        <li class="nav-item"><a class="nav-link" data-rankview="team" href="#rankTeam">Team</a></li>
                        <li class="nav-item"><a class="nav-link" data-rankview="overall" href="#rankOverall">Overall</a></li>
                    </ul>
                    <div class="tab-content pt-3">
                        <div class="rank-view active" id="rankIndividual">
                            <div class="table-responsive">
                                <table class="table table-hover align-middle">
                                    <thead class="table-light"><tr><th>#</th><th>Player ID</th><th>Name</th><th>Team</th><th>Played</th><th>W</th><th>D</th><th>L</th><th>Points</th></tr></thead>
                                    <tbody id="individualRankingBody"><tr><td colspan="9" class="text-center">Loading...</td></tr></tbody>
                                </table>
                            </div>
                        </div>
                        <div class="rank-view" id="rankTeam">
                            <div class="table-responsive">
                                <table class="table table-hover align-middle">
                                    <thead class="table-light"><tr><th>#</th><th>Team Name</th><th>Players</th><th>Total Points</th><th>Avg Points</th></tr></thead>
                                    <tbody id="teamRankingBody"><tr><td colspan="5" class="text-center">Loading...</td></tr></tbody>
                                </table>
                            </div>
                        </div>
                        <div class="rank-view" id="rankOverall">
                            <div class="table-responsive">
                                <table class="table table-hover align-middle">
                                    <thead class="table-light"><tr><th>#</th><th>Name</th><th>Type</th><th>Points</th></tr></thead>
                                    <tbody id="overallRankingBody"><tr><td colspan="4" class="text-center">Loading...</td></tr></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- ============================= -->
        <!-- TAB: STAFF PORTAL (RESULTS) -->
        <!-- ============================= -->
        <div class="view-tab" id="staff" role="tabpanel">
            <div class="card card-dashboard">
                <div class="card-header">Submit Match Result</div>
                <div class="card-body">
                    <form id="resultForm">
                        <div class="row">
                            <div class="col-md-6 mb-3">
                                <label class="form-label">Winner Player ID <span class="text-danger">*</span></label>
                                <input type="text" class="form-control" id="rWinner" placeholder="e.g. ADL-TZ-DAR-000001" required />
                            </div>
                            <div class="col-md-6 mb-3">
                                <label class="form-label">Loser Player ID <span class="text-danger">*</span></label>
                                <input type="text" class="form-control" id="rLoser" placeholder="e.g. ADL-TZ-DAR-000002" required />
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6 mb-3">
                                <label class="form-label">Score <span class="text-danger">*</span></label>
                                <input type="text" class="form-control" id="rScore" placeholder="e.g. 2-0 or 1-1" required />
                            </div>
                            <div class="col-md-6 mb-3">
                                <label class="form-label">Round</label>
                                <input type="number" class="form-control" id="rRound" value="1" />
                            </div>
                        </div>
                        <button type="submit" class="btn btn-primary">Submit Result</button>
                        <div id="resultMsg" class="mt-2"></div>
                    </form>
                    <hr />
                    <h6>Recent Results</h6>
                    <div id="recentResults" class="small"></div>
                </div>
            </div>
        </div>

        <!-- ============================= -->
        <!-- TAB: ADMIN DASHBOARD -->
        <!-- ============================= -->
        <div class="view-tab" id="admin" role="tabpanel">
            <!-- Dashboard Cards -->
            <div class="row">
                <div class="col-md-3"><div class="card text-white bg-primary mb-3"><div class="card-body"><h5 class="card-title"><i class="fas fa-users"></i> Total Players</h5><h2 id="totalPlayers">0</h2></div></div></div>
                <div class="col-md-3"><div class="card text-white bg-success mb-3"><div class="card-body"><h5 class="card-title"><i class="fas fa-user-tie"></i> Staff</h5><h2 id="totalStaff">0</h2></div></div></div>
                <div class="col-md-3"><div class="card text-white bg-warning mb-3"><div class="card-body"><h5 class="card-title"><i class="fas fa-clock"></i> Pending Verification</h5><h2 id="pendingVerification">0</h2></div></div></div>
                <div class="col-md-3"><div class="card text-white bg-info mb-3"><div class="card-body"><h5 class="card-title"><i class="fas fa-trophy"></i> Games Played</h5><h2 id="totalGames">0</h2></div></div></div>
            </div>

            <!-- Charts -->
            <div class="row">
                <div class="col-md-4"><div class="card card-dashboard"><div class="card-header">Gender Distribution</div><div class="card-body"><canvas id="genderChart" height="180"></canvas></div></div></div>
                <div class="col-md-4"><div class="card card-dashboard"><div class="card-header">Player Status</div><div class="card-body"><canvas id="statusChart" height="180"></canvas></div></div></div>
                <div class="col-md-4"><div class="card card-dashboard"><div class="card-header">Disability / Special Needs</div><div class="card-body"><canvas id="disabilityChart" height="180"></canvas></div></div></div>
            </div>
            <div class="row mt-3">
                <div class="col-md-4"><div class="card card-dashboard"><div class="card-header">Students vs Others</div><div class="card-body"><canvas id="studentChart" height="180"></canvas></div></div></div>
                <div class="col-md-8"><div class="card card-dashboard"><div class="card-header">Registrations by Region</div><div class="card-body"><canvas id="regionChart" height="180"></canvas></div></div></div>
            </div>

            <!-- Verification Center -->
            <div class="card card-dashboard mt-3">
                <div class="card-header">Verification Center</div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead><tr><th>Player ID</th><th>Name</th><th>Status</th><th>Action</th></tr></thead>
                            <tbody id="verifyTableBody"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Ads Manager -->
            <div class="card card-dashboard mt-3">
                <div class="card-header">Ads Manager</div>
                <div class="card-body">
                    <form id="adForm" class="row g-2">
                        <div class="col-md-3"><input type="text" class="form-control" id="adTitle" placeholder="Ad Title" required /></div>
                        <div class="col-md-3"><input type="text" class="form-control" id="adLink" placeholder="Link (optional)" /></div>
                        <div class="col-md-2"><input type="date" class="form-control" id="adStart" required /></div>
                        <div class="col-md-2"><input type="date" class="form-control" id="adEnd" required /></div>
                        <div class="col-md-1"><input type="file" class="form-control" id="adImage" accept="image/*" required /></div>
                        <div class="col-md-1"><button type="submit" class="btn btn-primary btn-sm">Upload Ad</button></div>
                    </form>
                    <div id="adList" class="mt-2"></div>
                </div>
            </div>

            <!-- Export Buttons -->
            <div class="card card-dashboard mt-3">
                <div class="card-header">Export Data</div>
                <div class="card-body">
                    <button class="btn btn-outline-success btn-sm me-2" onclick="exportPlayersExcel()"><i class="fas fa-file-excel"></i> Players Excel</button>
                    <button class="btn btn-outline-danger btn-sm me-2" onclick="exportPlayersPDF()"><i class="fas fa-file-pdf"></i> Players PDF</button>
                    <input type="text" id="playerIdForPdf" class="form-control mt-2" placeholder="Enter Player ID for single PDF" style="width:250px; display:inline-block;" />
                    <button class="btn btn-outline-warning btn-sm" onclick="exportPlayerPDF(document.getElementById('playerIdForPdf').value)">Export Single PDF</button>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- ============================================= -->
<!-- PIN LOGIN MODAL -->
<!-- ============================================= -->
<div class="modal fade" id="pinModal" tabindex="-1">
    <div class="modal-dialog">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title">Enter Access PIN</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <p class="text-muted small">Staff PIN: 10ZA811SX | Admin PIN: A771Z8min</p>
                <input type="password" id="pinInput" class="form-control" placeholder="••••••••" />
                <div id="pinMsg" class="mt-2"></div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-primary" onclick="verifyPin()">Submit</button>
            </div>
        </div>
    </div>
</div>

<!-- Footer -->
<footer class="text-center text-muted py-3 mt-5 border-top">
    &copy; 2026 AFMIND Draughts League | Built with Firebase
</footer>

<!-- Bootstrap JS -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

<!-- ============================================================ -->
<!-- ZOHO VALIDATION SCRIPTS (zilizobaki kama zilivyo) -->
<!-- ============================================================ -->
<script>
    // -----------------------------------------------------------------------
    // ZOHO VALIDATION (hazijabadilishwa)
    // -----------------------------------------------------------------------
    function zf_CheckMandatory() { return true; }
    function zf_ValidCheck() { return true; }
    function zf_ShowErrorMsg(id, msg) { /* ignored */ }
    function zf_ClearSignature(canvasId) { /* ignored */ }
    // ... (zoho functions zote zinaweza kuachwa hapa, hazitatumiwa)
</script>

<!-- ============================================================ -->
<!-- FIREBASE + MAIN APP LOGIC -->
<!-- ============================================================ -->
<script type="module">
    import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
    import { getAnalytics } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-analytics.js";
    import {
        getFirestore,
        collection,
        addDoc,
        getDocs,
        query,
        orderBy,
        onSnapshot,
        doc,
        updateDoc,
        where,
        deleteDoc,
        serverTimestamp,
        runTransaction,
        increment
    } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";
    import { getStorage, ref, uploadBytes, getDownloadURL } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-storage.js";

    const firebaseConfig = {
        apiKey: "AIzaSyCPOR-LUmTHl9tEPr_FJIXqfR0XeeEJX3w",
        authDomain: "tech-zone-18bd7.firebaseapp.com",
        projectId: "tech-zone-18bd7",
        storageBucket: "tech-zone-18bd7.firebasestorage.app",
        messagingSenderId: "1044876741263",
        appId: "1:1044876741263:web:d9f7d660d448c32f5afefa",
        measurementId: "G-ETV8ZP11BQ"
    };

    const app = initializeApp(firebaseConfig);
    const analytics = getAnalytics(app);
    const db = getFirestore(app);
    const storage = getStorage(app);

    const STAFF_PIN = "10ZA811SX";
    const ADMIN_PIN = "A771Z8min";

    let isAdmin = false;
    let isStaff = false;

    function showLoading() { document.getElementById('loadingOverlay').style.display = 'flex'; }
    function hideLoading() { document.getElementById('loadingOverlay').style.display = 'none'; }

    function updateUIForRole() {
        const userDisplay = document.getElementById('userDisplay');
        if (isAdmin) { userDisplay.textContent = 'Admin (PIN)'; document.getElementById('loginBtn').classList.add('d-none'); document.getElementById('logoutBtn').classList.remove('d-none'); }
        else if (isStaff) { userDisplay.textContent = 'Staff (PIN)'; document.getElementById('loginBtn').classList.add('d-none'); document.getElementById('logoutBtn').classList.remove('d-none'); }
        else { userDisplay.textContent = ''; document.getElementById('loginBtn').classList.remove('d-none'); document.getElementById('logoutBtn').classList.add('d-none'); }
    }

    window.switchTab = function(tabId) {
        document.querySelectorAll('.view-tab').forEach(el => el.classList.remove('active'));
        document.getElementById(tabId).classList.add('active');
        document.querySelectorAll('#mainTabs .nav-link').forEach(el => el.classList.remove('active'));
        const link = document.querySelector(`#mainTabs .nav-link[data-view="${tabId}"]`);
        if (link) link.classList.add('active');
    };

    window.openPinLogin = () => new bootstrap.Modal(document.getElementById('pinModal')).show();

    window.verifyPin = () => {
        const pin = document.getElementById('pinInput').value.trim();
        const msg = document.getElementById('pinMsg');
        if (pin === ADMIN_PIN) { isAdmin = true; isStaff = false; msg.innerHTML = '<div class="alert alert-success">Admin access granted.</div>'; updateUIForRole(); switchTab('admin'); bootstrap.Modal.getInstance(document.getElementById('pinModal')).hide(); }
        else if (pin === STAFF_PIN) { isAdmin = false; isStaff = true; msg.innerHTML = '<div class="alert alert-success">Staff access granted.</div>'; updateUIForRole(); switchTab('staff'); bootstrap.Modal.getInstance(document.getElementById('pinModal')).hide(); }
        else { msg.innerHTML = '<div class="alert alert-danger">Invalid PIN. Try again.</div>'; }
        document.getElementById('pinInput').value = '';
    };

    window.clearAccess = () => { isAdmin = false; isStaff = false; updateUIForRole(); switchTab('register'); };

    document.querySelectorAll('#mainTabs .nav-link[data-view]').forEach(el => {
        el.addEventListener('click', (e) => {
            e.preventDefault();
            const view = el.dataset.view;
            if ((view === 'staff' || view === 'results') && !isStaff && !isAdmin) { alert('Please login with Staff PIN first.'); openPinLogin(); return; }
            if (view === 'admin' && !isAdmin) { alert('Please login with Admin PIN first.'); openPinLogin(); return; }
            switchTab(view);
        });
    });

    // ------------------------------------------------------------
    // MAPS FOR PLAYER ID
    // ------------------------------------------------------------
    const countryCodeMap = { 'Tanzania':'TZ','Kenya':'KE','Uganda':'UG','Rwanda':'RW','Burundi':'BI','South Sudan':'SS' };
    const regionCodeMap = {
        'Dar es Salaam':'DAR','Arusha':'ARU','Dodoma':'DOD','Mwanza':'MWZ','Mbeya':'MBY','Morogoro':'MOR',
        'Kigoma':'KGM','Tanga':'TNG','Kilimanjaro':'KLJ','Iringa':'IRN','Kagera':'KGR','Mara':'MAR',
        'Manyara':'MNY','Njombe':'NJM','Pwani':'PWN','Rukwa':'RKW','Ruvuma':'RVM','Shinyanga':'SHY',
        'Simiyu':'SMY','Singida':'SGD','Songwe':'SGW','Tabora':'TBR','Katavi':'KTV','Geita':'GTA',
        'Lindi':'LND','Mtwara':'MTW','Unguja Kaskazini':'UNK','Unguja Kusini':'UNS','Mjini Magharibi':'MMG',
        'Pemba Kaskazini':'PBK','Pemba Kusini':'PBS','Nairobi City':'NBO','Mombasa':'MSA','Kisumu':'KSM'
    };

    function getRegionValue() {
        const map = { 'Tanzania':'Dropdown10','Kenya':'Dropdown11','Uganda':'Dropdown12','Rwanda':'Dropdown13','Burundi':'Dropdown14','South Sudan':'Dropdown15' };
        const select = document.querySelector(`[name="${map[document.getElementById('nationalitySelect').value]}"]`);
        return select ? select.value : '';
    }

    async function generatePlayerId(country, region) {
        const countryCode = countryCodeMap[country] || 'XX';
        const regionCode = regionCodeMap[region] || '000';
        const counterRef = doc(db, 'counters', 'playerId');
        const newNum = await runTransaction(db, async (transaction) => {
            const counterDoc = await transaction.get(counterRef);
            let last = counterDoc.exists ? counterDoc.data().lastNumber : 0;
            const next = last + 1;
            transaction.set(counterRef, { lastNumber: next }, { merge: true });
            return next;
        });
        return `ADL-${countryCode}-${regionCode}-${String(newNum).padStart(6, '0')}`;
    }

    // ------------------------------------------------------------
    // PLAYER REGISTRATION
    // ------------------------------------------------------------
    document.getElementById('afmindForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        // Validate custom signature
        const sigData = document.getElementById('playerSignatureData').value;
        if (!sigData) {
            alert('Tafadhali chora sahihi yako kwenye kisanduku cha sahihi.');
            document.querySelector('#playerSignatureCanvas').style.border = '2px solid #e74c3c';
            document.querySelector('#playerSignatureCanvas').scrollIntoView({ behavior: 'smooth', block: 'center' });
            return;
        }
        // Copy custom signature to Zoho hidden input so Zoho validation passes
        document.getElementById('hiddenSignInput-Signature-Zoho').value = sigData;

        // Let Zoho validation run (if any)
        if (typeof zf_CheckMandatory !== 'undefined' && !zf_CheckMandatory()) return;
        if (typeof zf_ValidCheck !== 'undefined' && !zf_ValidCheck()) return;

        showLoading();
        const submitBtn = document.querySelector('.zf-submitColor');
        submitBtn.textContent = 'Submitting...';
        submitBtn.disabled = true;

        try {
            const formData = {
                salutation: document.querySelector('[name="Name_Salutation"]').value,
                fullName: document.querySelector('[name="Name_Last"]').value,
                nickname: document.querySelector('[name="Name_First"]').value,
                gender: document.querySelector('[name="Dropdown8"]').value,
                nationality: document.querySelector('[name="Dropdown9"]').value,
                district: document.querySelector('[name="Address_City"]').value,
                street: document.querySelector('[name="Address_AddressLine1"]').value,
                region: getRegionValue(),
                dob: document.querySelector('[name="SingleLine3"]').value,
                idType: document.querySelector('[name="Dropdown"]').value,
                idNumber: document.querySelector('[name="SingleLine2"]').value,
                phone: document.querySelector('[name="PhoneNumber_countrycode"]').value,
                email: document.querySelector('[name="Email"]').value,
                participationType: document.querySelector('[name="Dropdown1"]').value,
                teamName: document.querySelector('[name="SingleLine"]').value,
                experience: document.querySelector('[name="Dropdown2"]').value,
                draughtsType: document.querySelector('[name="Dropdown3"]').value,
                playerLevel: document.querySelector('[name="Dropdown4"]').value,
                education: document.querySelector('[name="Dropdown5"]').value,
                employmentStatus: document.querySelector('[name="Dropdown6"]').value,
                school: document.querySelector('[name="SingleLine1"]').value,
                jobTitle: document.querySelector('[name="SingleLine5"]').value,
                disability: document.querySelector('input[name="Radio"]:checked')?.value || '',
                disabilityDetails: document.querySelector('[name="MultiLine"]').value,
                emergencyName: document.querySelector('[name="SingleLine6"]').value,
                emergencyRelation: document.querySelector('[name="Dropdown7"]').value,
                emergencyPhone: document.querySelector('[name="PhoneNumber1_countrycode"]').value,
                signature: sigData,
                status: 'pending',
                played: 0, wins: 0, draws: 0, losses: 0, points: 0
            };

            for (let f of [{ input:'[name="ImageUpload2"]', key:'profilePhotoURL' }, { input:'[name="ImageUpload"]', key:'idPhotoURL' }]) {
                const file = document.querySelector(f.input).files[0];
                if (file) {
                    const storageRef = ref(storage, `players/${Date.now()}_${file.name}`);
                    await uploadBytes(storageRef, file);
                    formData[f.key] = await getDownloadURL(storageRef);
                }
            }

            const playerId = await generatePlayerId(formData.nationality, formData.region);
            formData.playerId = playerId;
            await addDoc(collection(db, 'players'), { ...formData, createdAt: serverTimestamp() });
            alert(`✅ Registered! Your Player ID: ${playerId}`);
        } catch (err) { alert('Error: ' + err.message); }
        finally { submitBtn.textContent = 'Submit'; submitBtn.disabled = false; hideLoading(); }
    });

    // ------------------------------------------------------------
    // STAFF REGISTRATION
    // ------------------------------------------------------------
    document.getElementById('staffRegForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const sig = document.getElementById('staffSignatureData').value;
        if (!sig) { alert('Tafadhali chora sahihi yako.'); return; }
        showLoading();
        const data = {
            fullName: document.getElementById('staffFullName').value,
            email: document.getElementById('staffEmail').value,
            phone: document.getElementById('staffPhone').value,
            role: document.getElementById('staffRole').value,
            nationality: document.getElementById('staffNationality').value,
            region: document.getElementById('staffRegion').value,
            idNumber: document.getElementById('staffIdNumber').value,
            signature: sig,
            status: 'pending'
        };
        const file = document.getElementById('staffIdImage').files[0];
        if (file) {
            const storageRef = ref(storage, `staff_ids/${Date.now()}_${file.name}`);
            await uploadBytes(storageRef, file);
            data.idPhotoURL = await getDownloadURL(storageRef);
        }
        try { await addDoc(collection(db, 'staff'), data); alert('Staff registration submitted!'); document.getElementById('staffRegForm').reset(); clearStaffSig(); }
        catch (err) { alert(err.message); }
        hideLoading();
    });

    // ------------------------------------------------------------
    // TEAM REGISTRATION
    // ------------------------------------------------------------
    document.getElementById('teamRegistrationForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const sig = document.getElementById('teamSignatureData').value;
        if (!sig) { alert('Tafadhali chora sahihi ya kiongozi mkuu.'); return; }

        showLoading();
        const form = e.target;
        const fd = new FormData(form);
        const teamData = {
            teamName: fd.get('teamName'),
            region: fd.get('region'),
            district: fd.get('district'),
            ward: fd.get('ward'),
            yearFounded: fd.get('yearFounded'),
            address: fd.get('address'),
            teamPhone: fd.get('teamPhone'),
            email: fd.get('email'),
            chairpersonName: fd.get('chairpersonName'),
            chairpersonPhone: fd.get('chairpersonPhone'),
            secretaryName: fd.get('secretaryName'),
            secretaryPhone: fd.get('secretaryPhone'),
            ownerName: fd.get('ownerName'),
            ownerPhone: fd.get('ownerPhone'),
            leaderName: fd.get('leaderName'),
            pledgeDate: fd.get('pledgeDate'),
            signature: sig,
            players: [],
            status: 'pending'
        };

        for (let i = 1; i <= 8; i++) {
            const name = fd.get(`player${i}Name`);
            const reg = fd.get(`player${i}Reg`);
            const phone = fd.get(`player${i}Phone`);
            if (name && reg) teamData.players.push({ no: i, name, reg, phone });
        }

        const fileFields = ['attachment_logo','attachment_picha','attachment_nida','attachment_katiba','attachment_ceti','attachment_barua','stampAttachment'];
        for (const key of fileFields) {
            const file = fd.get(key);
            if (file && file.size > 0) {
                const storageRef = ref(storage, `teams/${Date.now()}_${file.name}`);
                await uploadBytes(storageRef, file);
                teamData[key] = await getDownloadURL(storageRef);
            }
        }

        try {
            await addDoc(collection(db, 'teams'), { ...teamData, createdAt: serverTimestamp() });
            alert('✅ Timu imesajiliwa kikamilifu!');
            form.reset();
            clearTeamSig();
        } catch (err) { alert('Error: ' + err.message); }
        hideLoading();
    });

    // ------------------------------------------------------------
    // RANKING
    // ------------------------------------------------------------
    function loadIndividualRanking() {
        const q = query(collection(db, 'players'), where('status', '==', 'approved'), orderBy('points', 'desc'));
        onSnapshot(q, (snap) => {
            const tbody = document.getElementById('individualRankingBody');
            if (snap.empty) { tbody.innerHTML = '<tr><td colspan="9">No approved players.</td></tr>'; return; }
            let html = '';
            snap.forEach((doc, i) => {
                const d = doc.data();
                const rank = i + 1;
                html += `<tr class="ranking-row ${rank===1?'table-warning':''}">
                    <td><strong>${rank}</strong></td>
                    <td><span class="player-id">${d.playerId||'—'}</span></td>
                    <td>${d.fullName}</td>
                    <td>${d.teamName||'—'}</td>
                    <td>${d.played||0}</td>
                    <td>${d.wins||0}</td>
                    <td>${d.draws||0}</td>
                    <td>${d.losses||0}</td>
                    <td><strong>${d.points||0}</strong></td>
                </tr>`;
            });
            tbody.innerHTML = html;
        });
    }

    async function loadTeamRanking() {
        const tbody = document.getElementById('teamRankingBody');
        tbody.innerHTML = '<tr><td colspan="5" class="text-center">Loading...</td></tr>';
        try {
            const teamsSnap = await getDocs(collection(db, 'teams'));
            if (teamsSnap.empty) { tbody.innerHTML = '<tr><td colspan="5">No teams registered.</td></tr>'; return; }
            const teamPoints = [];
            for (const teamDoc of teamsSnap.docs) {
                const t = teamDoc.data();
                const teamName = t.teamName || 'Unknown';
                const playersSnap = await getDocs(query(collection(db, 'players'), where('teamName', '==', teamName), where('status', '==', 'approved')));
                let total = 0, count = 0;
                playersSnap.forEach(p => { total += (p.data().points || 0); count++; });
                teamPoints.push({ name: teamName, players: count, totalPoints: total, avgPoints: count > 0 ? total/count : 0 });
            }
            teamPoints.sort((a,b) => b.totalPoints - a.totalPoints);
            let html = '';
            teamPoints.forEach((t, i) => {
                html += `<tr><td><strong>${i+1}</strong></td><td>${t.name}</td><td>${t.players}</td><td><strong>${t.totalPoints}</strong></td><td>${t.avgPoints.toFixed(1)}</td></tr>`;
            });
            tbody.innerHTML = html || '<tr><td colspan="5">No data.</td></tr>';
        } catch (e) { tbody.innerHTML = '<tr><td colspan="5">Error loading teams.</td></tr>'; }
    }

    async function loadOverallRanking() {
        const tbody = document.getElementById('overallRankingBody');
        tbody.innerHTML = '<tr><td colspan="4" class="text-center">Loading...</td></tr>';
        try {
            const playersSnap = await getDocs(query(collection(db, 'players'), where('status', '==', 'approved'), orderBy('points', 'desc')));
            const teamsSnap = await getDocs(collection(db, 'teams'));
            const overall = [];
            playersSnap.forEach(d => { const p = d.data(); overall.push({ name: p.fullName, type: 'Player', points: p.points || 0 }); });
            for (const teamDoc of teamsSnap.docs) {
                const t = teamDoc.data();
                const teamName = t.teamName || 'Unknown';
                const ps = await getDocs(query(collection(db, 'players'), where('teamName', '==', teamName), where('status', '==', 'approved')));
                let total = 0;
                ps.forEach(p => { total += (p.data().points || 0); });
                overall.push({ name: teamName, type: 'Team', points: total });
            }
            overall.sort((a,b) => b.points - a.points);
            let html = '';
            overall.slice(0, 50).forEach((item, i) => {
                html += `<tr><td><strong>${i+1}</strong></td><td>${item.name}</td><td>${item.type}</td><td><strong>${item.points}</strong></td></tr>`;
            });
            tbody.innerHTML = html || '<tr><td colspan="4">No data.</td></tr>';
        } catch (e) { tbody.innerHTML = '<tr><td colspan="4">Error loading overall.</td></tr>'; }
    }

    document.querySelectorAll('#rankingSubTabs .nav-link').forEach(el => {
        el.addEventListener('click', function(e) {
            e.preventDefault();
            document.querySelectorAll('#rankingSubTabs .nav-link').forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            const view = this.dataset.rankview;
            document.querySelectorAll('.rank-view').forEach(v => v.classList.remove('active'));
            document.getElementById('rank' + view.charAt(0).toUpperCase() + view.slice(1)).classList.add('active');
            if (view === 'team') loadTeamRanking();
            else if (view === 'overall') loadOverallRanking();
            else loadIndividualRanking();
        });
    });

    loadIndividualRanking();

    // ------------------------------------------------------------
    // RESULTS
    // ------------------------------------------------------------
    document.getElementById('resultForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!isStaff && !isAdmin) { alert('Staff only.'); return; }
        showLoading();
        const msg = document.getElementById('resultMsg');
        const winnerId = document.getElementById('rWinner').value.trim();
        const loserId = document.getElementById('rLoser').value.trim();
        const score = document.getElementById('rScore').value.trim();
        const round = parseInt(document.getElementById('rRound').value) || 1;
        try {
            const wSnap = await getDocs(query(collection(db, 'players'), where('playerId', '==', winnerId), where('status', '==', 'approved')));
            const lSnap = await getDocs(query(collection(db, 'players'), where('playerId', '==', loserId), where('status', '==', 'approved')));
            if (wSnap.empty || lSnap.empty) throw new Error('Player(s) not found or not approved.');
            await addDoc(collection(db, 'results'), { winnerId, loserId, score, round, submittedBy: 'pin-user', date: serverTimestamp() });
            await runTransaction(db, async (transaction) => {
                const wRef = wSnap.docs[0].ref, lRef = lSnap.docs[0].ref;
                const wData = (await transaction.get(wRef)).data();
                const lData = (await transaction.get(lRef)).data();
                const isDraw = score.includes('1-1') || score.includes('0-0') || score.toLowerCase().includes('draw');
                if (isDraw) {
                    transaction.update(wRef, { played: increment(1), draws: increment(1), points: increment(1) });
                    transaction.update(lRef, { played: increment(1), draws: increment(1), points: increment(1) });
                } else {
                    transaction.update(wRef, { played: increment(1), wins: increment(1), points: increment(3) });
                    transaction.update(lRef, { played: increment(1), losses: increment(1) });
                }
            });
            msg.innerHTML = '<div class="alert alert-success">✅ Result submitted! Ranking updated.</div>';
            document.getElementById('resultForm').reset();
        } catch (err) { msg.innerHTML = `<div class="alert alert-danger">${err.message}</div>`; }
        hideLoading();
    });

    // ------------------------------------------------------------
    // ADMIN STATS & VERIFICATION
    // ------------------------------------------------------------
    function loadAdminStats() {
        onSnapshot(collection(db, 'players'), (snap) => {
            document.getElementById('totalPlayers').textContent = snap.size;
            document.getElementById('pendingVerification').textContent = snap.docs.filter(d => d.data().status === 'pending').length;
        });
        onSnapshot(collection(db, 'staff'), (snap) => { document.getElementById('totalStaff').textContent = snap.size; });
        onSnapshot(collection(db, 'results'), (snap) => { document.getElementById('totalGames').textContent = snap.size; });
    }

    function loadVerificationList() {
        onSnapshot(query(collection(db, 'players'), where('status', 'in', ['pending','approved','rejected'])), (snap) => {
            const tbody = document.getElementById('verifyTableBody');
            if (snap.empty) { tbody.innerHTML = '<tr><td colspan="4">No players.</td></tr>'; return; }
            let html = '';
            snap.forEach(doc => {
                const d = doc.data();
                const cls = d.status === 'approved' ? 'text-success' : d.status === 'rejected' ? 'text-danger' : 'text-warning';
                html += `<tr>
                    <td><span class="player-id">${d.playerId||'—'}</span></td>
                    <td>${d.fullName}</td>
                    <td class="${cls}">${d.status}</td>
                    <td>${d.status === 'pending' ? `
                        <button class="btn btn-sm btn-success" onclick="verifyPlayer('${doc.id}','approve')">Approve</button>
                        <button class="btn btn-sm btn-danger" onclick="verifyPlayer('${doc.id}','reject')">Reject</button>` : '<span class="text-muted">Done</span>'}</td>
                </tr>`;
            });
            tbody.innerHTML = html;
        });
    }

    window.verifyPlayer = async (docId, action) => {
        if (!isAdmin) return;
        showLoading();
        const status = action === 'approve' ? 'approved' : 'rejected';
        await updateDoc(doc(db, 'players', docId), { status, verifiedAt: serverTimestamp() });
        await addDoc(collection(db, 'verification_logs'), { playerId: docId, status, verifiedBy: 'admin', timestamp: serverTimestamp() });
        hideLoading();
        alert(`Player ${status}!`);
    };

    // ------------------------------------------------------------
    // ADMIN CHARTS
    // ------------------------------------------------------------
    async function loadAdminCharts() {
        const snap = await getDocs(collection(db, 'players'));
        let male=0,female=0,other=0, approved=0,pending=0,rejected=0, withDis=0,withoutDis=0, students=0,nonStudents=0;
        const regions = {};
        snap.forEach(doc => {
            const d = doc.data();
            const g = (d.gender || '').toLowerCase();
            if (g.includes('male') || g.includes('mwanaume')) male++;
            else if (g.includes('female') || g.includes('mwanamke')) female++;
            else other++;

            if (d.status === 'approved') approved++;
            else if (d.status === 'pending') pending++;
            else if (d.status === 'rejected') rejected++;

            if (d.disability && d.disability.toLowerCase() === 'ndio') withDis++;
            else withoutDis++;

            const edu = (d.education || '').toLowerCase();
            const emp = (d.employmentStatus || '').toLowerCase();
            if (edu.includes('student') || emp.includes('student') || edu.includes('mwanafunzi') || emp.includes('mwanafunzi')) students++;
            else nonStudents++;

            const reg = d.region || 'Unknown';
            regions[reg] = (regions[reg] || 0) + 1;
        });

        new Chart(document.getElementById('genderChart'), { type:'doughnut', data:{ labels:['Wanaume','Wanawake','Nyingine'], datasets:[{ data:[male,female,other], backgroundColor:['#0d6efd','#d63384','#6c757d'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        new Chart(document.getElementById('statusChart'), { type:'doughnut', data:{ labels:['Approved','Pending','Rejected'], datasets:[{ data:[approved,pending,rejected], backgroundColor:['#198754','#ffc107','#dc3545'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        new Chart(document.getElementById('disabilityChart'), { type:'doughnut', data:{ labels:['Walemavu / Mahitaji Maalum','Wasio na Ulemavu'], datasets:[{ data:[withDis,withoutDis], backgroundColor:['#fd7e14','#20c997'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        new Chart(document.getElementById('studentChart'), { type:'doughnut', data:{ labels:['Wanafunzi','Wengine'], datasets:[{ data:[students,nonStudents], backgroundColor:['#0dcaf0','#6f42c1'] }] }, options:{ responsive:true, plugins:{ legend:{ position:'bottom' } } } });
        new Chart(document.getElementById('regionChart'), { type:'bar', data:{ labels:Object.keys(regions), datasets:[{ label:'Players', data:Object.values(regions), backgroundColor:'#0d6efd' }] }, options:{ responsive:true, plugins:{ legend:{ display:false } } } });
    }

    // ------------------------------------------------------------
    // ADS
    // ------------------------------------------------------------
    document.getElementById('adForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!isAdmin) return;
        showLoading();
        const file = document.getElementById('adImage').files[0];
        if (!file) { alert('Select image'); hideLoading(); return; }
        const storageRef = ref(storage, `ads/${Date.now()}_${file.name}`);
        await uploadBytes(storageRef, file);
        const url = await getDownloadURL(storageRef);
        await addDoc(collection(db, 'ads'), {
            title: document.getElementById('adTitle').value,
            link: document.getElementById('adLink').value,
            imageURL: url,
            startDate: document.getElementById('adStart').value,
            endDate: document.getElementById('adEnd').value,
            active: true,
            createdAt: serverTimestamp()
        });
        document.getElementById('adForm').reset();
        loadAdsList();
        hideLoading();
    });

    async function loadAdsList() {
        const snap = await getDocs(query(collection(db, 'ads'), where('active', '==', true)));
        const container = document.getElementById('adList');
        let html = '<div class="row">';
        snap.forEach(doc => {
            const d = doc.data();
            html += `<div class="col-md-3 mb-2"><div class="card"><img src="${d.imageURL}" class="card-img-top">
                <div class="card-body"><p class="small"><a href="${d.link||'#'}">${d.title}</a></p>
                <button class="btn btn-sm btn-danger" onclick="deleteAd('${doc.id}')">Delete</button></div></div></div>`;
        });
        container.innerHTML = html || '<p>No active ads.</p>';
    }
    window.deleteAd = async (id) => { if (isAdmin) { await deleteDoc(doc(db, 'ads', id)); loadAdsList(); } };

    // ------------------------------------------------------------
    // EXPORTS
    // ------------------------------------------------------------
    window.exportPlayersExcel = async () => {
        showLoading(); const snap = await getDocs(collection(db, 'players')); const data = snap.docs.map(d => d.data());
        const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Players'); XLSX.writeFile(wb, 'players.xlsx'); hideLoading();
    };
    window.exportRankingExcel = async () => {
        showLoading(); const snap = await getDocs(query(collection(db, 'players'), where('status','==','approved'), orderBy('points','desc')));
        const data = snap.docs.map(d => d.data()); const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Ranking'); XLSX.writeFile(wb, 'ranking.xlsx'); hideLoading();
    };
    window.exportPlayersPDF = async () => {
        showLoading(); const snap = await getDocs(collection(db, 'players')); const { jsPDF } = window.jspdf; const pdf = new jsPDF(); let y = 20;
        pdf.text('Players List', 20, y); y += 10;
        snap.forEach((doc, i) => { const d = doc.data(); pdf.text(`${i+1}. ${d.fullName} (${d.playerId||'—'})`, 20, y); y += 6; if (y>270) { pdf.addPage(); y=20; } });
        pdf.save('players.pdf'); hideLoading();
    };
    window.exportPlayerPDF = async (playerId) => {
        showLoading(); const q = query(collection(db, 'players'), where('playerId','==',playerId)); const snap = await getDocs(q);
        if (snap.empty) { alert('Player not found'); hideLoading(); return; }
        const d = snap.docs[0].data(); const { jsPDF } = window.jspdf; const pdf = new jsPDF();
        pdf.setFontSize(16); pdf.text('Player Profile', 20, 20); pdf.setFontSize(12);
        pdf.text(`ID: ${d.playerId}`, 20, 30); pdf.text(`Name: ${d.fullName}`, 20, 38);
        pdf.text(`Nationality: ${d.nationality}`, 20, 46); pdf.text(`Region: ${d.region}`, 20, 54);
        pdf.text(`Points: ${d.points}`, 20, 62); pdf.save(`${d.playerId}.pdf`); hideLoading();
    };

    // ------------------------------------------------------------
    // RECENT RESULTS
    // ------------------------------------------------------------
    (async () => {
        const q = query(collection(db, 'results'), orderBy('date', 'desc'));
        const snap = await getDocs(q);
        let html = '<ul class="list-unstyled">';
        snap.forEach(doc => { const d = doc.data(); html += `<li><span class="badge bg-secondary">${d.round||1}</span> ${d.winnerId} vs ${d.loserId} → ${d.score}</li>`; });
        document.getElementById('recentResults').innerHTML = html || '<p>No results yet.</p>';
    })();

    // ------------------------------------------------------------
    // INIT
    // ------------------------------------------------------------
    window.addEventListener('DOMContentLoaded', () => {
        loadAdminStats();
        loadVerificationList();
        loadAdsList();
        loadAdminCharts();

        // Ads banner
        onSnapshot(query(collection(db, 'ads'), where('active', '==', true)), (snap) => {
            const container = document.getElementById('adsBanner');
            if (snap.empty) { container.innerHTML = ''; return; }
            let html = '';
            snap.forEach(doc => {
                const d = doc.data();
                html += `<div class="col-md-3 col-6 mb-2"><a href="${d.link||'#'}" target="_blank"><img src="${d.imageURL}" alt="${d.title}" class="img-fluid rounded" style="max-height:80px;" /></a></div>`;
            });
            container.innerHTML = html;
        });
    });

    // ============================================================
    // PLAYER SIGNATURE (CUSTOM)
    // ============================================================
    (function() {
        const canvas = document.getElementById('playerSignatureCanvas');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        const hiddenInput = document.getElementById('playerSignatureData');
        let drawing = false, lastX = 0, lastY = 0;

        canvas.width = 500; canvas.height = 150;

        function getPos(e) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            const scaleY = canvas.height / rect.height;
            let cx, cy;
            if (e.touches) { cx = e.touches[0].clientX; cy = e.touches[0].clientY; }
            else { cx = e.clientX; cy = e.clientY; }
            return { x: (cx - rect.left) * scaleX, y: (cy - rect.top) * scaleY };
        }

        function start(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; ctx.beginPath(); ctx.moveTo(lastX, lastY); }
        function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
        function stop() { if (drawing) { drawing = false; ctx.beginPath(); update(); } }

        function update() {
            const data = ctx.getImageData(0,0,canvas.width,canvas.height).data;
            let empty = true;
            for (let i=0; i<data.length; i+=4) { if (data[i+3] !== 0) { empty = false; break; } }
            hiddenInput.value = empty ? '' : canvas.toDataURL('image/png');
            canvas.style.border = empty ? '1px solid #ddd' : '1px solid #198754';
        }

        window.clearPlayerSig = function() { ctx.clearRect(0,0,canvas.width,canvas.height); update(); };
        window.previewPlayerSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:90%;max-height:90%;border:2px solid #ccc;border-radius:8px;" />`);
            win.document.close();
        };
        window.printPlayerSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:100%;max-height:100%;" />`);
            win.document.close();
            win.print();
        };

        ctx.strokeStyle = '#1a3c6e'; ctx.lineWidth = 2.5; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
        canvas.addEventListener('mousedown', start); canvas.addEventListener('mousemove', draw); canvas.addEventListener('mouseup', stop);
        canvas.addEventListener('mouseleave', stop); canvas.addEventListener('touchstart', start, { passive: false });
        canvas.addEventListener('touchmove', draw, { passive: false }); canvas.addEventListener('touchend', stop); canvas.addEventListener('touchcancel', stop);
        clearPlayerSig();
    })();

    // ============================================================
    // STAFF SIGNATURE
    // ============================================================
    (function() {
        const canvas = document.getElementById('staffSignatureCanvas');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        const hiddenInput = document.getElementById('staffSignatureData');
        let drawing = false, lastX=0, lastY=0;
        canvas.width = 500; canvas.height = 150;

        function getPos(e) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            const scaleY = canvas.height / rect.height;
            let cx, cy;
            if (e.touches) { cx = e.touches[0].clientX; cy = e.touches[0].clientY; }
            else { cx = e.clientX; cy = e.clientY; }
            return { x: (cx - rect.left) * scaleX, y: (cy - rect.top) * scaleY };
        }
        function start(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; ctx.beginPath(); ctx.moveTo(lastX, lastY); }
        function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
        function stop() { if (drawing) { drawing = false; ctx.beginPath(); update(); } }
        function update() {
            const data = ctx.getImageData(0,0,canvas.width,canvas.height).data;
            let empty = true;
            for (let i=0; i<data.length; i+=4) { if (data[i+3] !== 0) { empty = false; break; } }
            hiddenInput.value = empty ? '' : canvas.toDataURL('image/png');
            canvas.style.border = empty ? '1px solid #ddd' : '1px solid #198754';
        }
        window.clearStaffSig = function() { ctx.clearRect(0,0,canvas.width,canvas.height); update(); };
        window.previewStaffSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:90%;max-height:90%;border:2px solid #ccc;border-radius:8px;" />`);
            win.document.close();
        };
        window.printStaffSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:100%;max-height:100%;" />`);
            win.document.close();
            win.print();
        };
        ctx.strokeStyle = '#1a3c6e'; ctx.lineWidth = 2.5; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
        canvas.addEventListener('mousedown', start); canvas.addEventListener('mousemove', draw); canvas.addEventListener('mouseup', stop);
        canvas.addEventListener('mouseleave', stop); canvas.addEventListener('touchstart', start, { passive: false });
        canvas.addEventListener('touchmove', draw, { passive: false }); canvas.addEventListener('touchend', stop); canvas.addEventListener('touchcancel', stop);
        clearStaffSig();
    })();

    // ============================================================
    // TEAM SIGNATURE
    // ============================================================
    (function() {
        const canvas = document.getElementById('teamSignatureCanvas');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        const hiddenInput = document.getElementById('teamSignatureData');
        let drawing = false, lastX=0, lastY=0;
        canvas.width = 500; canvas.height = 180;

        function getPos(e) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            const scaleY = canvas.height / rect.height;
            let cx, cy;
            if (e.touches) { cx = e.touches[0].clientX; cy = e.touches[0].clientY; }
            else { cx = e.clientX; cy = e.clientY; }
            return { x: (cx - rect.left) * scaleX, y: (cy - rect.top) * scaleY };
        }
        function start(e) { e.preventDefault(); drawing = true; const p = getPos(e); lastX = p.x; lastY = p.y; ctx.beginPath(); ctx.moveTo(lastX, lastY); }
        function draw(e) { if (!drawing) return; e.preventDefault(); const p = getPos(e); ctx.lineTo(p.x, p.y); ctx.stroke(); lastX = p.x; lastY = p.y; }
        function stop() { if (drawing) { drawing = false; ctx.beginPath(); update(); } }
        function update() {
            const data = ctx.getImageData(0,0,canvas.width,canvas.height).data;
            let empty = true;
            for (let i=0; i<data.length; i+=4) { if (data[i+3] !== 0) { empty = false; break; } }
            hiddenInput.value = empty ? '' : canvas.toDataURL('image/png');
            canvas.style.border = empty ? '1px solid #ddd' : '1px solid #198754';
        }
        window.clearTeamSig = function() { ctx.clearRect(0,0,canvas.width,canvas.height); update(); };
        window.previewTeamSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:90%;max-height:90%;border:2px solid #ccc;border-radius:8px;" />`);
            win.document.close();
        };
        window.printTeamSig = function() {
            const data = hiddenInput.value;
            if (!data) { alert('Hakuna sahihi.'); return; }
            const win = window.open('','_blank','width=600,height=400');
            win.document.write(`<img src="${data}" style="max-width:100%;max-height:100%;" />`);
            win.document.close();
            win.print();
        };
        ctx.strokeStyle = '#1a3c6e'; ctx.lineWidth = 2.5; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
        canvas.addEventListener('mousedown', start); canvas.addEventListener('mousemove', draw); canvas.addEventListener('mouseup', stop);
        canvas.addEventListener('mouseleave', stop); canvas.addEventListener('touchstart', start, { passive: false });
        canvas.addEventListener('touchmove', draw, { passive: false }); canvas.addEventListener('touchend', stop); canvas.addEventListener('touchcancel', stop);
        clearTeamSig();
    })();

    // ============================================================
    // PREVIEW TEAM FORM
    // ============================================================
    window.previewTeamForm = function() {
        const form = document.getElementById('teamRegistrationForm');
        const fd = new FormData(form);
        let html = '<html><head><title>Preview Fomu ya Timu</title><style>body{font-family:sans-serif;padding:20px;} table{width:100%;border-collapse:collapse;} td,th{border:1px solid #ccc;padding:8px;}</style></head><body>';
        html += '<h2>AFMIND CUP - Fomu ya Usajili wa Timu</h2><table>';
        for (let [key, value] of fd.entries()) {
            if (value instanceof File) {
                html += `<tr><td><strong>${key}</strong></td><td>Faili: ${value.name} (${(value.size/1024).toFixed(1)} KB)</td></tr>`;
            } else {
                html += `<tr><td><strong>${key}</strong></td><td>${value || '—'}</td></tr>`;
            }
        }
        html += '</table></body></html>';
        const win = window.open('', '_blank', 'width=800,height=600');
        win.document.write(html);
        win.document.close();
    };

</script>

</body>
</html>
