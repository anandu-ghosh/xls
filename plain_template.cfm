<cfset plain = spreadsheetNew()>
<cfset spreadsheetAddRow(plain, "First Name,Last Name,Address,Email,Phone,DOB,Role")>
<cfheader name="Content-Disposition" value="attachment; filename=Plain_Template.xls">
<cfcontent type="application/msexcel" variable="#spreadsheetReadBinary(plain)#" reset="true">




