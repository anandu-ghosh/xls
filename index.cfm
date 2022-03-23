<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>User</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
</head>
<body>
<cfset variables.message = structNew() />
<cfif structKeyExists(form, "form_submit")>
    <cfif form.xls_file != "">
        <cffile 
            action="upload"
            fileField="xls_file"
            destination="F:\ColdFusion2021\cfusion\wwwroot\xls\uploads"
            nameconflict="makeunique"
            result="data"
        >
        <cfset variables.path = "#data.serverdirectory#\#data.serverfile#">
        <cfspreadsheet    
            action="read" 
            src = "#variables.path#" 
            excludeHeaderRow = true
            query = "queryData"
            headerrow="1"
        >
        <cfinvoke component="components.xldata" method="xlfileRead" returnvariable="results" data="#queryData#">
    <cfelse>
        <cfset variables.message.errormsg = 'upload a excel file' />
    </cfif> 
</cfif>
<cfif structKeyExists(form, "download_data")>
    <cfinvoke component="components.xldata" method="downloadWithData" returnvariable="xlsdownload">
</cfif>
    <div class="container">
        <div class="row">
            <div class="col-12 text-center">
                <h2>User Details</h2>
            </div>
            <div class="col mt-4">
                <div class="row">
                    <div class="col">
                        <form action="" method="post">
                            <a href="plain_template.cfm" class="btn btn-primary" >Plain Template</a>
                            <button type="submit" class="btn btn-secondary" name="download_data">Template with data</button>
                        </form>
                    </div>
                    <div class="col">
                        <form  action="" method="post" enctype="multipart/form-data">
                            <div class="form-group row">
                                <div class="col-sm-8 col-form-label"> <input name="xls_file" type="file" /></div>
                                <div class="col-sm-4">
                                    <button type="submit" class="btn btn-success" name="form_submit">Upload</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>    
            </div>
            <cfinvoke component="components.xldata" method="userData" returnvariable="listData" >
            <div class="col-12 mt-5">
                <table class="table">
                    <thead>
                        <tr>
                        <th scope="col">First Name</th>
                        <th scope="col">Last Name</th>
                        <th scope="col">Address</th>
                        <th scope="col">Email</th>
                        <th scope="col">Phone</th>
                        <th scope="col">DOB</th>
                        <th scope="col">Role</th>
                        </tr>
                    </thead>
                    <tbody>
                        <cfoutput query="listData">
                            <tr>
                                <td>#listData.firstname#</td>
                                <td>#listData.lastname#</td>
                                <td>#listData.address#</td>
                                <td>#listData.email#</td>
                                <td>#listData.phone#</td>
                                <td>#listData.dob#</td>
                                <td>#listData.role#</td> 
                            </tr>
                        </cfoutput>    
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.12.9/dist/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
</body>
</html>