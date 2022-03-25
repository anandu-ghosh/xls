component{
    function userData(){
        allUser = queryExecute("SELECT * FROM user", {});
        return allUser;
    }

    function downloadWithData(){
        local.dataSheet = SpreadsheetNew();
        SpreadSheetAddRow(local.dataSheet,"First Name,Last Name,Address,Email,Phone,DOB,Role");
        local.userData = userData();
        SpreadSheetAddRows(local.dataSheet,local.userData);
        cfheader( name="Content-Disposition", value="attachment;filename=User_data.xls" );
        cfcontent( variable=SpreadSheetReadBinary(local.dataSheet), type="application/msexcel" );
    }

    function xlfileRead(data){
        local.check_sheet_rows   = checkSheetColumns(data);
        if(local.check_sheet_rows === false){
            local.sheetData = uploadFailDataCreate(data);
            local.mySheet = SpreadsheetNew();

            SpreadSheetAddRow(local.mySheet,"First Name,Last Name,Address,Email,Phone,DOB,Role,Result");
            spreadsheetAddRows(local.mySheet,local.sheetData);
            cfheader( name="Content-Disposition", value="attachment;filename=Upload_result.xls" );
            cfcontent( variable=SpreadSheetReadBinary(local.mySheet), type="application/msexcel" );
        }else if(local.check_sheet_rows === true){
            local.role = queryExecute("SELECT role FROM role");
            local.roles = ValueArray(local.role,"role");
            for(row in data){
                if(row["First Name"] != '' || row["Last Name"] != '' || row["Address"] != '' || row["Email"] != '' || row["Phone"] != '' || row["DOB"] != '' || row["Role"] != ''){
                    local.emailCheck = queryExecute("SELECT email FROM user WHERE email = :email;",{ 
                        email : { cfsqltype: "cf_sql_varchar", value: row["Email"]} 
                    });
                    local.userrole = listToArray(row["Role"]);
                    for(ros in userrole){
                        if(ArrayContains(roles,ros) === false){
                            local.roleStatus = false
                        }else{
                            local.roleStatus = true
                        }
                    }
                    if(local.emailCheck.RecordCount > 0){
                        if(local.roleStatus === false){
                            return 'role_error';
                        }else{
                            queryExecute("UPDATE user SET firstname = :firstname, lastname= :lastname, address= :address,phone= :phone, dob= :dob,role = :role WHERE email = :email;",{
                                firstname: { cfsqltype: "cf_sql_varchar", value: row["First Name"] },
                                lastname: { cfsqltype: "cf_sql_varchar", value: row["Last Name"] },
                                address: { cfsqltype: "cf_sql_varchar", value: row["Address"] },
                                phone: { cfsqltype: "cf_sql_varchar", value: row["Phone"] },
                                dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(row["DOB"],"yyy-mm-dd") },
                                role: { cfsqltype: "cf_sql_varchar", value: row["Role"] },
                                email: { cfsqltype: "cf_sql_varchar", value: row["Email"]}
                            });
                        }
                    }else{
                        if(local.roleStatus === false){
                            return 'role_error';
                        }else{
                            queryExecute("insert into user(firstname,lastname,address,email,phone,dob,role)values( :firstname, :lastname, :address, :email, :phone, :dob, :role )",{
                                firstname: { cfsqltype: "cf_sql_varchar", value: row["First Name"] },
                                lastname: { cfsqltype: "cf_sql_varchar", value: row["Last Name"] },
                                address: { cfsqltype: "cf_sql_varchar", value: row["Address"] },
                                email: { cfsqltype: "cf_sql_varchar", value: row["Email"] },
                                phone: { cfsqltype: "cf_sql_varchar", value: row["Phone"] },
                                dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(row["DOB"],"yyy-mm-dd") },
                                role: { cfsqltype: "cf_sql_varchar", value: row["Role"] }
                            });
                        }
                    }
                }    
            }
            return 'success';
        }
    }

    function checkSheetColumns(data){
        local.result = [];
        local.i = 1;
        for(row in data){
            if(row["First Name"] != '' || row["Last Name"] != '' || row["Address"] != '' || row["Email"] != '' || row["Phone"] != '' || row["DOB"] != '' || row["Role"] != ''){
                if(row["First Name"] === '' || row["Last Name"] === '' || row["Address"] === '' || row["Email"] === '' || row["Phone"] === '' || row["DOB"] === '' || row["Role"] === ''){
                    local.result[i] = false;
                }else{
                    local.result[i] = true;
                }
            }
            i++;
        }
        if(ArrayContains(result,false)){
            return false;
        }else{
            return true;
        }  
    }

    function uploadFailDataCreate(data){
        local.sheetData =queryNew("FirstName,LastName,Address,Email,Phone,DOB,Role,Result");
        for(row in data){
            queryAddRow(local.sheetData);
            local.nullMessage =[];
            if(row["First Name"] != '' || row["Last Name"] != '' || row["Address"] != '' || row["Email"] != '' || row["Phone"] != '' || row["DOB"] != '' || row["Role"] != ''){
                if(row["First Name"] === '' || row["Last Name"] === '' || row["Address"] === '' || row["Email"] === '' || row["Phone"] === '' || row["DOB"] === '' || row["Role"] === ''){

                    if(row["First Name"] === ''){
                        local.nullMessage[1] = 'First Name missing';
                    }
                    if(row["Last Name"] === ''){
                        local.nullMessage[2] = 'Second Name missing';
                    }
                    if(row["Address"] === ''){
                        local.nullMessage[3] = 'Address missing';
                    }
                    if(row["Email"] === ''){
                        local.nullMessage[4] = 'Email missing';
                    }
                    if(row["Phone"] === ''){
                        local.nullMessage[5] = 'Phone missing';
                    }
                    if(row["DOB"] === ''){
                        local.nullMessage[6] = 'DOB missing';
                    }
                    if(row["Role"] === ''){
                        local.nullMessage[7] = 'Role missing';
                    }
                    local.endresult = arrayToList(local.nullMessage)
                    querySetCell(local.sheetData, "FirstName", row["First Name"]);
                    querySetCell(local.sheetData, "LastName", row["Last Name"]);
                    querySetCell(local.sheetData, "Address", row["Address"]);
                    querySetCell(local.sheetData, "Email", row["Email"]);
                    querySetCell(local.sheetData, "Phone", row["Phone"]);
                    querySetCell(local.sheetData, "DOB", row["DOB"]);
                    querySetCell(local.sheetData, "Role", row["Role"]);
                    querySetCell(local.sheetData, "Result", local.endresult);
                }else{
                    querySetCell(local.sheetData, "FirstName", row["First Name"]);
                    querySetCell(local.sheetData, "LastName", row["Last Name"]);
                    querySetCell(local.sheetData, "Address", row["Address"]);
                    querySetCell(local.sheetData, "Email", row["Email"]);
                    querySetCell(local.sheetData, "Phone", row["Phone"]);
                    querySetCell(local.sheetData, "DOB", row["DOB"]);
                    querySetCell(local.sheetData, "Role", row["Role"]);
                    querySetCell(local.sheetData, "Result", 'Sucess');
                }
            }
        }
        return local.sheetData;
    }
}