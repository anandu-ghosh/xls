component{
    function userData(){
        allUser = queryExecute("SELECT firstname,lastname,address,email,phone,dob,GROUP_CONCAT(role.role)as roles FROM user inner join user_roles on user_roles.user_id = user.id inner join role on role.id = user_roles.role_id group by user.id", {});
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
            local.sheet = queryNew("FirstName,LastName,Address,Email,Phone,DOB,Role,Result");
            for(rows in local.sheetData){
                queryAddRow(local.sheet);
                querySetCell(local.sheet, "FirstName", rows["FirstName"]);
                querySetCell(local.sheet, "LastName", rows["LastName"]);
                querySetCell(local.sheet, "Address", rows["Address"]);
                querySetCell(local.sheet, "Email", rows["Email"]);
                querySetCell(local.sheet, "Phone", rows["Phone"]);
                querySetCell(local.sheet, "DOB", rows["DOB"]);
                querySetCell(local.sheet, "Role", rows["Role"]);
                querySetCell(local.sheet, "Result", rows["Result"]); 
            }
            local.mySheet = SpreadsheetNew();
            SpreadSheetAddRow(local.mySheet,"First Name,Last Name,Address,Email,Phone,DOB,Role,Result");
            spreadsheetAddRows(local.mySheet,local.sheet);
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
                            local.userId = queryExecute("SELECT id FROM user WHERE email = :email",{
                                email: { cfsqltype: "cf_sql_varchar", value: row["Email"]}
                            });        
                            queryExecute("DELETE FROM user_roles WHERE user_id = :userId",{
                                userId : { cfsqltype: "cf_sql_integer", value:local.userId.id }
                            })
                            ;
                            queryExecute("UPDATE user SET firstname = :firstname, lastname= :lastname, address= :address,phone= :phone, dob= :dob WHERE email = :email;",{
                                firstname: { cfsqltype: "cf_sql_varchar", value: row["First Name"] },
                                lastname: { cfsqltype: "cf_sql_varchar", value: row["Last Name"] },
                                address: { cfsqltype: "cf_sql_varchar", value: row["Address"] },
                                phone: { cfsqltype: "cf_sql_varchar", value: row["Phone"] },
                                dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(row["DOB"],"yyy-mm-dd") },
                                email: { cfsqltype: "cf_sql_varchar", value: row["Email"]}
                            });

                            for(ros in userrole){
                                local.roleId = queryExecute("SELECT id FROM role WHERE role = :roles",{
                                    roles: { cfsqltype: "cf_sql_varchar", value: ros}
                                }); 

                                queryExecute("insert into user_roles(user_id,role_id)values( :user_id, :role_id)",{
                                    user_id: { cfsqltype: "cf_sql_varchar", value: local.userId.id },
                                    role_id: { cfsqltype: "cf_sql_varchar", value: local.roleId.id }
                                },{ result="roledata" });
                            } 
                        }
                    }else{
                        if(local.roleStatus === false){
                            return 'role_error';
                        }else{
                            queryExecute("insert into user(firstname,lastname,address,email,phone,dob)values( :firstname, :lastname, :address, :email, :phone, :dob )",{
                                firstname: { cfsqltype: "cf_sql_varchar", value: row["First Name"] },
                                lastname: { cfsqltype: "cf_sql_varchar", value: row["Last Name"] },
                                address: { cfsqltype: "cf_sql_varchar", value: row["Address"] },
                                email: { cfsqltype: "cf_sql_varchar", value: row["Email"] },
                                phone: { cfsqltype: "cf_sql_varchar", value: row["Phone"] },
                                dob: { cfsqltype: "cf_sql_varchar", value: DateFormat(row["DOB"],"yyy-mm-dd") }
                            },{ result="resultdata" });
                            
                            for(ros in userrole){
                                local.roleId = queryExecute("SELECT id FROM role WHERE role = :roles",{
                                    roles: { cfsqltype: "cf_sql_varchar", value: ros}
                                }); 

                                queryExecute("insert into user_roles(user_id,role_id)values( :user_id, :role_id)",{
                                    user_id: { cfsqltype: "cf_sql_varchar", value: resultdata.generated_key },
                                    role_id: { cfsqltype: "cf_sql_varchar", value: local.roleId.id }
                                },{ result="roledata" });
                            }
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
        local.sheetData =queryNew("Id,FirstName,LastName,Address,Email,Phone,DOB,Role,Result");
        var i = 1;
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
                    querySetCell(local.sheetData, "Id", i++);
                    querySetCell(local.sheetData, "FirstName", row["First Name"]);
                    querySetCell(local.sheetData, "LastName", row["Last Name"]);
                    querySetCell(local.sheetData, "Address", row["Address"]);
                    querySetCell(local.sheetData, "Email", row["Email"]);
                    querySetCell(local.sheetData, "Phone", row["Phone"]);
                    querySetCell(local.sheetData, "DOB", row["DOB"]);
                    querySetCell(local.sheetData, "Role", row["Role"]);
                    querySetCell(local.sheetData, "Result", local.endresult);
                   
                }else{
                    querySetCell(local.sheetData, "Id", "");
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
        local.data = sortedQuery = queryExecute("SELECT * FROM sheetData ORDER BY Id DESC;", {}, {dbtype: "query"});
        return local.data;
    }
}