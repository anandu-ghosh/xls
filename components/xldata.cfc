component{
    function userData(){
        allUser = queryExecute("SELECT * FROM user", {});
        return allUser;
    }

    function xlfileRead(data){
        
    }
}