//
require(`winax`);

//const fs=require('fs');
const path=require('path');

const PATH_TO_MDB=path.resolve(`..\\access\\dbsion.accdb`);
console.log('path is '+PATH_TO_MDB);

var con=new ActiveXObject(`ADODB.Connection`);
con.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+PATH_TO_MDB);
var rs=con.Execute('select * from priceitem');
console.log(`*****record count is =====>${rs.RecordCount}`);
var fc=rs.Fields.Count;
console.log('fields count is '+fc);

rs.movefirst();
var rc=1;
while(!rs.eof) {
    rs.movenext();
    if(!rs.eof){rc=rc+1}
}
console.log(rc);
//at last
con.Close();