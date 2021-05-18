/**
 * node.jsでADOを利用してAccessのmdbファイルを利用するサンプル
 */
const fs = require('fs');
const path = require('path'); 
require('winax');  // npm i winax

// MDBファイルを作成する
// 要 Microsoft Access Database Engine 2016 Redistributable
// https://www.microsoft.com/en-us/download/details.aspx?id=54920
const filename = 'persons.mdb';
const data_path = path.join(__dirname, '/data/', filename);

if (fs.existsSync(data_path)) {
  fs.unlinkSync(data_path);
}

const constr = 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + data_path;
const cat = new ActiveXObject('ADOX.Catalog');
cat.Create(constr);
const con = cat.ActiveConnection
con.Execute('create Table persons (Name char(50), City char(50), Phone char(20), Zip decimal(5))');
con.Execute("insert into persons values('John', 'London','123-45-67','14589')");
con.Execute("insert into persons values('Andrew', 'Paris','333-44-55','38215')");
con.Execute("insert into persons values('Romeo', 'Rom','222-33-44','54323')");

var rs = con.Execute('Select * from persons'); 
var fields = rs.Fields;
console.log('Result field count: ' + fields.Count);
console.log('Result record count: ' + rs.RecordCount);

rs.MoveFirst();
while (!rs.EOF) {
    // Access as property by string key
    var name = fields['Name'].Value;

    // Access as method with string argument
    var town = fields('City').value;

    // Access as indexed array
    var phone = fields[2].value;
    var zip = fields[3].value;    

    console.log('> Person: '+name+' from ' + town + ' phone: ' + phone + ' zip: ' + zip);
    rs.MoveNext();
}

con.Close();