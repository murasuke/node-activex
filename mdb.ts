/**
 * node(typescript)でADOを利用してAccessのmdbファイルを利用するサンプル
 */
import fs from 'fs';
import path from 'path';

require('winax');  // npm i winax

// MDBファイルを作成する
// 要 Microsoft Access Database Engine 2016 Redistributable
// https://www.microsoft.com/en-us/download/details.aspx?id=54920
const filename = 'persons.mdb';
const data_path = path.join(__dirname, '/data/', filename);

// ファイルがあれば削除
if (fs.existsSync(data_path)) {
  fs.unlinkSync(data_path);
}
 
// mdbファイルを作成するため「ADODB.Connection」ではなく「ADOX.Catalog」を利用する
const constr = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${data_path}`;
const cat: ADOX.Catalog = new ActiveXObject('ADOX.Catalog');
cat.Create(constr);
const con = cat.ActiveConnection as ADODB.Connection;

try {
  // データ登録
  con.Execute('create Table persons (Name char(30), City char(30), Phone char(20), Zip decimal(5))');
  con.Execute("insert into persons values('John', 'London','123-45-67','14589')");
  con.Execute("insert into persons values('Andrew', 'Paris','333-44-55','38215')");
  con.Execute("insert into persons values('Romeo', 'Rom','222-33-44','54323')");

  // selectした結果を表示
  const rs: ADODB.Recordset = con.Execute('Select * from persons'); 
  console.log(`Result field count: ${rs.Fields.Count}`);

  while (!rs.EOF) {
    // 型指定の都合でrs.Fields['Name'] はコンパイルエラー
    const name = rs.Fields('Name').Value;  
    const town = rs.Fields('City').Value;
    const phone = rs.Fields(2).Value;
    const zip = rs.Fields(3).Value;    

    console.log(`> Person: ${name} from ${town} phone: ${phone} zip: ${zip} `);
    rs.MoveNext();
  }
} finally {
  con.Close();
}
      
