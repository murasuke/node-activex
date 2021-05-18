# node.jsからADO(AtvieX Data Object)を利用する

## 目的

* 10年ぶりにWSHを利用したところ、ふとTypescriptでActiveX使えたら幸せではないか？と思い付き調査を開始
  * node.jsからActiveXを利用するライブラリがあるのではないか？
  * ライブラリがあるなら、TypeScriptでも使えるのではないか？
  * Typescriptが使えるのであれば、ADOの型定義を誰か作っているのではないか？

## 現状

* node.jsでActiveXを利用する「winax」を利用してADO利用できるところまで確認

```bash
npm i winax
```

```javascript
const path = require('path'); 
require('winax');  // npm i winax

// MDBファイルを作成する
// 要 Microsoft Access Database Engine 2016 Redistributable
// https://www.microsoft.com/en-us/download/details.aspx?id=54920
const filename = 'persons.mdb';
const data_path = path.join(__dirname, '/data/', filename);

const constr = 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + data_path;
const cat = new ActiveXObject('ADOX.Catalog');
cat.Create(constr);
const con = cat.ActiveConnection
con.Execute('create Table persons (Name char(50), City char(50), Phone char(20), Zip decimal(5))');
con.Execute("insert into persons values('John', 'London','123-45-67','14589')");
```
