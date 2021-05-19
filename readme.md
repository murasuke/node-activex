# node.jsからADO(ActvieX Data Object)を利用する

## 目的

* 10年ぶりにWSHを利用したところ、ふとTypescriptでActiveX使えたら幸せではないか？と思い付き調査
  * node.jsからActiveXを利用するライブラリがあるのではないか？
  * ライブラリがあるなら、TypeScriptでも使えるのではないか？
  * Typescriptが使えるのであれば、ADOの型定義を誰か作っているのではないか？

## 調査結果

nodeでActiveXオブジェクトを生成し、型指定することは可能でした。

* node.jsでActiveXを利用するためには「[winax](https://www.npmjs.com/package/winax)」を使う
* TypeScript用の型定義もnpmに用意されている
  * @types/activex-adodb
  * @types/activex-adox


## 1：javascript(node)でADOを利用

* [winax](https://www.npmjs.com/package/winax)をインストール
```bash
npm i winax
```


### javascriptでADOを利用するサンプル(  mdb.js)

```javascript
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

// ファイルがあれば削除
if (fs.existsSync(data_path)) {
  fs.unlinkSync(data_path);
}

const constr = `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${data_path}`;
const cat = new ActiveXObject('ADOX.Catalog');
cat.Create(constr);
const con = cat.ActiveConnection;
try {
  // データ登録
  con.Execute('create Table persons (Name char(30), City char(30), Phone char(20), Zip decimal(5))');
  con.Execute("insert into persons values('John', 'London','123-45-67','14589')");
  con.Execute("insert into persons values('Andrew', 'Paris','333-44-55','38215')");
  con.Execute("insert into persons values('Romeo', 'Rom','222-33-44','54323')");

  // selectした結果を表示
  const rs = con.Execute('Select * from persons'); 
  const fields = rs.Fields;
  console.log(`Result field count: ${fields.Count}`);

  while (!rs.EOF) {
      // Access as property by string key
      const name = fields['Name'].Value;  
      // Access as method with string argument
      const town = fields('City').value;
      // Access as indexed array
      const phone = fields[2].value;
      // Access recordset
      const zip = rs[3].value;    

      console.log(`> Person: ${name} from ${town} phone: ${phone} zip: ${zip} `);
      rs.MoveNext();
  }
} finally {
  con.Close();
}
```

### 実行と確認(javascript)

  ADOでデータ登録、取得が出来ていることを確認

```bash
> node mdb.js

Result field count: 4
> Person: John                           from London                         phone: 123-45-67            zip: 14589
> Person: Andrew                         from Paris                          phone: 333-44-55            zip: 38215
> Person: Romeo                          from Rom                            phone: 222-33-44            zip: 54323
```

## 2：TypeScript(node)でADOを利用する

typescriptと型定義をインストールしてから初期化(ts-config生成)
```
npm i typescript -D
npm i ts-node
npm i @types/node @types/activex-adodb @types/activex-adox -D

npx tsc -init
```
  コンパイルが面倒なので、ts-nodeをインストールする(typescriptをコンパイルせずに実行できる)

### typescriptに書き直したソース(mdb.ts)

* ADODB.Recordsetなどで型指定行われるため、コーディングが楽になります

```typescript
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
```


### 実行と確認(typescript)

  ADOでデータを登録、取得が出来ていることを確認

```bash
> ts-node mdb.ts

Result field count: 4
> Person: John                           from London                         phone: 123-45-67            zip: 14589
> Person: Andrew                         from Paris                          phone: 333-44-55            zip: 38215
> Person: Romeo                          from Rom                            phone: 222-33-44            zip: 54323
```
