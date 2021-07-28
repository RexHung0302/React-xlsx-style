import logo from "./logo.svg";
import "./App.css";

import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-style";
import { saveAs } from "file-saver";

console.log(XLSX, XLSXStyle);

const title = ["id", "name", "age", "fav"];
const titleDisplay = { id: "編號", name: "姓名", age: "年齡", fav: "最愛" };

const defaultJson = [
  {
    id: 0,
    name: "張三",
    age: 18,
    fav: "玩遊戲",
  },
  {
    id: 1,
    name: "李四",
    age: 25,
    fav: "爬山",
  },
  {
    id: 2,
    name: "王五",
    age: 60,
    fav: "去夜店",
  },
];

const s2ab = (s) => {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
  return buf;
};

function App() {

  const handleDownloadExcel = () => {
    // 流程
    // 1. 將自定義表頭設定好
    // 2. 使用 xlsx 把 JSON 轉為 Sheet(xlsx-style 沒提供), 後方可帶入 option 設定表頭, 並放入宣告的 worksheet 變數中
    // 3. 修改 worksheet 樣式
    // 4. 宣告 workbook 這邊放入 worksheet 及一些設定
    // 5. 使用 xlsx-style 將 JS 轉為 excel 的格式(其中形式 type 必須設定為 binary, 如果設定為 file 環境必須為 node.js)
    // 6. 透過 s2ab function 將內容轉為 blob 供下載
    // 7. 使用套件 file-server 取代自己手寫 createElement(a) 的程式碼, 透過該套件的 api -> saveAs 去下載剛剛轉好的 blob 檔案
    // 備註 1：套件 xlsx-style 安裝後會出現 ./cptable 找不到的問題，React 請透過 rewired 的套件，
    //        新增 config-overrides.js 去解決，Vue 可透過 Vue.config.js 或 Webpack.config.js
    // 備註 2：Excel 的單位可從大到小分為 Workbook Object > Worksheet Object > Cell Object
    // 詳細可參考：https://zhuanlan.zhihu.com/p/257845606

    // 自訂表頭
    const newTitleData = [titleDisplay, ...defaultJson];

    // 宣告 Worksheet Object
    const worksheet = XLSX.utils.json_to_sheet(newTitleData, {
      header: title,
      skipHeader: true,
    });

    // 操作 Worksheet Object, 並修改裡面的 Cell Object
    worksheet["A1"].s = {
      font: {
        color: { rgb: "FF0187FA" },
        bold: true,
      },
      alignment: {
        horizontal: "center",
      },
    };

    // 宣告 Workbook Object
    const workbook = {
      SheetNames: ["sheet_1"],
      Sheets: {
        sheet_1: worksheet,
      },
    };

    // 設定檔案形式(這邊請避免使用 xlsx 去 write, 下載結束會發現沒有樣式出來)
    const workbookout = XLSXStyle.write(workbook, {
      bookType: "xlsx",
      type: "binary",
    });

    var blob = new Blob([s2ab(workbookout)], {
      type: "application/octet-stream",
    });

    saveAs(blob, `excel - ${new Date()}.xlsx`);
  };

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <button onClick={() => handleDownloadExcel()}>Download Excel</button>
      </header>
    </div>
  );
}

export default App;
