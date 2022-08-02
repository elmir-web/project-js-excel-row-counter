let xlsx = require('xlsx');

let wb = null;
let ws = null;
let data = null;
let count = null;

wb = xlsx.readFile('test1.XLSX', { cellDates: true });

ws = wb.Sheets['Sheet1'];

data = xlsx.utils.sheet_to_json(ws);

count = 0;

let newData = null;

newData = data.map((item) => {
  let tempNewItem = {
    НОМЕРСТРОКИ: '',
    ...item,
  };

  if (
    (tempNewItem['Имя пользователя'] === undefined ||
      !tempNewItem['Имя пользователя']) && // !СМОТРИМ, чтобы имя пользователя или пустое или не определено
    tempNewItem['№ документа счета'] !== '' // !И одновременно ПРИ ЭТОМ номер документа пустой
  ) {
    count++;

    tempNewItem[`НОМЕРСТРОКИ`] = `${count}`;
  }

  return tempNewItem;
});

let newWB = null;
let newWS = null;

newWB = xlsx.utils.book_new();

try {
  newWS = xlsx.utils.json_to_sheet(
    // [{ test: 1, aza: 2 }]
    newData
  );
} catch (error) {
  console.log(`Ошибка 11 (JSON to XLSX):`);
  console.log(error);
}

xlsx.utils.book_append_sheet(newWB, newWS, 'New Data');

xlsx.writeFile(newWB, 'NewFile.xlsx');

console.log(`Я обработовал: ${count}`);
console.log(`Я ВСЕ!!!11`);
