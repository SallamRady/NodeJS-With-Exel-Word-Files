const xlsx = require("xlsx");
const path = require("path");

//read exel file
const wf = xlsx.readFile(path.join("exel", "Friends.xlsx"));
//To get exel file sheet names
console.log(wf.SheetNames);
//to get sheet data
const ws = wf.Sheets["Sheet1"];
//convert exel file to json object then deal with it normally.
let jsonObject = xlsx.utils.sheet_to_json(ws);

jsonObject = jsonObject.map((record) => {
  if (record.Age >= 20 && record.Age < 30) record.AgeLevel = "young man";
  else record.AgeLevel = "man";
  return record;
});

//save json object in sheet in exel file
let neWB = xlsx.utils.book_new();
let neWS = xlsx.utils.json_to_sheet(jsonObject);
// let fileName = "Exel-File-Via-NodeJs-" + Date.now()+".xlsx";
let fileName = path.join(
  "exel",
  "f" + String((Math.random() * 10).toFixed(3)) + ".xlsx"
);
xlsx.utils.book_append_sheet(neWB, neWS);
xlsx.writeFile(neWB, fileName);

