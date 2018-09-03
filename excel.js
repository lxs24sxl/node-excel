let xlsx = require ('node-xlsx');

// console.log( xlsx );

let xlsxData = xlsx.parse ('./assets/test.xlsx');

console.log(JSON.stringify(xlsxData))

