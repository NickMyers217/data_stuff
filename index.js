const readline = require('readline');
const fs = require('fs');
const path = require('path');

const xlsx = require('node-xlsx');

const config = require('./config.json');


const inputSheet = path.join('./', 'data', config.inputSpreadsheet);
const hcadData = path.join('./', 'data', config.hcadData);

console.log('Loading spreadsheet', inputSheet, '...');
const workSheetsFromFile = xlsx.parse(inputSheet);
const data = workSheetsFromFile.slice(0, 1).shift().data;
console.log('Loaded!');

const convertDataToJson = (data) => {
  const columns = data.slice(0, 1).shift();
  const rows = data.slice(1);

  return rows.map(row => ({
    lastName: row[0],
    address: row[1],
    city: row[2],
    state: row[3],
    zip: row[4]
  }));
};

const convertResultsToXls = (results) => {
  const columns = [['HCADData','LastName','Address','City','State','ZipCode']];
  return columns.concat(results.map(result => {
    return [
      result.hcadData ? result.hcadData: 'NULL',
      result.lastName,
      result.address,
      result.city,
      result.state,
      result.zip
    ];
  }));
};

console.log('Parsing input spreadsheet...');
const people = convertDataToJson(data);
console.log('Parsing complete!');

console.log('Begin processing...');

const rl = readline.createInterface({
  input: fs.createReadStream(hcadData),
  crlfDelay: Infinity
});

let results = [];
rl.on('line', (line) => {
  const values = line.split('\t');
  const searchText = values.slice(2).join(', ');

  const matches = people.filter(person => {
    return searchText.toLowerCase().indexOf(person.address.toLowerCase()) > -1;
  });

  if (matches.length > 0) {
    console.log(`Found match for ${matches[0].address} ==> ${searchText}...`);
    const result = Object.assign({}, matches[0], { hcadData: searchText });
    results.push(result);
  }
}).on('close', () => {
  console.log('Complete!')
  const buffer = xlsx.build([{name: 'Sheet1', data: convertResultsToXls(results)}]);

  console.log('Building output file!');
  fs.writeFile(path.join('./', 'data', 'output.xlsx'), buffer, 'binary', function (err) {
    if (err) {
      console.log(err);
    } else {
      console.log('Complete!');
    }
  });
});