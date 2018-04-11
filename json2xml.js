const commander = require('commander');
const officegen = require('officegen');
const fs = require('fs');
const path = require('path');
const mkdirp = require('mkdirp');
const moment = require('moment');

commander
  .option('-i --input <input>', 'Path to input file. Input file must be a JSON')
  .option('-o --output <output>', 'Path to output directory. File name will be ${date}.xlsx')
  .parse(process.argv);

if (!commander.input) {
  console.log('Input file is required. See `json2xml.js --help`');
  process.exit(1);
}

if (!commander.output) {
  commander.output = __dirname;
}

if (!fs.statSync(commander.output).isDirectory()) {
  console.log('output must be a directory.');
  process.exit(2);
}

const raw = fs.readFileSync(commander.input);
const input = JSON.parse(raw);

const output = [];
for (k of Object.keys(input)) {
  const o = [k]
  for (kk of Object.keys(input[k])) {
    o.push(input[k][kk]);
  }
  output.push(o);
}

const filename = `${moment().format('YYYY-MM-DD')}.xlsx`;
const fullPath = path.join(commander.output, filename);

mkdirp.sync(commander.output);
const xlsx = officegen('xlsx');
const sheet = xlsx.makeNewSheet();
sheet.name = 'data';
sheet.data = output;
const stream = fs.createWriteStream(fullPath);
xlsx.generate(stream);
