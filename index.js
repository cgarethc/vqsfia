/* jshint esversion:6 */

(function(){
'use strict';

// deps
const officegen = require('officegen');
const log4js = require('log4js');
const fs = require('fs');
const parse = require('csv-parse');
const commandline = require('commandline-parser');

// column to start reading the CSV from (zero-based)
const FIRST_COLUMN = 2;
// source CSV file
const MATRIX_CSV = 'skillsmatrix_v2.1.csv';

// configure logger
var logger = log4js.getLogger();
logger.setLevel('DEBUG');

// configure command line parser
var clparser = new commandline.Parser({
  name : "vqSFIA",
  desc : 'SFIA progression process form generation'
});
clparser.addArgument('help' ,{
	flags : ['h','help'],
	desc : "show help",
	optional : true,
	action : function(value, parser){
    parser.printHelp();
    process.exit(0);
  }
});
clparser.addArgument('role' ,{
	flags : ['r','role'],
	desc : 'specify a role',
	optional : true,
});
clparser.addArgument('level' ,{
	flags : ['l','level'],
	desc : 'specify a level',
	optional : true,
});
clparser.addArgument('progression' ,{
	flags : ['p','progession'],
	desc : 'specify a level to progress to, if not provided a skills profile will be produced',
	optional : true,
});
clparser.exec();

let docx = officegen( 'docx' );

let parser = parse();
let header;
let fileName;
// construct the table header, setting the column widths in pixels
let table = [
  [
    {
      val:'Skill',
      opts: {
        b:true,
        cellColWidth: 800
      }
    },
    {
      val:'Requirement',
      opts: {
        b:true,
        cellColWidth: 6000
      }
    }
  ]
  ];

  // parse
  parser.on('data', function(row){
    if(!header){
      // read the CSV header
      header = row;
      logger.debug(header);

      // output a heading to the document
      let paragraph = docx.createP();
      paragraph.addText ('Skills Profile', {font_size: 24, bold: true});
    }
    else if((row[0] === clparser.get('role') && row[1] === clparser.get('level')) && !clparser.get('progression')){
      // create a skills profile table row for the matching role and level
      logger.debug(row);
      let paragraph = docx.createP();
      paragraph.addText (`${row[0]} - ${row[1]}`, {font_size: 12, bold: true});
      fileName = `Skills Profile ${row[0]} - ${row[1]}.docx`;

      for(var counter = FIRST_COLUMN; counter < row.length; counter++){
        if(row[counter] !== 'N/A'){
          table.push([header[counter], row[counter]]);
        }
      }
    }
  });

// Catch any error
parser.on('error', function(err){
  logger.error(err);
});

parser.on('finish', function(){
  logger.info('Done');
  let out = fs.createWriteStream(fileName);

  let tableStyle = {
    tableSize: 24,
    tableFontFamily: "Arial",
    tableAlign: "left",
    borders: true
  };

  docx.createTable (table, tableStyle);

  docx.generate( out );
  out.on('close', function (){
    logger.info('Finished creating file');
  });
});

fs.createReadStream(`${__dirname}/${MATRIX_CSV}`).pipe(parser);

})();
