/* jshint esversion:6 */

(function(){
  'use strict';

  // deps
  const officegen = require('officegen');
  const log4js = require('log4js');
  const fs = require('fs');
  const parse = require('csv-parse');
  const commandline = require('commandline-parser');
  const _ = require('lodash');

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
  clparser.addArgument('fromrole' ,{
    flags : ['f','fromrole'],
    desc : 'specify a role',
    optional : true,
  });
  clparser.addArgument('torole' ,{
    flags : ['t','torole'],
    desc : 'specify a role to progress to',
    optional : true,
  });
  clparser.addArgument('levelfrom' ,{
    flags : ['l','level'],
    desc : 'specify a current level',
    optional : true,
  });
  clparser.addArgument('progressionlevel' ,{
    flags : ['p','progessionlevel'],
    desc : 'specify a level to progress to, if not provided a skills profile will be produced',
    optional : true,
  });
  clparser.exec();

  let docx = officegen( 'docx' );

  let parser = parse();
  let header;
  let fileName;

  let fromData;
  let toData;
  let skills = [];

  // parse
  parser.on('data', function(row){
    if(!header){
      // read the CSV header
      header = row;
      logger.debug(header);


    }
    else{
      if(
        (row[0] === clparser.get('fromrole') &&
        row[1] === clparser.get('levelfrom')) ||
        (row[0] === clparser.get('torole') &&
        row[1] === clparser.get('progressionlevel'))
      ){
        // if the role, plus the from level or to level match
        let skillData = {};

        for(var counter = FIRST_COLUMN; counter < row.length; counter++){
          if(row[counter] !== 'N/A'){
            skills.push(header[counter]);
            skillData[header[counter]] = row[counter];
          }
        }

        logger.debug(`skillData: ${JSON.stringify(skillData)}`);

        if(row[1] === clparser.get('levelfrom')){
          // it's the from level
          logger.debug('setting to fromData');
          fromData = skillData;
        }
        else if(row[1] === clparser.get('progressionlevel')){
          // it's the to level
          logger.debug('setting to toData');
          toData = skillData;
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

    let table;

    // output a heading to the document and set filename
    if(!clparser.get('progressionlevel')){
      // just generate a Profile
      let paragraph = docx.createP();
      paragraph.addText ('Skills Profile', {font_size: 24, bold: true});
      paragraph = docx.createP();
      paragraph.addText (`${clparser.get('fromrole')} - ${clparser.get('levelfrom')}`, {font_size: 12, bold: true});
      fileName = `Skills Profile ${clparser.get('fromrole')} - ${clparser.get('levelfrom')}.docx`;

      table = [
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

      _.forEach(fromData, function(value, key){
        logger.debug(`Adding to table: ${key} = ${value}`);
        table.push([key, value]);
      });
    }
    else{
      // generate a matrix
      let paragraph = docx.createP();
      paragraph.addText ('Progression Matrix', {font_size: 24, bold: true});
      paragraph = docx.createP();
      paragraph.addText (
        `${clparser.get('fromrole')} - ${clparser.get('levelfrom')} to ${clparser.get('torole')} - ${clparser.get('progressionlevel')}`,
        {font_size: 12, bold: true}
      );
      fileName = `Progression Matrix ${clparser.get('fromrole')} - ${clparser.get('levelfrom')} to ${clparser.get('torole')} - ${clparser.get('progressionlevel')}.docx`;
      table = [
        [
          {
            val:'Skill',
            opts: {
              b:true,
              cellColWidth: 400
            }
          },
          {
            val:'Current',
            opts: {
              b:true,
              cellColWidth: 2000
            }
          },
          {
            val:'Requirement',
            opts: {
              b:true,
              cellColWidth: 2000
            }
          },
          {
            val:'Evidence',
            opts: {
              b:true,
              cellColWidth: 2000
            }
          }
        ]
      ];

      skills.forEach(function(skill){
        logger.debug(skill);
        let fromSkill = '';
        let toSkill = '';
        if(fromData[skill]){
          fromSkill = fromData[skill];
        }
        if(toData[skill]){
          toSkill = toData[skill];
        }
        if(fromSkill !== toSkill){
          table.push([skill, fromSkill, toSkill]);
        }
      });
    }

    let out = fs.createWriteStream(fileName);

    let tableStyle = {
      tableSize: 24,
      tableFontFamily: "Arial",
      tableAlign: "left",
      borders: true
    };

    docx.createTable (table, tableStyle);

    if(!fromData){
      logger.error(`No level was matched for ${clparser.get('fromrole')} - ${clparser.get('levelfrom')}`);
      System.exit(2);
    }

    docx.generate( out );
    out.on('close', function (){
      logger.info('Finished creating file');
    });
  });

  fs.createReadStream(`${__dirname}/${MATRIX_CSV}`).pipe(parser);

})();
