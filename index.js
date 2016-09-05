/* jshint esversion:6 */

(function(){
  'use strict';

  // deps
  const officegen = require('officegen');
  const log4js = require('log4js');
  const fs = require('fs');
  const parse = require('csv-parse');
  const _ = require('lodash');
  const express = require('express');
  const cli = require('commander');

  // column to start reading the CSV from (zero-based)
  const FIRST_COLUMN = 2;
  // source CSV file
  const MATRIX_CSV = 'skillsmatrix_v2.1.csv';

  // configure logger
  const logger = log4js.getLogger();
  logger.setLevel('INFO');

  // configure command line parser
  cli
    .version('0.0.1')
    .option('-f, --fromrole [role]', 'The role, e.g. "Test Engineer"')
    .option('-t, --torole [role]', 'Role to progress to, e.g. "Team Leader"')
    .option('-l, --levelfrom [level]', 'Current level, e.g. "Senior"')
    .option('-p, --progressionlevel [level]', 'Level to progress to, e.g. "Intermediate"')
    .option('-s, --server [port]', 'Run as a web server on specified port', parseInt)
    .parse(process.argv);

  /**
  * @out output stream to write to
  */
  function createDocument(fromrole, levelfrom, torole, progressionlevel, out){

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

          (row[0] === fromrole &&
            row[1] === levelfrom) ||
            (row[0] === torole &&
              row[1] === progressionlevel)
            ){
              // if the role, plus the from level or to level match
              let skillData = {};

              for(var counter = FIRST_COLUMN; counter < row.length; counter++){
                if(row[counter] !== 'N/A'){
                  if(_.indexOf(skills, header[counter]) < 0){
                    logger.debug(`Adding ${header[counter]}`);
                    skills.push(header[counter]);
                  }
                  skillData[header[counter]] = row[counter];
                }
              }

              logger.debug(`skillData: ${JSON.stringify(skillData)}`);

              if(row[1] === levelfrom){
                // it's the from level
                logger.debug('setting to fromData');
                fromData = skillData;
              }
              else if(row[1] === progressionlevel){
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
          if(!progressionlevel){
            // just generate a Profile
            let paragraph = docx.createP();
            paragraph.addText ('Skills Profile', {font_size: 24, bold: true});
            paragraph = docx.createP();
            paragraph.addText (`${fromrole} - ${levelfrom}`, {font_size: 12, bold: true});
            fileName = `Skills Profile ${fromrole} - ${levelfrom}.docx`;

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
              `${fromrole} - ${levelfrom} to ${torole} - ${progressionlevel}`,
              {font_size: 12, bold: true}
            );
            fileName = `Progression Matrix ${fromrole} - ${levelfrom} to ${torole} - ${progressionlevel}.docx`;
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
                table.push([skill, fromSkill, toSkill, '']);
              }
            });
          }



          let tableStyle = {
            tableSize: 24,
            tableFontFamily: "Arial",
            tableAlign: "left",
            borders: true
          };

          docx.createTable (table, tableStyle);

          if(!fromData){
            logger.error(`No level was matched for ${fromrole} - ${levelfrom}`);
            process.exit(2);
          }

          if(!out){
            out = fs.createWriteStream(fileName);
          }
          docx.generate(out);
          out.on('close', function (){
            logger.info('Finished creating file');
          });
        });

        fs.createReadStream(`${__dirname}/${MATRIX_CSV}`).pipe(parser);
      }


      if(cli.server){
        const app = express();
        app.get('/', function (req, res) {
          // TODO this isn't working yet - see Postman
          createDocument(
            req.params.fromrole,
            req.params.levelfrom,
            req.params.torole,
            req.params.progressionlevel,
            res
          );
        });
        app.listen(cli.server, function () {
          logger.info(`vqsfia listening on port ${cli.server}`);
        });
      }
      else{
        createDocument(
          cli.fromrole,
          cli.levelfrom,
          cli.torole,
          cli.progressionlevel
        );
      }


})();
