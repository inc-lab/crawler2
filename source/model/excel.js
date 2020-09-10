const path = require('path');
const root_path = path.dirname(require.main.filename || process.mainModule.filename);
const Excel = require('exceljs');
const moment = require('moment');
const { Console } = require('console');
//------------------------------------------------------//
module.exports = (async function(sheet, name) {



    let date = moment().format("YYYY-MM-DD HH-mm-ss");
    let file = root_path + "/tmp/" + " ["+date+"] " + name + " .xlsx";




    let workbook = new Excel.Workbook();
    workbook.creator = 'Robot';
    workbook.lastModifiedBy = 'robot';
    workbook.created = new Date();
    workbook.modified = new Date();
    let worksheet = workbook.addWorksheet('Мониторинг '+ date );

    //Формируем массив ключей
    let keys = {};
    let index = 7; // столбец с которого начинается конкуренты
    //--------------------------------------------------//
    for (let i = 0; i < sheet[0].length; i++) {
        try {
            let key = sheet[0][i].replace(/\s+/g, " ").replace(/^\s|\s$/g, "").toLowerCase();
            keys[ key ] = i;
            if(sheet[0][i]=="*") { index = i; }   
        } catch (error) {}

    }

    //-----------------------  ЗАГОЛОВОК ----------------------------//
    for (let i = 0; i < sheet[0].length; i++) {
        let defaultWidth = 200;
        let defaultHeight = 70;
        worksheet.getRow(1).getCell(i+1).value  = sheet[0][i];
        worksheet.getRow(1).getCell(i+1).style = {  width: defaultWidth , height: defaultHeight };
        worksheet.getRow(1).getCell(i+1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true  };
        worksheet.getRow(1).getCell(i+1).font = { color: { argb: 'FFFFFFFF' },};
        worksheet.getRow(1).getCell(i+1).fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'FF2962FF'},
            bgColor:{argb:'FF2962FF'}
        };
    }
    //-----------------------  ЗАГОЛОВОК ----------------------------//


    //--------------------------------------------------------------------------//
    //-----------------------  ПРОХОДИМ ПО СТРОКАМ ----------------------------//
    for (let i = 1; i < sheet.length; i++) {
        //------------------------------------------------------------------------//
        for (let j = 0;  j <= keys['*']; j++) {
            //-----------------------------------------//
            
            if(typeof sheet[i][j] !== "object"){
              //  console.log(i+'-'+j+':'+sheet[i][j]);
                worksheet.getRow(i+1).getCell(j+1).value  = sheet[i][j];
                if(j==4){

                    worksheet.getRow(i+1).getCell(j+1).fill = {
                        type: 'pattern', pattern:'solid',
                        fgColor:{argb:'FFFFF000'},
                        bgColor:{argb:'FFFFF000'}
                    };
                
                }
                //-----------------------------------------//
                /*if(j==3){
    
                        worksheet.getRow(i+1).getCell(j+1).fill = {
                            type: 'pattern', pattern:'solid',
                            fgColor:{argb:'FFFFF000'},
                            bgColor:{argb:'FFFFF000'}
                        };
                
                }*/
                //-----------------------------------------//
                continue;
            }else{
                worksheet.getRow(i+1).getCell(j+1).value  = "object";
                if(j==4){

                        worksheet.getRow(i+1).getCell(j+1).fill = {
                            type: 'pattern', pattern:'solid',
                            fgColor:{argb:'FFFFF000'},
                            bgColor:{argb:'FFFFF000'}
                        };
                
                }
                //-----------------------------------------//
               /* if(j==3){
    
                        worksheet.getRow(i+1).getCell(j+1).fill = {
                            type: 'pattern', pattern:'solid',
                            fgColor:{argb:'FFFFF000'},
                            bgColor:{argb:'FFFFF000'}
                        };
                
                }*/
                //-----------------------------------------//
               
            }

        }
        //Поля для конкурента
        //------------------------------------------------------------------------//
        for (let j = index+1; j < sheet[i].length; j++) {
            if(typeof sheet[i][j] !== "object"){
                worksheet.getRow(i+1).getCell(j+1).value  = sheet[i][j];
                continue;
            }
            if( sheet[i][j]['status'] == "ok" ){
                var pricemain = sheet[i][4].replace(/[^0-9]+/g,'');
                if(pricemain=='' || isNaN(pricemain)){
                    pricemain = sheet[i][7]['price'];
                    pricemain=pricemain.toString().replace(/[^0-9]+/g,'');
                }
                console.log(i+'--'+j+'--'+parseInt(sheet[i][j]['price'])+'--'+parseInt(pricemain)+'--'+sheet[i][j]['price']+'--'+pricemain);
                let price = sheet[i][j]['price'] ? sheet[i][j]['price'] : "0";
                worksheet.getRow(i+1).getCell(j+1).value  = { text: price, hyperlink: sheet[i][j]['url'] };
                if( parseInt(sheet[i][j]['price'])==0){
                    worksheet.getRow(i+1).getCell(j+1).fill = {
                        type: 'pattern', pattern:'solid',
                        fgColor:{argb:'FF4DB6AC'},
                        bgColor:{argb:'FF4DB6AC'}
                    };
                }
                
               else if( parseInt(sheet[i][j]['price'])>parseInt(pricemain)){
                    worksheet.getRow(i+1).getCell(j+1).fill = {
                        type: 'pattern', pattern:'solid',
                        fgColor:{argb:'FF4DB6AC'},
                        bgColor:{argb:'FFСССССС'}
                    };
                }
                else if(parseInt(sheet[i][j]['price'])==parseInt(pricemain)){
                    worksheet.getRow(i+1).getCell(j+1).fill = {
                        type: 'pattern', pattern:'solid',
                        fgColor:{argb:'FFFFF000'},
                        bgColor:{argb:'FFFFF000'}
                    };
                }
                else{
                    worksheet.getRow(i+1).getCell(j+1).fill = {
                        type: 'pattern', pattern:'solid',
                        fgColor:{argb:'FFFF8A65'},
                        bgColor:{argb:'FFСССССС'}
                    }; 
                }
                //-----------------------------------------------//
            }else{
                worksheet.getRow(i+1).getCell(j+1).value  = { text: "error", hyperlink: sheet[i][j]['url'] };
            }
        }
        //------------------------------------------------------------------------//



        //------------------------------------------------------------------------//
    }




    //----------------------/  ПРОХОДИМЯ ПО СТРОКАМ ----------------------------//
    //--------------------------------------------------------------------------//


    await workbook.xlsx.writeFile(file);
    return file;
});
//------------------------------------------------------//