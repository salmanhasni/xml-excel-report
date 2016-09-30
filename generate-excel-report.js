var excelBuilder = require('msexcel-builder-colorfix'),
    fs = require('fs'),
    nodemailer = require('nodemailer'),
    parser = require('xml2json');

const CONFIG = require('./config.json');

function generateMapping(CONFIG){
    var date = new Date();
    var fName = CONFIG.fileName ? CONFIG.fileName : date + '.xlsx';
    var workbook = excelBuilder.createWorkbook(CONFIG.OutputFolderName, fName);

    var sheet = workbook.createSheet(CONFIG.sheet.name, CONFIG.sheet.totalCols, CONFIG.sheet.totalRows);

    var startingRow = CONFIG.startingRow;
    var colorList = CONFIG.color.list;

    startingRow = setTitle(sheet, CONFIG.sheetTitle, {col:1,row:startingRow},{col:3,row:startingRow},
        startingRow, colorList.pop());
    startingRow++;
    setFields(sheet,
        CONFIG.sheetFields,
        ++startingRow);

    var xmlConclusion = parseTestResult(CONFIG.inputFolderName);

    startingRow = fillData(sheet, xmlConclusion, ++startingRow, colorList);
    ++startingRow;
    addSuccessRate(sheet, xmlConclusion, {col:1,row:startingRow},{col:3,row:startingRow}, startingRow);

    workbook.save(function(error){
        if (error){
            workbook.cancel();
        }
        else
        {
            console.log('Mapping Successful');
            attachAndSendReport(CONFIG, workbook.fpath + workbook.fname);
        }
    });
}

function fillData(sheet, xmlConclusion, startingRow, colorList){

    for (var title in xmlConclusion) {
        var dataColor = colorList.pop();

        startingRow = setSubTitle(sheet, title, {col:1,row:startingRow},{col:3,row:startingRow},
            startingRow, dataColor);
        startingRow++;

        var xmlConclusionData = xmlConclusion[title];

        var lightColor = shadeColor("#" + dataColor, 0.80);

        for(var i = 0; i < xmlConclusionData.length; i++){
            sheet.font(1, startingRow + i, {name:CONFIG.fontName,sz:'10'});
            sheet.fill(1, startingRow + i, {type:'bold',fgColor: lightColor});
            sheet.border(1, startingRow + i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
            sheet.set(1, startingRow + i, i + 1);
            sheet.align(1, startingRow + i, 'left');

            sheet.font(2, startingRow + i, {name:CONFIG.fontName,sz:'10'});
            sheet.fill(2, startingRow + i, {type:'bold',fgColor: lightColor});
            sheet.border(2, startingRow + i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
            var textToRemove = CONFIG.textToRemoveRegex ? CONFIG.textToRemoveRegex : "" ;
            var textToRemoveRegex = new RegExp(textToRemove, 'g');
            sheet.set(2, startingRow + i, xmlConclusionData[i].name.replace(textToRemoveRegex, ''));
            sheet.align(2, startingRow + i, 'left');

            sheet.font(3, startingRow + i, {name:CONFIG.fontName,sz:'10'});
            sheet.border(3, startingRow + i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
            var color = "";
            color = xmlConclusionData[i].status ? CONFIG.color.pass : CONFIG.color.fail ;
            sheet.fill(3, startingRow + i, {type:'bold',fgColor: color});
            sheet.align(3, startingRow + i, 'left');
        }

        startingRow += xmlConclusionData.length;
        colorList.unshift(dataColor);
    }

    return startingRow;
}

function setTitle(sheet, title, from, to, startingRow, color){
    sheet.font(1, startingRow, {name:CONFIG.fontName,sz:'14',bold:'true'});
    sheet.height(1, 20);
    sheet.set(1, startingRow, title);
    sheet.fill(1, startingRow, {type:'bold',fgColor: color});
    sheet.merge(from, to);
    sheet.align(1,startingRow, 'center');
    sheet.border(from.col, from.row, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
    sheet.border(to.col, to.row, {left:'thin',top:'thin',right:'thin',bottom:'thin'});

    return startingRow;
}

function setSubTitle(sheet, title, from, to, startingRow, color){
    sheet.font(1, startingRow, {name:CONFIG.fontName,sz:'10'});
    sheet.height(1, 25);
    sheet.set(1, startingRow, title);
    sheet.fill(1, startingRow, {type:'bold',fgColor: color});
    sheet.merge(from, to);
    sheet.align(1,startingRow, 'center');
    sheet.border(from.col, from.row, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
    sheet.border(to.col, to.row, {left:'thin',top:'thin',right:'thin',bottom:'thin'});

    return startingRow;
}

function setFields(sheet, fields, startingRow){
    for (var i = 0; i < fields.length; i++){
        sheet.font(i+1, startingRow, {name:CONFIG.fontName,sz:'10', bold: true});
        sheet.set(i+1, startingRow, fields[i].name);
        sheet.border(i+1, startingRow , {left:'thin',top:'thin',right:'thin',bottom:'thin'});
        sheet.width(i+1, fields[i].width);
        sheet.align(i+1,startingRow, 'left');
    }

    return startingRow;
}

function parseTestResult(path){
    var files = fs.readdirSync(path);
    var classConclusion = new Object();

    for(var i=0;i<files.length; i++){
        var filePath = path + "/" + files[i];
        if(fs.lstatSync(filePath).isFile()){
            var data = fs.readFileSync(filePath);

            var result = JSON.parse(parser.toJson(data));

            if(result && result.testsuite && result.testsuite.testcase){
                var fileNamesArray = files[i].split(".");
                var className = fileNamesArray[fileNamesArray.length-2];
                var tests = result.testsuite.testcase;
                var xmlConclusion = [];
                for(var t = 0; t<tests.length; t++){
                    var test = tests[t];
                    if(!test.failure){
                        xmlConclusion.push({status: true, name: test.name});
                    } else {
                        xmlConclusion.push({status: false, name: test.name});
                    }
                }
                classConclusion[className] =  xmlConclusion;
            }
        }
    }

    return classConclusion;
}

function addSuccessRate(sheet, xmlConclusion, from, to, startingRow){

    var totalTest = 0;
    var passedTest = 0;

    for (var title in xmlConclusion) {

        var xmlConclusionData = xmlConclusion[title];

        for(var i = 0; i<xmlConclusionData.length; i++){
            totalTest++;
            if(xmlConclusionData[i].status)
                passedTest++
        }
    }

    sheet.font(1, startingRow, {name:CONFIG.fontName,sz:'20',bold:'true'});
    sheet.height(startingRow, 40);
    var successPercentage = (passedTest/totalTest) * 100;
    sheet.set(1, startingRow, "Success Rate: " + successPercentage.toFixed(2) + "%");
    sheet.merge(from, to);
    sheet.align(1,startingRow, 'center');

    return startingRow;
}

function attachAndSendReport(CONFIG, path){
    var nodemailer = require('nodemailer');

    var transporter = nodemailer.createTransport('smtps://tiggee.test.email@gmail.com:tenpearls@smtp.gmail.com');

    var mailOptions = {
        from: '"Selenium Automation " <tiggee.test.email@gmail.com>', // sender address
        to: CONFIG.emails.to.join(", "), // list of receivers
        subject: CONFIG.emails.subject, // Subject line
        attachments: [ {path: path} ]
    };

    transporter.sendMail(mailOptions, function(error, info){
        if(error){
            return console.log(error);
        }
        console.log('Message sent: ' + info.response);
    });
}

function shadeColor(color, percent) {
    var f=parseInt(color.slice(1),16),t=percent<0?0:255,p=percent<0?percent*-1:percent,R=f>>16,G=f>>8&0x00FF,B=f&0x0000FF;
    return "" + (0x1000000+(Math.round((t-R)*p)+R)*0x10000+(Math.round((t-G)*p)+G)*0x100+(Math.round((t-B)*p)+B)).toString(16).slice(1);
}

generateMapping(CONFIG);
