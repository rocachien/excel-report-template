/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, __dirname, describe, before, it */
"use strict";

var buster       = require('buster'),
    XlsxTemplate = require('../lib'),
    fs           = require('fs'),
    path         = require('path'),
    etree        = require('elementtree');

buster.spec.expose();
buster.testRunner.timeout = 500;

function getSharedString(sharedStrings, sheet1, index) {
    return sharedStrings.findall("./si")[
        parseInt(sheet1.find("./sheetData/row/c[@r='" + index + "']/v").text, 10)
    ].find("t").text;
}

describe("CRUD operations", function() {

    before(function(done) {
        done();
    });

    describe('Excel Report Template', function() {

        it("can substitute values and generate a file", function(done) {

            fs.readFile(path.join(__dirname, 'templates', 'report.xlsx'), function(err, data) {
                buster.expect(err).toBeNull();

                var t = new XlsxTemplate(data);
                var d = [
                    {
                        'Company': 'Công ty CP DigiNet',
                        'Address': '341 Điện Biên Phủ, Quận Bình Thạnh',
                        'TaskID': '17TA00000000023',
                        'TaskNo': '',
                        'SalesPersonID': 'HONGTRANG',
                        'CompanyID': '',
                        'CompanyTypeID': '',
                        'ContactID': '',
                        'AddressID': '',
                        'TaskType': '',
                        'TaskDetails': 'Caäp nhaät coâng vieäc HT',
                        'TaskStatus': 'NEW',
                        'Priority': '',
                        'TaskDateFrom': '2008-10-14T00:00:00.000Z',
                        'TaskTimeFrom': '0920',
                        'TaskDateTo': '2008-10-14T00:00:00.000Z',
                        'TaskTimeTo': '1140',
                        'IsPrivate': 0,
                        'CreateUserID': 'LEMONADMIN',
                        'CreateDate': '2008-10-14T08:04:38.420Z',
                        'LastModifyUserID': 'LEMONADMIN',
                        'LastModifyDate': '2008-10-14T09:15:33.687Z',
                        'UserID': 'LEMONADMIN',
                        'Title': '2345 dA SUA',
                        'Results': '',
                        'Deadline': '2008-10-14T08:04:38.420Z',
                        'Status': 0,
                        'Type': '',
                        'FinishDate': null,
                        'Notes': '',
                        'NotesU': '',
                        'TitleU': 'Hẹn gặp khách hàng ABC, họp báo giá lần 1',
                        'ResultsU': '',
                        'TaskDetailsU': 'Cập nhật công việc HT',
                        'CaseID': '',
                        'FinishPercent': 0,
                        'ActEvaluation': '',
                        'ActEvaluationU': '',
                        'TaskColor': '',
                        'TaskRelatedTypeID': '',
                        'TaskRelatedID': '',
                        'TaskContactTypeID': '',
                        'TaskStatusU': '',
                        'LinkTaskID': '',
                        'TaskDisplayOrder': 0,
                        'AssignerTaskID': '',
                        'Location': '341-343 Điện Biên Phủ Phường 15 Quận Bình Thạnh TP Hồ Chí Minh',
                        'ScheduleType': ''
                    },
                    {
                        'Company': 'Công ty CP DigiNet',
                        'Address': '341 Điện Biên Phủ, Quận Bình Thạnh',
                        'TaskID': '17TA0J000000091',
                        'TaskNo': '',
                        'SalesPersonID': 'A004',
                        'CompanyID': null,
                        'CompanyTypeID': 'TN',
                        'ContactID': '',
                        'AddressID': '',
                        'TaskType': '',
                        'TaskDetails': 'DFSA',
                        'TaskStatus': 'NEW',
                        'Priority': '',
                        'TaskDateFrom': '2012-05-09T00:00:00.000Z',
                        'TaskTimeFrom': '1020',
                        'TaskDateTo': '2012-05-09T00:00:00.000Z',
                        'TaskTimeTo': '1320',
                        'IsPrivate': 0,
                        'CreateUserID': 'LEMONADMIN',
                        'CreateDate': '2012-05-09T10:20:36.583Z',
                        'LastModifyUserID': 'LEMONADMIN',
                        'LastModifyDate': '2012-08-30T13:43:55.393Z',
                        'UserID': 'LEMONADMIN',
                        'Title': 'Kieàu test',
                        'Results': 'DFS',
                        'Deadline': null,
                        'Status': 0,
                        'Type': '',
                        'FinishDate': null,
                        'Notes': 'SFAD',
                        'NotesU': 'SFAD',
                        'TitleU': 'Hẹn gặp khách hàng ABC, họp báo giá lần 1',
                        'ResultsU': 'DFS',
                        'TaskDetailsU': 'DFSA',
                        'CaseID': '',
                        'FinishPercent': 0,
                        'ActEvaluation': 'FSD',
                        'ActEvaluationU': 'FSD',
                        'TaskColor': '',
                        'TaskRelatedTypeID': 'TN',
                        'TaskRelatedID': '000091BT',
                        'TaskContactTypeID': '',
                        'TaskStatusU': '',
                        'LinkTaskID': '',
                        'TaskDisplayOrder': 0,
                        'AssignerTaskID': '',
                        'Location': '341-343 Điện Biên Phủ Phường 15 Quận Bình Thạnh TP Hồ Chí Minh',
                        'ScheduleType': ''
                    },
                    {
                        'Company': 'Công ty CP DigiNet',
                        'Address': '341 Điện Biên Phủ, Quận Bình Thạnh',
                        'TaskID': '17TA0J000000092',
                        'TaskNo': '',
                        'SalesPersonID': 'A001',
                        'CompanyID': null,
                        'CompanyTypeID': 'TN',
                        'ContactID': '',
                        'AddressID': '',
                        'TaskType': '',
                        'TaskDetails': 'GFH',
                        'TaskStatus': 'NEW',
                        'Priority': '',
                        'TaskDateFrom': '2012-05-09T00:00:00.000Z',
                        'TaskTimeFrom': '1025',
                        'TaskDateTo': null,
                        'TaskTimeTo': '1025',
                        'IsPrivate': 0,
                        'CreateUserID': 'LEMONADMIN',
                        'CreateDate': '2012-05-09T10:25:57.997Z',
                        'LastModifyUserID': 'LEMONADMIN',
                        'LastModifyDate': '2012-05-09T10:25:57.997Z',
                        'UserID': 'LEMONADMIN',
                        'Title': 'HFG',
                        'Results': 'HGF',
                        'Deadline': null,
                        'Status': 0,
                        'Type': '',
                        'FinishDate': null,
                        'Notes': 'HT',
                        'NotesU': 'HT',
                        'TitleU': 'Hẹn gặp khách hàng ABC, họp báo giá lần 1',
                        'ResultsU': 'HGF',
                        'TaskDetailsU': 'GFH',
                        'CaseID': '',
                        'FinishPercent': 0,
                        'ActEvaluation': 'HGF',
                        'ActEvaluationU': 'HGF',
                        'TaskColor': '',
                        'TaskRelatedTypeID': 'TN',
                        'TaskRelatedID': '0033DAD',
                        'TaskContactTypeID': '',
                        'TaskStatusU': '',
                        'LinkTaskID': '',
                        'TaskDisplayOrder': 0,
                        'AssignerTaskID': '',
                        'Location': '341-343 Điện Biên Phủ Phường 15 Quận Bình Thạnh TP Hồ Chí Minh',
                        'ScheduleType': ''
                    }
                ];

                t.substitute(1, d);

                var newData = t.generate();

                var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
                    sheet1        = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

                // planData placeholder - added rows and cells
                buster.expect(sheet1.find("./sheetData/row/c[@r='A13']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='A13']/v").text, 10)
                    ].find("t").text
                ).toEqual("17TA0J000000091");
                buster.expect(sheet1.find("./sheetData/row/c[@r='A14']").attrib.t).toEqual("s");
                buster.expect(
                    sharedStrings.findall("./si")[
                        parseInt(sheet1.find("./sheetData/row/c[@r='A14']/v").text, 10)
                    ].find("t").text
                ).toEqual("17TA0J000000092");

                buster.expect(sheet1.find("./sheetData/row/c[@r='A13']/v").text).toEqual("17TA0J000000091");
                buster.expect(sheet1.find("./sheetData/row/c[@r='A13']/v").text).toEqual("17TA0J000000092");


                // XXX: For debugging only
                fs.writeFileSync('test/output/report.xlsx', newData, 'binary');

                done();
            });

        });
    });
});
