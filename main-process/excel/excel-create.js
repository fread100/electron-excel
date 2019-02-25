// example : https://www.npmjs.com/package/excel4node
function initExcel(box){
    // Require library
    var excel = require('excel4node');
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('Sheet 1');
    var worksheet2 = workbook.addWorksheet('Sheet 2');
    // Create a reusable style
    var top_style = workbook.createStyle({
        font: {
            size: 22
        },
        alignment: {
            wrapText: true,
            horizontal: 'center',
            vertical: 'center'
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: 'double' //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
            },
            right: {
                style: 'double'
            },
            top: {
                style: 'double'
            },
            bottom: {
                style: 'thin'
            }
        }
    });
    var style = workbook.createStyle({
        font: {
            size: 22
        },
        alignment: {
            wrapText: true,
            horizontal: 'center',
            vertical: 'center'
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: 'double' //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
            },
            right: {
                style: 'double'
            },
            top: {
                style: 'thin'
            },
            bottom: {
                style: 'thin'
            }
        }
    });
    var inputStyle = workbook.createStyle({
        font: {
            size: 10
        },
        alignment: {
            wrapText: true,
            horizontal: 'center',
            vertical: 'center'
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: 'double' //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
            },
            right: {
                style: 'double'
            },
            top: {
                style: 'thin'
            },
            bottom: {
                style: 'thin'
            }
        }
    });
    var bottom_inputStyle = workbook.createStyle({
        font: {
            size: 10
        },
        alignment: {
            wrapText: true,
            horizontal: 'center',
            vertical: 'center'
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: 'double' //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
            },
            right: {
                style: 'double'
            },
            top: {
                style: 'thin'
            },
            bottom: {
                style: 'double'
            }
        }

    });
    var all_style = workbook.createStyle({
        border: {
            left: {
                style: 'double'
            },
            right: {
                style: 'double'
            },
            top: {
                style: 'double'
            },
            bottom: {
                style: 'double'
            }
        }
    });
    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    //worksheet.cell(1,1).number(100).style(style);

    for(var i=0; i<box.length; i++){
        //console.log(box[i]);
        //console.log($(box[i]).find(".sigong").val());
        var gongjeong1 = $(box[i]).find(".gongjeong1").val();
        var sigong1 = $(box[i]).find(".sigong1").val();
        var postion1 = $(box[i]).find(".position1").val();
        var imgsrc1 = $(box[i]).find(".drag-file-input").eq(0).val();

        var gongjeong2 = $(box[i]).find(".gongjeong2").val();
        var sigong2 = $(box[i]).find(".sigong2").val();
        var postion2 = $(box[i]).find(".position2").val();
        var imgsrc2 = $(box[i]).find(".drag-file-input").eq(1).val();

        var gongjeong3 = $(box[i]).find(".gongjeong3").val();
        var sigong3 = $(box[i]).find(".sigong3").val();
        var postion3 = $(box[i]).find(".position3").val();
        var imgsrc3 = $(box[i]).find(".drag-file-input").eq(2).val();

        var cnt18 = 1;
        for(var j=0; j < 3; j++){
            worksheet.row(i*18+(cnt18++)).setHeight(32);
            worksheet.row(i*18+(cnt18++)).setHeight(50);
            worksheet.row(i*18+(cnt18++)).setHeight(32);
            worksheet.row(i*18+(cnt18++)).setHeight(50);
            worksheet.row(i*18+(cnt18++)).setHeight(32);
            worksheet.row(i*18+(cnt18++)).setHeight(50);
        }


        worksheet.column(1).setWidth(19);
        //worksheet.row(2).setHeight(33);
        worksheet.column(2).setWidth(60);
        cnt18 = 1;

        worksheet.cell(i*18+(cnt18++),1).string('공종').style(top_style);
        worksheet.cell(i*18+(cnt18++),1).string(gongjeong1).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(sigong1).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('위치').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(postion1).style(bottom_inputStyle);

        worksheet.cell(i*18+(cnt18++),1).string('공종').style(top_style);
        worksheet.cell(i*18+(cnt18++),1).string(gongjeong2).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(sigong2).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('위치').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(postion2).style(bottom_inputStyle);

        worksheet.cell(i*18+(cnt18++),1).string('공종').style(top_style);
        worksheet.cell(i*18+(cnt18++),1).string(gongjeong3).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(sigong3).style(inputStyle);
        worksheet.cell(i*18+(cnt18++),1).string('위치').style(style);
        worksheet.cell(i*18+(cnt18++),1).string(postion3).style(bottom_inputStyle);

        worksheet.cell((i*18)+1, 2, (i*18)+6, 2, true).style(all_style);
        worksheet.cell((i*18)+7, 2, (i*18)+12, 2, true).style(all_style);
        worksheet.cell((i*18)+13, 2, (i*18)+18, 2, true).style(all_style);
        console.log("imgsrc1", imgsrc1);
        console.log("imgsrc2", imgsrc2);
        console.log("imgsrc3", imgsrc3);
        if(imgsrc1){
            worksheet.addImage({
                path: imgsrc1,
                type: 'picture',
                position: {
                    type: 'twoCellAnchor',
                    from: {
                        col: 2,
                        colOff: 0,
                        row: (i*18)+1,
                        rowOff: 0,
                    },
                    to: {
                        col: 3,
                        colOff: 0,
                        row: (i*18)+7,
                        rowOff: 0,
                    },
                },
            });
        }
        if(imgsrc2){
            worksheet.addImage({
                path: imgsrc2,
                type: 'picture',
                position: {
                    type: 'twoCellAnchor',
                    from: {
                        col: 2,
                        colOff: 0,
                        row: (i*18)+7,
                        rowOff: 0,
                    },
                    to: {
                        col: 3,
                        colOff: 0,
                        row: (i*18)+13,
                        rowOff: 0,
                    },
                },
            });
        }
        if(imgsrc3){
            worksheet.addImage({
                path: imgsrc3,
                type: 'picture',
                position: {
                    type: 'twoCellAnchor',
                    from: {
                        col: 2,
                        colOff: 0,
                        row: (i*18)+13,
                        rowOff: 0,
                    },
                    to: {
                        col: 3,
                        colOff: 0,
                        row: (i*18)+19,
                        rowOff: 0,
                    },
                },
            });
        }

        //worksheet.setPrintArea(1, 1, (i*6)+6, 2);
    }





// Set value of cell B1 to 300 as a number type styled with paramaters of style
    //worksheet.cell(1,2).number(200).style(style);

// Set value of cell C1 to a formula styled with paramaters of style
    //worksheet.cell(1,3).formula('A1 + B1').style(style);

// Set value of cell A2 to 'string' styled with paramaters of style
    //worksheet.cell(2,1).string('string').style(style);

// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    //worksheet.cell(3,1).bool(true).style(style).style({font: {size: 14}});

    workbook.write('Excel.xlsx');
    alert("생성완료");
}