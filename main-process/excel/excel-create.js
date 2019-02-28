// example : https://www.npmjs.com/package/excel4node
function initExcel(box){
    // Require library
    var excel = require('excel4node');
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('Sheet 1');
    //var worksheet2 = workbook.addWorksheet('Sheet 2');
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
    var bottom_top_style = workbook.createStyle({
        border: {
            top: {
                style: 'double'
            },
            bottom: {
                style: 'double'
            }
        }
    });
    var bottom_top_style2 = workbook.createStyle({
        border: {
            top: {
                style: 'double'
            },
            right: {
                style: 'double'
            },
            bottom: {
                style: 'double'
            }
        }
    });
    var bottom_double_style = workbook.createStyle({
        border: {
            bottom: {
                style: 'double'
            }
        }
    });
    var top_double_style = workbook.createStyle({
        border: {
            top: {
                style: 'double'
            }
        }
    });
    // Set value of cell A1 to 100 as a number type styled with paramaters of style
    //worksheet.cell(1,1).number(100).style(style);

    for(var i=0; i<box.length; i++){
        //console.log(box[i]);
        //console.log($(box[i]).find(".sigong").val());
        //입력된 값
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

        //세로 칸크기
        var cnt18 = 1;
        for(var j=0; j < 3; j++){
            worksheet.row(i*24+(cnt18++)).setHeight(2);
            worksheet.row(i*24+(cnt18++)).setHeight(30);
            worksheet.row(i*24+(cnt18++)).setHeight(49);
            worksheet.row(i*24+(cnt18++)).setHeight(30);
            worksheet.row(i*24+(cnt18++)).setHeight(49);
            worksheet.row(i*24+(cnt18++)).setHeight(30);
            worksheet.row(i*24+(cnt18++)).setHeight(49);
            //if(j == 2)
            worksheet.row(i*24+(cnt18++)).setHeight(5.5);
        }

        //가로 칸 크기
        worksheet.column(1).setWidth(17);
        worksheet.column(2).setWidth(0.2);
        worksheet.column(3).setWidth(58);
        worksheet.column(4).setWidth(0.2);
        cnt18 = 1;
        //공종,시공내용,위치영역
        //worksheet.cell(i*20+(cnt18++),1).string('').style(top_style);
        worksheet.cell(i*24+(cnt18),1, i*24+(cnt18+1),1 ,true).string('공종').style(top_style);
        cnt18++;cnt18++;
        worksheet.cell(i*24+(cnt18++),1).string(gongjeong1).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*24+(cnt18++),1).string(sigong1).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('위치').style(style);
        worksheet.cell(i*24+(cnt18),1, i*24+(cnt18+1),1, true).string(postion1).style(inputStyle);
        cnt18++;cnt18++;
        //worksheet.cell(i*20+(cnt18++),1).string('').style(bottom_inputStyle);


        //worksheet.cell(i*20+(cnt18++),1).string('').style(top_style);
        worksheet.cell(i*24+(cnt18),1, i*24+(cnt18+1),1 ,true).string('공종').style(top_style);
        cnt18++;cnt18++;
        worksheet.cell(i*24+(cnt18++),1).string(gongjeong2).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*24+(cnt18++),1).string(sigong2).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('위치').style(style);
        //worksheet.cell(i*20+(cnt18++),1).string(postion2).style(inputStyle);
        worksheet.cell(i*24+(cnt18),1, i*24+(cnt18+1),1, true).string(postion2).style(inputStyle);
        cnt18++;cnt18++;
        //worksheet.cell(i*20+(cnt18++),1).string('').style(bottom_inputStyle);

        //worksheet.cell(i*20+(cnt18++),1).string('').style(top_style);
        worksheet.cell(i*24+(cnt18),1, i*24+(cnt18+1),1 ,true).string('공종').style(top_style);
        cnt18++;cnt18++;
        worksheet.cell(i*24+(cnt18++),1).string(gongjeong3).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('시공내용').style(style);
        worksheet.cell(i*24+(cnt18++),1).string(sigong3).style(inputStyle);
        worksheet.cell(i*24+(cnt18++),1).string('위치').style(style);
        worksheet.cell(i*24+(cnt18),1,i*24+(cnt18+1),1,true).string(postion3).style(bottom_inputStyle);
        cnt18++;
        //worksheet.cell(i*20+(cnt18++),1).string('').style(bottom_inputStyle);

        //그림 넣고 빈공간을 위한 영역
        worksheet.cell((i*24)+1, 2, (i*24)+24, 2, true).style(bottom_top_style); //공종, 그림 사이
        worksheet.cell((i*24)+1, 4, (i*24)+24, 4, true).style(bottom_top_style2); //그림과 문서 오른쪽끝부분 사이
        worksheet.cell((i*24)+1, 3).style(top_double_style); //그림에서 문서 윗부분
        worksheet.cell((i*24)+24, 3).style(bottom_double_style); //그림에서 문서 아랫부분
        worksheet.cell((i*24)+9, 3).style(top_double_style); //첫번째 그림, 두번째 그림 사이
        worksheet.cell((i*24)+17, 3).style(top_double_style); //두번째 그림, 세번째 그림 사이

        //그림영역
        worksheet.cell((i*24)+2, 3, (i*24)+7, 3, true);
        worksheet.cell((i*24)+10, 3, (i*24)+15, 3, true);
        worksheet.cell((i*24)+18, 3, (i*24)+24, 3, true);
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
                        col: 3,
                        colOff: 0,
                        row: (i*24)+2,
                        rowOff: 0,
                    },
                    to: {
                        col: 4,
                        colOff: 0,
                        row: (i*24)+8,
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
                        col: 3,
                        colOff: 0,
                        row: (i*24)+10,
                        rowOff: 0,
                    },
                    to: {
                        col: 4,
                        colOff: 0,
                        row: (i*24)+16,
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
                        col: 3,
                        colOff: 0,
                        row: (i*24)+18,
                        rowOff: 0,
                    },
                    to: {
                        col: 4,
                        colOff: 0,
                        row: (i*24)+24,
                        rowOff: 0,
                    },
                },
            });
        }


    }
    worksheet.setPrintArea(1, 1, ((box.length-1)*24)+24, 4);




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