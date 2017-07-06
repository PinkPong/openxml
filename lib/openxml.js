var ver = '0.0.1';
var time = 0;

var _timer = function() {
	var t = new Date();
	var diff = (t.valueOf() - time)/1000 + ' sec';
	time = t.valueOf();
	
	return diff;
}	

var xDoc = new require('./XLSXExporter.js')();
var xChart = new require('./OpenXMLChart.js')();

var exports = module.exports = {};

exports.echo = function () {

	var rez = {'ver':ver,'module':'openxml'};
	return rez;
};

exports.init = function (tempFolder) {
	console.log('Xexp echo->',xDoc.echo(),_timer());
	xDoc.init(tempFolder);
};

exports.createStyle = function (style, font, fill, border) {
	return xDoc.create_style(style, font, fill, border);
};

exports.addColumn = function (sheet_id, col) {
	xDoc.add_col(sheet_id, col);
};

exports.split = function (sheet_id, x, y) {
	xDoc.split_view(sheet_id, x, y);
};

exports.mergeCells = function (sheet_id, mc) {
	xDoc.merge_cell(sheet_id, mc);
};

exports.headerFooter = function (sheet_id, data) {

	//console.log(sheet_id,data)
	for(var d in data) {
		xDoc['add_' + d].call(xDoc, sheet_id, data[d]);
	}
	
	/*switch (data.type.toLowerCase()) {
		case 'header':
			xDoc.add_header(sheet_id, data);
		break;
		case 'footer':
			xDoc.add_header(sheet_id, data);
		break;
	}*/
};

exports.createSheet = function (name, data) {
	console.log('create sheet->',name,_timer());
	var sh=xDoc.add_sheet(name);
	
	var dLn = data.length;
	console.log('add '+dLn+' data rows to sheet->',name);
	switch (name.toLowerCase()) {
		case 'info':
		case 'summary':
			for(var a=0;a<dLn;a++)
			{
				xDoc.add_row(sh,data[a]);
				data[a]=null;
				//console.log(a+' added');
			}
		break;
		default:
			for(var a=0;a<dLn;a++)
			{
				xDoc.add_row_no_check(sh,data[a],a+1,false,data[a].useOpts);
				data[a]=null;
				//console.log(a+' added. no check');
			}
		break;
	}
};


exports.addImage = function (sheet_id, data) {
	xDoc.add_image(sheet_id, data);//.imagePath, data.bbox);
};

exports.addChart = function (name, data) {
	console.log('create chart->',name);
	
	var chartColors=["CE4D45","0A3E83","2AA850","843E83","0E9C1D","3039D0","E10D1D","CE1645","0E9CE9","BD783A"];
	
	var singlePage=true;
	var chartSheetName=name;
	
	var chartSH=xDoc.get_sheet(chartSheetName);
	//console.log('check charts',chartSH);
	if(chartSH===undefined) {
		chartSH=xDoc.add_sheet(chartSheetName);
	}
	//console.log('charts',chartSH);
	xChart.OpenXMLChart(data.chartName,data.chartType);
	
	var dataSheetName = data.dataSheet;
	var dataSH=xDoc.get_sheet(dataSheetName);
	//console.log('check data',dataSH,dataSheetName);
	
	if(dataSH===undefined) {
		dataSH=xDoc.add_sheet(dataSheetName);
	}
	//console.log('data',dataSH);
/*
var chartCFG = {
			chartName:'main KPIs',
			chartType:'line',
			categoryColumn:0,
			nameRow:4,
			dataBox:[1,3,30,20+3],
			data:chartData,
			dataSheet:'mega KPIs',
			chartBox:[1,4,15,20]
	};
	*/
	
	var firstRow = 0;
	var chartHG=data.chartBox[3] - data.chartBox[1];
	
	if(dataSH!=chartSH) {
		singlePage=false;
		chartHG=0;
	} else {
		for(var a=0;a<data.dataBox[1];a++) {
			xDoc.add_row(dataSH,data.data[a]);
			firstRow ++;
		}
		for(var a=0;a<chartHG+2;a++) {
			xDoc.add_row(dataSH,['-']);
		}
		
	}
	
	//console.log(data);
	
	for(var a=firstRow;a<data.data.length;a++) {
		xDoc.add_row(dataSH,data.data[a]);
	}
	
	var serieID = 0;
	
	var category_col=xDoc.abc_id(data.categoryColumn);
	var name_row=data.nameRow+chartHG+firstRow;

//	var data_start_row=data.dataBox[1]+1+chartHG+firstRow;
//	var data_end_row=data.dataBox[3]+1+chartHG+firstRow;
	var data_start_row=data.dataBox[1]+0+chartHG+firstRow;
	var data_end_row=data.dataBox[3]+0+chartHG+firstRow;
	
	var series_num=data.dataBox[2];
				   
	for(var ss=0;ss<series_num;ss++) {
		var clr = ss;
		if (ss >= chartColors.length) {
			clr = ss%chartColors.length;
		}
		console.log(chartColors[Math.round(ss/chartColors.length)],chartColors[clr])
		var serie_col=xDoc.abc_id(ss+1);
		xChart.addSerie(0,
						//chartColors[Math.round(ss/chartColors.length)],
						chartColors[clr],
						"'"+dataSheetName+"'!$"+serie_col+"$"+name_row,
						"'"+dataSheetName+"'!$"+category_col+"$"+data_start_row+":$"+category_col+"$"+data_end_row,
						"'"+dataSheetName+"'!$"+serie_col+"$"+data_start_row+":$"+serie_col+"$"+data_end_row,
						serieID, null, data.chartType);
		serieID++;
	};
	xDoc.add_chart(chartSH,xChart.getChart(),data.chartBox);
};


exports.getDoc = function (cb, fName) {
	console.log('pack sheets');
	if (cb !== undefined) {
		console.log('pack with cb',_timer());
		xDoc.pack(cb, fName);
	} else {
		console.log('sync pack',_timer());
		var fNm=xDoc.pack();
	
		return fNm;
	}
};

var init = function () {
	// nothing to do yet
};
