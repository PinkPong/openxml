var ver = '0.0.1';
var dom = require('xmldom').DOMParser;

var chartTemplate = [
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'+
			'<c:lang val="en-US"/>\n'+
			
			'<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:noFill/></a:ln></c:spPr>\n'+
			
			'<c:chart>\n'+
				'<c:plotVisOnly val="0"/>\n'+
				'<c:dispBlanksAs val="gap"/>\n'+
				'<c:legend><c:legendPos val="b"/></c:legend>\n'+
			
				'<c:title>\n'+
					'<c:tx>\n'+
						'<c:rich>\n'+
							'<a:bodyPr/>\n'+
							'<a:p>\n'+
								'<a:pPr>\n'+
									'<a:defRPr sz="975" b="1" i="0" u="none" strike="noStrike" baseline="0">\n'+
										'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>\n'+
										'<a:latin typeface="Arial"/>\n'+
										'<a:ea typeface="Arial"/>\n'+
										'<a:cs typeface="Arial"/>\n'+
									'</a:defRPr>\n'+
								'</a:pPr>\n'+
								'<a:r><a:t>Chart</a:t></a:r>\n'+
							'</a:p>\n'+
						'</c:rich>\n'+
					'</c:tx>\n'+
					'<c:spPr><a:noFill/></c:spPr>\n'+
				'</c:title>\n'+
				
				'<c:plotArea>\n'+
					'<c:spPr><a:noFill/><a:ln w="25400"><a:noFill/></a:ln></c:spPr>\n',

					
					
													
					'<c:catAx>\n'+
						'<c:axPos val="b"/>\n'+
						'<c:axId val="0"/>\n'+
						'<c:delete val="0"/>\n'+
						'<c:crossAx val="1"/>\n'+
						'<c:crosses val="autoZero"/>\n'+
						'<c:numFmt formatCode="General" sourceLinked="1"/>\n'+
						'<c:scaling><c:orientation val="minMax"/></c:scaling>\n'+
						'<c:auto val="1"/>\n'+
						'<c:lblAlgn val="ctr"/>\n'+
						'<c:lblOffset val="100"/>\n'+
						'<!--c:tickLblSkip val="1"/>\n'+
						'<c:tickMarkSkip val="1"/-->\n'+
						'<c:tickLblPos val="nextTo"/>\n'+
						'<c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="880000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr>\n'+			
						'<c:majorGridlines>\n'+
							'<c:spPr>\n'+
								'<a:ln w="3175"><a:solidFill><a:srgbClr val="CCCCFF"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>\n'+
							'</c:spPr>\n'+
						'</c:majorGridlines>\n'+
						'<c:txPr>\n'+
							'<a:bodyPr rot="-5400000" vert="horz"/>\n'+
							'<a:p>\n'+
								'<a:pPr>\n'+
									'<a:defRPr sz="500" b="0" i="0" u="none" strike="noStrike" baseline="0">\n'+
										'<a:solidFill><a:srgbClr val="880000"/></a:solidFill>\n'+
										'<a:latin typeface="Arial"/>\n'+
										'<a:ea typeface="Arial"/>\n'+
										'<a:cs typeface="Arial"/>\n'+
									'</a:defRPr>\n'+
								'</a:pPr>\n'+
								'<a:endParaRPr lang="en-US"/>\n'+
							'</a:p>\n'+
						'</c:txPr>\n'+
					'</c:catAx>\n'+

					'<c:valAx>\n'+
						'<c:axPos val="l"/>\n'+
						'<c:axId val="1"/>\n'+
						'<c:delete val="0"/>\n'+
						'<c:crossAx val="0"/>\n'+
						'<c:crosses val="autoZero"/>\n'+
						'<c:numFmt formatCode="General" sourceLinked="1"/>\n'+
						'<c:scaling><c:orientation val="minMax"/></c:scaling>\n'+
						'<c:crossBetween val="between"/>\n'+
						'<c:tickLblPos val="nextTo"/>\n'+
						'<c:spPr>\n'+
							'<a:ln w="3175"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>\n'+
						'</c:spPr>\n'+
						'<c:majorGridlines>\n'+
							'<c:spPr>\n'+
								'<a:ln w="3175"><a:solidFill><a:srgbClr val="CCCCFF"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>\n'+
							'</c:spPr>\n'+
						'</c:majorGridlines>\n'+
						'<c:txPr>\n'+
							'<a:bodyPr rot="0" vert="horz"/>\n'+
							'<a:p>\n'+
								'<a:pPr>\n'+
									'<a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0">\n'+
										'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>\n'+
										'<a:latin typeface="Arial"/>\n'+
										'<a:ea typeface="Arial"/>\n'+
										'<a:cs typeface="Arial"/>\n'+
									'</a:defRPr>\n'+
								'</a:pPr>\n'+
								'<a:endParaRPr lang="en-US"/>\n'+
							'</a:p>\n'+
						'</c:txPr>\n'+
						'<c:title>\n'+
							'<c:tx>\n'+
								'<c:rich>\n'+
									'<a:bodyPr/>\n'+
									'<a:p>\n'+
										'<a:pPr>\n'+
											'<a:defRPr sz="975" b="1" i="0" u="none" strike="noStrike" baseline="0">\n'+
												'<a:solidFill>\n'+'<a:srgbClr val="000000"/>\n'+'</a:solidFill>\n'+
												'<a:latin typeface="Arial"/>\n'+
												'<a:ea typeface="Arial"/>\n'+
												'<a:cs typeface="Arial"/>\n'+
											'</a:defRPr>\n'+
										'</a:pPr>\n'+
										'<a:r><a:t></a:t></a:r>\n'+
									'</a:p>\n'+
								'</c:rich>\n'+
							'</c:tx>\n'+
						'</c:title>\n'+
					'</c:valAx>\n'+
	
				'</c:plotArea>\n'+
			'</c:chart>\n'+
		'</c:chartSpace>\n'];

var serieTemplate = '<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'+
										'<c:idx val="0"/>\n'+
										'<c:order val="0"/>\n'+
										'<c:smooth val="1"/>\n'+
										'<c:tx><c:strRef><c:f>serie</c:f></c:strRef></c:tx>\n'+	

										'<c:spPr>\n'+
											//'<a:ln w="3175"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:prstDash val="solid"/></a:ln>\n'+
											'<a:ln w="25400"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:prstDash val="solid"/></a:ln>\n'+
										'</c:spPr>\n'+

										'<c:marker>\n'+
											'<c:symbol val="circle"/>\n'+
											'<c:size val="6"/>\n'+
											'<c:spPr>\n'+
												'<a:solidFill><a:srgbClr val="c0c0c0"/></a:solidFill>\n'+
												'<a:ln><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:prstDash val="solid"/></a:ln>\n'+
											'</c:spPr>\n'+
										'</c:marker>\n'+
		
										'<c:cat><c:strRef><c:f>category</c:f></c:strRef></c:cat>\n'+
										'<c:val><c:numRef><c:f>data</c:f></c:numRef></c:val>\n'+
									'</c:ser>\n';
		
		
var chartTypes = {
										bar:'<c:barChart>\n'+
												'<c:barDir val="col"/>\n'+
												'<c:varyColors val="0"/>\n'+
												'<c:grouping val="standard"/>\n'+
												'<c:marker val="1"/>\n'+
												'<c:axId val="0"/>\n'+
												'<c:axId val="1"/>\n'+
											'</c:barChart>\n',
										line:'<c:lineChart>\n'+
												'<c:grouping val="standard"/>\n'+
												'<c:varyColors val="0"/>\n'+
												'<c:marker val="1"/>\n'+
												'<c:axId val="0"/>\n'+
												'<c:axId val="1"/>\n'+
											'</c:lineChart>\n',
										area:'<c:areaChart>\n'+
												'<c:grouping val="standard"/>\n'+
												'<c:varyColors val="0"/>\n'+
												'<c:marker val="1"/>\n'+
												'<c:axId val="0"/>\n'+
												'<c:axId val="1"/>\n'+
											'</c:areaChart>\n'
										};
											

		
									
var chartData;
var chartType;

var series;
var name;

var __c;
var __a;
		
module.exports = function () {
	return {
	
		echo: function () {
			var rez = {'ver': ver, 'module': 'OpenXMLChart'};
			return rez;
		},
		
		OpenXMLChart: function(n, type) {
			if(type===undefined) {
				type = 'line';
			}
			
			name=n;
			chartType=type;
			
			chartData = new dom().parseFromString(chartTemplate[0]+chartTypes[chartType]+chartTemplate[1]);
			
			__c='http://schemas.openxmlformats.org/drawingml/2006/chart';
			__a='http://schemas.openxmlformats.org/drawingml/2006/main';
			
			this.changeName();
			series=[];
		},
		
		changeName: function() {
			chartData.getElementsByTagNameNS(__a,'t').item(0).firstChild.replaceData(0,100,name);
		},
		
	
		addSerie: function(id,clr,name,category,data,order,markerType) {
	
			var serie_id = 0;
			var s=new dom().parseFromString(serieTemplate);

			if((order===null)||(order===undefined)) {
				order=id;
			}

			s.getElementsByTagNameNS(__c,'idx').item(0).setAttribute('val',id);
			s.getElementsByTagNameNS(__c,'order').item(0).setAttribute('val',order);
			s.getElementsByTagNameNS(__c,'f').item(0).firstChild.replaceData(0,100,name);

			if ((typeof (clr) === 'string') || (clr instanceof String)) {
				s.getElementsByTagNameNS(__a,'srgbClr').item(0).setAttribute('val',clr);
				s.getElementsByTagNameNS(__a,'srgbClr').item(1).setAttribute('val',clr);
			}else {
				s.getElementsByTagNameNS(__a,'srgbClr').item(0).setAttribute('val',clr[0]);
				s.getElementsByTagNameNS(__a,'srgbClr').item(1).setAttribute('val',clr[1]);
			}
			
			if((markerType!==null)&&(markerType!==undefined)) {
				s.getElementsByTagNameNS(__c,'symbol').item(0).setAttribute('val',markerType);
			}
			
			s.getElementsByTagNameNS(__c,'f').item(1).firstChild.replaceData(0,100,category);
			s.getElementsByTagNameNS(__c,'f').item(2).firstChild.replaceData(0,100,data);
			
			series.push(s);
			serieID=series.length-1;
			
			return serieID;
		},
		
		getChart: function() {
			for(var aa=0;aa<series.length;aa++) {
				chartData.getElementsByTagNameNS(__c,chartType+'Chart').item(0).appendChild(series[aa]);
			}
			
			return chartData;
		}
		
	}
};