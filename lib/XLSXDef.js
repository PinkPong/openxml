var ver = '0.0.1';
var dom = require('xmldom').DOMParser;

var _magic = 12700;
var templates = {
	content_types: {
		file_name: '[Content_Types]',
		file_ext: 'xml',
		xml: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
				'<Default Extension="xml" ContentType="application/xml"/>\n' +
				'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
				'<Default Extension="jpeg" ContentType="image/jpeg"/>\n' +
				'<Default Extension="png" ContentType="image/png"/>\n' +
				'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n' +
				'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n' +
				'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n' +
				'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n' +
				
				'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n' +
				
				'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n' +
				'</Types>',
		sheets: '<Override PartName="/xl/worksheets/sheet{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n',
		drawings: '<Override PartName="/xl/drawings/drawing{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n',
		charts: '<Override PartName="/xl/charts/chart{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n'
	},
	_rels: {
		file_name: '_rels/',
		file_ext: 'rels',
		xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
				'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n' +
				'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\n' +
				'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n' +
				'</Relationships>'
	},
	app: {
		file_name: 'docProps/app',
		file_ext: 'xml',
		xml: '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n' +
				'<Application>handmade by Pong</Application>\n' +
				'<DocSecurity>0</DocSecurity>\n' +
				'<ScaleCrop>false</ScaleCrop>\n' +
				'<Company>properex</Company>\n' +
				'<LinksUpToDate>false</LinksUpToDate>\n' +
				'<SharedDoc>false</SharedDoc>\n' +
				'<HyperlinksChanged>false</HyperlinksChanged>\n' +
				'<AppVersion>1.0</AppVersion>\n' +
				'</Properties>'
	},
	core: {
		file_name: 'docProps/core',
		file_ext: 'xml',
		xml: '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
				'<dc:creator>GTA</dc:creator>\n' +
				'<cp:lastModifiedBy>GTA</cp:lastModifiedBy>\n' +
				'<dcterms:created xsi:type="dcterms:W3CDTF">2015-01-01T00:00:00Z</dcterms:created>\n' +
				'<dcterms:modified xsi:type="dcterms:W3CDTF">2015-01-01T00:00:00Z</dcterms:modified>\n' +
				'</cp:coreProperties>'
	},
	workbook_rel: {
		file_name: 'xl/_rels/workbook',
		file_ext: 'xml.rels',
		xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
				'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" Id="rId1000"/>\n' +
				'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId1001"/>\n' +
				
				'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" Id="rId1002" />\n' +

				'</Relationships>',
		sheets: '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{id}.xml" Id="rId{id}"/>\n'
	},
	workbook: {
		file_name: 'xl/workbook',
		file_ext: 'xml',
		xml: '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n' +
				'</workbook>',
		sheets: '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="{name}" sheetId="{id}" r:id="rId{id}"/>\n',
		sheets_group: 'sheets'
	},
	
	theme: {
		file_name: 'xl/theme/theme1',
		file_ext: 'xml',
		xml: '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\n' +
				'<a:themeElements>\n' +

					'<a:clrScheme name="Office">\n' +
						'<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>\n' +
						'<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\n' +
						'<a:dk2><a:srgbClr val="44546A"/></a:dk2>\n' +
						'<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>\n' +
						'<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>\n' +
						'<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>\n' +
						'<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>\n' +
						'<a:accent4><a:srgbClr val="FFC000"/></a:accent4>\n' +
						'<a:accent5><a:srgbClr val="4472C4"/></a:accent5>\n' +
						'<a:accent6><a:srgbClr val="70AD47"/></a:accent6>\n' +
						'<a:hlink><a:srgbClr val="0563C1"/></a:hlink>\n' +
						'<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>\n' +
					'</a:clrScheme>\n' +

				'<a:fontScheme name="Office">\n' +
				
					'<a:majorFont>\n' +
						'<a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/>\n' +
						'<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/>\n' +
						'<a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/>\n' +
						'<a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/>\n' +
						'<a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/>\n' +
						'<a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/>\n' +
						'<a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/>\n' +
						'<a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>\n' +
						'<a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/>\n' +
						'<a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/>\n' +
						'<a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>\n' +
					'</a:majorFont>\n' +
					
					'<a:minorFont>\n' +
						'<a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/>\n' +
						'<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>\n' +
						'<a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/>\n' +
						'<a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/>\n' +
						'<a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/>\n' +
						'<a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/>\n' +
						'<a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/>\n' +
						'<a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/>\n' +
						'<a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/>\n' +
						'<a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/>\n' +
						'<a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/>\n' +
					'</a:minorFont>\n' +
					
				'</a:fontScheme>\n' +

				'<a:fmtScheme name="Office">\n' +
					'<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\n' +
						'<a:gradFill rotWithShape="1">\n' +
							'<a:gsLst>\n' +
								'<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>\n' +
								'<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>\n' +
								'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>\n' +
							'</a:gsLst>\n' +
							'<a:lin ang="5400000" scaled="0"/>\n' +
						'</a:gradFill>\n' +
						'<a:gradFill rotWithShape="1">\n' +
							'<a:gsLst>\n' +
								'<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>\n' +
								'<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>\n' +
								'<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>\n' +
							'</a:gsLst>\n' +
							'<a:lin ang="5400000" scaled="0"/>\n' +
						'</a:gradFill>\n' +
					'</a:fillStyleLst>\n' +

					'<a:lnStyleLst>\n' +
						'<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n' +
						'<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n' +
						'<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\n' +
					'</a:lnStyleLst>\n' +

					'<a:effectStyleLst>\n' +
						'<a:effectStyle><a:effectLst/></a:effectStyle>\n' +
						'<a:effectStyle><a:effectLst/></a:effectStyle>\n' +
						'<a:effectStyle>\n' +
							'<a:effectLst>\n' +
								'<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>\n' +
							'</a:effectLst>\n' +
						'</a:effectStyle>\n' +
					'</a:effectStyleLst>\n' +

					'<a:bgFillStyleLst>\n' +
						'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\n' +
						'<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>\n' +
						'<a:gradFill rotWithShape="1">\n' +
						'<a:gsLst>\n' +
							'<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>\n' +
							'<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>\n' +
							'<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>\n' +
						'</a:gsLst>\n' +
						'<a:lin ang="5400000" scaled="0"/>\n' +
						'</a:gradFill>\n' +
					'</a:bgFillStyleLst>\n' +
				'</a:fmtScheme>\n' +
				
				'</a:themeElements>\n' +
				'<a:objectDefaults/>\n' +
				'<a:extraClrSchemeLst/>\n' +

				'<a:extLst>\n' +
					'<a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">\n' +
						'<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>\n' +
					'</a:ext>\n' +
				'</a:extLst>\n' +

				'</a:theme>' 
				
	},
	
	styles: {
		file_name: 'xl/styles',
		file_ext: 'xml',
		xml: '<styleSheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n' +
				'</styleSheet>',
		fonts: '<font>\n' +
				'<b val="{b}"/>\n' +
				'<u val="{u}"/>\n' +
				'<i val="{i}"/>\n' +
				'<strike val="{strike}"/>\n' +
				'<sz val="{sz}"/>\n' +
				'<color rgb="{color}"/>\n' +
				'<name val="{name}"/>\n' +
				'<family val="{family}"/>\n' +
				'</font>\n',
		fonts_group: 'fonts',
		fills: '<fill><patternFill patternType="solid"><fgColor rgb="{fg_clr}"/><bgColor rgb="{bg_clr}"/></patternFill></fill>\n',
		fills_group: 'fills',
		borders: '<border>\n' +
				'<left style="{l_style}"><color rgb="{l_clr}"/></left>\n' + // thin,thick,medium,none
				'<right style="{r_style}"><color rgb="{r_clr}"/></right>\n' +
				'<top style="{t_style}"><color rgb="{t_clr}"/></top>\n' +
				'<bottom style="{b_style}"><color rgb="{b_clr}"/></bottom>\n' +
				'<diagonal style="{d_style}"><color rgb="{d_clr}"/></diagonal>\n' +
				'</border>\n',
		borders_group: 'borders',
		cellXfs: '<xf borderId="{border_id}" applyBorder="{use_border}" fillId="{fill_id}" applyFill="{use_fill}" fontId="{font_id}" applyFont="{use_font}" applyAlignment="{use_txt_aling}">\n' +
				'<alignment wrapText="{txt_wrap}" textRotation="{txt_rot}" vertical="{txt_v_align}" horizontal="{txt_h_align}"/>\n' +
				'</xf>\n',
		cellXfs_group: 'cellXfs'
	},
	shared_strings: {
		file_name: 'xl/sharedStrings',
		file_ext: 'xml',
		xml: '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n' +
				'</sst>',
		//strings: '<si><t>{str}</t></si>\n'
		strings: '{str}\n'
	},
	sheet: {
		file_name: 'xl/worksheets/sheet',
		file_ext: 'xml',
		xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\n' +
				//'<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>\n' +
				'</worksheet>',
		/*sheets_group:'sheetData',
		 merge_cell_group:'mergeCells',
		 marge_cell:'<mergeCell ref={ref}/>\n',*/
		drawings: '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{id}"/>\n'
	},
	sheet_rels: {
		file_name: 'xl/worksheets/_rels/sheet',
		file_ext: 'xml.rels',
		xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
				'</Relationships>',
		drawings: '<Relationship Id="rId{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{id}.xml"/>\n'
	},
	drawing_rels: {
		file_name: 'xl/drawings/_rels/drawing',
		file_ext: 'xml.rels',
		xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
				'</Relationships>',
		charts: 'media_selector',
		chart: '<Relationship Id="rId{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{id}.xml"/>\n',
		image: '<Relationship Id="rId{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_id}"/>\n'
		
	},
	drawing: {
		file_name: 'xl/drawings/drawing',
		file_ext: 'xml',
		xml: '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n' +
				'</xdr:wsDr>',
		drawings: 'media_selector',
		chart:	'<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n' +
							'<xdr:from>\n' +
								'<xdr:col>{col_from}</xdr:col>\n' +
								'<xdr:colOff>0</xdr:colOff>\n' +
								'<xdr:row>{row_from}</xdr:row>\n' +
								'<xdr:rowOff>0</xdr:rowOff>\n' +
							'</xdr:from>\n' +
							
							'<xdr:to>\n' +
								'<xdr:col>{col_to}</xdr:col>\n' +
								'<xdr:colOff>0</xdr:colOff>\n' +
								'<xdr:row>{row_to}</xdr:row>\n' +
								'<xdr:rowOff>0</xdr:rowOff>\n' +
							'</xdr:to>\n' +
							
							'<xdr:graphicFrame macro="">\n' +
								'<xdr:nvGraphicFramePr>\n' +
								'<xdr:cNvPr id="{chart_id}" name="chart {chart_id}"/>\n' +
								'<xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr>\n' +
								'</xdr:nvGraphicFramePr>\n' +
								'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>\n' +
								'<a:graphic>\n' +
								'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">\n' +
								'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{chart_id}"/>\n' +
								'</a:graphicData>\n' +
								'</a:graphic>\n' +
							'</xdr:graphicFrame>\n' +
							
							'<xdr:clientData/>\n' +
						'</xdr:twoCellAnchor>\n',
				
				
		image:	'<!--xdr:twoCellAnchor editAs="oneCell">\n' +
							'<xdr:from>\n' +
								'<xdr:col>{col_from}</xdr:col>\n' +
								'<xdr:colOff>0</xdr:colOff>\n' +
								'<xdr:row>{row_from}</xdr:row>\n' +
								'<xdr:rowOff>0</xdr:rowOff>\n' +
							'</xdr:from>\n' +

							'<xdr:to>\n' +
								'<xdr:col>{col_to}</xdr:col>\n' +
								'<xdr:colOff>0</xdr:colOff>\n' +
								'<xdr:row>{row_to}</xdr:row>\n' +
								'<xdr:rowOff>0</xdr:rowOff>\n' +
							'</xdr:to-->\n' +
							
						'<xdr:oneCellAnchor>\n' +
							'<xdr:from>\n' +
								'<xdr:col>{col_from}</xdr:col>\n' +
								'<xdr:colOff>0</xdr:colOff>\n' +
								'<xdr:row>{row_from}</xdr:row>\n' +
								'<xdr:rowOff>0</xdr:rowOff>\n' +
							'</xdr:from>\n' +
							'<xdr:ext cx="{image_wd}" cy="{image_hg}"/>\n' +

							'<xdr:pic>\n' +
								'<xdr:nvPicPr>\n' +
									'<xdr:cNvPr id="{image_id}" name="picture {image_id}"/>\n' +
										'<xdr:cNvPicPr>\n' +
										'<a:picLocks noChangeAspect="1"/>\n' +
									'</xdr:cNvPicPr>\n' +
								'</xdr:nvPicPr>\n' +
								'<xdr:blipFill>\n' +
									'<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId{image_id}"/>\n' +
									'<a:stretch>\n' +
									'<a:fillRect/>\n' +
									'</a:stretch>\n' +
								'</xdr:blipFill>\n' +
								'<xdr:spPr>\n' +
									'<a:xfrm>\n' +
										'<!--a:off x="0" y="0"/><a:ext cx="0" cy="0"/-->\n' +
									'</a:xfrm>\n' +
									'<a:prstGeom prst="rect">\n' +
										'<a:avLst/>\n' +
									'</a:prstGeom>\n' +
								'</xdr:spPr>\n' +
							'</xdr:pic>\n' +

							'<xdr:clientData/>\n' +
						'<!--/xdr:twoCellAnchor-->' +
						'</xdr:oneCellAnchor>' 
	}
};


module.exports = function () {
	return {
	
		echo: function () {
			var rez = {'ver': ver, 'module': 'XSLXDef'};
			return rez;
		},
		
		get_xml: function (_type, sheets, drawings, charts)
		{
			var media_type = null;
			//if()
			var o = templates[_type];
			var _xml = new dom().parseFromString(o.xml);
			
			var _vars = [sheets, drawings, charts];
			var _vars_name = ['sheets', 'drawings', 'charts'];

			for (var dd = 0; dd < _vars.length; dd++) {
				if (_vars[dd] != null) {
					
					var rez_str = '';
					var _item = _vars[dd];
					var _item_name = _vars_name[dd];
					/*
					var _item_names = _vars_name[dd].split(',');
					var _item_name = _item_names[0];
					
					while((o[_item_name] === undefined)&&(_item_names.length>0)) {
						_item_name = _item_names.shift();
					}
					*/
					//console.log(_type,_item_name);
					if ((typeof (_item) === 'string') || (_item instanceof String)) {
						rez_str = _item;
					} else {
						if (o[_item_name] !== undefined) {
							var ln = _item.length;
							
							for (var aa = 0; aa < ln; aa++) {
								var data_str = (o[_item_name]);//+'';//.toString();
								//console.log(aa,_item[aa]);
								if(data_str === 'media_selector') {
									data_str = o[_item[aa].media_selector];
								}
								var _obj = _item[aa];
								for (var bb in _obj) {
									if (((typeof (_obj[bb]) === 'string') ||
											(_obj[bb] instanceof String) ||
											(typeof (_obj[bb]) === 'number')) &&
											(data_str.indexOf('{' + bb + '}') !== -1)) {
										var ptrn = new RegExp('{' + bb + '}', 'g');
										if (bb === 'id') {
											data_str = data_str.replace(ptrn, (_obj[bb] + 1));
										} else {
											data_str = data_str.replace(ptrn, _obj[bb]);
										}
									}
								}
								rez_str = rez_str.concat(data_str);
							}
						}
					}
					if (rez_str !== '') {
						var data_xml = new dom().parseFromString('\n'+rez_str+'\n');
						if (o[_item_name + '_group'] === undefined) {
							_xml.documentElement.appendChild(data_xml);
						} else {
							var tgt = new dom().parseFromString('\n<' + o[_item_name + '_group'] + '/>\n');
							tgt.documentElement.appendChild(data_xml);
							_xml.documentElement.appendChild(tgt);
						}
					}
				}
			}
			return [o.file_name, o.file_ext, _xml];
		},
		
		get_extra_xml: function (_type, fonts, fills, borders, cellXfs, strings)
		{
			var o = templates[_type];
			var _xml = new dom().parseFromString(o.xml);
			var data = [fonts, fills, borders, cellXfs, strings];
			var data_refs = ['fonts', 'fills', 'borders', 'cellXfs', 'strings', 'theme'];
			for (var dd = 0; dd < data.length; dd++) {
				if (data[dd] != null) {
					var rez_str = '';
					var _data = data[dd];
					var d_ref = data_refs[dd];
					if ((typeof (_data) === 'string') || (_data instanceof String)) {
						rez_str = _data;
					} else {
						if (o[d_ref] !== undefined) {
							var ln = _data.length;
							for (var aa = 0; aa < ln; aa++) {
								var data_str = (o[d_ref])+'';//.toString();
								var _obj = _data[aa];
								for (var bb in _obj) {
									if (((typeof (_obj[bb]) === 'string') ||
											(_obj[bb] instanceof String) ||
											(typeof (_obj[bb]) === 'number')) &&
											(data_str.indexOf('{' + bb + '}') !== -1)) {
										var ptrn = new RegExp('{' + bb + '}', 'g');
										if (bb === 'id') {
											data_str = data_str.replace(ptrn, (_obj[bb] + 1));
										} else {
											data_str = data_str.replace(ptrn, _obj[bb]);
										}
									}
								}
								rez_str = rez_str.concat(data_str);
							}
						}
					}
					if (rez_str !== '') {
						var data_xml = new dom().parseFromString(rez_str);
						if (o[d_ref + '_group'] === undefined) {
							_xml.documentElement.appendChild(data_xml);
						} else {
							var tgt = new dom().parseFromString('\n<' + o[d_ref + '_group'] + '/>\n');
							tgt.documentElement.appendChild(data_xml);
							_xml.documentElement.appendChild(tgt);
						}
					}
				}
			}
			return [o.file_name, o.file_ext, _xml];
		}
		
	}

};