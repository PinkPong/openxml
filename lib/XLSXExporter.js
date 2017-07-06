var ver = '0.0.1';

var time = 0;

var _timer = function() {
	var t = new Date();
	var diff = (t.valueOf() - time)/1000 + ' sec';
	time = t.valueOf();
	
	return diff;
}	

var tmpFolder = 'tmp/';
var easy_view=true;//false;
var _zip=null;
		
var sheets=null;
var drawings=null;
var charts=null;
var images=null;
		
var s_fonts=null;
var s_fills=null;
var s_borders=null;
var styles=null;
		
var sh_str=null;
var totalShStr=0;
var _sh_str=null;
var _sh_str_big=null;

		
var abc='abcdefghijklmnopqrstuvwxyz';
var emu = 914400;

var x = require('./XLSXDef.js')();
var srl = require('xmldom').XMLSerializer;

var _orders = null;

module.exports = function() {
  return {
  
    echo: function() {
		var rez = {'ver':ver,'module':'XSLXExporter'};
		return rez;
    },

	init:function(tmpFolderPath)
		{
			if (tmpFolderPath != undefined) {
				tmpFolder = tmpFolderPath;
			}				
			_zip=new require('node-zip')();
			console.log('----------------XLSXExporter init',_timer());
			//for(var i=0;i<20000;i++) {console.log(i+'->'+this.abc_id(i))}
			//console.log('----------------XLSXExporter init');
			sheets=[];
			charts=[];
			images=[];
			drawings=[];
			
			// init some default styles
			
			s_fonts=[{b:0,i:0,u:'none',strike:0,family:2,sz:8,color:'0',name:'Arial'}];
			s_fills=[{fg_clr:'0',bg_clr:'0'},{fg_clr:'0',bg_clr:'0'},{fg_clr:'0',bg_clr:'0'}];
			s_borders=[{l_style:'none',r_style:'none',t_style:'none',b_style:'none',d_style:'none',l_clr:'0',r_clr:'0',t_clr:'0',b_clr:'0',d_clr:'0'}];
			styles=[{border_id:0,use_border:0,fill_id:0,use_fill:0,font_id:0,use_font:0,use_txt_aling:0,txt_rot:0,txt_v_align:'center',txt_h_align:'center'}];
			
			sh_str=[];
			_sh_str={};
			_sh_str_big={};
			_orders={};
			totalShStr=0;
		},
		
	//zip_io_err:function(event:IOErrorEvent):void
	/*zip_io_err:function(event)
		{
			trace("IO Error:",event.text);
		},*/
		
	//public function abc_id(num:int):String
	abc_id:function(num){//26
		var col_id=_orders[num];
		var n=num;
		if(col_id===undefined) {
			col_id='';
			var ln=abc.length;//26
			var val=num%ln;//0
			//console.log(num,ln,val)
			col_id=abc.charAt(val)+col_id;//'a'
			
			while(num>=ln)//26
			//while(num>0)//26
			{
				num-=val;//26
				val=num%ln;//0
				num/=ln;
				if(num>0) {
					col_id=abc.charAt(num-1)+col_id;
				}
			}
			_orders[n]=col_id;
		}
		
		return col_id;
	},
		
	//public function create_style(style,font:*=null,fill:*=null,border:*=null):uint
	create_style:function(style, font, fill, border) {
			var _id=0;
			
			var font_id=-1;
			var fill_id=-1;
			var border_id=-1;
		
			var prop;
			/* */
			/*<font> <b val="1"/> <u val="1"/> <i val="1"/> <strike val="1"/> <sz val="8"/> <color rgb="ffff0000"/> <name val="Arial"/> <family val="2"/> </font>*/
			if(font !== null) {
				var font_o={};//new Object();
				font_o={
							b:0,
							i:0,
							u:'none',
							strike:0,
							family:2,
							sz:10,
							color:'0',
							name:'Arial'	
						};
				
				for(prop in font) {
					if(prop === 'color') {
						font_o[prop]='ff'+font[prop];
					} else {
						font_o[prop]=font[prop];
					}
				}
				font_id=s_fonts.length;
				if(style.font_id === undefined) {
					style.font_id=font_id;
				}
				s_fonts.push(font_o);
			}
			
			/*<fill> <patternFill patternType="solid"> <fgColor rgb="ffaabbcc"/> <bgColor rgb="ffaabbcc"/> </patternFill> </fill>*/
			if(fill !== null) {
				var fill_o={};//:Object=new Object();
				fill_o={
							fg_clr:'ff000000',
							bg_clr:'ffffffff'	
						};
				for(prop in fill) {
					fill_o[prop] = 'ff' + fill[prop];
				}
				fill_id = s_fills.length;
				if(style.fill_id === undefined) {
					style.fill_id=fill_id;
				}
				s_fills.push(fill_o);
			}
				
			/*<border> 
				<left style="thin"> <color rgb="64"/> </left>
				<right style="thin"> <color rgb="64"/> </right>
				<top style="thin"> <color rgb="64"/> </top>
				<bottom style="thin"> <color rgb="64"/> </bottom>
				<diagonal/>
			</border>*/
			if(border !== null) {
				var border_o={};//:Object=new Object();
				border_o={
							l_style:'none',r_style:'none',
							t_style:'none',b_style:'none',
							d_style:'none',
							l_clr:'0',r_clr:'0',
							t_clr:'0',b_clr:'0',
							d_clr:'0'
						};
				for(prop in border)	{
					if(prop.indexOf('clr') !== -1)	{
						border_o[prop] = 'ff' + border[prop];
					}else {
						border_o[prop] = 'ff' + border[prop];
					}
				}
				border_id=s_borders.length;
				if(style.border_id === undefined) {
					style.border_id = border_id;
				}
				s_borders.push(border_o);
			}
			
			/*<xf borderId="1" fillId="3" fontId="1" numFmtId="177" applyAlignment="1" applyBorder="1" applyFill="1" applyFont="1" applyNumberFormat="1">
				<alignment wrapText="1" textRotation="90" vertical="center" horizontal="center"/>
			</xf>*/
			/*
			'<xf borderId="{border_id}" applyBorder="{use_border}" fillId="{fill_id}" applyFill="{use_fill}" fontId="{font_id}" applyFont="{use_font}" applyAlignment="{use_txt_aling}">'+
				'<alignment wrapText="{wrap_txt}" textRotation="{txt_rot}" vertical="{txt_v_align}" horizontal="{txt_h_align}"/>'+
				'</xf>'
			*/
		
			var style_o={};//:Object=new Object();
			style_o={
						border_id: 0,
						use_border: 0,
						fill_id: 0,
						use_fill: 0,
						font_id: 0,
						use_font: 0,
						use_txt_aling: 0,
						txt_rot: 0,
						txt_v_align: 'center',
						txt_h_align: 'center',
						txt_wrap: 0
					};
			
			if(style.border_id !== undefined) {
				style_o.border_id = style.border_id;
				style_o.use_border = 1;
			}
			if(style.fill_id !== undefined) {
				style_o.fill_id = style.fill_id;
				style_o.use_fill = 1;
			}
			if(style.font_id !== undefined) {
				style_o.font_id = style.font_id;
				style_o.use_font = 1;
			}
			
			if(style.txt_rot !== undefined){
				style_o.txt_rot = style.txt_rot;
				style_o.use_txt_align = 1;
			}
			if(style.txt_v_align !== undefined){
				style_o.txt_v_align = style.txt_v_align;
				style_o.use_txt_align = 1;
			}
			if(style.txt_h_align !== undefined){
				style_o.txt_h_align = style.txt_h_align;
				style_o.use_txt_align = 1;
			}
			if(style.txt_wrap !== undefined){
				style_o.txt_wrap = style.txt_wrap;
				style_o.use_txt_align = 1;
			}
			
			_id = styles.length;
			styles.push(style_o);
			
			return _id;
		},

	get_sheet:function(sheet_name) {
		return this.sheet_ref(sheet_name);
	},
		

	add_sheet:function(sheet_name) {
		var sheet_id=sheets.length;
		var sh={
				id:sheet_id,
				name:sheet_name,
				data:[],
				totCols:10,
				totRows:10
				//cols:[],//new ArrayList<Object>());
				//hyperlinks:[]//new ArrayList<Object>());
				};
		sheets.push(sh);
			
		return sheet_id;
	},
	
	//public void add_hyperlink(int sheetID,HashMap<String,String> hyperlink)
	add_hyperlink:function(sheet_id,hyperlink) {
		
		//Map<String,Object> sh=(HashMap<String,Object>) sheets.get(sheetID);
		var sh=sheets[this.sheet_ref(sheet_id)];
		//List<Object> sh_data=(List<Object>) sh.get("hyperlinks");
		if(sh.hyperlinks === undefined) {
			sh.hyperlinks = [];
		}
		sh.hyperlinks.push(hyperlink);
		//sh_data.add(hyperlink);
	},
	
	add_col:function(sheet_id, col) {

		var sh=sheets[this.sheet_ref(sheet_id)];
		if(sh.colls === undefined) {
			sh.colls = [];
		}
		sh.colls.push(col);
	},
	
	split_view:function(sheet_id, x, y) {
		
		var sh=sheets[this.sheet_ref(sheet_id)];
		sh.split_view = [x, y];
	},

	check_headerFooter:function(data) {
		
		//console.log('---',data);
		var checks = ['left', 'center', 'right'];
		for(var a=0; a<3; a++) {
			var check = checks[a];
			if((data[check] === undefined) || (data[check] === null)) {
				data[check] = '';
			}
		}
		//console.log('===',data);
		return data;
	},
	
	add_header:function(sheet_id, data) {

		var sh=sheets[this.sheet_ref(sheet_id)];
		data = this.check_headerFooter(data);
		
		sh.header = {
			left: '&amp;L' + data.left,
			center: '&amp;C' + data.center,
			right: '&amp;R' + data.right
			};
	},

	add_footer:function(sheet_id, data) {
		
		var sh=sheets[this.sheet_ref(sheet_id)];
		data = this.check_headerFooter(data);

		sh.footer = {
			left: '&amp;L' + data.left,
			center: '&amp;C' + data.center,
			right: '&amp;R' + data.right
			};
	},
	
	add_chart:function(sheet_id, chart, bbox) {
		
		sheet_id = this.sheet_ref(sheet_id);
		var chart_id = charts.length;
		
		charts.push({
					id: chart_id,
					sheet_id: sheet_id,
					data: chart,
					bbox: bbox,
					media_selector: 'chart'
				});
	},
		
	add_image:function(sheet_id, image_data) {//, bbox) {
		
		sheet_id = this.sheet_ref(sheet_id);
		var image_id = images.length;
		images.push({
					id:image_id,
					sheet_id:sheet_id,
					data: image_data.path,
					bbox: image_data.bbox,
					wd:	image_data.wd,
					hg: image_data.hg,
					dpi: image_data.dpi, 
					scale: image_data.scale,
					media_selector:'image'
				});
	},

	merge_cell:function(sheet_id, mc) {
		
		var sh=sheets[this.sheet_ref(sheet_id)];
		if(sh.merged_cells === undefined) {
			sh.merged_cells = [];
		}
			
		if(mc.constructor === Array) {
			mc = this.abc_id(mc[0]-1) + (mc[1]) + ':' + this.abc_id(mc[2] -1 + mc[0] - 1) + (mc[3] + mc[1] - 1);
		} 
			
		sh.merged_cells.push(mc);
	},
	
/*
public void add_table(final int sheet_id,final String ref,final List<Object> table)
	{
		final int table_id=tables.size();

		   //put("table_id",Integer.toString(table_id+1));
		   //put("ref",ref);
		for(int t=0;t<table.size();t++)
		{
			HashMap<String,String> tc=(HashMap<String,String>) table.get(t);
			tc.put("table_id",Integer.toString(table_id+1));
			tc.put("ref",ref);
			//tc.put("autoFilter","1");
			
		}
		
		tables.add(new HashMap<String,Object>(){{
						put("id",table_id);
						put("sheet_id",sheet_id);
						//put("autoFilter",table);
						put("tableColumns",table);
						//put("tableStyleInfo",true);			
						}});
		
		Map<String,Object> sh=(HashMap<String,Object>) sheets.get(sheet_id);
		sh.put("table_id",table_id);
		sh.put("tableParts","1");
	}*/
	
	//public function add_row(sheet:*,row:*,row_options:*=null):void
	add_row:function(sheet, row, order) {
		
		//<row r="1" spans="1:57" s="1" customFormat="1" ht="15" x14ac:dyDescent="0.2">
		//<row r="1" s="1" customFormat="1">
			
		var sh=sheets[this.sheet_ref(sheet)];
			
		if((typeof(row)==='string')||(row instanceof String)) {
			sh.data.push(row);
		} else {
			var opts = undefined;
			if ( (row.constructor === Object) && (row.opts!==undefined) ) {
				opts = row.opts;
				row = row.data;
			}
			var ht = '';
			var s = '';
			var row_style = '';
			if(opts !== undefined) {
				if(opts.hg !== undefined) {
					ht = ' ht="' + opts.hg + '" customHeight="1"';
				}
					
				if(opts.style !== undefined) {
					row_style = opts.style;
				}
					
				if(opts.row_style !== undefined) {
					s = ' s="' + opts.row_style + '" customFormat="1"';
				}
			}
				
			if (order === undefined) {
				order = sh.data.length + 1;
			}
				
			var row_str =['<row r="', order , '"' , ht , s , '>'];
			var _ln=row.length;
			if(_ln>sh.totCols) {
				sh.totCols=_ln;
			}
			if(order>sh.totRows) {
				sh.totRows=order;
			}
			for(var ee=0; ee<_ln; ee++)
			{
				var cell_order = ' r="' + this.abc_id(ee) + order +'"';
				var item = [''];	
				if ((row[ee]!==undefined)&&(typeof(row[ee])==='string')) {
					item=(row[ee]).split('~|~');
				}
				var val=item[0];
				var style='';
					
				if(row_style!=='') {
					style=' s="' + row_style + '" ';
				}
					
				if(item.length>1) {
					style = ' s="' + item[1] + '" ';
				}

				if((!(val <= 0) && !(val > 0))||(val>9.9E+100)){
					if(val.length>32000) {
						val = val.substr(0,32000)+'.....';
					}
					var s_str_id=_sh_str[val];
					if(s_str_id === undefined) {
						_sh_str[val] = totalShStr;
						s_str_id = totalShStr;
						totalShStr++;
					}
					row_str.push('<c', cell_order , ' t="s"' , style , '><v>' , s_str_id , '</v></c>');
				} else {
					var test0x = val.split('0x');
					if (test0x.length==1) {
						row_str.push('<c', cell_order , style , '><v>' , val , '</v></c>');
					} else {
						row_str.push('<c', cell_order , style , ' t="inlineStr"><is><t>' , val , '</t></is></c>');
					}
				}
			}
			row_str.push('</row>');
			sh.data.push(row_str.join(''));
			row_str=null;
		}
	},

	add_row_no_check:function(sheet, row, order, string, _opts) {
		
		var sh=sheets[this.sheet_ref(sheet)];
			
		if(string) {
			sh.data.push(row);
		}else
		{
			var _row = null;
			var ht = '';
			var s = '';
			var row_style = '';
			if(_opts) {
				var opts = row.opts;
				_row = row.data;
				if(opts.hg !== undefined) {
					ht = ' ht="' + opts.hg + '" customHeight="1"';
				}
					
				if(opts.style !== undefined) {
					row_style = opts.style;
					//s = ' s="' + opts.style + '" customFormat="1"';
				}
				
				if(opts.row_style !== undefined) {
					s = ' s="' + opts.row_style + '" customFormat="1"';
				}
			} else {
				_row = row;
			}
			var row_str =[ '<row r="' , order , '"' , ht , s , '>'];
			var _ln=_row.length;
			if(_ln>sh.totCols) {
				sh.totCols=_ln;
			}
			if(order>sh.totRows) {
				sh.totRows=order;
			}
	
			for(var ee=0; ee<_ln; ee++)
			{
				var cell_order = ' r="' + this.abc_id(ee) + order +'"';
				var __row=_row[ee];
					
				var style='';
				if(row_style!=='') {
					style=' s="' + row_style + '" ';
				}

				if (__row.indexOf('~|~')!==-1) {

					var item=__row.split('~|~');
					__row=item[0];
					style = ' s="' + item[1] + '" ';
				}

				if((!(__row <= 0) && !(__row > 0))||(__row>9.9E+100)){
					if(__row.length>32000) {
						__row = __row.substr(0,32000)+'.....';
					}
					var s_str_id=_sh_str[__row];
					if(s_str_id === undefined) {
						_sh_str[__row] = totalShStr;
						s_str_id = totalShStr;
						totalShStr++;
					}
					row_str.push('<c' , cell_order , ' t="s"' , style , '><v>' , s_str_id , '</v></c>');
				}else {
					var test0x = __row.split('0x');
					if (test0x.length==1) {
						row_str.push('<c', cell_order , style , '><v>' , __row , '</v></c>');
					} else {
						row_str.push('<c', cell_order , style , ' t="inlineStr"><is><t>' , __row , '</t></is></c>');
					}
				}
			}
			row_str.push( '</row>');
			sh.data.push(row_str.join(''));
			row_str=null;
		}
	},
		
	//private function sheet_ref(sheet:*):uint
	sheet_ref:function(sheet) {
			var sheet_id=undefined;
			//console.log('check sheet',sheet);
			if((typeof(sheet)==='string')||(sheet instanceof String)) {
				//console.log('string');
				sh=sheet.toLowerCase();
				for(var aa=0;aa<sheets.length;aa++) {
					if(sheets[aa].name.toLowerCase()===sh) {
						sheet_id=aa;
						break;
					}
				}
			} else {
				sheet_id=sheet;
			}
			
			return sheet_id;
		},
		
	get_xlsx:function() {
			pack();
			
			return _zip;	
		},
			
	add_file:function(fileName, data) {
			var fPath=fileName.split('/');
			var path='';
			for(var a=0;a<fPath.length-1;a++) {
				path+=fPath[a]+'/';
			}
			
			var fNm=fPath[fPath.length-1];
			if(fPath!==''){
				var f = _zip.folder(path);
				f.file(fNm, data);
			}		else{
				_zip.file(fNm, data);
			}
		},
	
	add_media:function(mediaId, fileName) {
			var fs = require('fs');
			var data = fs.readFileSync(fileName);
			
			var fNm = 'image' + mediaId + '.' + fileName.split('.')[1];
			var path = 'xl/media/';
			
			//console.log('file->>>>>>>',data);
			//console.log('media name->',fNm);
			//console.log('path->>>>>>>',path);
			
			var f = _zip.folder(path);
			f.file(fNm, data);

			return fNm;
		},

	xml2ba:function(xml, placeHolder, str) {
		
			//console.log('start xml2ba->',_timer());
			
			var rez='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
					new srl().serializeToString(xml, 'text/xml');;
			
			//console.log('serialization done->',_timer());
			
			if(!easy_view) {
				rez=rez.replace(/\n/g,'');
			}
			//console.log('regexp doen->',_timer());
			
			if ((placeHolder !== undefined) && (str !== undefined)) {
				var r = rez.split(placeHolder);
				if (r.length == 2) {
					rez = r[0] + str + r[1];
				}
			}
			//console.log('concat done->',_timer());
			return rez;
		},		
		
       /*
	xml2ba:function(xml) {
	                console.log('serialize start',new Date());
			var rez='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+//xml.toString();
					new srl().serializeToString(xml, 'text/xml');;
			//console.log('easy view');
			if(!easy_view) {
				//rez=rez.replace(/\n/g,'');
			}
			console.log('serialize end',new Date());
			//console.log('all done');
			return rez;
		},     */
		
	//private function pack():void 
	pack:function(cb, fName)	{
			// create xlsx structure and pack
			//var x:XLSXDef=new XLSXDef();
			
			//console.log('x',x);	//require('./XLSXDef.js');
			console.log('start sheeting!',_timer());
			if(charts.length==0) {
				charts=null;
			} else {
				for(var aa=0;aa<charts.length;aa++) {
					var chart=charts[aa];
					var drawing_id=-1;
					var bbox=chart.bbox;
						
					for(var nn=0;nn<drawings.length;nn++) {
						if(drawings[nn].sheet_id==chart.sheet_id) {
							drawing_id=nn;
							break;
						}
					}
					
					if(drawing_id==-1) {
						if(bbox==null) {
							bbox=[0,0,20,20];
						}
						drawing_id=drawings.length;
						drawings.push({
												id:drawing_id,
												sheet_id:chart.sheet_id,
												chart_id:chart.id,
												data:[{
															col_from:bbox[0],col_to:bbox[2],
															row_from:bbox[1],row_to:bbox[3],
															chart_id:chart.id+1,
															media_selector:'chart'
														}],
												media:[chart]
											});
					} else {
						var drawing_data_ln=drawings[drawing_id].data.length;
						if(bbox==null) {
							bbox=[0,drawing_data_ln*20,20,drawing_data_ln*20+20];
						}
						drawings[drawing_id].data.push({
																col_from:bbox[0],col_to:bbox[2],
																row_from:bbox[1],row_to:bbox[3],
																chart_id:chart.id+1,
																media_selector:'chart'
															});
						drawings[drawing_id].media.push(chart);
					}
					
					sheets[chart.sheet_id].drawing_id = drawing_id;
					
					this.add_file('xl/charts/chart'+(aa+1)+'.xml',this.xml2ba(charts[aa].data));
				}
			
			}
			
			if(images.length == 0) {
				images=null;
			} else {
				for(var aa=0; aa<images.length; aa++) {
					var image = images[aa];
					var drawing_id = -1;
					var bbox= image.bbox;
						
					for(var nn=0; nn<drawings.length; nn++) {
						if(drawings[nn].sheet_id == image.sheet_id) {
							drawing_id=nn;
							break;
						}
					}
					
					if(drawing_id == -1) {
						if(bbox == null) {
							bbox = [0,0,20,20];
						}
						drawing_id = drawings.length;
						drawings.push({
												id:drawing_id,
												sheet_id:image.sheet_id,
												image_id:image.id,
												data:[{
															col_from: bbox[0], col_to: bbox[2],
															row_from: bbox[1], row_to: bbox[3],
															image_wd: Math.round(image.wd/image.dpi*emu*image.scale),
															image_hg: Math.round(image.hg/image.dpi*emu*image.scale),
															image_id: image.id+1,
															media_selector: 'image'
														}],
												media:[image]
											});
					} else {
						var drawing_data_ln = drawings[drawing_id].data.length;
						if(bbox == null) {
							bbox=[0,drawing_data_ln*20,20,drawing_data_ln*20+20];
						}
						drawings[drawing_id].data.push({
																col_from: bbox[0], col_to: bbox[2],
																row_from: bbox[1], row_to: bbox[3],
																image_wd: Math.round(image.wd/image.dpi*emu*image.scale),
																image_hg: Math.round(image.hg/image.dpi*emu*image.scale),
																image_id: image.id+1,
																media_selector: 'image'
															});
						drawings[drawing_id].media.push(image);
					}
					
					sheets[image.sheet_id].drawing_id = drawing_id;
					
					images[aa].image_id=this.add_media((aa+1),images[aa].data);
					//this.add_file('xl/charts/chart'+(aa+1)+'.xml',this.xml2ba(charts[aa].data));
				}
			
			}
			
			if(drawings.length==0) {
				drawings=null;
			} else {
				for(var bb=0;bb<drawings.length;bb++) {
					var drawing_file=x.get_xml('drawing',null,drawings[bb].data);
					this.add_file(drawing_file[0]+(bb+1)+'.'+drawing_file[1],this.xml2ba(drawing_file[2]));
					
					var drawing_rels_file=x.get_xml('drawing_rels',null,null,drawings[bb].media);
					this.add_file(drawing_rels_file[0]+(bb+1)+'.'+drawing_rels_file[1],this.xml2ba(drawing_rels_file[2]));
				}
			}

			var xlsx_files=['_rels','content_types','app','core','workbook_rel','workbook'];
			var tmp_file=null;
			
			for(aa=0;aa<xlsx_files.length;aa++) {
				tmp_file=x.get_xml(xlsx_files[aa],sheets,drawings,charts);
				this.add_file(tmp_file[0]+'.'+tmp_file[1],this.xml2ba(tmp_file[2]));	
			}
			//console.log('get extra xml->',_timer());

			//var xlsx_extra_files=['styles', 'shared_strings', 'theme'];
			var xlsx_extra_files=['styles', 'theme'];
			var tmp_shstr=Object.keys(_sh_str);
			tmp_shstr=tmp_shstr.join('#@#').split('&').join('&amp;').split('<').join('&lt;').split('>').join('&gt;').split('#@#');

			var shared_strings_data='<si><t>'+tmp_shstr.join('</t></si><si><t>')+'</t></si>';
			console.log('shared strings ln---->',tmp_shstr.length,totalShStr);
			tmp_shstr = null;
			/*for(var ss=0;ss<sh_str.length;ss++) {
				sh_str[ss]={str:sh_str[ss]};	
			}*/
			
			for(aa=0; aa<xlsx_extra_files.length; aa++) {
				tmp_file = x.get_extra_xml(xlsx_extra_files[aa], s_fonts, s_fills, s_borders, styles, null);//sh_str);
				this.add_file(tmp_file[0]+'.'+tmp_file[1],this.xml2ba(tmp_file[2]));	
			}
			tmp_file = x.get_extra_xml('shared_strings', s_fonts, s_fills, s_borders, styles, '___shared-strings___');
			this.add_file(tmp_file[0]+'.'+tmp_file[1],this.xml2ba(tmp_file[2],'___shared-strings___',shared_strings_data));	
			//console.log('extra xml ok->',_timer());

			var sheet_file=null;
			
			for(aa=0;aa<sheets.length;aa++) {
					
				/*sheets_group:'sheetData',
				merge_cell_group:'mergeCells',
				marge_cell:'<mergeCell ref={ref}/>',*/
				
				//-----------------------------------------
				/*sheets_group:'sheetData',
				merge_cell_group:'mergeCells',
				marge_cell:'<mergeCell ref={ref}/>',*/
			
				//Map<String,Object> sh=(Map<String,Object>) sheets.get(aa);
				//List<Object> sh_data=(List<Object>) sh.get("data");
				var sh = sheets[aa];
				var sh_data = sh.data;
			
				var cols = '';
				//"<col min=\"{c_id}\" max=\"c_id\" width=\"{wd}\" bestFit=\"1\" collapsed=\"{clps}\" outlineLevel=\"{o_lvl}\" customWidth=\"1\"/>"
				//<col min="9" max="9" width="3.5" style="3" customWidth="1"/>
				//<col min="10" max="10" width="24" bestFit="1" customWidth="1"/>

				if(sh.colls !== undefined) {
					var mc_list = sh.colls;
				
					for(var mm=0; mm<mc_list.length; mm++) {
					
						var col = mc_list[mm];
						cols += '<col ' +
								'min="' + (mm+1) + '" ' +
								'max="' + (mm+1)+ '" ' +
								'bestFit="1" ';
								
						if(col.wd !== undefined) {
							cols += 'width="' + col.wd + '" customWidth="1" ';
						}
						if(col.clpsd !== undefined) {
							cols += 'collapsed="' + col.clpsd + '" ';
						}
						if(col.ol_lvl !== undefined) {
							cols += 'outlineLevel="' + col.ol_lvl + '" ';
						}
						if(col.style !== undefined) {
							cols += 'style="' + col.style + '" ';
						}
						
						cols += '/>';
				}
				
				if(cols != '') {
					cols = '<cols>' + cols + '</cols>';
				}
			}

			
			//console.log('columns->>>>>>',cols);
			sh_data.unshift('<sheetData>');
			sh_data.unshift(cols);
			
			if(sh.split_view !== undefined) {
				sh_data.unshift('<sheetViews>' +
					//'<sheetView tabSelected="1" workbookViewId="0">' +
					'<sheetView workbookViewId="0">' +
						'<pane xSplit="' + sh.split_view[0] + '" ySplit="' + sh.split_view[1] + '" topLeftCell="'+ this.abc_id(sh.split_view[0]+1) + (sh.split_view[1]+1) + '" activePane="bottomRight" state="frozen"/>' +
					'</sheetView>' +
				'</sheetViews>');
			}
			
			//sh_data.unshift('<dimension ref="A1:ZZ400000"/>');
			var shDm = 'A1:'+this.abc_id(sh.totCols+1)+(sh.totRows+1);
			sh_data.unshift('<dimension ref="'+shDm+'"/>');
			console.log('sheet dimension---',shDm);
			//String sheet_data="<sheetData>"+sh_data.toString()+"</sheetData>";
			//System.out.println("adding sheet data----------> "+sheet_data);
			//System.out.println("adding sh data-------------> "+sh_data);
			sh_data.push('</sheetData>');
			
			var mc='';
			if(sh.merged_cells !== undefined){
				var mc_list = sh.merged_cells;
				//console.log("adding merged cells-------------> "+mc_list);
				
				for(var mm=0; mm<mc_list.length; mm++) {
					mc += '<mergeCell ref="'+mc_list[mm]+'"/>';
				}
				mc = '<mergeCells>' + mc + '</mergeCells>';
			}
			sh_data.push(mc);
		
			
			
			var h = '';
			if(sh.header !== undefined) {
				h = '<oddHeader>' + sh.header.left + sh.header.center + sh.header.right + '</oddHeader>';
			}
			
			var f = '';
			if(sh.footer !== undefined) {
				f = '<oddFooter>' + sh.footer.left + sh.footer.center + sh.footer.right + '</oddFooter>';
			}
			
			var hf = '';
			if((h !== '')||(f !== '')) {
				hf = '<headerFooter>' + h + f + '</headerFooter>';
			}
			sh_data.push(hf);
			/*
			</sheetData>
			<mergeCells/>
			<headerFooter>
				<oddHeader>&amp;L left-top &amp;C center-top &amp;R right-top</oddHeader>
				<oddFooter>&amp;L btm-left &amp;F &amp;C btm-cntr &amp;A &amp;R  lft-cntr &amp;P of &amp;N</oddFooter>
			</headerFooter>
			<drawing/>
			*/
			//console.log('h->',h,'f->',f,'hf->',hf);
			
			var hyperlinks='';
			/*<hyperlinks>
			<hyperlink ref="A1" display="link" location="'Sheet Nr2'!a1" toolTip="Bla bla bla!!!!!"/>
			</hyperlinks>*/
			if(sh.hyperlinks !== undefined) {
				var hl_list = sh.hyperlinks;
				
				for(var mm=0; mm<hl_list.length; mm++) {
					var h_link = hl_list[mm];
					hyperlinks += '<hyperlink ' +
								'ref="' + h_link.ref + '" ' +
								'display="' + h_link.display + '" ' +
								'location="' + h_link.location + '" ' +
								'/>';
				}
				if(hyperlinks !== '') {
					hyperlinks = '<hyperlinks>' + hyperlinks + '</hyperlinks>';
				}
			}
			sh_data.push(hyperlinks);
				
				
				
				//var sheet_data = '<sheetData>'+sheets[aa].data.join('')+'</sheetData>';
				//console.log('join sheet data',sheets[aa].data.length);
				var sheet_data = sheets[aa].data.join('');
				//console.log('done',sheet_data.length);
				/*
				if(sheets[aa].merged_cells!=undefined) {
					var mc='';
					for(var mm=0;mm<sheets[aa].merged_cells.length;mm++) {
						mc+='<mergeCell ref="'+sheets[aa].merged_cells[mm]+'"/>';
					}
					sheet_data+='<mergeCells>'+mc+'</mergeCells>';
				}*/
				

				//console.log('get sheets xml->',_timer());
				if(sheets[aa].drawing_id !== undefined) {
					var sheet_rels_file=x.get_xml('sheet_rels',null,[drawings[sheets[aa].drawing_id]]);
					this.add_file(sheet_rels_file[0]+(aa+1)+'.'+sheet_rels_file[1],this.xml2ba(sheet_rels_file[2]));
					
					sheet_file=x.get_xml('sheet','___sheet-data___',[drawings[sheets[aa].drawing_id]]);
				} else {
					sheet_file=x.get_xml('sheet','___sheet-data___');
				}
				//console.log('xml ok->',_timer());
				this.add_file(sheet_file[0]+(aa+1)+'.'+sheet_file[1],this.xml2ba(sheet_file[2],'___sheet-data___',sheet_data));
			}
			console.log('go to zip it!',_timer());
			var fNM = this.close(cb, fName);
			return fNM;
		},
		
	close:function(cb, _fName)
		{
			sheets=null;
			drawings=null;
			charts=null;
			images=null;
			
			s_fonts=null;
			s_fills=null;
			s_borders=null;
			styles=null;

			sh_str=null;
			_sh_str=null;
			_sh_str_big=null;
			_orders=null;
			totalShStr=0;
			//console.log('zip it, fn->');
			var fs = require("fs");
			
			//var buffer = zip.generate({type:"nodebuffer"});
			//fs.writeFile("test.zip", buffer, function(err) {
			// 	if (err) throw err;
			//});

			if (cb !== undefined) {
				console.log('start zip',_timer());
				var data = _zip.generate({type:"nodebuffer"});//,compression:'DEFLATE'});//,compressionOptions:{level:9}});//{base64:false,compression:'DEFLATE'});
				var tempName = (_fName!=undefined)?_fName:('tmp_xlsx_'+Math.floor(new Date() / 1000).toString());
				var tmpFname = tmpFolder + tempName;
				console.log('zip done, fn->',tmpFname,_timer());
				//console.log('write file');
				//fs.writeFileSync(tmpFname+'.xlsx', data, 'binary');

				//fs.writeFileSync(tmpFname, data, 'binary');
				fs.writeFile(tmpFname, data, function(err) {
				 	if (err) {
						console.log('can not crate zip!',_timer());
						throw err;
					}
					console.log('xslxx exp. async pack done',tmpFname,_timer());
					_zip=null;
					cb(tmpFname);
				});
			} else {
				var data = _zip.generate({base64:false});//,compression:'DEFLATE'});
				var tempName = (_fName!=undefined)?_fName:('tmp_xlsx_'+Math.floor(new Date() / 1000).toString());
				var tmpFname = tmpFolder + tempName;

				//console.log('zip done, fn->',tmpFname);
				//console.log('write file');
				//fs.writeFileSync(tmpFname+'.xlsx', data, 'binary');

				fs.writeFileSync(tmpFname, data, 'binary');
				_zip=null;
			        console.log('xslxx exp. SYNC pack done',tmpFname,_timer());
		        	return tmpFname;
			}
			//console.log('write done');
			//_zip=null;
			
		        //return tmpFname;
		}
	}
}


/*
Column widths in Excel 2007/2010 don’t follow any of the above. The width units are defined in the OOXML Standard (section 3.13.1.2).
The column width number represents, roughly, the number of *characters* that would fit in that column, accounting for padding, and rounded to 1/256th.
Padding is always 2 pixels on each side of the cell, plus 1 pixel for the gridlines, so 5 pixels.
The width value you would use in the XML would be:

width = Truncate( (NumChars * MaxWidth + Padding) / MaxWidth * 256) / 256

Let’s say you want to fit an average of 10 characters into the cell.
First, you would need to know the maximum width of a digit (0-9) using the spreadsheet’s *default* font. For 11pt Calibri, the maximum width of a digit is 7 pixels at 96dpi. So:

width = Truncate( (10 * 7 + 5) / 7 * 256) / 256 = 10.7109375 



    Get the image's dimensions and resolution
    Compute the image's width in EMUs: wEmu = imgWidthPixels / imgHorizontalDpi * emuPerInch
    Compute the image's height in EMUs: hEmu = imgHeightPixels / imgVerticalDpi * emuPerInch
    Compute the max page width in EMUs (I've found that if the image is too wide, it won't show)

    If the image's width in EMUs is greater than the max page width's, I scale the image's width and height so that the width of image is equal to that of the page (when I say page, I'm referring to the "usable" page, that is minus the margins):

    var ratio = hEmu / wEmu;
    wEmu = maxPageWidthEmu;
    hEmu = wEmu * ratio;

    I then use the width as the value of Cx and height as the value of Cy, both in the DW.Extent and in the A.Extents (of the A.Transform2D of the PIC.ShapeProperties).

	Note that the emuPerInch value is 914400.

*/
