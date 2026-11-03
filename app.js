/* Văn phòng của Bùi — bản web 1.2. HTML/JS thuần, không framework build.
   Không thu thập dữ liệu, không nhúng key/secret. File luôn lưu ở máy trước. */
'use strict';
var FONTS=['Carlito','Be Vietnam Pro','Inter','Roboto','Noto Sans','Noto Serif'];
var $=function(id){return document.getElementById(id)};
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function uid(){return 'f'+Date.now().toString(36)+Math.floor(Math.random()*1e6).toString(36)}
function loadJSON(k,f){try{var v=JSON.parse(localStorage.getItem(k));return v==null?f:v}catch(e){return f}}
function saveJSON(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){alert('Không lưu được vào máy (bộ nhớ trình duyệt đầy hoặc bị chặn): '+e.message)}}
function downloadBlob(blob,name){var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},1000)}
function downloadText(text,name,type){downloadBlob(new Blob([text],{type:type||'text/plain;charset=utf-8'}),name)}

/* ---------- Chuyển tab (bấm bằng tay) ---------- */
document.querySelectorAll('#tabs button').forEach(function(b){
  b.addEventListener('click',function(){
    document.querySelectorAll('#tabs button').forEach(function(x){x.classList.remove('active')});
    b.classList.add('active');
    document.querySelectorAll('section.tab').forEach(function(s){s.classList.remove('active')});
    $('tab-'+b.dataset.tab).classList.add('active');
  });
});
function gotoTab(name){var m={editor:'word',sheet:'excel',slides:'ppt',pdf:'pdf',translate:'translate',files:'files',auth:'auth',settings:'settings'};if(m[name])switchApp(m[name]);}

/* ---------- Bottom sheet dùng chung (thay menu thả xuống nhỏ) ---------- */
function openSheet(title,items){
  $('sheetTitle').textContent=title;
  var body=$('sheetBody');body.innerHTML='';
  items.forEach(function(it){
    var btn=document.createElement('button');btn.innerHTML=it.label;
    btn.addEventListener('click',function(){closeSheet();it.onClick();});
    body.appendChild(btn);
  });
  $('sheetBackdrop').classList.add('open');$('sheet').classList.add('open');
}
function closeSheet(){$('sheetBackdrop').classList.remove('open');$('sheet').classList.remove('open')}
function openFontSheet(){openSheet('Chọn font chữ',FONTS.concat(['Font hệ thống (-apple-system)']).map(function(f){
  var v=(f.indexOf('hệ thống')>=0)?'system':f;
  return {label:esc(f),onClick:function(){applyFont(v);$('fontLabel').textContent=f}};
}))}
function openSizeSheet(){openSheet('Chọn cỡ chữ',[14,16,18,20,24,28].map(function(s){
  return {label:s+' pt',onClick:function(){applySize(String(s))}};
}))}
function openColorSheet(kind){
  var colors=[['#1a1a1a','Đen'],['#c00000','Đỏ'],['#0b6fa4','Xanh dương'],['#1a7a2e','Xanh lá'],['#ffeb3b','Vàng'],['#ffc9c9','Hồng nhạt'],['#cfe8ff','Xanh nhạt'],['#ffffff','Trắng']];
  openSheet(kind==='fore'?'Chọn màu chữ':'Chọn màu tô sáng',colors.map(function(c){
    return {label:'<span class="swatch" style="background:'+c[0]+'"></span>'+esc(c[1]),onClick:function(){
      if(kind==='fore')cmd('foreColor',c[0]);else{try{document.execCommand('hiliteColor',false,c[0])}catch(e){document.execCommand('backColor',false,c[0])}}
    }};
  }));
}
function openLineSheet(){openSheet('Khoảng cách dòng',[[ '1,15','1.15'],['1,5','1.5'],['2','2']].map(function(p){
  return {label:'Dòng '+p[0],onClick:function(){setLineHeight(p[1])}};
}))}
function openHeadingSheet(){openSheet('Tiêu đề / đoạn',[[ 'Đoạn thường','p'],['Tiêu đề 1','h1'],['Tiêu đề 2','h2'],['Tiêu đề 3','h3']].map(function(p){
  return {label:esc(p[0]),onClick:function(){cmd('formatBlock',p[1])}};
}))}
function openDocSheet(){
  var docs=loadJSON('vb_docs_v1',[]);
  if(!docs.length){alert('Chưa có tài liệu nào lưu trên máy.');return}
  openSheet('Mở tài liệu',docs.map(function(d){
    return {label:esc(d.name||'(không tên)'),onClick:function(){openDoc(d.id)}};
  }));
}

/* ---------- ZIP dạng stored (không nén) + CRC32, dùng cho .docx/.xlsx ---------- */
var CRC_TABLE=(function(){var t=[],c;for(var n=0;n<256;n++){c=n;for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c>>>0}return t})();
function crc32(bytes){var c=0xFFFFFFFF;for(var i=0;i<bytes.length;i++)c=CRC_TABLE[(c^bytes[i])&0xFF]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
function strBytes(s){return new TextEncoder().encode(s)}
function zipStored(entries){
  var chunks=[],central=[],offset=0;
  entries.forEach(function(e){
    var data=(typeof e.data==='string')?strBytes(e.data):e.data;
    var name=strBytes(e.name);var crc=crc32(data);
    var lh=new DataView(new ArrayBuffer(30));
    lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0,true);
    lh.setUint16(8,0,true);lh.setUint16(10,0,true);lh.setUint16(12,0,true);
    lh.setUint32(14,crc,true);lh.setUint32(18,data.length,true);lh.setUint32(22,data.length,true);
    lh.setUint16(26,name.length,true);lh.setUint16(28,0,true);
    chunks.push(new Uint8Array(lh.buffer),name,data);
    var ch=new DataView(new ArrayBuffer(46));
    ch.setUint32(0,0x02014b50,true);ch.setUint16(4,20,true);ch.setUint16(6,20,true);
    ch.setUint16(8,0,true);ch.setUint16(10,0,true);ch.setUint16(12,0,true);ch.setUint16(14,0,true);
    ch.setUint32(16,crc,true);ch.setUint32(20,data.length,true);ch.setUint32(24,data.length,true);
    ch.setUint16(28,name.length,true);ch.setUint16(32,0,true);ch.setUint16(34,0,true);
    ch.setUint16(36,0,true);ch.setUint16(38,0,true);ch.setUint32(42,offset,true);
    central.push(new Uint8Array(ch.buffer),name);
    offset+=30+name.length+data.length;
  });
  var csize=central.reduce(function(a,c){return a+c.length},0);
  var end=new DataView(new ArrayBuffer(22));
  end.setUint32(0,0x06054b50,true);end.setUint16(8,entries.length,true);end.setUint16(10,entries.length,true);
  end.setUint32(12,csize,true);end.setUint32(16,offset,true);
  var total=offset+csize+22;var out=new Uint8Array(total);var p=0;
  chunks.concat(central).forEach(function(c){out.set(c,p);p+=c.length});
  out.set(new Uint8Array(end.buffer),p);
  return new Blob([out],{type:'application/zip'});
}

/* ---------- Soạn văn bản ---------- */
var currentDocId=null;
function cmd(c,v){$('editor').focus();document.execCommand(c,false,v==null?null:v)}
function wrapSpan(style){
  var sel=window.getSelection();if(!sel.rangeCount)return;
  var r=sel.getRangeAt(0);if(r.collapsed)return;
  var sp=document.createElement('span');Object.keys(style).forEach(function(k){sp.style[k]=style[k]});
  try{r.surroundContents(sp)}catch(e){
    var frag=r.extractContents();sp.appendChild(frag);r.insertNode(sp);
  }
}
function applyFont(v){
  $('editor').focus();
  if(v==='system'){wrapSpan({fontFamily:'-apple-system, "Segoe UI", sans-serif'});return}
  // execCommand fontName giữ tên font trong HTML để lưu/mở lại và xuất docx đúng tên.
  document.execCommand('fontName',false,v);
  try{var sel2=window.getSelection();if(!sel2||!sel2.rangeCount||sel2.getRangeAt(0).collapsed)$('editor').style.fontFamily="'"+v+"', -apple-system, 'Segoe UI', sans-serif";}catch(e){}
  $('fontLabel').textContent=v;
}
function applySize(px){wrapSpan({fontSize:px+'px'})}
function setLineHeight(v){$('editor').style.lineHeight=v}
function newDoc(){currentDocId=null;$('docName').value='';$('editor').innerHTML='';$('editor').focus()}
function saveDoc(silent){
  var name=$('docName').value.trim()||('Tài liệu '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentDocId)name=uniqueName('doc',name);
  var html=$('editor').innerHTML;
  var docs=loadJSON('vb_docs_v1',[]);
  if(currentDocId){var d=docs.find(function(x){return x.id===currentDocId});
    if(d){d.name=name;d.html=html;d.updatedAt=Date.now()}else{docs.unshift({id:currentDocId,name:name,html:html,updatedAt:Date.now(),syncStatus:'local'})}
  }else{currentDocId=uid();docs.unshift({id:currentDocId,name:name,html:html,updatedAt:Date.now(),syncStatus:'local'})}
  saveJSON('vb_docs_v1',docs);renderDocList();renderFileList();
  if(!silent)alert('Đã lưu vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openDoc(id){
  var d=loadJSON('vb_docs_v1',[]).find(function(x){return x.id===id});if(!d)return;
  currentDocId=d.id;$('docName').value=d.name;$('editor').innerHTML=d.html||'';gotoTab('editor');$('editor').focus();
}
function deleteDoc(id){
  if(!confirm('Xóa tài liệu này khỏi máy?'))return;
  saveJSON('vb_docs_v1',loadJSON('vb_docs_v1',[]).filter(function(x){return x.id!==id}));
  if(currentDocId===id)newDoc();renderDocList();renderFileList();
}
function renderDocList(){
  var docs=loadJSON('vb_docs_v1',[]);
  $('docList').innerHTML=docs.length?docs.map(function(d){
    return '<div class="fileitem"><b>'+esc(d.name)+'</b><span class="small">'+new Date(d.updatedAt).toLocaleString('vi-VN')+'</span>'
      +'<button onclick="openDoc(\''+d.id+'\')">Mở</button>'
      +'<button onclick="deleteDoc(\''+d.id+'\')">Xóa</button></div>';
  }).join(''):'<p class="small">Chưa có tài liệu nào. Soạn rồi bấm “Lưu vào máy”.</p>';
}
function editorText(){var d=document.createElement('div');d.innerHTML=$('editor').innerHTML;return d.innerText||''}
function exportTxt(){downloadText(editorText(),(($('docName').value.trim()||'tai-lieu')+'.txt'))}
/* Phân tích HTML của editor thành đoạn + run có định dạng, để xuất docx */
function hexColor(val){
  if(!val)return null;
  var str=String(val).trim();
  var m=str.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
  if(m)return ((+m[1]).toString(16).padStart(2,'0')+(+m[2]).toString(16).padStart(2,'0')+(+m[3]).toString(16).padStart(2,'0')).toUpperCase();
  var h=str.match(/^#([0-9A-Fa-f]{6})$/);
  if(h)return h[1].toUpperCase();
  var h3=str.match(/^#([0-9A-Fa-f]{3})$/);
  if(h3)return h3[1].split('').map(function(c){return c+c}).join('').toUpperCase();
  return null;
}
function highlightName(hex){
  var map={'FFEB3B':'yellow','FFFF00':'yellow','00FF00':'green','00FFFF':'cyan','FF00FF':'magenta','0000FF':'blue','FF0000':'red'};
  return map[hex]||null;
}
function collectRuns(node,fmt,out){
  if(node.nodeType===3){var t=node.nodeValue;if(t)out.push({text:t,fmt:Object.assign({},fmt)});return}
  if(node.nodeType!==1)return;
  var f=Object.assign({},fmt);var tag=node.tagName.toLowerCase();
  if(tag==='b'||tag==='strong')f.bold=true;
  if(tag==='i'||tag==='em')f.italic=true;
  if(tag==='u')f.underline=true;
  if(tag==='s'||tag==='strike'||tag==='del')f.strike=true;
  if(tag==='font'){
    if(node.getAttribute('face'))f.font=node.getAttribute('face');
    var fc=hexColor(node.getAttribute('color'));if(fc)f.color=fc;
  }
  var st=node.style||{};
  if(st.fontFamily){
    var fam=st.fontFamily.replace(/["']/g,'').split(',')[0].trim();
    if(fam)f.font=(fam.indexOf('apple')>=0||fam.indexOf('Segoe')>=0)?'Carlito':fam;
  }
  if(st.fontSize){
    var px=parseFloat(st.fontSize);
    if(px)f.size=(String(st.fontSize).indexOf('pt')>=0)?Math.round(px*2):Math.round(px*1.5);
  }
  if(st.fontWeight&&parseInt(st.fontWeight)>=700)f.bold=true;
  if(st.fontStyle==='italic')f.italic=true;
  if(st.textDecoration&&st.textDecoration.indexOf('underline')>=0)f.underline=true;
  if(st.textDecoration&&st.textDecoration.indexOf('line-through')>=0)f.strike=true;
  var cc=hexColor(st.color);if(cc)f.color=cc;
  var bg=hexColor(st.backgroundColor);
  if(bg){var hl=highlightName(bg);if(hl)f.highlight=hl;else f.shd=bg;}
  for(var i=0;i<node.childNodes.length;i++)collectRuns(node.childNodes[i],f,out);
}
function docxParaXml(p){
  if(p.pageBreak&&(!p.runs||!p.runs.length))return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  var pPr='<w:pPr>';
  if(p.heading)pPr+='<w:pStyle w:val="'+p.heading+'"/>';
  if(p.align&&p.align!=='left')pPr+='<w:jc w:val="'+(p.align==='justify'?'both':p.align)+'"/>';
  if(p.bullet==='bullet')pPr+='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
  if(p.bullet==='decimal')pPr+='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
  if(p.pageBreak)pPr+='<w:pageBreakBefore/>';
  pPr+='</w:pPr>';
  var runs=p.runs.length?p.runs:[{text:'',fmt:{}}];
  var rs=runs.map(function(r){
    var f=r.fmt||{};var rPr='<w:rPr>';
    if(f.font)rPr+='<w:rFonts w:ascii="'+esc(f.font)+'" w:hAnsi="'+esc(f.font)+'"/>';
    if(f.size)rPr+='<w:sz w:val="'+f.size+'"/><w:szCs w:val="'+f.size+'"/>';
    if(f.bold)rPr+='<w:b/>';if(f.italic)rPr+='<w:i/>';
    if(f.underline)rPr+='<w:u w:val="single"/>';if(f.strike)rPr+='<w:strike/>';
    if(f.color&&/^[0-9A-Fa-f]{6}$/.test(f.color))rPr+='<w:color w:val="'+f.color.toUpperCase()+'"/>';
    if(f.highlight)rPr+='<w:highlight w:val="'+f.highlight+'"/>';
    if(f.shd)rPr+='<w:shd w:val="clear" w:color="auto" w:fill="'+f.shd+'"/>';
    rPr+='</w:rPr>';
    var txt=esc(r.text).replace(/\n/g,'<w:br/>');
    return '<w:r>'+rPr+'<w:t xml:space="preserve">'+txt+'</w:t></w:r>';
  }).join('');
  return '<w:p>'+pPr+rs+'</w:p>';
}
function emptyPara(){return {kind:'para',align:'left',heading:null,runs:[{text:'',fmt:{}}],bullet:null,pageBreak:false}}
function paraFromNode(node,opts){
  opts=opts||{};
  var align='left',heading=null;
  if(node.nodeType===1){
    var tg=node.tagName.toLowerCase();
    if(tg==='h1')heading='Heading1';if(tg==='h2')heading='Heading2';if(tg==='h3')heading='Heading3';
    var a=(node.style&&node.style.textAlign)||node.getAttribute('align')||'';
    if(a==='center'||a==='right'||a==='justify')align=a;
    if(node.style&&node.style.pageBreakBefore==='always')opts.pageBreak=true;
  }
  var runs=[];collectRuns(node,{},runs);
  return {kind:'para',align:align,heading:heading,runs:runs,bullet:opts.bullet||null,pageBreak:!!opts.pageBreak};
}
function docxBlocksFromNode(root){
  var blocks=[];
  function cellParas(tc){
    var out=[];
    Array.from(tc.childNodes||[]).forEach(function(cn){
      if(cn.nodeType===3){if(cn.nodeValue.trim())out.push({kind:'para',align:'left',heading:null,runs:[{text:cn.nodeValue,fmt:{}}],bullet:null,pageBreak:false});}
      else if(cn.nodeType===1){
        var ct=cn.tagName.toLowerCase();
        if(ct==='ul'||ct==='ol'){
          Array.from(cn.childNodes||[]).forEach(function(li){
            if(li.nodeType===1&&li.tagName.toLowerCase()==='li')out.push(paraFromNode(li,{bullet:ct==='ol'?'decimal':'bullet'}));
          });
        }else out.push(paraFromNode(cn));
      }
    });
    if(!out.length)out.push(emptyPara());
    return out;
  }
  Array.from(root.childNodes||[]).forEach(function(b){
    if(b.nodeType===3){if(b.nodeValue.trim())blocks.push({kind:'para',align:'left',heading:null,runs:[{text:b.nodeValue,fmt:{}}],bullet:null,pageBreak:false});return}
    if(b.nodeType!==1)return;
    var tg=b.tagName.toLowerCase();
    if(tg==='ul'||tg==='ol'){
      Array.from(b.childNodes||[]).forEach(function(li){
        if(li.nodeType===1&&li.tagName.toLowerCase()==='li')blocks.push(paraFromNode(li,{bullet:tg==='ol'?'decimal':'bullet'}));
      });
      return;
    }
    if(tg==='table'){
      var rows=[];
      Array.from(b.childNodes||[]).forEach(function(section){
        var stg=section.tagName?section.tagName.toLowerCase():'';
        var trs=(stg==='tr')?[section]:Array.from(section.childNodes||[]).filter(function(n){return n.nodeType===1&&n.tagName.toLowerCase()==='tr'});
        trs.forEach(function(tr){
          var cells=[];
          Array.from(tr.childNodes||[]).forEach(function(tc){
            if(tc.nodeType!==1)return;
            var ct=tc.tagName.toLowerCase();if(ct!=='td'&&ct!=='th')return;
            cells.push(cellParas(tc));
          });
          if(cells.length)rows.push(cells);
        });
      });
      if(rows.length)blocks.push({kind:'table',rows:rows});
      return;
    }
    if(tg==='hr'){blocks.push({kind:'para',align:'left',heading:null,runs:[],bullet:null,pageBreak:true});return}
    if(tg==='br')return;
    blocks.push(paraFromNode(b));
  });
  if(!blocks.length)blocks.push(emptyPara());
  return blocks;
}
function docxXmlFromNode(root){
  var blocks=docxBlocksFromNode(root);
  var body=blocks.map(function(b){
    if(b.kind==='table'){
      var ncols=Math.max.apply(null,b.rows.map(function(r){return r.length}));
      var grid='<w:tblGrid>'+Array.from({length:ncols}).map(function(){return '<w:gridCol w:w="2400"/>'}).join('')+'</w:tblGrid>';
      var trs=b.rows.map(function(row){
        var tcs=row.map(function(cparas){
          return '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'+cparas.map(docxParaXml).join('')+'</w:tc>';
        }).join('');
        return '<w:tr>'+tcs+'</w:tr>';
      }).join('');
      return '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'+grid+trs+'</w:tbl>';
    }
    return docxParaXml(b);
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'</w:body></w:document>';
}
function docxXml(){return docxXmlFromNode($('editor'));}
var DOCX_STYLES='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  +'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito" w:cs="Carlito"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
  +'<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  +'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Carlito" w:hAnsi="Carlito"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="480" w:after="160"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="2B579A"/></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="360" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="2B579A"/></w:rPr></w:style>'
  +'<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
  +'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style></w:styles>';
var DOCX_NUMBERING='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  +'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="0"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl></w:abstractNum>'
  +'<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
  +'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>';
function buildDocxBlobFromNode(root){
  return zipStored([
    {name:'[Content_Types].xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'},
    {name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'},
    {name:'word/_rels/document.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'},
    {name:'word/document.xml',data:docxXmlFromNode(root)},
    {name:'word/styles.xml',data:DOCX_STYLES},
    {name:'word/numbering.xml',data:DOCX_NUMBERING}
  ]);
}
function exportDocx(){
  var name=($('docName').value.trim()||'tai-lieu')+'.docx';
  downloadBlob(buildDocxBlobFromNode($('editor')),name);
}
function testFont(f){$('fontTest').style.fontFamily="'"+f+"', -apple-system, 'Segoe UI', sans-serif"}

/* ---------- Bảng tính: chạm chọn, chạm hai lần sửa, vuốt cuộn ---------- */
var SHEET_ROWS=30,SHEET_COLS=10;
var sheetRaw=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});
var currentSheetId=null,selCell={r:0,c:0},lastTap={key:'',t:0};
function colName(c){return String.fromCharCode(65+c)}
function cellRef(r,c){return colName(c)+(r+1)}
function parseRef(s){var m=/^([A-J])([1-9][0-9]*)$/.exec(s.trim().toUpperCase());if(!m)return null;return {r:parseInt(m[2],10)-1,c:m[1].charCodeAt(0)-65}}
function evalCell(r,c,seen){
  seen=seen||new Set();var key=r+','+c;
  if(seen.has(key))return '#LỖI';
  var raw=sheetRaw[r]&&sheetRaw[r][c]||'';
  if(raw.charAt(0)!=='=')return raw;
  seen.add(key);
  var expr=raw.slice(1).trim();
  var fm=/^(SUM|AVERAGE|COUNT|MAX|MIN)\s*\((.*)\)\s*$/i.exec(expr);
  if(fm){
    var fn=fm[1].toUpperCase();var vals=[];
    fm[2].split(',').forEach(function(part){
      var range=part.trim().split(':');
      if(range.length===2){
        var a=parseRef(range[0]),b=parseRef(range[1]);
        if(a&&b){for(var rr=Math.min(a.r,b.r);rr<=Math.max(a.r,b.r);rr++)for(var cc=Math.min(a.c,b.c);cc<=Math.max(a.c,b.c);cc++){
          var v=evalCell(rr,cc,new Set(seen));var n=parseFloat(v);if(!isNaN(n))vals.push(n);
        }}
      }else{var one=parseRef(part);if(one){var v2=evalCell(one.r,one.c,new Set(seen));var n2=parseFloat(v2);if(!isNaN(n2))vals.push(n2)}else{var n3=parseFloat(part);if(!isNaN(n3))vals.push(n3)}}
    });
    if(fn==='SUM')return vals.reduce(function(a,b){return a+b},0);
    if(fn==='AVERAGE')return vals.length?vals.reduce(function(a,b){return a+b},0)/vals.length:0;
    if(fn==='COUNT')return vals.length;
    if(fn==='MAX')return vals.length?Math.max.apply(null,vals):0;
    if(fn==='MIN')return vals.length?Math.min.apply(null,vals):0;
  }
  /* Bieu thuc so hoc: tham chieu o + - * / ngoac don, khong dung eval. */
  var toks=[];var i=0;var ok=true;
  while(i<expr.length){
    var ch=expr[i];
    if(ch===' '){i++;continue}
    if('+-*/()'.indexOf(ch)>=0){toks.push({t:'op',v:ch});i++;continue}
    var mm=/^[A-J][1-9][0-9]*/i.exec(expr.slice(i));
    if(mm){toks.push({t:'ref',v:mm[0]});i+=mm[0].length;continue}
    var nm=/^\d+(\.\d+)?/.exec(expr.slice(i));
    if(nm){toks.push({t:'num',v:parseFloat(nm[0])});i+=nm[0].length;continue}
    ok=false;break;
  }
  if(!ok||!toks.length)return raw;
  var pos=0;
  function peek(){return toks[pos]}
  function parseE(){
    var v=parseT();
    while(peek()&&(peek().v==='+'||peek().v==='-')){var op=peek().v;pos++;var rhs=parseT();v=(op==='+')?v+rhs:v-rhs;}
    return v;
  }
  function parseT(){
    var v=parseF();
    while(peek()&&(peek().v==='*'||peek().v==='/')){var op=peek().v;pos++;var rhs=parseF();if(op==='/'&&rhs===0)throw new Error('div0');v=(op==='*')?v*rhs:v/rhs;}
    return v;
  }
  function parseF(){
    var tk=peek();if(!tk)throw new Error('empty');
    if(tk.t==='op'&&tk.v==='-'){pos++;return -parseF();}
    if(tk.t==='op'&&tk.v==='('){pos++;var v=parseE();if(!peek()||peek().v!==')')throw new Error('paren');pos++;return v;}
    if(tk.t==='num'){pos++;return tk.v;}
    if(tk.t==='ref'){
      pos++;var rf=parseRef(tk.v);if(!rf)throw new Error('ref');
      var vv=evalCell(rf.r,rf.c,new Set(seen));var nn=parseFloat(vv);
      if(isNaN(nn))throw new Error('notnum');
      return nn;
    }
    throw new Error('tok');
  }
  try{
    var res=parseE();
    if(pos!==toks.length)return raw;
    return res;
  }catch(e){return '#LỖI'}
}
function displayCell(r,c){var v=evalCell(r,c);return (v==null?'':String(v))}
function renderSheet(){
  var html='<tr><th></th>';
  for(var c=0;c<SHEET_COLS;c++)html+='<th>'+colName(c)+'</th>';
  html+='</tr>';
  for(var r=0;r<SHEET_ROWS;r++){
    html+='<tr><th>'+(r+1)+'</th>';
    for(var cc=0;cc<SHEET_COLS;cc++){
      var sel=(selCell.r===r&&selCell.c===cc)?' selected':'';
      html+='<td class="cell'+sel+'" data-r="'+r+'" data-c="'+cc+'"><div class="cellval">'+esc(displayCell(r,cc))+'</div></td>';
    }
    html+='</tr>';
  }
  $('sheetTable').innerHTML=html;
  $('sheetTable').querySelectorAll('td.cell').forEach(function(td){
    td.addEventListener('click',function(){
      var r=+td.dataset.r,c=+td.dataset.c;
      var now=Date.now(),key=r+','+c;
      if(selCell.r===r&&selCell.c===c&&lastTap.key===key&&now-lastTap.t<450){editCell(td,r,c);return}
      selCell={r:r,c:c};lastTap={key:key,t:now};renderSheet();
    });
  });
  updateFormulaBar();
}
function editCell(td,r,c){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var inp=document.createElement('input');inp.value=sheetRaw[r][c]||'';
  if(inp.value.charAt(0)==='=')inp.className='formula';
  td.innerHTML='';td.appendChild(inp);inp.focus();
  var done=false;
  function commit(){if(done)return;done=true;pushSheetHistory();sheetRaw[r][c]=inp.value;renderSheet()}
  inp.addEventListener('blur',commit);
  inp.addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();commit()}});
}
function newSheet(){currentSheetId=null;$('sheetName').value='';sheetRaw=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});sheetUndo=[];sheetRedo=[];renderSheet();updateStatusBar();}
function saveSheet(silent){
  var name=$('sheetName').value.trim()||('Bảng tính '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentSheetId)name=uniqueName('sheet',name);
  var sheets=loadJSON('vb_sheets_v1',[]);
  var rec={id:currentSheetId||uid(),name:name,raw:sheetRaw,updatedAt:Date.now(),syncStatus:'local'};
  if(currentSheetId){var s=sheets.find(function(x){return x.id===currentSheetId});if(s){Object.assign(s,rec)}else sheets.unshift(rec)}
  else{currentSheetId=rec.id;sheets.unshift(rec)}
  saveJSON('vb_sheets_v1',sheets);renderSheetList();renderFileList();if(!silent)alert('Đã lưu bảng tính vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openSheetRec(id){
  var s=loadJSON('vb_sheets_v1',[]).find(function(x){return x.id===id});if(!s)return;
  currentSheetId=s.id;$('sheetName').value=s.name;sheetRaw=normalizeSheetRaw(s.raw);renderSheet();gotoTab('sheet');updateStatusBar();
}
function renderSheetList(){
  var sheets=loadJSON('vb_sheets_v1',[]);
  $('sheetList').innerHTML=sheets.length?sheets.map(function(s){
    return '<div class="fileitem"><b>'+esc(s.name)+'</b><span class="small">'+new Date(s.updatedAt).toLocaleString('vi-VN')+'</span>'
      +'<button onclick="openSheetRec(\''+s.id+'\')">Mở</button><button onclick="deleteSheet(\''+s.id+'\')">Xóa</button></div>';
  }).join(''):'<p class="small">Chưa có bảng tính nào.</p>';
}
function csvEscape(v){v=String(v==null?'':v);return (v.indexOf(',')>=0||v.indexOf('"')>=0||v.indexOf('\n')>=0)?'"'+v.replace(/"/g,'""')+'"':v}
function sheetCsv(){
  var lastR=0,lastC=0;
  for(var r=0;r<SHEET_ROWS;r++)for(var c=0;c<SHEET_COLS;c++)if(sheetRaw[r][c]){lastR=Math.max(lastR,r);lastC=Math.max(lastC,c)}
  var lines=[];
  for(var rr=0;rr<=lastR;rr++){var row=[];for(var cc=0;cc<=lastC;cc++)row.push(csvEscape(displayCell(rr,cc)));lines.push(row.join(','))}
  return lines.join('\n');
}
function exportCsv(){downloadText(sheetCsv(),(($('sheetName').value.trim()||'bang-tinh')+'.csv'),'text/csv;charset=utf-8')}
function normalizeSheetRaw(raw){
  var out=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});
  (raw||[]).forEach(function(row,r){
    if(r>=SHEET_ROWS)return;
    (row||[]).forEach(function(v,c){if(c<SHEET_COLS)out[r][c]=v});
  });
  return out;
}
function xlsxBlobFromRaw(raw){
  var old=sheetRaw;sheetRaw=normalizeSheetRaw(raw);
  try{
    var lastR=0,lastC=0;
    for(var r=0;r<SHEET_ROWS;r++)for(var c=0;c<SHEET_COLS;c++)if(sheetRaw[r][c]){lastR=Math.max(lastR,r);lastC=Math.max(lastC,c)}
    var rows='';
    for(var rr=0;rr<=lastR;rr++){
      rows+='<row r="'+(rr+1)+'">';
      for(var cc=0;cc<=lastC;cc++){
        var v=displayCell(rr,cc);var ref=cellRef(rr,cc);
        if(v===''||v==null)continue;
        var n=parseFloat(v);
        if(!isNaN(n)&&String(v).trim()===String(n))rows+='<c r="'+ref+'"><v>'+n+'</v></c>';
        else rows+='<c r="'+ref+'" t="inlineStr"><is><t xml:space="preserve">'+esc(v)+'</t></is></c>';
      }
      rows+='</row>';
    }
    var dim='A1:'+cellRef(lastR,lastC);
    return zipStored([
      {name:'[Content_Types].xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'},
      {name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'},
      {name:'xl/workbook.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>'},
      {name:'xl/_rels/workbook.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'},
      {name:'xl/styles.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="11"/><name val="Carlito"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'},
      {name:'xl/worksheets/sheet1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="'+dim+'"/><sheetViews><sheetView workbookViewId="0" showGridLines="1"/></sheetViews><sheetData>'+rows+'</sheetData></worksheet>'}
    ]);
  }finally{sheetRaw=old}
}
function exportXlsx(){
  downloadBlob(xlsxBlobFromRaw(sheetRaw),($('sheetName').value.trim()||'bang-tinh')+'.xlsx');
}
/* ---------- PDF: xem trang, chú thích bằng ngón tay, chụm phóng to ---------- */
var pdfDoc=null,pdfFileName='',pdfPageNum=1,pdfScale=1.5,pdfBaseScale=1.5;
var pdfArrayBuffer=null,annotDrawing=null,annotStart=null,pinchState=null;
function pdfAnnots(){return loadJSON('vb_pdf_annots_v1',{})}
function pageAnnots(){var a=pdfAnnots();return (a[pdfFileName]&&a[pdfFileName][pdfPageNum])||{highlights:[],notes:[],drawings:[]}}
function savePageAnnots(pa){var a=pdfAnnots();a[pdfFileName]=a[pdfFileName]||{};a[pdfFileName][pdfPageNum]=pa;saveJSON('vb_pdf_annots_v1',a)}
function loadScript(src){
  return new Promise(function(res,rej){
    if(document.querySelector('script[src="'+src+'"]')){res();return}
    var s=document.createElement('script');s.src=src;s.onload=res;s.onerror=function(){rej(new Error('Không tải được '+src))};
    document.head.appendChild(s);
  });
}
function loadPdfJs(){
  if(window.pdfjsLib)return Promise.resolve();
  return loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js').then(function(){
    if(window.pdfjsLib)window.pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    else throw new Error('pdf.js không khả dụng');
  });
}
function renderPdfPage(){
  if(!pdfDoc)return;
  pdfDoc.getPage(pdfPageNum).then(function(page){
    var vp=page.getViewport({scale:pdfScale});
    var cv=$('pdfCanvas'),ov=$('pdfOverlay');
    cv.width=vp.width;cv.height=vp.height;ov.width=vp.width;ov.height=vp.height;
    ov.style.width=vp.width+'px';ov.style.height=vp.height+'px';
    cv.style.width=vp.width+'px';cv.style.height=vp.height+'px';
    $('pdfWrap').style.display='inline-block';
    page.render({canvasContext:cv.getContext('2d'),viewport:vp}).promise.then(drawAnnots);
    $('pdfPageInfo').textContent='Trang '+pdfPageNum+' / '+pdfDoc.numPages+' — '+pdfFileName;
  }).catch(function(e){$('pdfStatus').textContent='Không vẽ được trang này: '+e.message});
}
function drawAnnots(){
  var ov=$('pdfOverlay'),ctx=ov.getContext('2d');ctx.clearRect(0,0,ov.width,ov.height);
  var pa=pageAnnots();
  pa.highlights.forEach(function(h){
    ctx.fillStyle='rgba(255,235,59,0.45)';
    ctx.fillRect(h.x*ov.width,h.y*ov.height,h.w*ov.width,h.h*ov.height);
  });
  pa.drawings.forEach(function(path){
    ctx.strokeStyle='#c00000';ctx.lineWidth=3;ctx.lineCap='round';ctx.beginPath();
    path.forEach(function(p,i){var x=p.x*ov.width,y=p.y*ov.height;if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y)});
    ctx.stroke();
  });
  pa.notes.forEach(function(n,i){
    ctx.fillStyle='#0b6fa4';ctx.beginPath();ctx.arc(n.x*ov.width,n.y*ov.height,14,0,Math.PI*2);ctx.fill();
    ctx.fillStyle='#fff';ctx.font='bold 16px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText(String(i+1),n.x*ov.width,n.y*ov.height);
  });
}
function relPos(e){
  var r=$('pdfOverlay').getBoundingClientRect();
  var t=(e.touches&&e.touches[0])||e;
  return {x:(t.clientX-r.left)/r.width,y:(t.clientY-r.top)/r.height};
}
$('pdfOverlay').addEventListener('pointerdown',function(e){
  if(e.pointerType==='touch'&&e.isPrimary===false)return;
  var tool=$('pdfTool').value;var p=relPos(e);
  if(tool==='highlight')annotStart=p;
  if(tool==='draw')annotDrawing=[p];
  if(tool==='note'){
    var txt=prompt('Nhập ghi chú cho vị trí này:');
    if(txt){var pa=pageAnnots();pa.notes.push({x:p.x,y:p.y,text:txt});savePageAnnots(pa);drawAnnots();
      $('pdfStatus').textContent='Đã lưu ghi chú số '+(pa.notes.length)+' ở trang '+pdfPageNum+'. Ghi chú: '+txt;}
  }
});
$('pdfOverlay').addEventListener('pointermove',function(e){
  if(!annotDrawing)return;
  annotDrawing.push(relPos(e));
  var ctx=$('pdfOverlay').getContext('2d');drawAnnots();
  ctx.strokeStyle='#c00000';ctx.lineWidth=3;ctx.lineCap='round';ctx.beginPath();
  annotDrawing.forEach(function(p,i){var x=p.x*$('pdfOverlay').width,y=p.y*$('pdfOverlay').height;if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y)});
  ctx.stroke();
});
$('pdfOverlay').addEventListener('pointerup',function(e){
  var tool=$('pdfTool').value;
  if(tool==='highlight'&&annotStart){
    var p=relPos(e);var pa=pageAnnots();
    pa.highlights.push({x:Math.min(annotStart.x,p.x),y:Math.min(annotStart.y,p.y),w:Math.abs(p.x-annotStart.x),h:Math.abs(p.y-annotStart.y)});
    savePageAnnots(pa);annotStart=null;drawAnnots();
    $('pdfStatus').textContent='Đã lưu tô sáng ở trang '+pdfPageNum+' (lưu kèm tên file, không ghi vào ruột PDF gốc).';
  }
  if(tool==='draw'&&annotDrawing){var pa2=pageAnnots();pa2.drawings.push(annotDrawing);savePageAnnots(pa2);annotDrawing=null;drawAnnots()}
});
function clearPageAnnots(){
  if(!pdfFileName)return;
  savePageAnnots({highlights:[],notes:[],drawings:[]});drawAnnots();
  $('pdfStatus').textContent='Đã xóa chú thích trang '+pdfPageNum+'.';
}
/* Chụm hai ngón để phóng to: theo dõi 2 điểm chạm trên khung PDF */
var pdfWrapEl=$('pdfWrap');
pdfWrapEl.addEventListener('touchstart',function(e){
  if(e.touches.length===2){
    pinchState={d:Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY),base:pdfScale};
  }
},{passive:true});
pdfWrapEl.addEventListener('touchmove',function(e){
  if(pinchState&&e.touches.length===2){
    var d=Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY);
    if(pinchState.d>0)pdfScale=Math.min(4,Math.max(0.5,pinchState.base*(d/pinchState.d)));
  }
},{passive:true});
pdfWrapEl.addEventListener('touchend',function(e){
  if(pinchState&&e.touches.length<2){pinchState=null;renderPdfPage();$('pdfStatus').textContent='Độ phóng hiện tại: '+Math.round(pdfScale/pdfBaseScale*100)+'%.'}
},{passive:true});
function openPdfBuffer(buf,name){
  $('pdfStatus').textContent='Đang tải pdf.js (cần mạng lần đầu)…';
  loadPdfJs().then(function(){
    return window.pdfjsLib.getDocument({data:buf}).promise;
  }).then(function(doc){
    pdfDoc=doc;pdfFileName=name;pdfArrayBuffer=buf;pdfPageNum=1;pdfScale=1.5;pdfBaseScale=1.5;
    renderPdfPage();$('pdfStatus').textContent='Đã mở '+name+' ('+doc.numPages+' trang). Chú thích lưu kèm tên file.';
  }).catch(function(e){
    $('pdfStatus').textContent='KHÔNG MỞ ĐƯỢC PDF: không tải được pdf.js (mất mạng hoặc CDN bị chặn). Kiểm tra mạng rồi thử lại — trang không hiện trang trắng khó hiểu. Chi tiết: '+e.message;
    $('pdfWrap').style.display='none';
  });
}
$('pdfFile').addEventListener('change',function(){
  var f=this.files[0];if(!f)return;
  f.arrayBuffer().then(function(buf){openPdfBuffer(buf.slice(0),f.name)});
});
function pdfPrev(){if(pdfDoc&&pdfPageNum>1){pdfPageNum--;renderPdfPage()}}
function pdfNext(){if(pdfDoc&&pdfPageNum<pdfDoc.numPages){pdfPageNum++;renderPdfPage()}}
function openPdfFromList(){
  var gens=loadJSON('vb_generated_pdfs_v1',[]);
  if(!gens.length){alert('Chưa có PDF nào tạo bằng camera. Chụp ảnh rồi bấm “Tạo file PDF”.');return}
  openSheet('Mở PDF đã tạo bằng camera',gens.map(function(g){
    return {label:esc(g.name),onClick:function(){getBlob(g.blobId).then(function(b){
      if(!b){alert('Không tìm thấy file PDF này trong máy (có thể đã xóa dữ liệu trình duyệt). Tạo lại từ ảnh.');return}
      b.arrayBuffer().then(function(buf){openPdfBuffer(buf,g.name)});
    })}};
  }));
}

/* ---------- IndexedDB: giữ Blob PDF đã tạo (localStorage không chứa được file lớn) ---------- */
function idb(){
  return new Promise(function(res,rej){
    var r=indexedDB.open('vb-office',1);
    r.onupgradeneeded=function(){r.result.createObjectStore('blobs')};
    r.onsuccess=function(){res(r.result)};r.onerror=function(){rej(r.error)};
  });
}
function putBlob(id,blob){
  return idb().then(function(db){return new Promise(function(res,rej){
    var tx=db.transaction('blobs','readwrite');tx.objectStore('blobs').put(blob,id);
    tx.oncomplete=function(){res()};tx.onerror=function(){rej(tx.error)};
  })});
}
function getBlob(id){
  return idb().then(function(db){return new Promise(function(res,rej){
    var tx=db.transaction('blobs','readonly');var q=tx.objectStore('blobs').get(id);
    q.onsuccess=function(){res(q.result||null)};q.onerror=function(){rej(q.error)};
  })}).catch(function(){return null});
}

/* ---------- OCR: Tesseract.js (vie+eng) trước, OCR.space dự phòng ---------- */
var ocrImageBlob=null;
function loadTesseract(){
  if(window.Tesseract)return Promise.resolve();
  return loadScript('https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js').then(function(){
    if(!window.Tesseract)throw new Error('Tesseract.js không khả dụng sau khi tải');
  });
}
function runOcr(blob){
  $('ocrStatus').textContent='Đang nhận dạng… (lần đầu cần mạng để tải script và dữ liệu vie+eng, có thể mất một lúc)';
  $('ocrFallback').style.display='none';
  loadTesseract().then(function(){
    return window.Tesseract.recognize(blob,'vie+eng',{logger:function(m){
      if(m.status)$('ocrStatus').textContent='OCR: '+m.status+(m.progress!=null?' — '+Math.round(m.progress*100)+'%':'');
    }});
  }).then(function(res){
    $('ocrResult').value=(res.data&&res.data.text)||'';
    $('ocrStatus').textContent='Xong bằng Tesseract.js (vie+eng). Bùi sửa tay trực tiếp trong ô kết quả nếu nhận dạng sai.';
  }).catch(function(e){
    var key=(loadJSON('vb_settings_v1',{}).ocrSpaceKey||'').trim();
    var msg='Tesseract.js không chạy được trên trình duyệt này ('+e.message+'). Phương án dự phòng: OCR.space qua mạng (language vie). ';
    if(!key){
      msg+='Chưa có key OCR.space miễn phí trong Cài đặt — Bùi tự tạo key miễn phí ở ocr.space rồi nhập vào Cài đặt, sau đó bấm OCR lại. Trang không nhúng key nào sẵn và không bịa kết quả.';
      $('ocrFallback').textContent=msg;$('ocrFallback').style.display='block';$('ocrStatus').textContent='Chưa nhận dạng được — xem phương án dự phòng bên dưới.';
      return;
    }
    $('ocrStatus').textContent='Đang thử OCR.space (cần mạng; bản miễn phí giới hạn số lần/ngày và kích thước file)…';
    var fd=new FormData();fd.append('file',blob,'anh-ocr.jpg');fd.append('apikey',key);fd.append('language','vie');fd.append('isOverlayRequired','false');
    return fetch('https://api.ocr.space/parse/image',{method:'POST',body:fd}).then(function(r){return r.json()}).then(function(j){
      if(j.ParsedResults&&j.ParsedResults[0]&&j.ParsedResults[0].ParsedText!==undefined){
        $('ocrResult').value=j.ParsedResults[0].ParsedText;
        $('ocrStatus').textContent='Xong bằng OCR.space (vie). Bùi sửa tay trực tiếp nếu nhận dạng sai.';
      }else{
        throw new Error((j.ErrorMessage&&j.ErrorMessage.join('; '))||'OCR.space không trả kết quả');
      }
    });
  }).catch(function(e2){
    if(e2&&e2.message)$('ocrStatus').textContent='OCR chưa xong: '+e2.message+' — không có kết quả bịa nào được ghi vào ô kết quả.';
  });
}
function ocrCurrentPage(){
  if(!pdfDoc){alert('Chưa mở file PDF nào. Chọn file PDF ở trên trước.');return}
  pdfDoc.getPage(pdfPageNum).then(function(page){
    var vp=page.getViewport({scale:2});
    var cv=document.createElement('canvas');cv.width=vp.width;cv.height=vp.height;
    return page.render({canvasContext:cv.getContext('2d'),viewport:vp}).promise.then(function(){
      cv.toBlob(function(b){if(b)runOcr(b)},'image/png');
    });
  }).catch(function(e){$('ocrStatus').textContent='Không xuất được trang thành ảnh để OCR: '+e.message});
}
$('ocrImage').addEventListener('change',function(){
  var f=this.files[0];if(!f)return;ocrImageBlob=f;
  var img=$('ocrPreview');img.src=URL.createObjectURL(f);img.style.display='block';
});
function ocrChosenImage(){
  if(!ocrImageBlob){alert('Chưa chọn ảnh nào.');return}
  runOcr(ocrImageBlob);
}
function copyOcr(){var t=$('ocrResult');t.select();try{document.execCommand('copy');$('ocrStatus').textContent='Đã sao chép kết quả OCR.'}catch(e){$('ocrStatus').textContent='Trình duyệt không cho sao chép tự động — Bùi bôi đen và sao chép tay.'}}
function ocrHtml(){
  return $('ocrResult').value.split('\n').map(function(line){return '<p>'+esc(line)+'</p>'}).join('');
}
function insertOcrToDoc(){
  if(!$('ocrResult').value.trim()){alert('Chưa có kết quả OCR.');return}
  $('editor').innerHTML+=ocrHtml();gotoTab('editor');alert('Đã chèn kết quả OCR vào tài liệu đang soạn.');
}
function saveOcrAsDoc(){
  var txt=$('ocrResult').value.trim();if(!txt){alert('Chưa có kết quả OCR.');return}
  var docs=loadJSON('vb_docs_v1',[]);
  var rec={id:uid(),name:'OCR '+(new Date()).toLocaleString('vi-VN'),html:ocrHtml(),updatedAt:Date.now(),syncStatus:'local'};
  docs.unshift(rec);saveJSON('vb_docs_v1',docs);currentDocId=rec.id;$('docName').value=rec.name;$('editor').innerHTML=rec.html;
  renderDocList();renderFileList();alert('Đã lưu kết quả OCR thành tài liệu mới.');
}

/* ---------- Chụp ảnh tạo PDF thật bằng JavaScript thuần (nhúng JPEG) ---------- */
var photos=[],generatedPdfBlob=null,generatedPdfName='';
var camStream=null;
function openCamera(){
  if(!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia){
    alert('Trình duyệt này không hỗ trợ mở camera trực tiếp. Dùng nút “chọn ảnh từ máy” (input capture) bên cạnh — cách ổn định nhất trên iPhone.');
    return;
  }
  navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'},audio:false}).then(function(s){
    camStream=s;var v=$('camVideo');v.srcObject=s;v.style.display='block';
    $('captureBtn').style.display='inline-block';$('closeCamBtn').style.display='inline-block';
  }).catch(function(e){alert('Không mở được camera: '+e.message+' — dùng nút chọn ảnh từ máy bên cạnh.')});
}
function closeCamera(){
  if(camStream)camStream.getTracks().forEach(function(t){t.stop()});
  camStream=null;$('camVideo').style.display='none';$('captureBtn').style.display='none';$('closeCamBtn').style.display='none';
}
function compressToJpeg(source){
  return new Promise(function(res,rej){
    var img=new Image();
    img.onload=function(){
      var max=1600,scale=Math.min(1,max/Math.max(img.naturalWidth,img.naturalHeight));
      var cv=document.createElement('canvas');
      cv.width=Math.round(img.naturalWidth*scale);cv.height=Math.round(img.naturalHeight*scale);
      cv.getContext('2d').drawImage(img,0,0,cv.width,cv.height);
      res(cv.toDataURL('image/jpeg',0.85));
    };
    img.onerror=function(){rej(new Error('Không đọc được ảnh'))};
    if(typeof source==='string')img.src=source;else img.src=URL.createObjectURL(source);
  });
}
function capturePhoto(){
  var v=$('camVideo');if(!v.videoWidth){alert('Camera chưa sẵn sàng, chờ một lúc rồi chụp lại.');return}
  var cv=document.createElement('canvas');cv.width=v.videoWidth;cv.height=v.videoHeight;
  cv.getContext('2d').drawImage(v,0,0);
  compressToJpeg(cv.toDataURL('image/jpeg',0.92)).then(function(d){photos.push(d);renderPhotos()});
}
$('camFiles').addEventListener('change',function(){
  var files=Array.from(this.files||[]);
  if(!files.length)return;
  Promise.all(files.map(function(f){return compressToJpeg(f)})).then(function(ds){
    photos=photos.concat(ds);renderPhotos();
    $('camStatus').textContent='Đã thêm '+ds.length+' ảnh. Tổng '+photos.length+' trang.';
  }).catch(function(e){$('camStatus').textContent='Không đọc được ảnh: '+e.message});
});
function renderPhotos(){
  $('photoList').innerHTML=photos.length?photos.map(function(d,i){
    return '<div class="fileitem"><img src="'+d+'" alt="Trang '+(i+1)+'" style="width:90px;height:auto;border:1px solid #ddd">'
      +'<b>Trang '+(i+1)+'</b>'
      +'<button onclick="movePhoto('+i+',-1)">Lên</button><button onclick="movePhoto('+i+',1)">Xuống</button>'
      +'<button onclick="removePhoto('+i+')">Xóa trang</button></div>';
  }).join(''):'<p class="small">Chưa có ảnh nào. Chụp bằng camera hoặc chọn ảnh từ máy.</p>';
}
function movePhoto(i,dir){
  var j=i+dir;if(j<0||j>=photos.length)return;
  var t=photos[i];photos[i]=photos[j];photos[j]=t;renderPhotos();
}
function removePhoto(i){photos.splice(i,1);renderPhotos()}
function jpegInfo(dataUrl){
  return new Promise(function(res,rej){
    var img=new Image();
    img.onload=function(){
      var bin=atob(dataUrl.split(',')[1]);var bytes=new Uint8Array(bin.length);
      for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
      res({w:img.naturalWidth,h:img.naturalHeight,bytes:bytes});
    };
    img.onerror=function(){rej(new Error('Không đọc được ảnh JPEG'))};
    img.src=dataUrl;
  });
}
function buildPdfFromJpegs(dataUrls){
  return Promise.all(dataUrls.map(jpegInfo)).then(function(imgs){
    var parts=[],offsets=[],pos=0;
    function pushBytes(b){parts.push(b);pos+=b.length}
    function pushStr(s){pushBytes(strBytes(s))}
    var n=imgs.length,totalObjs=2+n*3;
    pushStr('%PDF-1.4\n');
    var kids=imgs.map(function(_,i){return (3+i*3)+' 0 R'}).join(' ');
    offsets[1]=pos;pushStr('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n');
    offsets[2]=pos;pushStr('2 0 obj\n<< /Type /Pages /Kids ['+kids+'] /Count '+n+' >>\nendobj\n');
    imgs.forEach(function(im,i){
      var pageObj=3+i*3,imgObj=4+i*3,contentObj=5+i*3;
      var content='q '+im.w+' 0 0 '+im.h+' 0 0 cm /Im'+i+' Do Q';
      offsets[pageObj]=pos;
      pushStr(pageObj+' 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 '+im.w+' '+im.h+'] /Resources << /XObject << /Im'+i+' '+imgObj+' 0 R >> >> /Contents '+contentObj+' 0 R >>\nendobj\n');
      offsets[imgObj]=pos;
      pushStr(imgObj+' 0 obj\n<< /Type /XObject /Subtype /Image /Width '+im.w+' /Height '+im.h+' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length '+im.bytes.length+' >>\nstream\n');
      pushBytes(im.bytes);pushStr('\nendstream\nendobj\n');
      offsets[contentObj]=pos;
      pushStr(contentObj+' 0 obj\n<< /Length '+content.length+' >>\nstream\n'+content+'\nendstream\nendobj\n');
    });
    var xrefPos=pos;
    var xref='xref\n0 '+(totalObjs+1)+'\n0000000000 65535 f \n';
    for(var k=1;k<=totalObjs;k++)xref+=String(offsets[k]).padStart(10,'0')+' 00000 n \n';
    xref+='trailer\n<< /Size '+(totalObjs+1)+' /Root 1 0 R >>\nstartxref\n'+xrefPos+'\n%%EOF\n';
    pushStr(xref);
    return new Blob(parts,{type:'application/pdf'});
  });
}
function buildPdfFromPhotos(){
  if(!photos.length){alert('Chưa có ảnh nào để tạo PDF.');return}
  $('camStatus').textContent='Đang tạo file PDF thật bằng JavaScript thuần…';
  buildPdfFromJpegs(photos).then(function(blob){
    generatedPdfBlob=blob;
    generatedPdfName='Anh-chup-'+(new Date()).toISOString().slice(0,10)+'-'+Date.now().toString(36)+'.pdf';
    var blobId=uid();
    return putBlob(blobId,blob).then(function(){
      var gens=loadJSON('vb_generated_pdfs_v1',[]);
      gens.unshift({id:uid(),name:generatedPdfName,blobId:blobId,updatedAt:Date.now(),syncStatus:'local'});
      saveJSON('vb_generated_pdfs_v1',gens);renderFileList();
      $('camStatus').textContent='Đã tạo xong PDF thật ('+photos.length+' trang): '+generatedPdfName+' — đã lưu vào danh sách file để đồng bộ, và mở được bằng OCR ở mục 3a.';
      $('dlGenBtn').style.display='inline-block';$('openGenBtn').style.display='inline-block';
    }).catch(function(){
      $('camStatus').textContent='Đã tạo xong PDF nhưng trình duyệt không lưu được file vào IndexedDB — bấm “Tải PDF vừa tạo về máy” để giữ file.';
      $('dlGenBtn').style.display='inline-block';$('openGenBtn').style.display='inline-block';
    });
  }).catch(function(e){$('camStatus').textContent='Tạo PDF chưa xong: '+e.message});
}
function downloadGeneratedPdf(){if(generatedPdfBlob)downloadBlob(generatedPdfBlob,generatedPdfName)}
function openGeneratedPdfForOcr(){
  if(!generatedPdfBlob){alert('Chưa có PDF vừa tạo.');return}
  generatedPdfBlob.arrayBuffer().then(function(buf){gotoTab('pdf');openPdfBuffer(buf,generatedPdfName)});
}

/* ---------- Dịch qua MyMemory (cần mạng, không bịa) ---------- */
function doTranslate(){
  var text=$('transInput').value.trim();
  if(!text){$('transStatus').textContent='Chưa nhập văn bản cần dịch.';return}
  var pair=$('srcLang').value+'|'+$('dstLang').value;
  $('transStatus').textContent='Đang dịch qua MyMemory (cần mạng)…';
  $('transOutput').value='';
  fetch('https://api.mymemory.translated.net/get?q='+encodeURIComponent(text)+'&langpair='+encodeURIComponent(pair))
    .then(function(r){return r.json()}).then(function(j){
      if(j.responseData&&j.responseData.translatedText){
        $('transOutput').value=j.responseData.translatedText;
        $('transStatus').textContent='Dịch xong. Lưu ý: MyMemory là dịch máy miễn phí, nên đọc lại trước khi dùng cho văn bản quan trọng.';
      }else{throw new Error('MyMemory không trả bản dịch')}
    }).catch(function(e){
      $('transStatus').textContent='KHÔNG DỊCH ĐƯỢC (mất mạng hoặc MyMemory lỗi): '+e.message+' — trang không bịa bản dịch. Kiểm tra mạng rồi thử lại.';
    });
}
function translateSelection(){
  var sel=window.getSelection();var text=sel?sel.toString().trim():'';
  if(!text){alert('Bôi đen một đoạn trong tài liệu trước, rồi bấm nút này.');return}
  var pair=$('srcLang').value+'|'+$('dstLang').value;
  fetch('https://api.mymemory.translated.net/get?q='+encodeURIComponent(text)+'&langpair='+encodeURIComponent(pair))
    .then(function(r){return r.json()}).then(function(j){
      var tr=(j.responseData&&j.responseData.translatedText)||'';
      if(!tr)throw new Error('MyMemory không trả bản dịch');
      var range=sel.getRangeAt(0);var node=range.startContainer;
      var block=node.nodeType===1?node:node.parentNode;
      while(block&&block!==$('editor')&&!['P','H1','H2','H3','LI','DIV'].includes(block.tagName))block=block.parentNode;
      var p=document.createElement('p');p.className='ban-dich';p.textContent=tr;
      if(block&&block.parentNode)block.parentNode.insertBefore(p,block.nextSibling);
      else $('editor').appendChild(p);
      alert('Đã chèn bản dịch ngay dưới đoạn gốc.');
    }).catch(function(e){alert('KHÔNG DỊCH ĐƯỢC (mất mạng hoặc MyMemory lỗi): '+e.message+' — không chèn bản dịch bịa.')});
}
function insertTranslationBelow(){
  var tr=$('transOutput').value.trim();if(!tr){alert('Chưa có bản dịch. Bấm “Dịch” trước.');return}
  var p=document.createElement('p');p.className='ban-dich';p.textContent=tr;
  $('editor').appendChild(p);gotoTab('editor');alert('Đã chèn bản dịch vào cuối tài liệu đang soạn.');
}

/* ---------- Quản lý file: nhập, xuất, chia sẻ, trạng thái đồng bộ ---------- */
function syncBadge(s){
  if(s==='drive')return '<span class="badge drive">Đã lên Drive</span>';
  if(s==='onedrive')return '<span class="badge onedrive">Đã lên OneDrive</span>';
  if(s==='pending')return '<span class="badge pending">Chờ đồng bộ</span>';
  return '<span class="badge local">Đã lưu máy</span>';
}
function renderFileList(){
  var docs=loadJSON('vb_docs_v1',[]),sheets=loadJSON('vb_sheets_v1',[]),decks=loadJSON('vb_decks_v1',[]),gens=loadJSON('vb_generated_pdfs_v1',[]);
  var html='';
  docs.forEach(function(d){
    html+='<div class="fileitem"><b>'+esc(d.name)+'</b>'+syncBadge(d.syncStatus)
      +'<button onclick="openDoc(\''+d.id+'\')">Mở</button>'
      +'<button onclick="downloadDocFile(\''+d.id+'\')">Tải .docx</button>'
      +'<button onclick="shareDoc(\''+d.id+'\')">Chia sẻ</button>'
      +'<button onclick="syncFile(\'doc\',\''+d.id+'\',\'drive\')">Lên Drive</button>'
      +'<button onclick="syncFile(\'doc\',\''+d.id+'\',\'onedrive\')">Lên OneDrive</button></div>';
  });
  sheets.forEach(function(s){
    html+='<div class="fileitem"><b>'+esc(s.name)+'</b>'+syncBadge(s.syncStatus)
      +'<button onclick="openSheetRec(\''+s.id+'\')">Mở</button>'
      +'<button onclick="syncFile(\'sheet\',\''+s.id+'\',\'drive\')">Lên Drive</button>'
      +'<button onclick="syncFile(\'sheet\',\''+s.id+'\',\'onedrive\')">Lên OneDrive</button></div>';
  });
  decks.forEach(function(dk){
    html+='<div class="fileitem"><b>'+esc(dk.name)+'</b>'+syncBadge(dk.syncStatus)
      +'<button onclick="openDeck(\''+dk.id+'\')">Mở</button>'
      +'<button onclick="downloadDeckPptx(\''+dk.id+'\')">Tải .pptx</button>'
      +'<button onclick="syncFile(\'deck\',\''+dk.id+'\',\'drive\')">Lên Drive</button>'
      +'<button onclick="syncFile(\'deck\',\''+dk.id+'\',\'onedrive\')">Lên OneDrive</button></div>';
  });
  gens.forEach(function(g){
    html+='<div class="fileitem"><b>'+esc(g.name)+'</b>'+syncBadge(g.syncStatus)
      +'<button onclick="downloadGenFile(\''+g.blobId+'\',\''+esc(g.name)+'\')">Tải về</button><button onclick="deleteGen(\''+g.id+'\')">Xóa</button>'
      +'<button onclick="syncFile(\'pdf\',\''+g.id+'\',\'drive\')">Lên Drive</button>'
      +'<button onclick="syncFile(\'pdf\',\''+g.id+'\',\'onedrive\')">Lên OneDrive</button></div>';
  });
  $('fileList').innerHTML=html||'<p class="small">Chưa có file nào. Soạn, lưu, chụp ảnh tạo PDF hoặc nhập file từ máy.</p>';
}
function downloadDocFile(id){
  var d=loadJSON('vb_docs_v1',[]).find(function(x){return x.id===id});if(!d)return;
  var old=$('editor').innerHTML,oldName=$('docName').value;
  $('editor').innerHTML=d.html||'';$('docName').value=d.name;
  exportDocx();$('editor').innerHTML=old;$('docName').value=oldName;
}
function downloadGenFile(blobId,name){
  getBlob(blobId).then(function(b){if(b)downloadBlob(b,name);else alert('Không tìm thấy file PDF này trong máy.')});
}
function shareDoc(id){
  var d=loadJSON('vb_docs_v1',[]).find(function(x){return x.id===id});if(!d)return;
  var tmp=document.createElement('div');tmp.innerHTML=d.html||'';
  var file=new File([tmp.innerText||''],d.name+'.txt',{type:'text/plain'});
  if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
    navigator.share({files:[file],title:d.name}).catch(function(e){
      if(e&&e.name==='AbortError')return;
      downloadText(tmp.innerText||'',d.name+'.txt');
      alert('Chia sẻ chưa xong ('+(e&&e.message?e.message:'lỗi không rõ')+') — file đã tải về máy để Bùi chia sẻ tay.');
    });
  }else{
    downloadText(tmp.innerText||'',d.name+'.txt');
    alert('Trình duyệt này không hỗ trợ Web Share với file — file đã tải về máy để Bùi chia sẻ tay.');
  }
}
/* Giải nén docx bằng JS thuần: entry stored thì lấy thẳng, entry deflate thì dùng DecompressionStream nếu có */
function unzipInflate(data){
  if(typeof DecompressionStream==='undefined')return Promise.reject(new Error('Trình duyệt này không giải nén được docx nén (cần DecompressionStream).'));
  return new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer()
    .then(function(ab){return new TextDecoder().decode(ab)});
}
function unzipDocx(buf){
  return new Promise(function(resolve,reject){
    var bytes=new Uint8Array(buf);var dv=new DataView(buf);
    function readStr(o,len){var str='';for(var i=0;i<len;i++)str+=String.fromCharCode(bytes[o+i]);return str}
    function useEntries(entries){
      var target=entries.find(function(e){return e.name==='word/document.xml'});
      if(!target){reject(new Error('Không thấy word/document.xml trong file docx'));return}
      if(target.method===0){resolve(new TextDecoder().decode(target.data));return}
      if(target.method===8){unzipInflate(target.data).then(resolve).catch(reject);return}
      reject(new Error('Không giải nén được entry docx với phương pháp '+target.method));
    }
    function centralPath(){
      var eocd=-1;
      for(var i=bytes.length-22;i>=0;i--){if(dv.getUint32(i,true)===0x06054b50){eocd=i;break}}
      if(eocd<0)throw new Error('Không tìm thấy thư mục trung tâm (central directory) của file zip');
      var total=dv.getUint16(eocd+10,true);var cdOff=dv.getUint32(eocd+16,true);
      var entries=[];var p=cdOff;
      for(var n=0;n<total;n++){
        if(dv.getUint32(p,true)!==0x02014b50)break;
        var method=dv.getUint16(p+10,true),csize=dv.getUint32(p+20,true);
        var nlen=dv.getUint16(p+28,true),elen=dv.getUint16(p+30,true),clen=dv.getUint16(p+32,true);
        var lho=dv.getUint32(p+42,true);var name=readStr(p+46,nlen);
        var ln=dv.getUint16(lho+26,true),le=dv.getUint16(lho+28,true);
        var start=lho+30+ln+le;
        entries.push({name:name,method:method,data:bytes.slice(start,start+csize)});
        p+=46+nlen+elen+clen;
      }
      useEntries(entries);
    }
    try{
      var entries=[];var off=0;
      while(off+30<=bytes.length&&dv.getUint32(off,true)===0x04034b50){
        var method=dv.getUint16(off+8,true),csize=dv.getUint32(off+18,true);
        var nlen=dv.getUint16(off+26,true),elen=dv.getUint16(off+28,true);
        var name=readStr(off+30,nlen);var start=off+30+nlen+elen;
        entries.push({name:name,method:method,data:bytes.slice(start,start+csize)});
        if(csize===0)break;
        off=start+csize;
      }
      var target=entries.find(function(e){return e.name==='word/document.xml'});
      if(target&&(target.method===0||target.data.length>0))useEntries(entries);
      else centralPath();
    }catch(e){
      try{centralPath()}catch(e2){reject(e2)}
    }
  });
}
$('importFile').addEventListener('change',function(){
  var f=this.files[0];if(!f)return;
  $('importStatus').textContent='Đang nhập '+f.name+'…';
  var lower=f.name.toLowerCase();
  if(lower.endsWith('.txt')){
    f.text().then(function(t){
      var html=t.split('\n').map(function(l){return '<p>'+esc(l)+'</p>'}).join('');
      addImportedDoc(f.name.replace(/\.txt$/i,''),html);
    });
  }else if(lower.endsWith('.csv')){
    f.text().then(function(t){
      var rows=parseCsv(t).filter(function(rw){return rw.some(function(c){return String(c).trim()!==''})});
      newSheet();
      rows.forEach(function(rw,r){
        if(r>=SHEET_ROWS)return;
        rw.forEach(function(cell,c){if(c<SHEET_COLS)sheetRaw[r][c]=cell});
      });
      $('sheetName').value=f.name.replace(/\.csv$/i,'');renderSheet();gotoTab('sheet');
      $('importStatus').textContent='Đã nhập CSV thành bảng tính. Bấm “Lưu vào máy” để giữ lại.';
    });
  }else if(lower.endsWith('.docx')){
    f.arrayBuffer().then(unzipDocx).then(function(xml){
      var paras=xml.split('</w:p>').map(function(p){
        var texts=(p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g)||[]).map(function(m){return m.replace(/<[^>]+>/g,'')}).join('');
        return '<p>'+esc(texts)+'</p>';
      }).join('');
      addImportedDoc(f.name.replace(/\.docx$/i,''),paras);
    }).catch(function(e){
      $('importStatus').textContent='Nhập docx chưa xong: '+e.message+' (giới hạn: trang chỉ đọc được văn bản thuần của docx, không giữ nguyên định dạng phức tạp).';
    });
  }else if(lower.endsWith('.pdf')){
    $('importStatus').textContent='File PDF không lưu vào danh sách văn bản. Sang mục “PDF & OCR & Chụp ảnh”, chọn lại file này để xem, chú thích và OCR.';
    gotoTab('pdf');
  }else{$('importStatus').textContent='Định dạng chưa hỗ trợ nhập: chỉ txt, docx (văn bản thuần), csv, pdf.'}
});
function addImportedDoc(name,html){
  var docs=loadJSON('vb_docs_v1',[]);
  var rec={id:uid(),name:name,html:html,updatedAt:Date.now(),syncStatus:'local'};
  docs.unshift(rec);saveJSON('vb_docs_v1',docs);currentDocId=rec.id;$('docName').value=name;$('editor').innerHTML=html;
  renderDocList();renderFileList();$('importStatus').textContent='Đã nhập: '+name+'.';gotoTab('editor');
}

/* ---------- Đăng nhập & đồng bộ: NGAY TRONG TRANG NÀY ---------- */
var googleToken=null,msToken=null,msAccount=null;
function settings(){return loadJSON('vb_settings_v1',{})}
function saveSettings(){
  var s={googleClientId:$('googleClientId').value.trim(),msClientId:$('msClientId').value.trim(),ocrSpaceKey:$('ocrSpaceKey').value.trim()};
  saveJSON('vb_settings_v1',s);$('settingsStatus').textContent='Đã lưu cài đặt vào máy này.';
  updateAuthBar();
}
function isFileProto(){return location.protocol==='file:'}
function fileProtoNote(){
  return isFileProto()
    ?'⚠️ Trang đang mở bằng file:// trên Safari. Luồng OAuth CHỈ chạy trọn vẹn khi trang được phục vụ qua https với redirect URI đã đăng ký — mở file:// thì nút đăng nhập sẽ báo rõ giới hạn này chứ không hiện lỗi khó hiểu. Cách chạy thật: đưa thư mục web/ lên một host https (ví dụ GitHub Pages) rồi mở link https đó.'
    :'Trang đang chạy qua '+location.protocol+'. OAuth chỉ chạy trọn vẹn khi redirect URI đã đăng ký khớp đúng địa chỉ trang hiện tại.';
}
function updateAuthBar(){
  var g=googleToken?'Google: đã đăng nhập':'Google: chưa đăng nhập';
  var m=msToken?'Microsoft: đã đăng nhập':'Microsoft: chưa đăng nhập';
  $('authBar').textContent=g+' • '+m+' — vào mục Đăng nhập & Đồng bộ để đăng nhập ngay trong trang này.';
  var n=$('authNote');if(n)n.textContent=fileProtoNote();
  var st=$('authStatus');if(st)st.textContent=g+' • '+m;
}
function loginGoogle(){
  if(isFileProto()){alert(fileProtoNote());return}
  var cid=settings().googleClientId;
  if(!cid){alert('Chưa nhập Google Client ID. Vào Cài đặt nhập Client ID trước (xem hướng dẫn ở mục Đăng nhập & Đồng bộ), rồi bấm lại. Trang không bịa Client ID nào.');gotoTab('settings');return}
  loadScript('https://accounts.google.com/gsi/client').then(function(){
    if(!window.google||!google.accounts||!google.accounts.oauth2)throw new Error('Google Identity Services không tải được (mất mạng?)');
    var client=google.accounts.oauth2.initTokenClient({
      client_id:cid,scope:'https://www.googleapis.com/auth/drive.file',
      callback:function(resp){
        if(resp.access_token){googleToken=resp.access_token;sessionStorage.setItem('vb_google_token',resp.access_token);updateAuthBar();alert('Đã đăng nhập Google ngay trong trang này. Giờ bấm “Lên Drive” ở danh sách file.')}
        else alert('Google không trả access token: '+(resp.error||'không rõ lỗi'));
      }
    });
    client.requestAccessToken();
  }).catch(function(e){alert('Đăng nhập Google chưa xong (cần mạng + Client ID đúng + redirect URI đã đăng ký): '+e.message)});
}
function loginMicrosoft(){
  if(isFileProto()){alert(fileProtoNote());return}
  var cid=settings().msClientId;
  if(!cid){alert('Chưa nhập Microsoft Client ID. Vào Cài đặt nhập Client ID trước (xem hướng dẫn ở mục Đăng nhập & Đồng bộ), rồi bấm lại. Trang không bịa Client ID nào.');gotoTab('settings');return}
  loadScript('https://alcdn.msauth.net/browser/2.38.0/msal-browser.min.js').then(function(){
    if(!window.msal)throw new Error('MSAL.js không tải được (mất mạng?)');
    var app=new window.msal.PublicClientApplication({auth:{clientId:cid,redirectUri:location.href.split('#')[0]}});
    window._msalApp=app;
    return app.handleRedirectPromise().then(function(){
      return app.loginPopup({scopes:['User.Read','Files.ReadWrite']});
    }).then(function(res){
      msAccount=res.account;msToken=res.accessToken;updateAuthBar();
      alert('Đã đăng nhập Microsoft ngay trong trang này. Giờ bấm “Lên OneDrive” ở danh sách file.');
    });
  }).catch(function(e){alert('Đăng nhập Microsoft chưa xong (cần mạng + Client ID đúng + redirect URI đã đăng ký): '+e.message)});
}
function logoutAll(){
  var acct=msAccount;
  googleToken=null;msToken=null;msAccount=null;
  sessionStorage.removeItem('vb_google_token');
  var app=window._msalApp;
  if(app&&acct){
    try{
      if(app.logoutPopup)app.logoutPopup({account:acct}).catch(function(){});
      else if(app.logoutRedirect)app.logoutRedirect({account:acct});
    }catch(e){/* da xoa token trong trang; dang xuat MSAL la buoc them */}
  }
  updateAuthBar();alert('Đã đăng xuất trong trang này (token chỉ giữ trong phiên).');
}
function setSyncStatus(kind,id,status){
  var key=kind==='doc'?'vb_docs_v1':(kind==='sheet'?'vb_sheets_v1':(kind==='deck'?'vb_decks_v1':'vb_generated_pdfs_v1'));
  var arr=loadJSON(key,[]);var it=arr.find(function(x){return x.id===id});if(it){it.syncStatus=status;saveJSON(key,arr)}
  renderDocList();renderSheetList();renderFileList();
}
function fileBlobForSync(kind,id){
  if(kind==='deck'){
    var dk=loadJSON('vb_decks_v1',[]).find(function(x){return x.id===id});if(!dk)return null;
    return {name:dk.name+'.pptx',blob:buildPptxBlob(dk)};
  }
  if(kind==='doc'){
    var d=loadJSON('vb_docs_v1',[]).find(function(x){return x.id===id});if(!d)return null;
    var tmp=document.createElement('div');tmp.innerHTML=d.html||'';
    return {name:d.name+'.docx',blob:buildDocxBlobFromNode(tmp)};
  }
  if(kind==='sheet'){
    var s=loadJSON('vb_sheets_v1',[]).find(function(x){return x.id===id});if(!s)return null;
    return {name:s.name+'.xlsx',blob:xlsxBlobFromRaw(s.raw||[])};
  }
  var g=loadJSON('vb_generated_pdfs_v1',[]).find(function(x){return x.id===id});if(!g)return null;
  return getBlob(g.blobId).then(function(b){return b?{name:g.name,blob:b}:null});
}
function syncFile(kind,id,target){
  var needLogin=(target==='drive'&&!googleToken)||(target==='onedrive'&&!msToken);
  if(needLogin){
    setSyncStatus(kind,id,'pending');
    alert('Chưa đăng nhập '+(target==='drive'?'Google':'Microsoft')+'. File đã chuyển sang “Chờ đồng bộ” — đăng nhập ngay trong mục Đăng nhập & Đồng bộ rồi bấm lại nút này.');
    gotoTab('auth');return;
  }
  var fb=fileBlobForSync(kind,id);
  Promise.resolve(fb).then(function(f){
    if(!f){alert('Không tìm thấy file để đồng bộ.');return}
    if(target==='drive'){
      var fd=new FormData();
      fd.append('metadata',new Blob([JSON.stringify({name:f.name})],{type:'application/json'}));
      fd.append('file',f.blob);
      return fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',{method:'POST',headers:{Authorization:'Bearer '+googleToken},body:fd})
        .then(function(r){if(!r.ok)throw new Error('Drive trả mã '+r.status);return r.json()})
        .then(function(){setSyncStatus(kind,id,'drive');alert('Đã lên Drive thật: '+f.name)});
    }else{
      return fetch('https://graph.microsoft.com/v1.0/me/drive/root:/'+encodeURIComponent(f.name)+':/content',{method:'PUT',headers:{Authorization:'Bearer '+msToken},body:f.blob})
        .then(function(r){if(!r.ok)throw new Error('OneDrive trả mã '+r.status);return r.json()})
        .then(function(){setSyncStatus(kind,id,'onedrive');alert('Đã lên OneDrive thật: '+f.name)});
    }
  }).catch(function(e){
    setSyncStatus(kind,id,'pending');
    alert('Đồng bộ chưa thành công, file giữ trạng thái “Chờ đồng bộ”: '+e.message);
  });
}
function syncCurrentDoc(target){
  if(!currentDocId){alert('Chưa chọn tài liệu nào. Mở hoặc lưu một tài liệu trước.');return}
  syncFile('doc',currentDocId,target);
}
function listDriveFiles(){
  if(!googleToken){alert('Chưa đăng nhập Google. Vào mục Đăng nhập & Đồng bộ đăng nhập ngay trong trang này trước.');gotoTab('auth');return}
  fetch('https://www.googleapis.com/drive/v3/files?pageSize=20&fields=files(id,name)',{headers:{Authorization:'Bearer '+googleToken}})
    .then(function(r){if(!r.ok)throw new Error('Drive trả mã '+r.status);return r.json()})
    .then(function(j){
      var files=j.files||[];
      $('cloudList').innerHTML=files.length?files.map(function(f){
        return '<div class="fileitem"><b>'+esc(f.name)+'</b><button onclick="downloadDriveFile(\''+f.id+'\',\''+esc(f.name)+'\')">Tải về mở lại</button></div>';
      }).join(''):'<p class="small">Drive chưa có file nào.</p>';
    }).catch(function(e){$('cloudList').textContent='Không xem được Drive: '+e.message});
}
function downloadDriveFile(id,name){
  fetch('https://www.googleapis.com/drive/v3/files/'+id+'?alt=media',{headers:{Authorization:'Bearer '+googleToken}})
    .then(function(r){if(!r.ok)throw new Error('Drive trả mã '+r.status);return r.text()})
    .then(function(t){addImportedDoc(name.replace(/\.[^.]+$/,''),t.split('\n').map(function(l){return '<p>'+esc(l)+'</p>'}).join(''))})
    .catch(function(e){alert('Tải từ Drive chưa xong: '+e.message)});
}
function listOneDriveFiles(){
  if(!msToken){alert('Chưa đăng nhập Microsoft. Vào mục Đăng nhập & Đồng bộ đăng nhập ngay trong trang này trước.');gotoTab('auth');return}
  fetch('https://graph.microsoft.com/v1.0/me/drive/root/children',{headers:{Authorization:'Bearer '+msToken}})
    .then(function(r){if(!r.ok)throw new Error('OneDrive trả mã '+r.status);return r.json()})
    .then(function(j){
      cloudFiles=j.value||[];
      var files=cloudFiles;
      $('cloudList').innerHTML=files.length?files.map(function(f,i){
        return '<div class="fileitem"><b>'+esc(f.name)+'</b><button onclick="downloadOneDriveFileByIndex('+i+')">Tải về mở lại</button></div>';
      }).join(''):'<p class="small">OneDrive chưa có file nào.</p>';
    }).catch(function(e){$('cloudList').textContent='Không xem được OneDrive: '+e.message});
}
var cloudFiles=[];
function downloadOneDriveFileByIndex(i){
  var f=cloudFiles[i];if(!f){alert('Không tìm thấy file OneDrive này.');return}
  downloadOneDriveFile(f['@microsoft.graph.downloadUrl']||'',f.name);
}
function downloadOneDriveFile(url,name){
  if(!url){alert('OneDrive không trả link tải cho file này.');return}
  fetch(url).then(function(r){return r.text()}).then(function(t){
    addImportedDoc(name.replace(/\.[^.]+$/,''),t.split('\n').map(function(l){return '<p>'+esc(l)+'</p>'}).join(''));
  }).catch(function(e){alert('Tải từ OneDrive chưa xong: '+e.message)});
}

/* ---------- Trình chiếu (bản 1.1): tạo/sửa/xóa slide, chèn ảnh, trình chiếu, xuất .pptx ----------
   Lưu localStorage 'vb_decks_v1' như tài liệu/bảng tính. Ảnh nén tối đa 1280px (dùng lại
   compressToJpeg của mục chụp ảnh) để vừa bộ nhớ trình duyệt. Font trình chiếu: Carlito
   (1 trong 5 font đã kiểm chứng) + dự phòng hệ thống; không thêm font mới. */
var currentDeckId=null;
var deckSlides=[{title:'',bullets:[''],image:null,imageW:0,imageH:0}];
function emptySlide(){return {title:'',bullets:[''],image:null,imageW:0,imageH:0}}
function slideBullets(s){return (s.bullets||[]).filter(function(b){return String(b).trim()!==''})}
function renderSlideEditor(){
  var wrap=$('slideEditor');if(!wrap)return;
  if(!deckSlides.length)deckSlides=[emptySlide()];
  if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;
  if(currentSlide<0)currentSlide=0;
  var thumbs=deckSlides.map(function(s,i){
    var t=s.title||'(Chưa có tiêu đề)';var b=slideBullets(s)[0]||'';
    return '<button class="thumb'+(i===currentSlide?' active':'')+'" onclick="selectSlide('+i+')"><span class="tnum">Slide '+(i+1)+'</span><span class="ttitle">'+esc(t)+'</span><span class="tbody">'+esc(b)+'</span></button>';
  }).join('');
  var s=deckSlides[currentSlide];
  var imgHtml=s.image
    ?'<img class="slideimg" src="'+s.image+'" alt="Ảnh slide"><div class="row"><button onclick="removeSlideImage('+currentSlide+')">Xóa ảnh này</button></div>'
    :'<div class="row"><button onclick="pickSlideImage('+currentSlide+')">Chèn ảnh vào slide này</button></div>';
  var dis=deckLocked?' disabled':'';
  var transLabel={none:'Không',fade:'Mờ dần',slide:'Trượt ngang'}[deckTransition]||'Không';
  wrap.innerHTML='<div id="slideThumbs">'+thumbs+'</div><div id="slideMain"><div class="slideframe'+(deckSize==='4:3'?' ratio-43':'')+'" style="background:'+(s.bg||'#ffffff')+'">'
    +'<label>Tiêu đề slide '+(currentSlide+1)+':</label>'
    +'<input type="text" value="'+esc(s.title||'')+'" oninput="updateSlideTitle('+currentSlide+',this.value)" placeholder="Tiêu đề slide" aria-label="Tiêu đề slide"'+dis+'>'
    +'<label>Nội dung — mỗi dòng một ý gạch đầu dòng:</label>'
    +'<textarea oninput="updateSlideBullets('+currentSlide+',this.value)" placeholder="Ý thứ nhất&#10;Ý thứ hai" aria-label="Nội dung slide"'+dis+'>'+esc((s.bullets||[]).join('\n'))+'</textarea>'
    +imgHtml
    +'</div><p class="small">Slide '+(currentSlide+1)+' / '+deckSlides.length+' • Hiệu ứng chuyển: '+transLabel+(deckLocked?' • Đang bảo vệ (khóa sửa)':'')+'</p>'
    +'<div class="row"><button onclick="selectSlide('+(currentSlide-1)+')">Slide trước</button><button onclick="selectSlide('+(currentSlide+1)+')">Slide sau</button><button onclick="addSlide()">Thêm slide</button><button onclick="presentFromCurrent()">Trình chiếu từ slide này</button></div></div>';
}
function updateSlideTitle(i,v){if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}if(deckSlides[i])deckSlides[i].title=v;updateStatusBar()}
function updateSlideBullets(i,v){if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}if(deckSlides[i])deckSlides[i].bullets=String(v).split('\n')}
function addSlide(){pushDeckHistory();deckSlides.push(emptySlide());currentSlide=deckSlides.length-1;renderSlideEditor();updateStatusBar()}
function moveSlide(i,dir){
  var j=i+dir;if(j<0||j>=deckSlides.length)return;
  pushDeckHistory();
  var t=deckSlides[i];deckSlides[i]=deckSlides[j];deckSlides[j]=t;
  if(currentSlide===i)currentSlide=j;else if(currentSlide===j)currentSlide=i;
  renderSlideEditor();updateStatusBar();
}
function deleteSlide(i){
  if(deckSlides.length<=1){alert('Bài trình bày cần ít nhất 1 slide — không xóa slide cuối cùng này. Muốn làm lại thì bấm “Bài mới”.');return}
  if(!confirm('Xóa slide '+(i+1)+' này?'))return;
  pushDeckHistory();deckSlides.splice(i,1);if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;renderSlideEditor();updateStatusBar();
}
function newDeck(){currentDeckId=null;$('deckName').value='';deckSlides=[emptySlide()];currentSlide=0;deckSize='16:9';deckTransition='none';renderSlideEditor();updateStatusBar()}
function saveDeck(silent){
  var name=$('deckName').value.trim()||('Bài trình bày '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentDeckId)name=uniqueName('deck',name);
  var slides=deckSlides.map(function(s){return {title:s.title||'',bullets:s.bullets||[],image:s.image||null,imageW:s.imageW||0,imageH:s.imageH||0,bg:s.bg||''}});
  var decks=loadJSON('vb_decks_v1',[]);
  var rec={id:currentDeckId||uid(),name:name,slides:slides,updatedAt:Date.now(),syncStatus:'local',size:deckSize,transition:deckTransition};
  if(currentDeckId){var d=decks.find(function(x){return x.id===currentDeckId});if(d){Object.assign(d,rec)}else decks.unshift(rec)}
  else{currentDeckId=rec.id;decks.unshift(rec)}
  saveJSON('vb_decks_v1',decks);renderDeckList();renderFileList();if(!silent)alert('Đã lưu bài trình bày vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openDeck(id){
  var d=loadJSON('vb_decks_v1',[]).find(function(x){return x.id===id});if(!d)return;
  currentDeckId=d.id;$('deckName').value=d.name;
  deckSlides=(d.slides&&d.slides.length)?d.slides.map(function(s){return {title:s.title||'',bullets:s.bullets||[],image:s.image||null,imageW:s.imageW||0,imageH:s.imageH||0,bg:s.bg||''}}):[emptySlide()];
  deckSize=d.size||'16:9';deckTransition=d.transition||'none';currentSlide=0;
  renderSlideEditor();gotoTab('slides');
}
function deleteDeck(id){
  if(!confirm('Xóa bài trình bày này khỏi máy?'))return;
  saveJSON('vb_decks_v1',loadJSON('vb_decks_v1',[]).filter(function(x){return x.id!==id}));
  if(currentDeckId===id)newDeck();renderDeckList();renderFileList();
}
function renderDeckList(){
  var decks=loadJSON('vb_decks_v1',[]);
  $('deckList').innerHTML=decks.length?decks.map(function(d){
    return '<div class="fileitem"><b>'+esc(d.name)+'</b><span class="small">'+(d.slides?d.slides.length:0)+' slide • '+new Date(d.updatedAt).toLocaleString('vi-VN')+'</span>'
      +'<button onclick="openDeck(\''+d.id+'\')">Mở</button>'
      +'<button onclick="downloadDeckPptx(\''+d.id+'\')">Tải .pptx</button>'
      +'<button onclick="deleteDeck(\''+d.id+'\')">Xóa</button></div>';
  }).join(''):'<p class="small">Chưa có bài trình bày nào. Thêm slide rồi bấm “Lưu vào máy”.</p>';
}
/* Chèn ảnh vào slide: nén bằng compressToJpeg (có sẵn), ghi thêm kích thước để xuất .pptx đúng tỉ lệ */
var pendingImageSlide=-1;
function pickSlideImage(i){
  pendingImageSlide=i;
  var inp=$('slideImageInput');inp.value='';inp.click();
}
function removeSlideImage(i){
  if(!deckSlides[i])return;
  pushDeckHistory();
  deckSlides[i].image=null;deckSlides[i].imageW=0;deckSlides[i].imageH=0;renderSlideEditor();
}
$('slideImageInput').addEventListener('change',function(){
  var f=this.files[0];var idx=pendingImageSlide;pendingImageSlide=-1;
  if(!f||idx<0||!deckSlides[idx])return;
  compressToJpeg(f).then(function(dataUrl){
    var img=new Image();
    img.onload=function(){
      deckSlides[idx].image=dataUrl;deckSlides[idx].imageW=img.naturalWidth;deckSlides[idx].imageH=img.naturalHeight;
      renderSlideEditor();
    };
    img.onerror=function(){deckSlides[idx].image=dataUrl;renderSlideEditor()};
    img.src=dataUrl;
  }).catch(function(e){alert('Không đọc được ảnh: '+e.message)});
});

/* ---------- Trình chiếu toàn màn hình: vuốt + nút lớn ---------- */
var presentSlides=[],presentIndex=0,presentTouchX=null;
function presentDeck(){
  if(!deckSlides.length)deckSlides=[emptySlide()];
  presentSlides=JSON.parse(JSON.stringify(deckSlides));
  presentIndex=0;
  $('presentOverlay').style.display='flex';
  renderPresentSlide();
  var ov=$('presentOverlay');
  if(ov.requestFullscreen){ov.requestFullscreen().catch(function(){/* không vào được fullscreen thì vẫn xem bằng lớp phủ toàn màn hình */});}
}
function renderPresentSlide(){
  var s=presentSlides[presentIndex]||emptySlide();
  var bullets=slideBullets(s);
  var html='<h2>'+esc(s.title||'(Chưa có tiêu đề)')+'</h2>';
  if(bullets.length)html+='<ul>'+bullets.map(function(b){return '<li>'+esc(b)+'</li>'}).join('')+'</ul>';
  if(s.image)html+='<img class="presentimg" src="'+s.image+'" alt="Ảnh slide '+(presentIndex+1)+'">';
  if(!bullets.length&&!s.image)html+='<p style="font-size:21px;opacity:.75">Slide này chưa có nội dung. Bấm “✕ Thoát” để về soạn thêm.</p>';
  var ps=$('presentSlide');
  ps.innerHTML=html;
  ps.style.background=(s.bg||'#111111');
  ps.style.color=(s.bg&&s.bg!=='#ffffff')?'#1a1a1a':'#ffffff';
  ps.className=(deckTransition==='fade')?'trans-fade':(deckTransition==='slide'?'trans-slide':'');
  $('presentCounter').textContent='Slide '+(presentIndex+1)+' / '+presentSlides.length;
}
function presentNext(){if(presentIndex<presentSlides.length-1){presentIndex++;renderPresentSlide()}}
function presentPrev(){if(presentIndex>0){presentIndex--;renderPresentSlide()}}
function exitPresent(){
  $('presentOverlay').style.display='none';
  if(document.fullscreenElement&&document.exitFullscreen)document.exitFullscreen().catch(function(){});
}
(function(){
  var ov=document.getElementById('presentOverlay');if(!ov)return;
  ov.addEventListener('touchstart',function(e){
    if(e.touches.length===1)presentTouchX=e.touches[0].clientX;
  },{passive:true});
  ov.addEventListener('touchend',function(e){
    if(presentTouchX==null)return;
    var dx=e.changedTouches[0].clientX-presentTouchX;presentTouchX=null;
    if(dx<-60)presentNext();
    else if(dx>60)presentPrev();
  },{passive:true});
  document.addEventListener('keydown',function(e){
    if($('presentOverlay').style.display==='none')return;
    if(e.key==='ArrowRight')presentNext();
    else if(e.key==='ArrowLeft')presentPrev();
    else if(e.key==='Escape')exitPresent();
  });
})();

/* ---------- Xuất .pptx tối giản tự viết (zip stored + xml, JS thuần, không thư viện ngoài) ----------
   Cấu trúc tối thiểu PowerPoint/LibreOffice/WPS mở được: presentation + slide + layout + master + theme.
   Mỗi slide: 1 hộp tiêu đề, 1 hộp gạch đầu dòng, và 1 ảnh (nếu có) giữ đúng tỉ lệ. Font ghi
   "Carlito" (font mặc định, metric tương đương Calibri, đã kiểm chứng tiếng Việt); máy mở file không có font này sẽ tự dùng font dự phòng. */
function dataUrlToBytes(dataUrl){
  var bin=atob(dataUrl.split(',')[1]);var bytes=new Uint8Array(bin.length);
  for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
  return bytes;
}
function fitBox(maxW,maxH,iw,ih){
  if(!iw||!ih)return {w:maxW,h:maxH};
  var s=Math.min(maxW/iw,maxH/ih);
  return {w:Math.round(iw*s),h:Math.round(ih*s)};
}
function pptxSlideXml(slide,imgRelId){
  var bullets=slideBullets(slide);
  var bgXml=(slide.bg&&slide.bg!=='#ffffff')?'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="'+slide.bg.replace('#','')+'"/></a:solidFill></p:bgPr></p:bg>':'';
  var title='<p:sp><p:nvSpPr><p:cNvPr id="2" name="Tiêu đề"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr></p:nvSpPr>'
    +'<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    +'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="vi-VN" sz="3200" b="1"><a:latin typeface="Carlito"/></a:rPr><a:t>'+esc(slide.title||'')+'</a:t></a:r></a:p></p:txBody></p:sp>';
  var body='';
  var hasImg=!!slide.image;
  if(bullets.length){
    var bw=hasImg?4572000:8229600;
    var paras=bullets.map(function(b){
      return '<a:p><a:pPr marL="228600" indent="-228600"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>'
        +'<a:r><a:rPr lang="vi-VN" sz="2400"><a:latin typeface="Carlito"/></a:rPr><a:t>'+esc(b)+'</a:t></a:r></a:p>';
    }).join('');
    body='<p:sp><p:nvSpPr><p:cNvPr id="3" name="Nội dung"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr></p:nvSpPr>'
      +'<p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="'+bw+'" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
      +'<p:txBody><a:bodyPr/><a:lstStyle/>'+paras+'</p:txBody></p:sp>';
  }
  var pic='';
  if(hasImg){
    var box=bullets.length
      ?fitBox(3657600,4572000,slide.imageW,slide.imageH)
      :fitBox(8229600,4572000,slide.imageW,slide.imageH);
    var px=bullets.length?5486400:Math.round(457200+(8229600-box.w)/2);
    var py=1371600+Math.round((4572000-box.h)/2);
    pic='<p:pic><p:nvPicPr><p:cNvPr id="4" name="Ảnh"/><p:cNvPicPr/></p:nvPicPr>'
      +'<p:blipFill><a:blip r:embed="'+imgRelId+'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
      +'<p:spPr><a:xfrm><a:off x="'+px+'" y="'+py+'"/><a:ext cx="'+box.w+'" cy="'+box.h+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    +'<p:cSld>'+bgXml+'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
    +'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>'
    +title+body+pic+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}
function buildPptxBlob(deck){
  var slides=(deck.slides&&deck.slides.length)?deck.slides:[emptySlide()];
  var entries=[];
  var types='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    +'<Default Extension="xml" ContentType="application/xml"/>'
    +'<Default Extension="jpg" ContentType="image/jpeg"/>'
    +'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
  slides.forEach(function(_,i){types+='<Override PartName="/ppt/slides/slide'+(i+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';});
  types+='<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    +'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    +'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>';
  entries.push({name:'[Content_Types].xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+types});
  entries.push({name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>'});
  var presRels='',sldIds='';
  slides.forEach(function(_,i){
    presRels+='<Relationship Id="rIdS'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+(i+1)+'.xml"/>';
    sldIds+='<p:sldId id="'+(256+i)+'" r:id="rIdS'+(i+1)+'"/>';
  });
  presRels+='<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
  entries.push({name:'ppt/_rels/presentation.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+presRels+'</Relationships>'});
  entries.push({name:'ppt/presentation.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst>'+sldIds+'</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'});
  var imgCount=0;
  slides.forEach(function(s,i){
    var imgRelId=null,slideRels='<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>';
    if(s.image){
      imgCount++;
      imgRelId='rIdImg';
      slideRels+='<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+imgCount+'.jpg"/>';
      entries.push({name:'ppt/media/image'+imgCount+'.jpg',data:dataUrlToBytes(s.image)});
    }
    entries.push({name:'ppt/slides/_rels/slide'+(i+1)+'.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+slideRels+'</Relationships>'});
    entries.push({name:'ppt/slides/slide'+(i+1)+'.xml',data:pptxSlideXml(s,imgRelId)});
  });
  entries.push({name:'ppt/slideLayouts/slideLayout1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld name="Trống"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'});
  entries.push({name:'ppt/slideMasters/_rels/slideMaster1.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'});
  entries.push({name:'ppt/slideMasters/slideMaster1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="lt1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdL"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>'});
  entries.push({name:'ppt/theme/theme1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Văn phòng của Bùi"><a:themeElements><a:clrScheme name="Mặc định"><a:dk1><a:sysClr val="windowText" lastClr="1A1A1A"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="0B6FA4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="70AD47"/></a:accent5><a:accent6><a:srgbClr val="7030A0"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Mặc định"><a:majorFont><a:latin typeface="Carlito"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Carlito"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Mặc định"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'});
  return zipStored(entries);
}
function currentDeckForExport(){
  return {name:$('deckName').value.trim()||'bai-trinh-bay',slides:deckSlides};
}
function exportPptx(){
  downloadBlob(buildPptxBlob(currentDeckForExport()),(currentDeckForExport().name+'.pptx'));
}
function downloadDeckPptx(id){
  var d=loadJSON('vb_decks_v1',[]).find(function(x){return x.id===id});if(!d)return;
  downloadBlob(buildPptxBlob(d),(d.name+'.pptx'));
}


/* ---------- Khung Office Windows 1.2: app, ribbon, Backstage, status, undo ----------
   Chuẩn đối chiếu là Office bản desktop Windows (xem CHECKLIST_OFFICE_WINDOWS_1.2.md).
   Những phần không làm được y hệt trên trình duyệt điện thoại được ghi rõ trong checklist
   (ví dụ: nút cửa sổ Windows do trình duyệt giữ; ribbon cuộn ngang thay vì thu gọn). */
var currentApp='word';
var currentSlide=0;
var deckLocked=false;
var deckSize='16:9';
var deckTransition='none';
var sheetLocked=false;
var zoomLevel=100;
var rulerOn=true;
var gridOn=true;
var thumbsOn=true;
var RIBBON_TABS={
  word:[['file','Tệp'],['home','Trang đầu'],['insert','Chèn'],['design','Thiết kế'],['layout','Bố trí'],['ref','Tham chiếu'],['review','Xem lại'],['view','Xem']],
  excel:[['file','Tệp'],['home','Trang đầu'],['insert','Chèn'],['layout','Bố trí Trang'],['formulas','Công thức'],['data','Dữ liệu'],['review','Xem lại'],['view','Xem']],
  ppt:[['file','Tệp'],['home','Trang đầu'],['insert','Chèn'],['design','Thiết kế'],['transitions','Chuyển đổi'],['show','Trình chiếu'],['review','Xem lại'],['view','Xem']]
};
var RIBBON_KEYS={
  word:['home','insert','design','layout','ref','review','view'],
  excel:['home','insert','layout','formulas','data','review','view'],
  ppt:['home','insert','design','transitions','show','review','view']
};
var APP_TITLES={word:'Word',excel:'Excel',ppt:'PowerPoint',pdf:'PDF & OCR',translate:'Dịch',files:'Quản lý file',auth:'Đăng nhập & Đồng bộ',settings:'Cài đặt'};
var APP_SECTIONS={word:'editor',excel:'sheet',ppt:'slides',pdf:'pdf',translate:'translate',files:'files',auth:'auth',settings:'settings'};
var ribbonTab='home';
function switchApp(app){
  currentApp=app;
  try{document.body.className='app-'+(app==='word'?'word':app==='excel'?'excel':app==='ppt'?'ppt':'other')}catch(e){}
  /* Bản 1.3: chọn ứng dụng qua nút ☰ (Backstage → Ứng dụng). */
  Object.keys(APP_SECTIONS).forEach(function(k){
    var sec=$('tab-'+APP_SECTIONS[k]);if(sec)sec.classList.remove('active');
  });
  var sec=$('tab-'+APP_SECTIONS[app]);if(sec)sec.classList.add('active');
  var rt=$('ribbonTabs'),rb=$('ribbon');
  if(app==='word'||app==='excel'||app==='ppt'){
    if(rt)rt.style.display='flex';if(rb)rb.style.display='block';
    renderRibbonTabs();switchRibbon('home');
  }else{
    if(rt)rt.style.display='none';if(rb)rb.style.display='none';
  }
  updateTitleBar();updateStatusBar();
}
function renderRibbonTabs(){
  var rt=$('ribbonTabs');if(!rt)return;
  var tabs=RIBBON_TABS[currentApp]||RIBBON_TABS.word;
  rt.innerHTML=tabs.map(function(t){
    var cls=(t[0]==='file')?'filetab':(t[0]===ribbonTab?'active':'');
    return '<button class="'+cls+'" data-rtab="'+t[0]+'" onclick="switchRibbon(\''+t[0]+'\')">'+esc(t[1])+'</button>';
  }).join('');
}
function switchRibbon(tab){
  if(tab==='file'){openBackstage();return}
  ribbonTab=tab;
  try{document.querySelectorAll('.rpanel').forEach(function(p){p.style.display='none'})}catch(e){}
  var keys=RIBBON_KEYS[currentApp]||[];
  keys.forEach(function(k){
    var p=$('rp-'+currentApp+'-'+k);if(p)p.style.display=(k===tab?'flex':'none');
  });
  renderRibbonTabs();
}
function currentFileName(){
  if(currentApp==='word')return $('docName').value.trim()||'Tài liệu chưa đặt tên';
  if(currentApp==='excel')return $('sheetName').value.trim()||'Bảng tính chưa đặt tên';
  if(currentApp==='ppt')return $('deckName').value.trim()||'Bài trình bày chưa đặt tên';
  return APP_TITLES[currentApp]||'';
}
function updateTitleBar(){
  var an=$('appName');if(an)an.textContent=APP_TITLES[currentApp]||'Văn phòng';
  var dt=$('docTitle');if(dt)dt.textContent=currentFileName();
  var ss=$('saveState');
  if(ss){
    var saved=(currentApp==='word'&&!!currentDocId)||(currentApp==='excel'&&!!currentSheetId)||(currentApp==='ppt'&&!!currentDeckId);
    ss.textContent=(currentApp==='word'||currentApp==='excel'||currentApp==='ppt')?(saved?' • Đã lưu':' • Chưa lưu'):'';
  }
}
function updateStatusBar(){
  var l=$('statusLeft');if(!l)return;
  if(currentApp==='word'){
    var t=editorText();var words=t.trim()?t.trim().split(/\s+/).length:0;
    l.textContent='Trang 1 • '+words+' từ • '+t.length+' ký tự';
  }else if(currentApp==='excel'){
    l.textContent='Ô '+cellRef(selCell.r,selCell.c)+' • Sheet1'+(sheetLocked?' • Đã bảo vệ':'');
  }else if(currentApp==='ppt'){
    l.textContent='Slide '+(currentSlide+1)+' / '+deckSlides.length;
  }else{
    l.textContent=APP_TITLES[currentApp]||'';
  }
  var zl=$('zoomLabel');if(zl)zl.textContent=zoomLevel+'%';
  var zs=$('zoomSlider');if(zs)zs.value=String(zoomLevel);
  updateTitleBar();
}
function setZoom(v){
  zoomLevel=Math.max(50,Math.min(200,parseInt(v,10)||100));
  ['wordPage','sheetWrap','slideMain'].forEach(function(id){
    var el=$(id);if(el)el.style.zoom=zoomLevel+'%';
  });
  updateStatusBar();syncRuler();
}
function renderRuler(){
  var rn=$('rulerNums');if(!rn)return;
  var html='';
  for(var i=0;i<20;i++)html+='<span>'+(i===0?'':i)+'</span>';
  rn.innerHTML=html;
}
/* Bản 1.3: thước nằm TRONG #wordPage nên rộng đúng bằng trang giấy.
   syncRuler() đặt 2 vạch mốc lề tại đúng padding-left / padding-right
   tính được từ #editor (getComputedStyle), gọi khi tải trang, resize
   và sau mỗi lần đổi lề / khổ / hướng trang. */
function syncRuler(){
  var r=$('ruler'),ed=$('editor'),ml=$('rulerMarkLeft'),mr=$('rulerMarkRight');
  if(!r||!ed||!ml||!mr)return;
  if(r.style.display==='none')return;
  try{
    var cs=getComputedStyle(ed);
    var pl=parseFloat(cs.paddingLeft)||0,pr=parseFloat(cs.paddingRight)||0;
    var w=r.clientWidth;
    if(!w)return;
    ml.style.left=pl+'px';
    mr.style.left=Math.max(pl,w-pr)+'px';
  }catch(e){}
}
function toggleRuler(){rulerOn=!rulerOn;var r=$('ruler');if(r)r.style.display=rulerOn?'block':'none';if(rulerOn)syncRuler();}
function toggleGridlines(){
  gridOn=!gridOn;var w=$('sheetWrap');
  if(w){if(gridOn)w.classList.remove('nogrid');else w.classList.add('nogrid');}
}
function toggleThumbs(){
  thumbsOn=!thumbsOn;var l=$('slideLayout');
  if(l){if(thumbsOn)l.classList.remove('hide-thumbs');else l.classList.add('hide-thumbs');}
}
function saveCurrentApp(){
  if(currentApp==='word')saveDoc(true);
  else if(currentApp==='excel')saveSheet(true);
  else if(currentApp==='ppt')saveDeck(true);
  else{alert('Mục này không có file Office để lưu. Dùng nút Lưu riêng trong từng mục.');return}
  updateStatusBar();
}
function saveAsCurrent(){
  var nm=$('bsNewName');var v=nm?nm.value.trim():'';
  if(!v){alert('Nhập tên mới trước.');return}
  if(currentApp==='word'){currentDocId=null;$('docName').value=v;saveDoc();}
  else if(currentApp==='excel'){currentSheetId=null;$('sheetName').value=v;saveSheet();}
  else if(currentApp==='ppt'){currentDeckId=null;$('deckName').value=v;saveDeck();}
  else{alert('Mục này không có file Office để lưu thành.');return}
  closeBackstage();
}
/* Bản 1.3: ô "Tên tệp" nhìn thấy được trong Backstage → Thông tin, đồng bộ
   2 chiều với 3 input ẩn docName / sheetName / deckName (mở Thông tin thì
   điền từ input ẩn, đổi tên thì ghi lại input ẩn + updateTitleBar()). */
function currentFileInput(){
  if(currentApp==='word')return $('docName');
  if(currentApp==='excel')return $('sheetName');
  if(currentApp==='ppt')return $('deckName');
  return null;
}
function renameCurrentFile(v){
  var inp=currentFileInput();if(!inp)return;
  inp.value=v;updateTitleBar();
}
/* Backstage (tab Tệp): Mới / Mở / Lưu / Lưu thành / Thông tin — xem trước tên file. */
function openBackstage(){var b=$('backstage');if(!b)return;b.style.display='flex';backstageShow('open');}
function closeBackstage(){var b=$('backstage');if(b)b.style.display='none';}
function backstageShow(view){
  try{document.querySelectorAll('#backstageNav button').forEach(function(x){
    if(x.dataset.bs===view)x.classList.add('active');else x.classList.remove('active');
  })}catch(e){}
  var c=$('backstageContent');if(!c)return;
  if(view==='new'){
    c.innerHTML='<h3>Tạo mới</h3><p class="small">File mới luôn bắt đầu trống, không có nội dung mẫu. Định dạng mặc định: .docx / .xlsx / .pptx.</p>'
      +'<div class="row"><button onclick="closeBackstage();switchApp(\'word\');newDoc()">Tài liệu Word mới</button>'
      +'<button onclick="closeBackstage();switchApp(\'excel\');newSheet()">Bảng tính Excel mới</button>'
      +'<button onclick="closeBackstage();switchApp(\'ppt\');newDeck()">Bài trình bày PowerPoint mới</button></div>';
  }else if(view==='open'){
    var docs=loadJSON('vb_docs_v1',[]),sheets=loadJSON('vb_sheets_v1',[]),decks=loadJSON('vb_decks_v1',[]);
    var h='<h3>Mở file đã lưu trên máy này</h3><p class="small">Bấm Mở để xem và sửa tiếp. File chỉ lưu trong trình duyệt này, chưa đồng bộ lên Drive/OneDrive trừ khi Bùi bấm đồng bộ.</p>';
    if(!docs.length&&!sheets.length&&!decks.length)h+='<p class="small">Chưa có file nào lưu trên máy.</p>';
    docs.forEach(function(d){h+='<div class="bs-file"><b>'+esc(d.name)+'</b><span class="small">Word • '+new Date(d.updatedAt).toLocaleString('vi-VN')+'</span><button onclick="closeBackstage();openDoc(\''+d.id+'\')">Mở</button></div>'});
    sheets.forEach(function(sh){h+='<div class="bs-file"><b>'+esc(sh.name)+'</b><span class="small">Excel • '+new Date(sh.updatedAt).toLocaleString('vi-VN')+'</span><button onclick="closeBackstage();openSheetRec(\''+sh.id+'\')">Mở</button></div>'});
    decks.forEach(function(dk){h+='<div class="bs-file"><b>'+esc(dk.name)+'</b><span class="small">PowerPoint • '+dk.slides.length+' slide • '+new Date(dk.updatedAt).toLocaleString('vi-VN')+'</span><button onclick="closeBackstage();openDeck(\''+dk.id+'\')">Mở</button></div>'});
    c.innerHTML=h;
  }else if(view==='save'){
    saveCurrentApp();
    c.innerHTML='<h3>Đã lưu</h3><p>'+esc(currentFileName())+' đã lưu vào máy.</p>';
  }else if(view==='saveas'){
    c.innerHTML='<h3>Lưu thành tên mới</h3><p class="small">File cũ giữ nguyên, file mới lưu thêm với tên dưới đây.</p><div class="row"><input type="text" id="bsNewName" placeholder="Tên mới"><button onclick="saveAsCurrent()">Lưu thành</button></div>';
  }else if(view==='info'){
    var info='';
    if(currentApp==='word'){
      var t=editorText();
      info='Loại: Tài liệu Word (.docx)<br>Tên: '+esc(currentFileName())+'<br>Số từ: '+(t.trim()?t.trim().split(/\s+/).length:0)+'<br>Trạng thái: '+(currentDocId?'Đã lưu ở máy':'Chưa lưu');
    }else if(currentApp==='excel'){
      var used=0;sheetRaw.forEach(function(row){row.forEach(function(v){if(v)used++})});
      info='Loại: Bảng tính Excel (.xlsx)<br>Tên: '+esc(currentFileName())+'<br>Ô đã nhập: '+used+'<br>Trạng thái: '+(currentSheetId?'Đã lưu ở máy':'Chưa lưu');
    }else if(currentApp==='ppt'){
      info='Loại: Bài trình bày PowerPoint (.pptx)<br>Tên: '+esc(currentFileName())+'<br>Số slide: '+deckSlides.length+'<br>Trạng thái: '+(currentDeckId?'Đã lưu ở máy':'Chưa lưu');
    }else{
      info='Mục này là tiện ích (PDF / Dịch / File / Cài đặt), không phải file Office.';
    }
    var nameRow='';
    var fi=currentFileInput();
    if(fi)nameRow='<label>Tên tệp:</label><input type="text" id="bsFileName" value="'+esc(fi.value)+'" placeholder="Tên tệp" oninput="renameCurrentFile(this.value)">';
    /* Bản 1.3: toàn bộ hướng dẫn dùng bằng ngón tay / phím tắt trước đây nằm
       thường trực trên vùng làm việc nay gom về đây, đóng trong <details> —
       chỉ mở ra mới thấy, không chiếm chỗ soạn thảo. Nội dung giữ nguyên,
       kể cả danh sách công thức Excel và ghi chú nén ảnh slide. */
    var guide='<details><summary>Hướng dẫn dùng bằng ngón tay</summary>'
      +'<p><b>Word:</b> chạm vào trang giấy là gõ ngay. Phím tắt khi có bàn phím: Ctrl+B đậm, Ctrl+I nghiêng, Ctrl+U gạch chân, Ctrl+S lưu. Font mặc định cho tài liệu mới: Carlito 11pt (tương đương Calibri 11).</p>'
      +'<p><b>Excel:</b> chạm một lần để chọn ô, chạm hai lần liên tiếp để sửa ô, vuốt để cuộn bảng. Công thức hỗ trợ: =SUM(A1:A3), =AVERAGE(A1:A3), =COUNT, =MAX, =MIN, tham chiếu ô (=A1) và phép tính (+ − × ÷, ngoặc). Sửa xong chạm sang ô khác là tính lại ngay.</p>'
      +'<p><b>PowerPoint:</b> bấm nút trình chiếu là xem toàn màn hình; vuốt ngón tay sang trái/phải để chuyển slide, hoặc bấm hai nút Trước/Sau lớn ở dưới màn hình; bấm “Thoát” để về lại trang soạn. Ảnh chèn vào slide được nén nhỏ (tối đa 1280px) trước khi lưu để vừa bộ nhớ trình duyệt.</p>'
      +'<p><b>PDF:</b> chụm hai ngón để phóng to/thu nhỏ trang, vẽ chú thích bằng ngón tay trên trang, chạm để thêm ghi chú. Không cần chuột hay thao tác chính xác từng pixel.</p>'
      +'</details>';
    c.innerHTML='<h3>Thông tin</h3><p>'+info+'</p>'+nameRow+guide;
  }
}
/* Undo / Redo: Word dùng lệnh của trình duyệt; Excel/PowerPoint dùng ảnh chụp nhanh. */
var sheetUndo=[],sheetRedo=[];
function pushSheetHistory(){sheetUndo.push(JSON.stringify(sheetRaw));if(sheetUndo.length>50)sheetUndo.shift();sheetRedo=[];}
function undoSheet(){if(!sheetUndo.length)return;sheetRedo.push(JSON.stringify(sheetRaw));sheetRaw=JSON.parse(sheetUndo.pop());renderSheet();updateStatusBar();}
function redoSheet(){if(!sheetRedo.length)return;sheetUndo.push(JSON.stringify(sheetRaw));sheetRaw=JSON.parse(sheetRedo.pop());renderSheet();updateStatusBar();}
var deckUndo=[],deckRedo=[];
function pushDeckHistory(){deckUndo.push(JSON.stringify(deckSlides));if(deckUndo.length>50)deckUndo.shift();deckRedo=[];}
function undoDeck(){if(!deckUndo.length)return;deckRedo.push(JSON.stringify(deckSlides));deckSlides=JSON.parse(deckUndo.pop());if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;renderSlideEditor();updateStatusBar();}
function redoDeck(){if(!deckRedo.length)return;deckUndo.push(JSON.stringify(deckSlides));deckSlides=JSON.parse(deckRedo.pop());if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;renderSlideEditor();updateStatusBar();}
function undo(){if(currentApp==='word')cmd('undo');else if(currentApp==='excel')undoSheet();else if(currentApp==='ppt')undoDeck();}
function redo(){if(currentApp==='word')cmd('redo');else if(currentApp==='excel')redoSheet();else if(currentApp==='ppt')redoDeck();}
/* Ten file trung: them hau to (2), (3)... thay vi ghi de hay tao trung ten. */
function uniqueName(kind,name){
  var key=kind==='doc'?'vb_docs_v1':(kind==='sheet'?'vb_sheets_v1':'vb_decks_v1');
  var names=loadJSON(key,[]).map(function(x){return x.name});
  if(names.indexOf(name)<0)return name;
  var i=2;
  while(names.indexOf(name+' ('+i+')')>=0)i++;
  return name+' ('+i+')';
}
function parseCsv(text){
  var rows=[],row=[],field='',inq=false;
  for(var i=0;i<text.length;i++){
    var ch=text[i];
    if(inq){
      if(ch==='"'){if(text[i+1]==='"'){field+='"';i++;}else inq=false;}
      else field+=ch;
    }else if(ch==='"')inq=true;
    else if(ch===','){row.push(field);field='';}
    else if(ch==='\n'){row.push(field);rows.push(row);row=[];field='';}
    else if(ch!=='\r')field+=ch;
  }
  if(field!==''||row.length){row.push(field);rows.push(row);}
  return rows;
}
function deleteSheet(id){
  if(!confirm('Xóa bảng tính này khỏi máy?'))return;
  saveJSON('vb_sheets_v1',loadJSON('vb_sheets_v1',[]).filter(function(x){return x.id!==id}));
  if(currentSheetId===id)newSheet();renderSheetList();renderFileList();updateStatusBar();
}
function delBlob(id){
  return idb().then(function(db){return new Promise(function(res){
    try{var tx=db.transaction('blobs','readwrite');tx.objectStore('blobs').delete(id);tx.oncomplete=function(){res()};tx.onerror=function(){res()};}catch(e){res()}
  })}).catch(function(){});
}
function deleteGen(id){
  if(!confirm('Xóa file PDF này khỏi máy?'))return;
  var gens=loadJSON('vb_generated_pdfs_v1',[]);
  var g=gens.find(function(x){return x.id===id});
  saveJSON('vb_generated_pdfs_v1',gens.filter(function(x){return x.id!==id}));
  if(g)delBlob(g.blobId);
  if(g&&pdfFileName===g.name){pdfDoc=null;pdfFileName='';$('pdfWrap').style.display='none';$('pdfPageInfo').textContent='Chưa mở file nào';}
  renderFileList();
}
/* Word: chen bang / anh / ngay / duong ke / ngat trang / muc luc / nhan xet. */
function insertHtmlAtCursor(html){$('editor').focus();document.execCommand('insertHTML',false,html);updateStatusBar();}
function insertTableIntoDoc(){
  var row='<tr><td></td><td></td><td></td></tr>';
  insertHtmlAtCursor('<table><tbody>'+row+row+row+'</tbody></table><p></p>');
}
function pickDocImage(){var inp=$('docImageInput');if(inp){inp.value='';inp.click();}}
function insertDateIntoDoc(){insertHtmlAtCursor(esc((new Date()).toLocaleDateString('vi-VN')));}
function insertHRuleIntoDoc(){insertHtmlAtCursor('<hr><p></p>');}
function insertPageBreakIntoDoc(){insertHtmlAtCursor('<p style="page-break-before:always"></p>');}
function nodeText(node){
  if(!node)return '';
  if(node.nodeType===3)return node.nodeValue||'';
  var out='';
  for(var i=0;i<(node.childNodes||[]).length;i++)out+=nodeText(node.childNodes[i]);
  return out;
}
function insertTOC(){
  var heads=[];
  Array.from($('editor').childNodes||[]).forEach(function(n){
    if(n.nodeType===1){
      var tg=n.tagName.toLowerCase();
      if(tg==='h1'||tg==='h2'||tg==='h3')heads.push({lv:tg,txt:nodeText(n).trim()});
    }
  });
  if(!heads.length){alert('Chưa có tiêu đề nào (Tiêu đề 1/2/3) để lập mục lục. Đặt tiêu đề cho các phần trước, rồi bấm lại.');return}
  var html='<h2>Mục lục</h2>'+heads.map(function(h){
    var pad=h.lv==='h2'?'&nbsp;&nbsp;&nbsp;&nbsp;':(h.lv==='h3'?'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;':'');
    return '<p>'+pad+esc(h.txt)+'</p>';
  }).join('');
  insertHtmlAtCursor(html);
}
function insertFootnoteNote(){
  var txt=prompt('Nhập nội dung chú thích:');
  if(!txt)return;
  insertHtmlAtCursor('<p class="comment-para">Chú thích: '+esc(txt)+'</p>');
}
function insertCommentPara(){
  var txt=prompt('Nhập nhận xét:');
  if(!txt)return;
  insertHtmlAtCursor('<p class="comment-para">Nhận xét: '+esc(txt)+'</p>');
}
function toggleSpellcheck(){
  var ed=$('editor');if(!ed)return;
  var on=ed.getAttribute('spellcheck')!=='false';
  ed.setAttribute('spellcheck',on?'false':'true');
  alert(on?'Đã tắt kiểm tra chính tả của trình duyệt.':'Đã bật kiểm tra chính tả của trình duyệt.');
}
function setPageBg(color){var p=$('wordPage');if(p)p.style.background=color;}
function setPageMargins(kind){
  var ed=$('editor');if(!ed)return;
  ed.style.padding=(kind==='rong'?'96px':kind==='hep'?'32px':'64px');
  syncRuler();
}
function setPageSize(size){
  var p=$('wordPage');if(!p)return;
  p.classList.remove('size-a5','landscape');
  if(size==='A5')p.classList.add('size-a5');
  syncRuler();
}
function setPageOrientation(o){
  var p=$('wordPage');if(!p)return;
  if(o==='ngang')p.classList.add('landscape');else p.classList.remove('landscape');
  syncRuler();
}
/* Excel: thao tac o / hang / cot / cong thuc / sap xep / khoa. */
function updateFormulaBar(){
  var fb=$('formulaBar');var ca=$('cellAddr');
  if(ca)ca.textContent=cellRef(selCell.r,selCell.c);
  if(fb)fb.value=sheetRaw[selCell.r]?(sheetRaw[selCell.r][selCell.c]||''):'' ;
}
function commitFormulaBar(){
  var fb=$('formulaBar');if(!fb)return;
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');renderSheet();return}
  var v=fb.value;
  if(sheetRaw[selCell.r][selCell.c]!==v){pushSheetHistory();sheetRaw[selCell.r][selCell.c]=v;}
  renderSheet();
}
function insertRowAbove(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw.splice(selCell.r,0,Array(SHEET_COLS).fill(''));sheetRaw.length=SHEET_ROWS;renderSheet();
}
function insertColLeft(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();
  sheetRaw.forEach(function(row){row.splice(selCell.c,0,'');row.length=SHEET_COLS;});
  renderSheet();
}
function deleteRow(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw.splice(selCell.r,1);sheetRaw.push(Array(SHEET_COLS).fill(''));renderSheet();
}
function deleteCol(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw.forEach(function(row){row.splice(selCell.c,1);row.push('');});renderSheet();
}
function clearCell(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw[selCell.r][selCell.c]='';renderSheet();
}
function insertDateIntoCell(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw[selCell.r][selCell.c]=(new Date()).toLocaleDateString('vi-VN');renderSheet();
}
function insertFormula(fn){
  var fb=$('formulaBar');
  if(fb){fb.value='='+fn+'(';fb.focus();}
}
function sortSheet(dir){
  pushSheetHistory();
  var lastR=0;
  sheetRaw.forEach(function(row,r){row.forEach(function(v){if(v)lastR=Math.max(lastR,r)})});
  var c=selCell.c;
  var rows=sheetRaw.slice(0,lastR+1);
  rows.sort(function(a,b){
    var x=a[c]||'',y=b[c]||'';
    var nx=parseFloat(x),ny=parseFloat(y);
    var cmp=(!isNaN(nx)&&!isNaN(ny)&&String(x).trim()!==''&&String(y).trim()!=='')?(nx-ny):String(x).localeCompare(String(y),'vi');
    return dir==='desc'?-cmp:cmp;
  });
  for(var i=0;i<rows.length;i++)sheetRaw[i]=rows[i];
  renderSheet();
}
function showFormulaHelp(){
  alert('Cách viết công thức: bắt đầu bằng dấu =. Ví dụ: =SUM(A1:A3) cộng các ô A1 tới A3; =AVERAGE(A1:A3) tính trung bình; =COUNT(A1:A3) đếm ô có số; =A1+B1 cộng hai ô; =A1*2 nhân ô A1 với 2. Gõ công thức vào thanh công thức (fx) rồi Enter, hoặc chạm hai lần vào ô để sửa.');
}
function triggerImportCsv(){var f=$('importFile');if(f)f.click();else gotoTab('files');}
function toggleSheetLock(){
  sheetLocked=!sheetLocked;
  alert(sheetLocked?'Đã bảo vệ trang tính: không sửa được ô cho tới khi mở khóa.':'Đã mở khóa trang tính.');
  updateStatusBar();
}
/* PowerPoint: slide hien tai, nen, co, hieu ung, khoa, trinh chieu tu slide nay. */
function selectSlide(i){if(i<0||i>=deckSlides.length)return;currentSlide=i;renderSlideEditor();updateStatusBar();}
function deleteCurrentSlide(){deleteSlide(currentSlide);}
function moveCurrentSlide(dir){
  var j=currentSlide+dir;if(j<0||j>=deckSlides.length)return;
  pushDeckHistory();
  var t=deckSlides[currentSlide];deckSlides[currentSlide]=deckSlides[j];deckSlides[j]=t;currentSlide=j;
  renderSlideEditor();updateStatusBar();
}
function setSlideBg(color){if(deckSlides[currentSlide])deckSlides[currentSlide].bg=color;renderSlideEditor();}
function setSlideSize(r){deckSize=r;renderSlideEditor();}
function setTransition(t){deckTransition=t;renderSlideEditor();}
function insertDateBullet(){
  if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  if(!deckSlides[currentSlide])return;
  deckSlides[currentSlide].bullets.push((new Date()).toLocaleDateString('vi-VN'));
  renderSlideEditor();
}
function toggleDeckLock(){
  deckLocked=!deckLocked;
  alert(deckLocked?'Đã bảo vệ bài trình bày: không sửa được slide cho tới khi mở khóa.':'Đã mở khóa bài trình bày.');
  renderSlideEditor();
}
function presentFromCurrent(){
  if(!deckSlides.length)deckSlides=[emptySlide()];
  presentSlides=JSON.parse(JSON.stringify(deckSlides));
  presentIndex=Math.min(currentSlide,presentSlides.length-1);
  $('presentOverlay').style.display='flex';
  renderPresentSlide();
  var ov=$('presentOverlay');
  if(ov.requestFullscreen){ov.requestFullscreen().catch(function(){});}
}
/* Gan su kien cho thanh cong thuc va chen anh vao tai lieu (chay khi script nap xong). */
(function(){
  var fb=document.getElementById('formulaBar');
  if(fb&&fb.addEventListener){
    fb.addEventListener('change',function(){commitFormulaBar();});
    fb.addEventListener('keydown',function(e){if(e.key==='Enter'){e.preventDefault();commitFormulaBar();}});
  }
  var di=document.getElementById('docImageInput');
  if(di&&di.addEventListener)di.addEventListener('change',function(){
    var f=this.files[0];if(!f)return;
    compressToJpeg(f).then(function(d){insertHtmlAtCursor('<img src="'+d+'" alt="Ảnh chèn">');}).catch(function(e){alert('Không đọc được ảnh: '+e.message)});
  });
})();

/* ---------- Khởi động ---------- */
(function init(){
  var s=settings();
  $('googleClientId').value=s.googleClientId||'';$('msClientId').value=s.msClientId||'';$('ocrSpaceKey').value=s.ocrSpaceKey||'';
  googleToken=sessionStorage.getItem('vb_google_token')||null;
  renderDocList();renderSheetList();renderDeckList();renderSlideEditor();renderFileList();renderSheet();renderPhotos();updateAuthBar();
  try{renderRibbonTabs();switchApp('word');renderRuler();syncRuler();updateStatusBar();}catch(e){}
  if(window.addEventListener)window.addEventListener('resize',function(){syncRuler();});
  var ed=$('editor');if(ed&&ed.addEventListener)ed.addEventListener('input',function(){updateStatusBar();});
  document.addEventListener('keydown',function(e){
    if(!(e.ctrlKey||e.metaKey))return;
    var k=(e.key||'').toLowerCase();
    if(k==='s'){e.preventDefault();saveCurrentApp();return;}
    if(document.activeElement===$('editor')){
      if(k==='b'){e.preventDefault();cmd('bold');}
      else if(k==='i'){e.preventDefault();cmd('italic');}
      else if(k==='u'){e.preventDefault();cmd('underline');}
    }
  });
  if('serviceWorker' in navigator&&location.protocol!=='file:'){
    navigator.serviceWorker.register('sw.js').catch(function(){/* PWA là tùy chọn, app vẫn chạy không cần SW */});
  }
})();
