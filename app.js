/* Văn phòng của Bùi — bản web 1.6. HTML/JS thuần, không framework build.
   Đóng băng giao diện từ 1.5: bản này chỉ sửa cỗ máy bên trong (công thức
   Excel, giữ file khi xuất/nhập, hoàn tác, cảnh báo mất dữ liệu).
   Không thu thập dữ liệu, không nhúng key/secret. File luôn lưu ở máy trước. */
'use strict';
var FONTS=['Carlito','Be Vietnam Pro','Inter','Roboto','Noto Sans','Noto Serif'];
var $=function(id){return document.getElementById(id)};
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function uid(){return 'f'+Date.now().toString(36)+Math.floor(Math.random()*1e6).toString(36)}
function loadJSON(k,f){try{var v=JSON.parse(localStorage.getItem(k));return v==null?f:v}catch(e){return f}}
function saveJSON(k,v){try{var s=JSON.stringify(v);localStorage.setItem(k,s);return true}catch(e){alert('Không lưu được vào máy (bộ nhớ trình duyệt đầy hoặc bị chặn): '+e.message);return false}}
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
function cmd(c,v){$('editor').focus();document.execCommand(c,false,v==null?null:v);try{docNoteEdit()}catch(e){}}
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
function newDoc(){currentDocId=null;trackChangesOn=false;setPageBg('');$('docName').value='';$('editor').innerHTML='';savedDocSnap=docSnapNow();try{docHistoryReset()}catch(e){}$('editor').focus()}
function saveDoc(silent){
  var name=$('docName').value.trim()||('Tài liệu '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentDocId)name=uniqueName('doc',name);
  var html=$('editor').innerHTML;
  var docs=loadJSON('vb_docs_v1',[]);
  if(currentDocId){var d=docs.find(function(x){return x.id===currentDocId});
    if(d){d.name=name;d.html=html;d.bg=currentDocBg;d.updatedAt=Date.now()}else{docs.unshift({id:currentDocId,name:name,html:html,bg:currentDocBg,updatedAt:Date.now(),syncStatus:'local'})}
  }else{currentDocId=uid();docs.unshift({id:currentDocId,name:name,html:html,bg:currentDocBg,updatedAt:Date.now(),syncStatus:'local'})}
  if(!saveJSON('vb_docs_v1',docs)){updateTitleBar();return}
  savedDocSnap=docSnapNow();renderDocList();renderFileList();
  if(!silent)alert('Đã lưu vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openDoc(id){
  var d=loadJSON('vb_docs_v1',[]).find(function(x){return x.id===id});if(!d)return;
  currentDocId=d.id;$('docName').value=d.name;$('editor').innerHTML=d.html||'';setPageBg(d.bg||'');savedDocSnap=docSnapNow();try{docHistoryReset()}catch(e){}gotoTab('editor');$('editor').focus();
}
function deleteDoc(id){
  if(!confirm('Xóa tài liệu này khỏi máy?'))return;
  saveJSON('vb_docs_v1',loadJSON('vb_docs_v1',[]).filter(function(x){return x.id!==id}));
  if(currentDocId===id)newDoc();renderDocList();renderFileList();
}
function renderDocList(){
  /* Bản 1.4: khối danh sách dưới vùng làm việc đã chuyển vào khung Tệp → Mở;
     giữ hàm để mọi chỗ gọi cũ không vỡ khi phần tử không còn trong DOM. */
  var el=$('docList');if(!el)return;
  var docs=loadJSON('vb_docs_v1',[]);
  el.innerHTML=docs.length?docs.map(function(d){
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
  if(tag==='img'){
    var s2=node.getAttribute?node.getAttribute('src'):null;
    if(s2)out.push({image:{src:s2,w:node.naturalWidth||parseInt(node.getAttribute('width'))||0,h:node.naturalHeight||parseInt(node.getAttribute('height'))||0},fmt:Object.assign({},fmt)});
    return;
  }
  if(tag==='ins')f.ins=true;
  if(tag==='del')f.del=true;
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
function docxParaXml(p,ctx){
  ctx=ctx||{images:[],date:new Date().toISOString()};
  if(p.pageBreak&&(!p.runs||!p.runs.length))return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  var pPr='<w:pPr>';
  if(p.heading)pPr+='<w:pStyle w:val="'+p.heading+'"/>';
  if(p.align&&p.align!=='left')pPr+='<w:jc w:val="'+(p.align==='justify'?'both':p.align)+'"/>';
  if(p.bullet==='bullet')pPr+='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
  if(p.bullet==='decimal')pPr+='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>';
  if(p.indLeft||p.indFirst)pPr+='<w:ind'+(p.indLeft?' w:left="'+p.indLeft+'"':'')+(p.indFirst?' w:firstLine="'+p.indFirst+'"':'')+'/>';
  if(p.border)pPr+='<w:pBdr><w:top w:val="single" w:sz="4" w:color="'+(p.borderColor||'999999')+'"/><w:left w:val="single" w:sz="4" w:color="'+(p.borderColor||'999999')+'"/><w:bottom w:val="single" w:sz="4" w:color="'+(p.borderColor||'999999')+'"/><w:right w:val="single" w:sz="4" w:color="'+(p.borderColor||'999999')+'"/></w:pBdr>';
  if(ctx.lineSpacing)pPr+='<w:spacing w:line="'+ctx.lineSpacing+'" w:lineRule="auto"/>';
  if(p.pageBreak)pPr+='<w:pageBreakBefore/>';
  pPr+='</w:pPr>';
  var runs=p.runs.length?p.runs:[{text:'',fmt:{}}];
  var rs=runs.map(function(r){
    var f=r.fmt||{};
    if(r.image){
      var idx=ctx.images.indexOf(r.image.src);
      if(idx<0){ctx.images.push(r.image.src);idx=ctx.images.length-1}
      var rel='rIdImg'+(idx+1);
      var pw=r.image.w||0,ph=r.image.h||0;
      if(!pw&&!ph){pw=320;ph=240}
      else if(!pw)pw=Math.round(ph*4/3);
      else if(!ph)ph=Math.round(pw*3/4);
      var cx=pw*9525,cy=ph*9525;
      var picId=100+idx;
      return '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="'+cx+'" cy="'+cy+'"/><wp:docPr id="'+picId+'" name="Ảnh '+(idx+1)+'"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="'+picId+'" name=""/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="'+rel+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    }
    var rPr='<w:rPr>';
    if(f.font)rPr+='<w:rFonts w:ascii="'+esc(f.font)+'" w:hAnsi="'+esc(f.font)+'"/>';
    if(f.size)rPr+='<w:sz w:val="'+f.size+'"/><w:szCs w:val="'+f.size+'"/>';
    if(f.bold)rPr+='<w:b/>';if(f.italic)rPr+='<w:i/>';
    if(f.underline)rPr+='<w:u w:val="single"/>';if(f.strike)rPr+='<w:strike/>';
    if(f.color&&/^[0-9A-Fa-f]{6}$/.test(f.color))rPr+='<w:color w:val="'+f.color.toUpperCase()+'"/>';
    if(f.highlight)rPr+='<w:highlight w:val="'+f.highlight+'"/>';
    if(f.shd)rPr+='<w:shd w:val="clear" w:color="auto" w:fill="'+f.shd+'"/>';
    rPr+='</w:rPr>';
    var txt=esc(r.text).replace(/\n/g,'<w:br/>');
    var inner='<w:r>'+rPr+'<w:t xml:space="preserve">'+txt+'</w:t></w:r>';
    if(f.del)return '<w:del w:author="Văn phòng của Bùi" w:date="'+ctx.date+'"><w:r>'+rPr+'<w:delText xml:space="preserve">'+txt+'</w:delText></w:r></w:del>';
    if(f.ins)return '<w:ins w:author="Văn phòng của Bùi" w:date="'+ctx.date+'">'+inner+'</w:ins>';
    return inner;
  }).join('');
  return '<w:p>'+pPr+rs+'</w:p>';
}
function emptyPara(){return {kind:'para',align:'left',heading:null,runs:[{text:'',fmt:{}}],bullet:null,pageBreak:false}}
function paraFromNode(node,opts){
  opts=opts||{};
  var align='left',heading=null,indLeft=0,indFirst=0,border=false,borderColor=null;
  if(node.nodeType===1){
    var tg=node.tagName.toLowerCase();
    if(tg==='h1')heading='Heading1';if(tg==='h2')heading='Heading2';if(tg==='h3')heading='Heading3';
    var a=(node.style&&node.style.textAlign)||node.getAttribute('align')||'';
    if(a==='center'||a==='right'||a==='justify')align=a;
    if(node.style&&node.style.pageBreakBefore==='always')opts.pageBreak=true;
    if(node.style&&node.style.marginLeft){
      var ml=parseFloat(node.style.marginLeft);
      if(!isNaN(ml)&&ml!==0)indLeft=Math.round(ml*15);
    }
    if(node.style&&node.style.textIndent){
      var ti=parseFloat(node.style.textIndent);
      if(!isNaN(ti)&&ti!==0)indFirst=Math.round(ti*15);
    }
    if(node.style&&node.style.border&&node.style.border.indexOf('none')<0){
      border=true;
      borderColor=hexColor(node.style.borderColor);
      if(!borderColor){
        var bm=/([0-9A-Fa-f]{6})/.exec(node.style.border);
        if(bm)borderColor=bm[1].toUpperCase();
      }
      if(!borderColor)borderColor='999999';
    }
  }
  var runs=[];collectRuns(node,{},runs);
  return {kind:'para',align:align,heading:heading,runs:runs,bullet:opts.bullet||null,pageBreak:!!opts.pageBreak,indLeft:indLeft,indFirst:indFirst,border:border,borderColor:borderColor};
}
function nodeClass(node){
  if(!node)return '';
  try{
    if(node.getAttribute){var c=node.getAttribute('class');if(c)return c}
  }catch(e){}
  if(typeof node.className==='string')return node.className;
  return '';
}
function docxBlocksFromNode(root,ctxIn){
  var ctxB=ctxIn||{};
  if(!ctxB.headerParas)ctxB.headerParas=[];
  if(!ctxB.footerParas)ctxB.footerParas=[];
  if(!ctxB.comments)ctxB.comments=[];
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
    var clsB=nodeClass(b);
    if(clsB.indexOf('doc-header')>=0){ctxB.headerParas.push(paraFromNode(b));return}
    if(clsB.indexOf('doc-footer')>=0){ctxB.footerParas.push(paraFromNode(b));return}
    if(clsB.indexOf('doc-toc')>=0){blocks.push({kind:'toc'});return}
    if(clsB.indexOf('comment-para')>=0){
      var cp=paraFromNode(b);
      var ctext='';(cp.runs||[]).forEach(function(r){if(r.text)ctext+=r.text});
      blocks.push({kind:'comment',text:ctext});
      return;
    }
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
function docxXmlFromNode(root,ctxIn){
  var ctx=ctxIn||{images:[],date:new Date().toISOString()};
  if(!ctx.images)ctx.images=[];
  if(!ctx.date)ctx.date=new Date().toISOString();
  var lh=(root&&root.style&&root.style.lineHeight)?parseFloat(root.style.lineHeight):NaN;
  if(!isNaN(lh)&&lh>0)ctx.lineSpacing=Math.round(lh*240);
  if(!ctx.comments)ctx.comments=[];
  if(!ctx.headerParas)ctx.headerParas=[];
  if(!ctx.footerParas)ctx.footerParas=[];
  var blocks=docxBlocksFromNode(root,ctx);
  var body=blocks.map(function(b){
    if(b.kind==='toc'){
      var heads=blocks.filter(function(h){return h.heading==='Heading1'||h.heading==='Heading2'||h.heading==='Heading3'});
      var entries=heads.map(function(h){
        var lv=parseInt(h.heading.slice(-1),10);
        var txt='';(h.runs||[]).forEach(function(r){if(r.text)txt+=r.text});
        return '<w:p><w:pPr><w:ind w:left="'+((lv-1)*360)+'"/></w:pPr><w:r><w:t xml:space="preserve">'+esc(txt)+'</w:t></w:r></w:p>';
      }).join('');
      return '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Mục lục</w:t></w:r></w:p>'
        +'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>'
        +'<w:p><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r></w:p>'
        +'<w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        +entries
        +'<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    }
    if(b.kind==='comment'){
      var ci=ctx.comments.length;
      ctx.comments.push({text:b.text||''});
      return '<w:p><w:commentRangeStart w:id="'+ci+'"/><w:r><w:commentReference w:id="'+ci+'"/></w:r><w:commentRangeEnd w:id="'+ci+'"/></w:p>';
    }
    if(b.kind==='table'){
      var ncols=Math.max.apply(null,b.rows.map(function(r){return r.length}));
      var grid='<w:tblGrid>'+Array.from({length:ncols}).map(function(){return '<w:gridCol w:w="2400"/>'}).join('')+'</w:tblGrid>';
      var trs=b.rows.map(function(row){
        var tcs=row.map(function(cparas){
          return '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'+cparas.map(function(pp){return docxParaXml(pp,ctx)}).join('')+'</w:tc>';
        }).join('');
        return '<w:tr>'+tcs+'</w:tr>';
      }).join('');
      return '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'+grid+trs+'</w:tbl>';
    }
    return docxParaXml(b,ctx);
  }).join('');
  var bgXml='';
  if(ctx.bg){
    var bgHex=hexColor(ctx.bg);
    if(bgHex)bgXml='<w:background w:color="'+bgHex+'"/>';
  }
  var sectPr='<w:sectPr>'
    +((ctx.headerParas&&ctx.headerParas.length)?'<w:headerReference w:type="default" r:id="rIdHdr"/>':'')
    +((ctx.footerParas&&ctx.footerParas.length)?'<w:footerReference w:type="default" r:id="rIdFtr"/>':'')
    +'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+bgXml+'<w:body>'+body+sectPr+'</w:body></w:document>';
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
function bytesToBase64(bytes){
  var bin='';
  var CH=0x8000;
  for(var i=0;i<bytes.length;i+=CH)bin+=String.fromCharCode.apply(null,bytes.subarray(i,i+CH));
  return btoa(bin);
}
function buildDocxBlobFromNode(root,opts){
  opts=opts||{};
  var ctx={images:[],date:new Date().toISOString(),bg:opts.bg||''};
  var docXml=docxXmlFromNode(root,ctx);
  /* Ban 1.7: dau/chan trang that + ghi chu that. header/footer la part
     rieng (header1.xml/footer1.xml) tham chieu tu sectPr; ghi chu la
     comments.xml tham chieu tu document rels. Anh trong dau/chan trang
     (neu co) dung chung rel anh cua tai lieu (rIdImg*). */
  var headerXml=null,footerXml=null,commentsXml=null;
  if(ctx.headerParas&&ctx.headerParas.length){
    headerXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>'
      +ctx.headerParas.map(function(pp){return docxParaXml(pp,ctx)}).join('')+'</w:body></w:document>';
  }
  if(ctx.footerParas&&ctx.footerParas.length){
    footerXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>'
      +ctx.footerParas.map(function(pp){return docxParaXml(pp,ctx)}).join('')+'</w:body></w:document>';
  }
  if(ctx.comments&&ctx.comments.length){
    commentsXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      +ctx.comments.map(function(cm,ci){
        return '<w:comment w:id="'+ci+'" w:author="Văn phòng của Bùi" w:date="'+ctx.date+'"><w:p><w:r><w:t xml:space="preserve">'+esc(cm.text)+'</w:t></w:r></w:p></w:comment>';
      }).join('')+'</w:comments>';
  }
  var rels='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
  var typesExtra='';
  var entries=[
    {name:'[Content_Types].xml',data:''},
    {name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'},
    {name:'word/document.xml',data:docXml},
    {name:'word/styles.xml',data:DOCX_STYLES},
    {name:'word/numbering.xml',data:DOCX_NUMBERING},
    {name:'word/settings.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:displayBackgroundShape/></w:settings>'}
  ];
  ctx.images.forEach(function(src,i){
    var rel='rIdImg'+(i+1);
    var mime='image/jpeg';var ext='jpg';
    var mm=/^data:([^;]+);base64,(.*)$/.exec(src||'');
    if(mm){
      mime=mm[1];
      var sub=mime.split('/')[1]||'jpg';
      sub=sub.replace(/[^a-z0-9]/g,'');
      ext=(sub==='jpeg')?'jpg':(sub||'jpg');
    }
    rels+='<Relationship Id="'+rel+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image'+(i+1)+'.'+ext+'"/>';
    typesExtra+='<Default Extension="'+ext+'" ContentType="'+mime+'"/>';
    var bytes=null;
    if(mm){try{bytes=dataUrlToBytes(src)}catch(e){bytes=null}}
    if(bytes)entries.push({name:'word/media/image'+(i+1)+'.'+ext,data:bytes});
  });
  var imgRelsXml='';
  ctx.images.forEach(function(src,i){
    var mm2=/^data:([^;]+);base64,/.exec(src||'');
    var ext2='jpg';
    if(mm2){var sub2=mm2[1].split('/')[1]||'jpg';sub2=sub2.replace(/[^a-z0-9]/g,'');ext2=(sub2==='jpeg')?'jpg':(sub2||'jpg')}
    imgRelsXml+='<Relationship Id="rIdImg'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image'+(i+1)+'.'+ext2+'"/>';
  });
  if(headerXml){
    rels+='<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>';
    typesExtra+='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
    entries.push({name:'word/header1.xml',data:headerXml});
    entries.push({name:'word/_rels/header1.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+imgRelsXml+'</Relationships>'});
  }
  if(footerXml){
    rels+='<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>';
    typesExtra+='<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
    entries.push({name:'word/footer1.xml',data:footerXml});
    entries.push({name:'word/_rels/footer1.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+imgRelsXml+'</Relationships>'});
  }
  if(commentsXml){
    rels+='<Relationship Id="rIdCmt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';
    typesExtra+='<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>';
    entries.push({name:'word/comments.xml',data:commentsXml});
  }
  entries.push({name:'word/_rels/document.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels+'</Relationships>'});
  entries[0].data='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'+typesExtra+'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>';
  return zipStored(entries);
}
function exportDocx(){
  if(!confirmExportWarnings(exportWarnings('word')))return;
  var name=($('docName').value.trim()||'tai-lieu')+'.docx';
  downloadBlob(buildDocxBlobFromNode($('editor'),{bg:currentDocBg}),name);
}
function testFont(f){$('fontTest').style.fontFamily="'"+f+"', -apple-system, 'Segoe UI', sans-serif"}

/* ---------- Bảng tính: chạm chọn, chạm hai lần sửa, vuốt cuộn ---------- */
var SHEET_ROWS=100,SHEET_COLS=26;
var sheetMerges=[];
var sheetRaw=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});
var currentSheetId=null,selCell={r:0,c:0},lastTap={key:'',t:0};
function colName(c){var n=c+1,out='';while(n>0){var m=(n-1)%26;out=String.fromCharCode(65+m)+out;n=Math.floor((n-1)/26)}return out}
function colIndexFromName(nm){var v=0;for(var i=0;i<nm.length;i++)v=v*26+(nm.charCodeAt(i)-64);return v-1}
function cellRef(r,c){return colName(c)+(r+1)}
function parseRef(s){var m=/^([A-Z]+)([1-9][0-9]*)$/.exec(s.trim().toUpperCase());if(!m)return null;return {r:parseInt(m[2],10)-1,c:colIndexFromName(m[1])}}
/* Ban 1.7: gop o. sheetMerges: danh sach {r1,c1,r2,c2} (o trai-tren giu gia
   tri, cac o con lai trong vung trong va khong ve rieng). Cong thuc tro toi
   o trong vung gop tinh theo gia tri o trai-tren (qua cellValue), con hien
   thi (displayCell) cua o con lai la rong — dung nhu Excel ve mat hinh. */
function normalizeMerges(list){
  var out=[],seen={};
  (list||[]).forEach(function(m){
    if(!m)return;
    var r1=Math.min(+m.r1,+m.r2),r2=Math.max(+m.r1,+m.r2);
    var c1=Math.min(+m.c1,+m.c2),c2=Math.max(+m.c1,+m.c2);
    if(isNaN(r1)||isNaN(r2)||isNaN(c1)||isNaN(c2))return;
    if(r1<0||c1<0||r2>=SHEET_ROWS||c2>=SHEET_COLS)return;
    if(r1===r2&&c1===c2)return;
    var k=r1+','+c1+','+r2+','+c2;
    if(seen[k])return;seen[k]=true;
    out.push({r1:r1,c1:c1,r2:r2,c2:c2});
  });
  return out;
}
function mergeAtTop(r,c){
  for(var i=0;i<sheetMerges.length;i++){var m=sheetMerges[i];if(m.r1===r&&m.c1===c)return m}
  return null;
}
function mergeTopLeft(r,c){
  for(var i=0;i<sheetMerges.length;i++){var m=sheetMerges[i];
    if(r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2)return {r:m.r1,c:m.c1,merge:m}}
  return null;
}
function isMergeSlave(r,c){var t=mergeTopLeft(r,c);return !!t&&(t.r!==r||t.c!==c)}
function mergeRangeInternal(r1,c1,r2,c2){
  r1=Math.min(r1,r2);r2=Math.max(r1,r2);c1=Math.min(c1,c2);c2=Math.max(c1,c2);
  if(r1<0||c1<0||r2>=SHEET_ROWS||c2>=SHEET_COLS)return false;
  if(r1===r2&&c1===c2)return false;
  sheetMerges=sheetMerges.filter(function(m){return m.r2<r1||m.r1>r2||m.c2<c1||m.c1>c2});
  for(var r=r1;r<=r2;r++)for(var cc=c1;cc<=c2;cc++){
    if(r===r1&&cc===c1)continue;
    if(sheetRaw[r])sheetRaw[r][cc]='';
    delete cellFmt[cellFmtKey(r,cc)];
  }
  sheetMerges.push({r1:r1,c1:c1,r2:r2,c2:c2});
  sheetMerges=normalizeMerges(sheetMerges);
  return true;
}
function mergeRange(r1,c1,r2,c2){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return false}
  pushSheetHistory();
  var ok=mergeRangeInternal(r1,c1,r2,c2);
  if(!ok){undoSheet();}
  renderSheet();updateStatusBar();
  return ok;
}
function unmergeAtInternal(r,c){
  var t=mergeTopLeft(r,c);if(!t)return false;
  sheetMerges=sheetMerges.filter(function(m){return m!==t.merge});
  return true;
}
function unmergeAt(r,c){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return false}
  pushSheetHistory();
  var ok=unmergeAtInternal(r,c);
  renderSheet();updateStatusBar();
  return ok;
}
function mergeSelectedCells(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var def=cellRef(selCell.r,selCell.c)+':'+cellRef(selCell.r,Math.min(selCell.c+1,SHEET_COLS-1));
  var txt=prompt('Nhập vùng cần gộp (ví dụ A1:C1, A1:B3):',def);if(!txt)return;
  var parts=txt.split(':');
  var a=parseRef(parts[0]||''),b=parts.length>1?parseRef(parts[1]):a;
  if(!a||!b){alert('Vùng gộp không đúng dạng (ví dụ A1:C1).');return}
  if(!mergeRange(a.r,a.c,b.r,b.c))alert('Không gộp được vùng này (cần ít nhất 2 ô trong lưới).');
  else selCell={r:Math.min(a.r,b.r),c:Math.min(a.c,b.c)};
}
function unmergeSelectedCell(){
  if(!unmergeAt(selCell.r,selCell.c))alert('Ô đang chọn không nằm trong vùng gộp nào.');
}
function shiftMergesRows(at,delta){
  var out=[];
  sheetMerges.forEach(function(m){
    if(delta>0&&m.r1>=at)out.push({r1:m.r1+delta,c1:m.c1,r2:m.r2+delta,c2:m.c2});
    else if(delta>0&&m.r1<at&&m.r2>=at)out.push({r1:m.r1,c1:m.c1,r2:m.r2+delta,c2:m.c2});
    else if(delta<0&&m.r2<at)out.push(m);
    else if(delta<0&&m.r1>at)out.push({r1:m.r1+delta,c1:m.c1,r2:m.r2+delta,c2:m.c2});
    else if(delta<0&&m.r1<=at&&m.r2>=at){
      if(m.r1===m.r2){/* hang duy nhat bi xoa: bo gop, giu gia tri */}
      else if(m.r1===at)out.push({r1:m.r1,c1:m.c1,r2:m.r2+delta,c2:m.c2});
      else out.push({r1:m.r1,c1:m.c1,r2:m.r2+delta,c2:m.c2});
    }else out.push(m);
  });
  sheetMerges=normalizeMerges(out);
}
function shiftMergesCols(at,delta){
  var out=[];
  sheetMerges.forEach(function(m){
    if(delta>0&&m.c1>=at)out.push({r1:m.r1,c1:m.c1+delta,r2:m.r2,c2:m.c2+delta});
    else if(delta>0&&m.c1<at&&m.c2>=at)out.push({r1:m.r1,c1:m.c1,r2:m.r2,c2:m.c2+delta});
    else if(delta<0&&m.c2<at)out.push(m);
    else if(delta<0&&m.c1>at)out.push({r1:m.r1,c1:m.c1+delta,r2:m.r2,c2:m.c2+delta});
    else if(delta<0&&m.c1<=at&&m.c2>=at){
      if(m.c1===m.c2){}
      else if(m.c1===at)out.push({r1:m.r1,c1:m.c1,r2:m.r2,c2:m.c2+delta});
      else out.push({r1:m.r1,c1:m.c1,r2:m.r2,c2:m.c2+delta});
    }else out.push(m);
  });
  sheetMerges=normalizeMerges(out);
}
function remapMergesAfterRowReorder(oldToNew){
  var out=[];
  sheetMerges.forEach(function(m){
    if(m.r1===m.r2){
      var nr=oldToNew[m.r1];
      if(nr>=0)out.push({r1:nr,c1:m.c1,r2:nr,c2:m.c2});
      return;
    }
    var mapped=[];
    for(var r=m.r1;r<=m.r2;r++){var nr2=oldToNew[r];if(nr2<0)return;mapped.push(nr2)}
    mapped.sort(function(a,b){return a-b});
    var cont=true;
    for(var i=1;i<mapped.length;i++)if(mapped[i]!==mapped[0]+i)cont=false;
    if(cont)out.push({r1:mapped[0],c1:m.c1,r2:mapped[mapped.length-1],c2:m.c2});
    /* Vung gop nhieu hang bi tach khi sap xep/xoa trung: bo gop, giu gia tri. */
  });
  sheetMerges=normalizeMerges(out);
}

/* ---------- Ban 1.6: gia tri so THAT tach khoi dinh dang hien thi ----------
   LOI CU: o go "1.000.000 ₫" hay "50%" la chuoi hien thi; cong thuc tro toi
   o do dung parseFloat nen chi lay duoc 1 (cat o dau cham) hoac 50, tinh sai.
   SUA TAN GOC: moi o giu nguyen chuoi nguoi go (sheetRaw). Khi can so, ham
   parseCellNumber() dich chuoi do thanh so THAT theo cach viet tieng Viet:
   dau cham phan tach hang nghin, dau phay la thap phan, hau to ₫/đ/VND bi
   bo, hau to % chia cho 100. Moi cong thuc tinh tren so that; hien thi thi
   ap dinh dang (tien / % / so thap phan) len so that do. Vi du: o go "50%"
   co gia tri that 0,5 — hien thi 50%, cong thuc =A1*2 ra 1 (khong phai 100).
   Quy uoc giong Excel: o dinh dang % chua 0,5 thi hien 50%; o go 50 roi moi
   ap dinh dang % thi hien 5000% (vi 50 = 5000%). */
function parseViNumber(t){
  var s=String(t==null?'':t).replace(/ /g,' ').trim();
  if(!s)return null;
  var neg=false;
  if(s.charAt(0)==='-'){neg=true;s=s.slice(1)}
  else if(s.charAt(0)==='+'){s=s.slice(1)}
  if(!s)return null;
  if(s.indexOf(',')>=0&&s.indexOf('.')>=0){s=s.replace(/\./g,'').replace(',','.')}
  else if(s.indexOf(',')>=0){
    if(/^\d{1,3}(,\d{3}){2,}$/.test(s))s=s.replace(/,/g,'');
    else s=s.replace(',','.');
  }else if(s.indexOf('.')>=0){
    if(/^\d{1,3}(\.\d{3})+$/.test(s))s=s.replace(/\./g,'');
  }
  if(!/^\d+(\.\d+)?$/.test(s))return null;
  var n=parseFloat(s);
  return neg?-n:n;
}
function parseCellNumber(s){
  if(s==null)return null;
  var t=String(s).replace(/ /g,' ').trim();
  if(!t)return null;
  if(t.charAt(t.length-1)==='%'){
    var inner=parseViNumber(t.slice(0,-1));
    return inner==null?null:inner/100;
  }
  t=t.replace(/(₫|vnđ|vnd|đ)$/i,'').trim();
  return parseViNumber(t);
}
/* Gia tri cua mot o: so that neu dich duoc, khong thi chuoi goc ('' neu rong).
   O cong thuc: ket qua tinh; o khong cong thuc: so that neu la so, khong thi
   chuoi goc. Ham nay la cho DUY NHAT moi noi can gia tri (cong thuc, bieu do,
   sap xep, xuat file) duoc dung — khong dung parseFloat truc tiep nua. */
function lookupSheetEntry(name){
  var want=String(name==null?'':name).toLowerCase();
  if(typeof wbSheets!=='undefined'&&wbSheets){
    for(var i=0;i<wbSheets.length;i++){
      if(String(wbSheets[i].name||'').toLowerCase()===want)return {raw:wbSheets[i].raw,merges:wbSheets[i].merges||[]};
    }
  }
  return null;
}
function sheetEntryForEval(name){
  /* Trang dang sua luon dung du lieu that trong bien toan cuc (khong dung
     ban sao cu trong wbSheets); trang khac dung ban da luu luc chuyen trang. */
  if(String(name==null?'':name).toLowerCase()===String(currentSheetName()).toLowerCase())
    return {raw:sheetRaw,merges:sheetMerges};
  return lookupSheetEntry(name);
}
function cellValueSheet(sheetName,r,c,seen){
  var entry=sheetEntryForEval(sheetName);
  if(!entry||!entry.raw)return '#LỖI';
  if(r<0||r>=entry.raw.length||c<0)return '';
  var row=entry.raw[r]||[];
  if(c>=row.length)return '';
  var mg=entry.merges||[];
  for(var i=0;i<mg.length;i++){var m=mg[i];
    if(r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2&&!(r===m.r1&&c===m.c1))
      return cellValueSheet(sheetName,m.r1,m.c1,seen);}
  var raw=row[c]||'';
  if(raw.charAt(0)==='=')return evalCellOn(sheetName,r,c,seen);
  var n=parseCellNumber(raw);
  return n==null?raw:n;
}
function cellValue(r,c,seen){
  return cellValueSheet(currentSheetName(),r,c,seen);
}
function cellNumericValue(r,c,seen){
  var v=cellValue(r,c,seen);
  if(typeof v==='number')return v;
  if(typeof v==='string')return parseCellNumber(v);
  return null;
}
function formulaErr(msg){var e=new Error(msg);throw e}
function evalCellOn(sheetName,r,c,seen){
  seen=seen||new Set();
  var key=String(sheetName==null?'':sheetName).toLowerCase()+'!'+r+','+c;
  if(seen.has(key))return '#LỖI';
  var entry=sheetEntryForEval(sheetName);
  if(!entry||!entry.raw)return '#LỖI';
  if(r<0||r>=entry.raw.length||c<0)return '';
  var row=entry.raw[r]||[];
  if(c>=row.length)return '';
  var mg=entry.merges||[];
  for(var i2=0;i2<mg.length;i2++){var mm=mg[i2];
    if(r>=mm.r1&&r<=mm.r2&&c>=mm.c1&&c<=mm.c2&&!(r===mm.r1&&c===mm.c1))
      return cellValueSheet(sheetName,mm.r1,mm.c1,seen);}
  var raw=row[c]||'';
  if(raw.charAt(0)!=='=')return raw;
  seen.add(key);
  var formulaSheet=sheetName;
  var expr=raw.slice(1).trim();
  if(!expr)return raw;
  var toks=[],i=0;
  try{
    while(i<expr.length){
      var ch=expr[i];
      if(ch===' '||ch==='\t'){i++;continue}
      if(ch==='$'){i++;continue}
      /* Tham chieu lien trang: 'Ten trang'!A1 hoac Sheet2!A1. Ten co dau
         cach/dau tieng Viet phai dung dang co ngoac don; ten khong dau
         viet lien duoc dung dang khong ngoac. */
      if(ch==="'"){
        var qm=/^'([^']+)'!([A-Za-z]+[0-9]+)/i.exec(expr.slice(i));
        if(qm){toks.push({t:'sref',sheet:qm[1],ref:qm[2].toUpperCase()});i+=qm[0].length;continue}
        formulaErr('tok');
      }
      var sm=/^([A-Za-z_][A-Za-z0-9_]*)!([A-Za-z]+[0-9]+)/i.exec(expr.slice(i));
      if(sm){toks.push({t:'sref',sheet:sm[1],ref:sm[2].toUpperCase()});i+=sm[0].length;continue}
      if('+-*/(),;:'.indexOf(ch)>=0){toks.push({t:'op',v:ch});i++;continue}
      if(ch==='<'||ch==='>'){
        if(expr[i+1]==='='){toks.push({t:'op',v:ch+'='});i+=2;continue}
        if(ch==='<'&&expr[i+1]==='>'){toks.push({t:'op',v:'<>'});i+=2;continue}
        toks.push({t:'op',v:ch});i++;continue;
      }
      if(ch==='='){toks.push({t:'op',v:'='});i++;continue}
      if(ch==='%'){toks.push({t:'op',v:'%'});i++;continue}
      if(ch==='"'){
        var j=i+1,buf='';
        while(j<expr.length){
          if(expr[j]==='"'){
            if(expr[j+1]==='"'){buf+='"';j+=2;continue}
            break;
          }
          buf+=expr[j];j++;
        }
        if(j>=expr.length)formulaErr('str');
        toks.push({t:'str',v:buf});i=j+1;continue;
      }
      var rm=/^[A-Z]+[1-9][0-9]*/i.exec(expr.slice(i));
      if(rm){toks.push({t:'ref',v:rm[0].toUpperCase()});i+=rm[0].length;continue}
      var nm=/^(\d+(\.\d+)?|\.\d+)/.exec(expr.slice(i));
      if(nm){toks.push({t:'num',v:parseFloat(nm[0])});i+=nm[0].length;continue}
      var im=/^[A-Za-z_][A-Za-z0-9_]*/.exec(expr.slice(i));
      if(im){toks.push({t:'ident',v:im[0].toUpperCase()});i+=im[0].length;continue}
      formulaErr('tok');
    }
  }catch(e){return '#LỖI'}
  if(!toks.length)return raw;
  var pos=0;
  function pk(){return toks[pos]}
  function isOp(v){var t=pk();return t&&t.t==='op'&&t.v===v}
  function eatOp(v){if(isOp(v)){pos++;return true}return false}
  function expectOp(v){if(!eatOp(v))formulaErr('paren')}
  function parseCmp(){
    var l=parseAdd();
    var t=pk();
    if(t&&t.t==='op'&&['=','<>','<','>','<=','>='].indexOf(t.v)>=0){
      var op=t.v;pos++;
      var r2=parseAdd();
      return {t:'cmp',op:op,l:l,r:r2};
    }
    return l;
  }
  function parseAdd(){
    var v=parseMul();
    while(pk()&&(isOp('+')||isOp('-'))){var op=pk().v;pos++;var rhs=parseMul();v={t:'bin',op:op,l:v,r:rhs}}
    return v;
  }
  function parseMul(){
    var v=parseUnary();
    while(pk()&&(isOp('*')||isOp('/'))){var op=pk().v;pos++;var rhs=parseUnary();v={t:'bin',op:op,l:v,r:rhs}}
    return v;
  }
  function parseUnary(){
    if(isOp('-')){pos++;return {t:'neg',v:parseUnary()}}
    if(isOp('+')){pos++;return parseUnary()}
    return parsePostfix();
  }
  function parsePostfix(){
    var v=parsePrimary();
    while(isOp('%')){pos++;v={t:'pct',v:v}}
    return v;
  }
  function parsePrimary(){
    var tk=pk();
    if(!tk)formulaErr('empty');
    if(tk.t==='num'){pos++;return {t:'num',v:tk.v}}
    if(tk.t==='str'){pos++;return {t:'str',v:tk.v}}
    if(isOp('(')){pos++;var v=parseCmp();expectOp(')');return v}
    if(tk.t==='ref'){
      pos++;
      if(isOp(':')){
        pos++;
        var tk2=pk();
        if(!tk2||tk2.t!=='ref')formulaErr('range');
        pos++;
        return {t:'range',a:tk.v,b:tk2.v};
      }
      return {t:'ref',v:tk.v};
    }
    if(tk.t==='sref'){
      pos++;
      if(isOp(':')){
        pos++;
        var tk3=pk();
        if(!tk3)formulaErr('range');
        if(tk3.t==='ref'){pos++;return {t:'srange',sheet:tk.sheet,a:tk.ref,b:tk3.v}}
        if(tk3.t==='sref'){pos++;return {t:'srange',sheet:tk.sheet,a:tk.ref,b:tk3.ref}}
        formulaErr('range');
      }
      return {t:'sref',sheet:tk.sheet,ref:tk.ref};
    }
    if(tk.t==='ident'){
      var name=tk.v;
      if((name==='TRUE'||name==='FALSE')&&!isOp('(')){pos++;return {t:'bool',v:(name==='TRUE')}}
      pos++;
      if(!isOp('('))formulaErr('ident');
      pos++;
      var args=[];
      if(!isOp(')')){
        while(true){
          args.push(parseCmp());
          if(isOp(',')||isOp(';')){pos++;continue}
          break;
        }
      }
      expectOp(')');
      return {t:'call',name:name,args:args};
    }
    formulaErr('tok');
  }
  function toNum(v){
    if(typeof v==='number')return v;
    if(typeof v==='boolean')return v?1:0;
    if(typeof v==='string'){
      var n=parseCellNumber(v);
      if(n!=null)return n;
    }
    formulaErr('notnum');
  }
  function isTruthy(v){
    if(typeof v==='boolean')return v;
    if(typeof v==='number')return v!==0;
    if(typeof v==='string')return v.trim().toUpperCase()==='TRUE';
    return !!v;
  }
  function rangeCells(node){
    /* Tra ve danh sach {sheet,r,c} cua mot vung (cung trang hoac lien trang). */
    var sheet, a, b;
    if(node.t==='range'){sheet=formulaSheet;a=parseRef(node.a);b=parseRef(node.b)}
    else if(node.t==='srange'){sheet=node.sheet;a=parseRef(node.a);b=parseRef(node.b)}
    else formulaErr('range');
    if(!a||!b)formulaErr('range');
    var out=[];
    for(var rr=Math.min(a.r,b.r);rr<=Math.max(a.r,b.r);rr++)
      for(var cc=Math.min(a.c,b.c);cc<=Math.max(a.c,b.c);cc++)
        out.push({sheet:sheet,r:rr,c:cc});
    return out;
  }
  function expandArg(node,s){
    if(node.t==='range'||node.t==='srange'){
      var out=[];
      rangeCells(node).forEach(function(cell){
        var cvv=cellValueSheet(cell.sheet,cell.r,cell.c,new Set(s));
        if(typeof cvv==='string'&&cvv.charAt(0)==='#')formulaErr('notnum');
        out.push({v:cvv,fromRef:true});
      });
      return out;
    }
    if(node.t==='ref'){
      var rf=parseRef(node.v);
      if(!rf)formulaErr('ref');
      var svv=cellValueSheet(formulaSheet,rf.r,rf.c,new Set(s));
      if(typeof svv==='string'&&svv.charAt(0)==='#')formulaErr('notnum');
      return [{v:svv,fromRef:true}];
    }
    if(node.t==='sref'){
      var rf2=parseRef(node.ref);
      if(!rf2)formulaErr('ref');
      var svv2=cellValueSheet(node.sheet,rf2.r,rf2.c,new Set(s));
      if(typeof svv2==='string'&&svv2.charAt(0)==='#')formulaErr('notnum');
      return [{v:svv2,fromRef:true}];
    }
    return [{v:evalNode(node,s),fromRef:false}];
  }
  var _viCollator=null;
function cmpTextVi(a,b){
  try{
    if(typeof Intl!=='undefined'&&Intl.Collator){
      if(!_viCollator)_viCollator=new Intl.Collator('vi',{sensitivity:'accent'});
      if(_viCollator)return _viCollator.compare(String(a),String(b));
    }
  }catch(e){}
  var al=String(a).toLowerCase(),bl=String(b).toLowerCase();
  return al<bl?-1:(al>bl?1:0);
}
function criteriaMatches(cellVal,crit){
    /* Tieu chi COUNTIF/SUMIF: so, chu (khong phan biet hoa thuong), hoac
       chuoi so sanh dang ">5", ">=10", "<>An". O rong chi dem khi tieu
       chi la chuoi rong. */
    if(cellVal===''||cellVal==null)return crit==='';
    var critStr=(typeof crit==='string')?crit.trim():null;
    if(critStr!=null){
      var cm=/^(>=|<=|<>|>|<|=)(.*)$/.exec(critStr);
      if(cm){
        var op=cm[1],operand=cm[2].trim();
        var on=parseCellNumber(operand);
        var cn=(typeof cellVal==='number')?cellVal:parseCellNumber(cellVal);
        if(on!=null&&cn!=null){
          if(op==='>')return cn>on;if(op==='<')return cn<on;
          if(op==='>=')return cn>=on;if(op==='<=')return cn<=on;
          if(op==='=')return Math.abs(cn-on)<1e-9;if(op==='<>')return Math.abs(cn-on)>=1e-9;
        }
        var cs=String(cellVal).trim().toLowerCase(),os=operand.toLowerCase();
        if(op==='=')return cs===os;if(op==='<>')return cs!==os;
        if(op==='>')return cs>os;if(op==='<')return cs<os;
        if(op==='>=')return cs>=os;if(op==='<=')return cs<=os;
        return false;
      }
      if(typeof cellVal==='number'){
        var pn=parseCellNumber(critStr);
        if(pn!=null)return Math.abs(cellVal-pn)<1e-9;
      }
      return String(cellVal).trim().toLowerCase()===critStr.toLowerCase();
    }
    if(typeof crit==='number'&&typeof cellVal==='number')return Math.abs(cellVal-crit)<1e-9;
    if(typeof crit==='number'&&typeof cellVal==='string'){
      var pn2=parseCellNumber(cellVal);
      return pn2!=null&&Math.abs(pn2-crit)<1e-9;
    }
    if(typeof crit==='boolean')return cellVal===crit;
    return String(cellVal)===String(crit);
  }
  function evalNode(node,s){
    s=new Set(s);
    if(node.t==='num')return node.v;
    if(node.t==='str')return node.v;
    if(node.t==='bool')return node.v;
    if(node.t==='pct')return toNum(evalNode(node.v,s))/100;
    if(node.t==='neg')return -toNum(evalNode(node.v,s));
    if(node.t==='bin'){
      var l=toNum(evalNode(node.l,s)),r2=toNum(evalNode(node.r,s));
      if(node.op==='+')return l+r2;
      if(node.op==='-')return l-r2;
      if(node.op==='*')return l*r2;
      if(node.op==='/'){if(r2===0)formulaErr('div0');return l/r2}
    }
    if(node.t==='cmp'){
      var cl=evalNode(node.l,s),cr=evalNode(node.r,s);
      var bothNum=(typeof cl==='number')&&(typeof cr==='number');
      if(!bothNum&&(typeof cl==='string'||typeof cr==='string')){
        cl=String(cl).trim().toLowerCase();cr=String(cr).trim().toLowerCase();
      }
      if(node.op==='=')return cl===cr;
      if(node.op==='<>')return cl!==cr;
      if(node.op==='<')return cl<cr;
      if(node.op==='>')return cl>cr;
      if(node.op==='<=')return cl<=cr;
      if(node.op==='>=')return cl>=cr;
    }
    if(node.t==='ref'){
      var rf3=parseRef(node.v);
      if(!rf3)formulaErr('ref');
      var vv=cellValueSheet(formulaSheet,rf3.r,rf3.c,s);
      if(vv===''||vv==null)return 0;
      /* Ban 1.7: tham chieu tra ve dung gia tri o (so/chu/boolean) nhu
         Excel; phep tinh se tu ep kieu qua toNum va bao #LOI neu chu
         khong phai so. Cu: tham chieu o chu luon #LOI ke ca khi chi
         can lay chu (vi du =VLOOKUP(A1,...) voi A1 la chu). */
      return vv;
    }
    if(node.t==='sref'){
      var rf4=parseRef(node.ref);
      if(!rf4)formulaErr('ref');
      var vv2=cellValueSheet(node.sheet,rf4.r,rf4.c,s);
      if(vv2===''||vv2==null)return 0;
      return vv2;
    }
    if(node.t==='range'||node.t==='srange')formulaErr('range');
    if(node.t==='call'){
      var fn=node.name;
      if(fn==='IF'){
        if(node.args.length<2)formulaErr('args');
        var cond=isTruthy(evalNode(node.args[0],s));
        if(cond)return evalNode(node.args[1],s);
        if(node.args.length>=3)return evalNode(node.args[2],s);
        return false;
      }
      if(fn==='ROUND'){
        if(!node.args.length)formulaErr('args');
        var xv=toNum(evalNode(node.args[0],s));
        var nd=node.args.length>=2?Math.round(toNum(evalNode(node.args[1],s))):0;
        var f=Math.pow(10,nd);
        var rv=Math.sign(xv)*Math.round(Math.abs(xv)*f+1e-9)/f;
        return (nd>0)?Number(rv.toFixed(nd)):rv;
      }
      if(fn==='ABS'){
        if(node.args.length!==1)formulaErr('args');
        return Math.abs(toNum(evalNode(node.args[0],s)));
      }
      if(fn==='COUNTIF'){
        if(node.args.length<2)formulaErr('args');
        var crit=evalNode(node.args[1],s);
        var cnt=0;
        rangeCells(node.args[0]).forEach(function(cell){
          var cv=cellValueSheet(cell.sheet,cell.r,cell.c,new Set(s));
          if(criteriaMatches(cv,crit))cnt++;
        });
        return cnt;
      }
      if(fn==='SUMIF'){
        if(node.args.length<2)formulaErr('args');
        var crit2=evalNode(node.args[1],s);
        var rng=rangeCells(node.args[0]);
        var sumRng=(node.args.length>=3)?rangeCells(node.args[2]):rng;
        var total=0;
        rng.forEach(function(cell,idx){
          var cv2=cellValueSheet(cell.sheet,cell.r,cell.c,new Set(s));
          if(!criteriaMatches(cv2,crit2))return;
          var sc=sumRng[Math.min(idx,sumRng.length-1)];
          if(!sc)return;
          var sv=cellValueSheet(sc.sheet,sc.r,sc.c,new Set(s));
          if(typeof sv==='number')total+=sv;
          else if(typeof sv==='string'){var pn3=parseCellNumber(sv);if(pn3!=null)total+=pn3}
        });
        return total;
      }
      if(fn==='VLOOKUP'){
        if(node.args.length<3)formulaErr('args');
        var lk=evalNode(node.args[0],s);
        var tbl=node.args[1];
        if(tbl.t!=='range'&&tbl.t!=='srange')formulaErr('range');
        var cells=rangeCells(tbl);
        if(!cells.length)formulaErr('range');
        var colIdx=Math.floor(toNum(evalNode(node.args[2],s)));
        var rTop=Math.min.apply(null,cells.map(function(x){return x.r}));
        var rBot=Math.max.apply(null,cells.map(function(x){return x.r}));
        var cLeft=Math.min.apply(null,cells.map(function(x){return x.c}));
        var cRight=Math.max.apply(null,cells.map(function(x){return x.c}));
        var tblSheet=(tbl.t==='srange')?tbl.sheet:formulaSheet;
        if(colIdx<1||cLeft+colIdx-1>cRight)formulaErr('range');
        var approx=true;
        if(node.args.length>=4){
          var a4=evalNode(node.args[3],s);
          if(a4===false||a4===0)approx=false;
          else if(a4===true||a4===1)approx=true;
          else if(typeof a4==='string')approx=(a4.trim().toUpperCase()==='TRUE');
        }
        var bestRow=-1;
        for(var vr=rTop;vr<=rBot;vr++){
          var first=cellValueSheet(tblSheet,vr,cLeft,new Set(s));
          var match=false;
          if(approx){
            var lkN=(typeof lk==='number')?lk:parseCellNumber(lk);
            var fN=(typeof first==='number')?first:parseCellNumber(first);
            if(lkN!=null&&fN!=null){
              /* Tim gan dung cho cot so da sap xep tang dan (dung quy
                 uoc Excel/Calc). */
              if(Math.abs(fN-lkN)<1e-9){match=true;bestRow=vr;break}
              if(fN<lkN)bestRow=vr;
              else break;
              continue;
            }
            if(typeof lk==='string'&&typeof first==='string'){
              /* Tim gan dung cho cot chu da sap xep theo tieng Viet
                 (chuan doi chieu LibreOffice Calc): hang cuoi co gia tri
                 <= gia tri tim; duoi gia tri nho nhat -> khong thay.
                 Luu y: gan dung chi co nghia khi cot da sap xep; cot
                 chua sap xep thi ket qua chi la lan quet xac dinh nay,
                 Calc co the tra khac vi no tim nhi phan. */
              var cmp=cmpTextVi(String(first).trim(),String(lk).trim());
              if(cmp===0){bestRow=vr;break}
              if(cmp<0){bestRow=vr;continue}
              break;
            }
            if(typeof lk==='number'&&typeof first==='number')match=Math.abs(lk-first)<1e-9;
            else if(typeof lk==='number'&&typeof first==='string')match=parseCellNumber(first)!=null&&Math.abs(lk-parseCellNumber(first))<1e-9;
            else if(typeof lk==='string'&&typeof first==='number')match=parseCellNumber(lk)!=null&&Math.abs(parseCellNumber(lk)-first)<1e-9;
            else match=String(lk).trim().toLowerCase()===String(first).trim().toLowerCase();
          }else{
            if(typeof lk==='number'&&typeof first==='number')match=Math.abs(lk-first)<1e-9;
            else if(typeof lk==='number'&&typeof first==='string')match=parseCellNumber(first)!=null&&Math.abs(lk-parseCellNumber(first))<1e-9;
            else if(typeof lk==='string'&&typeof first==='number')match=parseCellNumber(lk)!=null&&Math.abs(parseCellNumber(lk)-first)<1e-9;
            else match=String(lk).trim().toLowerCase()===String(first).trim().toLowerCase();
          }
          if(match){bestRow=vr;break}
        }
        if(bestRow<0)formulaErr('notfound');
        var got=cellValueSheet(tblSheet,bestRow,cLeft+colIdx-1,new Set(s));
        if(got===''||got==null)return 0;
        return got;
      }
      if(fn==='SUM'||fn==='AVERAGE'||fn==='COUNT'||fn==='MAX'||fn==='MIN'){
        var vals=[];
        node.args.forEach(function(arg){
          expandArg(arg,s).forEach(function(it){
            if(typeof it.v==='number')vals.push(it.v);
            else if(it.fromRef){/* bo qua chu/rong trong tham chieu, giong Excel */}
            else if(typeof it.v==='boolean')vals.push(it.v?1:0);
            else if(typeof it.v==='string'){
              var pn=parseCellNumber(it.v);
              if(pn!=null)vals.push(pn);else formulaErr('notnum');
            }
          });
        });
        if(fn==='SUM')return vals.reduce(function(a,b){return a+b},0);
        if(fn==='COUNT')return vals.length;
        if(fn==='AVERAGE'){if(!vals.length)formulaErr('avg-empty');return vals.reduce(function(a,b){return a+b},0)/vals.length}
        if(fn==='MAX')return vals.length?Math.max.apply(null,vals):0;
        if(fn==='MIN')return vals.length?Math.min.apply(null,vals):0;
      }
      formulaErr('fn');
    }
    formulaErr('node');
  }
  try{
    var ast=parseCmp();
    if(pos!==toks.length)formulaErr('trailing');
    return evalNode(ast,seen);
  }catch(e){
    /* Ban 1.7: thong nhat loi — chia cho 0, vong tron, khong tim thay,
       sai cu phap, ham khong co deu ra #LOI, khong hien lai chu cong
       thuc nhu ban 1.6 nua. */
    return '#LỖI';
  }
}
function evalCell(r,c,seen){
  return evalCellOn(currentSheetName(),r,c,seen);
}
function fmtNumForCell(num,f){
  if(f.num==='money'){
    if(f.dec!=null)return num.toLocaleString('vi-VN',{minimumFractionDigits:f.dec,maximumFractionDigits:f.dec})+' ₫';
    return num.toLocaleString('vi-VN')+' ₫';
  }
  if(f.num==='percent'){
    var pc=num*100;
    if(f.dec!=null)return pc.toLocaleString('vi-VN',{minimumFractionDigits:f.dec,maximumFractionDigits:f.dec})+'%';
    return pc.toLocaleString('vi-VN')+'%';
  }
  if(f.dec!=null)return num.toFixed(f.dec);
  return String(num);
}
function displayCell(r,c){
  /* Ban 1.6: hien thi KHONG lam mat gia tri. O thuong khong dinh dang thi
     hien dung chu nguoi go ("1.000.000" van hien "1.000.000"); chi khi co
     dinh dang tien/%/thap phan moi dich sang so that roi dinh dang lai. */
  if(isMergeSlave(r,c))return '';
  if(showFormulasFlag)return sheetRaw[r]?(sheetRaw[r][c]||''):'' ;
  var raw=sheetRaw[r]?(sheetRaw[r][c]||''):'';
  var f=cellFmt[cellFmtKey(r,c)]||{};
  if(raw.charAt(0)==='='){
    var v=evalCell(r,c);
    if(typeof v==='boolean')return v?'TRUE':'FALSE';
    if(typeof v==='number')return fmtNumForCell(v,f);
    return String(v==null?'':v);
  }
  if(f.num||f.dec!=null){
    var num=parseCellNumber(raw);
    if(num==null)return raw;
    return fmtNumForCell(num,f);
  }
  return raw;
}
function renderSheet(){
  var html='<tr><th></th>';
  for(var c=0;c<SHEET_COLS;c++)html+='<th>'+colName(c)+'</th>';
  html+='</tr>';
  for(var r=0;r<SHEET_ROWS;r++){
    /* Bản 1.4: lọc ẩn hàng không khớp (display:none nhưng vẫn giữ hàng để
       chỉ số hàng không lệch); định dạng ô áp qua style nội tuyến. */
    var hideRow=sheetFilter&&String(sheetRaw[r][sheetFilter.c]||'').toLowerCase().indexOf(sheetFilter.text)<0;
    html+='<tr'+(hideRow?' style="display:none"':'')+'><th>'+(r+1)+'</th>';
    for(var cc=0;cc<SHEET_COLS;cc++){
      if(isMergeSlave(r,cc))continue;
      var mgTop=mergeAtTop(r,cc);
      var spanAttr=mgTop?(' colspan="'+(mgTop.c2-mgTop.c1+1)+'" rowspan="'+(mgTop.r2-mgTop.r1+1)+'"'):'';
      var sel=(selCell.r===r&&selCell.c===cc)?' selected':'';
      var fm=cellFmt[cellFmtKey(r,cc)]||{};
      var st='';
      if(fm.b)st+='font-weight:700;';
      if(fm.i)st+='font-style:italic;';
      if(fm.u)st+='text-decoration:underline;';
      if(fm.color)st+='color:'+fm.color+';';
      if(fm.bg)st+='background-color:'+fm.bg+';';
      else if(dupKeys[cellFmtKey(r,cc)])st+='background-color:#fde68a;';
      if(fm.align)st+='text-align:'+fm.align+';';
      if(fm.valign)st+='vertical-align:'+fm.valign+';';
      var wrapStyle=fm.wrap?'white-space:normal;word-break:break-word;':'';
      var noteMark=fm.note?' ◐':'';
      html+='<td class="cell'+sel+'" data-r="'+r+'" data-c="'+cc+'"'+spanAttr+' style="'+st+'"><div class="cellval" style="'+wrapStyle+'"'+(fm.note?' title="'+esc(fm.note)+'"':'')+'>'+esc(displayCell(r,cc))+noteMark+'</div></td>';
    }
    html+='</tr>';
  }
  $('sheetTable').innerHTML=html;
  renderSheetTabs();
  updateTitleBar();
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
function newSheet(){currentSheetId=null;sheetLocked=false;freezeTopRow=false;gridOn=true;var gw=$('sheetWrap');if(gw)gw.classList.remove('nogrid');$('sheetName').value='';sheetRaw=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});cellFmt={};sheetMerges=[];wbSheets=[{name:'Sheet1',raw:sheetRaw,fmt:cellFmt,merges:[]}];activeSheetIndex=0;dupKeys={};sheetFilter=null;sheetUndo=[];sheetRedo=[];savedSheetSnap=sheetSnapNow();renderSheet();updateStatusBar();}
function saveSheet(silent){
  var name=$('sheetName').value.trim()||('Bảng tính '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentSheetId)name=uniqueName('sheet',name);
  var sheets=loadJSON('vb_sheets_v1',[]);
  syncCurrentToWb();
  var rec={id:currentSheetId||uid(),name:name,raw:sheetRaw,fmt:cellFmt,merges:sheetMerges,sheets:workbookSnapshotSheets(),locked:sheetLocked,freeze:freezeTopRow,grid:gridOn,updatedAt:Date.now(),syncStatus:'local'};
  if(currentSheetId){var s=sheets.find(function(x){return x.id===currentSheetId});if(s){Object.assign(s,rec)}else sheets.unshift(rec)}
  else{currentSheetId=rec.id;sheets.unshift(rec)}
  if(!saveJSON('vb_sheets_v1',sheets)){updateTitleBar();return}
  savedSheetSnap=sheetSnapNow();renderSheetList();renderFileList();if(!silent)alert('Đã lưu bảng tính vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openSheetRec(id){
  var s=loadJSON('vb_sheets_v1',[]).find(function(x){return x.id===id});if(!s)return;
  currentSheetId=s.id;$('sheetName').value=s.name;
  if(s.sheets&&s.sheets.length){
    wbSheets=s.sheets.map(function(sh){return {name:sh.name||'Sheet1',raw:normalizeSheetRaw(sh.raw),fmt:sh.fmt||{},merges:normalizeMerges(sh.merges||[]),locked:!!sh.locked,freeze:!!sh.freeze,grid:sh.grid!==false}});
    loadSheetFromWb(0);
  }else{
    wbSheets=[{name:'Sheet1',raw:normalizeSheetRaw(s.raw),fmt:s.fmt||{},merges:normalizeMerges(s.merges||[]),locked:!!s.locked,freeze:!!s.freeze,grid:s.grid!==false}];
    loadSheetFromWb(0);
  }
  savedSheetSnap=sheetSnapNow();renderSheet();gotoTab('sheet');updateStatusBar();
}
function renderSheetList(){
  var el=$('sheetList');if(!el)return;
  var sheets=loadJSON('vb_sheets_v1',[]);
  el.innerHTML=sheets.length?sheets.map(function(s){
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
/* Ban 1.6: xuat xlsx GIU cong thuc + dinh dang that.
   Cu: chi ghi gia tri hien thi (displayCell) thanh chu/so — mo bang
   LibreOffice/Excel that thi mat cong thuc, mat dinh dang o.
   Moi: moi o ghi dung loai — cong thuc ghi <f> kem gia tri tinh san (<v>),
   so that ghi <v> (o tien "1.000.000 ₫" ghi 1000000, o % "50%" ghi 0,5),
   dinh dang ghi that vao styles.xml (font dam/nghieng/gach chan/mau chu,
   to nen, dinh dang tien/%/thap phan, can le, xuong dong) va gan s="..."
   vao tung o. Dong bang hang dau ghi <pane> that, khoa trang ghi
   <sheetProtection> that. */
function xlsxFormulaForFile(raw){
  /* Dinh dang file xlsx dung dau phay ngan cach doi so; app chap nhan ca
     dau cham phay kieu Viet Nam nen khi xuat phai doi ';' -> ',' ngoai
     chuoi trong ngoac kep. */
  var out='',inStr=false;
  for(var i=0;i<raw.length;i++){
    var ch=raw[i];
    if(ch==='"'){
      if(inStr&&raw[i+1]==='"'){out+='""';i++;continue}
      inStr=!inStr;out+=ch;continue;
    }
    if(ch===';'&&!inStr){out+=',';continue}
    out+=ch;
  }
  return out;
}
function xlsxBlobFromSheets(sheetsIn){
  /* Ban 1.7: xuat workbook nhieu trang. Styles gom chung cho ca workbook
     (font/to nen/dinh dang so) de chi so s="..." dung chung; moi trang co
     sheetData + mergeCells + legacyDrawing (ghi chu) rieng. Cong thuc lien
     trang giu nguyen dang Sheet2!A1 / 'Ten trang'!A1 trong <f>; gia tri
     tinh san <v> lay tu bo may cong thuc cua app (co ho tro lien trang). */
  var sheets=(sheetsIn||[]).map(function(sh,i){
    return {name:sh.name||('Sheet'+(i+1)),raw:normalizeSheetRaw(sh.raw),fmt:sh.fmt||{},
      merges:normalizeMerges(sh.merges||[]),locked:!!sh.locked,freeze:!!sh.freeze,grid:sh.grid!==false};
  });
  if(!sheets.length)sheets=[{name:'Sheet1',raw:normalizeSheetRaw([]),fmt:{},merges:[],locked:false,freeze:false,grid:true}];
  var oldRaw=sheetRaw,oldFmt=cellFmt,oldMerges=sheetMerges,oldWb=wbSheets,oldActive=activeSheetIndex;
  var oldLock=sheetLocked,oldFreeze=freezeTopRow,oldGrid=gridOn;
  try{
    wbSheets=sheets;
    var fonts=[{b:false,i:false,u:false,color:null}];
    var fontIdx={};
    var fills=[null,null];
    var fillIdx={};
    var numFmts=[];
    var numFmtIdx={};
    function hex6(col){
      if(!col)return null;
      var h=String(col).replace('#','').toUpperCase();
      return /^[0-9A-F]{6}$/.test(h)?h:null;
    }
    function getFontId(f){
      var key=JSON.stringify([!!f.b,!!f.i,!!f.u,hex6(f.color)]);
      if(fontIdx[key]!==undefined)return fontIdx[key];
      fonts.push({b:!!f.b,i:!!f.i,u:!!f.u,color:hex6(f.color)});
      fontIdx[key]=fonts.length-1;return fontIdx[key];
    }
    function getFillId(f){
      var bg=hex6(f.bg);
      if(!bg)return 0;
      if(fillIdx[bg]!==undefined)return fillIdx[bg];
      fills.push(bg);fillIdx[bg]=fills.length-1;return fillIdx[bg];
    }
    function getNumFmtId(f){
      var code=null;
      if(f.num==='money')code=(f.dec!=null)?('#,##0.'+new Array(f.dec+1).join('0')+'" ₫"'):'#,##0" ₫"';
      else if(f.num==='percent')code=(f.dec!=null)?('0.'+new Array(f.dec+1).join('0')+'%'):'0%';
      else if(f.dec!=null)code=(f.dec===0)?'0':('0.'+new Array(f.dec+1).join('0'));
      if(!code)return 0;
      if(numFmtIdx[code]!==undefined)return numFmtIdx[code];
      var id=164+numFmts.length;
      numFmts.push({id:id,code:code});numFmtIdx[code]=id;return id;
    }
    var xfIdx={};var xfs=[{numFmtId:0,fontId:0,fillId:0,align:null,valign:null,wrap:false}];
    function getXfId(f){
      var key=JSON.stringify([getNumFmtId(f),getFontId(f),getFillId(f),f.align||null,f.valign||null,!!f.wrap]);
      if(xfIdx[key]!==undefined)return xfIdx[key];
      xfs.push({numFmtId:getNumFmtId(f),fontId:getFontId(f),fillId:getFillId(f),align:f.align||null,valign:f.valign||null,wrap:!!f.wrap});
      xfIdx[key]=xfs.length-1;return xfIdx[key];
    }
    function sheetRowsXml(sh){
      var lastR=0,lastC=0;
      for(var r=0;r<SHEET_ROWS;r++)for(var c=0;c<SHEET_COLS;c++)if(sh.raw[r][c]){lastR=Math.max(lastR,r);lastC=Math.max(lastC,c)}
      var rows='';
      for(var rr=0;rr<=lastR;rr++){
        var rowXml='';
        for(var cc=0;cc<=lastC;cc++){
          var vraw=sh.raw[rr][cc];
          if(!vraw)continue;
          var fm=sh.fmt[cellFmtKey(rr,cc)]||{};
          var xfId=getXfId(fm);
          var sAttr=(xfId>0)?(' s="'+xfId+'"'):'';
          var ref=cellRef(rr,cc);
          if(vraw.charAt(0)==='='){
            var ev=evalCell(rr,cc);
            var fXml='<f>'+esc(xlsxFormulaForFile(vraw.slice(1)))+'</f>';
            if(typeof ev==='number')rowXml+='<c r="'+ref+'"'+sAttr+'>'+fXml+'<v>'+ev+'</v></c>';
            else if(typeof ev==='boolean')rowXml+='<c r="'+ref+'"'+sAttr+'" t="b">'+fXml+'<v>'+(ev?1:0)+'</v></c>';
            else if(typeof ev==='string'&&ev!==''&&ev.charAt(0)!=='#')rowXml+='<c r="'+ref+'"'+sAttr+'" t="str">'+fXml+'<v>'+esc(ev)+'</v></c>';
            else rowXml+='<c r="'+ref+'"'+sAttr+'">'+fXml+'</c>';
          }else{
            var num=parseCellNumber(vraw);
            if(num!=null)rowXml+='<c r="'+ref+'"'+sAttr+'><v>'+num+'</v></c>';
            else rowXml+='<c r="'+ref+'"'+sAttr+'" t="inlineStr"><is><t xml:space="preserve">'+esc(vraw)+'</t></is></c>';
          }
        }
        if(rowXml)rows+='<row r="'+(rr+1)+'">'+rowXml+'</row>';
      }
      return {rows:rows,lastR:lastR,lastC:lastC};
    }
    var sheetXmls=[],commentParts=[];
    sheets.forEach(function(sh,si){
      activeSheetIndex=si;
      sheetRaw=sh.raw;cellFmt=sh.fmt;sheetMerges=sh.merges;
      sheetLocked=sh.locked;freezeTopRow=sh.freeze;gridOn=sh.grid;
      var built=sheetRowsXml(sh);
      var views='<sheetViews><sheetView workbookViewId="0" showGridLines="'+(gridOn?'1':'0')+'">';
      if(freezeTopRow)views+='<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
      else views+='<selection activeCell="A1" sqref="A1"/>';
      views+='</sheetView></sheetViews>';
      var prot=sheetLocked?'<sheetProtection sheet="1" objects="1" scenarios="1"/>':'';
      var mergeXml='';
      if(sh.merges.length){
        mergeXml='<mergeCells count="'+sh.merges.length+'">'+sh.merges.map(function(m){
          return '<mergeCell ref="'+cellRef(m.r1,m.c1)+':'+cellRef(m.r2,m.c2)+'"/>';
        }).join('')+'</mergeCells>';
      }
      /* Ghi chu o dang legacy comments (Excel/LibreOffice hien duoc): moi
         o co fmt.note thanh 1 comment + 1 hinh ve dinh vi trong drawing. */
      var notes=[];
      Object.keys(sh.fmt).forEach(function(k){
        var f=sh.fmt[k];
        if(f&&f.note){
          var pp=k.split(',');
          notes.push({r:+pp[0],c:+pp[1],text:String(f.note)});
        }
      });
      notes.sort(function(a,b){return (a.r-b.r)||(a.c-b.c)});
      var legacyXml='';
      var cmt=null;
      if(notes.length){
        legacyXml='<legacyDrawing r:id="rIdDraw"/>';
        var cmtList=notes.map(function(n){
          return '<comment ref="'+cellRef(n.r,n.c)+'" authorId="0"><text><t xml:space="preserve">'+esc(n.text)+'</t></text></comment>';
        }).join('');
        var commentsXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Văn phòng của Bùi</author></authors><commentList>'+cmtList+'</commentList></comments>';
        var anchors=notes.map(function(n,ni){
          return '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>'+n.c+'</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>'+n.r+'</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>'+(n.c+2)+'</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>'+(n.r+2)+'</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="'+(ni+1)+'" name="Ghi chú '+(ni+1)+'"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFE1"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr><xdr:clientData/></xdr:sp></xdr:twoCellAnchor>';
        }).join('');
        var drawingXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'+anchors+'</xdr:wsDr>';
        var relsXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdCmt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments'+(si+1)+'.xml"/><Relationship Id="rIdDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'+(si+1)+'.xml"/></Relationships>';
        cmt={commentsXml:commentsXml,drawingXml:drawingXml,relsXml:relsXml};
      }
      commentParts.push(cmt);
      var dim='A1:'+cellRef(built.lastR,built.lastC);
      sheetXmls.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="'+dim+'"/>'+views+'<sheetData>'+built.rows+'</sheetData>'+prot+mergeXml+legacyXml+'</worksheet>');
    });
    var fontsXml=fonts.map(function(fn){
      var x='<font><sz val="11"/><name val="Carlito"/>';
      if(fn.b)x+='<b/>';if(fn.i)x+='<i/>';if(fn.u)x+='<u/>';
      if(fn.color)x+='<color rgb="FF'+fn.color+'"/>';
      return x+'</font>';
    }).join('');
    var fillsXml=fills.map(function(bg,i){
      if(i===0)return '<fill><patternFill patternType="none"/></fill>';
      if(i===1)return '<fill><patternFill patternType="gray125"/></fill>';
      return '<fill><patternFill patternType="solid"><fgColor rgb="FF'+bg+'"/><bgColor rgb="FF'+bg+'"/></patternFill></fill>';
    }).join('');
    var numFmtsXml=numFmts.length?('<numFmts count="'+numFmts.length+'">'+numFmts.map(function(nf){return '<numFmt numFmtId="'+nf.id+'" formatCode="'+esc(nf.code)+'"/>'}).join('')+'</numFmts>'):'';
    var xfsXml=xfs.map(function(xf){
      var a='<xf numFmtId="'+xf.numFmtId+'" fontId="'+xf.fontId+'" fillId="'+xf.fillId+'" borderId="0" xfId="0"';
      if(xf.numFmtId!==0)a+=' applyNumberFormat="1"';
      if(xf.fontId!==0)a+=' applyFont="1"';
      if(xf.fillId!==0)a+=' applyFill="1"';
      if(xf.align||xf.valign||xf.wrap)a+=' applyAlignment="1"';
      a+='>';
      if(xf.align||xf.valign||xf.wrap){
        a+='<alignment';
        if(xf.align)a+=' horizontal="'+xf.align+'"';
        if(xf.valign)a+=' vertical="'+xf.valign+'"';
        if(xf.wrap)a+=' wrapText="1"';
        a+='/>';
      }
      return a+'</xf>';
    }).join('');
    var stylesXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      +numFmtsXml+'<fonts count="'+fonts.length+'">'+fontsXml+'</fonts>'
      +'<fills count="'+fills.length+'">'+fillsXml+'</fills>'
      +'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
      +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      +'<cellXfs count="'+xfs.length+'">'+xfsXml+'</cellXfs>'
      +'<cellStyles><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    var sheetsXml=sheets.map(function(sh,si){
      return '<sheet name="'+esc(sh.name)+'" sheetId="'+(si+1)+'" r:id="rIdS'+(si+1)+'"/>';
    }).join('');
    var wbRels=sheets.map(function(sh,si){
      return '<Relationship Id="rIdS'+(si+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(si+1)+'.xml"/>';
    }).join('')+'<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    var typesExtra=sheets.map(function(sh,si){
      return '<Override PartName="/xl/worksheets/sheet'+(si+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }).join('');
    commentParts.forEach(function(cmt,si){
      if(!cmt)return;
      typesExtra+='<Override PartName="/xl/comments'+(si+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'
        +'<Override PartName="/xl/drawings/drawing'+(si+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.drawing+xml"/>';
    });
    var entries=[
      {name:'[Content_Types].xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'+typesExtra+'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'},
      {name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'},
      {name:'xl/workbook.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheetsXml+'</sheets></workbook>'},
      {name:'xl/_rels/workbook.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+wbRels+'</Relationships>'},
      {name:'xl/styles.xml',data:stylesXml}
    ];
    sheetXmls.forEach(function(xml,si){
      entries.push({name:'xl/worksheets/sheet'+(si+1)+'.xml',data:xml});
      if(commentParts[si]){
        entries.push({name:'xl/worksheets/_rels/sheet'+(si+1)+'.xml.rels',data:commentParts[si].relsXml});
        entries.push({name:'xl/comments'+(si+1)+'.xml',data:commentParts[si].commentsXml});
        entries.push({name:'xl/drawings/drawing'+(si+1)+'.xml',data:commentParts[si].drawingXml});
      }
    });
    return zipStored(entries);
  }finally{
    sheetRaw=oldRaw;cellFmt=oldFmt;sheetMerges=oldMerges;wbSheets=oldWb;activeSheetIndex=oldActive;
    sheetLocked=oldLock;freezeTopRow=oldFreeze;gridOn=oldGrid;
  }
}
function xlsxBlobFromRaw(raw,opts){
  opts=opts||{};
  return xlsxBlobFromSheets([{name:'Sheet1',raw:raw,fmt:opts.fmt||cellFmt,merges:opts.merges||sheetMerges,
    locked:opts.locked!==undefined?opts.locked:sheetLocked,
    freeze:opts.freeze!==undefined?opts.freeze:freezeTopRow,
    grid:opts.grid!==undefined?opts.grid:gridOn}]);
}
function exportXlsx(){
  if(!confirmExportWarnings(exportWarnings('excel')))return;
  downloadBlob(xlsxBlobFromSheets(workbookSnapshotSheets()),($('sheetName').value.trim()||'bang-tinh')+'.xlsx');
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
/* Ban 1.6: giai nen zip tra ve MOI entry (can cho nhap docx/xlsx/pptx giu
   anh/dinh dang). Entry stored lay thang; entry deflate dung
   DecompressionStream cua trinh duyet (Node kiem chung co san API nay). */
function unzipAll(buf){
  return new Promise(function(resolve,reject){
    try{
      var bytes=new Uint8Array(buf);var dv=new DataView(buf);
      function readStr(o,len){var s='';for(var i=0;i<len;i++)s+=String.fromCharCode(bytes[o+i]);return s}
      var eocd=-1;
      for(var i=bytes.length-22;i>=Math.max(0,bytes.length-65558);i--){
        if(dv.getUint32(i,true)===0x06054b50){eocd=i;break}
      }
      if(eocd<0){reject(new Error('Không tìm thấy thư mục trung tâm (central directory) của file zip'));return}
      var total=dv.getUint16(eocd+10,true);var cdOff=dv.getUint32(eocd+16,true);
      var metas=[];var p=cdOff;
      for(var n=0;n<total;n++){
        if(dv.getUint32(p,true)!==0x02014b50)break;
        var method=dv.getUint16(p+10,true),csize=dv.getUint32(p+20,true);
        var nlen=dv.getUint16(p+28,true),elen=dv.getUint16(p+30,true),clen=dv.getUint16(p+32,true);
        var lho=dv.getUint32(p+42,true);
        var name=new TextDecoder().decode(bytes.slice(p+46,p+46+nlen));
        var ln=dv.getUint16(lho+26,true),le=dv.getUint16(lho+28,true);
        var start=lho+30+ln+le;
        metas.push({name:name,method:method,data:bytes.slice(start,start+csize)});
        p+=46+nlen+elen+clen;
      }
      var out={};var idx=0;
      function next(){
        if(idx>=metas.length){resolve(out);return}
        var m=metas[idx++];
        if(m.method===0){out[m.name]=m.data;next();return}
        if(m.method===8){
          if(typeof DecompressionStream!=='undefined'){
            new Response(new Blob([m.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer()
              .then(function(ab){out[m.name]=new Uint8Array(ab);next()}).catch(reject);
            return;
          }
          if(typeof globalThis!=='undefined'&&globalThis.__inflateRaw){
            try{out[m.name]=globalThis.__inflateRaw(m.data);next()}catch(e){reject(e)}
            return;
          }
          reject(new Error('Trình duyệt này không giải nén được entry zip nén (cần DecompressionStream).'));
          return;
        }
        reject(new Error('Không giải nén được entry zip với phương pháp '+m.method));
      }
      next();
    }catch(e){reject(e)}
  });
}
function zipEntryText(e){
  if(e==null)return '';
  if(typeof e==='string')return e;
  return new TextDecoder().decode(e);
}
function xmlUnesc(s){
  return String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
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
    f.arrayBuffer().then(unzipAll).then(function(entries){
      var docXml=zipEntryText(entries['word/document.xml']);
      var relsXml=zipEntryText(entries['word/_rels/document.xml.rels']);
      var numXml=zipEntryText(entries['word/numbering.xml']);
      var media={};
      Object.keys(entries).forEach(function(k){
        if(k.indexOf('word/media/')===0){
          var ext=k.split('.').pop().toLowerCase();
          var mime=(ext==='png')?'image/png':(ext==='gif')?'image/gif':(ext==='webp')?'image/webp':'image/jpeg';
          media[k]='data:'+mime+';base64,'+bytesToBase64(entries[k]);
        }
      });
      var partByType=function(keyword){var re=new RegExp('<Relationship[^>]*Type="[^"]*'+keyword+'[^"]*"[^>]*Target="([^"]+)"');var mm=re.exec(relsXml||'');if(!mm)return '';var t=mm[1].replace(/^\/+/, '');if(t.indexOf('word/')!==0)t='word/'+t;return t;};
      var hdrPath=partByType('header'),ftrPath=partByType('footer'),cmtPath=partByType('comments');
      var parsed=parseDocxXmlToHtml(docXml,relsXml,media,numXml,{headerXml:zipEntryText(entries[hdrPath]||''),footerXml:zipEntryText(entries[ftrPath]||''),commentsXml:zipEntryText(entries[cmtPath]||'')});
      addImportedDoc(f.name.replace(/\.docx$/i,''),parsed.html,parsed.bg);
      $('importStatus').textContent='Đã nhập docx: giữ tiêu đề, định dạng chữ, bảng, ảnh, danh sách, theo dõi thay đổi, đầu trang/chân trang, ghi chú Word, mục lục (dạng trường, quét lại khi xuất).';
    }).catch(function(e){
      $('importStatus').textContent='Nhập docx chưa xong: '+e.message+' — file có thể không phải docx chuẩn.';
    });
  }else if(lower.endsWith('.xlsx')){
    f.arrayBuffer().then(unzipAll).then(function(entries){
      var parsedWb=parseXlsxToWorkbook(entries);
      newSheet();
      wbSheets=parsedWb.sheets.map(function(sh){return {name:sh.name,raw:sh.raw,fmt:sh.fmt,merges:sh.merges||[],locked:false,freeze:false,grid:true}});
      if(!wbSheets.length)wbSheets=[{name:'Sheet1',raw:normalizeSheetRaw([]),fmt:{},merges:[],locked:false,freeze:false,grid:true}];
      loadSheetFromWb(0);
      $('sheetName').value=f.name.replace(/\.xlsx$/i,'');
      renderSheet();gotoTab('sheet');
      $('importStatus').textContent='Đã nhập XLSX: '+parsedWb.sheets.length+' trang tính, '+parsedWb.totalUsed+' ô có nội dung (giữ công thức + định dạng cơ bản + gộp ô + ghi chú ô).'
        +(parsedWb.totalSkipped?(' Bỏ qua '+parsedWb.totalSkipped+' ô ngoài lưới 26 cột × 100 hàng.'): '')
        +' Bấm “Lưu vào máy” để giữ lại.';
    }).catch(function(e){
      $('importStatus').textContent='Nhập xlsx chưa xong: '+e.message;
    });
  }else if(lower.endsWith('.pptx')){
    f.arrayBuffer().then(unzipAll).then(function(entries){
      var parsed=parsePptxToDeck(entries);
      deckSlides=parsed.slides;deckSize=parsed.size;deckTransition='none';deckLocked=false;currentDeckId=null;currentSlide=0;
      $('deckName').value=f.name.replace(/\.pptx$/i,'');
      renderSlideEditor();gotoTab('slides');
      $('importStatus').textContent='Đã nhập PPTX: '+parsed.slides.length+' slide theo đúng thứ tự, giữ tiêu đề, nội dung, mọi ảnh mỗi slide, bảng đầu mỗi slide, slide ẩn. Biểu đồ, hình dạng, ghi chú, hiệu ứng không nhập được. Bấm “Lưu vào máy” để giữ lại.';
    }).catch(function(e){
      $('importStatus').textContent='Nhập pptx chưa xong: '+e.message;
    });
  }else if(lower.endsWith('.pdf')){
    $('importStatus').textContent='File PDF không lưu vào danh sách văn bản. Sang mục “PDF & OCR & Chụp ảnh”, chọn lại file này để xem, chú thích và OCR.';
    gotoTab('pdf');
  }else{$('importStatus').textContent='Định dạng chưa hỗ trợ nhập: chỉ txt, docx, xlsx, pptx, csv, pdf.'}
});
/* Ban 1.7: nhap xlsx NHIEU TRANG (1.6 chi lay trang dau, toi da 10x30).
   Giu: cong thuc (ghi lai dang =...), gia tri so, chu; dinh dang co ban
   (dam/nghieng/gach chan/mau chu/to nen/tien/%/thap phan/can le/xuong
   dong); gop o (mergeCell); ghi chu o (comments legacy). Luoi app nay la
   26 cot (A-Z) x 100 hang; o ngoai luoi van dem vao skipped. Bieu do
   khong nhap duoc. Dinh dang nhap vao duoc sao chep (khong dung chung
   doi tuong) de sua 1 o khong lam lan sang o khac cung kieu. */
function parseXlsxSharedAndStyles(entries){
  var shared=[];
  var ssXml=zipEntryText(entries['xl/sharedStrings.xml']||'');
  if(ssXml){
    var siRe=/<si>([\s\S]*?)<\/si>/g;var siM;
    while((siM=siRe.exec(ssXml))){
      var txt='';var tRe=/<t[^>]*>([^<]*)<\/t>/g;var tM;
      while((tM=tRe.exec(siM[1])))txt+=xmlUnesc(tM[1]);
      shared.push(txt);
    }
  }
  var stylesXml=zipEntryText(entries['xl/styles.xml']||'');
  var numFmtCode={9:'0%',10:'0.00%',14:'',44:'#,##0',164:'#,##0'};
  var nfRe=/<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;var nfM;
  while((nfM=nfRe.exec(stylesXml)))numFmtCode[+nfM[1]]=xmlUnesc(nfM[2]);
  function splitTags(xml,tag){
    var out=[];var re=new RegExp('<'+tag+'(?:\\s[^>]*)?>([\\s\\S]*?)<\/'+tag+'>','g');var m;
    while((m=re.exec(xml)))out.push(m[0]);
    return out;
  }
  var fonts=splitTags(stylesXml,'font').map(function(fx){
    var col=/<color[^>]*rgb="([0-9A-Fa-f]+)"/.exec(fx);
    var rgb=col?col[1]:null;
    if(rgb&&rgb.length===8)rgb=rgb.slice(2);
    return {b:/<b[\s\/>]/.test(fx),i:/<i[\s\/>]/.test(fx),u:/<u[\s\/>]/.test(fx),color:rgb};
  });
  var fills=splitTags(stylesXml,'fill').map(function(fx){
    if(/patternType="none"/.test(fx)||/patternType="gray125"/.test(fx))return null;
    var col=/<fgColor[^>]*rgb="([0-9A-Fa-f]+)"/.exec(fx);
    var rgb=col?col[1]:null;
    if(rgb&&rgb.length===8)rgb=rgb.slice(2);
    if(!rgb||rgb==='000000')return null;
    return rgb;
  });
  var xfs=[];
  var xfSection=/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
  if(xfSection){
    var xfRe=/<xf\s([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g;var xfM;
    while((xfM=xfRe.exec(xfSection[1]))){
      var attrs=xfM[1];
      var numId=+( /numFmtId="(\d+)"/.exec(attrs)||[0,0])[1];
      var fontId=+( /fontId="(\d+)"/.exec(attrs)||[0,0])[1];
      var fillId=+( /fillId="(\d+)"/.exec(attrs)||[0,0])[1];
      var inner=xfM[2]||'';
      var al=/<alignment[^>]*horizontal="([^"]+)"/.exec(inner);
      var va=/<alignment[^>]*vertical="([^"]+)"/.exec(inner);
      var wr=/<alignment[^>]*wrapText="1"/.test(inner);
      var f={};
      var code=numFmtCode[numId];
      if(code!==undefined&&code!==''){
        if(code.indexOf('%')>=0){
          f.num='percent';
          var decM=/0\.(0+)%/.exec(code);
          if(decM)f.dec=decM[1].length;
        }else if(code.indexOf('₫')>=0||/VND/i.test(code)){
          f.num='money';
          var decM2=/0\.(0+)/.exec(code);
          if(decM2)f.dec=decM2[1].length;
        }else{
          var decM3=/^0\.(0+)$/.exec(code);
          if(decM3)f.dec=decM3[1].length;
        }
      }
      var fn=fonts[fontId];
      if(fn){if(fn.b)f.b=true;if(fn.i)f.i=true;if(fn.u)f.u=true;if(fn.color)f.color='#'+fn.color}
      var fl=fills[fillId];
      if(fl)f.bg='#'+fl;
      if(al&&['left','center','right'].indexOf(al[1])>=0)f.align=al[1];
      if(va&&['top','middle','bottom'].indexOf(va[1])>=0)f.valign=va[1];
      if(wr)f.wrap=true;
      xfs.push(f);
    }
  }
  return {shared:shared,xfs:xfs};
}
function parseXlsxOneSheet(entries,sheetPath,shared,xfs){
  var raw=Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')});
  var fmt={};var skipped=0,used=0;
  var shXml=zipEntryText(entries[sheetPath]||'');
  var cRe=/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;var cM;
  while((cM=cRe.exec(shXml))){
    var attrs=cM[1];var inner=cM[2]||'';
    var rM=/r="([A-Z]+[0-9]+)"/.exec(attrs);
    if(!rM)continue;
    var rf=parseRef(rM[1]);
    if(!rf||rf.r>=SHEET_ROWS){skipped++;continue}
    if(rf.c>=SHEET_COLS){skipped++;continue}
    var tM=/t="([^"]+)"/.exec(attrs);
    var t=tM?tM[1]:'n';
    var fM=/<f[^>]*>([^<]*)<\/f>/.exec(inner);
    var vM=/<v>([^<]*)<\/v>/.exec(inner);
    var val='';
    if(fM)val='='+xmlUnesc(fM[1]);
    else if(t==='s'&&vM)val=shared[parseInt(vM[1],10)]||'';
    else if(t==='inlineStr'){
      var isM=/<is>([\s\S]*?)<\/is>/.exec(inner);
      var tt='';
      if(isM){var tRe2=/<t[^>]*>([^<]*)<\/t>/g;var tM2;while((tM2=tRe2.exec(isM[1])))tt+=xmlUnesc(tM2[1])}
      val=tt;
    }else if(t==='b'&&vM)val=(vM[1]==='1')?'TRUE':'FALSE';
    else if(vM)val=xmlUnesc(vM[1]);
    if(val!=='')used++;
    raw[rf.r][rf.c]=val;
    var sM=/s="(\d+)"/.exec(attrs);
    if(sM&&xfs[+sM[1]]&&Object.keys(xfs[+sM[1]]).length)fmt[rf.r+','+rf.c]=Object.assign({},xfs[+sM[1]]);
  }
  var merges=[];
  var mgRe=/<mergeCell[^>]*ref="([^"]+)"/g;var mgM;
  while((mgM=mgRe.exec(shXml))){
    var parts=mgM[1].split(':');
    var a=parseRef(parts[0]||''),b=parts.length>1?parseRef(parts[1]):a;
    if(a&&b)merges.push({r1:Math.min(a.r,b.r),c1:Math.min(a.c,b.c),r2:Math.max(a.r,b.r),c2:Math.max(a.c,b.c)});
  }
  merges=normalizeMerges(merges);
  /* Ghi chu o: tim file comments qua rels cua trang tinh nay. */
  try{
    var relsPath=sheetPath.replace('xl/worksheets/','xl/worksheets/_rels/')+'.rels';
    var relsXml=zipEntryText(entries[relsPath]||'');
    var cmtTarget='';
    var relTagRe=/<Relationship\b[^>]*>/g;var relTag;
    while((relTag=relTagRe.exec(relsXml))){
      if(relTag[0].indexOf('comments')>=0){
        var tgM=/Target="([^"]+)"/.exec(relTag[0]);
        if(tgM)cmtTarget=tgM[1];
        break;
      }
    }
    if(cmtTarget){
      var ct=cmtTarget.replace(/^\.\.\//,'').replace(/^\/+/,'');
      if(ct.indexOf('xl/')!==0)ct='xl/'+ct;
      var cmtXml=zipEntryText(entries[ct]||'');
      var cmRe=/<comment\b[^>]*ref="([^"]+)"[^>]*>([\s\S]*?)<\/comment>/g;var cmM;
      while((cmM=cmRe.exec(cmtXml))){
        var rf2=parseRef(cmM[1]);
        if(!rf2||rf2.r>=SHEET_ROWS||rf2.c>=SHEET_COLS)continue;
        var txt='';
        var tRe3=/<t[^>]*>([^<]*)<\/t>/g;var tM3;
        while((tM3=tRe3.exec(cmM[2])))txt+=xmlUnesc(tM3[1]);
        if(!txt)continue;
        var key=rf2.r+','+rf2.c;
        fmt[key]=Object.assign({},fmt[key]||{});
        fmt[key].note=txt;
      }
    }
  }catch(e){}
  return {raw:raw,fmt:fmt,merges:merges,used:used,skipped:skipped};
}
function parseXlsxToWorkbook(entries){
  function relsMap(xml){
    var m={};var re=/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;var mm;
    while((mm=re.exec(xml)))m[mm[1]]=mm[2];
    return m;
  }
  function normTarget(t,base){
    if(!t)return '';
    t=t.replace(/^\/+/,'');
    if(t.indexOf('xl/')===0)return t;
    return base+t;
  }
  var wbRels=relsMap(zipEntryText(entries['xl/_rels/workbook.xml.rels']||''));
  var wbXml=zipEntryText(entries['xl/workbook.xml']||'');
  var list=[];
  var tagRe=/<sheet\b[^>]*>/g;var tagM;
  while((tagM=tagRe.exec(wbXml))){
    var nm=/name="([^"]+)"/.exec(tagM[0]);
    var rid=/r:id="([^"]+)"/.exec(tagM[0]);
    if(nm)list.push({name:xmlUnesc(nm[1]),path:rid?normTarget(wbRels[rid[1]]||'worksheets/sheet1.xml','xl/'):'xl/worksheets/sheet1.xml'});
  }
  if(!list.length)list=[{name:'Sheet1',path:'xl/worksheets/sheet1.xml'}];
  var sd=parseXlsxSharedAndStyles(entries);
  var sheets=list.map(function(it){
    var one=parseXlsxOneSheet(entries,it.path,sd.shared,sd.xfs);
    return {name:it.name,raw:one.raw,fmt:one.fmt,merges:one.merges,used:one.used,skipped:one.skipped};
  });
  var totalUsed=0,totalSkipped=0;
  sheets.forEach(function(sh){totalUsed+=sh.used;totalSkipped+=sh.skipped});
  return {sheets:sheets,totalUsed:totalUsed,totalSkipped:totalSkipped};
}
function parseXlsxToSheet(entries){
  var wb=parseXlsxToWorkbook(entries);
  var f=wb.sheets[0]||{raw:Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')}),fmt:{},name:'Sheet1',used:0,skipped:0};
  return {raw:f.raw,fmt:f.fmt,sheetName:f.name,used:f.used,skipped:f.skipped};
}
/* Ban 1.7: nhap docx GIU dinh dang + dau/chan trang + ghi chu + muc luc.
   Giu: tieu de Heading1-3, dam/nghieng/gach chan/gach ngang, mau chu,
   co chu, font, to sang, thu le, vien doan, bang, anh (qua media + rels),
   danh sach gach dau dong/danh so (qua numbering.xml), theo doi thay doi
   (w:ins/w:del thanh the ins/del cua app), lien ket, mau nen trang,
   dau trang/chan trang (thanh div doc-header/doc-footer), ghi chu Word
   that (thanh doan comment-para), muc luc truong TOC (thanh div doc-toc
   voi cac muc da luu san; xuat lai se quet lai tu tieu de).
   Khong nhap duoc: kieu doan phuc tap; anh trong dau/chan trang co the
   khong dung anh goc vi rel anh cua header/footer khac rel cua tai lieu. */
function parseDocxXmlToHtml(docXml,relsXml,mediaByName,numberingXml,extra){
  extra=extra||{};
  mediaByName=mediaByName||{};
  var rels={};
  var relRe=/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;var relM;
  while((relM=relRe.exec(relsXml||''))){
    var t=relM[2].replace(/^\/+/,'');
    if(t.indexOf('word/')!==0)t='word/'+t;
    rels[relM[1]]=t;
  }
  var numFmtById={};
  if(numberingXml){
    var absFmt={};
    var absRe=/<w:abstractNum[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g;var absM;
    while((absM=absRe.exec(numberingXml))){
      var fm=/<w:numFmt[^>]*w:val="([^"]+)"/.exec(absM[2]);
      absFmt[absM[1]]=fm?fm[1]:'decimal';
    }
    var numRe=/<w:num[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g;var numM;
    while((numM=numRe.exec(numberingXml))){
      var av=/<w:abstractNumId[^>]*w:val="(\d+)"/.exec(numM[2]);
      if(av)numFmtById[numM[1]]=absFmt[av[1]]||'decimal';
    }
  }
  var hlMap={yellow:'#FFFF00',green:'#00FF00',cyan:'#00FFFF',magenta:'#FF00FF',blue:'#0000FF',red:'#FF0000',darkYellow:'#808000',darkGreen:'#008000',darkCyan:'#008080',darkMagenta:'#800080',darkBlue:'#000080',darkRed:'#800000',darkGray:'#808080',lightGray:'#C0C0C0',black:'#000000'};
  function runHtml(runXml){
    var rPr=/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(runXml);
    var pr=rPr?rPr[1]:'';
    var bold=/<w:b[\s\/>]/.test(pr),ital=/<w:i[\s\/>]/.test(pr);
    var und=/<w:u[\s\/>]/.test(pr),strk=/<w:strike[\s\/>]/.test(pr);
    var colM=/<w:color[^>]*w:val="([0-9A-Fa-f]+)"/.exec(pr);
    var szM=/<w:sz[^>]*w:val="(\d+)"/.exec(pr);
    var fontM=/<w:rFonts[^>]*w:ascii="([^"]+)"/.exec(pr);
    var hlM=/<w:highlight[^>]*w:val="([^"]+)"/.exec(pr);
    var shdM=/<w:shd[^>]*w:fill="([0-9A-Fa-f]+)"/.exec(pr);
    var style='';
    if(fontM)style+='font-family:'+fontM[1]+';';
    if(szM)style+='font-size:'+(+szM[1]/2)+'pt;';
    if(colM)style+='color:#'+colM[1]+';';
    if(hlM&&hlMap[hlM[1]])style+='background-color:'+hlMap[hlM[1]]+';';
    else if(shdM&&shdM[1].toUpperCase()!=='AUTO')style+='background-color:#'+shdM[1]+';';
    var inner='';
    var blipM=/<a:blip[^>]*r:embed="([^"]+)"/.exec(runXml);
    if(blipM){
      var part=rels[blipM[1]];
      var dataUrl=part?mediaByName[part]:null;
      if(dataUrl)inner+='<img src="'+dataUrl+'" alt="Ảnh chèn">';
    }
    var tRe=/<w:t[^>]*>([^<]*)<\/w:t>/g;var tM2;
    while((tM2=tRe.exec(runXml)))inner+=esc(xmlUnesc(tM2[1]));
    inner=inner.replace(/<w:tab\/>/g,' ');
    if(/<w:tab\/>/.test(runXml)&&!inner)inner=' ';
    if(/<w:br[\s\/>]/.test(runXml))inner+='<br>';
    var core=inner;
    if(strk)core='<s>'+core+'</s>';
    if(und)core='<u>'+core+'</u>';
    if(ital)core='<i>'+core+'</i>';
    if(bold)core='<b>'+core+'</b>';
    if(style)core='<span style="'+style+'">'+core+'</span>';
    return core;
  }
  function paraObj(paraXml){
    var pPrM=/<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(paraXml);
    var pr=pPrM?pPrM[1]:'';
    var heading=null;
    var stM=/<w:pStyle[^>]*w:val="Heading([123])"/.exec(pr);
    if(stM)heading='h'+stM[1];
    var align='left';
    var jcM=/<w:jc[^>]*w:val="(center|right|both)"/.exec(pr);
    if(jcM)align=(jcM[1]==='both')?'justify':jcM[1];
    var bullet=null;
    var numM2=/<w:numId[^>]*w:val="(\d+)"/.exec(pr);
    if(numM2){
      var fmt=numFmtById[numM2[1]];
      if(!fmt)fmt=(numM2[1]==='1')?'bullet':'decimal';
      bullet=(fmt==='bullet')?'bullet':'decimal';
    }
    var indLeft=0,indFirst=0;
    var indM=/<w:ind[^>]*w:left="(-?\d+)"/.exec(pr);
    if(indM)indLeft=Math.round(+indM[1]/15);
    var firstM=/<w:ind[^>]*w:firstLine="(-?\d+)"/.exec(pr);
    if(firstM)indFirst=Math.round(+firstM[1]/15);
    var border=/<w:pBdr>/.test(pr);
    var runs='';
    var segRe=/<w:hyperlink(\s[^>]*)?>([\s\S]*?)<\/w:hyperlink>|<w:ins(\s[^>]*)?>([\s\S]*?)<\/w:ins>|<w:del(\s[^>]*)?>([\s\S]*?)<\/w:del>|<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;var segM;
    while((segM=segRe.exec(paraXml))){
      if(segM[2]!==undefined){
        var href=null;
        var idM=/r:id="([^"]+)"/.exec(segM[1]||'');
        if(idM){
          var tgt=rels[idM[1]];
          if(tgt)tgt=tgt.replace(/^word\//,'');
          if(tgt&&/^https?:\/\//.test(tgt))href=tgt;
          else if(tgt&&tgt.indexOf('http')===0)href=tgt;
        }
        var innerRuns='';
        var rRe2=/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;var rM2;
        while((rM2=rRe2.exec(segM[2])))innerRuns+=runHtml(rM2[0]);
        runs+=href?('<a href="'+esc(href)+'" target="_blank" rel="noopener">'+innerRuns+'</a>'):innerRuns;
      }else if(segM[4]!==undefined){
        var innerRuns2='';
        var rRe3=/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;var rM3;
        while((rM3=rRe3.exec(segM[4])))innerRuns2+=runHtml(rM3[0]);
        runs+='<ins class="track-ins">'+innerRuns2+'</ins>';
      }else if(segM[6]!==undefined){
        var delTxt='';
        var dRe=/<w:delText[^>]*>([^<]*)<\/w:delText>/g;var dM;
        while((dM=dRe.exec(segM[6])))delTxt+=esc(xmlUnesc(dM[1]));
        runs+='<del class="track-del">'+delTxt+'</del>';
      }else if(segM[7]!==undefined){
        runs+=runHtml(segM[7]);
      }
    }
    return {runs:runs,heading:heading,align:align,bullet:bullet,indLeft:indLeft,indFirst:indFirst,border:border};
  }
  function paraHtml(p){
    var tag=p.heading||'p';
    var style='';
    if(p.align!=='left')style+='text-align:'+p.align+';';
    if(p.indLeft)style+='margin-left:'+p.indLeft+'px;';
    if(p.indFirst)style+='text-indent:'+p.indFirst+'px;';
    if(p.border)style+='border:1px solid #999;';
    return '<'+tag+(style?' style="'+style+'"':'')+'>'+(p.runs||'<br>')+'</'+tag+'>';
  }
  /* Ban 1.7: ghi chu that (comments.xml) thanh comment-para; dau/chan
     trang (header/footer parts) thanh div doc-header/doc-footer de xuat
     lai that; muc luc truong TOC thanh div doc-toc. */
  var commentMap={};
  if(extra.commentsXml){
    var cmRe2=/<w:comment\b[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/g;var cmM2;
    while((cmM2=cmRe2.exec(extra.commentsXml))){
      var ctext='';
      var ctRe=/<w:t[^>]*>([^<]*)<\/w:t>/g;var ctM;
      while((ctM=ctRe.exec(cmM2[2])))ctext+=xmlUnesc(ctM[1]);
      commentMap[cmM2[1]]=ctext;
    }
  }
  function headerFooterHtml(partXml,cls){
    if(!partXml)return '';
    var texts=[];
    var hpRe=/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;var hpM;
    while((hpM=hpRe.exec(partXml))){
      var pp=paraObj(hpM[0]);
      var tx=(pp.runs||'').replace(/<[^>]*>/g,'').trim();
      if(tx)texts.push(tx);
    }
    if(!texts.length)return '';
    return '<div class="'+cls+'">'+texts.join(' ')+'</div>';
  }
  var headerHtml=headerFooterHtml(extra.headerXml||'','doc-header');
  var footerHtml=headerFooterHtml(extra.footerXml||'','doc-footer');
  var bodyM=/<w:body>([\s\S]*)<\/w:body>/.exec(docXml||'');
  var body=bodyM?bodyM[1]:(docXml||'');
  var html='';var listKind=null;
  var commentQueue=[];
  function flushList(){if(listKind){html+='</'+listKind+'>';listKind=null}if(commentQueue.length){html+=commentQueue.join('');commentQueue=[]}}
  var tocMode=0,tocEntries=[];
  var topRe=/<w:tbl>[\s\S]*?<\/w:tbl>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;var topM;
  while((topM=topRe.exec(body))){
    var chunk=topM[0];
    var isBegin=/fldCharType="begin"/.test(chunk);
    var isSep=/fldCharType="separate"/.test(chunk);
    var isEnd=/fldCharType="end"/.test(chunk);
    if(tocMode===1){
      if(isEnd){flushList();html+='<div class="doc-toc"><h2>Mục lục</h2></div>';tocMode=0;}
      else if(isSep){tocMode=2;tocEntries=[];}
      continue;
    }
    if(tocMode===2){
      if(isEnd){flushList();html+='<div class="doc-toc"><h2>Mục lục</h2>'+tocEntries.join('')+'</div>';tocMode=0;continue;}
      if(chunk.indexOf('<w:tbl>')===0)continue;
      var ep=paraObj(chunk);
      if((ep.runs||'').replace(/<[^>]*>/g,'').trim())tocEntries.push('<p>'+ep.runs+'</p>');
      continue;
    }
    if(isBegin){tocMode=1;continue;}
    if(chunk.indexOf('<w:tbl>')===0){
      flushList();
      var rowsHtml='';
      var trRe=/<w:tr(?:\s[^>]*)?>([\s\S]*?)<\/w:tr>/g;var trM;
      while((trM=trRe.exec(chunk))){
        var cellsHtml='';
        var tcRe=/<w:tc(?:\s[^>]*)?>([\s\S]*?)<\/w:tc>/g;var tcM;
        while((tcM=tcRe.exec(trM[1]))){
          var span='';
          var gsM=/<w:gridSpan[^>]*w:val="(\d+)"/.exec(tcM[1]);
          if(gsM&&+gsM[1]>1)span=' colspan="'+gsM[1]+'"';
          var cellParas='';
          var cpRe=/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;var cpM;
          while((cpM=cpRe.exec(tcM[1]))){
            var cp=paraObj(cpM[0]);
            cellParas+='<p>'+(cp.runs||'<br>')+'</p>';
          }
          if(!cellParas)cellParas='<p><br></p>';
          cellsHtml+='<td'+span+'>'+cellParas+'</td>';
        }
        if(cellsHtml)rowsHtml+='<tr>'+cellsHtml+'</tr>';
      }
      if(rowsHtml)html+='<table><tbody>'+rowsHtml+'</tbody></table>';
    }else{
      var p=paraObj(chunk);
      var cmtHtmls=[];
      var crRe=/commentReference\b[^>]*w:id="(\d+)"/g;var crM;
      while((crM=crRe.exec(chunk))){
        var ct=commentMap[crM[1]];
        if(ct===undefined)continue;
        var label=(ct.indexOf('Nhận xét:')===0||ct.indexOf('Chú thích:')===0)?ct:('Nhận xét: '+ct);
        cmtHtmls.push('<p class="comment-para">'+esc(label)+'</p>');
      }
      if(p.bullet){
        var want=(p.bullet==='bullet')?'ul':'ol';
        if(listKind!==want){flushList();html+='<'+want+'>';listKind=want}
        html+='<li>'+(p.runs||'<br>')+'</li>';
        if(cmtHtmls.length)commentQueue=commentQueue.concat(cmtHtmls);
      }else{
        flushList();
        var hasText=(p.runs||'').replace(/<[^>]*>/g,'').trim()!=='';
        if(cmtHtmls.length&&!hasText&&!p.heading){html+=cmtHtmls.join('');}
        else{html+=paraHtml(p);if(cmtHtmls.length)html+=cmtHtmls.join('');}
      }
    }
  }
  flushList();
  if(headerHtml)html=headerHtml+html;
  if(footerHtml)html+=footerHtml;
  var bg='';
  var bgM=/<w:background[^>]*w:color="([0-9A-Fa-f]+)"/.exec(docXml||'');
  if(bgM)bg='#'+bgM[1];
  return {html:html||'<p><br></p>',bg:bg};
}
function addImportedDoc(name,html,bg){
  var docs=loadJSON('vb_docs_v1',[]);
  var rec={id:uid(),name:name,html:html,bg:bg||'',updatedAt:Date.now(),syncStatus:'local'};
  docs.unshift(rec);saveJSON('vb_docs_v1',docs);currentDocId=rec.id;$('docName').value=name;$('editor').innerHTML=html;setPageBg(bg||'');savedDocSnap=docSnapNow();try{docHistoryReset()}catch(e){}
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
    return {name:d.name+'.docx',blob:buildDocxBlobFromNode(tmp,{bg:d.bg||''})};
  }
  if(kind==='sheet'){
    var s=loadJSON('vb_sheets_v1',[]).find(function(x){return x.id===id});if(!s)return null;
    if(s.sheets&&s.sheets.length)return {name:s.name+'.xlsx',blob:xlsxBlobFromSheets(s.sheets)};
    return {name:s.name+'.xlsx',blob:xlsxBlobFromRaw(s.raw||[],{fmt:s.fmt||{},merges:s.merges||[],locked:!!s.locked,freeze:!!s.freeze,grid:s.grid!==false})};
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
function emptySlide(){return {title:'',bullets:[''],image:null,imageW:0,imageH:0,images:null,table:null}}
/* Ban 1.7: anh trong slide: images = [{src,w,h}] (w/h theo px, tu EMU
   cua file nhap hoac tu anh chon); slide cu chi co image van hien qua
   duong dan cu. Bang trong slide: table = {rows:[[o,...],...]}. */
function slideImagesOf(slide){
  if(slide.images&&slide.images.length){
    return slide.images.map(function(im){
      if(typeof im==='string')return {src:im,w:0,h:0};
      return {src:im.src||im.image||'',w:im.w||im.imageW||0,h:im.h||im.imageH||0};
    }).filter(function(im){return !!im.src});
  }
  if(slide.image)return [{src:slide.image,w:slide.imageW||0,h:slide.imageH||0}];
  return [];
}
function slideTableRows(slide){
  if(!slide.table)return null;
  var rows=slide.table.rows||slide.table;
  if(!Array.isArray(rows)||!rows.length)return null;
  return rows;
}
function slideBullets(s){return (s.bullets||[]).filter(function(b){return String(b).trim()!==''})}
function renderSlideEditor(){
  var wrap=$('slideEditor');if(!wrap)return;
  if(!deckSlides.length)deckSlides=[emptySlide()];
  if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;
  if(currentSlide<0)currentSlide=0;
  var thumbs=deckSlides.map(function(s,i){
    var t=(s.hidden?'(Ẩn) ':'')+(s.title||'(Chưa có tiêu đề)');var b=slideBullets(s)[0]||'';
    return '<button class="thumb'+(i===currentSlide?' active':'')+'" onclick="selectSlide('+i+')"><span class="tnum">Slide '+(i+1)+'</span><span class="ttitle">'+esc(t)+'</span><span class="tbody">'+esc(b)+'</span></button>';
  }).join('');
  var s=deckSlides[currentSlide];
  var imgHtml=s.image
    ?'<img class="slideimg" src="'+s.image+'" alt="Ảnh slide"><div class="row"><button onclick="removeSlideImage('+currentSlide+')">Xóa ảnh này</button></div>'
    :'<div class="row"><button onclick="pickSlideImage('+currentSlide+')">Chèn ảnh vào slide này</button></div>';
  var _imgs=slideImagesOf(s);
  var extraImgsHtml=(_imgs.length>1)?'<div class="row">Ảnh khác trong slide ('+(_imgs.length-1)+' ảnh thêm):</div>'+_imgs.slice(1).map(function(im){return '<img class="slideimg" src="'+im.src+'" alt="Ảnh slide">'}).join(''):'';
  var _trows=slideTableRows(s);
  var tableHtml=_trows?'<div class="row">Bảng trong slide:</div><table><tbody>'+_trows.map(function(row){return '<tr>'+(row||[]).map(function(cell){return '<td>'+esc(cell)+'</td>'}).join('')+'</tr>'}).join('')+'</tbody></table>':'';
  var dis=deckLocked?' disabled':'';
  var transLabel={none:'Không',fade:'Mờ dần',slide:'Trượt ngang',up:'Trượt dọc',zoom:'Phóng to'}[deckTransition]||'Không';
  var titleStyle='color:'+(s.titleColor||'#1a1a1a')+';'+(s.titleB?'font-weight:700;':'')+(s.titleI?'font-style:italic;':'')+(s.titleU?'text-decoration:underline;':'');
  var bodyStyle='text-align:'+(s.align||'left')+';';
  var shapeRow=(s.shapes&&s.shapes.length)?'<div class="row"><button onclick="removeLastSlideShape()">Xóa hình vừa chèn</button></div>':'';
  var commentHtml=s.comment?'<p class="small">Nhận xét: '+esc(s.comment)+'</p>':'';
  wrap.innerHTML='<div id="slideThumbs">'+thumbs+'</div><div id="slideMain"><div class="slideframe'+(deckSize==='4:3'?' ratio-43':'')+'" style="background:'+(s.bg||'#ffffff')+';'+(s.font?'font-family:\''+s.font+'\', sans-serif;':'')+'">'
    +'<label>Tiêu đề slide '+(currentSlide+1)+':</label>'
    +'<input type="text" value="'+esc(s.title||'')+'" oninput="updateSlideTitle('+currentSlide+',this.value)" placeholder="Tiêu đề slide" aria-label="Tiêu đề slide" style="'+titleStyle+'"'+dis+'>'
    +'<label>Nội dung — mỗi dòng một ý'+(s.bulletsOn===false?' (đang tắt gạch đầu dòng)':' gạch đầu dòng')+':</label>'
    +'<textarea oninput="updateSlideBullets('+currentSlide+',this.value)" placeholder="Ý thứ nhất&#10;Ý thứ hai" aria-label="Nội dung slide" style="'+bodyStyle+'" spellcheck="'+(pptSpellOn?'true':'false')+'"'+dis+'>'+esc((s.bullets||[]).join('\n'))+'</textarea>'
    +imgHtml+extraImgsHtml+tableHtml
    +slideShapesHtml(s)+shapeRow
    +slideChartHtml(s)
    +commentHtml
    +'</div></div>';
  updateTitleBar();
}
function updateSlideTitle(i,v){if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}if(!deckTypingPushed){pushDeckHistory();deckTypingPushed=true}if(deckSlides[i])deckSlides[i].title=v;updateStatusBar()}
function updateSlideBullets(i,v){if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}if(!deckTypingPushed){pushDeckHistory();deckTypingPushed=true}if(deckSlides[i])deckSlides[i].bullets=String(v).split('\n');updateStatusBar()}
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
function newDeck(){currentDeckId=null;deckLocked=false;deckTypingPushed=false;$('deckName').value='';deckSlides=[emptySlide()];currentSlide=0;deckSize='16:9';deckTransition='none';savedDeckSnap=deckSnapNow();renderSlideEditor();updateStatusBar()}
function saveDeck(silent){
  var name=$('deckName').value.trim()||('Bài trình bày '+(new Date()).toLocaleDateString('vi-VN'));
  if(!currentDeckId)name=uniqueName('deck',name);
  var slides=deckSlides.map(function(s){return {title:s.title||'',bullets:s.bullets||[],image:s.image||null,imageW:s.imageW||0,imageH:s.imageH||0,images:s.images||null,table:s.table||null,bg:s.bg||'',font:s.font||'',titleB:!!s.titleB,titleI:!!s.titleI,titleU:!!s.titleU,titleColor:s.titleColor||'',align:s.align||'left',bulletsOn:s.bulletsOn!==false,shapes:s.shapes||[],chart:s.chart||null,comment:s.comment||'',hidden:!!s.hidden}});
  var decks=loadJSON('vb_decks_v1',[]);
  var rec={id:currentDeckId||uid(),name:name,slides:slides,updatedAt:Date.now(),syncStatus:'local',size:deckSize,transition:deckTransition,locked:deckLocked};
  if(currentDeckId){var d=decks.find(function(x){return x.id===currentDeckId});if(d){Object.assign(d,rec)}else decks.unshift(rec)}
  else{currentDeckId=rec.id;decks.unshift(rec)}
  if(!saveJSON('vb_decks_v1',decks)){updateTitleBar();return}
  savedDeckSnap=deckSnapNow();renderDeckList();renderFileList();if(!silent)alert('Đã lưu bài trình bày vào máy: '+name);updateTitleBar();updateStatusBar();
}
function openDeck(id){
  var d=loadJSON('vb_decks_v1',[]).find(function(x){return x.id===id});if(!d)return;
  currentDeckId=d.id;$('deckName').value=d.name;
  deckSlides=(d.slides&&d.slides.length)?d.slides.map(function(s){return {title:s.title||'',bullets:s.bullets||[],image:s.image||null,imageW:s.imageW||0,imageH:s.imageH||0,images:s.images||null,table:s.table||null,bg:s.bg||'',font:s.font||'',titleB:!!s.titleB,titleI:!!s.titleI,titleU:!!s.titleU,titleColor:s.titleColor||'',align:s.align||'left',bulletsOn:s.bulletsOn!==false,shapes:s.shapes||[],chart:s.chart||null,comment:s.comment||'',hidden:!!s.hidden}}):[emptySlide()];
  deckSize=d.size||'16:9';deckTransition=d.transition||'none';deckLocked=!!d.locked;deckTypingPushed=false;currentSlide=0;savedDeckSnap=deckSnapNow();
  renderSlideEditor();gotoTab('slides');
}
function deleteDeck(id){
  if(!confirm('Xóa bài trình bày này khỏi máy?'))return;
  saveJSON('vb_decks_v1',loadJSON('vb_decks_v1',[]).filter(function(x){return x.id!==id}));
  if(currentDeckId===id)newDeck();renderDeckList();renderFileList();
}
function renderDeckList(){
  var el=$('deckList');if(!el)return;
  var decks=loadJSON('vb_decks_v1',[]);
  el.innerHTML=decks.length?decks.map(function(d){
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
  deckSlides[i].image=null;deckSlides[i].imageW=0;deckSlides[i].imageH=0;deckSlides[i].images=null;renderSlideEditor();
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
  var vis=visibleSlides();
  if(!vis.length){alert('Tất cả slide đang bị ẩn — hiện lại ít nhất 1 slide rồi trình chiếu.');return}
  presentSlides=JSON.parse(JSON.stringify(vis.map(function(v){return v.slide})));
  presentIndex=0;
  $('presentOverlay').style.display='flex';
  renderPresentSlide();startAutoAdvance();
  var ov=$('presentOverlay');
  if(ov.requestFullscreen){ov.requestFullscreen().catch(function(){/* không vào được fullscreen thì vẫn xem bằng lớp phủ toàn màn hình */});}
}
function renderPresentSlide(){
  var s=presentSlides[presentIndex]||emptySlide();
  var bullets=slideBullets(s);
  var titleColor=s.titleColor||((s.bg&&s.bg!=='#ffffff')?'#1a1a1a':'#ffffff');
  var html='<h2 style="color:'+titleColor+'">'+esc(s.title||'(Chưa có tiêu đề)')+'</h2>';
  if(bullets.length){
    if(s.bulletsOn===false)html+='<div>'+bullets.map(function(b){return '<p style="font-size:24px;line-height:1.6">'+esc(b)+'</p>'}).join('')+'</div>';
    else html+='<ul>'+bullets.map(function(b){return '<li>'+esc(b)+'</li>'}).join('')+'</ul>';
  }
  if(s.image)html+='<img class="presentimg" src="'+s.image+'" alt="Ảnh slide '+(presentIndex+1)+'">';
  html+=slideShapesHtml(s,true)+slideChartHtml(s);
  if(!bullets.length&&!s.image&&!(s.shapes&&s.shapes.length)&&!(s.chart&&s.chart.length))html+='<p style="font-size:21px;opacity:.75">Slide này chưa có nội dung. Bấm “✕ Thoát” để về soạn thêm.</p>';
  var ps=$('presentSlide');
  ps.innerHTML=html;
  ps.style.background=(s.bg||'#111111');
  ps.style.color=(s.bg&&s.bg!=='#ffffff')?'#1a1a1a':'#ffffff';
  if(s.bg==='#222222')ps.style.color='#ffffff';
  ps.className=(deckTransition==='fade')?'trans-fade':(deckTransition==='slide'?'trans-slide':(deckTransition==='up'?'trans-up':(deckTransition==='zoom'?'trans-zoom':'')));
  $('presentCounter').textContent='Slide '+(presentIndex+1)+' / '+presentSlides.length;
}
function presentNext(){if(presentIndex<presentSlides.length-1){presentIndex++;renderPresentSlide()}}
function presentPrev(){if(presentIndex>0){presentIndex--;renderPresentSlide()}}
function exitPresent(){
  if(presentTimer)clearInterval(presentTimer);
  presentTimer=null;
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
/* Ban 1.6: slide xuat GIU dinh dang chu (dam/nghieng/gach chan/mau/font
   tieu de), hinh dang don gian, hieu ung chuyen (fade/push), co slide.
   Cu: tieu de luon dam mac dinh, khong mau/font; hieu ung/co chi o trong
   app, file xuat mat het. Zoom khong co dang transition chuan don gian
   nen xuat thanh fade (ghi ro trong code, khong bia kieu transition). */
function pptxSlideXml(slide,imgList,opts){
  opts=opts||{};
  var slideW=(opts.size==='4:3')?9144000:12192000;
  var slideH=6858000;
  var bullets=slideBullets(slide);
  var bgXml=(slide.bg&&slide.bg!=='#ffffff')?'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="'+slide.bg.replace('#','')+'"/></a:solidFill></p:bgPr></p:bg>':'';
  var fontName=slide.font||'Carlito';
  var titleRPr='<a:rPr lang="vi-VN" sz="3200"'+(slide.titleB?' b="1"':'')+(slide.titleI?' i="1"':'')+(slide.titleU?' u="sng"':'')+'>';
  if(slide.titleColor&&/^#[0-9A-Fa-f]{6}$/.test(slide.titleColor))titleRPr+='<a:solidFill><a:srgbClr val="'+slide.titleColor.replace('#','').toUpperCase()+'"/></a:solidFill>';
  titleRPr+='<a:latin typeface="'+esc(fontName)+'"/></a:rPr>';
  var title='<p:sp><p:nvSpPr><p:cNvPr id="2" name="Tiêu đề"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr></p:nvSpPr>'
    +'<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    +'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>'+titleRPr+'<a:t>'+esc(slide.title||'')+'</a:t></a:r></a:p></p:txBody></p:sp>';
  var body='';
  var hasImg=!!slide.image;
  if(bullets.length){
    var bw=hasImg?4572000:8229600;
    var paras=bullets.map(function(b){
      var pPr=(slide.bulletsOn===false)?'<a:pPr/>':'<a:pPr marL="228600" indent="-228600"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
      return '<a:p>'+pPr
        +'<a:r><a:rPr lang="vi-VN" sz="2400"><a:latin typeface="'+esc(fontName)+'"/></a:rPr><a:t>'+esc(b)+'</a:t></a:r></a:p>';
    }).join('');
    body='<p:sp><p:nvSpPr><p:cNvPr id="3" name="Nội dung"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr></p:nvSpPr>'
      +'<p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="'+bw+'" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
      +'<p:txBody><a:bodyPr/><a:lstStyle/>'+paras+'</p:txBody></p:sp>';
  }
  var pic='';
  var _imgList=(imgList&&imgList.length)?imgList:(hasImg?[{rel:(typeof imgList==='string'?imgList:null),src:slide.image,w:slide.imageW||0,h:slide.imageH||0}]:[]);
  if(_imgList.length===1){
    var box=bullets.length
      ?fitBox(3657600,4572000,_imgList[0].w,_imgList[0].h)
      :fitBox(8229600,4572000,_imgList[0].w,_imgList[0].h);
    var px=bullets.length?5486400:Math.round(457200+(8229600-box.w)/2);
    var py=1371600+Math.round((4572000-box.h)/2);
    pic='<p:pic><p:nvPicPr><p:cNvPr id="4" name="Ảnh"/><p:cNvPicPr/></p:nvPicPr>'
      +'<p:blipFill><a:blip r:embed="'+_imgList[0].rel+'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
      +'<p:spPr><a:xfrm><a:off x="'+px+'" y="'+py+'"/><a:ext cx="'+box.w+'" cy="'+box.h+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  }else if(_imgList.length>1){
    /* Nhieu anh: xep chong theo chieu doc trong cot anh (ben phai neu co
       gach dau dong, khong thi toan chieu rong); anh nhap tu file giu
       ti le qua w/h tu EMU, anh khong ro kich thuoc thi vua khung. */
    var slotH=Math.floor(4572000/_imgList.length);
    var bw2=bullets.length?3657600:8229600;
    var px2=bullets.length?5486400:457200;
    pic=_imgList.map(function(im,idx){
      var fb=fitBox(bw2,slotH,im.w,im.h);
      var py2=1371600+idx*slotH+Math.round((slotH-fb.h)/2);
      return '<p:pic><p:nvPicPr><p:cNvPr id="'+(4+idx)+'" name="Ảnh '+(idx+1)+'"/><p:cNvPicPr/></p:nvPicPr>'
        +'<p:blipFill><a:blip r:embed="'+im.rel+'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        +'<p:spPr><a:xfrm><a:off x="'+px2+'" y="'+py2+'"/><a:ext cx="'+fb.w+'" cy="'+fb.h+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    }).join('');
  }
  var tableXml='';
  var _trows2=slideTableRows(slide);
  if(_trows2){
    var ncols=Math.max.apply(null,_trows2.map(function(r){return (r||[]).length}));
    if(ncols>0){
      var colW=Math.floor(8229600/ncols);
      var frameH=1371600;
      var rowH=Math.max(200000,Math.floor(frameH/_trows2.length));
      var gridXml='';
      for(var gi=0;gi<ncols;gi++)gridXml+='<a:gridCol w="'+colW+'"/>';
      var trsXml=_trows2.map(function(row){
        var tcs='';
        for(var ci2=0;ci2<ncols;ci2++){
          var cell=(row||[])[ci2];
          tcs+='<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="vi-VN" sz="1800"><a:latin typeface="'+esc(fontName)+'"/></a:rPr><a:t>'+esc(cell==null?'':cell)+'</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
        }
        return '<a:tr h="'+rowH+'">'+tcs+'</a:tr>';
      }).join('');
      tableXml='<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="Bảng"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr></p:nvGraphicFramePr>'
        +'<p:xfrm><a:off x="457200" y="5486400"/><a:ext cx="8229600" cy="'+frameH+'"/></p:xfrm>'
        +'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>'+gridXml+'</a:tblGrid>'+trsXml+'</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
    }
  }
  var shapes='';
  if(slide.shapes&&slide.shapes.length){
    shapes=slide.shapes.map(function(sh,i){
      var prst=(sh.type==='circle')?'ellipse':(sh.type==='arrow'?'rightArrow':'rect');
      var sx=457200+i*2000000;
      return '<p:sp><p:nvSpPr><p:cNvPr id="'+(10+i)+'" name="Hình '+(i+1)+'"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr></p:nvSpPr>'
        +'<p:spPr><a:xfrm><a:off x="'+sx+'" y="5486400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="'+prst+'"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0B6FA4"/></a:solidFill></p:spPr></p:sp>';
    }).join('');
  }
  var trans='';
  if(opts.transition==='fade')trans='<p:transition spd="med"><p:fade thruBlk="1"/></p:transition>';
  else if(opts.transition==='slide')trans='<p:transition spd="med"><p:push dir="l"/></p:transition>';
  else if(opts.transition==='up')trans='<p:transition spd="med"><p:push dir="u"/></p:transition>';
  else if(opts.transition==='zoom')trans='<p:transition spd="med"><p:fade thruBlk="1"/></p:transition>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    +'<p:cSld>'+bgXml+'<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
    +'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+slideW+'" cy="'+slideH+'"/></a:xfrm></p:grpSpPr>'
    +title+body+pic+shapes+tableXml+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'+trans+'</p:sld>';
}
function buildPptxBlob(deck){
  var slides=(deck.slides&&deck.slides.length)?deck.slides:[emptySlide()];
  var entries=[];
  var types='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    +'<Default Extension="xml" ContentType="application/xml"/>'
    +'<Default Extension="jpg" ContentType="image/jpeg"/>'
    +'<Default Extension="png" ContentType="image/png"/>'
    +'<Default Extension="gif" ContentType="image/gif"/>'
    +'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
  slides.forEach(function(_,i){types+='<Override PartName="/ppt/slides/slide'+(i+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';});
  types+='<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    +'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    +'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>';
  entries.push({name:'[Content_Types].xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+types});
  entries.push({name:'_rels/.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>'});
  var presRels='',sldIds='';
  slides.forEach(function(s,i){
    presRels+='<Relationship Id="rIdS'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+(i+1)+'.xml"/>';
    sldIds+='<p:sldId id="'+(256+i)+'" r:id="rIdS'+(i+1)+'"'+(s.hidden?' show="0"':'')+'/>';
  });
  presRels+='<Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>';
  entries.push({name:'ppt/_rels/presentation.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+presRels+'</Relationships>'});
  var sldSz=(deck.size==='4:3')?'<p:sldSz cx="9144000" cy="6858000"/>':'<p:sldSz cx="12192000" cy="6858000"/>';
  entries.push({name:'ppt/presentation.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst><p:sldIdLst>'+sldIds+'</p:sldIdLst>'+sldSz+'<p:notesSz cx="6858000" cy="9144000"/></p:presentation>'});
  var imgCount=0;
  slides.forEach(function(s,i){
    var slideRels='<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>';
    var imgList=slideImagesOf(s).map(function(im,k){
      imgCount++;
      var imgExt='jpg';
      var imgMm=/^data:([^;]+);base64,/.exec(im.src||'');
      if(imgMm){var sub2=imgMm[1].split('/')[1]||'jpg';sub2=sub2.replace(/[^a-z0-9]/g,'');imgExt=(sub2==='jpeg')?'jpg':(sub2||'jpg')}
      slideRels+='<Relationship Id="rIdImg'+(k+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+imgCount+'.'+imgExt+'"/>';
      entries.push({name:'ppt/media/image'+imgCount+'.'+imgExt,data:dataUrlToBytes(im.src)});
      return {rel:'rIdImg'+(k+1),src:im.src,w:im.w,h:im.h};
    });
    entries.push({name:'ppt/slides/_rels/slide'+(i+1)+'.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+slideRels+'</Relationships>'});
    entries.push({name:'ppt/slides/slide'+(i+1)+'.xml',data:pptxSlideXml(s,imgList,{size:deck.size,transition:deck.transition})});
  });
  entries.push({name:'ppt/slideLayouts/slideLayout1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld name="Trống"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'});
  entries.push({name:'ppt/slideMasters/_rels/slideMaster1.xml.rels',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'});
  entries.push({name:'ppt/slideMasters/slideMaster1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="lt1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdL"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Carlito"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>'});
  entries.push({name:'ppt/theme/theme1.xml',data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Văn phòng của Bùi"><a:themeElements><a:clrScheme name="Mặc định"><a:dk1><a:sysClr val="windowText" lastClr="1A1A1A"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="0B6FA4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="70AD47"/></a:accent5><a:accent6><a:srgbClr val="7030A0"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Mặc định"><a:majorFont><a:latin typeface="Carlito"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Carlito"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Mặc định"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'});
  return zipStored(entries);
}
/* Ban 1.7: nhap pptx GIU them bang + moi anh (1.6 chi lay anh dau, bo
   bang). Giu thu tu slide theo presentation.xml, tieu de, noi dung
   (moi doan = 1 y), moi anh cua moi slide, bang dau tien cua moi
   slide, slide an (show="0"). Khong nhap duoc: bieu do, hinh dang,
   ghi chu, hieu ung, chu de; slide co nhieu bang thi chi giu bang dau. */
function parsePptxToDeck(entries){
  function relsMap(xml){
    var m={};var re=/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;var mm;
    while((mm=re.exec(xml||'')))m[mm[1]]=mm[2];
    return m;
  }
  function normMedia(t){
    t=(t||'').replace(/^\.\.\//,'').replace(/^\/+/,'');
    if(t.indexOf('ppt/')===0)return t;
    return 'ppt/'+t;
  }
  function normSlide(t){
    t=(t||'').replace(/^\/+/,'');
    if(t.indexOf('ppt/')===0)return t;
    return 'ppt/'+t;
  }
  var presXml=zipEntryText(entries['ppt/presentation.xml']||'');
  var presRels=relsMap(zipEntryText(entries['ppt/_rels/presentation.xml.rels']||''));
  var size='16:9';
  var szM=/<p:sldSz[^>]*cx="(\d+)"/.exec(presXml);
  if(szM&&+szM[1]===9144000)size='4:3';
  var order=[];
  var idRe=/<p:sldId([^>]*)\/>/g;var idM;
  while((idM=idRe.exec(presXml))){
    var ridM=/r:id="([^"]+)"/.exec(idM[1]);
    var showM=/show="([^"]+)"/.exec(idM[1]);
    if(ridM)order.push({rid:ridM[1],hidden:showM&&showM[1]==='0'});
  }
  if(!order.length){
    Object.keys(entries).forEach(function(k){
      var sm=/^ppt\/slides\/slide(\d+)\.xml$/.exec(k);
      if(sm)order.push({rid:'',path:k,hidden:false});
    });
    order.sort(function(a,b){return a.path<b.path?-1:1});
  }
  var slides=order.map(function(o){
    var path=o.path||normSlide(presRels[o.rid]||'');
    var slideXml=zipEntryText(entries[path]||'');
    var relsPath=path.replace('ppt/slides/','ppt/slides/_rels/')+'.rels';
    var slideRels=relsMap(zipEntryText(entries[relsPath]||''));
    var title='';var bullets=[];
    var spRe=/<p:sp>([\s\S]*?)<\/p:sp>/g;var spM;var firstShape=true;
    while((spM=spRe.exec(slideXml))){
      var spXml=spM[1];
      var phM=/<a:ph[^>]*type="([^"]+)"/.exec(spXml);
      var phType=phM?phM[1]:'';
      var paras=[];
      var pRe=/<a:p>([\s\S]*?)<\/a:p>/g;var pM;
      while((pM=pRe.exec(spXml))){
        var txt='';
        var tRe=/<a:t>([^<]*)<\/a:t>/g;var tM;
        while((tM=tRe.exec(pM[1])))txt+=xmlUnesc(tM[1]);
        if(txt.trim()!=='')paras.push(txt);
      }
      if(!paras.length)continue;
      if(phType==='title'||phType==='ctrTitle')title=paras.join(' ');
      else if(phType==='body'||phType==='subTitle'||phType===''||firstShape){
        if(title===''&&firstShape&&phType===''){title=paras.join(' ')}
        else bullets=bullets.concat(paras);
      }else bullets=bullets.concat(paras);
      firstShape=false;
    }
    var images=[];
    var picRe=/<p:pic>([\s\S]*?)<\/p:pic>/g;var picM2;
    while((picM2=picRe.exec(slideXml))){
      var blipM=/<a:blip[^>]*r:embed="([^"]+)"/.exec(picM2[1]);
      if(!blipM)continue;
      var mediaPath=normMedia(slideRels[blipM[1]]||'');
      var bytes=entries[mediaPath];
      if(!bytes)continue;
      var ext=mediaPath.split('.').pop().toLowerCase();
      var mime=(ext==='png')?'image/png':(ext==='gif')?'image/gif':(ext==='webp')?'image/webp':'image/jpeg';
      var exM=/<a:ext cx="(\d+)" cy="(\d+)"/.exec(picM2[1]);
      var iw=exM?Math.round(+exM[1]/9525):0;
      var ih=exM?Math.round(+exM[2]/9525):0;
      images.push({src:'data:'+mime+';base64,'+bytesToBase64(bytes),w:iw,h:ih});
    }
    var image=images.length?images[0].src:null;
    var imageW=images.length?images[0].w:0;
    var imageH=images.length?images[0].h:0;
    /* Bang trong slide: lay bang a:tbl dau tien (slide co nhieu bang thi
       chi giu bang dau — ghi ro o thong bao nhap). */
    var table=null;
    var tblM=/<a:tbl>([\s\S]*?)<\/a:tbl>/.exec(slideXml);
    if(tblM){
      var trows=[];
      var trRe2=/<a:tr(?:\s[^>]*)?>([\s\S]*?)<\/a:tr>/g;var trM2;
      while((trM2=trRe2.exec(tblM[1]))){
        var cells=[];
        var tcRe2=/<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/g;var tcM2;
        while((tcM2=tcRe2.exec(trM2[1]))){
          var cellParts=[];
          var cpRe2=/<a:p>([\s\S]*?)<\/a:p>/g;var cpM2;
          while((cpM2=cpRe2.exec(tcM2[1]))){
            var txt2='';
            var tRe3=/<a:t>([^<]*)<\/a:t>/g;var tM3;
            while((tM3=tRe3.exec(cpM2[1])))txt2+=xmlUnesc(tM3[1]);
            cellParts.push(txt2);
          }
          cells.push(cellParts.join(' '));
        }
        if(cells.length)trows.push(cells);
      }
      if(trows.length)table={rows:trows};
    }
    return {title:title,bullets:bullets.length?bullets:[''],image:image,imageW:imageW,imageH:imageH,images:images.length?images:null,table:table,hidden:!!o.hidden};
  });
  if(!slides.length)slides=[emptySlide()];
  return {slides:slides,size:size};
}
function currentDeckForExport(){
  return {name:$('deckName').value.trim()||'bai-trinh-bay',slides:deckSlides,size:deckSize,transition:deckTransition};
}
function exportPptx(){
  if(!confirmExportWarnings(exportWarnings('ppt')))return;
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
    if(rt)rt.style.display='flex';if(rb)rb.style.display='';
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
  /* Bản 1.4: tab đang chọn luôn cuộn vào tầm nhìn (không bị kẹt ngoài rìa
     phải trên điện thoại). block:'nearest' để không cuộn trang dọc theo. */
  try{var act=rt.querySelector('button.active');if(act)act.scrollIntoView({inline:'center',block:'nearest'});}catch(e){}
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
  layoutRibbon();
}
/* ---------- Bản 1.5: ribbon không giấu nhóm trong im lặng ----------
   Desktop (từ 700px trở lên, gồm cả 1024px+): đo chiều rộng thật của các
   nhóm; nhóm nào tràn thì ẩn khỏi hàng và gom vào menu tràn "⋯" kèm tên
   nhóm, bấm nút trong menu sẽ bấm đúng nút gốc. Điện thoại (<700px): giữ
   vuốt ngang + hàng chip tên nhóm đủ mọi nhóm, bấm chip cuộn tới nhóm. */
function activeRibbonPanel(){
  var ps=document.querySelectorAll('.rpanel');
  for(var i=0;i<ps.length;i++){if(ps[i].style.display==='flex')return ps[i]}
  return null;
}
function closeRibbonOverflow(){
  var rb=$('ribbon');if(rb)rb.classList.remove('overflow-open');
  var ov=$('ribbonOverflow');if(ov)ov.style.display='none';
}
function toggleRibbonOverflow(){
  var rb=$('ribbon');if(!rb)return;
  if(rb.classList.contains('overflow-open')){closeRibbonOverflow();return}
  rb.classList.add('overflow-open');
  var ov=$('ribbonOverflow');if(ov)ov.style.display='block';
}
document.addEventListener('click',function(e){
  if(!e.target||!e.target.closest||!e.target.closest('#ribbon'))closeRibbonOverflow();
});
function highlightActiveChip(panel){
  var chips=$('groupChips');if(!chips)return;
  var groups=panel.querySelectorAll('.rgroup');
  var cur=0;
  for(var i=0;i<groups.length;i++){
    if(groups[i].offsetLeft-panel.scrollLeft<=48)cur=i;
  }
  chips.querySelectorAll('.gchip').forEach(function(ch,idx){ch.classList.toggle('active',idx===cur)});
}
function renderGroupChips(panel){
  var chips=$('groupChips');if(!chips)return;
  if(window.innerWidth>=700){chips.innerHTML='';return}
  var groups=panel.querySelectorAll('.rgroup');
  chips.innerHTML=Array.prototype.map.call(groups,function(g,i){
    var lb=g.querySelector('.glabel');
    return '<button type="button" class="gchip" data-gi="'+i+'">'+esc(lb?lb.textContent:'')+'</button>';
  }).join('');
  chips.querySelectorAll('.gchip').forEach(function(ch){
    ch.addEventListener('click',function(){
      var g=groups[+ch.dataset.gi];if(!g)return;
      try{g.scrollIntoView({inline:'start',block:'nearest'})}catch(e){panel.scrollLeft=g.offsetLeft}
      highlightActiveChip(panel);
    });
  });
  panel.onscroll=function(){highlightActiveChip(panel)};
  highlightActiveChip(panel);
}
var ribbonOverflowTargets=[];
function layoutRibbon(){
  var rb=$('ribbon');if(!rb)return;
  var panel=activeRibbonPanel();if(!panel)return;
  closeRibbonOverflow();
  var groups=panel.querySelectorAll('.rgroup');
  groups.forEach(function(g){g.style.display=''});
  var ovBtn=$('ribbonOverflowBtn'),ov=$('ribbonOverflow');
  if(window.innerWidth<700){
    if(ovBtn)ovBtn.style.display='none';
    if(ov)ov.innerHTML='';
    renderGroupChips(panel);
    return;
  }
  var chips=$('groupChips');if(chips)chips.innerHTML='';
  if(ov)ov.innerHTML='';
  if(ovBtn)ovBtn.style.display='none';
  var panelStyle=getComputedStyle(panel);
  var pad=(parseFloat(panelStyle.paddingLeft)||0)+(parseFloat(panelStyle.paddingRight)||0);
  var total=pad;
  groups.forEach(function(g){total+=g.offsetWidth});
  if(total<=rb.clientWidth)return;
  if(ovBtn)ovBtn.style.display='block';
  var btnW=ovBtn?(ovBtn.offsetWidth+8):56;
  var avail=rb.clientWidth-btnW;
  var used=pad,kept=0;
  for(var i=0;i<groups.length;i++){
    if(used+groups[i].offsetWidth<=avail){used+=groups[i].offsetWidth;kept++}
    else break;
  }
  if(kept===0)kept=1;
  var hidden=[];
  for(var j=kept;j<groups.length;j++){groups[j].style.display='none';hidden.push(groups[j])}
  if(!hidden.length){if(ovBtn)ovBtn.style.display='none';return}
  ribbonOverflowTargets=[];
  var html='';
  hidden.forEach(function(g){
    var lb=g.querySelector('.glabel');
    html+='<div class="ogroup"><div class="oglabel">'+esc(lb?lb.textContent:'')+'</div><div class="obtns">';
    g.querySelectorAll('.rbtn').forEach(function(b){
      var lab=b.querySelector('.rlabel');var ic=b.querySelector('.ricon');
      var idx=ribbonOverflowTargets.length;ribbonOverflowTargets.push(b);
      html+='<button type="button" class="obtn" data-oi="'+idx+'">'
        +'<span class="ricon">'+(ic?ic.innerHTML:'')+'</span>'
        +'<span class="rlabel">'+esc(lab?lab.textContent:(ic?ic.textContent:''))+'</span></button>';
    });
    html+='</div></div>';
  });
  ov.innerHTML=html;
  ov.querySelectorAll('.obtn').forEach(function(ob){
    ob.addEventListener('click',function(){
      var t=ribbonOverflowTargets[+ob.dataset.oi];
      closeRibbonOverflow();
      if(t)t.click();
    });
  });
}
function currentFileName(){
  if(currentApp==='word')return $('docName').value.trim()||'Tài liệu chưa đặt tên';
  if(currentApp==='excel')return $('sheetName').value.trim()||'Bảng tính chưa đặt tên';
  if(currentApp==='ppt')return $('deckName').value.trim()||'Bài trình bày chưa đặt tên';
  return APP_TITLES[currentApp]||'';
}
/* Ban 1.6: chi bao "Chua luu" chinh xac. Cu: chi can co id la bao "Da luu"
   du nguoi dung sua tiep sau do. Moi: anh chup noi dung luc luu/mo; tieu de
   so sanh noi dung hien tai voi anh chup — sua roi ma chua luu thi bao
   "Chua luu", luu xong thi het bao. */
var savedDocSnap='',savedSheetSnap='',savedDeckSnap='';
var currentDocBg='';
function docSnapNow(){return ($('docName').value||'')+'\n'+($('editor').innerHTML||'')+'\nBG:'+(currentDocBg||'')}
function sheetSnapNow(){return ($('sheetName').value||'')+'\n'+sheetSnapshot()+'\nLOCK:'+sheetLocked+'\nFREEZE:'+freezeTopRow+'\nGRID:'+gridOn}
function deckSnapNow(){return ($('deckName').value||'')+'\n'+deckSnapshot()}
function updateTitleBar(){
  var an=$('appName');if(an)an.textContent=APP_TITLES[currentApp]||'Văn phòng';
  var dt=$('docTitle');if(dt)dt.textContent=currentFileName();
  var ss=$('saveState');
  if(ss){
    var saved=false;
    if(currentApp==='word')saved=!!currentDocId&&docSnapNow()===savedDocSnap;
    else if(currentApp==='excel')saved=!!currentSheetId&&sheetSnapNow()===savedSheetSnap;
    else if(currentApp==='ppt')saved=!!currentDeckId&&deckSnapNow()===savedDeckSnap;
    ss.textContent=(currentApp==='word'||currentApp==='excel'||currentApp==='ppt')?(saved?' • Đã lưu':' • Chưa lưu'):'';
  }
}
function updateStatusBar(){
  var l=$('statusLeft');if(!l)return;
  if(currentApp==='word'){
    var t=editorText();var words=t.trim()?t.trim().split(/\s+/).length:0;
    l.textContent='Trang 1 • '+words+' từ • '+t.length+' ký tự';
  }else if(currentApp==='excel'){
    l.textContent='Ô '+cellRef(selCell.r,selCell.c)+' • '+currentSheetName()+(sheetLocked?' • Đã bảo vệ':'');
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
  pushSheetHistory();gridOn=!gridOn;var w=$('sheetWrap');
  if(w){if(gridOn)w.classList.remove('nogrid');else w.classList.add('nogrid');}
}
function toggleThumbs(){
  thumbsOn=!thumbsOn;var l=$('slideLayout');
  if(l){if(thumbsOn)l.classList.remove('hide-thumbs');else l.classList.add('hide-thumbs');}
}
/* ---------- Ban 1.6 (cap nhat 1.7): bao ro muc nao CHUA luu theo file truoc khi xuat ----------
   Nguyen tac: khong de nguoi dung tuong da luu ma thuc ra mat. Moi muc chi
   o trong phien hoac chi luu o may (localStorage) deu phai xuat hien o day.
   Danh sach duoc kiem tra that tu noi dung hien tai (khong phai nhan cung):
   - Word da luu theo docx: chu, dinh dang chu, tieu de, bang, anh, thut le,
     vien doan, theo doi thay doi (w:ins/w:del), mau nen trang, khoang cach
     dong. Ban 1.7 them: dau trang/chan trang that (header/footer parts),
     ghi chu that (comments.xml), muc luc dang truong TOC (quet lai tu
     tieu de khi xuat). CHUA luu: bo cuc nhieu cot.
   - Excel da luu theo xlsx: gia tri, cong thuc, dinh dang o, dong bang,
     khoa trang. Ban 1.7 them: gop o, nhieu trang tinh, ghi chu o that
     (comments legacy). CHUA luu: bieu do cot (chi hien trong phien).
   - PowerPoint da luu theo pptx: chu, dinh dang chu tieu de, nen slide,
     anh, hinh dang don gian, hieu ung, co slide, slide an. CHUA luu:
     nhan xet slide, bieu do slide, khoa bai (chi co tac dung trong app).
   Luu vao may (localStorage) thi GIU DU cac muc tren, ke ca muc chua xuat
   duoc — chi khi XUAT file thi muc do moi mat, nen canh bao dat o cua xuat. */
function exportWarnings(app){
  var out=[];
  if(app==='word'){
    var html=$('editor')?($('editor').innerHTML||''):''; 
    /* Ban 1.7: dau trang/chan trang/ghi chu/muc luc nay xuat that theo
       docx (header/footer parts, comments.xml, truong TOC) nen khong
       con canh bao cho cac muc nay nua; chi con bo cuc nhieu cot. */
    if(html.indexOf('column-count')>=0)out.push('Bố cục nhiều cột: khi xuất .docx chỉ còn là văn bản một cột');
  }else if(app==='excel'){
    /* Ban 1.7: ghi chu o nay xuat that sang .xlsx (comments legacy) nen
       khong con canh bao; chi con bieu do cot trong phien. */
    var ch=$('sheetChart');
    if(ch&&ch.style&&ch.style.display!=='none'&&ch.innerHTML)out.push('Biểu đồ cột đang hiện chỉ ở trong phiên, không xuất sang .xlsx và không lưu theo file');
  }else if(app==='ppt'){
    var hasComment=false,hasChart=false;
    deckSlides.forEach(function(s,i){
      if(s.comment)hasComment=true;
      if(s.chart)hasChart=true;
    });
    if(hasComment)out.push('Nhận xét slide: chỉ lưu trong app, không xuất sang .pptx');
    if(hasChart)out.push('Biểu đồ slide: chỉ hiện trong app, không xuất sang .pptx');
    if(deckLocked)out.push('Khóa bài: chỉ có tác dụng trong app, file .pptx xuất ra vẫn sửa được');
  }
  return out;
}
function confirmExportWarnings(list){
  if(!list||!list.length)return true;
  return confirm('LƯU Ý — các mục sau CHƯA lưu theo file (chỉ ở trong app), xuất file sẽ mất dạng gốc của chúng:\n- '+list.join('\n- ')+'\n\nBấm OK để vẫn xuất file, Hủy để quay lại giữ nguyên trong app.');
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
    var warns=exportWarnings(currentApp);
    if(warns.length)info+='<br><b>Chưa lưu theo file khi xuất:</b> '+esc(warns.join('; '));
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
      +'<p><b>Chính tả:</b> nút Chính tả chỉ bật/tắt gạch chân cơ học của trình duyệt — CHƯA có từ điển tiếng Việt, không phải kiểm tra chính tả thật.</p>'
      +'<p><b>Excel:</b> công thức tính trên GIÁ TRỊ SỐ THẬT của ô, không tính trên chữ hiển thị: ô gõ “1.000.000 ₫” có giá trị 1000000, ô gõ “50%” có giá trị 0,5. Hàm có: SUM, AVERAGE, COUNT, MAX, MIN, IF, ROUND, ABS, VLOOKUP (gần đúng cho cột đã sắp xếp — số và chữ theo thứ tự tiếng Việt, FALSE để khớp chính xác), COUNTIF, SUMIF; công thức liên trang dạng Sheet2!A1.</p>'
      +'<p><b>PDF:</b> chụm hai ngón để phóng to/thu nhỏ trang, vẽ chú thích bằng ngón tay trên trang, chạm để thêm ghi chú. Không cần chuột hay thao tác chính xác từng pixel.</p>'
      +'</details>';
    c.innerHTML='<h3>Thông tin</h3><p>'+info+'</p>'+nameRow+guide;
  }
}
/* Undo / Redo (1.7): Word dùng ngăn xếp tự làm (wordHistory); Excel/PowerPoint dùng ảnh chụp nhanh. */
var sheetUndo=[],sheetRedo=[];
/* Ban 1.6: anh chup gom ca khoa trang/dong bang/luoi de undo tra dung trang thai;
   dinh dang o phai di theo hang/cot khi chen/xoa/sap xep — cu: dinh dang o
   lai so hang cu trong khi noi dung doi cho, sai cap doi. */
function sheetSnapshot(){return JSON.stringify({raw:sheetRaw,fmt:cellFmt,merges:sheetMerges,sheets:workbookSnapshotSheets(),active:activeSheetIndex,locked:sheetLocked,freeze:freezeTopRow,grid:gridOn})}
function sheetRestore(snap){
  try{var o=JSON.parse(snap);
    if(o&&o.sheets&&o.sheets.length){
      wbSheets=o.sheets.map(function(sh){return {name:sh.name||'Sheet1',raw:normalizeSheetRaw(sh.raw),fmt:sh.fmt||{},merges:normalizeMerges(sh.merges||[]),locked:!!sh.locked,freeze:!!sh.freeze,grid:sh.grid!==false}});
      activeSheetIndex=Math.min(Math.max(0,o.active||0),wbSheets.length-1);
      var sh2=wbSheets[activeSheetIndex];
      sheetRaw=normalizeSheetRaw(sh2.raw);cellFmt=sh2.fmt||{};sheetMerges=normalizeMerges(sh2.merges||[]);
      sheetLocked=!!sh2.locked;freezeTopRow=!!sh2.freeze;gridOn=sh2.grid!==false;
      var w2=$('sheetWrap');if(w2){if(gridOn)w2.classList.remove('nogrid');else w2.classList.add('nogrid')}
      return;
    }
    if(o&&o.raw){sheetRaw=o.raw;cellFmt=o.fmt||{};sheetMerges=normalizeMerges(o.merges||[]);
    if(o.locked!==undefined)sheetLocked=!!o.locked;
    if(o.freeze!==undefined)freezeTopRow=!!o.freeze;
    if(o.grid!==undefined){gridOn=o.grid!==false;var w=$('sheetWrap');if(w){if(gridOn)w.classList.remove('nogrid');else w.classList.add('nogrid')}}
    return}}catch(e){}
  sheetRaw=JSON.parse(snap);cellFmt={};sheetMerges=[];
}
function shiftCellFmtRows(at,delta){
  var nf={};
  Object.keys(cellFmt).forEach(function(k){
    var p=k.split(',');var r=+p[0],c=+p[1];
    if(delta>0&&r>=at)nf[(r+delta)+','+c]=cellFmt[k];
    else if(delta<0&&r===at){/* hang bi xoa: bo dinh dang */}
    else if(delta<0&&r>at)nf[(r+delta)+','+c]=cellFmt[k];
    else nf[k]=cellFmt[k];
  });
  cellFmt=nf;
}
function shiftCellFmtCols(at,delta){
  var nf={};
  Object.keys(cellFmt).forEach(function(k){
    var p=k.split(',');var r=+p[0],c=+p[1];
    if(delta>0&&c>=at)nf[r+','+(c+delta)]=cellFmt[k];
    else if(delta<0&&c===at){}
    else if(delta<0&&c>at)nf[r+','+(c+delta)]=cellFmt[k];
    else nf[k]=cellFmt[k];
  });
  cellFmt=nf;
}
function pushSheetHistory(){sheetUndo.push(sheetSnapshot());if(sheetUndo.length>50)sheetUndo.shift();sheetRedo=[];}
function undoSheet(){if(!sheetUndo.length)return;sheetRedo.push(sheetSnapshot());sheetRestore(sheetUndo.pop());renderSheet();updateStatusBar();}
function redoSheet(){if(!sheetRedo.length)return;sheetUndo.push(sheetSnapshot());sheetRestore(sheetRedo.pop());renderSheet();updateStatusBar();}
/* Ban 1.6: anh chup PPT gom co slide/hieu ung/khoa; go tieu de/noi dung cung
   co 1 buoc undo cho moi lan dung lai o slide (khong phai moi phim) — cu:
   go tieu de khong day lich su nen undo khong tra lai duoc, doi nen/co/
   hieu ung cung khong co undo. */
var deckUndo=[],deckRedo=[],deckTypingPushed=false;
function deckSnapshot(){return JSON.stringify({slides:deckSlides,size:deckSize,transition:deckTransition,locked:deckLocked})}
function deckRestore(snap){
  try{var o=JSON.parse(snap);
    if(o&&o.slides){deckSlides=o.slides;if(o.size!==undefined)deckSize=o.size;if(o.transition!==undefined)deckTransition=o.transition;if(o.locked!==undefined)deckLocked=!!o.locked;return}
    deckSlides=o;return}catch(e){deckSlides=[emptySlide()]}
}
function pushDeckHistory(){deckTypingPushed=false;deckUndo.push(deckSnapshot());if(deckUndo.length>50)deckUndo.shift();deckRedo=[];}
function undoDeck(){if(!deckUndo.length)return;deckTypingPushed=false;deckRedo.push(deckSnapshot());deckRestore(deckUndo.pop());if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;if(currentSlide<0)currentSlide=0;renderSlideEditor();updateStatusBar();}
function redoDeck(){if(!deckRedo.length)return;deckTypingPushed=false;deckUndo.push(deckSnapshot());deckRestore(deckRedo.pop());if(currentSlide>=deckSlides.length)currentSlide=deckSlides.length-1;if(currentSlide<0)currentSlide=0;renderSlideEditor();updateStatusBar();}
/* Ban 1.7: hoan tac Word TU LAM, khong dung lenh native cua trinh duyet
   (cu: cmd('undo')/cmd('redo') phu thuoc execCommand nen khong kiem
   chung tu dong duoc). Loi lich su la ham thuan makeHistoryStack:
   past giu cac anh chup TRUOC moi lan sua, current giu o ngoai; undo
   day current vao future roi lay past, redo nguoc lai; push moi xoa
   future. Tich hop: docNoteEdit() duoc goi khi editor doi (su kien
   input, lenh dinh dang, chen noi dung, dau/chan trang, chap nhan/tu
   choi thay doi) — moi lan doi day anh chup truoc do vao past. Khoi
   phuc con tro chi co gang dat ve cuoi tai lieu (anh chup la chuoi
   HTML, khong luu vi tri con tro chinh xac) — ghi ro, khong bia la
   giu dung vi tri cu. */
function makeHistoryStack(limit){
  var past=[],future=[];
  var L=limit||50;
  return {
    past:past,future:future,
    push:function(state){
      if(past.length&&past[past.length-1]===state)return;
      past.push(state);
      if(past.length>L)past.shift();
      future.length=0;
    },
    undo:function(current){
      if(!past.length)return {state:current,changed:false};
      future.push(current);
      return {state:past.pop(),changed:true};
    },
    redo:function(current){
      if(!future.length)return {state:current,changed:false};
      past.push(current);
      return {state:future.pop(),changed:true};
    },
    canUndo:function(){return past.length>0},
    canRedo:function(){return future.length>0},
    clear:function(){past.length=0;future.length=0}
  };
}
var wordHistory=makeHistoryStack(50);
var docHistoryCurrent='';
function docNoteEdit(){
  if(typeof wordHistory==='undefined')return;
  var ed=$('editor');if(!ed)return;
  var html=ed.innerHTML||'';
  if(html!==docHistoryCurrent){wordHistory.push(docHistoryCurrent);docHistoryCurrent=html;}
}
function docHistoryReset(){
  wordHistory.clear();
  var ed=$('editor');
  docHistoryCurrent=ed?(ed.innerHTML||''):'';
}
function placeCaretAtEnd(){
  try{
    var ed=$('editor');if(!ed)return;
    ed.focus();
    var sel=window.getSelection&&window.getSelection();
    if(!sel||!sel.removeAllRanges||!document.createRange)return;
    var range=document.createRange();
    range.selectNodeContents(ed);range.collapse(false);
    sel.removeAllRanges();sel.addRange(range);
  }catch(e){}
}
function wordUndo(){
  if(!wordHistory.canUndo())return;
  var ed=$('editor');if(!ed)return;
  var res=wordHistory.undo(ed.innerHTML||'');
  if(!res.changed)return;
  ed.innerHTML=res.state;docHistoryCurrent=res.state;
  placeCaretAtEnd();updateStatusBar();
}
function wordRedo(){
  if(!wordHistory.canRedo())return;
  var ed=$('editor');if(!ed)return;
  var res=wordHistory.redo(ed.innerHTML||'');
  if(!res.changed)return;
  ed.innerHTML=res.state;docHistoryCurrent=res.state;
  placeCaretAtEnd();updateStatusBar();
}
function undo(){if(currentApp==='word')wordUndo();else if(currentApp==='excel')undoSheet();else if(currentApp==='ppt')undoDeck();}
function redo(){if(currentApp==='word')wordRedo();else if(currentApp==='excel')redoSheet();else if(currentApp==='ppt')redoDeck();}
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
function insertHtmlAtCursor(html){$('editor').focus();document.execCommand('insertHTML',false,html);try{docNoteEdit()}catch(e){}updateStatusBar();}
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
  /* Bản 1.4: quét mọi tiêu đề trong tài liệu (kể cả tiêu đề lồng trong khối),
     không chỉ nút con trực tiếp — mục lục cũ bỏ sót tiêu đề khi con trỏ đang
     ở trong đoạn có định dạng lồng nhau. */
  var heads=[];
  Array.from($('editor').querySelectorAll('h1,h2,h3')).forEach(function(n){
    var tg=n.tagName.toLowerCase();
    if(nodeText(n).trim())heads.push({lv:tg,txt:nodeText(n).trim()});
  });
  if(!heads.length){alert('Chưa có tiêu đề nào (Tiêu đề 1/2/3) để lập mục lục. Đặt tiêu đề cho các phần trước, rồi bấm lại.');return}
  var html='<div class="doc-toc"><h2>Mục lục</h2>'+heads.map(function(h){
    var pad=h.lv==='h2'?'&nbsp;&nbsp;&nbsp;&nbsp;':(h.lv==='h3'?'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;':'');
    return '<p>'+pad+esc(h.txt)+'</p>';
  }).join('')+'</div>';
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
  alert(on?'Đã tắt gạch chân cơ học của trình duyệt.':'Đã bật gạch chân cơ học của trình duyệt — LƯU Ý: đây chỉ là kiểm tra cơ học, CHƯA có từ điển tiếng Việt, không phải kiểm tra chính tả thật.');
}
function setPageBg(color){currentDocBg=color||'';var p=$('wordPage');if(p)p.style.background=color||'';}
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
  pushSheetHistory();shiftCellFmtRows(selCell.r,1);shiftMergesRows(selCell.r,1);sheetRaw.splice(selCell.r,0,Array(SHEET_COLS).fill(''));sheetRaw.length=SHEET_ROWS;renderSheet();
}
function insertColLeft(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();shiftCellFmtCols(selCell.c,1);shiftMergesCols(selCell.c,1);
  sheetRaw.forEach(function(row){row.splice(selCell.c,0,'');row.length=SHEET_COLS;});
  renderSheet();
}
function deleteRow(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();shiftCellFmtRows(selCell.r,-1);shiftMergesRows(selCell.r,-1);sheetRaw.splice(selCell.r,1);sheetRaw.push(Array(SHEET_COLS).fill(''));renderSheet();
}
function deleteCol(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();shiftCellFmtCols(selCell.c,-1);shiftMergesCols(selCell.c,-1);sheetRaw.forEach(function(row){row.splice(selCell.c,1);row.push('');});renderSheet();
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
  var indexed=sheetRaw.slice(0,lastR+1).map(function(row,i){return {row:row,i:i}});
  indexed.sort(function(a,b){
    var x=a.row[c]||'',y=b.row[c]||'';
    var nx=parseCellNumber(x),ny=parseCellNumber(y);
    var cmp=(nx!=null&&ny!=null)?(nx-ny):String(x).localeCompare(String(y),'vi');
    return dir==='desc'?-cmp:cmp;
  });
  var nf={};
  indexed.forEach(function(it,ni){
    Object.keys(cellFmt).forEach(function(k){
      var p=k.split(',');
      if(+p[0]===it.i)nf[ni+','+p[1]]=cellFmt[k];
    });
  });
  cellFmt=nf;dupKeys={};
  var oldToNewSort=indexed.map(function(it){return it.i});
  var mapSort=[];indexed.forEach(function(it,ni){mapSort[it.i]=ni});
  for(var rr2=lastR+1;rr2<SHEET_ROWS;rr2++)mapSort[rr2]=rr2;
  remapMergesAfterRowReorder(mapSort);
  for(var i=0;i<indexed.length;i++)sheetRaw[i]=indexed[i].row;
  renderSheet();
}
function showFormulaHelp(){
  alert('Cách viết công thức: bắt đầu bằng dấu =. Ví dụ: =SUM(A1:A3) cộng các ô A1 tới A3; =AVERAGE, =COUNT (chỉ đếm ô có số), =MAX, =MIN; =A1+B1; =A1*2; =IF(A1>5,"Đạt","Chưa"); =ROUND(A1,2) làm tròn; =VLOOKUP(A1,C1:D5,2) tìm gần đúng theo cột đã sắp xếp (thêm ,FALSE để khớp chính xác); =COUNTIF(A1:A9,">5"); =SUMIF(A1:A9,"An",B1:B9). Liên trang: =Sheet2!A1, tên có dấu cách thì =\'Tên trang\'!A1. Gõ 50% thì giá trị thật là 0,5; gõ 1.000.000 ₫ thì giá trị thật là 1000000 — công thức luôn tính trên giá trị thật. Lỗi (chia cho 0, vòng tròn, sai cú pháp, không tìm thấy) đều hiện #LỖI. Gõ công thức vào thanh công thức (fx) rồi Enter, hoặc chạm hai lần vào ô để sửa.');
}
function triggerImportCsv(){var f=$('importFile');if(f)f.click();else gotoTab('files');}
function toggleSheetLock(){
  pushSheetHistory();sheetLocked=!sheetLocked;
  alert(sheetLocked?'Đã bảo vệ trang tính: không sửa được ô cho tới khi mở khóa.':'Đã mở khóa trang tính.');
  updateStatusBar();
}
/* PowerPoint: slide hien tai, nen, co, hieu ung, khoa, trinh chieu tu slide nay. */
function selectSlide(i){if(i<0||i>=deckSlides.length)return;deckTypingPushed=false;currentSlide=i;renderSlideEditor();updateStatusBar();}
function deleteCurrentSlide(){deleteSlide(currentSlide);}
function moveCurrentSlide(dir){
  var j=currentSlide+dir;if(j<0||j>=deckSlides.length)return;
  pushDeckHistory();
  var t=deckSlides[currentSlide];deckSlides[currentSlide]=deckSlides[j];deckSlides[j]=t;currentSlide=j;
  renderSlideEditor();updateStatusBar();
}
function setSlideBg(color){pushDeckHistory();if(deckSlides[currentSlide])deckSlides[currentSlide].bg=color;renderSlideEditor();}
function setSlideSize(r){pushDeckHistory();deckSize=r;renderSlideEditor();}
function setTransition(t){pushDeckHistory();deckTransition=t;renderSlideEditor();}
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
  var vis=visibleSlides();
  if(!vis.length){alert('Tất cả slide đang bị ẩn — hiện lại ít nhất 1 slide rồi trình chiếu.');return}
  presentSlides=JSON.parse(JSON.stringify(vis.map(function(v){return v.slide})));
  var pos=0;vis.forEach(function(v,i){if(v.orig<=currentSlide)pos=i});
  if(vis[pos].orig<currentSlide&&pos<vis.length-1&&vis[pos+1])pos=pos+1;
  /* Nếu slide hiện tại đang ẩn thì bắt đầu từ slide hiện gần nhất sau nó. */
  if(deckSlides[currentSlide]&&deckSlides[currentSlide].hidden){
    var nxt=-1;vis.forEach(function(v,i){if(v.orig>currentSlide&&nxt<0)nxt=i});
    pos=nxt>=0?nxt:vis.length-1;
  }
  presentIndex=pos;
  $('presentOverlay').style.display='flex';
  renderPresentSlide();startAutoAdvance();
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

/* ================= BẢN 1.4 — nút ribbon bổ sung (chỉ thêm, không sửa nút cũ) ================= */
/* ---------- Word: bảng tạm / cỡ chữ / đoạn / tìm-thay / chèn / chủ đề / cột / tham chiếu ---------- */
function docCut(){try{$('editor').focus();document.execCommand('cut')}catch(e){alert('Trình duyệt không cho cắt tự động — Bùi bôi đen rồi dùng phím tắt của máy.')}}
function docCopy(){try{$('editor').focus();document.execCommand('copy');alert('Đã sao chép đoạn đang chọn.')}catch(e){alert('Trình duyệt không cho sao chép tự động — Bùi bôi đen rồi dùng phím tắt của máy.')}}
function docPaste(){try{$('editor').focus();document.execCommand('paste')}catch(e){alert('Trình duyệt chặn dán tự động vì lý do bảo mật — Bùi bấm vào trang rồi dùng phím tắt Dán của máy (Ctrl+V).')}}
var storedDocFormat=null,formatWaiting=false;
function copyDocFormat(){
  var sel=window.getSelection();if(!sel||!sel.rangeCount)return;
  if(formatWaiting&&storedDocFormat){
    /* Bấm lần 2: dán định dạng đã lấy vào đoạn đang chọn. */
    var r=sel.getRangeAt(0);
    if(!r.collapsed){try{var sp=document.createElement('span');Object.keys(storedDocFormat).forEach(function(k){sp.style[k]=storedDocFormat[k]});var frag=r.extractContents();sp.appendChild(frag);r.insertNode(sp);}catch(e){}}
    formatWaiting=false;alert('Đã dán định dạng vào đoạn đang chọn.');return;
  }
  var node=sel.anchorNode;if(!node)return;
  var el=node.nodeType===1?node:node.parentElement;if(!el)return;
  var cs=getComputedStyle(el);
  storedDocFormat={fontFamily:cs.fontFamily,fontSize:cs.fontSize,fontWeight:cs.fontWeight,fontStyle:cs.fontStyle,textDecoration:cs.textDecoration,color:cs.color,backgroundColor:cs.backgroundColor};
  formatWaiting=true;
  alert('Đã lấy định dạng của đoạn đang chọn. Bôi đen đoạn khác rồi bấm lại nút này để dán định dạng.');
}
function changeFontSize(delta){
  var sel=window.getSelection();if(!sel||!sel.rangeCount)return;
  var r=sel.getRangeAt(0);
  if(r.collapsed){var ed=$('editor');var cur=parseFloat(getComputedStyle(ed).fontSize)||16;ed.style.fontSize=Math.max(8,cur+delta)+'px';return}
  var node=r.startContainer;var el=node.nodeType===1?node:node.parentElement;
  var cur2=el?parseFloat(getComputedStyle(el).fontSize)||16:16;
  wrapSpan({fontSize:Math.max(8,cur2+delta)+'px'});
}
function docBlockEl(){
  var sel=window.getSelection();if(!sel||!sel.rangeCount)return null;
  var node=sel.anchorNode;var el=node.nodeType===1?node:node.parentElement;
  while(el&&el!==$('editor')){
    var tg=el.tagName.toLowerCase();
    if(tg==='p'||tg==='h1'||tg==='h2'||tg==='h3'||tg==='li'||tg==='div')return el;
    el=el.parentElement;
  }
  return null;
}
function openParaBgSheet(){
  var colors=[['#ffffff','Không tô'],['#ffeb3b','Vàng'],['#ffc9c9','Hồng nhạt'],['#cfe8ff','Xanh nhạt'],['#e6f4ea','Xanh lá nhạt']];
  openSheet('Chọn màu nền đoạn',colors.map(function(c){
    return {label:'<span class="swatch" style="background:'+c[0]+'"></span>'+esc(c[1]),onClick:function(){var b=docBlockEl();if(b)b.style.backgroundColor=c[0]==='#ffffff'?'':c[0];}};
  }));
}
function toggleParaBorder(){
  var b=docBlockEl();if(!b)return;
  if(b.style.border)b.style.border='';else b.style.border='1px solid #999';
}
function clearFindMarks(){
  Array.from($('editor').querySelectorAll('mark.findmark')).forEach(function(m){
    var p=m.parentNode;while(m.firstChild)p.insertBefore(m.firstChild,m);p.removeChild(m);p.normalize();
  });
}
function findInDoc(){
  clearFindMarks();
  var term=prompt('Nhập từ cần tìm:');if(!term)return;
  var count=0,first=null;
  var walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT,null,false);
  var nodes=[];var n;while((n=walker.nextNode()))nodes.push(n);
  nodes.forEach(function(tn){
    var txt=tn.nodeValue;var idx=txt.indexOf(term);if(idx<0)return;
    var frag=document.createDocumentFragment();var rest=txt;
    while(true){
      var i=rest.indexOf(term);
      if(i<0){frag.appendChild(document.createTextNode(rest));break}
      if(i>0)frag.appendChild(document.createTextNode(rest.slice(0,i)));
      var mk=document.createElement('mark');mk.className='findmark';mk.textContent=term;frag.appendChild(mk);
      count++;if(!first)first=mk;
      rest=rest.slice(i+term.length);
    }
    tn.parentNode.replaceChild(frag,tn);
  });
  if(!count){alert('Không tìm thấy “'+term+'” trong tài liệu.');return}
  if(first)first.scrollIntoView({block:'center'});
  alert('Tìm thấy '+count+' chỗ có “'+term+'” (đã tô sáng vàng). Bấm Tìm lại để tìm từ khác.');
}
function replaceInDoc(){
  var term=prompt('Nhập từ cần thay:');if(!term)return;
  var rep=prompt('Thay bằng từ:');if(rep==null)return;
  var walker=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT,null,false);
  var nodes=[];var nn;while((nn=walker.nextNode()))nodes.push(nn);
  for(var k=0;k<nodes.length;k++){
    var tn=nodes[k];var idx=tn.nodeValue.indexOf(term);if(idx<0)continue;
    var before=tn.nodeValue.slice(0,idx),after=tn.nodeValue.slice(idx+term.length);
    var parent=tn.parentNode;var wasMark=parent&&parent.tagName==='MARK';
    tn.nodeValue=before+rep+after;
    if(wasMark){var gp=parent.parentNode;while(parent.firstChild)gp.insertBefore(parent.firstChild,parent);gp.removeChild(parent);gp.normalize();}
    var remain=0;var w2=document.createTreeWalker($('editor'),NodeFilter.SHOW_TEXT,null,false);var t2;while((t2=w2.nextNode())){var v=t2.nodeValue;var p=0;while((p=v.indexOf(term,p))>=0){remain++;p+=term.length}}
    alert('Đã thay 1 chỗ “'+term+'” thành “'+rep+'”. Còn '+remain+' chỗ nữa — bấm lại để thay tiếp.');updateStatusBar();return;
  }
  alert('Không còn chỗ nào có “'+term+'” nữa.');
}
function insertPageNumber(){
  /* Trang soạn là một trang giấy duy nhất nên số trang là nhãn tĩnh “Trang 1”
     chèn tại con trỏ, không phải trường số trang tự nhảy khi in nhiều trang. */
  insertHtmlAtCursor('<span>Trang 1</span>');
}
function insertDocHeader(){
  if($('editor').querySelector('.doc-header')){alert('Tài liệu đã có đầu trang rồi — sửa trực tiếp dòng đầu trang.');return}
  var txt=prompt('Nhập nội dung đầu trang:');if(!txt)return;
  $('editor').insertAdjacentHTML('afterbegin','<div class="doc-header">'+esc(txt)+'</div>');try{docNoteEdit()}catch(e){}
}
function insertDocFooter(){
  if($('editor').querySelector('.doc-footer')){alert('Tài liệu đã có chân trang rồi — sửa trực tiếp dòng chân trang.');return}
  var txt=prompt('Nhập nội dung chân trang:');if(!txt)return;
  $('editor').insertAdjacentHTML('beforeend','<div class="doc-footer">'+esc(txt)+'</div>');try{docNoteEdit()}catch(e){}
}
function insertDocLink(){
  var url=prompt('Nhập địa chỉ liên kết (ví dụ https://example.com):');if(!url)return;
  if(url.indexOf('://')<0)url='https://'+url;
  var sel=window.getSelection();var txt=(sel&&sel.toString())||'';
  var label=prompt('Nhập chữ hiển thị cho liên kết:',txt||url);if(!label)return;
  insertHtmlAtCursor('<a href="'+esc(url)+'" target="_blank" rel="noopener">'+esc(label)+'</a>');
}
function openSymbolSheet(){
  var chars=['©','®','™','°','±','×','÷','§','¶','•','–','—','‘','’','“','”','…','←','→','↑','↓','✓','★','♥','Ω','α','β','π','√','∞','≠','≤','≥','€','£','¥','₫'];
  openSheet('Chọn ký tự đặc biệt để chèn',chars.map(function(ch){
    return {label:ch+' — ký tự '+ch,onClick:function(){insertHtmlAtCursor(esc(ch))}};
  }));
}
function applyDocTheme(name){
  var themes={giay:['#ffffff','#1a1a1a','Carlito'],kem:['#fdf6e3','#5b4636','Noto Serif'],xanh:['#eaf3fb','#0b3d66','Be Vietnam Pro'],xanhla:['#e6f4ea','#14532d','Inter']};
  var t=themes[name]||themes.giay;
  setPageBg(t[0]);var ed=$('editor');ed.style.color=t[1];ed.style.fontFamily="'"+t[2]+"', -apple-system, 'Segoe UI', sans-serif";
  $('fontLabel').textContent=t[2];
}
function setThemeFontSize(pt){$('editor').style.fontSize=pt}
function setDocColumns(n){var ed=$('editor');ed.style.columnCount=n;ed.style.columnGap='32px'}
function insertColumnBreak(){insertHtmlAtCursor('<p style="break-before:column"></p>')}
function insertBookmark(){
  var name=prompt('Nhập tên đánh dấu (không dấu, không khoảng trắng):');if(!name)return;
  name=name.replace(/\s+/g,'-');
  insertHtmlAtCursor('<span id="bm-'+esc(name)+'">⚓ '+esc(name)+'</span>');
}
function insertCitation(){
  var tg=prompt('Nhập tác giả:');if(tg==null)return;
  var nam=prompt('Nhập năm xuất bản:');if(nam==null)return;
  var ten=prompt('Nhập tên tài liệu:');if(ten==null)return;
  insertHtmlAtCursor('<p>('+esc(tg)+', '+esc(nam)+'). '+esc(ten)+'.</p>');
}
function insertFigureList(){
  var imgs=$('editor').querySelectorAll('img');
  if(!imgs.length){alert('Tài liệu chưa có ảnh nào để lập danh mục hình. Chèn ảnh trước (tab Chèn → Ảnh), rồi bấm lại.');return}
  var html='<h2>Danh mục hình</h2>';
  Array.from(imgs).forEach(function(im,i){
    html+='<p>Hình '+(i+1)+' — '+(esc(im.getAttribute('alt')||'Ảnh chèn'))+'</p>';
  });
  insertHtmlAtCursor(html);
}
/* Theo dõi thay đổi: chặn ở beforeinput (trước khi chữ vào tài liệu) nên chữ
   gõ mới được bọc <ins>, chữ xóa được giữ lại dạng <del> — không dùng diff
   sau sự kiện nên không bịa lại vị trí xóa. Gõ bằng bộ gõ tiếng Việt (IME)
   được cho qua mặc định để không vỡ dấu. */
var trackChangesOn=false;
(function(){
  var ed=document.getElementById('editor');if(!ed)return;
  ed.addEventListener('beforeinput',function(e){
    if(!trackChangesOn)return;
    var t=e.inputType||'';
    if(t.indexOf('Composition')>=0)return;
    if(t==='insertText'||t==='insertFromPaste'||t==='insertFromDrop'){
      var txt='';
      if(t==='insertText')txt=e.data||'';
      else if(e.dataTransfer){try{txt=e.dataTransfer.getData('text/plain')||''}catch(err){}}
      if(!txt)return;
      e.preventDefault();document.execCommand('insertHTML',false,'<ins class="track-ins">'+esc(txt)+'</ins>');
      return;
    }
    if(t.indexOf('delete')===0){
      var sel=window.getSelection();if(!sel||!sel.rangeCount)return;
      var r=sel.getRangeAt(0);
      if(!r.collapsed){
        var delTxt=r.toString();if(!delTxt)return;
        e.preventDefault();r.deleteContents();document.execCommand('insertHTML',false,'<del class="track-del">'+esc(delTxt)+'</del>');
        return;
      }
      /* Xóa 1 ký tự khi không bôi đen: chỉ xử lý khi ký tự nằm trong cùng nút chữ. */
      try{
        var node=r.startContainer,off=r.startOffset;
        if(node.nodeType===3){
          if(t.indexOf('Backward')>=0&&off>0){
            var ch=node.nodeValue.charAt(off-1);e.preventDefault();
            var rr=document.createRange();rr.setStart(node,off-1);rr.setEnd(node,off);rr.deleteContents();
            document.execCommand('insertHTML',false,'<del class="track-del">'+esc(ch)+'</del>');
          }else if(t.indexOf('Forward')>=0&&off<node.nodeValue.length){
            var ch2=node.nodeValue.charAt(off);e.preventDefault();
            var rr2=document.createRange();rr2.setStart(node,off);rr2.setEnd(node,off+1);rr2.deleteContents();
            document.execCommand('insertHTML',false,'<del class="track-del">'+esc(ch2)+'</del>');
          }
        }
      }catch(err){}
    }
  });
  ed.addEventListener('input',function(){if(splitViewOn)syncSplitView()});
})();
function toggleTrackChanges(){
  trackChangesOn=!trackChangesOn;
  alert(trackChangesOn?'Đã bật Theo dõi thay đổi: chữ gõ mới hiện xanh gạch chân, chữ xóa giữ lại dạng gạch ngang.':'Đã tắt Theo dõi thay đổi: chữ gõ/xóa từ giờ trở lại bình thường (phần đã đánh dấu vẫn giữ nguyên).');
}
function acceptChanges(){
  var ed=$('editor');
  Array.from(ed.querySelectorAll('ins.track-ins')).forEach(function(el){var p=el.parentNode;while(el.firstChild)p.insertBefore(el.firstChild,el);p.removeChild(el)});
  Array.from(ed.querySelectorAll('del.track-del')).forEach(function(el){el.parentNode.removeChild(el)});
  try{docNoteEdit()}catch(e){}updateStatusBar();
}
function rejectChanges(){
  var ed=$('editor');
  Array.from(ed.querySelectorAll('ins.track-ins')).forEach(function(el){el.parentNode.removeChild(el)});
  Array.from(ed.querySelectorAll('del.track-del')).forEach(function(el){var p=el.parentNode;while(el.firstChild)p.insertBefore(el.firstChild,el);p.removeChild(el)});
  try{docNoteEdit()}catch(e){}updateStatusBar();
}
/* Chế độ đọc: ẩn ribbon + dải tab, chỉ còn trang giấy; nút Thoát hiện cố định. */
var readingOn=false;
function toggleReadingMode(){
  readingOn=!readingOn;
  /* Không gọi switchApp() ở đây: switchApp ghi đè body.className và sẽ
     xóa mất class 'reading' vừa bật (lỗi phát hiện khi kiểm khói 1.4). */
  document.body.classList.toggle('reading',readingOn);
  var b=$('exitReadingBtn');if(b)b.style.display=readingOn?'block':'none';
}
var splitViewOn=false;
function syncSplitView(){var sp=$('splitPane');if(sp)sp.innerHTML=$('editor').innerHTML}
function toggleSplitView(){
  splitViewOn=!splitViewOn;
  var sp=$('splitPane');if(!sp)return;
  sp.style.display=splitViewOn?'block':'none';
  document.body.classList.toggle('split-on',splitViewOn);
  if(splitViewOn)syncSplitView();
}

/* ---------- Excel 1.4: định dạng ô / bảng tạm / biểu đồ / lọc / tách cột ---------- */
var cellFmt={};
var wbSheets=[{name:'Sheet1',raw:null,fmt:null,merges:[]}];
var activeSheetIndex=0;
function currentSheetName(){return (wbSheets&&wbSheets[activeSheetIndex]&&wbSheets[activeSheetIndex].name)||'Sheet1'}
function syncCurrentToWb(){
  if(!wbSheets.length)wbSheets=[{name:'Sheet1',raw:sheetRaw,fmt:cellFmt,merges:sheetMerges}];
  if(activeSheetIndex<0)activeSheetIndex=0;
  if(activeSheetIndex>=wbSheets.length)activeSheetIndex=wbSheets.length-1;
  wbSheets[activeSheetIndex].raw=sheetRaw;
  wbSheets[activeSheetIndex].fmt=cellFmt;
  wbSheets[activeSheetIndex].merges=sheetMerges;
  wbSheets[activeSheetIndex].locked=sheetLocked;
  wbSheets[activeSheetIndex].freeze=freezeTopRow;
  wbSheets[activeSheetIndex].grid=gridOn;
}
function workbookSnapshotSheets(){
  syncCurrentToWb();
  return wbSheets.map(function(sh){
    return {name:sh.name,raw:sh.raw,fmt:sh.fmt||{},merges:normalizeMerges(sh.merges||[]),locked:!!sh.locked,freeze:!!sh.freeze,grid:sh.grid!==false};
  });
}
function loadSheetFromWb(idx){
  if(!wbSheets.length)return;
  if(idx<0)idx=0;if(idx>=wbSheets.length)idx=wbSheets.length-1;
  activeSheetIndex=idx;
  var sh=wbSheets[idx];
  sheetRaw=normalizeSheetRaw(sh.raw);
  cellFmt=sh.fmt||{};
  sheetMerges=normalizeMerges(sh.merges||[]);
  sheetLocked=!!sh.locked;freezeTopRow=!!sh.freeze;
  if(sh.grid!==undefined)gridOn=sh.grid!==false;
  dupKeys={};sheetFilter=null;
  var gw=$('sheetWrap');if(gw){if(gridOn)gw.classList.remove('nogrid');else gw.classList.add('nogrid')}
  var fw=$('sheetWrap');if(fw)fw.classList.toggle('freeze-first',freezeTopRow);
}
function switchSheet(idx){
  if(idx===activeSheetIndex)return;
  syncCurrentToWb();
  loadSheetFromWb(idx);
  renderSheetTabs();renderSheet();updateStatusBar();
}
function unusedSheetName(base){
  var names={};wbSheets.forEach(function(sh){names[String(sh.name).toLowerCase()]=true});
  if(!names[String(base).toLowerCase()])return base;
  var i=2;
  while(names[String(base+i).toLowerCase()])i++;
  return base+i;
}
/* Ban 1.7: Excel cam ky tu []:*?/\ trong ten trang (mo file se loi),
   nen chan ngay khi tao/doi ten, bao ro chu khong de file hong. */
function sheetNameBad(nm){return /[\[\]:*?\/\\]/.test(nm)}
function addSheet(name){
  if(name&&sheetNameBad(name)){alert('Tên trang tính không được chứa ký tự: [ ] : * ? / \\');return false}
  if(wbSheets.length>=20){alert('Đã đủ 20 trang tính.');return false}
  syncCurrentToWb();
  var nm=name||unusedSheetName('Sheet'+(wbSheets.length+1));
  pushSheetHistory();
  wbSheets.push({name:nm,raw:Array.from({length:SHEET_ROWS},function(){return Array(SHEET_COLS).fill('')}),fmt:{},merges:[],locked:false,freeze:false,grid:true});
  loadSheetFromWb(wbSheets.length-1);
  renderSheetTabs();renderSheet();updateStatusBar();
  return true;
}
function addSheetPrompt(){
  var nm=prompt('Nhập tên trang tính mới:','Sheet'+(wbSheets.length+1));
  if(nm==null)return;
  nm=nm.trim();if(!nm)nm=unusedSheetName('Sheet'+(wbSheets.length+1));
  var exists=wbSheets.some(function(sh){return String(sh.name).toLowerCase()===nm.toLowerCase()});
  if(exists){alert('Đã có trang tính tên “'+nm+'”. Đặt tên khác.');return}
  addSheet(nm);
}
function renameSheetPrompt(){
  var nm=prompt('Đổi tên trang tính “'+currentSheetName()+'” thành:',currentSheetName());
  if(nm==null)return;
  nm=nm.trim();if(!nm){alert('Tên trang tính không được để trống.');return}
  if(sheetNameBad(nm)){alert('Tên trang tính không được chứa ký tự: [ ] : * ? / \\');return}
  var exists=wbSheets.some(function(sh,i){return i!==activeSheetIndex&&String(sh.name).toLowerCase()===nm.toLowerCase()});
  if(exists){alert('Đã có trang tính tên “'+nm+'”. Đặt tên khác.');return}
  pushSheetHistory();
  wbSheets[activeSheetIndex].name=nm;
  renderSheetTabs();renderSheet();updateStatusBar();
}
function deleteSheetTab(){
  if(wbSheets.length<=1){alert('Bảng tính cần ít nhất 1 trang — không xóa trang cuối cùng này.');return}
  if(!confirm('Xóa trang tính “'+currentSheetName()+'” và toàn bộ nội dung của nó?'))return;
  pushSheetHistory();
  wbSheets.splice(activeSheetIndex,1);
  loadSheetFromWb(Math.min(activeSheetIndex,wbSheets.length-1));
  renderSheetTabs();renderSheet();updateStatusBar();
}
function renderSheetTabs(){
  var el=$('sheettabs');if(!el)return;
  var btns=wbSheets.map(function(sh,i){
    return '<button class="sheet-tab'+(i===activeSheetIndex?' active':'')+'" onclick="switchSheet('+i+')">'+esc(sh.name)+'</button>';
  }).join('');
  var note=(wbSheets.length>1)?('Bảng tính này có '+wbSheets.length+' trang tính.'):('Bảng tính này dùng một trang tính ('+esc(currentSheetName())+').');
  el.innerHTML=btns+'<button class="sheet-tab" onclick="addSheetPrompt()" title="Thêm trang tính">＋</button>'
    +'<button class="sheet-tab" onclick="renameSheetPrompt()">Đổi tên trang</button>'
    +'<button class="sheet-tab" onclick="deleteSheetTab()">Xóa trang này</button>'
    +'<span class="small" id="sheetTabNote">'+note+'</span>';
}

function cellFmtKey(r,c){return r+','+c}
function getCellFmt(r,c){var k=cellFmtKey(r,c);if(!cellFmt[k])cellFmt[k]={};return cellFmt[k]}
var cellClip=null;
function cellCut(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  cellClip={v:sheetRaw[selCell.r][selCell.c],fmt:JSON.parse(JSON.stringify(getCellFmt(selCell.r,selCell.c)))};
  pushSheetHistory();sheetRaw[selCell.r][selCell.c]='';delete cellFmt[cellFmtKey(selCell.r,selCell.c)];renderSheet();
}
function cellCopy(){cellClip={v:sheetRaw[selCell.r][selCell.c],fmt:JSON.parse(JSON.stringify(getCellFmt(selCell.r,selCell.c)))};alert('Đã sao chép ô '+cellRef(selCell.r,selCell.c)+'.')}
function cellPaste(){
  if(!cellClip){alert('Chưa sao chép ô nào. Chọn một ô rồi bấm Sao chép trước.');return}
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();sheetRaw[selCell.r][selCell.c]=cellClip.v;cellFmt[cellFmtKey(selCell.r,selCell.c)]=JSON.parse(JSON.stringify(cellClip.fmt||{}));renderSheet();
}
function toggleCellFmt(flag){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();var f=getCellFmt(selCell.r,selCell.c);
  if(flag==='b')f.b=!f.b;if(flag==='i')f.i=!f.i;if(flag==='u')f.u=!f.u;
  renderSheet();
}
function openCellColorSheet(kind){
  var colors=[['#1a1a1a','Đen'],['#c00000','Đỏ'],['#0b6fa4','Xanh dương'],['#1a7a2e','Xanh lá'],['#ffeb3b','Vàng'],['#ffc9c9','Hồng nhạt'],['#cfe8ff','Xanh nhạt'],['#e6f4ea','Xanh lá nhạt'],['#ffffff','Trắng']];
  openSheet(kind==='fore'?'Chọn màu chữ ô':'Chọn màu nền ô',colors.map(function(c){
    return {label:'<span class="swatch" style="background:'+c[0]+'"></span>'+esc(c[1]),onClick:function(){
      if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
      pushSheetHistory();var f=getCellFmt(selCell.r,selCell.c);
      if(kind==='fore')f.color=c[0];else f.bg=c[0];
      renderSheet();
    }};
  }));
}
function setCellAlign(a){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();getCellFmt(selCell.r,selCell.c).align=a;renderSheet()}
function setCellValign(v){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();getCellFmt(selCell.r,selCell.c).valign=v;renderSheet()}
function toggleCellWrap(){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();var f=getCellFmt(selCell.r,selCell.c);f.wrap=!f.wrap;renderSheet()}
function setCellNumFmt(kind){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();var f=getCellFmt(selCell.r,selCell.c);f.num=(f.num===kind)?null:kind;renderSheet()}
function changeCellDecimals(delta){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();var f=getCellFmt(selCell.r,selCell.c);var d=(f.dec==null)?(delta>0?1:0):f.dec;f.dec=Math.max(0,Math.min(6,d+delta));renderSheet()}
function clearCellFmt(){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();delete cellFmt[cellFmtKey(selCell.r,selCell.c)];renderSheet()}
function insertRowBelow(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();shiftCellFmtRows(selCell.r+1,1);shiftMergesRows(selCell.r+1,1);sheetRaw.splice(selCell.r+1,0,Array(SHEET_COLS).fill(''));sheetRaw.length=SHEET_ROWS;renderSheet();
}
function insertColRight(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushSheetHistory();shiftCellFmtCols(selCell.c+1,1);shiftMergesCols(selCell.c+1,1);
  sheetRaw.forEach(function(row){row.splice(selCell.c+1,0,'');row.length=SHEET_COLS;});
  renderSheet();
}
function insertCellLink(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var url=prompt('Nhập địa chỉ liên kết:');if(!url)return;
  if(url.indexOf('://')<0)url='https://'+url;
  pushSheetHistory();sheetRaw[selCell.r][selCell.c]=url;renderSheet();
}
function insertCellNote(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var old=getCellFmt(selCell.r,selCell.c).note||'';
  var txt=prompt('Nhập ghi chú cho ô '+cellRef(selCell.r,selCell.c)+':',old);if(txt==null)return;
  pushSheetHistory();if(txt)getCellFmt(selCell.r,selCell.c).note=txt;else delete getCellFmt(selCell.r,selCell.c).note;
  renderSheet();
}
function deleteCellNote(){if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}pushSheetHistory();delete getCellFmt(selCell.r,selCell.c).note;renderSheet()}
function hideSheetChart(){var c=$('sheetChart');if(c){c.style.display='none';c.innerHTML=''}}
function insertColumnChart(){
  var vals=[];
  for(var r=0;r<SHEET_ROWS;r++){var v=parseFloat(displayCellRaw(r,selCell.c));if(!isNaN(v))vals.push({label:'H'+(r+1),v:v})}
  if(!vals.length){alert('Cột '+colName(selCell.c)+' chưa có số nào để vẽ biểu đồ. Nhập số vào cột này rồi bấm lại.');return}
  var max=Math.max.apply(null,vals.map(function(x){return Math.abs(x.v)}))||1;
  var bars=vals.map(function(x){
    var h=Math.round(Math.abs(x.v)/max*120);
    return '<div class="chartbar-wrap"><div class="chartbar" style="height:'+h+'px" title="'+esc(String(x.v))+'"></div><span>'+esc(x.label)+'</span><span>'+esc(String(x.v))+'</span></div>';
  }).join('');
  var c=$('sheetChart');
  c.innerHTML='<h3>Biểu đồ cột — cột '+colName(selCell.c)+'</h3><div class="chartbars">'+bars+'</div><div class="row"><button onclick="hideSheetChart()">Ẩn biểu đồ</button></div>';
  c.style.display='block';
  c.scrollIntoView({block:'nearest'});
}
function displayCellRaw(r,c){return cellValue(r,c)}
var sheetFilter=null;
function filterSheet(){
  var txt=prompt('Chỉ hiện hàng mà cột '+colName(selCell.c)+' chứa chữ/số:');if(txt==null)return;
  sheetFilter={c:selCell.c,text:String(txt).toLowerCase()};
  renderSheet();
  alert(txt?('Đang lọc cột '+colName(selCell.c)+' theo “'+txt+'”. Bấm “Bỏ lọc” để hiện lại tất cả.'):'Đã bỏ lọc (ô tìm để trống).');
  if(!txt)sheetFilter=null;
}
function clearSheetFilter(){sheetFilter=null;renderSheet()}
function removeDuplicates(){
  pushSheetHistory();
  var seen={},out=[],kept=[];
  sheetRaw.forEach(function(row,i){
    var key=row.join('\u0001');
    if(key.trim()===''||!seen[key]){seen[key]=true;out.push(row);kept.push(i)}
  });
  while(out.length<SHEET_ROWS)out.push(Array(SHEET_COLS).fill(''));
  var nf2={};
  kept.forEach(function(oi,ni){
    Object.keys(cellFmt).forEach(function(k){
      var p=k.split(',');
      if(+p[0]===oi)nf2[ni+','+p[1]]=cellFmt[k];
    });
  });
  cellFmt=nf2;dupKeys={};
  var mapDup=[];kept.forEach(function(oi,ni){mapDup[oi]=ni});
  for(var rd=SHEET_ROWS-1;rd>=0;rd--)if(mapDup[rd]===undefined)mapDup[rd]=-1;
  remapMergesAfterRowReorder(mapDup);
  sheetRaw=out.slice(0,SHEET_ROWS);
  renderSheet();
  alert('Đã xóa các hàng trùng nhau hoàn toàn (giữ lại hàng đầu tiên).');
}
function splitCellColumn(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var v=sheetRaw[selCell.r][selCell.c]||'';
  var parts=String(v).split(/[,\s;]+/).filter(function(x){return x!==''});
  if(parts.length<2){alert('Ô đang chọn không tách được — cần ít nhất 2 phần cách nhau bằng dấu phẩy, khoảng trắng hoặc chấm phẩy.');return}
  pushSheetHistory();
  parts.forEach(function(p,i){if(selCell.c+i<SHEET_COLS)sheetRaw[selCell.r][selCell.c+i]=p});
  renderSheet();
}
var dupKeys={};
function highlightDuplicates(){
  if(Object.keys(dupKeys).length){dupKeys={};renderSheet();return}
  var counts={};
  for(var r=0;r<SHEET_ROWS;r++){var v=sheetRaw[r][selCell.c];if(v){counts[v]=(counts[v]||0)+1}}
  for(var r2=0;r2<SHEET_ROWS;r2++){var v2=sheetRaw[r2][selCell.c];if(v2&&counts[v2]>1)dupKeys[r2+','+selCell.c]=true}
  if(!Object.keys(dupKeys).length){alert('Cột '+colName(selCell.c)+' không có giá trị trùng nhau.');return}
  renderSheet();
}
var freezeTopRow=false,showFormulasFlag=false;
function toggleFreezeTopRow(){
  pushSheetHistory();freezeTopRow=!freezeTopRow;
  var w=$('sheetWrap');if(w)w.classList.toggle('freeze-first',freezeTopRow);
  alert(freezeTopRow?'Đã đóng băng hàng đầu: cuộn bảng, hàng 1 vẫn đứng yên dưới hàng tiêu đề cột.':'Đã bỏ đóng băng hàng đầu.');
}
function toggleShowFormulas(){showFormulasFlag=!showFormulasFlag;renderSheet()}
function insertColumnSum(){
  if(sheetLocked){alert('Trang tính đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var c=selCell.c;var last=-1;
  for(var r=0;r<SHEET_ROWS;r++){if(r!==selCell.r&&sheetRaw[r][c])last=r}
  if(last<0){alert('Cột '+colName(c)+' chưa có số nào để tính tổng.');return}
  pushSheetHistory();
  sheetRaw[selCell.r][selCell.c]='=SUM('+colName(c)+'1:'+colName(c)+(last+1)+')';
  renderSheet();
}

/* ---------- PowerPoint 1.4: phông / đoạn / hình dạng / biểu đồ / chủ đề / xem lại ---------- */
function curSlideObj(){return deckSlides[currentSlide]||emptySlide()}
function openSlideFontSheet(){
  openSheet('Chọn phông chữ cho slide này',FONTS.map(function(f){
    return {label:esc(f),onClick:function(){pushDeckHistory();curSlideObj().font=f;renderSlideEditor()}};
  }));
}
function toggleSlideTitleFmt(flag){
  pushDeckHistory();var s=curSlideObj();
  if(flag==='b')s.titleB=!s.titleB;if(flag==='i')s.titleI=!s.titleI;if(flag==='u')s.titleU=!s.titleU;
  renderSlideEditor();
}
function openSlideTitleColorSheet(){
  var colors=[['#1a1a1a','Đen'],['#c00000','Đỏ'],['#0b6fa4','Xanh dương'],['#1a7a2e','Xanh lá'],['#ffffff','Trắng']];
  openSheet('Chọn màu chữ tiêu đề',colors.map(function(c){
    return {label:'<span class="swatch" style="background:'+c[0]+'"></span>'+esc(c[1]),onClick:function(){pushDeckHistory();curSlideObj().titleColor=c[0];renderSlideEditor()}};
  }));
}
function setSlideAlign(a){pushDeckHistory();curSlideObj().align=a;renderSlideEditor()}
function toggleSlideBullets(){
  pushDeckHistory();var s=curSlideObj();s.bulletsOn=s.bulletsOn===false?true:false;renderSlideEditor();
  alert(s.bulletsOn?'Đã bật gạch đầu dòng cho slide này.':'Đã tắt gạch đầu dòng: nội dung trình chiếu hiện dạng đoạn văn.');
}
function insertSlideShape(type){
  if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  pushDeckHistory();var s=curSlideObj();if(!s.shapes)s.shapes=[];
  s.shapes.push({type:type});renderSlideEditor();
}
function removeLastSlideShape(){pushDeckHistory();var s=curSlideObj();if(s.shapes)s.shapes.pop();renderSlideEditor()}
function insertSlideChart(){
  if(deckLocked){alert('Bài trình bày đang được bảo vệ. Mở khóa ở tab Xem lại trước.');return}
  var txt=prompt('Nhập các số cho biểu đồ, cách nhau bằng dấu phẩy (ví dụ 5, 8, 3):');if(!txt)return;
  var vals=txt.split(',').map(function(x){return parseFloat(x)}).filter(function(x){return !isNaN(x)});
  if(!vals.length){alert('Chưa nhập được số nào.');return}
  pushDeckHistory();curSlideObj().chart=vals;renderSlideEditor();
}
function slideShapesHtml(s,forPresent){
  if(!s.shapes||!s.shapes.length)return '';
  return '<div class="slide-shapes">'+s.shapes.map(function(sh){
    if(sh.type==='rect')return '<div class="shape-rect"></div>';
    if(sh.type==='circle')return '<div class="shape-circle"></div>';
    return '<div class="shape-arrow">➜</div>';
  }).join('')+'</div>';
}
function slideChartHtml(s){
  if(!s.chart||!s.chart.length)return '';
  var max=Math.max.apply(null,s.chart.map(function(v){return Math.abs(v)}))||1;
  return '<div class="slide-chart">'+s.chart.map(function(v,i){
    return '<div class="chartbar-wrap"><div class="chartbar" style="height:'+Math.round(Math.abs(v)/max*100)+'px"></div><span>'+esc(String(v))+'</span></div>';
  }).join('')+'</div>';
}
function applySlideTheme(name){
  var themes={sang:['#ffffff','#2b579a'],kem:['#fdf6e3','#5b4636'],la:['#e6f4ea','#14532d'],toi:['#222222','#ffffff']};
  var t=themes[name]||themes.sang;
  pushDeckHistory();var s=curSlideObj();s.bg=t[0];s.titleColor=t[1];renderSlideEditor();
}
var pptSpellOn=true;
function toggleSlideSpellcheck(){
  pptSpellOn=!pptSpellOn;renderSlideEditor();
  alert(pptSpellOn?'Đã bật gạch chân cơ học của trình duyệt cho ô nội dung slide — LƯU Ý: chỉ là kiểm tra cơ học, CHƯA có từ điển tiếng Việt, không phải kiểm tra chính tả thật.':'Đã tắt gạch chân cơ học cho ô nội dung slide.');
}
function toggleSlideHidden(){
  pushDeckHistory();var s=curSlideObj();s.hidden=!s.hidden;renderSlideEditor();
  alert(s.hidden?'Đã ẩn slide này: khi trình chiếu sẽ bỏ qua slide này (vẫn còn trong bài, bấm lại để hiện).':'Đã hiện lại slide này.');
}
var autoAdvanceOn=false,presentTimer=null;
function startAutoAdvance(){
  if(presentTimer)clearInterval(presentTimer);
  presentTimer=null;
  if(!autoAdvanceOn)return;
  presentTimer=setInterval(function(){
    if(presentIndex<presentSlides.length-1)presentNext();
    else{clearInterval(presentTimer);presentTimer=null;}
  },5000);
}
function toggleAutoAdvance(){
  autoAdvanceOn=!autoAdvanceOn;
  if($('presentOverlay').style.display!=='none')startAutoAdvance();
  alert(autoAdvanceOn?'Đã bật tự chuyển slide mỗi 5 giây (áp dụng khi đang trình chiếu).':'Đã tắt tự chuyển slide.');
}
function visibleSlides(){
  var out=[];
  deckSlides.forEach(function(s,i){if(!s.hidden)out.push({slide:s,orig:i})});
  return out;
}
function insertSlideComment(){
  var s=curSlideObj();
  var txt=prompt('Nhập nhận xét cho slide này (để trống để xóa nhận xét cũ):',s.comment||'');if(txt==null)return;
  pushDeckHistory();s.comment=txt;renderSlideEditor();
}

/* ---------- Khởi động ---------- */
(function init(){
  var s=settings();
  $('googleClientId').value=s.googleClientId||'';$('msClientId').value=s.msClientId||'';$('ocrSpaceKey').value=s.ocrSpaceKey||'';
  googleToken=sessionStorage.getItem('vb_google_token')||null;
  renderDocList();renderSheetList();renderDeckList();renderSlideEditor();renderFileList();renderSheet();renderPhotos();updateAuthBar();
  try{renderRibbonTabs();switchApp('word');renderRuler();syncRuler();updateStatusBar();}catch(e){}
  if(window.addEventListener)window.addEventListener('resize',function(){syncRuler();layoutRibbon();});
  if(document.fonts&&document.fonts.ready)document.fonts.ready.then(function(){layoutRibbon()});
  var ed=$('editor');if(ed&&ed.addEventListener)ed.addEventListener('input',function(){try{docNoteEdit()}catch(e){}updateStatusBar();});
  try{docHistoryReset()}catch(e){}
  document.addEventListener('keydown',function(e){
    if(!(e.ctrlKey||e.metaKey))return;
    var k=(e.key||'').toLowerCase();
    if(k==='s'){e.preventDefault();saveCurrentApp();return;}
    if(k==='z'||k==='y'){
      var ae=document.activeElement;var tg=ae&&ae.tagName?ae.tagName.toLowerCase():'';
      if(tg==='input'||tg==='textarea')return;
      e.preventDefault();
      if(k==='z'){if(e.shiftKey)redo();else undo();}
      else redo();
      return;
    }
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
