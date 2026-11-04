"use strict";
/* ================= Fayldan avtomatik doldurma (protokol, plenum, müqavilə) =================
   Word (.docx), PDF (mətnli), .txt və .rtf fayllarının mətni oxunur, gündəlik, qərarlar, tarix və s. tapılır.
   Skan edilmiş səhifələr və şəkillər mətn tanıma ilə (ocr qovluğu) oxunur. */
const AZ_MONTHS=["yanvar","fevral","mart","aprel","may","iyun","iyul","avqust","sentyabr","oktyabr","noyabr","dekabr"];
function xmlText(x){return x.replace(/<w:tab\/>/g,"\t").replace(/<w:br\/>|<\/w:p>/g,"\n").replace(/<[^>]+>/g,"")
  .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(m,n)=>String.fromCharCode(+n)).replace(/&amp;/g,"&")}
function rtfText(s){return s.replace(/\\par[d]?\b/g,"\n").replace(/\{\\\*[^{}]*\}/g,"").replace(/\\u(-?\d+)\??/g,(m,n)=>String.fromCharCode(n<0?65536+ +n:+n))
  .replace(/\\'([0-9a-f]{2})/gi,(m,h)=>new TextDecoder("windows-1254").decode(new Uint8Array([parseInt(h,16)]))).replace(/\\[a-z]+-?\d* ?/gi,"").replace(/[{}]/g,"")}
/* Köhnə Word (.doc, Word 97–2003): «piece table» vasitəsilə mətni oxumaq */
function docText(bytes){const z=XLSX.CFB.read(bytes,{type:"array"});const wd=XLSX.CFB.find(z,"WordDocument");if(!wd)return null;
  const U=c=>c instanceof Uint8Array?c:Uint8Array.from(c);const W=U(wd.content),dv=new DataView(W.buffer,W.byteOffset,W.byteLength);const flags=dv.getUint16(0x0A,true);
  const tbl=XLSX.CFB.find(z,(flags&0x0200)?"1Table":"0Table");if(!tbl)return null;const Tb=U(tbl.content),tv=new DataView(Tb.buffer,Tb.byteOffset,Tb.byteLength);
  const fcClx=dv.getUint32(0x01A2,true),lcbClx=dv.getUint32(0x01A6,true);let i=fcClx,end=fcClx+lcbClx,out="";
  while(i<end&&Tb[i]===1){i+=3+tv.getUint16(i+1,true)}
  if(Tb[i]!==2)return null;const lcb=tv.getUint32(i+1,true);i+=5;const n=(lcb-4)/12;
  const cp=k=>tv.getUint32(i+4*k,true),pcd=k=>i+4*(n+1)+8*k;
  const dec1252=new TextDecoder("windows-1252");
  for(let k=0;k<n;k++){const len=cp(k+1)-cp(k);let fc=tv.getUint32(pcd(k)+2,true);
    if(fc&0x40000000){fc=(fc&~0x40000000)/2;out+=dec1252.decode(W.subarray(fc,fc+len))}
    else{let t="";for(let j=0;j<len;j++)t+=String.fromCharCode(dv.getUint16(fc+2*j,true));out+=t}}
  return out.replace(/\x13[^\x14\x15]*\x14?/g,"").replace(/[\x15\x01\x08]/g,"").replace(/\x07/g,"\t").replace(/[\r\x0b\x0c]/g,"\n")}
/* Bəzi PDF-lərdə ə, ğ, ş… başqa şriftlə ayrıca yazılır, əsas mətndə isə onların yerində boşluq qalır.
   Belə hərfləri x mövqeyinə görə ən yaxın boşluqlara (sıranı saxlamaqla) yerləşdiririk. */
const CW=c=>/[ilj.,:;'!|ıIİtfr]/.test(c)?0.45:c===" "?0.5:/[mwMW]/.test(c)?1.45:/[A-ZƏĞŞÇÖÜ]/.test(c)?1.25:1;
function charXs(t){const ws=[...t.s].map(CW),tot=ws.reduce((a,b)=>a+b,0)||1;let acc=0;return ws.map(w=>{const x=t.x+t.w*(acc+w/2)/tot;acc+=w;return x})}
const AZ_SPECIAL=/^[ƏəĞğİıŞşÇçÖöÜü\s]+$/;
function fixOverlay(r){const main=r.items.filter(t=>!AZ_SPECIAL.test(t.s));
  /* yalnız başqa parçanın ÜSTÜNƏ düşən hərf örtükdür; sözün öz yerində ayrıca parça olan hərfə («Ş»+«ə»+«mkir») toxunulmur */
  const ov=r.items.filter(t=>{if(!AZ_SPECIAL.test(t.s)||!/\S/.test(t.s))return false;const c=t.x+t.w/2;return main.some(q=>c>q.x&&c<q.x+q.w)});
  if(!ov.length||!main.length)return;
  const chars=[];for(const t of ov){const n=t.s.length;for(let i=0;i<n;i++)if(/\S/.test(t.s[i]))chars.push({c:t.s[i],x:t.x+t.w*(i+0.5)/n})}
  const slots=[];main.forEach(t=>{const xs=charXs({...t,s:t.s.replace(/ /g,"e")});for(let i=0;i<t.s.length;i++)if(t.s[i]===" ")slots.push({t,i,x:xs[i]})});
  chars.sort((a,b)=>a.x-b.x);slots.sort((a,b)=>a.x-b.x);const k=chars.length,m=slots.length;if(!k||k>m)return;
  const INF=1e18,D=Array.from({length:k+1},()=>new Float64Array(m+1).fill(INF));for(let j=0;j<=m;j++)D[0][j]=0;
  for(let i=1;i<=k;i++)for(let j=i;j<=m;j++)D[i][j]=Math.min(D[i][j-1],D[i-1][j-1]+Math.abs(chars[i-1].x-slots[j-1].x));
  const pick=[];for(let i=k,j=m;i>0;){if(D[i][j]===D[i][j-1]&&j>i)j--;else{pick.push([chars[i-1],slots[j-1]]);i--;j--}}
  const ed=new Map();for(const[c,sl]of pick){const a=ed.get(sl.t)||sl.t.s.split("");a[sl.i]=c.c;ed.set(sl.t,a)}
  for(const[t,a]of ed)t.s=a.join("");r.items=r.items.filter(t=>!ov.includes(t))}
/* ---------- Mətn tanıma: mühərrik yalnız lazım olanda yüklənir, internet tələb etmir ---------- */
let _ocr=null;
const ocrScript=src=>new Promise((res,rej)=>{const s=document.createElement("script");s.src=src;s.onload=res;s.onerror=()=>rej(new Error("yüklənmədi: "+src));document.head.append(s)});
function ocrEngine(){return _ocr||(_ocr=(async()=>{await ocrScript("ocr/tesseract-core-lstm.wasm.js");await ocrScript("ocr/aze.js");
  const M=await TesseractCore();M.FS.writeFile("aze.traineddata",Uint8Array.from(atob(window.__OCR_AZE),c=>c.charCodeAt(0)));window.__OCR_AZE=null;
  const api=new M.TessBaseAPI();if(api.Init(null,"aze"))throw new Error("Init");return{M,api}})().catch(e=>{_ocr=null;throw e}))}
async function ocrCanvas(cv){const{M,api}=await ocrEngine();const blob=await new Promise(r=>cv.toBlob(r,"image/png"));
  M.FS.writeFile("/input",new Uint8Array(await blob.arrayBuffer()));api.SetImageFile(1,0);const t=api.GetUTF8Text();api.Clear();return t||""}
const ocrNote=t=>{if(typeof toast==="function")toast(t);return new Promise(r=>setTimeout(r,60))};
const OCR_MAX=20,OCR_SCALE=300/72;
async function fileText(rel,name,maxPages=15){
  if(!window.api||!window.api.fileRead)return null;const ext=(String(name).split(".").pop()||"").toLowerCase();
  if(!["docx","doc","pdf","txt","rtf","png","jpg","jpeg"].includes(ext))return null;
  const raw=await window.api.fileRead(rel);if(!raw)return null;const bytes=raw instanceof Uint8Array?raw:new Uint8Array(raw);
  try{
    if(ext==="png"||ext==="jpg"||ext==="jpeg"){await ocrNote("Şəkil oxunur…");const bm=await createImageBitmap(new Blob([bytes]));const cv=document.createElement("canvas");cv.width=bm.width;cv.height=bm.height;cv.getContext("2d").drawImage(bm,0,0);return await ocrCanvas(cv)}
    if(ext==="docx"){const z=XLSX.CFB.read(bytes,{type:"array"});const f=XLSX.CFB.find(z,"word/document.xml")||XLSX.CFB.find(z,"/word/document.xml");
      return f?xmlText(new TextDecoder("utf-8").decode(f.content)):null}
    if(ext==="doc"){const t=docText(bytes);if(t)return t;if(String.fromCharCode(...bytes.slice(0,5))==="{\\rtf")return rtfText(new TextDecoder("latin1").decode(bytes));return null}
    if(ext==="pdf"){if(!window.pdfjsLib)return null;const task=pdfjsLib.getDocument({data:bytes,isEvalSupported:false});const doc=await task.promise;let out="",ocrN=0,ocrFail=false;
      /* mətn parçalarını mövqeyə görə sətirlərə yığırıq (ə, ğ kimi hərflər bəzən ayrıca parça olur) */
      for(let i=1;i<=Math.min(doc.numPages,maxPages);i++){const pg=await doc.getPage(i);const c=await pg.getTextContent();
        /* mətni olmayan səhifə skandır: şəkil kimi çəkib tanıyırıq */
        if(c.items.map(t=>t.str||"").join("").trim().length<20&&ocrN<OCR_MAX&&!ocrFail){ocrN++;
          try{await ocrNote(`Skan səhifə oxunur: ${i} / ${doc.numPages}…`);const vp=pg.getViewport({scale:OCR_SCALE}),cv=document.createElement("canvas");cv.width=vp.width;cv.height=vp.height;
            await pg.render({canvasContext:cv.getContext("2d"),viewport:vp}).promise;out+=await ocrCanvas(cv)+"\n\n"}catch(e){console.warn("ocr",e);ocrFail=true}
          continue}
        const it=c.items.filter(t=>t.str!==undefined&&t.str!=="").map(t=>({s:t.str,x:t.transform[4],y:t.transform[5],w:t.width,h:Math.abs(t.transform[3])||10}));
        const rows=[];for(const t of it.sort((a,b)=>b.y-a.y)){const r=rows.find(r=>Math.abs(r.y-t.y)<=Math.max(2,t.h*0.35));if(r)r.items.push(t);else rows.push({y:t.y,items:[t]})}
        rows.sort((a,b)=>b.y-a.y);
        for(const r of rows){fixOverlay(r);r.items.sort((a,b)=>a.x-b.x);let line="",end=null;
          for(const t of r.items){if(end!==null&&t.x-end>t.h*0.22&&!/\s$/.test(line)&&!/^\s/.test(t.s))line+=" ";line+=t.s;end=t.x+t.w}
          out+=line+"\n"}out+="\n"}
      try{await task.destroy()}catch(e){}return out}
    if(ext==="txt")return new TextDecoder("utf-8").decode(bytes);
    if(ext==="rtf")return rtfText(new TextDecoder("latin1").decode(bytes));
  }catch(e){console.warn("fileText",e)}return null}
/* «İ/ı» hərflərini «I/i» ilə əvəz edib axtarırıq (uzunluq dəyişmir, ona görə mətnin özündən kəsmək olur) */
function parseMinutes(text){
  /* «E Ş İ D İ L D İ», «Q Ə R A R A A L I R» kimi aralıqlı yazılmış başlıqları birləşdiririk */
  const joinSpaced=t=>t.replace(/(^|[\s:])((?:[A-ZƏÖÜĞİIŞÇ] ){2,}[A-ZƏÖÜĞİIŞÇ])(?=[\s:.]|$)/gm,(m,p,w)=>p+w.replace(/ /g,""));
  const src=joinSpaced(String(text||"").replace(/\r/g,"").replace(/[ \u00a0\t]+/g," ").replace(/\n\s*\n+/g,"\n")).replace(/QƏRARAAL/g,"QƏRARA AL");if(src.trim().length<20)return{};
  const T=src.replace(/İ/g,"I").replace(/ı/g,"i");const r={};
  const head=T.slice(0,900);
  let m=head.match(/(?:protokol|plenum)\w*\s*(?:№|N[oº°]?\.?|nömrə)\s*([0-9]+[0-9a-zа-я\/\-]*)/i)||head.match(/№\s*([0-9]+[0-9a-zа-я\/\-]*)/i);if(m)r.no=m[1];
  const dm=[];let x;
  const reD=/\b(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})\b/g;while((x=reD.exec(head)))dm.push({i:x.index,v:`${x[3]}-${x[2].padStart(2,"0")}-${x[1].padStart(2,"0")}`});
  const reM=/[«"“]?\s*(\d{1,2})\s*[»"”]?\s+([a-zəöüğşç]+)\s+(\d{4})/gi;while((x=reM.exec(head))){const k=AZ_MONTHS.findIndex(n=>x[2].toLowerCase().startsWith(n.slice(0,3)));if(k>=0)dm.push({i:x.index,v:`${x[3]}-${String(k+1).padStart(2,"0")}-${x[1].padStart(2,"0")}`})}
  dm.sort((a,b)=>a.i-b.i);if(dm[0]&&!isNaN(Date.parse(dm[0].v)))r.date=dm[0].v;
  const lines=src.split("\n").slice(0,12),TL=T.split("\n").slice(0,12);
  let bi=-1;for(const re of[/r.?yas.?t\s*hey.?t|icray.?\s*komit/i,/plenum/i,/komissiya/i,/iclas|yığıncaq|konfrans/i]){bi=TL.findIndex(l=>re.test(l)&&!/gündəlik/i.test(l));if(bi>=0)break}
  if(bi>=0)r.body=lines[bi].replace(/^\s*protokol\w*\s*(№\s*\S+)?/i,"").replace(/\s*(№|N)\s*\d+\S*/,"").replace(/[-–:,]\s*$/,"").trim().slice(0,160);
  m=T.match(/(\d{1,3})\s*nəfərdən\s*(\d{1,3})/i);if(m){r.total=m[1];r.present=m[2]}
  if(!r.present){m=T.match(/iştirak\s*ed\w*\s*[:\-–]?\s*(?:[^\n\d]{0,25})?(\d{1,3})/i);if(m)r.present=m[1]}
  if(!r.total){m=T.match(/(?:komitə|plenum|heyət)\w*\s*üzvlərinin\s*(?:ümumi\s*)?say\w*\s*[:\-–]?\s*(\d{1,3})/i);if(m)r.total=m[1]}
  const cut=(a,b)=>src.slice(a,b).replace(/^[\s:.\-–]+/,"").replace(/\s+$/,"");
  const STOP=/\n\s*(?:\d+\s*[.)].{0,2})?(?:dinl.{0,2}nildi|e.{0,2}idildi|.{0,2}.{0,2}x.{0,2}.{0,2}\s+ed|q.{0,2}rara.{0,2}al.{0,2}(?:nd.{0,2}|ir)|q.{0,2}rar\s+q.{0,2}bul|q.{0,2}rar\s*:|q.{0,2}rar\s+verildi|iclas.{0,2}n\s+s.{0,2}dri|s.{0,2}dr\s*:|katib|.{0,2}.{0,2}x.{0,2}.{0,2}lar)/i;
  m=T.match(/g.{0,2}nd.{0,2}lik\w*/i)||T.match(/(?:^|\n)\s*m.{0,2}s.{0,2}l.{1,2}(?:l.{0,2}r)?\s*:/i);if(m){const a=m.index+m[0].length;const rest=T.slice(a);const s2=rest.search(STOP);r.agenda=cut(a,s2>=0?a+s2:a+1500).slice(0,2000)}
  const dec=[];const reQ=/q.{0,2}rara.{0,2}al.{0,2}(?:nd.{0,2}|ir)|q.{0,2}rar\s+verildi|q.{0,2}rar\s*:/gi;
  while((x=reQ.exec(T))){const a=x.index+x[0].length;const rest=T.slice(a);const s2=rest.search(/\n\s*(?:\d+\s*[.)].{0,2})?(?:dinl.{0,2}nildi|e.{0,2}idildi|iclas.{0,2}n\s+s.{0,2}dri|s.{0,2}dr\s*:|katib|imza|q.{0,2}rar\s+q.{0,2}bul)/i);dec.push(cut(a,s2>=0?a+s2:a+1500))}
  if(dec.length)r.decisions=dec.filter(Boolean).map((d,i)=>dec.length>1?`${i+1}. ${d}`:d).join("\n").slice(0,4000);
  return r}
/* tapılanı yalnız boş xanalara yazırıq; orqan adını mövcud komissiyalarla tutuşdururuq */
const AF_LABEL={type:"növü",org:"ilk təşkilat",party:"digər tərəf",subject:"mövzusu",dateFrom:"qüvvəyə minmə",dateTo:"bitmə tarixi",regNo:"qeydiyyat",note:"qeyd",no:"№",date:"tarix",body:"orqan",present:"iştirakçılar",total:"üzvlərin sayı",agenda:"gündəlik",decisions:"qərarlar"};
function applyMinutes(rec,found,kind){const set=[];const L={no:"№",date:"tarix",body:"orqan",present:"iştirakçılar",total:"üzvlərin sayı",agenda:"gündəlik",decisions:"qərarlar"};
  if(found.body){const b=found.body.replace(/İ/g,"i").toLowerCase();
    const strip=t=>t.replace(/[əğışçöü\s]/g,"");const KW=["rəyasət","təftiş","maddi","gender","əməyin","gənc","fərdi","mədəni","hüquq"].map(strip),bs=strip(b);
    const hit=/plenum/.test(b)?"Rayon komitəsinin plenumu":(DB.bodies.find(x=>{const n=strip(x.name.replace(/İ/g,"i").toLowerCase());return KW.some(k=>n.includes(k)&&bs.includes(k))})||{}).name;
    if(hit)found.body=hit}
  const keys=REG[kind].fields.map(f=>f[0]);
  for(const k of Object.keys(L))if(found[k]&&keys.includes(k)&&!String(rec[k]||"").trim()){rec[k]=found[k];set.push(k)}
  return set}
/* ---------- Müqavilə: növü, №, tərəflər, məzmun, tarixlər, müddət ---------- */
const MON_RE="(yanvar|fevral|mart|aprel|may|iyun|iyul|avqust|sentyabr|oktyabr|noyabr|dekabr)";
const isoOf=(d,mo,y)=>`${y}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
/* «after» sözündən bilavasitə əvvəl gələn tarix */
function dateBefore(T,after){let m=T.match(new RegExp(`(\\d{1,2})\\D{0,6}?${MON_RE}\\D{0,10}(\\d{4})[^.]{0,90}?${after}`,"i"));
  if(m)return isoOf(m[1],AZ_MONTHS.indexOf(m[2].toLowerCase())+1,m[3]);
  m=T.match(new RegExp(`\\b(\\d{1,2})[.\\/](\\d{1,2})[.\\/](\\d{4})\\b[^.]{0,90}?${after}`,"i"));return m?isoOf(m[1],m[2],m[3]):""}
function addTerm(iso,n,unit){const[y,mo,d]=iso.split("-").map(Number);if(unit==="il")return isoOf(d,mo,y+n);
  const t=new Date(y,mo-1+n,d);return isoOf(t.getDate(),t.getMonth()+1,t.getFullYear())}
function parseContract(text){const src=String(text||"").replace(/[\s\u00a0]+/g," ").trim();if(src.length<20)return{};
  const T=src.replace(/İ/g,"I").replace(/ı/g,"i"),head=T.slice(0,1500),r={};const clean=x=>x.replace(/^[\s,:.\-–]+|[\s,:;\-–]+$/g,"");
  if(/kollektiv\s+müqavil/i.test(head))r.type="Kollektiv müqavilə";else if(/saziş/i.test(head))r.type="Saziş";else if(/əməkdaşl.q\s+müqavil/i.test(head))r.type="Əməkdaşlıq müqaviləsi";else if(/xidmət\s+müqavil/i.test(head))r.type="Xidmət müqaviləsi";
  let m=head.match(/müqavil\S*\s*(?:№|N[oeaº°2]\.?)\s*([0-9][0-9A-Za-z\/\-.]*)/i);if(m)r.no=m[1].replace(/[.\-]+$/,"");
  /* tərəflər: «X və Y arasında»; adın içində də «və» ola bilər («Gənclər və İdman»), ona görə Y-in X kimi başladığı «və» seçilir */
  const lw=t=>t.toLocaleLowerCase("az");
  const firstParty=t=>{t=clean(t).replace(/^(?:[^\p{L}]*\p{L}?[^\p{L}]+)+(?=\p{L}{2})/u,"");const P=t.split(/\s+və\s+/);if(P.length<2)return"";
    const nz=x=>lw(x).replace(/[^\p{L}\p{N}]+/gu,"");const st=nz(lw(P[0]).split(/\s+/).slice(0,2).join(" "));for(let i=1;i<P.length;i++)if(nz(P[i]).startsWith(st))return clean(P.slice(0,i).join(" və "));return clean(P[0])};
  let m2=/adland.r\S*\s+(.{3,420}?)\s+aras.nda/di.exec(T);
  if(m2){const pp=firstParty(src.slice(...m2.indices[1]));if(!/kollektiv\s+müqavil/i.test(pp))r.party=pp; /* səhifə sırası qarışıbsa, ad titul vərəqindən götürülür */
    const a=m2.indices[0][1],e=T.slice(a,a+1200).search(/\s+məqsədilə/i);if(e>0){const t=clean(src.slice(a,a+e)).replace(/(\p{L})- (\p{L})/gu,"$1-$2");if(t)r.subject=(t[0].toLocaleUpperCase("az")+t.slice(1)).slice(0,800)}}
  /* titul vərəqi: «X və Y arasında Kollektiv Müqavilə». Mətndəki ad ondan fərqlənirsə (şablondan köçürülmə), titul əsas götürülür */
  const cv=/^(.{3,420}?)\s+aras.nda\b.{0,12}?\s*kollektiv/di.exec(head);
  if(cv){const title=firstParty(src.slice(...cv.indices[1]));const W=t=>new Set(lw(t).replace(/[^\p{L}\p{N}]+/gu," ").trim().split(" "));
    if(title){if(!r.party)r.party=title;else{const A=W(title),B=W(r.party);if([...A].filter(w=>B.has(w)).length/Math.max(A.size,B.size)<0.5){r.conflict=r.party;r.party=title}}}}
  r.dateFrom=dateBefore(T,"q.vv.y[ə]?\\s+min");r.date=dateBefore(T,"(?:imzaland|bağland|imzalan)\\S*\\s+(?:və|tarix)")||r.dateFrom;
  m=T.match(/(\d+)\s*(?:\([^)]*\)\s*)?(il|ay)\s*müddətin[əd]\S*\s*(?:hüquqi\s*)?(?:qüvvəyə\s*malik|bağlan|qüvvədədir)/i);
  const st=r.dateFrom||r.date;if(m&&st)r.dateTo=addTerm(st,+m[1],m[2].toLowerCase());
  m=T.match(new RegExp(`(\\d{1,2})\\D{0,6}?${MON_RE}\\D{0,10}(\\d{4})\\s*il\\S*\\s*tarixli,?\\s*(\\S+)\\s*(?:№|N[oeaº°2])\\S*\\s*protokol`,"i"));
  if(m)r.note=`Təsdiq: ${m[4]} №-li protokol, ${m[1].padStart(2,"0")}.${String(AZ_MONTHS.indexOf(m[2].toLowerCase())+1).padStart(2,"0")}.${m[3]}`;
  return r}
/* digər tərəfin adını İlk təşkilatlarla tutuşdururuq; şübhə varsa boş qalır */
function matchOrg(party){const norm=t=>String(t||"").replace(/İ/g,"i").replace(/I/g,"ı").toLocaleLowerCase("az").replace(/[^\p{L}\p{N}]+/gu," ").trim();
  const P=new Set(norm(party).split(" "));let best=null,top=0,tie=false;
  for(const o of DB.orgs||[]){const W=norm(o.name).split(" ").filter(Boolean);if(!W.length)continue;const sc=W.filter(w=>P.has(w)).length/W.length;if(sc>top){top=sc;best=o;tie=false}else if(sc===top)tie=true}
  if(best&&top>=0.8&&!tie)return best.name;
  /* ad tapılmadı: təşkilatın «Açar sözlər»i, yalnız biri uyğun gələrsə */
  const n=norm(party),K=(DB.orgs||[]).filter(o=>String(o.keys||"").split(",").map(norm).some(k=>k.length>3&&n.includes(k)));return K.length===1?K[0].name:""}
function applyContract(rec,f){const set=[];const put=k=>{if(f[k]&&!String(rec[k]||"").trim()){rec[k]=f[k];set.push(k)}};
  if(f.type&&f.type!==rec.type&&(!rec.type||rec.type===CT_TYPES[0])){rec.type=f.type;set.push("type")}
  if(f.conflict)f.note=[f.note,`Diqqət: titul vərəqində «${f.party}», müqavilənin mətnində isə «${f.conflict}» yazılıb.`].filter(Boolean).join("\n");
  ["no","date","party","subject","dateFrom","dateTo","note"].forEach(put);
  if(f.party&&!rec.org){const o=matchOrg(f.party);if(o){rec.org=o;set.push("org")}}
  return set}
async function autoFillFrom(kind,rec,files){
  if(kind==="contracts"){let read=false,all=[];for(const f of files){const t=await fileText(f.rel,f.name,80);if(!t||t.trim().length<20)continue;read=true;all=all.concat(applyContract(rec,parseContract(t)))}return read?all:null}
  if(!["prots","plens"].includes(kind))return[];
  let read=false;for(const f of files){const t=await fileText(f.rel,f.name);if(!t||t.trim().length<20)continue;read=true;const got=applyMinutes(rec,parseMinutes(t),kind);if(got.length)return got}
  return read?[]:null}
