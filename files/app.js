"use strict";
const UI_VER=30; /* ekran hissəsinin versiyası — hər yeniləmədə artırılır */
/* ================= HİK-təşkilat — Həmkarlar İttifaqı rayon komitəsinin təşkilati işi ================= */
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const num=v=>{const n=Number(v);return isFinite(n)?n:0};
const newId=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,8);
const todayISO=()=>{const d=new Date();return`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`};
const dmy=s=>{if(!s)return"";const m=String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);return m?`${m[3]}.${m[2]}.${m[1]}`:String(s)};
const normName=s=>String(s??"").replace(/\s+/g," ").trim();
const lc=s=>normName(s).toLocaleLowerCase("az");
const nkey=s=>lc(s).replace(/\b(qızı|oğlu)\b/g,"").replace(/[.,]/g," ").split(" ").filter(Boolean).slice(0,3).join(" ");
const normFin=s=>String(s??"").toUpperCase().replace(/[^0-9A-Z]/g,"");
const plural=(n,w)=>`${n} ${w}`;
const MONTHS=["yanvar","fevral","mart","aprel","may","iyun","iyul","avqust","sentyabr","oktyabr","noyabr","dekabr"];
function xlDate(v){
  if(v==null||v==="")return"";
  if(typeof v==="number"&&window.XLSX){const d=XLSX.SSF.parse_date_code(v);if(d&&d.y)return`${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;return""}
  if(v instanceof Date&&!isNaN(v))return`${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  const s=String(v).trim();let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(m)return`${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  m=s.match(/^(\d{1,2})[.\/\-\s]+(\d{1,2})[.\/\-\s]+0?(\d{4})$/);if(!m)return"";
  let a=+m[1],b=+m[2],d,mo;if(a>12){d=a;mo=b}else if(b>12){mo=a;d=b}else{d=a;mo=b}
  if(mo<1||mo>12||d<1||d>31)return"";return`${m[3]}-${String(mo).padStart(2,"0")}-${String(d).padStart(2,"0")}`}
const yearsBetween=(a,b)=>{if(!a)return null;const[y1,m1,d1]=a.split("-").map(Number),[y2,m2,d2]=b.split("-").map(Number);return y2-y1-((m2<m1||(m2===m1&&d2<d1))?1:0)};
const ageOf=b=>yearsBetween(b,todayISO());
/* Təqaüd yaşı (65) tamam olanlar qırmızı göstərilir */
const RET_AGE=65,isRet=b=>{const x=ageOf(b);return x!=null&&x>=RET_AGE};
const retName=(n,b)=>isRet(b)?`<span class="retn" title="Təqaüd yaşı (${RET_AGE}) tamam olub — ${ageOf(b)} yaş">${n}</span> <span class="retb">${ageOf(b)} yaş · təqaüd</span>`:n;
try{if(!document.getElementById("retCss")){const st=document.createElement("style");st.id="retCss";st.textContent=".retn{color:#c62828!important;font-weight:700}.retb{display:inline-block;font-size:.72em;font-weight:600;color:#fff;background:#c62828;border-radius:9px;padding:0 7px;line-height:1.6;vertical-align:1px;white-space:nowrap}tr.ret td{background:rgba(198,40,40,.06)}.retd{color:#c62828;font-weight:700}@media print{.retn,.retd{color:#b00000!important}.retb{border:1px solid #b00000;color:#b00000;background:none}}";document.head.appendChild(st)}}catch(e){}
const fmtSize=n=>n>1048576?(n/1048576).toFixed(1)+" MB":Math.max(1,Math.round(n/1024))+" KB";

/* ---------- state ---------- */
const DEF_ORG="Mədəniyyət İşçiləri Həmkarlar İttifaqı Şəmkir Rayon Komitəsi";
/* İlk təşkilatlar: açar sözlər üzv siyahısındakı «Struktur bölmə» ilə tutuşdurulur (birinci uyğun gələn götürülür) */
const ORG_SEED=[
  ["Şəmkir rayon Mədəniyyət Mərkəzi","mədəniyyət mərkəz, mədəniyyət evi, folklor evi, klub"],
  ["Şəmkir rayon Mərkəzləşdirilmiş Kitabxana Sistemi","kitabxana, mks"],
  ["Şəmkir Uşaq Gənclər İdman İdarəsi və Məktəbi (UGİİ + UGİM)","uşaq gənclər idman, ugii, u.g.i.i, ugim, gənclər və idman"],
  ["Şəmkir şəhər Uşaq Musiqi Məktəbi","musiqi"],
  ["Şəmkir rayon Heydər Əliyev Mərkəzi","heydər əliyev"],
  ["Qədim Şəmkir Şəhər Dövlət Tarix-Mədəniyyət Qoruğu","qoruğ"],
  ["Şəmkir rayon Tarix Diyarşünaslıq Muzeyi","diyarşünaslıq muzey"],
  ["Şəmkir rayon Çənlibel Xalq Tətbiqi Sənət Muzeyi","çənlibel"],
  ["MİHİ Şəmkir Rayon Komitəsi","mihi, rayon komitəsi","Rayon komitəsi"]];
/* «Rayon komitəsi» — MİHİ RK aparatı: siyahıda nömrəsiz göstərilir və ilk təşkilatların sayına daxil edilmir */
const ORG_ST=["Fəaliyyətdədir","Qoşulacaq","Rayon komitəsi"];
const isRK=o=>o.state==="Rayon komitəsi";
const orgActive=o=>!isRK(o)&&(o.state||"Fəaliyyətdədir")!=="Qoşulacaq";
const BODY_ORDER=["Rayon komitəsinin üzvləri","Rəyasət Heyəti (İcrayə komitəsi)","Təftiş komissiyası","Maddi yardım üzrə komissiya","Gender bərabərliyi komissiyası","Əməyin mühafizəsi komissiyası","Gənclərlə iş üzrə komissiya","Fərdi əmək mübahisələri üzrə komissiya","Mədəni-kütləvi işlər üzrə komissiya","Hüquq məsələləri üzrə komissiya"];
let DB={v:1,org:DEF_ORG,chair:"",mem:[],bodies:[],orgs:[],events:[],prots:[],plens:[],contracts:[]};
const UI={tab:"home",mem:{view:"in",org:"",st:"",g:"",q:"",sel:null,draft:null},body:{sel:null,person:null},vet:{},
  orgs:{sel:null,q:""},events:{sel:null,q:"",y:""},prots:{sel:null,q:"",y:""},plens:{sel:null,q:"",y:""},contracts:{sel:null,q:"",y:""}};
const HAS_API=!!window.api;
const okYear=v=>!v||(v>="1900-01-01"&&v<="2100-12-31");
function load(){let d=null;try{d=HAS_API?window.api.load():JSON.parse(localStorage.getItem("hik-teskilat")||"null")}catch(e){}
  if(d)DB={...DB,...d};
  if(!DB.orgs.length)DB.orgs=ORG_SEED.map(([name,keys,state])=>({id:newId(),name,keys,state:state||"Fəaliyyətdədir",chair:"",phone:"",founded:"",note:"",files:[]}));
  /* köhnə versiya: MİHİ RK UGİİ ilə birləşdirilmişdisə — ayırmaq */
  const mix=DB.orgs.find(o=>/^MİHİ/.test(o.name)&&/ugii/.test(o.keys||""));
  if(mix){mix.keys="mihi, rayon komitəsi";mix.state="Rayon komitəsi";mix.name="MİHİ Şəmkir Rayon Komitəsi";if(/UGİİ/.test(mix.note||""))mix.note="";
    if(!DB.orgs.some(o=>/ugii/.test(o.keys||""))){const[n,k]=ORG_SEED[2];DB.orgs.splice(DB.orgs.indexOf(mix),0,{id:newId(),name:n,keys:k,state:"Fəaliyyətdədir",chair:"",phone:"",founded:"",note:"",files:[]})}}
  if(!DB.contracts)DB.contracts=[];
  /* Şəmkir RK-nın istəyi ilə bir dəfəlik: Excel vərəqindən gəlmiş köhnə xaric siyahısı silinir (əvvəl ehtiyat nüsxə). Başqa təşkilatlara toxunmur */
  if(!DB.outCut1&&/şəmkir/.test(lc(DB.org))){const keep=DB.mem.filter(m=>!(m.out&&/^Excel:/.test(m.outNote||"")));
    if(keep.length<DB.mem.length){if(HAS_API)window.api.snapshot("xaric-excel-silinmeden-once");DB.mem=keep}DB.outCut1=1;persist()}
  DB.mem.forEach(memDerive);
  /* il yazılarkən səhvən yadda qalmış tarixləri təmizləmək (məs. 0002-09-01) */
  for(const b of DB.bodies)if(b.elected&&!okYear(b.elected))b.elected=""}
function persist(){const j=JSON.stringify(DB);if(HAS_API){if(!window.api.save(j)){toast("Yadda saxlanmadı!");return false}}else{try{localStorage.setItem("hik-teskilat",j)}catch(e){}}return true}

/* ---------- lisenziya ---------- */
/* Telefon / brauzer versiyası (WEB): lisenziya açarı brauzerdə yoxlanılır, cihaz kodu brauzerdə saxlanılır */
const WEB=!HAS_API;if(WEB)document.documentElement.classList.add("web");
const LIC_PUB_RAW="+b7VMkUF8iAXePA9g8IRQFHFEVpOyhmMoxsh4ModHQA=";
const lsGetK=k=>{try{return localStorage.getItem(k)||""}catch(e){return""}},lsSetK=(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}};
function webMid(){let v=lsGetK("ht_mid");if(!/^[0-9A-F]{4}(-[0-9A-F]{4}){3}$/.test(v)){v=Array.from(crypto.getRandomValues(new Uint8Array(8)),x=>x.toString(16).padStart(2,"0")).join("").toUpperCase().match(/.{4}/g).join("-");lsSetK("ht_mid",v)}return v}
function webLic(key){const mid=webMid();let today=todayISO();const seen=lsGetK("ht_seen");if(seen>today)today=seen;else lsSetK("ht_seen",today);
  key=String(key||"").replace(/\s+/g,"");if(!key)return{status:"none",today,mid};
  const m=key.match(/^HTS1\.([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/);if(!m)return{status:"invalid",reason:/^HIK1\./.test(key)?"hik":"format",today,mid};
  const ub=s=>Uint8Array.from(atob(s.replace(/-/g,"+").replace(/_/g,"/")+"===".slice((s.length+3)%4)),c=>c.charCodeAt(0));
  let p;try{const data=ub(m[1]);if(!window.nacl||!nacl.sign.detached.verify(data,ub(m[2]),Uint8Array.from(atob(LIC_PUB_RAW),c=>c.charCodeAt(0))))return{status:"invalid",reason:"sign",today,mid};
    p=JSON.parse(new TextDecoder().decode(data))}catch(e){return{status:"invalid",reason:"format",today,mid}}
  if(p.p!=="HIK-teskilat"||!p.until)return{status:"invalid",reason:"format",today,mid};if(!p.mid)return{status:"invalid",reason:"nomid",today,mid};
  if(p.mid!==mid)return{status:"invalid",reason:"othermachine",today,mid};
  return{status:today>p.until?"expired":(p.from&&today<p.from?"notyet":"ok"),org:p.org||"",from:p.from||"",until:p.until,id:p.id||"",demo:!!p.demo,today,mid}}
let LIC=HAS_API?window.api.licGet():webLic(lsGetK("ht_lic"));
/* telefonda confirm() pəncərəsi işləməyə bilər: eyni düyməni ikinci dəfə basmaq təsdiq sayılır */
if(WEB){let pend=null,pt=0;window.confirm=msg=>{const now=Date.now();if(pend===msg&&now-pt<8000){pend=null;return true}pend=msg;pt=now;toast(String(msg).split("\n")[0]+" — təsdiq üçün düyməni yenidən basın");return false}}
/* mətn soruşmaq üçün səhifədaxili pəncərə (prompt() əvəzinə) */
function askText(title,def,cb){const md=document.createElement("div");md.className="lic-modal";
  md.innerHTML=`<div class="lic-box" role="dialog" aria-modal="true"><h2>${esc(title)}</h2><input id="askIn" class="ask-in" value="${esc(def||"")}"><div class="form-actions"><button class="btn primary" type="button" id="askOk">Yadda saxla</button><button class="btn" type="button" id="askNo">Ləğv et</button></div></div>`;
  document.body.append(md);const i=md.querySelector("#askIn");i.focus();i.select();const done=v=>{md.remove();if(v!=null&&v.trim())cb(v)};
  md.querySelector("#askOk").onclick=()=>done(i.value);md.querySelector("#askNo").onclick=()=>done(null);i.onkeydown=e=>{if(e.key==="Enter")done(i.value);if(e.key==="Escape")done(null)}}
const licOk=()=>LIC.status==="ok";
const licMsg=()=>LIC.status==="notyet"?`Lisenziya hələ qüvvəyə minməyib — ${dmy(LIC.from)}-dən başlayır.`:LIC.status==="expired"&&LIC.demo?"Demo müddəti bitib — dəyişiklik etmək bağlıdır. Tam lisenziya üçün AzerSoft LLC ilə əlaqə saxlayın.":LIC.status==="expired"?"Lisenziyanın müddəti bitib — dəyişiklik etmək bağlıdır.":"Proqram lisenziyasızdır — dəyişiklik etmək bağlıdır. «Parametrlər» bölməsində açarı daxil edin.";
function guard(){if(licOk())return true;toast(licMsg());return false}
function commit(msg){if(!guard())return false;const ok=persist();if(ok&&msg)toast(msg);render();return ok}

/* ---------- toast ---------- */
let tT;function toast(t){const el=$("#toast");el.textContent=t;el.classList.add("on");clearTimeout(tT);tT=setTimeout(()=>el.classList.remove("on"),3200)}

/* ---------- üzvlər: köməkçi ---------- */
const memKey=m=>normFin(m.fin)+"|"+nkey(m.name);
const active=()=>DB.mem.filter(m=>!m.out);
let _mIdx=null,_mIdxVer=-1,memVer=0;
function memByKey(name){if(_mIdxVer!==memVer){_mIdx=new Map();for(const m of DB.mem){const k=nkey(m.name);if(!_mIdx.has(k)||(_mIdx.get(k).out&&!m.out))_mIdx.set(k,m)}_mIdxVer=memVer}return _mIdx.get(nkey(name))||null}
/* Adın yazılışı bir az fərqlidirsə (Mehtiyev / Mehdiyev, Fərzail / Fərzayıl): ad eyni, soyad və ata adında cəmi 1–2 hərf fərq */
function lev(a,b){if(Math.abs(a.length-b.length)>2)return 9;const d=Array.from({length:a.length+1},(_,i)=>[i]);for(let j=1;j<=b.length;j++)d[0][j]=j;
  for(let i=1;i<=a.length;i++)for(let j=1;j<=b.length;j++)d[i][j]=Math.min(d[i-1][j]+1,d[i][j-1]+1,d[i-1][j-1]+(a[i-1]===b[j-1]?0:1));return d[a.length][b.length]}
let _fz=new Map(),_fzv=-1;
function memMatch(name,fin){const exact=memByKey(name);if(exact)return{m:exact,fuzzy:false};
  if(fin){const f=normFin(fin);const byFin=DB.mem.find(m=>m.fin===f);if(byFin)return{m:byFin,fuzzy:true}}
  if(_fzv!==memVer){_fz=new Map();_fzv=memVer}const k=nkey(name);if(_fz.has(k))return _fz.get(k);
  const[s0,...rest]=k.split(" ");let res=null;
  if(s0&&rest.length){const c=DB.mem.filter(m=>{const mk=nkey(m.name),[,...mr]=mk.split(" ");return mr[0]===rest[0]&&lev(mk,k)<=2});
    const act=c.filter(m=>!m.out);const one=act.length===1?act[0]:c.length===1?c[0]:null;if(one)res={m:one,fuzzy:true}}
  _fz.set(k,res);return res}
function orgOf(text){const t=lc(text);if(!t)return null;const head=lc(String(text).split(/[,\n]/)[0]);
  for(const pass of[head,t])for(const o of DB.orgs){const ks=String(o.keys||"").split(",").map(k=>lc(k)).filter(Boolean);if(ks.some(k=>pass.includes(k)))return o}
  return null}
/* Üzvə ilk təşkilat əl ilə (və ya təşkilatın öz Excel-i ilə) təyin olunubsa, o əsas götürülür; yoxsa struktur bölməyə görə */
const memOrg=m=>(m.orgId&&DB.orgs.find(o=>o.id===m.orgId))||orgOf(m.unit);
const orgName=m=>{const o=memOrg(m);return o?o.name:"Digər"};
function eduLevel(e){const t=lc(e);if(!t)return"Qeyd olunmayıb";if(/^ali|^magistr|^bakalavr/.test(t))return"Ali";if(/orta ixtisas|kollec|texnikum|tibb məktəbi|musiqi məktəbi|pedaqoji məktəb/.test(t))return"Orta ixtisas";
  if(/peşə/.test(t))return"Peşə";
  /* səviyyə yazılmayıb, ali məktəbin adı yazılıb («Naxçıvan Dövlət Universiteti») */
  if(!/natamam/.test(t)&&/universitet|institut|akademiya|konservatoriya/.test(t))return"Ali";if(/ümumi|orta|natamam/.test(t))return"Ümumi orta";return"Digər"}
const EXIT_BASIS=[
  "Əmək Məcəlləsi, maddə 69 — işçinin təşəbbüsü ilə (öz ərizəsi)",
  "Əmək Məcəlləsi, maddə 70 — işəgötürənin təşəbbüsü ilə",
  "Əmək Məcəlləsi, maddə 70, «a» bəndi — müəssisənin ləğvi",
  "Əmək Məcəlləsi, maddə 70, «b» bəndi — işçilərin sayının və ya ştatın ixtisarı",
  "Əmək Məcəlləsi, maddə 70, «c» bəndi — peşə hazırlığının vəzifəyə uyğun olmaması",
  "Əmək Məcəlləsi, maddə 70, «ç» bəndi — əmək vəzifələrinin yerinə yetirilməməsi / kobud pozulması",
  "Əmək Məcəlləsi, maddə 73 — əmək müqaviləsinin müddətinin bitməsi",
  "Əmək Məcəlləsi, maddə 74 — tərəflərin iradəsindən asılı olmayan hallar",
  "Həmkarlar ittifaqı üzvlüyündən öz ərizəsi ilə çıxma",
  "Pensiyaya çıxma",
  "Vəfat"];
const STATUSES=["Veteran","Qazi","Şəhid ailəsi","Əlil","Müharibə iştirakçısı","Çoxuşaqlı ana"];
/* «Statusu» xanasında bir neçə status ola bilər: «Qazi, ilk HİT Sədri» */
const stLow=s=>String(s||"").trim().toLocaleLowerCase("az");
const statusParts=st=>{const seen=new Set();return String(st||"").split(/[,;\/]+/).map(x=>x.trim()).filter(Boolean).map(p=>STATUSES.find(s=>stLow(s)===stLow(p))||p).filter(p=>{const k=stLow(p);if(seen.has(k))return false;seen.add(k);return true})};
const statusOthers=(st,cur)=>statusParts(st).filter(p=>stLow(p)!==stLow(cur));
function bodiesOf(m){const k=nkey(m.name),r=[];for(const b of DB.bodies)for(const p of b.people)if(nkey(p.name)===k)r.push({b,p});return r}
/* Təhsil qısa: «Ali», «Orta ixtisas», «Ümumi təhsil», «Kollec», «İlk peşə» … (mətnin vergülə qədər hissəsi) */
function eduShort(e){const t=normName(String(e||"").split(/[,;(]/)[0]);if(!t)return"";const l=lc(t);
  if(/^ali\b/.test(l))return"Ali";if(/^natamam ali/.test(l))return"Natamam ali";if(/^orta ixtisas/.test(l))return"Orta ixtisas";if(/^ilk peşə|^peşə/.test(l))return"İlk peşə";
  if(/^ümumi|^tam orta|^orta təhsil|^orta$/.test(l))return"Ümumi təhsil";if(/^qeyri-ali/.test(l))return"Qeyri-ali";if(/magistr/.test(l))return"Ali (magistr)";if(/bakalavr/.test(l))return"Ali (bakalavr)";
  /* səviyyə yazılmayıb, müəssisənin adı yazılıb: növünə görə */
  const all=lc(e);if(/peşə/.test(all))return"İlk peşə";if(/universitet|institut|akademiya|konservatoriya|incəsənət/.test(all))return"Ali";if(/kollec/.test(all))return"Kollec";if(/texnikum|tibb məktəbi|musiqi məktəbi|pedaqoji məktəb/.test(all))return"Orta ixtisas";
  return t.length>24?t.slice(0,22)+"…":t[0].toLocaleUpperCase("az")+t.slice(1)}
/* Orqanda vəzifəyə görə sıra: Sədr, Sədr müavini, Katib, sonra üzvlər (öz ardıcıllığı ilə) */
const roleRank=r=>{const l=lc(r);return/^sədr$/.test(l)?0:/^sədr(in)? müavin/.test(l)?1:/^katib/.test(l)?2:3};
const sortPeople=L=>L.map((p,i)=>[p,i]).sort((a,b)=>roleRank(a[0].role)-roleRank(b[0].role)||a[1]-b[1]).map(x=>x[0]);
/* Komissiya üzvlərinin boş xanalarını «Üzvlər» siyahısından doldurmaq */
function fillFromMembers(bodies){let n=0,p=0;for(const b of bodies)for(const x of b.people){const mm=memMatch(x.name,x.fin),m=mm&&mm.m;if(!m)continue;let ch=false;
    const set=(k,v)=>{if(!String(x[k]||"").trim()&&v){x[k]=v;n++;ch=true}};
    set("fin",m.fin);set("birth",m.birth);set("work",orgName(m)!=="Digər"?orgName(m):normName(String(m.unit||"").split(/[,\n]/)[0]));set("pos",m.pos);set("edu",eduShort(m.edu));set("phone",m.phone);if(ch)p++}
  return{n,p}}
const personView=p=>{const mm=memMatch(p.name,p.fin),m=mm&&mm.m;return{m,fuzzy:!!(mm&&mm.fuzzy),fin:p.fin||(m&&m.fin)||"",birth:p.birth||(m&&m.birth)||"",edu:p.edu||(m&&eduLevel(m.edu)!=="Qeyd olunmayıb"?m.edu:"")||"",
  work:p.work||(m&&m.unit)||"",pos:p.pos||(m&&m.pos)||""}};

/* ================= Excel idxal ================= */
function readWb(file){return file.arrayBuffer().then(b=>XLSX.read(b,{type:"array",cellDates:false}))}
/* Reyestr forması üçün əlavə sütunlar. Xanadakı cavab «Bəli»/«Xeyr» kimi oxunur; tanınmayan və boş xana boş qalır */
const MEM_YN=[["dis","Əlillik"],["war","Müharibə əlilliyi"],["shehid","Şəhid ailəsi"],["vet","Veteran"],["poor","Aztəminatlı ailə"],["aid","Ünvanlı sosial yardım"],["pens","Yaşa görə pensiya"]];
/* «Statusu» sütunundan reyestr xanaları: Şəhid ailəsi, Veteran (Qazi, müharibə iştirakçısı daxil), Əlil. Ayrıca yazılmış cavab dəyişdirilmir */
const MEM_ST=[["shehid",/şəhid/],["vet",/veteran|qazi|müharibə iştirak/],["war",/müharibə əlil/],["dis",/əlil/]];
function memDerive(m){const s=lc(m.status);for(const[k,re]of MEM_ST)if(!m[k])m[k]=re.test(s)?"Bəli":"Xeyr";return m}
const ynVal=v=>{const s=lc(v).replace(/[.!]+$/,"");if(/^(bəli|beli|hə|he|var|yes|\+|1)$/.test(s))return"Bəli";if(/^(xeyr|xeir|yox|no|-|–|—|0)$/.test(s))return"Xeyr";return""};
function parseMembers(wb){const out=[];
  for(const sn of wb.SheetNames){const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:""});
    const hi=rows.findIndex(r=>r&&r.some(c=>/soyad/i.test(String(c)))&&r.some(c=>/^\s*f[iİı]n\s*$/i.test(String(c))));if(hi<0)continue;
    const H=rows[hi].map(c=>lc(c));const col=re=>H.findIndex(h=>re.test(h));
    const C={no:col(/^№|^n[oº]/),unit:col(/struktur|bölmə|iş yeri/),pos:col(/vəzifə/),name:col(/soyad/),g:col(/cins/),fin:col(/^f[iı]n/),birth:col(/təvəllüd|doğum/),edu:col(/təhsil/),hired:col(/təyinat|işə qəbul/),st:H.findIndex(h=>/status/.test(h)&&!/üzvlük|əlil|şəhid|veteran/.test(h)),
      war:col(/müharibə.*əlil/),dis:H.findIndex(h=>/əlil/.test(h)&&!/müharibə/.test(h)),shehid:col(/şəhid/),vet:col(/veteran/),poor:col(/aztəminat/),aid:col(/ünvanlı|sosial yardım/),pens:col(/pensiya|müavinət/),phone:col(/telefon|əlaqə/),el:H.findIndex(isElHead),term:col(/seçil\S*\s*müddət|müddət/)};
    out.rbCols=[...new Set([...(out.rbCols||[]),...MEM_YN.filter(([k])=>C[k]>=0).map(x=>x[1]),...(C.phone>=0?["Telefon"]:[]),...(C.el>=0?["Seçilmə tarixi"]:[]),...(C.term>=0?["Seçilmə müddəti"]:[])])];
    /* Xaric olunanlar vərəqi yüklənmir: siyahı yeni yüklənən gündən başlayır, xaric yalnız proqramda qeyd olunur */
    if(/ləğv|xaric|çıx/i.test(sn)){out.rbSkip=[...(out.rbSkip||[]),sn];continue}
    const isOut=false;
    for(const r of rows.slice(hi+1)){const g=k=>C[k]>=0?r[C[k]]:"";const name=normName(g("name"));if(!name||/soyad/i.test(name))continue;
      const gen=lc(g("g"));
      /* yalnız doldurulmuş xanalar: boş xana əl ilə yazılmış məlumatı silmir */
      const extra={};for(const[k]of MEM_YN){const v=ynVal(g(k));if(v)extra[k]=v;else if(normName(g(k)))out.rbBad=(out.rbBad||0)+1}const tel=normName(g("phone"));if(tel)extra.phone=tel;const el=elDate(g("el"));if(el)extra.elected=el;const tm=normName(g("term"));if(tm)extra.term=/^\d+$/.test(tm)?tm+" il":tm;
      out.push({...extra,no:Number(g("no"))||"",unit:normName(g("unit")),pos:normName(g("pos")),name,gender:/^q/.test(gen)?"qadın":/^k/.test(gen)?"kişi":"",fin:normFin(g("fin")),
        birth:xlDate(g("birth")),edu:normName(g("edu")),hired:xlDate(g("hired")),status:normName(g("st")),out:isOut,outDate:"",outNote:isOut?"Excel: «"+sn+"» vərəqi":""})}}
  out.forEach(memDerive);return out}
const BODY_MAP=[[/komitə/i,0],[/rəyasət|icrayə/i,1],[/təftiş/i,2],[/maddi/i,3],[/gender/i,4],[/əməyin|mühafiz/i,5],[/gənc/i,6],[/fərdi|mübahisə/i,7],[/məd|kütl/i,8],[/hüquq/i,9]];
/* Seçilmə tarixi müxtəlif yazıla bilər: «15.03.2024», «15.03.2024-cü il», «15 mart 2024», «SEÇİLMƏ TARIXI» başlığı və s. */
const azFold=s=>lc(s).replace(/ı/g,"i").replace(/ə/g,"e").replace(/ş/g,"s").replace(/ç/g,"c").replace(/ğ/g,"g").replace(/ö/g,"o").replace(/ü/g,"u");
const isElHead=s=>{const t=azFold(s);return t.length<45&&(/tarix/.test(t)||/^secilib/.test(t))&&!/tevell|dogum|teyinat|qebul|bitme|bitdiyi|ise /.test(t)};
function elDate(v){if(v==null||v==="")return"";let d=xlDate(v);if(!d){const s=String(v);let m=s.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);if(m)d=xlDate(m[0]);else{m=lc(s).match(/(\d{1,2})\s+([a-zəöüğışç]+)\s+(\d{4})/);if(m){const mi=MONTHS.findIndex(x=>m[2].startsWith(x.slice(0,3)));if(mi>=0)d=`${m[3]}-${String(mi+1).padStart(2,"0")}-${m[1].padStart(2,"0")}`}}}return d&&okYear(d)?d:""}
function parseBodies(wb){const res=[];
  for(const sn of wb.SheetNames){const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:""});
    const data=rows.filter(r=>r&&typeof r[0]==="number"&&typeof r[1]==="string"&&r[1].trim());if(!data.length)continue;
    const title=normName((rows.slice(0,3).flat().find(c=>typeof c==="string"&&c.trim().length>8&&!/^№|soyad/i.test(c))||""));
    const mp=BODY_MAP.find(([re])=>re.test(sn));const name=mp?BODY_ORDER[mp[1]]:(title||sn);
    let elected="";const dm=title.match(/(\d{1,2})\s+([A-Za-zƏəÖöÜüĞğİıŞşÇç]+)\s+(\d{4})/);if(dm){const mi=MONTHS.indexOf(lc(dm[2]));if(mi>=0)elected=`${dm[3]}-${String(mi+1).padStart(2,"0")}-${dm[1].padStart(2,"0")}`}
    /* Reyestr üçün əlavə sütunlar başlığa görə tapılır: seçilmə tarixi, müddəti, cins, statuslar */
    const X={},firstData=rows.indexOf(data[0]);const XR=[["elected",/seçil\S*\s*tarix|seçki\s*tarix/],["term",/müddət/],["gender",/^cins/],["war",/müharibə.*əlil/],["dis",/əlil/],["shehid",/şəhid/],["vet",/veteran/],["poor",/aztəminat/],["aid",/ünvanlı|sosial yardım/],["pens",/pensiya|müavinət/]];
    let topEl="",topTerm="";rows.slice(0,firstData).forEach(row=>(row||[]).forEach((c,i)=>{if(typeof c!=="string"||!c.trim())return;const t=lc(c);const hit=isElHead(c)?XR[0]:XR.find(([,re])=>re.test(t));if(!hit)return;
      /* «Seçilmə tarixi: 15.03.2024» başlıqda və ya yanındakı xanada yazıla bilər */
      const after=normName(c.split(/[:：]/).slice(1).join(":"));const nx=row[i+1];
      if(hit[0]==="elected"&&(elDate(after)||elDate(c)||elDate(nx))){topEl=elDate(after)||elDate(c)||elDate(nx);return}
      if(hit[0]==="term"&&(after||/^\d+(\s*(il|ay))?$/i.test(normName(nx)))){topTerm=after||String(nx);return}
      if(!(hit[0] in X))X[hit[0]]=i}));
    /* Başlıq tanınmasa: bütün sətirlərdə eyni tarix yazılmış sütun seçilmə tarixidir (təvəllüd hamıda eyni olmur) */
    if(!("elected" in X)&&!topEl&&data.length>1){const w=Math.max(...data.map(r=>r.length));for(let i=4;i<w;i++){if(Object.values(X).includes(i))continue;const V=data.map(r=>elDate(r[i]));if(V.every(Boolean)&&new Set(V).size===1){X.elected=i;break}}}
    const skip=new Set(Object.values(X));const termOf=v=>{const s=normName(v);return/^\d+$/.test(s)?s+" il":s};
    const people=data.map(r=>{const p={id:newId(),name:normName(r[1]).replace(/\s+(qızı|oğlu)$/i,""),work:normName(r[2]),pos:normName(r[3]),phone:"",birth:"",fin:"",edu:"",role:""};
      if("elected" in X){const v=elDate(r[X.elected]);if(v)p.elected=v}if("term" in X){const v=termOf(r[X.term]);if(v)p.term=v}
      if("gender" in X){const g=lc(r[X.gender]);if(/^q/.test(g))p.gender="qadın";else if(/^k/.test(g))p.gender="kişi"}
      for(const k of["war","dis","shehid","vet","poor","aid","pens"])if(k in X){const v=ynVal(r[X[k]]);if(v)p[k]=v}
      for(const[ci,v]of r.map((v,i)=>[i,v]).slice(4)){if(skip.has(ci))continue;if(v===""||v==null)continue;
        if(typeof v==="number"){if(!p.birth)p.birth=xlDate(v);continue}
        const s=normName(v);
        if(/^(sədr|üzv|katib|sədr müavini|sədrin müavini)$/i.test(s))p.role=s[0].toLocaleUpperCase("az")+s.slice(1).toLocaleLowerCase("az");
        else if((s.match(/\d/g)||[]).length>=7&&/^[\d\s()+-]+$/.test(s))p.phone=s;
        else if(/^[0-9A-Z]{7}$/.test(s))p.fin=s;
        else if(/^(ali|orta|ümumi|peşə|natamam|magistr)/i.test(s))p.edu=s;
        else if(xlDate(s)&&!p.birth)p.birth=xlDate(s)}
      return p});
    const prev=DB.bodies.find(b=>b.name===name);
    if(!elected)elected=topEl||(people.find(p=>p.elected)||{}).elected||"";const fileTerm=termOf(topTerm)||(people.find(p=>p.term)||{}).term||"";
    if(prev)for(const p of people){const q=prev.people.find(x=>nkey(x.name)===nkey(p.name));if(q&&q.rbSent)p.rbSent=q.rbSent}
    res.push({id:prev?prev.id:newId(),name,elected:elected||(prev&&prev.elected)||"",note:"",people,...(prev&&prev.rbOrgan?{rbOrgan:prev.rbOrgan}:{}),...(fileTerm||(prev&&prev.rbTerm)?{rbTerm:fileTerm||prev.rbTerm}:{}),...(Object.keys(X).length?{xlCols:Object.keys(X)}:{})})}
  res.sort((a,b)=>{const i=BODY_ORDER.indexOf(a.name),j=BODY_ORDER.indexOf(b.name);return(i<0?99:i)-(j<0?99:j)});
  return res}

/* ================= render ================= */
const TABS=[["home","Ümumi"],["mem","Üzvlər"],["bodies","Komitə və komissiyalar"],["orgs","İlk təşkilatlar"],["contracts","Müqavilələr"],["events","Tədbirlər planı"],["prots","Protokollar"],["settings","Parametrlər"]];
/* Alt bölmələr: səhifənin yuxarısında düymə kimi */
const SUB={mem:[["mem","Üzvlər"],["vet","Veteran, Qazi, Şəhid ailəsi"]],prots:[["prots","Protokollar"],["plens","Plenumlar"]]};
const PARENT={vet:"mem",plens:"prots"};
function subNav(cnt){const g=SUB[PARENT[UI.tab]||UI.tab];if(!g)return"";
  return`<div class="subnav" role="tablist">${g.map(([k,l])=>`<button class="subtab" role="tab" data-tab="${k}" aria-selected="${UI.tab===k}">${l}${cnt[k]?` <span class="n">${cnt[k]}</span>`:""}</button>`).join("")}</div>`}
function render(){
  document.documentElement.dataset.sec=PARENT[UI.tab]||UI.tab;
  $("#hdrOrg").textContent=DB.org||DEF_ORG;
  const cnt={mem:active().length,bodies:DB.bodies.length,vet:active().filter(m=>m.status).length,orgs:DB.orgs.filter(orgActive).length,events:DB.events.length,prots:DB.prots.length,plens:DB.plens.length,contracts:DB.contracts.length};
  $("#tabs").innerHTML=TABS.map(([k,l])=>`<button class="tab" role="tab" data-tab="${k}" aria-selected="${UI.tab===k||PARENT[UI.tab]===k}">${l}${cnt[k]?`<span class="n">${cnt[k]}</span>`:""}</button>`).join("");
  licBar();
  const v=$("#view");
  const R={home:homeHTML,mem:memHTML,bodies:bodiesHTML,vet:vetHTML,orgs:()=>regHTML("orgs"),events:()=>regHTML("events"),prots:()=>regHTML("prots"),plens:()=>regHTML("plens"),contracts:()=>regHTML("contracts"),settings:settingsHTML}[UI.tab];
  v.innerHTML=subNav(cnt)+R();
  ({home:bindHome,mem:bindMem,bodies:bindBodies,vet:bindVet,orgs:()=>bindReg("orgs"),events:()=>bindReg("events"),prots:()=>bindReg("prots"),plens:()=>bindReg("plens"),contracts:()=>bindReg("contracts"),settings:bindSettings})[UI.tab]();
  addPrintBtns()}
/* lisenziya: qırmızı zolaq (yoxdur/bitib), sarı zolaq (30 gündən az qalıb), açılışda aktivləşdirmə pəncərəsi */
const licDays=()=>LIC.until&&LIC.today?Math.round((Date.parse(LIC.until)-Date.parse(LIC.today))/864e5):null;
const LIC_REASON={othermachine:"Bu açar başqa kompüter üçün yaradılıb.",hik:"Bu HİK-hesabatın açarıdır. HİK-təşkilat üçün ayrıca açar lazımdır.",nomid:"Açarda kompüter kodu yoxdur.",sign:"Açar etibarsızdır (imza uyğun gəlmir).",format:"Açar düzgün deyil — tam köçürüldüyünü yoxlayın."};
const midHTML=()=>LIC.mid?`<div class="lic-mid"><span class="mut">${WEB?"Bu cihazın kodu":"Bu kompüterin kodu"}</span><div class="lic-mid-row"><b class="mono">${esc(LIC.mid)}</b><button class="btn" type="button" data-copymid="${esc(LIC.mid)}">Kopyala</button></div></div>`:"";
function bindMid(root){root.querySelectorAll("[data-copymid]").forEach(b=>b.onclick=async()=>{try{await navigator.clipboard.writeText(b.dataset.copymid);toast("Kompüter kodu kopyalandı")}catch(e){toast("Kodu əl ilə köçürün: "+b.dataset.copymid)}})}
function licActivate(key){const k=String(key||"").trim();const r=HAS_API?window.api.licSet(k):webLic(k);if(r.status!=="invalid"){LIC=r;if(WEB)lsSetK("ht_lic",k.replace(/\s+/g,""))}return r}
let licDismissed=false;
function licBar(){let b=$("#licBar");const d=licDays(),warn=licOk()&&d!=null&&(d<=30||LIC.demo);
  if(licOk()&&!warn){b&&b.remove();$("#licModal")?.remove();return}
  if(!b){b=document.createElement("div");b.id="licBar";document.body.prepend(b)}
  b.className="lic-bar"+(warn?" warn":"");
  b.innerHTML=warn&&LIC.demo?`<b>Demo versiya</b> — bitməsinə <b>${d+1} gün</b> qalıb (${dmy(LIC.until)}-dək). Tam lisenziya üçün AzerSoft LLC ilə əlaqə saxlayın. <button class="btn" type="button" id="licGo">Açar daxil et</button>`:warn?`Lisenziyanın bitməsinə <b>${d} gün</b> qalıb (${dmy(LIC.until)}). Uzatmaq üçün AzerSoft LLC ilə əlaqə saxlayın.`:`${esc(licMsg())} <button class="btn" type="button" id="licGo">Lisenziya açarı daxil et</button>`;
  $("#licGo")?.addEventListener("click",()=>{UI.tab="settings";render();setTimeout(()=>$("#lic-key")?.focus(),50)});
  if((LIC.status==="none"||LIC.status==="invalid")&&!licDismissed&&!$("#licModal")){
    const md=document.createElement("div");md.id="licModal";md.className="lic-modal";
    md.innerHTML=`<div class="lic-box" role="dialog" aria-modal="true" aria-labelledby="licT"><img src="logo.png" alt="" style="width:72px;height:72px"><h2 id="licT">Proqramı aktivləşdirin</h2>
      <p class="mut">HİK-təşkilat AzerSoft LLC-yə məxsusdur. Lisenziya açarı almaq üçün <b>${WEB?"bu cihazın kodunu":"bu kompüterin kodunu"}</b> AzerSoft LLC-yə göndərin, aldığınız açarı aşağıya yapışdırın.</p>
      ${midHTML()}<textarea id="licModalKey" rows="4" spellcheck="false" placeholder="HTS1.…"></textarea><div class="form-msg" id="licModalMsg"></div>
      <div class="form-actions"><button class="btn primary" type="button" id="licModalGo">Aktivləşdir</button><button class="btn ghost" type="button" id="licModalLater">Yalnız baxış</button></div></div>`;
    document.body.append(md);bindMid(md);
    $("#licModalGo").onclick=()=>{const r=licActivate($("#licModalKey").value);
      if(r.status==="invalid"){const m=$("#licModalMsg");m.className="form-msg err";m.textContent=LIC_REASON[r.reason]||"Açar etibarsızdır.";return}
      md.remove();toast(r.status==="ok"?"Lisenziya aktivləşdi":r.status==="expired"?"Bu açarın müddəti artıq bitib":"Açar qəbul edildi");render()};
    $("#licModalLater").onclick=()=>{licDismissed=true;md.remove()}}}

/* ---------- Ümumi ---------- */
function barRows(list,total){const mx=Math.max(1,...list.map(x=>x[1]));
  return list.map(([l,n])=>`<tr><td>${esc(l)}</td><td class="num">${n}</td><td class="bar-c"><div class="hb"><i style="width:${(n/mx*100).toFixed(1)}%"></i></div></td><td class="num mut">${total?Math.round(n/total*100)+"%":""}</td></tr>`).join("")}
function homeHTML(){
  const A=active(),T=A.length,today=todayISO(),Y=+today.slice(0,4),M=today.slice(5,7);
  if(!DB.mem.length)return`<div class="card"><div class="empty"><b>Hələ məlumat yoxdur</b>Başlamaq üçün «Üzvlər» bölməsində üzvlərin Excel siyahısını, «Komitə və komissiyalar» bölməsində isə komissiyalar faylını yükləyin.
    <div class="form-actions" style="justify-content:center;margin-top:14px"><button class="btn primary" data-go="mem">Üzvlər siyahısını yüklə</button><button class="btn" data-go="bodies">Komissiyalar faylını yüklə</button></div></div></div>`;
  const q=A.filter(m=>m.gender==="qadın").length,k=A.filter(m=>m.gender==="kişi").length;
  const byOrg=new Map();for(const m of A){const n=orgName(m);byOrg.set(n,(byOrg.get(n)||0)+1)}
  const orgL=[...byOrg].sort((a,b)=>b[1]-a[1]);
  const ages=[["30 yaşadək",0,29],["30–39",30,39],["40–49",40,49],["50–59",50,59],["60 və yuxarı",60,200]].map(([l,a,b])=>[l,A.filter(m=>{const x=ageOf(m.birth);return x!=null&&x>=a&&x<=b}).length]);
  const noAge=A.filter(m=>ageOf(m.birth)==null).length;if(noAge)ages.push(["Təvəllüd yoxdur",noAge]);
  const eduM=new Map();for(const m of A){const l=eduLevel(m.edu);eduM.set(l,(eduM.get(l)||0)+1)}
  const eduL=["Ali","Orta ixtisas","Peşə","Ümumi orta","Digər","Qeyd olunmayıb"].filter(l=>eduM.get(l)).map(l=>[l,eduM.get(l)]);
  const jub=A.filter(m=>m.birth&&[50,55,60,65,70,75,80].includes(Y-+m.birth.slice(0,4))).sort((a,b)=>a.birth.slice(5).localeCompare(b.birth.slice(5)));
  const retL=A.filter(m=>isRet(m.birth)).sort((a,b)=>a.birth.localeCompare(b.birth));
  const bday=A.filter(m=>m.birth&&m.birth.slice(5,7)===M).sort((a,b)=>a.birth.slice(8).localeCompare(b.birth.slice(8)));
  const soon=DB.events.filter(e=>e.date>=today&&e.status!=="İcra olunub"&&e.status!=="Ləğv edilib").sort((a,b)=>a.date.localeCompare(b.date)).slice(0,6);
  const ctWarn=DB.contracts.map(c=>({c,s:ctState(c)})).filter(x=>x.s.d!=null&&x.s.d<=60).sort((a,b)=>a.s.d-b.s.d);
  const inBodies=new Set();DB.bodies.forEach(b=>b.people.forEach(p=>inBodies.add(nkey(p.name))));
  /* yoxlama */
  const finC=new Map();for(const m of DB.mem){if(m.fin)finC.set(m.fin,[...(finC.get(m.fin)||[]),m])}
  const dupF=[...finC].filter(([,l])=>l.length>1);
  const noFin=A.filter(m=>!m.fin),badFin=A.filter(m=>m.fin&&!/^[0-9A-Z]{7}$/.test(m.fin)),noBirth=A.filter(m=>!m.birth),noG=A.filter(m=>!m.gender);
  const notIn=[],fz=[];DB.bodies.forEach(b=>b.people.forEach(p=>{const mm=memMatch(p.name,p.fin);if(!mm)notIn.push({b,p});else if(mm.fuzzy&&nkey(mm.m.name)!==nkey(p.name))fz.push({b,p,m:mm.m})}));
  const outIn=[];DB.bodies.forEach(b=>b.people.forEach(p=>{const m=memByKey(p.name);if(m&&m.out)outIn.push({b,p})}));
  const chk=[[dupF.length,"Eyni FİN iki və daha çox şəxsdə",dupF.map(([f,l])=>`${f}: ${l.map(m=>esc(m.name)).join(", ")}`)],
    [notIn.length,"Komitə/komissiya üzvü üzv siyahısında tapılmadı (adın yazılışını yoxlayın)",notIn.map(x=>`${esc(x.p.name)} — ${esc(x.b.name)}`)],
    [fz.length,"Ad komissiyada və üzvlər siyahısında fərqli yazılıb («Adı düzəlt» ilə eyniləşdirin)",fz.map(x=>`${esc(x.p.name)} → ${esc(x.m.name)} — ${esc(x.b.name)}`)],
    [outIn.length,"Üzvlükdən xaric edilmiş şəxs komitə/komissiyadadır",outIn.map(x=>`${esc(x.p.name)} — ${esc(x.b.name)}`)],
    [noFin.length,"FİN yazılmayıb",noFin.map(m=>esc(m.name))],[badFin.length,"FİN 7 simvol deyil",badFin.map(m=>`${esc(m.name)} (${esc(m.fin)})`)],
    [noBirth.length,"Təvəllüd yazılmayıb",noBirth.map(m=>esc(m.name))],[noG.length,"Cinsi yazılmayıb",noG.map(m=>esc(m.name))]].filter(x=>x[0]);
  const mt=(l,v,s)=>`<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  const pl=m=>`<tr><td>${esc(m.name)}<div class="mut">${esc(orgName(m))}</div></td><td class="num">${dmy(m.birth)}</td><td class="num"><b>${Y-+m.birth.slice(0,4)}</b></td></tr>`;
  return`<div class="kpis">${mt("Üzvlər",T,`${DB.mem.length-T} nəfər xaric olunub`)}${mt("Qadın / kişi",`${q} / ${k}`,T?`qadınlar ${Math.round(q/T*100)}%`:"")}
    ${mt("İlk təşkilatlar",DB.orgs.filter(orgActive).length,`${new Set(A.map(m=>m.unit)).size} struktur bölmə`)}${mt("Seçkili orqanlarda",inBodies.size,`${DB.bodies.length} komitə və komissiya`)}</div>
  <div class="grid2">
    <section class="card printable" data-ptitle="Üzvlərin ilk təşkilatlar üzrə bölgüsü"><h2>İlk təşkilatlar üzrə</h2><div class="tbl-wrap"><table class="reg st-t"><thead><tr><th>Təşkilat</th><th class="num">Üzv</th><th></th><th class="num">Pay</th></tr></thead><tbody>${barRows(orgL,T)}</tbody></table></div></section>
    <div class="stack">
    <section class="card printable" data-ptitle="Üzvlərin yaş bölgüsü"><h2>Yaş üzrə</h2><table class="reg st-t"><tbody>${barRows(ages,T)}</tbody></table></section>
    <section class="card printable" data-ptitle="Üzvlərin təhsil üzrə bölgüsü"><h2>Təhsil üzrə</h2><table class="reg st-t"><tbody>${barRows(eduL,T)}</tbody></table></section></div>
    <section class="card printable" data-ptitle="${Y}-ci ildə yubileyi olan üzvlər"><h2>${Y}-ci ildə yubileylər <span class="badge">${jub.length}</span></h2><p class="hint">50, 55, 60, 65, 70, 75, 80 yaşı tamam olanlar.</p>
      <div class="tbl-wrap"><table class="reg"><thead><tr><th>S.A.A.</th><th class="num">Təvəllüd</th><th class="num">Yaş</th></tr></thead><tbody>${jub.map(pl).join("")||`<tr><td colspan="3" class="mut">Yoxdur</td></tr>`}</tbody></table></div></section>
    <section class="card printable" data-ptitle="Təqaüd yaşı (${RET_AGE}) tamam olan üzvlər"><h2 class="retd">Təqaüd yaşında (${RET_AGE}+) <span class="badge">${retL.length}</span></h2><p class="hint">${RET_AGE} yaşı tamam olmuş üzvlər — bütün siyahılarda qırmızı rənglə göstərilir.</p>
      <div class="tbl-wrap scroll-y"><table class="reg"><thead><tr><th>S.A.A.</th><th class="num">Təvəllüd</th><th class="num">Yaş</th></tr></thead><tbody>${retL.map(m=>`<tr class="ret" data-gomem="${m.id}" tabindex="0"><td><span class="retn">${esc(m.name)}</span><div class="mut">${esc(orgName(m))}</div></td><td class="num retd">${dmy(m.birth)}</td><td class="num retd">${ageOf(m.birth)}</td></tr>`).join("")||`<tr><td colspan="3" class="mut">Yoxdur</td></tr>`}</tbody></table></div></section>
    <section class="card printable" data-ptitle="Bu ay ad günü olan üzvlər"><h2>Bu ay ad günləri <span class="badge">${bday.length}</span></h2>
      <div class="tbl-wrap scroll-y"><table class="reg"><thead><tr><th>S.A.A.</th><th class="num">Təvəllüd</th><th class="num">Yaş</th></tr></thead><tbody>${bday.map(pl).join("")||`<tr><td colspan="3" class="mut">Yoxdur</td></tr>`}</tbody></table></div></section>
    <section class="card"><h2>Yaxın tədbirlər</h2>${soon.length?`<table class="reg"><tbody>${soon.map(e=>`<tr data-goev="${e.id}" tabindex="0"><td class="num">${dmy(e.date)}</td><td>${esc(e.title)}<div class="mut">${esc(e.resp||"")}</div></td></tr>`).join("")}</tbody></table>`:`<p class="mut">Planlaşdırılmış tədbir yoxdur. «Tədbirlər planı» bölməsində əlavə edin.</p>`}</section>
    <section class="card"><h2>Müqavilələr</h2>${ctWarn.length?`<table class="reg"><tbody>${ctWarn.map(({c,s})=>`<tr data-goct="${c.id}" tabindex="0"><td>${esc(c.type||"")}${c.no?" № "+esc(c.no):""}<div class="mut">${esc(c.org||c.party||"")}</div></td><td class="num nowrap">${dmy(c.dateTo)}<div><span class="st ${s.c}">${esc(s.t)}</span></div></td></tr>`).join("")}</tbody></table>`:`<p class="mut">${DB.contracts.length?`${DB.contracts.length} müqavilə; müddəti bitən və ya 60 gün ərzində bitəcək müqavilə yoxdur.`:"Müqavilə əlavə olunmayıb. «Müqavilələr» bölməsində kollektiv müqavilə və sazişləri faylları ilə birlikdə saxlaya bilərsiniz."}</p>`}</section>
    <section class="card"><h2>Məlumatların yoxlanması</h2>${chk.length?chk.map(([n,l,items])=>`<details class="chk"><summary><span class="st pend">${n}</span> ${l}</summary><ul>${items.slice(0,60).map(i=>`<li>${i}</li>`).join("")}${items.length>60?`<li>… və daha ${items.length-60}</li>`:""}</ul></details>`).join(""):`<p class="mut">Problem tapılmadı.</p>`}</section>
  </div>`}
function bindHome(){$$("[data-go]").forEach(b=>b.onclick=()=>{UI.tab=b.dataset.go;render()});
  $$("[data-goev]").forEach(r=>r.onclick=()=>{UI.tab="events";UI.events.sel=r.dataset.goev;render()});
  $$("[data-gomem]").forEach(r=>r.onclick=()=>{UI.tab="mem";Object.assign(UI.mem,{view:"in",org:"",st:"",g:"",q:"",sel:r.dataset.gomem});render()});
  $$("[data-goct]").forEach(r=>r.onclick=()=>{UI.tab="contracts";UI.contracts.sel=r.dataset.goct;render()})}

/* ---------- Üzvlər ---------- */
function memHTML(){
  const S=UI.mem,ed=S.sel&&DB.mem.find(m=>m.id===S.sel);
  const o=ed||S.draft||{name:"",gender:"",fin:"",birth:"",unit:"",pos:"",hired:"",edu:"",status:"",out:false,outDate:"",outNote:""};
  let L=DB.mem.filter(m=>S.view==="all"||(S.view==="out"?m.out:!m.out));
  if(S.org)L=L.filter(m=>orgName(m)===S.org);if(S.st)L=L.filter(m=>S.st==="—"?!m.status:statusParts(m.status).some(p=>stLow(p)===stLow(S.st)));if(S.g)L=L.filter(m=>m.gender===S.g);
  const q=lc(S.q);if(q)L=L.filter(m=>lc([m.name,m.fin,m.unit,m.pos,m.status,m.edu].join(" ")).includes(q));
  L.sort((a,b)=>((Number(a.no)||1e9)-(Number(b.no)||1e9))||a.name.localeCompare(b.name,"az"));
  const orgs=[...new Set(DB.mem.map(orgName))].sort((a,b)=>a.localeCompare(b,"az")),sts=[...new Set(DB.mem.flatMap(m=>statusParts(m.status)))].sort((a,b)=>a.localeCompare(b,"az"));
  const sel=(id,v,opts,all)=>`<select id="${id}"><option value="">${all}</option>${opts.map(x=>`<option${x===v?" selected":""}>${esc(x)}</option>`).join("")}</select>`;
  const shown=L.slice(0,600);
  const rows=shown.map((m,i)=>`<tr tabindex="0" data-id="${m.id}" class="${m.id===S.sel?"sel":""} ${m.out?"out":""} ${!m.out&&isRet(m.birth)?"ret":""}"><td class="num">${i+1}</td>
    <td>${m.out?esc(m.name):retName(esc(m.name),m.birth)}<div class="mut">${esc(m.fin)}${m.birth?` · <span class="${!m.out&&isRet(m.birth)?"retd":""}">${dmy(m.birth)}</span>`:""}${m.status?` · <b class="sst">${esc(m.status)}</b>`:""}${m.out?' · <span class="st rej">xaric</span>'+(m.outDate?" "+dmy(m.outDate):""):""}</div>${m.out&&m.outBasis?`<div class="mut">${esc(m.outBasis)}</div>`:""}</td>
    <td>${esc(m.unit)}<div class="mut">${esc(m.pos)}</div></td><td class="num">${dmy(m.hired)}</td></tr>`).join("");
  const B=ed?bodiesOf(ed):[];
  return`<div class="split">
  <section class="card${ed?" printable":""}" data-ptitle="Üzvün kartı — ${esc(o.name)}"><h2>${ed?"Üzvün kartı":"Yeni üzv"}</h2>
    <p class="hint">${ed?"Məlumatları dəyişib «Yadda saxla» düyməsini basın.":"Siyahıdan şəxsi seçin və ya yeni üzv əlavə edin."}</p>
    <form class="entry" id="mForm" novalidate>
      <div class="fld"><label for="m-name">Soyadı, adı, ata adı *</label><input id="m-name" name="name" value="${esc(o.name)}" autocomplete="off"></div>
      <div class="row2"><div class="fld"><label for="m-fin">FİN</label><input id="m-fin" name="fin" value="${esc(o.fin)}" maxlength="7" class="fin"></div>
        <div class="fld"><label for="m-gender">Cinsi</label><select id="m-gender" name="gender"><option value=""></option>${["qadın","kişi"].map(g=>`<option${o.gender===g?" selected":""}>${g}</option>`).join("")}</select></div></div>
      <div class="row2"><div class="fld"><label for="m-birth">Təvəllüdü</label><input id="m-birth" name="birth" type="date" value="${esc(o.birth)}"><span class="help${isRet(o.birth)?" retd":""}">${o.birth?ageOf(o.birth)+" yaş"+(isRet(o.birth)?" — təqaüd yaşı tamam olub":""):""}</span></div>
        <div class="fld"><label for="m-hired">Təyinat tarixi</label><input id="m-hired" name="hired" type="date" value="${esc(o.hired)}"></div></div>
      <div class="fld"><label for="m-unit">Struktur bölmə</label><textarea id="m-unit" name="unit" rows="2">${esc(o.unit)}</textarea></div>
      <div class="fld"><label for="m-orgId">İlk təşkilat</label><select id="m-orgId" name="orgId"><option value="">Avtomatik (struktur bölməyə görə)${!o.orgId&&o.unit?": "+esc(orgName({unit:o.unit})):""}</option>${DB.orgs.map(x=>`<option value="${x.id}"${o.orgId===x.id?" selected":""}>${esc(x.name)}</option>`).join("")}</select></div>
      <div class="fld"><label for="m-pos">Vəzifə</label><input id="m-pos" name="pos" value="${esc(o.pos)}"></div>
      <div class="fld"><label for="m-edu">Təhsili, təhsil müəssisəsi, ixtisası</label><textarea id="m-edu" name="edu" rows="2">${esc(o.edu)}</textarea></div>
      <div class="fld"><label for="m-status">Statusu</label><input id="m-status" name="status" list="stList" value="${esc(o.status)}" placeholder="Veteran, Qazi, Şəhid ailəsi…"><datalist id="stList">${STATUSES.map(s=>`<option>${s}</option>`).join("")}</datalist></div>
      <details class="m-rb"${MEM_YN.some(([k])=>o[k])||o.phone?" open":""}><summary>Reyestr üçün məlumatlar</summary><div class="fld"><label for="m-phone">Telefon</label><input id="m-phone" name="phone" value="${esc(o.phone||"")}"></div>
        <div class="row2 m-rb-g">${MEM_YN.map(([k,l])=>`<div class="fld"><label for="m-${k}">${l}</label><select id="m-${k}" name="${k}"><option value=""></option>${RB_YN.map(x=>`<option${o[k]===x?" selected":""}>${x}</option>`).join("")}</select></div>`).join("")}</div></details>
      <label class="chkl"><input type="checkbox" name="out" id="m-out"${o.out?" checked":""}> Üzvlükdən xaric edilib</label>
      <div class="entry" id="m-outBox"${o.out?"":" hidden"}>${(()=>{const inL=EXIT_BASIS.includes(o.outBasis||"");const other=!!o.outBasis&&!inL;return`
        <div class="fld"><label for="m-outBasis">Əmək Məcəlləsinin hansı bəndinə uyğun olaraq</label><select id="m-outBasis" name="outBasisSel"><option value="">— seçin —</option>${EXIT_BASIS.map(b=>`<option${o.outBasis===b?" selected":""}>${esc(b)}</option>`).join("")}<option value="__other"${other?" selected":""}>Digər (özünüz yazın)…</option></select>
          <input id="m-outOther" name="outBasisOther" value="${esc(other?o.outBasis:"")}" placeholder="məs.: maddə 70, «d» bəndi"${other?"":" hidden"} style="margin-top:6px"></div>`})()}
        <div class="row2"><div class="fld"><label for="m-outDate">Xaric tarixi</label><input id="m-outDate" name="outDate" type="date" value="${esc(o.outDate)}"></div>
        <div class="fld"><label for="m-outNote">Qeyd (əmr №, tarix)</label><input id="m-outNote" name="outNote" value="${esc(o.outNote)}"></div></div></div>
      ${ed?`<div class="fld"><label>Seçkili orqanlarda</label>${B.length?`<ul class="blist">${B.map(x=>`<li><a href="#" data-body="${x.b.id}">${esc(x.b.name)}</a>${x.p.role?` — <b>${esc(x.p.role)}</b>`:""}</li>`).join("")}</ul>`:`<span class="mut">Yoxdur</span>`}</div>`:""}
      <div class="form-msg" id="mMsg"></div>
      <div class="form-actions"><button class="btn primary" type="submit">${ed?"Yadda saxla":"Əlavə et"}</button>${ed?`<button class="btn" type="button" id="mClose">Bağla</button><button class="btn ghost" type="button" id="mDel">Sil</button>`:""}</div>
    </form></section>
  <section class="card printable" data-ptitle="Həmkarlar ittifaqı üzvlərinin siyahısı"><h2>Üzvlərin siyahısı</h2>
    <div class="reg-tools">
      <select id="mView"><option value="in"${S.view==="in"?" selected":""}>Üzvlər (${active().length})</option><option value="out"${S.view==="out"?" selected":""}>Xaric olunanlar (${DB.mem.length-active().length})</option><option value="all"${S.view==="all"?" selected":""}>Hamısı (${DB.mem.length})</option></select>
      ${sel("mOrg",S.org,orgs,"Bütün ilk təşkilatlar")}${sel("mSt",S.st,["—",...sts],"Bütün statuslar")}${sel("mG",S.g,["qadın","kişi"],"Cinsi: hamısı")}
      <input id="mq" class="srch" type="search" placeholder="Axtar: ad, FİN, vəzifə…" value="${esc(S.q)}"></div>
    <div class="reg-tools"><button class="btn" id="mImp" title="Üzvlər Excel faylı. «Ləğv edilənlər» vərəqi yüklənmir">Excel-dən yüklə</button><input type="file" id="mFile" accept=".xlsx,.xls" hidden>${DB.mem.length?`<button class="btn danger" id="mClear" title="Bütün üzv siyahısını silir. Sonra yeni Excel faylını yükləyə bilərsiniz">Siyahını sil</button>`:""}
      ${HAS_API&&window.api.hikMembers?`<button class="btn" id="mHik" title="Bu kompüterdəki HİK-hesabat proqramının «Uçot kartı» siyahısından üzvləri götürür">HİK-hesabatdan götür</button>`:""}
      <button class="btn" id="mExp">Excel-ə çıxar</button></div>
    <div class="sum"><b>${L.length}</b> nəfər${S.org?" · "+esc(S.org):""}${L.length>shown.length?` · ilk ${shown.length} göstərilir, axtarışla daraldın`:""}</div>
    <div class="tbl-wrap"><table class="reg"><thead><tr><th class="num">№</th><th>S.A.A.</th><th>Struktur bölmə / vəzifə</th><th class="num">Təyinat</th></tr></thead><tbody>
    ${rows||`<tr><td colspan="4"><div class="empty"><b>${DB.mem.length?"Uyğun qeyd yoxdur":"Siyahı boşdur"}</b>${DB.mem.length?"":"«Excel-dən yüklə» ilə üzvlərin siyahısını yükləyin."}</div></td></tr>`}</tbody></table></div></section></div>`}
function readMemForm(){const f=new FormData($("#mForm"));const g=k=>String(f.get(k)??"").trim();
  return rbStFix({name:normName(g("name")),fin:normFin(g("fin")),gender:g("gender"),birth:g("birth"),hired:g("hired"),unit:normName(g("unit")),pos:normName(g("pos")),edu:normName(g("edu")),status:normName(g("status")),orgId:g("orgId"),phone:normName(g("phone")),dis:g("dis"),war:g("war"),shehid:g("shehid"),vet:g("vet"),poor:g("poor"),aid:g("aid"),pens:g("pens"),out:!!f.get("out"),outDate:g("outDate"),outNote:g("outNote"),
    outBasis:g("outBasisSel")==="__other"?normName(g("outBasisOther")):g("outBasisSel")})}
/* Kartda status yazılanda uyğun xana «Bəli» olur */
function rbStFix(m){const s=lc(m.status);for(const[k,re]of MEM_ST)if(re.test(s))m[k]="Bəli";return memDerive(m)}
function bindMem(){const S=UI.mem;const f=$("#mForm");
  bindRows(id=>{S.sel=id;S.draft=null;render()});
  f.oninput=e=>{if(!S.sel)S.draft=readMemForm();if(e.target.name==="out"){$("#m-outBox").hidden=!e.target.checked;if(e.target.checked&&!$("#m-outDate").value)$("#m-outDate").value=todayISO()}
    if(e.target.name==="outBasisSel"){const x=e.target.value==="__other";$("#m-outOther").hidden=!x;if(x)$("#m-outOther").focus()}if(e.target.name==="fin")e.target.value=normFin(e.target.value)};
  f.onsubmit=e=>{e.preventDefault();const o=readMemForm(),msg=$("#mMsg");
    if(!o.name){msg.className="form-msg err";msg.textContent="Soyadı, adı, ata adını yazın.";return}
    if(o.fin&&!/^[0-9A-Z]{7}$/.test(o.fin)){msg.className="form-msg err";msg.textContent="FİN 7 hərf və rəqəmdən ibarət olmalıdır.";return}
    if(!o.out){o.outDate="";o.outNote="";o.outBasis=""}
    else if(!o.outBasis){msg.className="form-msg err";msg.textContent="Üzvlükdən xaric etmə əsasını (Əmək Məcəlləsinin bəndini) seçin.";return}
    if(!guard())return;let pn=0;
    if(S.sel){const i=DB.mem.findIndex(m=>m.id===S.sel),prev=DB.mem[i];DB.mem[i]={...prev,...o};pn=propagate(prev,DB.mem[i])}else{const r={...o,id:newId(),no:""};DB.mem.push(r);S.sel=r.id;S.draft=null}
    memVer++;commit(pn?`Yadda saxlandı · komissiyalarda da yeniləndi (${pn})`:"Yadda saxlandı")};
  $("#mClose")?.addEventListener("click",()=>{S.sel=null;render()});
  $("#mDel")?.addEventListener("click",()=>{const m=DB.mem.find(x=>x.id===S.sel);if(!m||!confirm(`«${m.name}» siyahıdan silinsin?\n(Üzvlükdən xaric etmək üçün «Üzvlükdən xaric edilib» işarəsindən istifadə edin.)`))return;if(!guard())return;DB.mem=DB.mem.filter(x=>x.id!==S.sel);S.sel=null;memVer++;commit("Silindi")});
  $$("[data-body]").forEach(a=>a.onclick=e=>{e.preventDefault();UI.tab="bodies";UI.body.sel=a.dataset.body;UI.body.person=null;render()});
  for(const[id,k]of[["mView","view"],["mOrg","org"],["mSt","st"],["mG","g"]])$("#"+id).onchange=e=>{S[k]=e.target.value;render()};
  searchBox("#mq",v=>S.q=v);
  $("#mImp").onclick=()=>$("#mFile").click();
  $("#mClear")?.addEventListener("click",()=>{const n=DB.mem.length,x=DB.mem.filter(m=>m.out).length;
    if(!confirm(`Üzv siyahısı silinsin?\n${n} nəfər: ${n-x} üzv, ${x} xaric olunan.\n\nKomitə və komissiyalara, protokollara və digər bölmələrə toxunulmur. Silinmədən əvvəl ehtiyat nüsxə saxlanılır.`)||!guard())return;
    if(HAS_API)window.api.snapshot("uzvler-silinmeden-once");DB.mem=[];memVer++;S.sel=null;S.draft=null;commit("Üzv siyahısı silindi. Yeni Excel faylını yükləyin")});
  $("#mFile").onchange=async e=>{const file=e.target.files[0];e.target.value="";if(!file||!guard())return;
    let L;try{L=parseMembers(await readWb(file))}catch(x){toast("Excel faylı oxunmadı");return}
    if(!L.length){toast("Faylda «Soyadı, adı, ata adı» və «FİN» sütunları tapılmadı");return}
    const old=new Map(DB.mem.map(m=>[memKey(m),m]));let same=0;
    const next=L.map(r=>{const ex=old.get(memKey(r));if(ex){same++;return{...ex,...r,id:ex.id,out:r.out||ex.out,outDate:r.out?ex.outDate:(ex.out?ex.outDate:""),outNote:r.out?(ex.outNote||r.outNote):(ex.out?ex.outNote:"")}}return{...r,id:newId()}});
    const inFile=new Set(next.map(m=>m.id));const kept=DB.mem.filter(m=>!inFile.has(m.id)&&m.orgId);next.push(...kept);
    const inN=next.filter(m=>!m.out).length-kept.filter(m=>!m.out).length,outN=next.length-kept.length-inN,gone=DB.mem.length-same-kept.length;
    if(!confirm(`Faylda: ${inN} üzv${outN?` və ${outN} xaric olunan («Ləğv edilənlər»)`:""}.\n${L.rbSkip?`Yüklənməyən vərəq: ${L.rbSkip.join(", ")}.\n`:""}${L.rbCols&&L.rbCols.length?`Reyestr sütunları tapıldı: ${L.rbCols.join(", ")}.\n`:"Reyestr sütunları tapılmadı.\n"}${L.rbBad?`DİQQƏT: ${L.rbBad} xanada «Bəli»/«Xeyr» yazılmayıb, onlar boş qalacaq.\n`:""}${DB.mem.length?`Hazırkı siyahı (${DB.mem.length}) bu faylla əvəz olunacaq; ${same} nəfər tanındı${gone>0?`, ${gone} nəfər siyahıdan çıxacaq`:""}.\nƏl ilə qeyd etdiyiniz xaric məlumatları saxlanılır.\n${kept.length?`Ayrıca yüklənmiş ilk təşkilat üzvləri (${kept.length}) saxlanılır.\n`:""}`:""}Davam edilsin?`))return;
    if(HAS_API)window.api.snapshot("uzvler-oncesi");
    DB.mem=next;memVer++;S.sel=null;S.view="in";commit(`${next.length} nəfər yükləndi`)};
  $("#mExp").onclick=()=>exportSheets([memSheet()],"Üzvlərin siyahısı");
  $("#mHik")?.addEventListener("click",()=>{const L=window.api.hikMembers();if(!L||!L.length){toast("HİK-hesabatda üzv siyahısı tapılmadı");return}if(!guard())return;
    const{next,upd,add}=mergeFromHik(L);
    if(!confirm(`HİK-hesabatın Uçot kartında ${L.length} nəfər var.\nYenilənəcək: ${upd}\nYeni əlavə olunacaq: ${add}\n\nCinsi və burada əl ilə yazdığınız digər məlumatlar saxlanılır. Davam edilsin?`))return;
    window.api.snapshot("hik-oncesi");DB.mem=next;memVer++;commit(`HİK-hesabatdan götürüldü: ${upd} yeniləndi, ${add} əlavə edildi`)})}

/* Üzvün adı / FİN / təvəllüdü dəyişəndə komissiyalarda, ilk təşkilatlarda və tədbirlərdə də yenilənir */
function propagate(a,b){let n=0;const k=nkey(a.name);
  for(const bd of DB.bodies)for(const p of bd.people)if(nkey(p.name)===k){
    if(a.name!==b.name){p.name=b.name;n++}
    for(const[f,v]of[["fin",b.fin],["birth",b.birth]])if(p[f]&&p[f]===a[f]&&v!==a[f]){p[f]=v;n++}}
  if(a.name!==b.name){for(const o of DB.orgs)if(nkey(o.chair)===k){o.chair=b.name;n++}for(const e of DB.events)if(nkey(e.resp)===k){e.resp=b.name;n++}}
  memVer++;return n}
/* HİK-hesabatın «Uçot kartı» siyahısından üzvləri götürmək (bu kompüterdə HİK-hesabat varsa) */
function mergeFromHik(L){const old=new Map(DB.mem.map(m=>[memKey(m),m]));const byName=new Map(DB.mem.map(m=>[nkey(m.name),m]));let upd=0,add=0;const next=DB.mem.slice();
  for(const h of L){const r={name:normName(h.name),fin:normFin(h.fin),birth:h.birth||"",unit:normName(h.unit),pos:normName(h.pos),edu:normName(h.edu),hired:h.hired||"",
      status:normName(h.sst||""),out:h.status==="Üzvlükdən xaric",outDate:h.exitDate||"",outNote:normName(h.exitBasis||"")};if(!r.name)continue;
    const ex=old.get(memKey(r))||(!r.fin&&byName.get(nkey(r.name)));
    if(ex){const i=next.indexOf(ex);next[i]={...ex,...Object.fromEntries(Object.entries(r).filter(([k,v])=>v!==""&&v!=null&&!(k==="out"&&!v&&!ex.out))),gender:ex.gender};upd++}
    else{next.push({...r,id:newId(),no:"",gender:""});add++}}
  return{next,upd,add}}

/* ---------- Komitə və komissiyalar ---------- */
/* ================= «Seçkili Orqanların Reyestri» (Google Forms): hər üzv üçün doldurulmuş forma =================
   HİK-hesabatdakı kimi: forma brauzerdə məlumatla doldurulmuş açılır, «Отправить» düyməsini istifadəçi özü basır. */
const RB_FORM="https://docs.google.com/forms/d/e/1FAIpQLSeEsD5toJ9bFRY36K_mG3UNr1c3wS37pxLZassJYm0UeNrl7Q/viewform";
const RB_E={kom:2123589413,organ:610827886,name:383472460,gender:418783316,birth:1340692376,edu:688539874,work:380120294,pos:1879360134,role:396298816,elected:1729432726,term:45502339,dis:951780494,shehid:1851206019,vet:1068990131,poor:15911477,war:1630428811,aid:2117229602,pens:195221422};
const RB_LABEL={kom:"Rayon komitəsi",organ:"Seçkili orqan",name:"Ad",gender:"Cinsi",birth:"Doğum tarixi",edu:"Təhsil",work:"İş yeri",pos:"Vəzifə",role:"Orqanda vəzifə",elected:"Seçilmə tarixi",term:"Seçilmə müddəti",dis:"Əlillik",shehid:"Şəhid ailəsi",vet:"Veteran",poor:"Aztəminatlı ailə",war:"Müharibə əlilliyi",aid:"Ünvanlı yardım",pens:"Yaşa görə pensiya",fin:"FİN",phone:"Telefon"};
const RB_ORGANS=["İlk Həmkarlar İttifaqı Təşkilatının Həmkarlar İttifaqı Komitəsi","İcraiyyə Komitəsi","Təftiş Komissiyası","Gender məsələləri üzrə Komissiya","Gənclər siyasəti üzrə Komissiya","Əməyin mühafizəsi üzrə Komissiya","Mədəni kütləvi tədbirlər üzrə Komissiya","Hüquq məsələləri üzrə Komissiya","Fərdi əmək mübahisələri üzrə Komissiya","İlk Həmkarlar İttifaqı Təşkilatının Yığıncağı","Rayon/Şəhər Komitəsinin Konfransı","Rayon/Şəhər Komitəsinin Həmkarlar İttifaqı Komitəsi"];
const RB_KOM=["BAKI ŞƏHƏRİ",...["Abşeron rayon komitəsi MİHİ", "Ağcabədi rayon komitəsi MİHİ", "Ağdam rayon komitəsi MİHİ", "Ağdaş rayon komitəsi MİHİ", "Ağstafa rayon komitəsi MİHİ", "Ağsu rayon komitəsi MİHİ", "Astara rayon komitəsi MİHİ", "Balakən rayon komitəsi MİHİ", "Bərdə rayon komitəsi MİHİ", "Beyləqan rayon komitəsi MİHİ", "Biləsuvar rayon komitəsi MİHİ", "Cəbrayıl rayon komitəsi MİHİ", "Cəlilabad rayon komitəsi MİHİ", "Daşkəsən rayon komitəsi MİHİ", "Füzuli rayon komitəsi MİHİ", "Gədəbəy rayon komitəsi MİHİ", "Gəncə şəhər komitəsi MİHİ", "Goranboy rayon komitəsi MİHİ", "Göyçay rayon komitəsi MİHİ", "Göygöl rayon komitəsi MİHİ", "Hacıqabul rayon komitəsi MİHİ", "İmişli rayon komitəsi MİHİ", "İsmayıllı rayon komitəsi MİHİ", "Kəlbəcər rayon komitəsi MİHİ", "Kürdəmir rayon komitəsi MİHİ", "Laçın rayon komitəsi MİHİ", "Lənkəran rayon komitəsi MİHİ", "Lerik rayon komitəsi MİHİ", "Masallı rayon komitəsi MİHİ", "Mingəçevir şəhər komitəsi MİHİ", "Naftalan şəhər komitəsi MİHİ", "Neftçala rayon komitəsi MİHİ", "Oğuz rayon komitəsi MİHİ", "Qax rayon komitəsi MİHİ", "Qazax rayon komitəsi MİHİ", "Qəbələ rayon komitəsi MİHİ", "Qobustan rayon komitəsi MİHİ", "Quba rayon komitəsi MİHİ", "Qubadlı rayon komitəsi MİHİ", "Qusar rayon komitəsi MİHİ", "Saatlı rayon komitəsi MİHİ", "Sabirabad rayon komitəsi MİHİ", "Şabran rayon komitəsi MİHİ", "Salyan rayon komitəsi MİHİ", "Şamaxı rayon komitəsi MİHİ", "Samux rayon komitəsi MİHİ", "Şəki şəhər komitəsi MİHİ", "Şəmkir rayon komitəsi MİHİ", "Şirvan şəhər komitəsi MİHİ", "Siyəzən rayon komitəsi MİHİ", "Sumqayıt şəhər komitəsi MİHİ", "Şuşa rayon komitəsi MİHİ", "Tərtər rayon komitəsi MİHİ", "Tovuz rayon komitəsi MİHİ", "Ucar rayon komitəsi MİHİ", "Xaçmaz rayon komitəsi MİHİ", "Xızı rayon komitəsi MİHİ", "Xocalı rayon komitəsi MİHİ", "Xocavənd rayon komitəsi MİHİ", "Yardımlı rayon komitəsi MİHİ", "Yevlax rayon komitəsi MİHİ", "Zaqatala rayon komitəsi MİHİ", "Zəngilan rayon komitəsi MİHİ", "Zərdab rayon komitəsi MİHİ"]];
/* Formanın quruluşu: Parametrlərdən/pəncərədən yeni link verilibsə DB.rbForm, yoxsa proqramdakı forma */
const RB_YN=["Bəli","Xeyr"];
const RB_DEF=(()=>{const O={kom:RB_KOM,organ:RB_ORGANS,gender:["Kişi","Qadın"],edu:["Orta","Orta ixtisas","Ali","Bakalavr","Magistr","Doktorantura","Digər"],role:["Sədr","Sədr müavini","Digər vəzifə","Seçkili orqan üzvü"],dis:RB_YN,shehid:RB_YN,vet:RB_YN,poor:RB_YN,war:RB_YN,aid:RB_YN,pens:RB_YN};
  const fields={};for(const k of Object.keys(RB_E))fields[k]={id:RB_E[k],opts:O[k]||null};return{url:RB_FORM,fields,other:[]}})();
const rbF=()=>DB.rbForm&&DB.rbForm.fields?DB.rbForm:RB_DEF;
/* Formanın sualları mətninə görə proqramın məlumatları ilə uyğunlaşdırılır. Sıra vacibdir: birinci uyğun gələn qayda götürülür */
const RB_RULES=[["name",/soyad/],["role",/(seçkili )?orqanda.*vəzifə|tutduğu vəzifə/],["organ",/seçkili (orqan|organ)/,1],["pos",/vəzifə/],["work",/müəssisə|iş yeri|işlədiyi/],["kom",/(rayon|şəhər).*komitə/,1],["gender",/cins/],["birth",/doğum|təvəllüd/],["edu",/təhsil/],["elected",/seçilmə tarix|seçildiyi tarix/],["term",/seçilmə müddət|müddət/],["war",/müharibə.*əlil/],["dis",/əlil/],["shehid",/şəhid/],["vet",/veteran/],["poor",/aztəminat/],["aid",/ünvanlı|sosial yardım/],["pens",/pensiya/],["fin",/^fin\b|\bfin\b|fərdi identifik/],["phone",/telefon|əlaqə/]];
function rbParse(data,url){const items=(data&&data[1]&&data[1][1])||[],fields={},other=[];
  for(const it of items){const type=it[3];if(![0,1,2,3,4,9].includes(type)||!it[4]||!it[4][0])continue;const title=normName(it[1]),t=lc(title),q=it[4][0],opts=Array.isArray(q[1])?q[1].map(o=>normName(o[0])).filter(Boolean):null;
    const r=RB_RULES.find(([k,re,needOpts])=>!fields[k]&&re.test(t)&&(!needOpts||(opts&&opts.length)));
    if(r)fields[r[0]]={id:q[0],opts:opts&&opts.length?opts:null,title};else if(!other.includes(title))other.push(title)}
  return{url,title:normName(data&&data[3])||normName(data&&data[1]&&data[1][8]),fields,other,read:todayISO()}}
/* Dəyər formanın variantlarında yoxdursa boş qalır: yanlış cavab yazılmır */
function rbFit(k,v){const f=rbF().fields[k];if(!f||!v)return"";if(!f.opts)return v;const l=lc(v),O=f.opts;let h=O.find(o=>lc(o)===l);if(h)return h;
  if(k==="role"&&/üzv/.test(l))h=O.find(o=>/üzv/.test(lc(o)));else if(k==="role"&&/müavin/.test(l))h=O.find(o=>/müavin/.test(lc(o)));else if(k==="role"&&/digər/.test(l))h=O.find(o=>/digər/.test(lc(o)));else h=O.find(o=>lc(o).startsWith(l)||l.startsWith(lc(o)));
  return h||""}
const rbKomOpts=()=>(rbF().fields.kom&&rbF().fields.kom.opts)||[];
const rbOrganOpts=()=>(rbF().fields.organ&&rbF().fields.organ.opts)||[];
const rbKom=()=>{const L=rbKomOpts();if(L.includes(DB.rbKom))return DB.rbKom;const o=lc(DB.org);return L.find(k=>o.includes(lc(k).split(" ")[0]))||""};
const rbOrgan=b=>{const L=rbOrganOpts();if(L.includes(b.rbOrgan))return b.rbOrgan;const n=lc(b.name);
  const R=[[/rayon komitəsi/,/(rayon|şəhər).*(həmkarlar|komitəsi$)/],[/rəyasət|icra/,/icra|rəyasət/],[/təftiş/,/təftiş/],[/maddi/,/maddi/],[/gender/,/gender/],[/gənc/,/gənc/],[/əməyin/,/əməyin/],[/mədəni/,/mədəni/],[/hüquq/,/hüquq/],[/fərdi|mübahisə/,/fərdi|mübahisə/]].find(([x])=>x.test(n));
  return R?L.find(o=>R[1].test(lc(o)))||"":""};
function rbGender(name,m){const g=lc(m&&m.gender);if(/^q/.test(g))return"Qadın";if(/^k/.test(g))return"Kişi";const s=lc(name);if(/\bqızı\b/.test(s))return"Qadın";if(/\boğlu\b/.test(s))return"Kişi";const f=s.split(" ")[0]||"";return/(ova|eva)$/.test(f)?"Qadın":/(ov|ev)$/.test(f)?"Kişi":""}
function rbEdu(e){const s=eduShort(e);return{"Ali":"Ali","Ali (magistr)":"Magistr","Ali (bakalavr)":"Bakalavr","Orta ixtisas":"Orta ixtisas","Kollec":"Orta ixtisas","Ümumi təhsil":"Orta","Natamam ali":"Digər","İlk peşə":"Digər","Qeyri-ali":"Digər"}[s]||""}
function rbRole(r){const l=lc(r);if(/^sədr$/.test(l))return"Sədr";if(/^sədr(in)? müavin/.test(l))return"Sədr müavini";if(!l||/^üzv/.test(l))return"Seçkili orqan üzvü";return"Digər vəzifə"}
/* Formanın dəyərləri. Proqramda məlumatı olmayan sual boş qalır: yanlış cavab yazmaq əvəzinə formada özünüz seçirsiniz */
function rbValues(b,p){const v=personView(p),m=v.m,st=m?lc(m.status):null,yn=(k,re)=>st===null?"":m[k]||(re.test(st)?"Bəli":"Xeyr"),unk=DB.rbUnk==="Xeyr"?"Xeyr":"",ex=k=>(m&&m[k])||unk;
  const st2=k=>p[k]||"";
  return{kom:rbKom(),organ:rbOrgan(b),name:p.name,gender:rbGender(p.name,p.gender?{gender:p.gender}:m),birth:/^\d{4}-\d{2}-\d{2}$/.test(v.birth)?v.birth:"",edu:rbEdu(v.edu),
    work:p.work||(m?(orgName(m)!=="Digər"?orgName(m):normName(String(m.unit||"").split(/[,\n]/)[0])):""),pos:v.pos,role:rbRole(p.role),elected:p.elected||b.elected||(m&&m.elected)||"",term:p.term||b.rbTerm||(m&&m.term)||"",
    dis:st2("dis")||yn("dis",/əlil/),shehid:st2("shehid")||yn("shehid",/şəhid/),vet:st2("vet")||yn("vet",/veteran|qazi|müharibə iştirak/),poor:st2("poor")||ex("poor"),war:st2("war")||yn("war",/müharibə əlil/),aid:st2("aid")||ex("aid"),pens:st2("pens")||(m&&m.pens)||(isRet(v.birth)?"":unk),fin:v.fin,phone:p.phone||(m&&m.phone)||""}}
function rbUrl(vals){const F=rbF(),u=new URLSearchParams();u.set("usp","pp_url");for(const k of Object.keys(F.fields)){const v=rbFit(k,vals[k]);if(v)u.append("entry."+F.fields[k].id,v)}return F.url+"?"+u.toString()}
async function rbOpen(url){
  try{if(HAS_API&&window.api.openForm){if(await window.api.openForm(url))return true}
    else if(!HAS_API&&window.open(url,"_blank","noopener"))return true}catch(e){}
  try{await navigator.clipboard.writeText(url);toast("Link kopyalandı — brauzerdə yapışdırın");return true}catch(e){toast("Forma açılmadı");return false}}
function rbDialog(b){let pend=null,prev=null,linkMsg="";const md=document.createElement("div");md.className="lic-modal rb-modal";document.body.append(md);
  const opt=(L,v,ph)=>`<option value="">${esc(ph)}</option>`+L.map(x=>`<option${x===v?" selected":""}>${esc(x)}</option>`).join("");
  const close=()=>{md.remove();document.removeEventListener("keydown",onKey);render()};
  const onKey=e=>{if(e.key==="Escape")close()};document.addEventListener("keydown",onKey);
  const save=()=>{if(licOk())persist()};
  const rbMiss=p=>{const vals=rbValues(b,p),miss=Object.keys(rbF().fields).filter(k=>!rbFit(k,vals[k])).map(k=>RB_LABEL[k]);return miss.length?"Formada özünüz doldurun: "+miss.join(", "):"Bütün suallar doldurulacaq"};
  /* Parametr dəyişəndə düymələr yenidən yaradılmır, yoxsa eyni anda basılan düymə işləməz */
  const upd=()=>{save();md.querySelectorAll("[data-rb-miss]").forEach(el=>{const p=b.people.find(y=>y.id===el.dataset.rbMiss);if(p)el.textContent=rbMiss(p)})};
  const draw=()=>{const L=sortPeople(b.people),done=L.filter(p=>p.rbSent).length,next=L.find(p=>!p.rbSent);
    md.innerHTML=`<div class="lic-box rb-box" role="dialog" aria-modal="true" aria-label="Seçkili Orqanların Reyestri">
      <div class="rb-head"><div><h2>Seçkili Orqanların Reyestri</h2><p class="mut">${esc(b.name)} · ${done} / ${L.length} göndərilib</p></div><button class="btn ghost" type="button" id="rbClose">Bağla</button></div>
      <details class="rb-link"${prev||linkMsg?" open":""}><summary>Forma linki${DB.rbForm?` · yenilənib ${dmy(DB.rbForm.read)}`:""}</summary>
        <div class="fld"><label for="rbUrlIn">Yeni formanın linkini yapışdırın</label><input id="rbUrlIn" value="${esc(prev?prev.url:rbF().url)}"></div>
        <div class="form-actions"><button class="btn" type="button" id="rbRead">Formanı oxu</button>${DB.rbForm?`<button class="btn ghost" type="button" id="rbDefault">Əvvəlki linkə qaytar</button>`:""}</div>
        ${linkMsg?`<div class="form-msg err">${esc(linkMsg)}</div>`:""}
        ${prev?`<div class="rb-prev"><b>Uyğunlaşdırılan suallar: ${Object.keys(prev.fields).length}</b><div class="mut rb-miss">${esc(Object.keys(prev.fields).map(k=>RB_LABEL[k]).join(", "))}</div>
          ${prev.other.length?`<b>Proqram doldurmayacaq (formada özünüz doldurun): ${prev.other.length}</b><div class="mut rb-miss">${esc(prev.other.join("; "))}</div>`:""}
          <div class="form-actions"><button class="btn primary" type="button" id="rbApply">Bu formanı istifadə et</button><button class="btn" type="button" id="rbCancel">Ləğv et</button></div></div>`:""}</details>
      <div class="rb-set">
        <div class="fld"><label for="rbKom">Rayon / şəhər komitəsi</label><select id="rbKom">${opt(rbKomOpts(),rbKom(),"Formada seçin")}</select></div>
        <div class="fld"><label for="rbOrg">Seçkili orqan (formada)</label><select id="rbOrg">${opt(rbOrganOpts(),rbOrgan(b),"Formada seçin")}</select></div>
        <div class="fld"><label for="rbEl">Seçilmə tarixi</label><input id="rbEl" type="date" min="1900-01-01" max="2100-12-31" value="${esc(b.elected||"")}"></div>
        <div class="fld"><label for="rbTerm">Seçilmə müddəti</label><input id="rbTerm" value="${esc(b.rbTerm||"")}" placeholder="məs.: 5 il"></div>
        <div class="fld rb-wide"><label for="rbUnk">Proqramda məlumatı olmayan suallar (aztəminatlı, ünvanlı yardım, pensiya)</label><select id="rbUnk"><option value="">Boş qalsın, formada özüm seçim</option><option value="Xeyr"${DB.rbUnk==="Xeyr"?" selected":""}>«Xeyr» yazılsın</option></select></div>
      </div>
      <p class="hint">Forma brauzerdə açılır. Səhifələri «Далее» ilə keçin, məlumatı yoxlayın, sonda «Отправить» düyməsini özünüz basın. Proqram heç nəyi özü göndərmir.</p>
      ${next?`<button class="btn primary" type="button" data-rb-open="${next.id}">Növbəti: ${esc(next.name)}</button>`:L.length?`<div class="st done">Hamısı göndərilib</div>`:`<div class="empty"><b>Bu orqanda hələ heç kim yoxdur</b></div>`}
      <div class="rb-list">${L.map((p,i)=>{return`<div class="rb-row${pend===p.id?" on":""}"><span class="num">${i+1}</span><div class="rb-who"><b>${esc(p.name)}</b> ${p.rbSent?`<span class="st done">göndərilib ${dmy(p.rbSent)}</span>`:""}
          <div class="mut rb-miss" data-rb-miss="${p.id}">${esc(rbMiss(p))}</div>
          ${pend===p.id?`<div class="rb-ask"><b>Forma açıldı. Göndərdinizmi?</b> <button class="btn primary sm" type="button" data-rb-yes="${p.id}">Bəli, göndərdim</button> <button class="btn sm" type="button" data-rb-no>Hələ yox</button></div>`:""}</div>
          <div class="rb-act"><button class="btn sm" type="button" data-rb-open="${p.id}">${p.rbSent||pend===p.id?"Yenidən aç":"Formanı aç"}</button>${p.rbSent?`<button class="btn ghost sm" type="button" data-rb-undo="${p.id}">İşarəni götür</button>`:""}</div></div>`}).join("")}</div></div>`;
    const q=s=>md.querySelector(s),all=s=>[...md.querySelectorAll(s)];
    q("#rbClose").onclick=close;
    q("#rbRead").onclick=async()=>{linkMsg="";prev=null;const raw=q("#rbUrlIn").value.trim();
      if(!HAS_API){linkMsg="Formanı oxumaq yalnız kompüter proqramında işləyir.";draw();return}
      if(!window.api.formRead){linkMsg="Bunun üçün proqramın yeni EXE faylı lazımdır (EXE_yarat.bat).";draw();return}
      q("#rbRead").disabled=true;q("#rbRead").textContent="Oxunur…";const r=await window.api.formRead(raw).catch(()=>({ok:false,reason:"net"}));
      if(!r.ok)linkMsg={link:"Bu, formanın doldurma linki deyil. Google Forms-da «Göndər» bölməsindəki linki və ya formanı açanda ünvan sətrindəki linki yapışdırın.",login:"Forma Google hesabına giriş tələb edir, ona görə proqram onu oxuya bilmir. Forma sahibindən girişi tələb etməməsini xahiş edin.",net:"Forma oxunmadı. İnternet bağlantısını və linki yoxlayın.",read:"Linkdə forma tapılmadı. Linki yoxlayın."}[r.reason]||"Forma oxunmadı.";
      else{let d=null;try{d=JSON.parse(r.data)}catch(e){}const f=d&&rbParse(d,r.url);if(!f||!f.fields.name||Object.keys(f.fields).length<3)linkMsg="Bu formada reyestr sualları tapılmadı (ad-soyad sualı olmalıdır). Köhnə forma dəyişmədi.";else prev=f}
      draw()};
    q("#rbCancel")&&(q("#rbCancel").onclick=()=>{prev=null;draw()});
    const reset=()=>{const same=prev&&prev.url===rbF().url;if(same)return;const n=DB.bodies.reduce((s,x)=>s+x.people.filter(p=>p.rbSent).length,0);
      if(n&&confirm(`Yeni forma üçün «göndərilib» işarələri silinsin? (${n} nəfər)\n\n«OK» — hamı yenidən göndəriləcək. «Cancel» — işarələr qalır.`))for(const x of DB.bodies)for(const p of x.people)delete p.rbSent};
    q("#rbApply")&&(q("#rbApply").onclick=()=>{if(!guard())return;reset();DB.rbForm=prev;prev=null;persist();toast("Yeni forma yadda saxlandı");draw()});
    q("#rbDefault")&&(q("#rbDefault").onclick=()=>{if(!guard()||!confirm("Proqramdakı əvvəlki forma linkinə qaytarılsın?"))return;delete DB.rbForm;prev=null;persist();draw()});
    q("#rbKom").onchange=e=>{DB.rbKom=e.target.value;upd()};
    q("#rbOrg").onchange=e=>{b.rbOrgan=e.target.value;upd()};
    q("#rbEl").onchange=e=>{if(okYear(e.target.value)){b.elected=e.target.value;upd()}};q("#rbEl").onblur=e=>{if(!okYear(e.target.value))e.target.value=b.elected||""};
    q("#rbTerm").oninput=e=>{b.rbTerm=normName(e.target.value)};q("#rbTerm").onchange=upd;
    q("#rbUnk").onchange=e=>{DB.rbUnk=e.target.value;upd()};
    all("[data-rb-open]").forEach(x=>x.onclick=async()=>{const p=b.people.find(y=>y.id===x.dataset.rbOpen);if(p&&await rbOpen(rbUrl(rbValues(b,p)))){pend=p.id;draw()}});
    all("[data-rb-yes]").forEach(x=>x.onclick=()=>{if(!guard())return;const p=b.people.find(y=>y.id===x.dataset.rbYes);if(p){p.rbSent=todayISO();persist()}pend=null;draw()});
    all("[data-rb-no]").forEach(x=>x.onclick=()=>{pend=null;draw()});
    all("[data-rb-undo]").forEach(x=>x.onclick=()=>{if(!guard())return;const p=b.people.find(y=>y.id===x.dataset.rbUndo);if(p){delete p.rbSent;persist()}draw()});
  };draw()}
try{const st=document.createElement("style");st.textContent=`.rb-modal{align-items:start;overflow:auto}.rb-box{max-width:820px;justify-items:stretch;padding:22px}.rb-head{display:flex;justify-content:space-between;gap:12px;align-items:start}.rb-set{display:grid;grid-template-columns:1fr 1fr;gap:10px}.rb-set select,.rb-set input{width:100%}.rb-wide{grid-column:1/-1}.rb-list{display:grid;gap:0;border-top:1px solid var(--line)}.rb-row{display:grid;grid-template-columns:28px 1fr auto;gap:10px;padding:9px 4px;border-bottom:1px solid var(--line);align-items:start}.rb-row.on{background:var(--gold-soft)}.rb-miss{font-size:12px;margin-top:2px}.rb-ask{margin-top:6px}.ask-term{width:110px;font:inherit;padding:7px 10px;border:1px solid var(--line-strong);border-radius:8px;background:var(--surface);color:var(--ink)}.body-meta .pterm{display:none}@media print{.ask-term{display:none!important}.body-meta .pterm{display:inline}}.m-rb{border:1px solid var(--line);border-radius:10px;padding:8px 12px;margin:6px 0}.m-rb summary{cursor:pointer;font-weight:600}.m-rb-g{grid-template-columns:1fr 1fr;margin-top:8px}.rb-link{border:1px solid var(--line);border-radius:10px;padding:10px 12px}.rb-link summary{cursor:pointer;font-weight:600}.rb-link .fld{margin-top:8px}.rb-link input{width:100%}.rb-prev{display:grid;gap:4px;margin-top:10px}.rb-act{display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end}@media(max-width:600px){.rb-set{grid-template-columns:1fr}.rb-row{grid-template-columns:24px 1fr}.rb-act{grid-column:2;justify-content:flex-start}}`;document.head.append(st)}catch(e){}

function bodiesHTML(){
  const S=UI.body;if(!S.sel||!DB.bodies.find(b=>b.id===S.sel))S.sel=DB.bodies[0]&&DB.bodies[0].id;
  const b=DB.bodies.find(x=>x.id===S.sel);
  const chips=DB.bodies.map(x=>`<button class="pchip" data-b="${x.id}" aria-pressed="${x.id===S.sel}">${esc(x.name)} <span class="cnum">${x.people.length}</span></button>`).join("");
  const tools=`<div class="reg-tools"><button class="btn" id="bImp" title="Komissiyalar Excel faylı: hər vərəq bir orqan/komissiya">Komissiyalar Excel-dən yüklə</button><input type="file" id="bFile" accept=".xlsx,.xls" hidden>
    <button class="btn" id="bNew">+ Yeni komissiya</button>${DB.bodies.length?`<button class="btn" id="bFill" title="Boş FİN, təvəllüd, iş yeri, vəzifə və təhsil xanalarını «Üzvlər» siyahısından doldurur">Üzvlər siyahısından doldur</button><button class="btn" id="bExp">Hamısını Excel-ə çıxar</button><button class="btn danger" id="bClear" title="Excel-dən yüklənmiş bütün komitə və komissiyaları silir">Excel-i sil</button>`:""}</div>`;
  if(!b)return`<div class="card"><h2>Komitə və komissiyalar</h2>${tools}<div class="empty"><b>Hələ heç bir komissiya yoxdur</b>«Komissiyalar Excel-dən yüklə» ilə faylınızı yükləyin və ya yeni komissiya yaradın.</div></div>`;
  const pe=S.person&&b.people.find(p=>p.id===S.person);const po=pe||{name:"",work:"",pos:"",phone:"",birth:"",fin:"",edu:"",role:"Üzv"};
  const rows=sortPeople(b.people).map((p,i)=>{const v=personView(p);const d=(own,val)=>own?esc(val):val?`<span class="from" title="Üzv siyahısından">${esc(val)}</span>`:"";
    return`<tr tabindex="0" data-id="${p.id}" class="${p.id===S.person?"sel":""} ${isRet(v.birth)?"ret":""}"><td class="num">${i+1}</td>
    <td>${retName(esc(p.name),v.birth)}${v.m?(v.m.out?' <span class="st rej">xaric</span>':"")+(v.fuzzy&&nkey(v.m.name)!==nkey(p.name)?`<div class="fz">Üzvlər siyahısında: <b>${esc(v.m.name)}</b> <button class="btn sm" type="button" data-fixname="${p.id}" title="Adı üzvlər siyahısındakı kimi yaz">Adı düzəlt</button></div>`:""):' <span class="st pend" title="Bu ad üzv siyahısında tapılmadı">siyahıda yoxdur</span>'}${v.m&&statusParts(v.m.status).length?`<div class="mut">${esc(statusParts(v.m.status).join(", "))}</div>`:""}</td>
    <td>${d(p.work,v.work&&v.m&&!p.work?orgName(v.m):v.work)}<div class="mut">${d(p.pos,v.pos)}</div></td><td class="nowrap">${esc(p.phone)}</td>
    <td class="num${isRet(v.birth)?" retd":""}">${d(p.birth,dmy(v.birth))}</td><td>${d(p.fin,v.fin)}</td><td>${d(p.edu,eduShort(v.edu))}</td><td>${p.role?`<b>${esc(p.role)}</b>`:""}</td></tr>`}).join("");
  return`<div class="chips">${chips}</div>
  <div class="split">
  <section class="card"><h2>${pe?"Şəxsin məlumatı":"Şəxs əlavə et"}</h2><p class="hint">Adı yazanda üzv siyahısından seçə bilərsiniz — FİN, təvəllüd, iş yeri avtomatik dolur.</p>
    <form class="entry" id="pForm" novalidate>
      <div class="fld"><label for="p-name">Soyadı, adı, ata adı *</label><input id="p-name" name="name" list="memList" value="${esc(po.name)}" autocomplete="off"><datalist id="memList">${active().map(m=>`<option value="${esc(m.name)}">${esc(m.pos||"")} — ${esc(orgName(m))}</option>`).join("")}</datalist></div>
      <div class="fld"><label for="p-role">Vəzifəsi orqanda</label><input id="p-role" name="role" list="roleList" value="${esc(po.role)}"><datalist id="roleList"><option>Sədr</option><option>Sədr müavini</option><option>Katib</option><option>Üzv</option></datalist></div>
      <div class="fld"><label for="p-work">Çalışdığı müəssisə</label><input id="p-work" name="work" value="${esc(po.work)}"></div>
      <div class="row2"><div class="fld"><label for="p-pos">Vəzifəsi</label><input id="p-pos" name="pos" value="${esc(po.pos)}"></div><div class="fld"><label for="p-phone">Əlaqə nömrəsi</label><input id="p-phone" name="phone" value="${esc(po.phone)}"></div></div>
      <div class="row2"><div class="fld"><label for="p-birth">Təvəllüd</label><input id="p-birth" name="birth" type="date" value="${esc(po.birth)}"></div><div class="fld"><label for="p-fin">FİN</label><input id="p-fin" name="fin" maxlength="7" class="fin" value="${esc(po.fin)}"></div></div>
      <div class="fld"><label for="p-edu">Təhsil</label><input id="p-edu" name="edu" value="${esc(po.edu)}"></div>
      <div class="form-msg" id="pMsg"></div>
      <div class="form-actions"><button class="btn primary" type="submit">${pe?"Yadda saxla":"Əlavə et"}</button>${pe?`<button class="btn" type="button" id="pClose">Bağla</button><button class="btn ghost" type="button" id="pDel">Çıxar</button>`:""}</div></form></section>
  <section class="card printable" data-ptitle="${esc(b.name)}"><h2>${esc(b.name)}</h2>
    <div class="body-meta"><span class="mut">Seçilib:</span> <input type="date" id="bElected" min="1900-01-01" max="2100-12-31" value="${esc(b.elected)}"><span class="pdate">${dmy(b.elected)||"—"}</span> <span class="mut">Müddəti:</span> <input id="bTerm" class="ask-term" value="${esc(b.rbTerm||"")}" placeholder="məs.: 5 il" aria-label="Seçilmə müddəti"><span class="pdate pterm">${esc(b.rbTerm||"—")}</span> ${DB.bodies.length>1?`<button class="btn ghost" id="bAllDates" type="button" title="Bu tarixi və müddəti tarixi və ya müddəti boş olan digər orqanlara yazır">Boş orqanlara da yaz</button>`:""} <span class="mut">${b.people.length} nəfər${b.people.find(p=>/^sədr$/i.test(p.role))?" · Sədr: <b>"+esc(b.people.find(p=>/^sədr$/i.test(p.role)).name)+"</b>":""}</span>
      <span class="push"></span><button class="btn primary" id="bRb" title="Üzvlərin məlumatını «Seçkili Orqanların Reyestri» formasına doldurur">🔗 Reyestr formasına doldur</button><button class="btn ghost" id="bRen">Adını dəyiş</button><button class="btn ghost" id="bDel">Sil</button></div>
    ${tools}
    <div class="tbl-wrap"><table class="reg"><thead><tr><th class="num">№</th><th>Soyadı, adı, ata adı</th><th>Müəssisə / vəzifə</th><th>Əlaqə</th><th class="num">Təvəllüd</th><th>FİN</th><th>Təhsil</th><th>Orqanda</th></tr></thead><tbody>
    ${rows||`<tr><td colspan="8"><div class="empty"><b>Bu orqanda hələ heç kim yoxdur</b>Soldakı formadan əlavə edin.</div></td></tr>`}</tbody></table></div>
    <p class="mut legend"><span class="from">Boz mail yazı</span> — Excel-də boş olan xana üzv siyahısından götürülüb.</p></section></div>`}
function readPerson(){const f=new FormData($("#pForm"));const g=k=>normName(f.get(k));return{name:g("name"),role:g("role"),work:g("work"),pos:g("pos"),phone:g("phone"),birth:g("birth"),fin:normFin(g("fin")),edu:g("edu")}}
function bindBodies(){const S=UI.body;
  $$("[data-b]").forEach(c=>c.onclick=()=>{S.sel=c.dataset.b;S.person=null;render()});
  $("#bImp").onclick=()=>$("#bFile").click();
  $("#bFile").onchange=async e=>{const file=e.target.files[0];e.target.value="";if(!file||!guard())return;
    let L;try{L=parseBodies(await readWb(file))}catch(x){toast("Excel faylı oxunmadı");return}
    if(!L.length){toast("Faylda komissiya üzvləri tapılmadı");return}
    if(!confirm(`Faylda ${L.length} orqan/komissiya, cəmi ${L.reduce((a,b)=>a+b.people.length,0)} qeyd tapıldı:\n\n${L.map(b=>"• "+b.name+" — "+b.people.length).join("\n")}\n\nSeçilmə tarixi tapıldı: ${L.filter(b=>b.elected).length} orqan. Seçilmə müddəti tapıldı: ${L.filter(b=>b.rbTerm).length} orqan.\n\n${DB.bodies.length?"Hazırkı komitə və komissiyalar bu faylla əvəz olunacaq. ":""}Davam edilsin?`))return;
    if(HAS_API)window.api.snapshot("komissiyalar-oncesi");
    const f=fillFromMembers(L);DB.bodies=L;S.sel=L[0].id;S.person=null;commit(`Komissiyalar yükləndi${f.n?` · ${f.n} boş xana üzv siyahısından dolduruldu`:""}`)};
  $("#bFill")?.addEventListener("click",()=>{if(!guard())return;const f=fillFromMembers(DB.bodies);if(!f.n){toast("Doldurulacaq boş xana tapılmadı");return}commit(`${f.p} nəfər üçün ${f.n} xana «Üzvlər» siyahısından dolduruldu`)});
  $("#bClear")?.addEventListener("click",()=>{const n=DB.bodies.reduce((a,b)=>a+b.people.length,0);
    if(!confirm(`Bütün komitə və komissiyalar silinsin?\n${DB.bodies.length} orqan, ${n} qeyd.\n\nÜzvlər siyahısına toxunulmur. Sonra Excel faylını yenidən yükləyə bilərsiniz.`)||!guard())return;
    if(HAS_API)window.api.snapshot("komissiyalar-silinmeden-once");DB.bodies=[];S.sel=null;S.person=null;commit("Komitə və komissiyalar silindi")});
  $("#bNew").onclick=()=>askText("Yeni komissiyanın adı","",n=>{if(!guard())return;const b={id:newId(),name:normName(n),elected:"",note:"",people:[]};DB.bodies.push(b);S.sel=b.id;S.person=null;commit("Komissiya yaradıldı")});
  $("#bExp")?.addEventListener("click",()=>exportSheets(DB.bodies.map(bodySheet),"Komitə və komissiyalar"));
  const b=DB.bodies.find(x=>x.id===S.sel);if(!b)return;
  $("#bRb").onclick=()=>rbDialog(b);
  $("#bTerm").onchange=e=>{if(!guard()){e.target.value=b.rbTerm||"";return}b.rbTerm=normName(e.target.value);const t=$(".body-meta .pterm");if(t)t.textContent=b.rbTerm||"—";persist()};
  $("#bAllDates")?.addEventListener("click",()=>{const el=$("#bElected").value,tm=normName($("#bTerm").value);if(!okYear(el)||(!el&&!tm)){toast("Əvvəlcə tarixi və ya müddəti yazın");return}if(!guard())return;
    b.elected=el;b.rbTerm=tm;let n=0;for(const x of DB.bodies)if(x!==b){let ch=false;if(el&&!x.elected){x.elected=el;ch=true}if(tm&&!x.rbTerm){x.rbTerm=tm;ch=true}if(ch)n++}
    commit(n?`${n} orqanın boş xanaları dolduruldu`:"Boş xana yoxdur")});
  /* İl yazılarkən (0002, 0020 …) tarix yadda saxlanmır və səhifə yenidən qurulmur, yoxsa yazı yarımçıq kəsilir */
  $("#bElected").onchange=e=>{const v=e.target.value;if(!okYear(v))return;if(!guard()){e.target.value=b.elected;return}b.elected=v;const pd=$(".body-meta .pdate");if(pd)pd.textContent=dmy(v)||"—";persist()};
  $("#bElected").onblur=e=>{const v=e.target.value;if(okYear(v)){if(v!==(b.elected||"")&&guard()){b.elected=v;persist();toast("Seçki tarixi yadda saxlandı")}}else{e.target.value=b.elected||"";toast("Tarixin ilini tam yazın, məsələn 2024")}};
  $("#bRen").onclick=()=>askText("Komissiyanın adı",b.name,n=>{if(!guard())return;b.name=normName(n);commit("Ad dəyişdirildi")});
  $("#bDel").onclick=()=>{if(!confirm(`«${b.name}» (${b.people.length} nəfər) silinsin?`)||!guard())return;DB.bodies=DB.bodies.filter(x=>x!==b);S.sel=null;commit("Silindi")};
  bindRows(id=>{S.person=id;render()});
  $$("[data-fixname]").forEach(btn=>btn.onclick=e=>{e.stopPropagation();if(!guard())return;const p=b.people.find(x=>x.id===btn.dataset.fixname);const v=p&&personView(p);if(!v||!v.m)return;
    const old=p.name;p.name=v.m.name;fillFromMembers([b]);for(const bd of DB.bodies)for(const q of bd.people)if(q!==p&&nkey(q.name)===nkey(old))q.name=v.m.name;commit(`Ad düzəldildi: ${v.m.name}`)});
  const f=$("#pForm");
  $("#p-name").addEventListener("change",e=>{const m=memByKey(e.target.value);if(!m)return;const set=(id,v)=>{const el=$(id);if(el&&!el.value&&v)el.value=v};
    e.target.value=m.name;set("#p-work",orgName(m));set("#p-pos",m.pos);set("#p-birth",m.birth);set("#p-fin",m.fin);set("#p-edu",eduShort(m.edu))});
  f.onsubmit=e=>{e.preventDefault();const p=readPerson();if(!p.name){$("#pMsg").className="form-msg err";$("#pMsg").textContent="Adı yazın.";return}if(!guard())return;
    if(S.person){const i=b.people.findIndex(x=>x.id===S.person);b.people[i]={...b.people[i],...p}}else b.people.push({...p,id:newId()});
    S.person=null;commit("Yadda saxlandı")};
  $("#pClose")?.addEventListener("click",()=>{S.person=null;render()});
  $("#pDel")?.addEventListener("click",()=>{const p=b.people.find(x=>x.id===S.person);if(!p||!confirm(`«${p.name}» bu orqandan çıxarılsın?`)||!guard())return;b.people=b.people.filter(x=>x!==p);S.person=null;commit("Çıxarıldı")})}

/* ---------- Veteran, Qazi, Şəhid ailəsi ---------- */
/* Bir neçə statusu olan şəxs hər statusun bölməsində göstərilir, digər statusları alt yazıda */
function vetGroups(){const g=new Map();for(const m of active())for(const p of statusParts(m.status)){const k=stLow(p);if(!g.has(k))g.set(k,[p,[]]);g.get(k)[1].push(m)}
  const ord=s=>{const i=STATUSES.findIndex(x=>stLow(x)===stLow(s));return i<0?99:i};return[...g.values()].sort((a,b)=>ord(a[0])-ord(b[0])||a[0].localeCompare(b[0],"az"))}
function vetHTML(){const G=vetGroups(),T=new Set(G.flatMap(x=>x[1].map(m=>m.id))).size,N=G.reduce((a,x)=>a+x[1].length,0);
  return`<section class="card printable" data-ptitle="Veteran, Qazi, Şəhid ailəsi statuslu üzvlər"><h2>Veteran, Qazi, Şəhid ailəsi</h2>
    <p class="hint">Üzvlər siyahısındakı «Statusu» sütunundan avtomatik yığılır. Statusu dəyişmək üçün «Üzvlər» bölməsində şəxsin kartını açın.</p>
    <div class="reg-tools"><button class="btn" id="vExp">Excel-ə çıxar</button></div>
    <div class="sum">Cəmi <b>${T}</b> nəfər: ${G.map(([s,l])=>`${esc(s)} — <b>${l.length}</b>`).join(" · ")||"—"}${N>T?`<div class="mut">Bir neçə statusu olan şəxs hər bölmədə göstərilir, ona görə bölmələrin cəmi (${N}) nəfər sayından çoxdur.</div>`:""}</div>
    ${G.map(([s,l])=>`<h3 class="grp-h">${esc(s)} <span class="badge">${l.length}</span></h3><div class="tbl-wrap"><table class="reg"><thead><tr><th class="num">№</th><th>Soyadı, adı, ata adı</th><th>FİN</th><th class="num">Təvəllüd</th><th class="num">Yaş</th><th>İlk təşkilat / vəzifə</th></tr></thead><tbody>
      ${l.sort((a,b)=>a.name.localeCompare(b.name,"az")).map((m,i)=>`<tr tabindex="0" data-id="${m.id}" class="${isRet(m.birth)?"ret":""}"><td class="num">${i+1}</td><td>${retName(esc(m.name),m.birth)}${statusOthers(m.status,s).length?`<div class="mut">${esc(statusOthers(m.status,s).join(", "))}</div>`:""}</td><td>${esc(m.fin)}</td><td class="num${isRet(m.birth)?" retd":""}">${dmy(m.birth)}</td><td class="num${isRet(m.birth)?" retd":""}">${ageOf(m.birth)??""}</td><td>${esc(orgName(m))}<div class="mut">${esc(m.pos)}</div></td></tr>`).join("")}</tbody></table></div>`).join("")
    ||`<div class="empty"><b>Statuslu üzv yoxdur</b>Üzvlər Excel faylındakı «Statusu» sütununda Veteran, Qazi, Şəhid ailəsi yazılanlar burada görünür.</div>`}</section>`}
function bindVet(){bindRows(id=>{UI.tab="mem";UI.mem.sel=id;render()});$("#vExp").onclick=()=>exportSheets([vetSheet()],"Veteran, Qazi, Şəhid ailəsi")}

/* ---------- Ümumi reyestr: İlk təşkilatlar, Tədbirlər planı, Protokollar, Plenumlar (fayllarla) ---------- */
/* Müqavilənin vəziyyəti bitmə tarixinə görə avtomatik hesablanır */
const CT_TYPES=["Kollektiv müqavilə","Saziş","Əməkdaşlıq müqaviləsi","Xidmət müqaviləsi","Digər"];
function ctState(o){if(!o.dateTo)return{t:"Müddətsiz",c:"info",d:null};const d=yearsDays(todayISO(),o.dateTo);
  if(d<0)return{t:"Müddəti bitib",c:"rej",d};if(d<=60)return{t:`${d} gün qalıb`,c:"pend",d};return{t:"Qüvvədədir",c:"done",d}}
function yearsDays(a,b){return Math.round((Date.parse(b)-Date.parse(a))/864e5)}
const EV_ST=["Planlaşdırılıb","İcra olunub","Təxirə salınıb","Ləğv edilib"];
const REG={
  orgs:{title:"İlk təşkilatlar",one:"İlk təşkilat",sheet:"İlk Təşkilatlar",
    fields:[["name","Təşkilatın adı *","text",1],["chair","İlk təşkilatın sədri","person"],["phone","Əlaqə nömrəsi"],["state","Vəziyyəti","select",0,"«Qoşulacaq» və «Rayon komitəsi» ilk təşkilatların sayına daxil edilmir; rayon komitəsi siyahıda nömrəsiz göstərilir.",ORG_ST],["founded","Yaradılma tarixi","date"],
      ["keys","Açar sözlər (üzvləri saymaq üçün)","text",0,"Struktur bölmədə bu sözlərdən biri olan üzvlər bu təşkilata aid sayılır. Vergüllə ayırın."],["note","Qeyd","area"]],
    cols:[["Təşkilat",o=>`${isRK(o)?"<b>"+esc(o.name)+"</b>":esc(o.name)}<div class="mut">${o.chair?"Sədr: "+esc(o.chair):""}${o.phone?" · "+esc(o.phone):""}</div>`],
      ["Üzv",o=>orgStat(o).n,"num"],["Qadın",o=>orgStat(o).q,"num"],["Kişi",o=>orgStat(o).k,"num"],["Fayl",o=>fileBadge(o),"num"]],
    xl:[["Təşkilatın adı",o=>o.name],["Vəziyyəti",o=>o.state||"Fəaliyyətdədir"],["Sədr",o=>o.chair],["Əlaqə nömrəsi",o=>o.phone],["Yaradılma tarixi",o=>dmy(o.founded)],["Üzvlərin sayı",o=>orgStat(o).n],["Qadın",o=>orgStat(o).q],["Kişi",o=>orgStat(o).k],["Struktur bölmələr",o=>orgStat(o).u],["Qeyd",o=>o.note],["Fayllar",o=>(o.files||[]).map(f=>f.name).join("; ")]],
    sort:(a,b)=>(isRK(b)-isRK(a))||orgStat(b).n-orgStat(a).n,noNum:isRK,
    /* cəm sətirləri: ilk təşkilatlar üzrə və MİHİ RK ilə birlikdə */
    totals:L=>{const sum=(A,f)=>A.reduce((a,o)=>a+f(o),0);const T=A=>[A.length,sum(A,o=>orgStat(o).n),sum(A,o=>orgStat(o).q),sum(A,o=>orgStat(o).k),sum(A,o=>(o.files||[]).length)];
      const io=L.filter(o=>!isRK(o)),act=io.filter(orgActive),rows=[[`İlk təşkilatlar üzrə cəmi (${io.length})`,...T(io).slice(1)]];
      if(L.some(isRK))rows.push(["Cəmi (MİHİ RK ilə birlikdə)",...T(L).slice(1)]);return rows},
    text:o=>[o.name,o.chair,o.note].join(" ")},
  events:{title:"Tədbirlər planı",one:"Tədbir",sheet:"Tədbirlər Planı",year:"date",
    fields:[["date","Tarix *","date",1],["dateTo","Bitmə tarixi (çoxgünlükdürsə)","date"],["title","Tədbirin adı *","area",1],["place","Keçiriləcəyi yer"],["resp","Məsul şəxs","person"],["part","İştirakçılar (sayı və ya kimlər)"],["status","İcra vəziyyəti","select",0,"",EV_ST],["note","Qeyd / icra haqqında","area"]],
    cols:[["Tarix",o=>`${dmy(o.date)}${o.dateTo?"<div class=mut>– "+dmy(o.dateTo)+"</div>":""}`,"num nowrap"],["Tədbir",o=>`${esc(o.title)}<div class="mut">${esc([o.place,o.resp].filter(Boolean).join(" · "))}</div>`],["Vəziyyət",o=>stBadge(o.status)],["Fayl",o=>fileBadge(o),"num"]],
    xl:[["Tarix",o=>dmy(o.date)],["Bitmə tarixi",o=>dmy(o.dateTo)],["Tədbirin adı",o=>o.title],["Yer",o=>o.place],["Məsul şəxs",o=>o.resp],["İştirakçılar",o=>o.part],["İcra vəziyyəti",o=>o.status||"Planlaşdırılıb"],["Qeyd",o=>o.note],["Fayllar",o=>(o.files||[]).map(f=>f.name).join("; ")]],
    sort:(a,b)=>(a.date||"").localeCompare(b.date||""),text:o=>[o.title,o.place,o.resp,o.note,o.status].join(" ")},
  prots:{title:"Protokollar",one:"Protokol",sheet:"Protokollar",year:"date",
    fields:[["no","Protokol №","text"],["date","Tarix *","date",1],["body","Hansı orqanın iclası","body"],["present","İştirak edənlər (sayı)","number"],["agenda","Gündəlik (müzakirə olunan məsələlər)","area",1],["decisions","Qəbul edilən qərarlar","area"],["note","Qeyd","area"]],
    cols:[["Protokol №",o=>esc(o.no),"num nowrap"],["Tarix",o=>dmy(o.date),"num nowrap"],["Orqan / gündəlik",o=>`${esc(o.body||"")}<div class="mut clamp">${esc(o.agenda)}</div>`],["Fayl",o=>fileBadge(o),"num"]],
    xl:[["Protokol №",o=>o.no],["Tarix",o=>dmy(o.date)],["Orqan",o=>o.body],["İştirak edənlər",o=>o.present],["Gündəlik",o=>o.agenda],["Qərarlar",o=>o.decisions],["Qeyd",o=>o.note],["Fayllar",o=>(o.files||[]).map(f=>f.name).join("; ")]],
    sort:(a,b)=>(b.date||"").localeCompare(a.date||"")||(parseInt(b.no)||0)-(parseInt(a.no)||0),text:o=>[o.no,o.body,o.agenda,o.decisions,o.note].join(" ")},
  contracts:{title:"Müqavilələr",one:"Müqavilə",sheet:"Müqavilələr",year:"date",
    fields:[["type","Növü","select",0,"",CT_TYPES],["no","Müqavilə №","text"],["date","Bağlanma tarixi *","date",1],["org","İlk təşkilat","org"],["party","Digər tərəf (işəgötürən, təşkilat)","text"],
      ["subject","Mövzusu / qısa məzmunu","area"],["dateFrom","Qüvvəyə minmə tarixi","date"],["dateTo","Bitmə tarixi","date",0,"Bitmə tarixinə 60 gün qalanda «Ümumi» bölməsində xəbərdarlıq görünür."],
      ["regNo","Qeydiyyat (№, tarix, orqan)","text"],["note","Qeyd / icrası","area"]],
    cols:[["Tarix",o=>dmy(o.date),"num nowrap"],["Müqavilə",o=>`${esc(o.type||"")}${o.no?" № "+esc(o.no):""}<div class="mut">${esc([o.org,o.party].filter(Boolean).join(" — "))}</div>`],
      ["Müddət",o=>{const st=ctState(o);return`${o.dateTo?dmy(o.dateTo):"—"}<div><span class="st ${st.c}">${esc(st.t)}</span></div>`},"nowrap"],["Fayl",o=>fileBadge(o),"num"]],
    xl:[["Növü",o=>o.type],["Müqavilə №",o=>o.no],["Bağlanma tarixi",o=>dmy(o.date)],["İlk təşkilat",o=>o.org],["Digər tərəf",o=>o.party],["Mövzusu",o=>o.subject],["Qüvvəyə minmə",o=>dmy(o.dateFrom)],["Bitmə tarixi",o=>dmy(o.dateTo)],["Vəziyyəti",o=>ctState(o).t],["Qeydiyyat",o=>o.regNo],["Qeyd",o=>o.note],["Fayllar",o=>(o.files||[]).map(f=>f.name).join("; ")]],
    sort:(a,b)=>(b.date||"").localeCompare(a.date||""),text:o=>[o.type,o.no,o.org,o.party,o.subject,o.note].join(" ")},
  plens:{title:"Plenumlar",one:"Plenum",sheet:"Plenumlar",year:"date",
    fields:[["no","Plenumun №","text"],["date","Tarix *","date",1],["total","Komitə üzvlərinin sayı","number"],["present","İştirak edənlər","number"],["agenda","Gündəlik","area",1],["decisions","Qərarlar","area"],["guests","Dəvət olunanlar","area"],["note","Qeyd","area"]],
    cols:[["Plenum №",o=>esc(o.no),"num nowrap"],["Tarix",o=>dmy(o.date),"num nowrap"],["Gündəlik",o=>`<div class="clamp">${esc(o.agenda)}</div>`],["İştirak",o=>o.total?`${num(o.present)}/${num(o.total)}${quorum(o)}`:esc(o.present||""),"num nowrap"],["Fayl",o=>fileBadge(o),"num"]],
    xl:[["Plenum №",o=>o.no],["Tarix",o=>dmy(o.date)],["Komitə üzvləri",o=>o.total],["İştirak edənlər",o=>o.present],["Gündəlik",o=>o.agenda],["Qərarlar",o=>o.decisions],["Dəvət olunanlar",o=>o.guests],["Qeyd",o=>o.note],["Fayllar",o=>(o.files||[]).map(f=>f.name).join("; ")]],
    sort:(a,b)=>(b.date||"").localeCompare(a.date||""),text:o=>[o.no,o.agenda,o.decisions,o.note].join(" ")}};
const quorum=o=>{const t=num(o.total),p=num(o.present);if(!t)return"";return p*3>=t*2?' <span class="st done" title="Üzvlərin ən azı 2/3-i">kvorum</span>':' <span class="st rej" title="2/3-dən az">kvorum yoxdur</span>'};
const stBadge=s=>{s=s||"Planlaşdırılıb";return`<span class="st ${s==="İcra olunub"?"done":s==="Ləğv edilib"?"rej":s==="Təxirə salınıb"?"info":"pend"}">${esc(s)}</span>`};
const fileBadge=o=>(o.files||[]).length?`<span class="fcount" title="${esc(o.files.map(f=>f.name).join("\n"))}">📎 ${o.files.length}</span>`:"";
let _os=null,_osv=-1;
function orgStat(o){if(_osv!==memVer+DB.orgs.length*1000+DB.orgs.map(x=>x.keys).join().length){_os=new Map();_osv=memVer+DB.orgs.length*1000+DB.orgs.map(x=>x.keys).join().length}
  if(!_os.has(o.id)||_os.get(o.id).keys!==o.keys){const L=active().filter(m=>memOrg(m)===o);_os.set(o.id,{keys:o.keys,n:L.length,q:L.filter(m=>m.gender==="qadın").length,k:L.filter(m=>m.gender==="kişi").length,u:new Set(L.map(m=>m.unit)).size})}
  return _os.get(o.id)}
const DOCK={prots:1,plens:1};
/* əl ilə sıralama (sürüşdürməklə): qeydlərdə «ord» saxlanılır; sırası verilməmiş (yeni) qeydlər yuxarıda */
const DRAGK={prots:1,plens:1};
const regSort=kind=>{const C=REG[kind];if(!DRAGK[kind]||!DB[kind].some(x=>x.ord!=null))return C.sort;
  return(a,b)=>{const x=a.ord==null,y=b.ord==null;if(x&&y)return C.sort(a,b);if(x)return -1;if(y)return 1;return a.ord-b.ord}};
function regHTML(kind){const C=REG[kind],S=UI[kind],list=DB[kind];
  const ed=S.sel&&list.find(x=>x.id===S.sel);const o=ed||S.draft||{};
  let L=list.slice();if(C.year&&S.y)L=L.filter(x=>(x[C.year]||"").startsWith(S.y));
  const q=lc(S.q);if(q)L=L.filter(x=>lc(C.text(x)).includes(q));L.sort(regSort(kind));const MAN=!!DRAGK[kind];
  const years=C.year?[...new Set(list.map(x=>(x[C.year]||"").slice(0,4)).filter(Boolean))].sort().reverse():[];
  const fld=([k,l,t="text",req,help,opts])=>{const v=o[k]??(k==="status"?"Planlaşdırılıb":k==="state"?"Fəaliyyətdədir":k==="type"?"Kollektiv müqavilə":"");const id=`r-${k}`;
    let inp;if(t==="area")inp=`<textarea id="${id}" name="${k}" rows="${k==="title"?2:3}">${esc(v)}</textarea>`;
    else if(t==="select")inp=`<select id="${id}" name="${k}">${opts.map(x=>`<option${x===v?" selected":""}>${esc(x)}</option>`).join("")}</select>`;
    else if(t==="body")inp=`<input id="${id}" name="${k}" list="bodyList" value="${esc(v)}"><datalist id="bodyList">${["Rayon komitəsinin plenumu",...DB.bodies.map(b=>b.name)].map(x=>`<option>${esc(x)}</option>`).join("")}</datalist>`;
    else if(t==="org")inp=`<select id="${id}" name="${k}"><option value=""></option>${[...DB.orgs.map(x=>x.name),...(v&&!DB.orgs.some(x=>x.name===v)?[v]:[])].map(x=>`<option${x===v?" selected":""}>${esc(x)}</option>`).join("")}</select>`;
    else if(t==="person")inp=`<input id="${id}" name="${k}" list="memList" value="${esc(v)}" autocomplete="off">`;
    else inp=`<input id="${id}" name="${k}" type="${t==="date"?"date":t==="number"?"number":"text"}"${t==="number"?' min="0"':""} value="${esc(v)}">`;
    const au=ed&&(S.auto||[]).includes(k);
    return`<div class="fld${au?" auto":""}"><label for="${id}">${l}${au?' <span class="autob">fayldan</span>':""}</label>${inp}${help?`<span class="help">${help}</span>`:""}</div>`};
  let nn=0;const body=L.map(x=>`<tr tabindex="0" data-id="${x.id}" class="${x.id===S.sel?"sel":""}">${MAN?`<td class="drag-h" title="Sıranı dəyişmək üçün tutub yuxarı-aşağı sürüşdürün" aria-label="Sürüşdür">⠿</td>`:""}<td class="num">${C.noNum&&C.noNum(x)?"":++nn}</td>${C.cols.map(([,f,cls])=>`<td class="${cls||""}">${f(x)}</td>`).join("")}</tr>`).join("");
  const extra=kind==="orgs"&&ed?(()=>{const st=orgStat(ed);const U=new Map();active().filter(m=>memOrg(m)===ed).forEach(m=>U.set(m.unit,(U.get(m.unit)||0)+1));
    return`<details class="units"><summary>Struktur bölmələr: ${U.size} · üzvlər: ${st.n} (qadın ${st.q}, kişi ${st.k})</summary><ul>${[...U].sort((a,b)=>b[1]-a[1]).map(([u,n])=>`<li>${esc(u)} — <b>${n}</b></li>`).join("")}</ul></details>`})():"";
  const unmatched=kind==="orgs"?active().filter(m=>!memOrg(m)):[];
  return`<datalist id="memList">${active().map(m=>`<option value="${esc(m.name)}">`).join("")}</datalist><div class="split">
  <section class="card${ed&&!DOCK[kind]?" printable":""}${ed&&DOCK[kind]?" has-doc":""}" data-ptitle="${esc(C.one)}">${ed&&DOCK[kind]?`<button class="doc-btn" type="button" data-docview="${kind}:${ed.id}" title="Yalnız nömrə, tarix, gündəlik və qərarlar — ayrıca səhifədə">📄 Səhifədə aç</button>`:""}<h2>${ed?C.one:"Yeni: "+C.one.toLocaleLowerCase("az")}</h2>
    <div class="fld top-files"><label>Fayllar${ed&&(ed.files||[]).length?` (${ed.files.length})`:""}</label>${filesHTML(kind,ed)}</div>
    <form class="entry" id="rForm" novalidate>${C.fields.map(fld).join("")}${extra}
      ${kind==="orgs"?(ed?`<div class="fld noweb"><label>Bu təşkilatın üzvləri</label><div class="form-actions"><button class="btn" type="button" id="oImp">Üzvlərini Excel-dən yüklə</button><input type="file" id="oFile" accept=".xlsx,.xls" hidden>${orgStat(ed).n?`<button class="btn ghost" type="button" id="oShow">Siyahıya bax (${orgStat(ed).n})</button>`:""}</div>
        <span class="help">Yeni qoşulan təşkilatın üzv siyahısı (Excel) bu təşkilata əlavə olunur; digər təşkilatların üzvlərinə toxunulmur.</span></div>`:`<div class="fld noweb"><label>Üzvlərin siyahısı</label><div class="form-actions"><button class="btn primary" type="button" id="oImp">Üzv siyahısını Excel-dən yüklə</button><input type="file" id="oFile" accept=".xlsx,.xls" hidden></div>
        <span class="help">Təşkilat avtomatik yaradılır (adı yuxarıda yazılmayıbsa, fayldan götürülür), üzvləri əlavə olunur və ilk təşkilatların sayına qoşulur.</span></div>`):""}
      <div class="form-msg" id="rMsg"></div>
      <div class="form-actions"><button class="btn primary" type="submit">${ed?"Yadda saxla":"Əlavə et"}</button>${ed?`<button class="btn" type="button" id="rClose">Bağla</button><button class="btn ghost" type="button" id="rDel">Sil</button>`:""}</div></form></section>
  <section class="card printable" data-ptitle="${esc(C.title)}"><h2>${esc(C.title)}</h2>
    <div class="reg-tools">${kind==="contracts"?`<button class="btn primary noweb" id="ctPdf" title="PDF, Word faylını seçin — yeni qeyd yaranır, xanalar fayldan doldurulur">📄 Fayldan yeni müqavilə</button>`:""}<button class="btn" id="rNew">+ Yeni ${esc(C.one.toLocaleLowerCase("az"))}</button>
      <button class="btn" id="rFiles" title="${ed?"Seçilmiş qeydə fayl əlavə et":"Formadakı yeni qeydi yadda saxlayıb fayl əlavə edir"}">📎 Fayl əlavə et${ed?": "+esc((t=>t.length>28?t.slice(0,27)+"…":t)(ed.no?"№ "+ed.no:ed.name||ed.title||dmy(ed.date))):""}</button>
      ${kind==="orgs"?`<button class="btn" id="oNewXl" title="Üzv siyahısı (Excel) ilə yeni ilk təşkilat yaratmaq">📥 Excel-dən yeni təşkilat</button><input type="file" id="oNewFile" accept=".xlsx,.xls" hidden>`:""}
      <button class="btn" id="rExp">Excel-ə çıxar</button></div>
    <div class="reg-tools">${C.year?`<select id="rYear"><option value="">Bütün illər</option>${years.map(y=>`<option${y===S.y?" selected":""}>${y}</option>`).join("")}</select>`:""}
      <input id="rq" class="srch" type="search" placeholder="Axtar…" value="${esc(S.q)}"></div>
    ${kind==="events"&&L.length?`<div class="sum">${EV_ST.map(s=>`${s}: <b>${L.filter(x=>(x.status||"Planlaşdırılıb")===s).length}</b>`).join(" · ")}</div>`:`<div class="sum"><b>${L.length}</b> qeyd</div>`}
    ${MAN&&L.length>1?`<div class="sum mut drag-tip">✋ Sıranı dəyişmək üçün sətrin solundakı <b>⠿</b> işarəsini tutub yuxarı-aşağı sürüşdürün.${list.some(x=>x.ord!=null)?` <button class="btn sm" type="button" id="rOrdReset" title="Əl ilə verilmiş sıranı ləğv edib tarixə görə düzür">Tarixə görə düz</button>`:""}</div>`:""}
    <div class="tbl-wrap"><table class="reg${MAN?" dragt":""}"><thead><tr>${MAN?'<th class="drag-th"></th>':""}<th class="num">№</th>${C.cols.map(([h,,cls])=>`<th class="${cls||""}">${h}</th>`).join("")}</tr></thead><tbody>
    ${body&&C.totals?`</tbody><tfoot>${C.totals(L).map(r=>`<tr class="tot"><td></td><td>${esc(r[0])}</td>${r.slice(1).map(v=>`<td class="num">${v||""}</td>`).join("")}</tr>`).join("")}</tfoot><tbody>`:""}
    ${body||`<tr><td colspan="${C.cols.length+1+(MAN?1:0)}"><div class="empty"><b>Hələ qeyd yoxdur</b>Soldakı formadan əlavə edin. Hər qeydə fayl (PDF, Word, şəkil) əlavə etmək olar.</div></td></tr>`}</tbody></table></div>
    ${unmatched.length?`<details class="chk"><summary><span class="st pend">${unmatched.length}</span> üzv heç bir ilk təşkilata aid edilmədi</summary><ul>${[...new Set(unmatched.map(m=>m.unit))].map(u=>`<li>${esc(u)}</li>`).join("")}</ul><p class="mut">Uyğun təşkilatı açıb «Açar sözlər» xanasına bu bölmənin adından bir söz əlavə edin.</p></details>`:""}</section></div>`}
function filesHTML(kind,rec){
  if(!rec)return`<div class="files empty-f">${HAS_API?`<button class="btn" type="button" id="fAddNew">+ Fayl əlavə et</button> <span class="mut">— qeyd avtomatik yadda saxlanılacaq</span>`:"Fayl əlavə etmək yalnız proqramda (EXE) işləyir."}</div>`;
  const F=rec.files||[];
  return`<div class="files" id="dropZone">${F.map((f,i)=>`<div class="file"><span class="fn" title="${esc(f.name)}">${esc(f.name)}</span><span class="mut">${fmtSize(f.size||0)} · ${dmy((f.added||"").slice(0,10))}</span>
    <span class="fa"><button class="btn ghost sm" type="button" data-fopen="${i}">Aç</button><button class="btn ghost sm" type="button" data-fshow="${i}" title="Qovluqda göstər">Qovluq</button><button class="btn ghost sm" type="button" data-frep="${i}" title="Yeni fayl seçin — köhnəsi silinəcək">Əvəz et</button><button class="btn ghost sm" type="button" data-fmove="${i}" title="Faylı başqa qeydə köçür">Köçür</button><button class="btn ghost sm" type="button" data-fdel="${i}" title="Sil">Sil</button></span></div>`).join("")}
    <div class="drop">${HAS_API?`<button class="btn" type="button" id="fAdd">+ Fayl əlavə et</button> <span class="mut">və ya faylları bura sürüşdürün</span>`:`<span class="mut">Fayl əlavə etmək yalnız proqramda (EXE) işləyir.</span>`}</div></div>`}
function readReg(kind){const f=new FormData($("#rForm"));const o={};for(const[k,,t]of REG[kind].fields){const v=f.get(k);o[k]=t==="area"?String(v??"").trim():normName(v)}return o}
/* sətirləri tutub sürüşdürmək (siçan və barmaqla) */
function bindDrag(kind){const tb=$("#view table.dragt tbody");if(!tb)return;
  $("#rOrdReset")?.addEventListener("click",()=>{if(!guard())return;DB[kind].forEach(x=>delete x.ord);persist();toast("Tarixə görə düzüldü");render()});
  tb.querySelectorAll(".drag-h").forEach(h=>{h.addEventListener("pointerdown",e=>{if(e.button>0)return;e.preventDefault();e.stopPropagation();
    const row=h.closest("tr");const before=[...tb.querySelectorAll("tr[data-id]")].map(r=>r.dataset.id);
    row.classList.add("dragging");document.body.classList.add("drag-on");
    const move=ev=>{const rows=[...tb.querySelectorAll("tr[data-id]")].filter(r=>r!==row);let tgt=null;
      for(const r of rows){const b=r.getBoundingClientRect();if(ev.clientY<b.top+b.height/2){tgt=r;break}}
      if(tgt?row.nextElementSibling!==tgt:tb.lastElementChild!==row)tgt?tb.insertBefore(row,tgt):tb.appendChild(row);
      if(ev.clientY<60)window.scrollBy(0,-12);else if(ev.clientY>innerHeight-60)window.scrollBy(0,12)};
    const up=()=>{window.removeEventListener("pointermove",move);window.removeEventListener("pointerup",up);window.removeEventListener("pointercancel",up);
      row.classList.remove("dragging");document.body.classList.remove("drag-on");window.__dragJust=true;setTimeout(()=>window.__dragJust=false,300);
      const after=[...tb.querySelectorAll("tr[data-id]")].map(r=>r.dataset.id);if(after.join()===before.join())return;
      if(!guard()){render();return}
      /* görünən (süzülmüş) sətirlərin yeni sırasını bütün siyahıya köçürürük */
      const full=DB[kind].slice().sort(regSort(kind)).map(x=>x.id),vis=new Set(after);let k=0;
      const order=full.map(id=>vis.has(id)?after[k++]:id);order.forEach((id,i)=>{const r=DB[kind].find(x=>x.id===id);if(r)r.ord=i+1});
      persist();toast("Sıra yadda saxlandı");render()};
    window.addEventListener("pointermove",move);window.addEventListener("pointerup",up);window.addEventListener("pointercancel",up)})})}
function bindReg(kind){const C=REG[kind],S=UI[kind];
  bindRows(id=>{if(window.__dragJust){window.__dragJust=false;return}S.sel=id;S.draft=null;S.auto=[];render()});
  if(DRAGK[kind])bindDrag(kind);
  const f=$("#rForm");f.oninput=()=>{if(!S.sel)S.draft=readReg(kind)};
  /* qeydi yadda saxlamaq; uğurlu olsa qeydin özünü qaytarır */
  const AF=["prots","plens","contracts"].includes(kind);
  const save=(quiet,lenient)=>{const o=readReg(kind),msg=$("#rMsg");
    const miss=C.fields.filter(([k,,,req])=>req&&!o[k]).map(([,l])=>l.replace(" *",""));
    if(miss.length&&!lenient){msg.className="form-msg err";msg.textContent=(quiet?"Fayl əlavə etmək üçün əvvəlcə doldurun: ":"Doldurun: ")+miss.join(", ");msg.scrollIntoView({block:"center"});return null}
    if(!guard())return null;const isNew=!S.sel;
    if(S.sel){const i=DB[kind].findIndex(x=>x.id===S.sel),prev=DB[kind][i];DB[kind][i]={...prev,...o};
      if(kind==="orgs"&&prev.name!==o.name)for(const c of DB.contracts)if(c.org===prev.name)c.org=o.name}else{const r={...o,id:newId(),files:[]};DB[kind].push(r);S.sel=r.id;S.draft=null}
    if(!persist())return null;if(!quiet){toast(isNew?"Əlavə edildi — indi fayl əlavə edə bilərsiniz":"Yadda saxlandı");render()}
    return DB[kind].find(x=>x.id===S.sel)};
  /* fayl əlavə olunandan sonra: protokol/plenumda boş xanalar faylın mətnindən doldurulur */
  const finish=async(r,got)=>{r.files=[...(r.files||[]),...got];let msg=`${got.length} fayl əlavə edildi`;S.auto=[];
    if(AF){const filled=await autoFillFrom(kind,r,got);
      if(filled&&filled.length){S.auto=filled;msg=`Fayldan dolduruldu: ${filled.map(k=>AF_LABEL[k]).join(", ")} — yoxlayın`}
      else if(filled===null)msg+=" · faylın mətni oxunmadı — xanaları özünüz doldurun";
      else msg+=" · faylda doldurmaq üçün yeni məlumat tapılmadı"}
    commit(msg);const miss=C.fields.filter(([k,,,req])=>req&&!r[k]).map(([,l])=>l.replace(" *",""));
    if(miss.length){const m=$("#rMsg");if(m){m.className="form-msg err";m.textContent="Doldurun: "+miss.join(", ")}}
    /* müqavilədə titul vərəqi çox vaxt skan olur: tapılmayan xanaları formada göstəririk */
    else if(kind==="contracts"&&S.auto.length){const lack=["no"].filter(k=>!r[k]).map(k=>AF_LABEL[k]);const m=$("#rMsg");if(m&&lack.length){m.className="form-msg";m.textContent=`Faylda tapılmadı: ${lack.join(", ")} (titul vərəqi skan ola bilər). Özünüz yazın və «Yadda saxla» basın.`}}};
  const saveAndAdd=async()=>{const wasNew=!S.sel;const r=save(true,AF);if(!r)return;const got=await window.api.filesAdd(kind,r.id);
    if(got&&got.length){await finish(r,got);return}
    if(wasNew&&AF&&!C.fields.some(([k])=>k!=="type"&&String(r[k]||"").trim())){DB[kind]=DB[kind].filter(x=>x!==r);S.sel=null;persist();render();return}
    toast("Qeyd yadda saxlandı");render()};
  f.onsubmit=e=>{e.preventDefault();S.auto=[];save(false)};
  $("#fAddNew")?.addEventListener("click",saveAndAdd);
  $("#rClose")?.addEventListener("click",()=>{S.sel=null;render()});
  $("#rDel")?.addEventListener("click",async()=>{const r=DB[kind].find(x=>x.id===S.sel);if(!r)return;const n=(r.files||[]).length;
    if(!confirm(`Bu qeyd silinsin?${n?`\nOna əlavə olunmuş ${n} fayl da kompüterdən silinəcək.`:""}`)||!guard())return;
    if(HAS_API)for(const fl of r.files||[])await window.api.fileDel(fl.rel);
    DB[kind]=DB[kind].filter(x=>x!==r);S.sel=null;commit("Silindi")});
  $("#rYear")?.addEventListener("change",e=>{S.y=e.target.value;render()});
  searchBox("#rq",v=>S.q=v);
  $("#rExp").onclick=()=>exportSheets([regSheet(kind)],C.title);
  $("#rNew").onclick=()=>{S.sel=null;S.draft=null;render();$("#rForm [name]")?.focus()};
  $("#ctPdf")?.addEventListener("click",()=>{S.sel=null;S.draft=null;S.auto=[];render();$("#fAddNew")?.click()});
  $("#rFiles").onclick=()=>{const r=S.sel&&DB[kind].find(x=>x.id===S.sel);
    if(!HAS_API){toast("Fayl əlavə etmək yalnız proqramda (EXE) işləyir");return}
    if(!r){saveAndAdd();return}$("#fAdd").click()};
  if(kind==="orgs"){$("#oNewXl")?.addEventListener("click",()=>$("#oNewFile").click());$("#oNewFile")?.addEventListener("change",e=>{const f=e.target.files[0];e.target.value="";importOrgFile(f,null)})}
  const rec=S.sel&&DB[kind].find(x=>x.id===S.sel);
  if(kind==="orgs")bindOrgImport(rec||null);if(!rec)return;
  /* fayl əlavə etməzdən əvvəl formada yazılanları yadda saxlayırıq ki, itməsin */
  const addFiles=async(getP)=>{if(!guard())return;const r=save(true,true)||rec;const got=await getP(r);if(!got||!got.length){render();return}await finish(r,got)};
  $("#fAdd")?.addEventListener("click",()=>addFiles(r=>window.api.filesAdd(kind,r.id)));
  const dz=$("#dropZone");if(dz&&HAS_API){dz.ondragover=e=>{e.preventDefault();dz.classList.add("over")};dz.ondragleave=()=>dz.classList.remove("over");
    dz.ondrop=e=>{e.preventDefault();dz.classList.remove("over");if(e.dataTransfer.files.length){const fl=e.dataTransfer.files;addFiles(r=>window.api.filesAddDropped(kind,r.id,fl))}}}
  $$("[data-fopen]").forEach(b=>b.onclick=async()=>{const r=await window.api?.fileOpen(rec.files[+b.dataset.fopen].rel);if(r)toast(r)});
  $$("[data-fshow]").forEach(b=>b.onclick=()=>window.api?.fileShow(rec.files[+b.dataset.fshow].rel));
  $$("[data-frep]").forEach(b=>b.onclick=async()=>{if(!guard())return;const i=+b.dataset.frep,old=rec.files[i];
    const nf=await window.api.fileReplace(kind,rec.id,old.rel);if(!nf)return;rec.files=rec.files.map((f,j)=>j===i?nf:f);commit(`«${old.name}» → «${nf.name}» ilə əvəz edildi`)});
  $$("[data-fmove]").forEach(b=>b.onclick=()=>{if(!guard())return;const fl=rec.files[+b.dataset.fmove];moveDialog(kind,rec,fl)});
  $$("[data-fdel]").forEach(b=>b.onclick=async()=>{const fl=rec.files[+b.dataset.fdel];if(!confirm(`«${fl.name}» faylı kompüterdən silinsin?`)||!guard())return;
    if(HAS_API)await window.api.fileDel(fl.rel);rec.files=rec.files.filter(x=>x!==fl);commit("Fayl silindi")})}

/* Faylı başqa bölməyə / qeydə köçürmək pəncərəsi */
const recLabel=(k,x)=>k==="orgs"?x.name:k==="events"?`${dmy(x.date)} — ${x.title||""}`:k==="contracts"?`${dmy(x.date)} — ${x.type||""}${x.no?" № "+x.no:""}${x.org?" · "+x.org:""}`:`${x.no?"№ "+x.no+" · ":""}${dmy(x.date)}${x.body?" · "+x.body:""}${x.agenda?" · "+String(x.agenda).slice(0,50):""}`;
function moveDialog(kind,rec,fl){const md=document.createElement("div");md.className="lic-modal";
  const kinds=["prots","plens","orgs","contracts","events"];
  const opts=k=>DB[k].slice().sort(REG[k].sort).filter(x=>x!==rec).map(x=>`<option value="${x.id}">${esc(recLabel(k,x).slice(0,110))}</option>`).join("");
  md.innerHTML=`<div class="lic-box" role="dialog" aria-modal="true"><h2>Faylı köçür</h2><p class="mut">«${esc(fl.name)}» hazırda: <b>${esc(REG[kind].one)}</b> — ${esc(recLabel(kind,rec).slice(0,80))}</p>
    <div class="fld" style="width:100%"><label for="mvKind">Hansı bölməyə</label><select id="mvKind" class="ask-in">${kinds.map(k=>`<option value="${k}"${k===kind?" selected":""}>${REG[k].title}</option>`).join("")}</select></div>
    <div class="fld" style="width:100%"><label for="mvRec">Hansı qeydə</label><select id="mvRec" class="ask-in"></select></div>
    <div class="form-msg" id="mvMsg"></div><div class="form-actions"><button class="btn primary" type="button" id="mvOk">Köçür</button><button class="btn" type="button" id="mvNo">Ləğv et</button></div></div>`;
  document.body.append(md);const fillRecs=()=>{const k=md.querySelector("#mvKind").value;const o=opts(k);md.querySelector("#mvRec").innerHTML=o||'<option value="">— bu bölmədə qeyd yoxdur —</option>'};fillRecs();
  md.querySelector("#mvKind").onchange=fillRecs;md.querySelector("#mvNo").onclick=()=>md.remove();
  md.querySelector("#mvOk").onclick=async()=>{const k=md.querySelector("#mvKind").value,id=md.querySelector("#mvRec").value;const tgt=DB[k].find(x=>x.id===id);
    if(!tgt){const m=md.querySelector("#mvMsg");m.className="form-msg err";m.textContent="Əvvəlcə həmin bölmədə qeyd yaradın və ya başqa qeyd seçin.";return}
    const nf=await window.api.fileMove(fl.rel,k,tgt.id);if(!nf){toast("Fayl köçürülmədi");return}
    rec.files=rec.files.filter(x=>x!==fl);tgt.files=[...(tgt.files||[]),nf];md.remove();
    commit(`«${nf.name}» köçürüldü: ${REG[k].one} — ${recLabel(k,tgt).slice(0,60)}`)}}

/* Yeni qoşulan ilk təşkilatın üzvlərini yükləmək: mövcud siyahıya əlavə edir / yeniləyir */
/* Üzv siyahısını (Excel) təşkilata yükləmək. org=null olduqda yeni ilk təşkilat yaradılır:
   adı formadan, yoxdursa fayldakı «Struktur bölmə»nin ən çox təkrarlanan adından, o da yoxdursa faylın adından götürülür */
function guessOrgName(L,file){const c=new Map();for(const m of L){const h=normName(String(m.unit||"").split(/[,\n]/)[0]).replace(/Mərkəzinin$/,"Mərkəzi");if(h)c.set(h,(c.get(h)||0)+1)}
  const top=[...c].sort((a,b)=>b[1]-a[1])[0];return top?top[0]:normName(String(file.name).replace(/\.[^.]+$/,"").replace(/[_]+/g," "))}
async function importOrgFile(file,org){
  if(!file||!guard())return;
  let L;try{L=parseMembers(await readWb(file))}catch(x){toast("Excel faylı oxunmadı");return}
  L=L.filter(m=>!m.out);
  if(!L.length){toast("Faylda «Soyadı, adı, ata adı» və «FİN» sütunları tapılmadı");return}
  let isNew=false;
  if(!org){const typed=normName($("#r-name")?.value||"");let name=typed||guessOrgName(L,file);
    const same=DB.orgs.find(o=>lc(o.name)===lc(name));
    if(same)org=same;
    else{if(!typed){const nn=prompt("Yeni ilk təşkilatın adı (fayldan götürüldü, dəyişə bilərsiniz):",name);if(nn===null)return;name=normName(nn)||name}
      const f=new FormData($("#rForm")||document.createElement("form"));const g=k=>normName(f.get(k)||"");
      org={id:newId(),name,keys:lc(name),state:"Fəaliyyətdədir",chair:g("chair"),phone:g("phone"),founded:g("founded"),note:String(f.get("note")||"").trim(),files:[]};isNew=true}}
  const byKey=new Map(DB.mem.map(m=>[memKey(m),m]));let add=0,upd=0,moved=0;const other=[];
  for(const r of L){const ex=byKey.get(memKey(r));if(ex){const was=memOrg(ex);if(was&&was!==org){moved++;other.push(`${ex.name} (${was.name})`)}upd++}else add++}
  if(!confirm(`${isNew?`Yeni ilk təşkilat yaradılacaq: «${org.name}»\n\n`:`«${org.name}» üçün `}faylda ${L.length} nəfər tapıldı.\nYeni üzv kimi əlavə olunacaq: ${add}\nYenilənəcək (artıq siyahıda olan): ${upd}${moved?`\n\nDİQQƏT: ${moved} nəfər hazırda başqa təşkilatdadır və bu təşkilata keçiriləcək:\n${other.slice(0,10).join("\n")}${moved>10?"\n…":""}`:""}\n\nDavam edilsin?`))return;
  if(HAS_API)window.api.snapshot("ilk-teskilat-oncesi");
  if(isNew)DB.orgs.push(org);
  for(const r of L){const ex=byKey.get(memKey(r));const rec={...r,unit:r.unit||org.name,orgId:org.id};
    if(ex){const i=DB.mem.indexOf(ex);DB.mem[i]={...ex,...Object.fromEntries(Object.entries(rec).filter(([,v])=>v!==""&&v!=null)),id:ex.id,gender:rec.gender||ex.gender}}
    else DB.mem.push({...rec,id:newId()})}
  memVer++;UI.orgs.sel=org.id;UI.orgs.draft=null;
  commit(isNew?`«${org.name}» yaradıldı: ${L.length} üzv. İlk təşkilatlar: ${DB.orgs.filter(orgActive).length}`:`${org.name}: ${add} əlavə edildi, ${upd} yeniləndi`)}
function bindOrgImport(org){
  $("#oShow")?.addEventListener("click",()=>{UI.tab="mem";Object.assign(UI.mem,{view:"in",org:org.name,st:"",g:"",q:"",sel:null});render()});
  $("#oImp")?.addEventListener("click",()=>$("#oFile").click());
  $("#oFile")?.addEventListener("change",e=>{const f=e.target.files[0];e.target.value="";importOrgFile(f,org)})}

/* ---------- Parametrlər ---------- */
function settingsHTML(){const bi=HAS_API?window.api.backupInfo():null;
  const ls=LIC.status==="ok"?`<span class="st done">Aktivdir${LIC.demo?" (demo)":""}</span> ${esc(LIC.org||"")}${LIC.until?(LIC.until>="9999"?" · müddətsiz":" · "+dmy(LIC.until)+"-dək"):""}${LIC.id?` <span class="mut">· № ${esc(LIC.id)}</span>`:""}`:LIC.status==="expired"?`<span class="st rej">Müddəti bitib</span>`:LIC.status==="invalid"?`<span class="st rej">${esc(LIC_REASON[LIC.reason]||"Etibarsız açar")}</span>`:`<span class="st pend">Lisenziya yoxdur</span>`;
  return`<div class="stack narrow">
  <section class="card"><h2>Təşkilat</h2><form class="entry" id="sForm"><div class="fld"><label for="s-org">Rayon komitəsinin adı</label><input id="s-org" value="${esc(DB.org)}"></div>
    <div class="fld"><label for="s-chair">Sədr (çap üçün)</label><input id="s-chair" value="${esc(DB.chair||"")}"></div><div class="form-actions"><button class="btn primary">Yadda saxla</button></div></form></section>
  <section class="card noweb"><h2>Bütün məlumatlar Excel-də</h2><p class="hint">Komissiyalar faylınızın quruluşunda bir Excel kitabı: hər komissiya, Veteran/Qazi/Şəhid ailəsi, Üzvlərin siyahısı, İlk təşkilatlar, Tədbirlər planı, Protokollar, Plenumlar, Müqavilələr ayrıca vərəqdə.</p>
    <button class="btn primary" id="allExp">Hamısını Excel-ə çıxar</button></section>
  ${HAS_API?`<section class="card"><h2>Ehtiyat nüsxə</h2><p class="hint">Məlumatlar: <code>${esc(window.api.dataPath())}</code>. Hər gün avtomatik nüsxə saxlanılır. Fləş və ya OneDrive qovluğu seçsəniz, proqram bağlananda məlumatlar və bütün fayllar ora da köçürülür.</p>
    <p>${bi.dir?`Qovluq: <b>${esc(bi.dir)}</b>${bi.exists?"":' <span class="st rej">tapılmadı</span>'}${bi.last?` · son nüsxə: ${dmy(bi.last.slice(0,10))}`:""}`:'<span class="mut">Xarici qovluq seçilməyib.</span>'}</p>
    <div class="form-actions"><button class="btn" id="bkChoose">${bi.dir?"Qovluğu dəyiş":"Qovluq seç"}</button>${bi.dir?`<button class="btn" id="bkNow">İndi nüsxə çıxar</button><button class="btn ghost" id="bkClear">Ləğv et</button>`:""}<button class="btn ghost" id="bkOpen">Nüsxələrə bax</button><button class="btn ghost" id="fDir">Fayllar qovluğu</button><button class="btn ghost" id="restore">Nüsxədən bərpa et…</button><button class="btn ghost" id="bkJson" title="Telefona köçürmək üçün">Ehtiyat nüsxə çıxar (.json)</button></div></section>`:`<section class="card"><h2>Məlumatlar (telefon)</h2>
    <p class="hint">Məlumatlar yalnız bu telefonun brauzerində saxlanılır. Kompüterdəki məlumatları köçürmək üçün: kompüterdə <b>Parametrlər → Ehtiyat nüsxə → «Ehtiyat nüsxə çıxar (.json)»</b>, alınan <b>HT_ehtiyat_….json</b> faylını telefona göndərin (Telegram, WhatsApp, e-poçt) və aşağıdan yükləyin.</p>
    <div class="form-actions"><button class="btn primary" type="button" id="wbIn">Ehtiyat nüsxəni yüklə (.json)</button><input type="file" id="wbFile" accept=".json,application/json" hidden></div>
    <div id="wbAsk" class="notice" hidden style="margin-top:12px"></div></section>`}
  <section class="card"><h2>Lisenziya</h2><p>${ls}</p>${midHTML()}
    <p class="hint">Açar almaq üçün kompüter kodunu AzerSoft LLC-yə göndərin. Açar yalnız bu kompüterdə və yalnız HİK-təşkilatda işləyir.</p>
    <div class="fld"><label for="lic-key">Lisenziya açarı</label><textarea id="lic-key" rows="3" class="mono" placeholder="HTS1.…"></textarea></div>
    <div class="form-actions" style="margin-top:8px"><button class="btn primary" id="licSave">Aktivləşdir</button><span class="form-msg" id="licMsg"></span></div></section>
  <p class="mut" style="text-align:center">HİK-təşkilat 1.0.0 · yeniləmə ${UI_VER}${HAS_API&&window.api.uiInfo&&window.api.uiInfo().fromUpdate?" (yeniləmədən)":""} · AzerSoft LLC</p></div>`}
function bindSettings(){
  $("#sForm").onsubmit=e=>{e.preventDefault();if(!guard())return;DB.org=normName($("#s-org").value)||DEF_ORG;DB.chair=normName($("#s-chair").value);commit("Yadda saxlandı")};
  $("#allExp").onclick=()=>exportSheets([...DB.bodies.map(bodySheet),vetSheet(),memSheet(),regSheet("orgs"),regSheet("events"),regSheet("prots"),regSheet("plens"),regSheet("contracts")],"Komissiyalar-Protokol");
  bindMid($("#view"));
  $("#bkJson")?.addEventListener("click",async()=>{const d=todayISO();if(await window.api.saveFile(`HT_ehtiyat_${d}.json`,new TextEncoder().encode(JSON.stringify(DB)),"json"))toast("Ehtiyat nüsxə saxlanıldı")});
  $("#wbIn")?.addEventListener("click",()=>$("#wbFile").click());
  $("#wbFile")?.addEventListener("change",async e=>{const f=e.target.files[0];e.target.value="";if(!f)return;let d;try{d=JSON.parse(await f.text())}catch(x){toast("Fayl oxunmadı");return}
    if(!d||!Array.isArray(d.mem)){toast("Bu HİK-təşkilat ehtiyat nüsxəsi deyil");return}
    const box=$("#wbAsk");box.hidden=false;box.innerHTML=`<p>Faylda: <b>${d.mem.length}</b> üzv, <b>${(d.bodies||[]).length}</b> komissiya, <b>${(d.orgs||[]).length}</b> ilk təşkilat, <b>${(d.prots||[]).length}</b> protokol, <b>${(d.plens||[]).length}</b> plenum, <b>${(d.contracts||[]).length}</b> müqavilə, <b>${(d.events||[]).length}</b> tədbir.<br>Telefondakı hazırkı məlumatlar bu faylla əvəz olunacaq. Əlavə olunmuş fayllar (sənədlər) telefona köçmür, yalnız qeydlər köçür.</p>
      <div class="form-actions"><button class="btn primary" type="button" id="wbYes">Əvəz et</button><button class="btn" type="button" id="wbNo">Ləğv et</button></div>`;
    $("#wbNo").onclick=()=>{box.hidden=true};$("#wbYes").onclick=()=>{DB={...DB,...d};memVer++;persist();toast("Məlumatlar yükləndi");render()}});
  $("#licSave").onclick=()=>{const k=$("#lic-key").value.trim();if(!k)return;const r=licActivate(k);const m=$("#licMsg");
    if(r.status==="invalid"){m.className="form-msg err";m.textContent=LIC_REASON[r.reason]||"Açar etibarsızdır.";return}toast(r.status==="ok"?"Lisenziya aktivləşdi":r.status==="expired"?"Bu açarın müddəti artıq bitib":"Açar qəbul edildi");render()};
  if(!HAS_API)return;
  const after=r=>{if(r&&r.ok)toast("Nüsxə çıxarıldı");else if(r&&r.reason==="missing")toast("Qovluq tapılmadı");render()};
  $("#bkChoose").onclick=async()=>after(await window.api.chooseBackupDir());
  $("#bkNow")?.addEventListener("click",async()=>after(await window.api.backupNow()));
  $("#bkClear")?.addEventListener("click",async()=>{await window.api.clearBackupDir();render()});
  $("#bkOpen").onclick=()=>window.api.openBackupDir();$("#fDir").onclick=()=>window.api.filesDir();
  $("#restore").onclick=async()=>{const t=await window.api.openFile();if(!t)return;let d;try{d=JSON.parse(t)}catch(e){toast("Fayl oxunmadı");return}
    if(!d||!Array.isArray(d.mem)){toast("Bu HİK-təşkilat nüsxəsi deyil");return}
    if(!confirm(`Hazırkı məlumatlar əvəz olunacaq (${d.mem.length} üzv, ${(d.bodies||[]).length} komissiya, ${(d.prots||[]).length} protokol). Davam edilsin?`)||!guard())return;
    window.api.snapshot("berpa-oncesi");DB={...DB,...d};memVer++;commit("Bərpa edildi")}}

/* ---------- Excel ixrac ---------- */
const memSheet=()=>{const L=DB.mem.slice().sort((a,b)=>(a.out-b.out)||((Number(a.no)||1e9)-(Number(b.no)||1e9))||a.name.localeCompare(b.name,"az"));
  return{name:"Üzvlərin siyahısı",head:["№","Struktur bölmə","İlk təşkilat","Vəzifə","Soyadı, adı, ata adı","Cinsi","FİN","Təvəllüdü","Təhsili","Təyinat tarixi","Statusu","Üzvlük","Xaric tarixi","Əmək Məcəlləsinin bəndi","Qeyd","Telefon",...MEM_YN.map(x=>x[1])],
    rows:L.map((m,i)=>[i+1,m.unit,orgName(m),m.pos,m.name,m.gender,m.fin,dmy(m.birth),m.edu,dmy(m.hired),m.status,m.out?"Xaric":"Üzv",dmy(m.outDate),m.outBasis||"",m.outNote,m.phone||"",...MEM_YN.map(x=>m[x[0]]||"")]),text:[6,15],w:[5,40,30,24,32,7,9,11,40,11,14,8,11,45,24,14,10,10,10,10,10,10,10]}};
const vetSheet=()=>{const rows=[];vetGroups().forEach(([s,l])=>l.forEach(m=>rows.push([rows.length+1,s,statusOthers(m.status,s).join(", "),m.name,m.fin,dmy(m.birth),ageOf(m.birth)??"",orgName(m),m.unit,m.pos])));
  return{name:"Veteran, Qazi, Şəhid ailəsi",head:["№","Statusu","Əlavə status","Soyadı, adı, ata adı","FİN","Təvəllüd","Yaş","İlk təşkilat","Struktur bölmə","Vəzifə"],rows,text:[4],w:[5,14,22,32,9,11,5,30,40,24]}};
const bodySheet=b=>({name:b.name.slice(0,31),title:`${b.name}${b.elected?" — seçilib: "+dmy(b.elected):""}`,head:["№","Soyadı, Adı, Ata adı","Çalışdığı müəssisə","Vəzifəsi","Əlaqə nömrəsi","Təvəllüd","FİN","Təhsil","Orqanda"],
  rows:sortPeople(b.people).map((p,i)=>{const v=personView(p);return[i+1,p.name,v.work,v.pos,p.phone,dmy(v.birth),v.fin,eduShort(v.edu),p.role]}),text:[4,6],w:[5,32,30,24,15,11,9,16,12]});
const regSheet=kind=>{const C=REG[kind];const L=DB[kind].slice().sort(C.sort);let nn=0;return{name:C.sheet,head:["№",...C.xl.map(x=>x[0])],rows:[...L.map(x=>[C.noNum&&C.noNum(x)?"":++nn,...C.xl.map(([,f])=>f(x)??"")]),
    ...(C.totals&&L.length?C.totals(L).map(r=>{const row=["",r[0],...C.xl.slice(1).map(()=>"")];C.xl.forEach(([h],i)=>{const j={"Üzvlərin sayı":1,"Qadın":2,"Kişi":3}[h];if(j)row[i+1]=r[j]});return row}):[])],w:[5,...C.xl.map(()=>22)]}};
async function exportSheets(sheets,fname){if(!window.XLSX){toast("Excel kitabxanası yüklənmədi");return}
  const wb=XLSX.utils.book_new(),used=new Set();
  for(const s of sheets){const A=[];if(s.title)A.push([s.title],[]);A.push(s.head,...s.rows);const ws=XLSX.utils.aoa_to_sheet(A);
    ws["!cols"]=(s.w||[]).map(w=>({wch:w}));const off=s.title?3:1;
    for(const c of s.text||[])for(let r=off;r<off+s.rows.length;r++){const a=XLSX.utils.encode_cell({r,c});if(ws[a]){ws[a].t="s";ws[a].z="@"}}
    let n=s.name.replace(/[\\\/?*\[\]:]/g,"").slice(0,31)||"Vərəq",k=2;while(used.has(n))n=n.slice(0,28)+" "+k++;used.add(n);XLSX.utils.book_append_sheet(wb,ws,n)}
  const name=`${fname} ${dmy(todayISO())}.xlsx`;
  if(HAS_API){const buf=XLSX.write(wb,{bookType:"xlsx",type:"array"});if(await window.api.saveFile(name,buf,"xlsx"))toast("Excel faylı saxlanıldı")}else XLSX.writeFile(wb,name)}

/* ---------- ümumi bağlamalar ---------- */
function bindRows(fn){$$("#view tr[data-id]").forEach(r=>{r.onclick=()=>fn(r.dataset.id);r.onkeydown=e=>{if(e.key==="Enter")fn(r.dataset.id)}})}
function searchBox(sel,set){const s=$(sel);if(!s)return;s.oninput=()=>{set(s.value);const p=s.selectionStart;render();const n=$(sel);n.focus();try{n.setSelectionRange(p,p)}catch(e){}}}
const PRINT_SVG='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M6 9V3h12v6"/><rect x="6" y="14" width="12" height="7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/></svg>';
function addPrintBtns(){$$("#view .printable").forEach(s=>{if(s.querySelector(":scope>.print-btn"))return;const b=document.createElement("button");b.className="print-btn";b.type="button";b.innerHTML=PRINT_SVG+"Çap et";b.onclick=e=>{e.stopPropagation();printSection(s)};s.prepend(b)})}
function endPrint(){document.body.classList.remove("printing");$$(".print-target").forEach(x=>x.classList.remove("print-target"));$$(".print-head").forEach(x=>x.remove())}
function printSection(s){endPrint();const h=document.createElement("div");h.className="print-head";
  h.innerHTML=`<b>${esc(DB.org||"")}</b><br>Çap tarixi: ${dmy(todayISO())}`;s.querySelector(":scope>h2")?.after(h);
  s.classList.add("print-target");document.body.classList.add("printing");
  const t=document.title;document.title=s.dataset.ptitle||"HİK-təşkilat";
  if(HAS_API&&window.api.printPreview){window.api.printPreview(document.title).catch(()=>toast("Çap faylı hazırlanmadı")).finally(()=>{document.title=t;endPrint()});return}
  setTimeout(()=>{window.print();document.title=t;endPrint()},50)}
if(HAS_API)window.api.onMenuPrint(()=>{if($("#docView")){$("#docPrint").click();return}const s=document.querySelector("#view .printable");if(s)printSection(s)});
/* ---------- Protokol / plenum: yalnız №, tarix, gündəlik və qərarlar ayrıca səhifədə ---------- */
try{if(!document.getElementById("docCss")){const st=document.createElement("style");st.id="docCss";st.textContent=`
#docView{position:fixed;inset:0;z-index:70;background:#5f6b6e;overflow:auto;padding:0 12px 40px}
#docView .doc-bar .btn{white-space:nowrap}
table.reg td.drag-h{width:26px;padding-left:4px;padding-right:2px;text-align:center;color:var(--muted);font-size:18px;line-height:1.2;cursor:grab;touch-action:none;user-select:none}
table.reg td.drag-h:hover{color:var(--acc,var(--brand))}
table.reg th.drag-th{width:26px;padding:0}
body.drag-on,body.drag-on *{cursor:grabbing!important;user-select:none}
table.reg tr.dragging{background:color-mix(in srgb,var(--acc,var(--brand)) 18%,var(--surface))!important;box-shadow:0 6px 18px -8px rgba(0,0,0,.35);position:relative;z-index:2}
table.reg tr.dragging td.drag-h{color:var(--acc,var(--brand))}
.drag-tip{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
@media print{table.reg td.drag-h,table.reg th.drag-th,.drag-tip{display:none!important}}
.has-doc{position:relative}.has-doc>h2{padding-right:150px}
.doc-btn{position:absolute;top:14px;right:14px;z-index:1;font:inherit;font-size:13px;font-weight:600;display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:8px;border:1px solid var(--acc,var(--brand));background:var(--acc,var(--brand));color:#fff;cursor:pointer;white-space:nowrap}
.doc-btn:hover{filter:brightness(1.1)}.doc-btn:focus-visible{outline:2px solid var(--brand);outline-offset:2px}
@media print{.doc-btn{display:none!important}}
#docView .doc-bar{position:sticky;top:0;z-index:1;display:flex;gap:8px;justify-content:center;flex-wrap:wrap;padding:12px 0;background:#5f6b6e}
#docView .doc-page{background:#fff;color:#000;width:210mm;max-width:100%;min-height:297mm;margin:0 auto;padding:20mm 18mm 20mm 25mm;box-shadow:0 6px 30px rgba(0,0,0,.35);font-family:"Times New Roman",Times,serif;font-size:14pt;line-height:1.45}
#docView .doc-h{text-align:center;font-weight:700;font-size:16pt;letter-spacing:.04em;margin:0 0 4px}
#docView .doc-d{display:flex;justify-content:space-between;margin:0 0 18px;font-size:13pt}
#docView .doc-s{font-weight:700;text-align:center;margin:18px 0 8px;letter-spacing:.06em}
#docView .doc-t{white-space:pre-wrap;text-align:justify;margin:0}
#docView .doc-t.empty{color:#888;font-style:italic}
@media (max-width:700px){#docView .doc-page{padding:14px 14px 24px;min-height:0;font-size:12pt}}
@media print{body.docprint>*:not(#docView){display:none!important}body.docprint #docView{position:static;background:#fff;padding:0;overflow:visible}
 body.docprint #docView .doc-bar{display:none!important}body.docprint #docView .doc-page{box-shadow:none;width:auto;min-height:0;margin:0;padding:0}
 @page{size:A4;margin:20mm 15mm 20mm 25mm}}`;document.head.appendChild(st)}}catch(e){}
/* gündəlik sahəsinə protokolun qalan hissəsi də düşübsə, «Eşidildi / Çıxışlar / Qərara alır» başlığından kəsilir */
function agendaOnly(t){const L=String(t||"").split(/\r?\n/),out=[];
  for(const l of L){const k=l.replace(/[\s:.\-–]+/g,"").toLocaleUpperCase("az").replace(/İ/g,"I");
    if(out.some(x=>x.trim())&&/^(EŞIDILDI|DINLƏNILDI|DINLƏDILDI|ÇIXIŞ|QƏRARAALIR|QƏRARAALINIR|QƏRARALIR|QƏRARVERIR|QƏRARVERDI)/.test(k))break;out.push(l)}
  return out.join("\n").trim().replace(/^(gündəlik|məsələ)\s*:?\s*\n?/i,"").trim()}
/* qərarlar sahəsi natamamdırsa, gündəlikdəki «Qərara alır» hissəsindən götürülür */
function decisionsFrom(t){const L=String(t||"").split(/\r?\n/),out=[];let on=false;
  for(const l of L){const k=l.replace(/[\s:.\-–]+/g,"").toLocaleUpperCase("az").replace(/İ/g,"I");
    if(!on){if(/^(QƏRARAALIR|QƏRARAALINIR|QƏRARALIR|QƏRARVERIR|QƏRARVERDI)/.test(k))on=true;continue}
    if(/^(QƏRARQƏBULOLUNUR|QƏRARQƏBULEDILDI|ICLASINSƏDRI|ICLASSƏDRI|SƏDR|KATIB|IMZA)/.test(k))break;out.push(l)}
  return out.join("\n").trim()}
function docView(kind,id){const o=(DB[kind]||[]).find(x=>x.id===id);if(!o)return;closeDoc();
  const pl=kind==="plens",T=(t,l)=>`<div class="doc-s">${l}</div><p class="doc-t${t&&t.trim()?"":" empty"}">${t&&t.trim()?esc(t.trim()):"(yazılmayıb)"}</p>`;
  const v=document.createElement("div");v.id="docView";v.setAttribute("role","dialog");v.setAttribute("aria-label",pl?"Plenum":"Protokol");
  v.innerHTML=`<div class="doc-bar"><button class="btn primary noweb" type="button" id="docPrint">${PRINT_SVG}Çap et</button><button class="btn" type="button" id="docClose">✕ Bağla</button></div>
  <div class="doc-page"><p class="doc-h">${pl?"PLENUM":"PROTOKOL"}${o.no?" № "+esc(o.no):""}</p>
  <div class="doc-d"><span></span><span>${o.date?dmy(o.date):""}</span></div>
  ${T(agendaOnly(o.agenda),pl?"GÜNDƏLİK":"GÜNDƏLİK (MÜZAKİRƏ OLUNAN MƏSƏLƏLƏR)")}${T((d=>d.length>String(o.decisions||"").trim().length+3?d:o.decisions)(decisionsFrom(o.agenda)),pl?"QƏRARLAR":"QƏBUL EDİLƏN QƏRARLAR")}</div>`;
  document.body.appendChild(v);document.body.style.overflow="hidden";
  v.querySelector("#docClose").onclick=closeDoc;v.querySelector("#docPrint").onclick=()=>printDoc(`${pl?"Plenum":"Protokol"}${o.no?" № "+o.no:""}${o.date?" — "+dmy(o.date):""}`);
  v.querySelector("#docClose").focus()}
function closeDoc(){$("#docView")?.remove();document.body.style.overflow="";document.body.classList.remove("docprint")}
function printDoc(title){document.body.classList.add("docprint");const t=document.title;document.title=title;
  const done=()=>{document.title=t;document.body.classList.remove("docprint")};
  if(HAS_API&&window.api.printPreview){window.api.printPreview(title).catch(()=>toast("Çap faylı hazırlanmadı")).finally(done);return}
  setTimeout(()=>{window.print();done()},50)}
document.addEventListener("keydown",e=>{if(e.key==="Escape"&&$("#docView"))closeDoc()});
document.addEventListener("click",e=>{const d=e.target.closest("[data-docview]");if(d){e.preventDefault();const[k,id]=d.dataset.docview.split(":");docView(k,id)}});
document.addEventListener("click",e=>{const t=e.target.closest(".tab,.subtab");if(t){UI.tab=t.dataset.tab;render();window.scrollTo(0,0)}});
load();render();
/* yeniləmə hazır olanda zolaq göstərilir; düymə ilə proqram bağlanmadan yenilənir (yadda saxlanmamış forma itə bilər) */
if(HAS_API&&window.api.onUiUpdated)window.api.onUiUpdated(v=>{let b=$("#updBar");if(!b){b=document.createElement("div");b.id="updBar";b.className="upd-bar";b.style.cssText="position:sticky;top:0;z-index:70;background:#B45309;color:#fff;border-bottom:0;padding:10px 14px;font-size:15px;font-weight:600;box-shadow:0 3px 12px rgba(0,0,0,.3)";document.body.prepend(b)}
  b.innerHTML=`Proqramın yeni versiyası hazırdır. Formada yazdıqlarınızı yadda saxlayın və <button class="btn" type="button" id="updGo" style="background:#fff;color:#B45309;border-color:#fff;font-weight:700">Yenilə</button>`;$("#updGo").onclick=()=>window.api.uiReload()});
/* proqram yeni versiya ilə açılanda (zolaq görünməsə də) bu barədə qısa xəbər verilir */
if(HAS_API)setTimeout(()=>{try{const k="ht_ui_seen",o=localStorage.getItem(k);if(o!==String(UI_VER)){localStorage.setItem(k,String(UI_VER));toast(`Proqram yeniləndi: yeniləmə ${UI_VER}`)}}catch(e){}},800);
