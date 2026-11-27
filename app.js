const state={token:null,profile:null,adminData:null,exam:null,questionIndex:0,timer:null,remaining:0,saveQueue:Promise.resolve(),list:{},edit:{},importPrev:null,importRes:null,hasilFilter:{jadwalId:'',kelasId:''}};
const $=id=>document.getElementById(id);
function show(id){['loginView','adminView','studentView','examView'].forEach(x=>$(x).classList.add('hidden'));$(id).classList.remove('hidden')}
function toast(msg){const t=$('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3500)}
const API_URL='https://script.google.com/macros/s/AKfycbwc79e30TT28ngUuIyLBgqoee0s3LCqdu9hP5wD4bY2RQlL96p3IMkmovgYKXBF4w0d/exec';
async function call(fn,...args){
  let res;
  try{res=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({fn,args}),redirect:'follow'})}
  catch(e){throw new Error('Tidak dapat terhubung ke server. Periksa koneksi internet.')}
  let out;try{out=await res.json()}catch(e){throw new Error('Respon server tidak valid. Pastikan web app sudah di-deploy dengan akses Anyone.')}
  if(!out.ok)throw new Error(out.error||'Terjadi kesalahan pada server.');
  return out.data;
}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function dateText(v){if(!v)return '-';const d=new Date(v);return isNaN(d)?String(v):d.toLocaleString('id-ID')}
function toLocalInput(v){const d=new Date(v);if(isNaN(d))return '';const p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`}
function opt(value,label,selected){return `<option value="${esc(value)}"${selected?' selected':''}>${esc(label)}</option>`}
function pillStatus(s){return `<span class="pill ${String(s)==='NONAKTIF'?'off':''}">${esc(s||'-')}</span>`}

/* ---------- Unduh / baca file tabel ---------- */
function downloadTable(name,rows){
  if(window.XLSX){
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=rows[0].map((_,i)=>({wch:Math.min(40,Math.max(...rows.map(r=>String(r[i]??'').length))+2)}));
    XLSX.utils.book_append_sheet(wb,ws,'Data');XLSX.writeFile(wb,name+'.xlsx');return;
  }
  const csv=rows.map(r=>r.map(c=>'"'+String(c??'').replace(/"/g,'""')+'"').join(';')).join('\r\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));a.download=name+'.csv';document.body.appendChild(a);a.click();a.remove();
}
function parseCSV(text){
  text=text.replace(/^\uFEFF/,'');
  const first=text.split(/\r?\n/)[0]||'';
  const cnt={',':first.split(',').length,';':first.split(';').length,'\t':first.split('\t').length};
  const delim=Object.keys(cnt).sort((a,b)=>cnt[b]-cnt[a])[0];
  const rows=[];let row=[],cell='',q=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(q){if(c==='"'){if(text[i+1]==='"'){cell+='"';i++}else q=false}else cell+=c}
    else if(c==='"')q=true;
    else if(c===delim){row.push(cell);cell=''}
    else if(c==='\n'){row.push(cell);rows.push(row);row=[];cell=''}
    else if(c!=='\r')cell+=c;
  }
  if(cell!==''||row.length){row.push(cell);rows.push(row)}
  return rows;
}
async function readFileRows(file){
  if(/\.csv$/i.test(file.name))return parseCSV(await file.text());
  if(!window.XLSX)throw new Error('Pustaka Excel belum termuat. Simpan file sebagai CSV atau muat ulang halaman.');
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:false,defval:''});
}

/* ---------- Login / logout ---------- */
$('loginForm').addEventListener('submit',async e=>{e.preventDefault();try{const r=await call('apiLogin',$('username').value,$('password').value,$('loginRole').value);state.token=r.token;state.profile=r.profile;if(r.profile.role==='SISWA'){await loadStudent();}else{ $('adminName').textContent=r.profile.nama;$('adminAvatar').textContent=(r.profile.nama||'A').charAt(0).toUpperCase();show('adminView');await adminPage('dashboard');}}catch(err){toast(err.message)}});
$('adminLogout').onclick=logout;$('studentLogout').onclick=logout;
async function logout(){try{await call('apiLogout',state.token)}catch(e){}state.token=null;state.profile=null;state.exam=null;state.dash=null;state.hasil=null;state.adminData=null;state.list={};state.edit={};state.importPrev=null;state.importRes=null;state.hasilFilter={jadwalId:'',kelasId:''};clearInterval(state.timer);show('loginView');$('loginForm').reset()}
$('adminNav').addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b)adminPage(b.dataset.page)});

/* ---------- Navigasi admin ---------- */
async function adminPage(page,force){
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.page===page));
  const names={dashboard:'Dashboard',students:'Data Siswa',classes:'Data Kelas',questions:'Bank Soal',packages:'Paket Ujian',schedules:'Jadwal Ujian',results:'Hasil Ujian',account:'Akun Admin'};
  $('pageTitle').textContent=names[page]||'Dashboard';
  try{
    const PD={students:['siswa','kelas'],classes:['kelas'],questions:['soal','opsi'],packages:['paket'],schedules:['paket','jadwal'],results:['jadwal','kelas']};
    const tk=state.navTicket=(state.navTicket||0)+1;
    if(force||!state.adminData)state.adminData={};
    if(page==='dashboard'&&!state.dash)$('adminContent').innerHTML='<div class="empty">Memuat data...</div>';
    const need=(PD[page]||[]).filter(k=>!(k in state.adminData));
    if(need.length){
      $('adminContent').innerHTML='<div class="empty">Memuat data...</div>';
      const r=await call('apiGetAdminData',state.token,need);
      if(!r)throw new Error('Data dari server kosong. Deploy ulang web app dengan versi baru.');
      if(tk!==state.navTicket)return;
      Object.assign(state.adminData,r);
    }
    if(page==='dashboard')await renderDashboard(tk);
    if(page==='students')renderStudents();
    if(page==='classes')renderClasses();
    if(page==='questions')renderQuestions();
    if(page==='packages')renderPackages();
    if(page==='schedules')renderSchedules();
    if(page==='results')await loadResults(true);
    if(page==='account')renderAccount();
  }catch(e){toast(e.message)}
}
function paintDashboard(d){$('adminContent').innerHTML=`<div class="page-head"><div><h3>Ringkasan aktivitas</h3><p>Pantau data peserta dan pelaksanaan ujian sekolah.</p></div><span class="pill">Sistem aktif</span></div><div class="stats">${[['Total Siswa',d.siswa,'♙'],['Bank Soal',d.soal,'✎'],['Paket Ujian',d.paket,'▣'],['Sesi Berlangsung',d.sesiAktif,'◷']].map(x=>`<div class="panel stat"><span class="stat-icon">${x[2]}</span><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join('')}</div><div class="panel table-panel"><h3>Hasil ujian terbaru</h3>${table(['Nama Siswa','Mata Pelajaran','Nilai','Waktu'],d.hasilTerbaru.map(x=>[esc(x.nama),esc(x.mapel),x.nilai,dateText(x.tanggal)]))}</div>`}
async function renderDashboard(tk){if(state.dash)paintDashboard(state.dash);const d=await call('apiAdminDashboard',state.token);state.dash=d;if(tk===state.navTicket)paintDashboard(d)}
function table(headers,rows){if(!rows.length)return '<div class="empty">Belum ada data.</div>';return `<div class="table-wrap"><table class="data-table"><thead><tr>${headers.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(x=>`<td>${x??'-'}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`}
function pageHead(title,desc,button=''){return `<div class="page-head"><div><h3>${title}</h3><p>${desc}</p></div>${button}</div>`}

/* ---------- Daftar dengan pencarian & halaman ---------- */
const LISTS={};
function listPanel(key,title,headers,items,textFn,rowFn){
  LISTS[key]={title,headers,items,textFn,rowFn};
  const L=state.list[key]=state.list[key]||{q:'',p:1};
  return `<div class="panel table-panel"><div class="list-head"><h3 id="lt_${key}"></h3><input class="search" data-search="${key}" placeholder="Cari..." value="${esc(L.q)}"></div><div id="lb_${key}"></div></div>`;
}
function paintList(key){
  const D=LISTS[key],L=state.list[key],tt=$('lt_'+key),bb=$('lb_'+key);if(!D||!tt||!bb)return;
  const q=L.q.trim().toLowerCase(),f=q?D.items.filter(it=>D.textFn(it).toLowerCase().includes(q)):D.items;
  const per=25,pages=Math.max(1,Math.ceil(f.length/per));if(L.p>pages)L.p=pages;if(L.p<1)L.p=1;
  tt.textContent=D.title+' ('+f.length+(q?' dari '+D.items.length:'')+')';
  const pager=pages>1?`<div class="pager"><button class="btn light small-btn" data-lk="${key}" data-go="${L.p-1}" ${L.p<=1?'disabled':''}>‹ Sebelumnya</button><span class="muted tiny">Halaman ${L.p} dari ${pages}</span><button class="btn light small-btn" data-lk="${key}" data-go="${L.p+1}" ${L.p>=pages?'disabled':''}>Berikutnya ›</button></div>`:'';
  bb.innerHTML=table(D.headers,f.slice((L.p-1)*per,L.p*per).map(D.rowFn))+pager;
}
function actBtns(t,id,status,extra=''){
  const off=String(status)==='NONAKTIF',a=`data-t="${t}" data-id="${esc(id)}"`;
  return `<div class="actions"><button class="btn light mini" data-act="edit" ${a}>Edit</button>${extra}<button class="btn light mini" data-act="status" ${a} data-to="${off?'AKTIF':'NONAKTIF'}">${off?'Aktifkan':'Nonaktifkan'}</button><button class="btn danger mini" data-act="del" ${a}>Hapus</button></div>`;
}
const TBL={
  siswa:{key:'siswa',id:'SiswaID',page:'students',name:r=>r.NamaSiswa},
  kelas:{key:'kelas',id:'KelasID',page:'classes',name:r=>r.NamaKelas},
  soal:{key:'soal',id:'SoalID',page:'questions',name:r=>String(r.Pertanyaan).slice(0,50)},
  paket:{key:'paket',id:'PaketID',page:'packages',name:r=>r.NamaPaket},
  jadwal:{key:'jadwal',id:'JadwalID',page:'schedules',name:r=>r.NamaJadwal}
};
function formTitle(t,label){return state.edit[t]?`Edit ${label}`:`Tambah ${label}`}
function cancelBtn(t){return state.edit[t]?`<button type="button" class="btn light" data-act="cancel" data-t="${t}">Batal edit</button>`:''}

$('adminContent').addEventListener('input',e=>{const k=e.target.dataset&&e.target.dataset.search;if(k){state.list[k].q=e.target.value;state.list[k].p=1;paintList(k)}});
$('adminContent').addEventListener('click',async e=>{
  const b=e.target.closest('[data-act],[data-go]');if(!b||b.disabled)return;
  if(b.dataset.go!==undefined){state.list[b.dataset.lk].p=Number(b.dataset.go);paintList(b.dataset.lk);return}
  const act=b.dataset.act,t=b.dataset.t,id=b.dataset.id,T=TBL[t];
  try{
    if(act==='edit'){state.edit[t]=state.adminData[T.key].find(r=>String(r[T.id])===id)||null;await adminPage(T.page);window.scrollTo({top:0,behavior:'smooth'})}
    else if(act==='cancel'){state.edit[t]=null;await adminPage(TBL[t].page)}
    else if(act==='status'){
      const rec=state.adminData[T.key].find(r=>String(r[T.id])===id),to=b.dataset.to;
      if(to==='NONAKTIF'&&!confirm(`Nonaktifkan "${T.name(rec)}"? Data tetap tersimpan dan bisa diaktifkan lagi.`))return;
      await call('apiSetStatus',state.token,t,id,to);toast(to==='AKTIF'?'Data diaktifkan':'Data dinonaktifkan');await adminPage(T.page,true);
    }
    else if(act==='del'){
      const rec=state.adminData[T.key].find(r=>String(r[T.id])===id);
      if(!confirm(`Hapus permanen "${T.name(rec)}"? Tindakan ini tidak dapat dibatalkan.`))return;
      await call('apiDeleteData',state.token,t,id);
      if(state.edit[t]&&String(state.edit[t][T.id])===id)state.edit[t]=null;
      toast('Data dihapus');await adminPage(T.page,true);
    }
    else if(act==='resetpw'){
      const rec=state.adminData.siswa.find(r=>r.SiswaID===id);
      const pw=prompt(`Password baru untuk ${rec.NamaSiswa} (${rec.Username}).\nKosongkan untuk dibuat acak (6 karakter). Minimal 6 karakter.`);
      if(pw===null)return;
      const r=await call('apiResetStudentPassword',state.token,id,pw);
      alert(`Password baru berhasil dibuat.\n\nNama: ${r.nama}\nUsername: ${r.username}\nPassword: ${r.password}\n\nCatat sekarang, password tidak dapat dilihat lagi.`);
    }
    else if(act==='dlakun'){downloadAkun()}
    else if(act==='clearimport'){state.importRes=null;renderStudents()}
    else if(act==='tpl'){downloadTable('template_import_siswa',[['Nama','Username','NISN','NIS','Kelas','JK','Password'],['Contoh Siswa','contoh01','0012345678','1234','VI-A','L','']])}
    else if(act==='gentoken'){const f=b.closest('form').elements.token,c='ABCDEFGHJKLMNPQRSTUVWXYZ23456789';let s='';for(let i=0;i<6;i++)s+=c[Math.floor(Math.random()*c.length)];f.value=s}
  }catch(err){toast(err.message)}
});

/* ---------- Siswa ---------- */
function renderStudents(){
  const d=state.adminData,ed=state.edit.siswa,v=ed||{};
  const kelas=d.kelas.filter(k=>k.Status!=='NONAKTIF'||k.KelasID===v.KelasID);
  const kelasName=id=>(d.kelas.find(k=>k.KelasID===id)||{}).NamaKelas||'-';
  $('adminContent').innerHTML=pageHead('Data peserta didik','Tambah, ubah, nonaktifkan, dan hapus akun siswa.')
  +`<div class="panel form-panel"><h3>${formTitle('siswa','siswa')}</h3>${kelas.length?'':'<div class="notice">Buat data kelas terlebih dahulu sebelum menambah siswa.</div>'}<form id="studentForm" class="form-grid"><div><label>Nama lengkap *</label><input name="nama" required value="${esc(v.NamaSiswa||'')}"></div><div><label>Username *</label><input name="username" required value="${esc(v.Username||'')}"></div><div><label>NISN</label><input name="nisn" value="${esc(v.NISN||'')}"></div><div><label>NIS</label><input name="nis" value="${esc(v.NIS||'')}"></div><div><label>Kelas *</label><select name="kelasId" required>${kelas.map(k=>opt(k.KelasID,k.NamaKelas,k.KelasID===v.KelasID)).join('')}</select></div><div><label>Jenis kelamin</label><select name="jk">${opt('','Pilih',!v.JenisKelamin)}${opt('L','L',v.JenisKelamin==='L')}${opt('P','P',v.JenisKelamin==='P')}</select></div><div><label>${ed?'Password baru':'Password awal'}</label><input name="password" placeholder="${ed?'Kosongkan = tidak diubah':'Kosongkan = Siswa123!'}"></div><div class="form-actions"><button class="btn primary">${ed?'Simpan perubahan':'Simpan siswa'}</button>${cancelBtn('siswa')}</div></form></div>`
  +importPanel(kelas)
  +listPanel('siswa','Daftar siswa',['Nama','NISN','Kelas','Username','Status','Aksi'],d.siswa,s=>[s.NamaSiswa,s.NISN,s.Username,kelasName(s.KelasID)].join(' '),
    s=>[esc(s.NamaSiswa),esc(s.NISN),esc(kelasName(s.KelasID)),esc(s.Username),pillStatus(s.Status),actBtns('siswa',s.SiswaID,s.Status,`<button class="btn light mini" data-act="resetpw" data-id="${esc(s.SiswaID)}">Reset PW</button>`)]);
  paintList('siswa');
  $('studentForm').onsubmit=async e=>{e.preventDefault();const x=Object.fromEntries(new FormData(e.target));if(ed)x.id=ed.SiswaID;try{const r=await call('apiSaveStudent',state.token,x);toast(ed?'Perubahan tersimpan':'Siswa tersimpan. '+(r.defaultPassword?'Password awal: '+r.defaultPassword:''));state.edit.siswa=null;adminPage('students',true)}catch(err){toast(err.message)}};
  bindImport();
}
function importPanel(kelas){
  const res=state.importRes;
  const resultHtml=res?`<div class="info-box"><b>Hasil import:</b> ${res.ok.length} berhasil, ${res.fail.length} gagal.${res.ok.length?`<div class="form-actions"><button type="button" class="btn primary small-btn" data-act="dlakun">Unduh daftar akun (username &amp; password)</button><button type="button" class="btn light small-btn" data-act="clearimport">Tutup</button></div><p class="tiny muted">Simpan file ini. Password tidak dapat dilihat lagi setelah halaman ditutup.</p>`:`<div class="form-actions"><button type="button" class="btn light small-btn" data-act="clearimport">Tutup</button></div>`}${res.fail.length?table(['Baris','Nama','Penyebab'],res.fail.map(f=>[f.baris,esc(f.nama||'-'),esc(f.error)])):''}</div>`:'';
  return `<div class="panel form-panel"><h3>Import siswa dari Excel / CSV</h3><p class="muted tiny">Baris pertama berisi judul kolom: <b>Nama, Username, NISN, NIS, Kelas, JK, Password</b>. Username kosong = memakai NISN. Password kosong = dibuat acak otomatis. Kolom Kelas diisi nama kelas persis seperti di Data Kelas.</p><div class="form-grid"><div><label>File (.xlsx, .xls, .csv)</label><input type="file" id="importFile" accept=".xlsx,.xls,.csv"></div><div><label>Kelas default (jika kolom Kelas kosong)</label><select id="importKelas">${opt('','- tidak ada -',true)}${kelas.map(k=>opt(k.KelasID,k.NamaKelas,false)).join('')}</select></div></div><div class="form-actions"><button type="button" class="btn light" data-act="tpl">Unduh template</button><button type="button" class="btn primary" id="importGo" disabled>Mulai import</button></div><div id="importInfo">${resultHtml}</div></div>`;
}
function bindImport(){
  const f=$('importFile'),go=$('importGo');if(!f)return;
  f.onchange=async()=>{
    state.importPrev=null;go.disabled=true;$('importInfo').innerHTML='';
    if(!f.files[0])return;
    try{
      const rows=await readFileRows(f.files[0]);
      const norm=s=>String(s??'').toLowerCase().replace(/[^a-z]/g,'');
      const alias={nama:'nama',namasiswa:'nama',namalengkap:'nama',username:'username',user:'username',nisn:'nisn',nis:'nis',kelas:'kelas',rombel:'kelas',jk:'jk',jeniskelamin:'jk',gender:'jk',password:'password',katasandi:'password',sandi:'password'};
      const head=(rows[0]||[]).map(h=>alias[norm(h)]||'');
      if(head.indexOf('nama')<0)throw new Error('Kolom "Nama" tidak ditemukan di baris pertama. Gunakan template.');
      const items=[];
      rows.slice(1).forEach((r,i)=>{
        if(!r.some(c=>String(c??'').trim()!==''))return;
        const o={baris:i+2};head.forEach((k,j)=>{if(k)o[k]=String(r[j]??'').trim()});
        if(!o.kelas&&$('importKelas').value)o.kelasId=$('importKelas').value;items.push(o);
      });
      if(!items.length)throw new Error('Tidak ada baris data pada file.');
      state.importPrev=items;go.disabled=false;
      $('importInfo').innerHTML=`<div class="info-box"><b>${items.length} baris</b> siap diimport. Contoh 5 baris pertama:${table(['Baris','Nama','Username','NISN','Kelas','Password'],items.slice(0,5).map(x=>[x.baris,esc(x.nama),esc(x.username||x.nisn||'-'),esc(x.nisn||'-'),esc(x.kelas||(x.kelasId?'(kelas default)':'-')),esc(x.password||'(acak)')]))}</div>`;
    }catch(err){toast(err.message)}
  };
  $('importKelas').onchange=()=>{if(state.importPrev)state.importPrev.forEach(x=>{if(!x.kelas){if($('importKelas').value)x.kelasId=$('importKelas').value;else delete x.kelasId}})};
  go.onclick=async()=>{
    const items=state.importPrev;if(!items||!items.length)return;go.disabled=true;f.disabled=true;
    const res={ok:[],fail:[]},size=10;
    for(let i=0;i<items.length;i+=size){
      $('importInfo').innerHTML=`<div class="info-box">Mengimport... ${Math.min(i+size,items.length)} / ${items.length}<div class="progress"><span style="width:${Math.round(Math.min(i+size,items.length)/items.length*100)}%"></span></div></div>`;
      const chunk=items.slice(i,i+size);
      try{
        const r=await call('apiImportStudents',state.token,chunk);
        r.results.forEach(x=>x.ok?res.ok.push(x):res.fail.push(x));
      }catch(err){chunk.forEach(x=>res.fail.push({baris:x.baris,nama:x.nama,error:err.message}))}
    }
    state.importPrev=null;state.importRes=res;toast(`Import selesai: ${res.ok.length} berhasil, ${res.fail.length} gagal`);adminPage('students',true);
  };
}
function downloadAkun(){
  const r=state.importRes;if(!r||!r.ok.length)return;
  downloadTable('daftar_akun_siswa',[['Nama','Username','Password','Kelas'],...r.ok.map(x=>[x.nama,x.username,x.password,x.kelas||''])]);
}

/* ---------- Kelas ---------- */
function renderClasses(){
  const d=state.adminData,ed=state.edit.kelas,v=ed||{};
  $('adminContent').innerHTML=pageHead('Data kelas','Buat dan kelola rombongan belajar peserta.')
  +`<div class="panel form-panel"><h3>${formTitle('kelas','kelas')}</h3><form id="classForm" class="form-grid"><div><label>Nama kelas *</label><input name="nama" placeholder="Contoh: VI-A" required value="${esc(v.NamaKelas||'')}"></div><div><label>Tahun ajaran</label><input name="tahun" value="${esc(v.TahunAjaran||'2026/2027')}"></div><div><label>Wali kelas</label><input name="wali" value="${esc(v.WaliKelas||'')}"></div><div class="form-actions"><button class="btn primary">${ed?'Simpan perubahan':'Simpan kelas'}</button>${cancelBtn('kelas')}</div></form></div>`
  +listPanel('kelas','Daftar kelas',['Nama Kelas','Tahun Ajaran','Wali Kelas','Status','Aksi'],d.kelas,k=>[k.NamaKelas,k.WaliKelas,k.TahunAjaran].join(' '),
    k=>[esc(k.NamaKelas),esc(k.TahunAjaran),esc(k.WaliKelas),pillStatus(k.Status),actBtns('kelas',k.KelasID,k.Status)]);
  paintList('kelas');
  $('classForm').onsubmit=async e=>{e.preventDefault();const x=Object.fromEntries(new FormData(e.target));if(ed)x.id=ed.KelasID;try{await call('apiSaveClass',state.token,x);toast(ed?'Perubahan tersimpan':'Kelas tersimpan');state.edit.kelas=null;adminPage('classes',true)}catch(err){toast(err.message)}};
}

/* ---------- Bank soal ---------- */
const MAPEL=['Bahasa Indonesia','Matematika'],KODE=['A','B','C','D','E'];
function renderQuestions(){
  const d=state.adminData,ed=state.edit.soal,v=ed||{};
  const ops=ed?d.opsi.filter(o=>o.SoalID===ed.SoalID):[];
  const teks=k=>(ops.find(o=>o.KodeOpsi===k)||{}).TeksOpsi||'';
  const kunci=(ops.find(o=>o.IsKunci==='YA')||{}).KodeOpsi||'A';
  $('adminContent').innerHTML=pageHead('Bank soal','Buat dan kelola soal pilihan ganda Bahasa Indonesia dan Matematika.')
  +`<div class="panel form-panel"><h3>${formTitle('soal','soal pilihan ganda')}</h3><form id="questionForm" class="form-grid"><div><label>Mata pelajaran *</label><select name="mapel">${MAPEL.map(m=>opt(m,m,m===v.MataPelajaran)).join('')}</select></div><div><label>Materi</label><input name="materi" value="${esc(v.Materi||'')}"></div><div class="wide"><label>Pertanyaan *</label><textarea name="pertanyaan" rows="3" required>${esc(v.Pertanyaan||'')}</textarea></div><div class="wide"><label>URL gambar (opsional)</label><input name="gambar" placeholder="https://..." value="${esc(v.GambarURL||'')}"></div>${KODE.map((k,i)=>`<div><label>Opsi ${k}${i<2?' *':' (opsional)'}</label><input name="opsi${k}" ${i<2?'required':''} value="${esc(teks(k))}"></div>`).join('')}<div><label>Kunci jawaban</label><select name="benar">${KODE.map(k=>opt(k,k,k===kunci)).join('')}</select></div><div><label>Tingkat kesulitan</label><select name="level">${['Mudah','Sedang','Sulit'].map(l=>opt(l,l,l===(v.TingkatKesulitan||'Sedang'))).join('')}</select></div><div class="wide"><label>Pembahasan (opsional)</label><textarea name="pembahasan" rows="2">${esc(v.Pembahasan||'')}</textarea></div><div class="form-actions"><button class="btn primary">${ed?'Simpan perubahan':'Simpan soal'}</button>${cancelBtn('soal')}</div></form></div>`
  +listPanel('soal','Daftar soal',['Mata Pelajaran','Materi','Pertanyaan','Status','Aksi'],d.soal,q=>[q.MataPelajaran,q.Materi,q.Pertanyaan].join(' '),
    q=>{const t=String(q.Pertanyaan||'');return [esc(q.MataPelajaran),esc(q.Materi),esc(t.length>100?t.slice(0,100)+'…':t),pillStatus(q.Status),actBtns('soal',q.SoalID,q.Status)]});
  paintList('soal');
  $('questionForm').onsubmit=async e=>{e.preventDefault();const f=new FormData(e.target);const x={mapel:f.get('mapel'),materi:f.get('materi'),pertanyaan:f.get('pertanyaan'),gambar:f.get('gambar'),level:f.get('level'),pembahasan:f.get('pembahasan'),opsi:KODE.map(k=>({kode:k,teks:f.get('opsi'+k),benar:f.get('benar')===k}))};if(ed)x.id=ed.SoalID;try{await call('apiSaveQuestion',state.token,x);toast(ed?'Perubahan tersimpan':'Soal berhasil disimpan');state.edit.soal=null;adminPage('questions',true)}catch(err){toast(err.message)}};
}

/* ---------- Paket ---------- */
function renderPackages(){
  const d=state.adminData,ed=state.edit.paket,v=ed||{};
  const yn=(name,cur)=>`<select name="${name}">${opt('true','Ya',cur)}${opt('','Tidak',!cur)}</select>`;
  $('adminContent').innerHTML=pageHead('Paket ujian','Kelompokkan soal menjadi latihan atau simulasi TKA.')
  +`<div class="panel form-panel"><h3>${formTitle('paket','paket')}</h3><form id="packageForm" class="form-grid"><div><label>Nama paket *</label><input name="nama" required value="${esc(v.NamaPaket||'')}"></div><div><label>Jenis</label><select name="jenis">${opt('LATIHAN','Latihan',v.JenisPaket!=='SIMULASI')}${opt('SIMULASI','Simulasi ujian',v.JenisPaket==='SIMULASI')}</select></div><div><label>Mata pelajaran *</label><select name="mapel">${MAPEL.map(m=>opt(m,m,m===v.MataPelajaran)).join('')}</select></div><div><label>Durasi (menit)</label><input name="durasi" type="number" min="1" value="${esc(v.DurasiMenit||60)}"></div><div><label>Jumlah soal (0 = semua)</label><input name="jumlah" type="number" min="0" value="${esc(v.JumlahSoal||0)}"></div><div><label>Acak soal</label>${yn('acak',ed?v.AcakSoal==='YA':true)}</div><div><label>Acak opsi</label>${yn('acakOpsi',ed?v.AcakOpsi==='YA':true)}</div><div><label>Tampilkan hasil ke siswa</label>${yn('tampil',ed?v.TampilkanHasil==='YA':true)}</div><div class="wide"><label>Deskripsi</label><textarea name="deskripsi">${esc(v.Deskripsi||'')}</textarea></div><div class="form-actions"><button class="btn primary">${ed?'Simpan perubahan':'Simpan paket'}</button>${cancelBtn('paket')}</div></form></div>`
  +listPanel('paket','Daftar paket',['Nama Paket','Jenis','Mapel','Durasi','Jumlah Soal','Status','Aksi'],d.paket,p=>[p.NamaPaket,p.JenisPaket,p.MataPelajaran].join(' '),
    p=>[esc(p.NamaPaket),esc(p.JenisPaket),esc(p.MataPelajaran),p.DurasiMenit+' mnt',p.JumlahSoal>0?p.JumlahSoal:'Semua',pillStatus(p.Status),actBtns('paket',p.PaketID,p.Status)]);
  paintList('paket');
  $('packageForm').onsubmit=async e=>{e.preventDefault();const f=new FormData(e.target);const x={nama:f.get('nama'),jenis:f.get('jenis'),mapel:f.get('mapel'),durasi:f.get('durasi'),jumlah:f.get('jumlah'),acak:f.get('acak')==='true',acakOpsi:f.get('acakOpsi')==='true',tampil:f.get('tampil')==='true',deskripsi:f.get('deskripsi')};if(ed)x.id=ed.PaketID;try{await call('apiSavePackage',state.token,x);toast(ed?'Perubahan tersimpan':'Paket tersimpan');state.edit.paket=null;adminPage('packages',true)}catch(err){toast(err.message)}};
}

/* ---------- Jadwal ---------- */
function renderSchedules(){
  const d=state.adminData,ed=state.edit.jadwal,v=ed||{};
  const paket=d.paket.filter(p=>p.Status!=='NONAKTIF'||p.PaketID===v.PaketID);
  const pName=id=>(d.paket.find(p=>p.PaketID===id)||{}).NamaPaket||'-';
  $('adminContent').innerHTML=pageHead('Jadwal ujian','Tentukan rentang waktu peserta dapat memulai ujian.')
  +`<div class="panel form-panel"><h3>${formTitle('jadwal','jadwal')}</h3>${paket.length?'':'<div class="notice">Buat paket ujian terlebih dahulu.</div>'}<form id="scheduleForm" class="form-grid"><div><label>Paket ujian *</label><select name="paketId" required>${paket.map(p=>opt(p.PaketID,p.NamaPaket,p.PaketID===v.PaketID)).join('')}</select></div><div><label>Nama jadwal *</label><input name="nama" required placeholder="Contoh: Simulasi TKA Gelombang 1" value="${esc(v.NamaJadwal||'')}"></div><div><label>Mulai *</label><input name="mulai" type="datetime-local" required value="${ed?toLocalInput(v.TanggalMulai):''}"></div><div><label>Selesai *</label><input name="selesai" type="datetime-local" required value="${ed?toLocalInput(v.TanggalSelesai):''}"></div><div><label>Token (opsional)</label><div class="inline-field"><input name="token" value="${esc(v.TokenUjian||'')}"><button type="button" class="btn light small-btn" data-act="gentoken">Acak</button></div></div><div class="form-actions"><button class="btn primary">${ed?'Simpan perubahan':'Simpan jadwal'}</button>${cancelBtn('jadwal')}</div></form></div>`
  +listPanel('jadwal','Daftar jadwal',['Nama Jadwal','Paket','Mulai','Selesai','Token','Status','Aksi'],d.jadwal,j=>[j.NamaJadwal,pName(j.PaketID)].join(' '),
    j=>[esc(j.NamaJadwal),esc(pName(j.PaketID)),dateText(j.TanggalMulai),dateText(j.TanggalSelesai),esc(j.TokenUjian||'-'),pillStatus(j.Status),actBtns('jadwal',j.JadwalID,j.Status)]);
  paintList('jadwal');
  $('scheduleForm').onsubmit=async e=>{e.preventDefault();const x=Object.fromEntries(new FormData(e.target));if(ed)x.id=ed.JadwalID;try{await call('apiSaveSchedule',state.token,x);toast(ed?'Perubahan tersimpan':'Jadwal tersimpan');state.edit.jadwal=null;adminPage('schedules',true)}catch(err){toast(err.message)}};
}

/* ---------- Hasil ujian ---------- */
async function loadResults(reset){
  if(reset||!state.hasil){state.hasil={rows:[],total:0};$('adminContent').innerHTML='<div class="empty">Memuat data...</div>'}
  try{
    const r=await call('apiHasilPage',state.token,30,state.hasil.rows.length,state.hasilFilter);
    if(!r)throw new Error('Data dari server kosong. Deploy ulang web app dengan versi baru.');
    state.hasil.rows=state.hasil.rows.concat(r.rows);state.hasil.total=r.total;
    paintResults();
  }catch(e){toast(e.message)}
}
function paintResults(){
  const h=state.hasil,ad=state.adminData||{},F=state.hasilFilter;
  const filters=`<div class="panel form-panel filters"><div><label>Jadwal</label><select id="fJadwal">${opt('','Semua jadwal',!F.jadwalId)}${(ad.jadwal||[]).map(j=>opt(j.JadwalID,j.NamaJadwal,j.JadwalID===F.jadwalId)).join('')}</select></div><div><label>Kelas</label><select id="fKelas">${opt('','Semua kelas',!F.kelasId)}${(ad.kelas||[]).map(k=>opt(k.KelasID,k.NamaKelas,k.KelasID===F.kelasId)).join('')}</select></div><div class="filter-actions"><button class="btn light" onclick="loadResults(true)">Muat ulang</button><button class="btn primary" onclick="exportHasil()">Ekspor Excel</button><button class="btn light" onclick="tutupSesi()">Tutup sesi kedaluwarsa</button></div></div>`;
  $('adminContent').innerHTML=pageHead('Hasil ujian','Menampilkan '+h.rows.length+' dari '+h.total+' hasil'+((F.jadwalId||F.kelasId)?' (terfilter)':' terbaru')+'.')+filters
  +`<div class="panel table-panel"><h3>Rekap hasil</h3>${table(['Nama Siswa','Kelas','Mata Pelajaran','Benar','Salah','Kosong','Nilai','Waktu selesai'],h.rows.map(x=>[esc(x.namaSiswa),esc(x.namaKelas),esc(x.MataPelajaran),x.JumlahBenar,x.JumlahSalah,x.JumlahKosong,x.Nilai,dateText(x.WaktuSelesai)]))}${h.rows.length<h.total?'<div class="center" style="margin-top:14px"><button class="btn light" onclick="loadResults()">Muat lebih banyak</button></div>':''}</div>`;
  const apply=()=>{state.hasilFilter={jadwalId:$('fJadwal').value,kelasId:$('fKelas').value};loadResults(true)};
  $('fJadwal').onchange=apply;$('fKelas').onchange=apply;
}
async function exportHasil(){
  try{toast('Menyiapkan file...');const r=await call('apiExportHasil',state.token,state.hasilFilter);
    if(!r.rows.length){toast('Tidak ada data untuk diekspor.');return}
    downloadTable('hasil_ujian_'+new Date().toISOString().slice(0,10),[r.header,...r.rows]);
  }catch(e){toast(e.message)}
}
async function tutupSesi(){
  if(!confirm('Tutup semua sesi ujian yang waktunya sudah habis? Jawaban yang sudah tersimpan akan dinilai.'))return;
  try{const r=await call('apiTutupSesiKedaluwarsa',state.token);state.dash=null;toast(r.ditutup+' sesi ditutup');loadResults(true)}catch(e){toast(e.message)}
}

/* ---------- Akun admin ---------- */
function renderAccount(){
  $('adminContent').innerHTML=pageHead('Akun administrator','Ganti password akun yang sedang digunakan.')
  +`<div class="panel form-panel"><h3>Ganti password</h3><form id="pwForm" class="form-grid"><div class="wide"><label>Password lama *</label><input type="password" name="lama" required autocomplete="current-password"></div><div><label>Password baru * (min. 6 karakter)</label><input type="password" name="baru" required minlength="6" autocomplete="new-password"></div><div><label>Ulangi password baru *</label><input type="password" name="baru2" required minlength="6" autocomplete="new-password"></div><div class="form-actions"><button class="btn primary">Simpan password</button></div></form></div>`;
  $('pwForm').onsubmit=async e=>{e.preventDefault();const x=Object.fromEntries(new FormData(e.target));if(x.baru!==x.baru2){toast('Password baru dan pengulangannya tidak sama.');return}try{await call('apiChangeAdminPassword',state.token,x.lama,x.baru);toast('Password berhasil diganti');e.target.reset()}catch(err){toast(err.message)}};
}

/* ---------- Siswa & ujian ---------- */
async function loadStudent(){const [d,results]=await Promise.all([call('apiStudentHome',state.token),call('apiStudentResults',state.token)]);$('studentName').textContent=d.profile.nama;show('studentView');$('studentContent').innerHTML=`<div class="welcome"><p class="eyebrow" style="color:#d8e4ff">HALO, PESERTA TKA</p><h2>${esc(d.profile.nama)}</h2><p>${esc(d.profile.kelas)} · NISN ${esc(d.profile.nisn||'-')}</p></div><div class="page-head"><div><h3>Ujian tersedia</h3><p>Pilih latihan atau simulasi yang sudah dijadwalkan sekolah.</p></div></div>${d.jadwal.length?d.jadwal.map(j=>`<div class="panel exam-card"><div><span class="pill ${j.status==='SELESAI'?'warn':''}">${esc(j.jenis)}</span><h3>${esc(j.nama)}</h3><p>${esc(j.paket)} · ${esc(j.mapel)} · ${j.durasi} menit</p></div><button class="btn ${j.status==='SELESAI'?'light':'primary'}" ${j.status==='SELESAI'?'disabled':''} onclick="startExam('${esc(j.id)}')">${j.status==='SELESAI'?'Sudah selesai':'Mulai ujian'}</button></div>`).join(''):'<div class="panel empty">Belum ada ujian yang tersedia saat ini.</div>'}<div class="panel student-results"><h3>Riwayat hasil ujian</h3>${table(['Mata Pelajaran','Benar','Total Soal','Nilai','Tanggal'],results.map(r=>[esc(r.mapel),r.benar,r.total,r.nilai,dateText(r.tanggal)]))}</div>`}
async function startExam(id){const token=prompt('Masukkan token ujian jika diberikan sekolah (kosongkan jika tidak ada):')||'';try{state.exam=await call('apiStartExam',state.token,id,token);state.questionIndex=0;show('examView');renderQuestion();startTimer()}catch(e){toast(e.message)}}
function startTimer(){clearInterval(state.timer);const ex=state.exam;const off=(ex.serverNow||Date.now())-Date.now();const end=ex.started+ex.duration*60000;function tick(){state.remaining=Math.max(0,Math.ceil((end-(Date.now()+off))/1000));$('timer').textContent=String(Math.floor(state.remaining/60)).padStart(2,'0')+':'+String(state.remaining%60).padStart(2,'0');if(state.remaining<=0){clearInterval(state.timer);submitExam(true)}}tick();state.timer=setInterval(tick,1000)}
function renderQuestion(){const ex=state.exam,q=ex.questions[state.questionIndex];if(!q)return;$('examSubject').textContent='Ujian berlangsung · '+ex.questions.length+' soal';$('questionCount').textContent=`Soal ${q.no} dari ${ex.questions.length}`;$('questionText').textContent=q.pertanyaan;$('questionImage').classList.toggle('hidden',!q.gambar);if(q.gambar)$('questionImage').src=q.gambar;$('saveState').textContent='Jawaban tersimpan';$('options').innerHTML=q.opsi.map(o=>`<div class="option ${q.jawaban===o.id?'selected':''}" data-option="${o.id}"><span class="option-key">${esc(o.kode)}</span><span>${esc(o.teks)}</span></div>`).join('');$('options').onclick=e=>{const opt=e.target.closest('[data-option]');if(opt)chooseOption(opt.dataset.option)};$('markBtn').textContent=q.ragu?'★ Hapus tanda ragu-ragu':'☆ Tandai ragu-ragu';$('prevBtn').disabled=state.questionIndex===0;$('nextBtn').textContent=state.questionIndex===ex.questions.length-1?'Kembali ke awal':'Berikutnya';$('numberGrid').innerHTML=ex.questions.map((x,i)=>`<button class="number-btn ${x.jawaban?'answered':''} ${x.ragu?'marked':''} ${i===state.questionIndex?'current':''}" data-index="${i}">${i+1}</button>`).join('');$('numberGrid').onclick=e=>{const b=e.target.closest('[data-index]');if(b){state.questionIndex=Number(b.dataset.index);renderQuestion()}}}
function saveCurrent(){const q=state.exam.questions[state.questionIndex],args=[state.exam.sessionId,state.exam.accessToken,q.id,q.jawaban,q.ragu];$('saveState').textContent='Menyimpan...';state.saveQueue=state.saveQueue.then(()=>call('apiSaveAnswer',state.token,...args)).then(()=>{$('saveState').textContent='Jawaban tersimpan'}).catch(e=>{$('saveState').textContent='Gagal menyimpan';toast(e.message)});return state.saveQueue}
function chooseOption(id){const q=state.exam.questions[state.questionIndex];q.jawaban=id;renderQuestion();saveCurrent()}
$('markBtn').onclick=()=>{const q=state.exam.questions[state.questionIndex];q.ragu=!q.ragu;renderQuestion();saveCurrent()};
$('prevBtn').onclick=()=>{if(state.questionIndex>0){state.questionIndex--;renderQuestion()}};
$('nextBtn').onclick=()=>{state.questionIndex=(state.questionIndex+1)%state.exam.questions.length;renderQuestion()};
$('submitExam').onclick=()=>submitExam(false);
async function submitExam(auto){if(!state.exam)return;if(!auto&&!confirm('Yakin ingin menyelesaikan ujian? Jawaban tidak dapat diubah setelah dikirim.'))return;clearInterval(state.timer);try{await state.saveQueue;const r=await call('apiSubmitExam',state.token,state.exam.sessionId,state.exam.accessToken);state.exam=null;show('studentView');if(r.tampilkan){$('studentContent').innerHTML=`<div class="panel result-box"><span class="pill">UJIAN SELESAI</span><h2>Hasil ujian</h2><div class="result-score">${r.nilai}</div><p>Benar: ${r.benar} · Salah: ${r.salah} · Kosong: ${r.kosong}</p><button class="btn primary" onclick="loadStudent()">Kembali ke dashboard</button></div>`}else{toast('Ujian berhasil dikirim. Nilai disimpan oleh sekolah.');await loadStudent()}}catch(e){toast(e.message);if(state.exam&&!auto)startTimer()}}
