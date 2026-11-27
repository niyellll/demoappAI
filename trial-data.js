export function mountTrial(root) {
  const byId = id => root.getElementById(id);
  const formats = {
    plan: ['Call plan', ['SALES ID','NAMA SALESMAN','CustID','Nama Toko','Jadwal Kunjungan'], ['custid','jadwalkunjungan']],
    stores: ['Master toko', ['custid','custname','address','region','village','latitude','longitude'], ['custid','custname','village','latitude','longitude']],
    sales: ['Data penjualan', ['invoice','tanggal','div_id','div_name','custid','custname','itemid','itemname','qty','sub_amount'], ['invoice','tanggal','custid','itemid','subamount']],
    items: ['Master item', ['itemid','struc3','struc4','itemname','KONVERSI','price','FOTO PRODUK'], ['itemid','itemname']]
  };
  const data = {plan:[],stores:[],sales:[],items:[]};
  const filesVersion = {};
  let photoUrls = [], displayUrl = '', optimized = false, baseDistance = 0;
  const checked = new Set();
  const key = value => String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g,'');
  const cell = value => String(value ?? '').trim();
  const code = value => cell(value).toUpperCase();
  const rupiah = new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',maximumFractionDigits:0});
  const km = n => new Intl.NumberFormat('id-ID',{maximumFractionDigits:2}).format(n)+' km';
  const number = value => {
    const text=cell(value).replace(/\s/g,'');
    if(!text) return NaN;
    return Number(text.includes(',') ? (text.lastIndexOf(',')>text.lastIndexOf('.') ? text.replace(/\./g,'').replace(',','.') : text.replace(/,/g,'')) : text);
  };
  function monthOf(value){
    if(value instanceof Date && !isNaN(value)) return value.toISOString().slice(0,7);
    if(typeof value==='number'){
      const date=XLSX.SSF.parse_date_code(value);
      return date ? date.y+'-'+String(date.m).padStart(2,'0') : '';
    }
    const text=cell(value);
    const iso=/^(\d{4})-(\d{2})-\d{2}/.exec(text);
    if(iso && +iso[2]>=1 && +iso[2]<=12) return iso[1]+'-'+iso[2];
    const name=/^\d{1,2}[- /]([a-z]{3})[- /](\d{2}|\d{4})$/i.exec(text);
    if(name){
      const index='jan feb mar apr may jun jul aug sep oct nov dec'.split(' ').indexOf(name[1].toLowerCase().replace('okt','oct').replace('mei','may').replace('agu','aug').replace('des','dec'));
      if(index>=0) return (name[2].length===2?'20':'')+name[2]+'-'+String(index+1).padStart(2,'0');
    }
    return '';
  }
  function downloadTemplate(type){
    const url=URL.createObjectURL(new Blob([formats[type][1].join(',')+'\r\n'],{type:'text/csv'}));
    const anchor=document.createElement('a');
    anchor.href=url; anchor.download=type+'-gpp.csv'; anchor.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  for(const [type,[title]] of Object.entries(formats)){
    const field=document.createElement('div');
    field.className='field';
    field.innerHTML='<label for="file-'+type+'">'+title+'</label><input id="file-'+type+'" type="file" accept=".xlsx,.csv" '+(type==='sales'?'multiple':'')+'><button type="button">Unduh CSV</button><p class="note" role="status"></p>';
    byId('imports').append(field);
    field.querySelector('button').onclick=()=>downloadTemplate(type);
    field.querySelector('input').onchange=event=>loadFiles(type,[...event.target.files],field.querySelector('p'));
  }
  function workbookRows(book,required){
    for(const name of book.SheetNames){
      const rows=XLSX.utils.sheet_to_json(book.Sheets[name],{header:1,defval:'',raw:true});
      const index=rows.slice(0,25).findIndex(row=>required.every(column=>row.map(key).includes(column)));
      if(index<0) continue;
      const header=rows[index].map(key);
      return rows.slice(index+1).map((row,i)=>Object.assign(Object.fromEntries(header.map((column,j)=>[column,row[j]??''])),{sheetRow:i+index+2}))
        .filter(row=>cell(row.custid)||cell(row.itemid));
    }
    throw Error('Kolom wajib tidak ditemukan. Cek header file CSV.');
  }
  function embeddedImages(book,rows){
    const files=book.files ?? {};
    function xml(path){
      const contents=files[path]?.content;
      if(!contents) return null;
      return new DOMParser().parseFromString(typeof contents==='string'?contents:new TextDecoder().decode(contents),'application/xml');
    }
    const children=(node,name)=>[...(node?.getElementsByTagNameNS('*',name)??[])];
    const images=[];
    const rowIds=new Map(rows.map(row=>[row.sheetRow,code(row.itemid)]));
    const drawingNames=Object.keys(files).filter(path=>/^xl\/drawings\/drawing\d+\.xml$/.test(path));
    for(const path of drawingNames){
      const rels=xml(path.replace('/drawings/','/drawings/_rels/')+'.rels');
      const targets=new Map(children(rels,'Relationship').map(node=>[node.getAttribute('Id'),node.getAttribute('Target')]));
      const drawing=xml(path);
      for(const anchor of [...children(drawing,'oneCellAnchor'),...children(drawing,'twoCellAnchor')]){
        const from=children(anchor,'from')[0];
        const row=Number(children(from,'row')[0]?.textContent)+1;
        const id=rowIds.get(row);
        if(!id) continue;
        const blip=children(anchor,'blip')[0];
        const relation=blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed');
        const target=targets.get(relation);
        if(!target) continue;
        const name=target.split('/').pop();
        const match=/\.(png|jpe?g|webp|gif)$/i.exec(name);
        const image=files['xl/media/'+name]?.content;
        if(match && image) images.push({id,blob:new Blob([image],{type:'image/'+(match[1].toLowerCase().startsWith('jp')?'jpeg':match[1].toLowerCase())})});
      }
    }
    return images;
  }
  async function loadFiles(type,files,status){
    if(!files.length) return;
    const version=filesVersion[type]=(filesVersion[type]??0)+1;
    status.textContent='Membaca file…';
    try{
      if(!globalThis.XLSX) throw Error('Pembaca Excel belum termuat. Muat ulang halaman dan coba lagi.');
      const rows=[], images=[];
      for(const file of files){
        if(file.size>20*1024*1024) throw Error('Maksimal 20 MB per file.');
        const workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true,bookFiles:type==='items',raw:/\.csv$/i.test(file.name)});
        const parsed=workbookRows(workbook,formats[type][2]);
        rows.push(...parsed);
        if(type==='items') images.push(...embeddedImages(workbook,parsed));
      }
      if(version!==filesVersion[type]) return;
      if(type==='sales'){
        const invalid=rows.filter(row=>!monthOf(row.tanggal)||!Number.isFinite(number(row.subamount))||!cell(row.invoice)||!code(row.custid)||!code(row.itemid));
        if(invalid.length) throw Error(invalid.length+' baris penjualan punya tanggal, omzet, atau kode yang belum valid.');
      }
      data[type]=rows;
      if(type==='items'){
        photoUrls.forEach(photo=>URL.revokeObjectURL(photo.url));
        photoUrls=images.map(photo=>({id:photo.id,url:URL.createObjectURL(photo.blob)}));
      }
      optimized=false;
      status.textContent=rows.length+' baris dimuat'+(type==='items'?' · '+images.length+' foto SKU':'')+'.';
      render();
    }catch(error){
      if(version===filesVersion[type]) status.textContent=error.message;
    }
  }
  function options(id,values,all=''){
    const select=byId(id), previous=select.value;
    if(!values.length){
      const placeholders={schedule:'Belum ada jadwal',shop:'Unggah Master Toko',division:'Unggah Data Penjualan',month:'Belum ada bulan','photo-shop':'Unggah Master Toko'};
      select.replaceChildren(new Option(placeholders[id]??'Belum ada pilihan',''));
      select.disabled=true;
      return;
    }
    select.disabled=false;
    select.replaceChildren(...(all?[new Option(all,'')]:[]),...values.map(([value,label])=>new Option(label,value)));
    if([...select.options].some(option=>option.value===previous)) select.value=previous;
  }
  const shops = ()=>new Map(data.stores.map(row=>[code(row.custid),row]));
  const items = ()=>new Map(data.items.map(row=>[code(row.itemid),row]));
  const coords = row=>{
    const lat=number(row?.latitude), lon=number(row?.longitude);
    return Number.isFinite(lat)&&Number.isFinite(lon)&&Math.abs(lat)<=90&&Math.abs(lon)<=180?[lat,lon]:null;
  };
  function distance(a,b){
    const radians=v=>v*Math.PI/180;
    const [lat1,lon1]=coords(a),[lat2,lon2]=coords(b);
    const h=Math.sin(radians(lat2-lat1)/2)**2+Math.cos(radians(lat1))*Math.cos(radians(lat2))*Math.sin(radians(lon2-lon1)/2)**2;
    return 12742*Math.asin(Math.min(1,Math.sqrt(h)));
  }
  function drawMap(id,rows,display){
    const root=byId(id);
    root.replaceChildren();
    const points=rows.map(row=>({row,point:coords(row)})).filter(item=>item.point);
    if(!points.length) return;
    const axes=[0,1].map(index=>points.map(item=>item.point[index]));
    const min=axes.map(axis=>Math.min(...axis)), max=axes.map(axis=>Math.max(...axis));
    const fraction=(value,index)=>max[index]===min[index]?50:10+80*(value-min[index])/(max[index]-min[index]);
    points.forEach(({row,point},index)=>{
      const pin=document.createElement('button');
      pin.type='button'; pin.textContent=index+1;
      pin.title=cell(row.custname); pin.setAttribute('aria-label',cell(row.custname));
      pin.style.top=(100-fraction(point[0],0))+'%';
      pin.style.left=fraction(point[1],1)+'%';
      pin.onclick=()=>display(row);
      root.append(pin);
    });
  }
  function route(){
    const storeMap=shops(), schedule=byId('schedule').value;
    const rows=data.plan.filter(row=>!schedule||cell(row.jadwalkunjungan)===schedule);
    const valid=rows.filter(row=>coords(storeMap.get(code(row.custid))));
    const missing=rows.length-valid.length;
    const store=row=>storeMap.get(code(row.custid));
    const total=list=>list.reduce((sum,row,i)=>i?sum+distance(store(list[i-1]),store(row)):sum,0);
    baseDistance=total(valid);
    let ordered=valid;
    if(optimized && valid.length>2){
      const rest=valid.slice(1), trial=[valid[0]];
      while(rest.length){
        const last=store(trial.at(-1));
        let nearest=0;
        for(let i=1;i<rest.length;i++) if(distance(last,store(rest[i]))<distance(last,store(rest[nearest]))) nearest=i;
        trial.push(rest.splice(nearest,1)[0]);
      }
      if(total(trial)<baseDistance) ordered=trial;
    }
    byId('route-info').textContent=rows.length?valid.length+' kunjungan berkoordinat · '+km(total(ordered))+(optimized?' (awal '+km(baseDistance)+')':'')+(missing?' · '+missing+' kunjungan tanpa koordinat master':''):'Unggah call plan dan master toko.';
    byId('route').replaceChildren(...[...ordered,...rows.filter(row=>!valid.includes(row))].map(row=>{
      const li=document.createElement('li'), shop=store(row);
      li.textContent=(cell(shop?.custname)||cell(row.namatoko)||code(row.custid))+' · '+code(row.custid)+(coords(shop)?'':' · koordinat belum tersedia');
      return li;
    }));
    drawMap('map',ordered.map(store),row=>{byId('store-info').textContent=cell(row.custname)+' · '+cell(row.village)+' · '+coords(row).join(', ');});
  }
  function chart(id,entries,format){
    const root=byId(id);
    if(!root) return;
    if(id!=='coverage') root.closest('.tile').hidden=!entries.length;
    const max=Math.max(1,...entries.map(([,value])=>Math.abs(value)));
    root.replaceChildren(...entries.map(([label,value])=>{
      const li=document.createElement('li'),bar=document.createElement('span');
      li.textContent=label+' · '+format(value);
      bar.className='bar';bar.style.width=(Math.abs(value)/max*100)+'%';
      li.append(bar);
      return li;
    }));
  }
  function analysis(){
    const division=byId('division').value, month=byId('month').value, shop=byId('shop').value;
    const storeMap=shops(), itemMap=items();
    const divisionRows=data.sales.filter(row=>!division||(cell(row.divname)||cell(row.divid))===division);
    const current=divisionRows.filter(row=>monthOf(row.tanggal)===month);
    byId('sales-results').hidden=!current.length;
    const sales=current.filter(row=>code(row.custid)===shop);
    const prior=divisionRows.filter(row=>monthOf(row.tanggal)<month&&code(row.custid)===shop);
    const observed=new Set(divisionRows.filter(row=>monthOf(row.tanggal)<=month&&code(row.custid)===shop).map(row=>code(row.itemid)));
    const dated=new Set(divisionRows.map(row=>monthOf(row.tanggal)));
    const sold=new Set(sales.filter(row=>number(row.qty)>0||(!cell(row.qty)&&number(row.subamount)>0)).map(row=>code(row.itemid)));
    const previous=new Set(prior.filter(row=>number(row.qty)>0||(!cell(row.qty)&&number(row.subamount)>0)).map(row=>code(row.itemid)));
    const repeat=[...previous].filter(id=>!sold.has(id)).length;
    const invoiceCount=new Set(sales.map(row=>cell(row.invoice))).size;
    byId('sales-info').textContent=!data.sales.length?'Unggah data penjualan untuk mulai.':
      (sales.length?invoiceCount+' invoice unik untuk toko ini.':'Tidak ada baris untuk toko dan bulan yang dipilih.')+
      (prior.length?' '+repeat+' SKU dari bulan sebelumnya belum terlihat dipesan di bulan ini.':'')+
      (dated.size<3?' Riwayat tersedia untuk '+dated.size+' bulan; untuk analisis tiga bulan, unggah periode lengkap.':'');
    function omzet(rows){
      const sums=new Map();
      for(const row of rows){
        const id=code(row.itemid), entry=sums.get(id)??[cell(itemMap.get(id)?.itemname)||cell(row.itemname)||id,0,id];
        entry[1]+=number(row.subamount);
        sums.set(id,entry);
      }
      return [...sums.values()].sort((a,b)=>b[1]-a[1]);
    }
    const elapsedMonths=new Set(divisionRows.filter(row=>monthOf(row.tanggal)<=month).map(row=>monthOf(row.tanggal))).size;
    const purchases=new Map();
    for(const row of sales){
      const id=code(row.itemid);
      if(!purchases.has(id)) purchases.set(id,new Set());
      purchases.get(id).add(cell(row.invoice));
    }
    chart('products',omzet(sales),value=>rupiah.format(value));
    chart('product-average',omzet([...prior,...sales]).map(([name,total])=>[name,total/Math.max(1,elapsedMonths)]),value=>rupiah.format(value));
    chart('product-invoices',omzet(sales).map(([name,,id])=>[name,purchases.get(id).size]),String);
    chart('suggestions',observed.size?omzet(current).filter(row=>row[1]>0).slice(0,10).filter(row=>!observed.has(row[2])):[],value=>rupiah.format(value));
    const active=new Set(current.filter(row=>number(row.qty)>0||(!cell(row.qty)&&number(row.subamount)>0)).map(row=>code(row.custid)));
    const mapped=[...active].map(id=>storeMap.get(id)).filter(row=>cell(row?.village));
    const areas=new Map();
    mapped.forEach(row=>{const area=cell(row.village);areas.set(area,(areas.get(area)??0)+1);});
    byId('coverage-info').textContent=current.length?mapped.length+' toko aktif terpetakan · '+(active.size-mapped.length)+' belum punya kelurahan yang cocok.':'Belum ada data penjualan untuk pilihan ini.';
    chart('coverage',[...areas].sort((a,b)=>b[1]-a[1]),String);
    byId('coverage-map').hidden=!mapped.some(row=>coords(row));
    drawMap('coverage-map',mapped,row=>{byId('coverage-info').textContent=cell(row.custname)+' · '+cell(row.village);});
  }
  function references(){
    const shop=byId('photo-shop').value, master=items();
    const history=new Map(data.sales.filter(row=>code(row.custid)===shop&&(number(row.qty)>0||(!cell(row.qty)&&number(row.subamount)>0))).map(row=>[code(row.itemid),cell(row.itemname)]));
    const labels=byId('checklist'), images=byId('images');
    labels.replaceChildren();images.replaceChildren();
    byId('checklist-note').hidden=!history.size;
    if(!photoUrls.length){
      const message=document.createElement('p');
      message.className='note';
      message.textContent=data.items.length?'Belum ada foto di Master Item.':'Unggah Master Item untuk melihat foto SKU.';
      images.append(message);
    }
    function result(){
      const missing=[...history].filter(([id])=>!checked.has(id)).map(([id,name])=>cell(master.get(id)?.itemname)||name||id);
      byId('missing').textContent=!history.size?'Belum ada riwayat yang cocok untuk toko ini.':missing.length?'Belum ditandai terlihat: '+missing.join(', ')+'. Konfirmasi stok di toko.':'Semua SKU dalam riwayat ditandai terlihat.';
    }
    for(const [id,name] of history){
      const input=document.createElement('input'), label=document.createElement('label');
      input.type='checkbox';input.checked=checked.has(id);
      input.onchange=()=>{if(input.checked) checked.add(id);else checked.delete(id);result();};
      label.style.display='block';
      label.append(input,' '+(cell(master.get(id)?.itemname)||name||id));
      labels.append(label);
    }
    for(const {id,url} of photoUrls){
      const figure=document.createElement('figure'),img=new Image(),caption=document.createElement('figcaption');
      img.src=url;img.alt='Foto SKU '+id;
      caption.textContent=cell(master.get(id)?.itemname)||id;
      figure.append(img,caption);images.append(figure);
    }
    result();
  }
  function render(){
    const storeMap=shops();
    const shopList=[...storeMap].map(([id,row])=>[id,cell(row.custname)||id]);
    options('schedule',[...new Set(data.plan.map(row=>cell(row.jadwalkunjungan)).filter(Boolean))].map(value=>[value,value]),'Semua jadwal');
    options('shop',shopList);options('photo-shop',shopList);
    options('division',[...new Set(data.sales.map(row=>cell(row.divname)||cell(row.divid)).filter(Boolean))].map(value=>[value,value]),'Semua divisi');
    options('month',[...new Set(data.sales.map(row=>monthOf(row.tanggal)).filter(Boolean))].sort().reverse().map(value=>[value,new Intl.DateTimeFormat('id-ID',{month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(value+'-01T00:00:00Z'))]));
    const callIds=new Set(data.plan.map(row=>code(row.custid))), salesIds=new Set(data.sales.map(row=>code(row.custid)));
    const count=ids=>[...ids].filter(id=>storeMap.has(id)).length;
    byId('matches').textContent=storeMap.size?'Kode toko yang cocok dengan master: call plan '+count(callIds)+'/'+callIds.size+' · penjualan '+count(salesIds)+'/'+salesIds.size+'.':'';
    route();analysis();references();
  }
  byId('schedule').onchange=()=>{optimized=false;route();};
  byId('optimize').onclick=()=>{optimized=true;route();};
  byId('reset-route').onclick=()=>{optimized=false;route();};
  for(const id of ['shop','division','month'])byId(id).onchange=analysis;
  byId('photo-shop').onchange=()=>{checked.clear();references();};
  byId('pick-display').onclick=()=>byId('display').click();
  byId('display').onchange=event=>{
    const file=event.target.files[0];
    if(!file||!file.type.startsWith('image/')) return;
    if(displayUrl) URL.revokeObjectURL(displayUrl);
    displayUrl=URL.createObjectURL(file);
    byId('display-name').textContent=file.name;
    const img=new Image();img.src=displayUrl;img.alt='Foto rak toko';
    byId('display-preview').replaceChildren(img);
  };
  render();
}
