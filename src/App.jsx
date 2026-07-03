import { useState, useRef, useEffect } from "react";

// ── CONFIG ───────────────────────────────────────────────────────────────
// Vercel: agregar variable de entorno VITE_GAS_URL con la URL del deployment
const GAS_URL = import.meta.env.VITE_GAS_URL || '';

// ── TOKENS ───────────────────────────────────────────────────────────────
const C = {
  cellar:'#0D0B09', barrel:'#1E1914', cork:'#2E2820', border:'#3A322A',
  gold:'#C4A84F', goldDim:'#7A6630', wine:'#8C2030', wineBg:'#1A0A0D',
  text:'#F0EDE6', muted:'#9B9087', dim:'#5A5248', green:'#4A8C5C', greenBg:'#0A1A0F',
};

// ── DATOS ────────────────────────────────────────────────────────────────
const USUARIOS = [
  { id:'mati',   nombre:'Mati',   emoji:'🔪', rol:'Socio' },
  { id:'andres', nombre:'Andrés', emoji:'⛷️', rol:'Socio' },
  { id:'lucas',  nombre:'Lucas',  emoji:'⛷️', rol:'Socio' },
  { id:'santi',  nombre:'Santi',  emoji:'🏠', rol:'Socio' },
];
const PRODUCTOS = [
  { id:'mb21', label:'Malbec 2021',     stock:66,   max:2964, hex:'#6B2030' },
  { id:'mb22', label:'Malbec 2022',     stock:3138, max:3818, hex:'#9B2335' },
  { id:'mb23', label:'Malbec 2023',     stock:3960, max:3960, hex:'#C04050' },
  { id:'cf21', label:'Cab. Franc 2021', stock:0,    max:972,  hex:'#4A1528' },
  { id:'cf22', label:'Cab. Franc 2022', stock:1158, max:2834, hex:'#7B2040' },
  { id:'ch22', label:'Chardonnay 2022', stock:0,    max:2208, hex:'#C4A84F' },
];
const CLIENTES = ['Angela San Rafael','Bahia Blanca','Carlos De Aquín','Adriana Laos','Chacho Andia','Mosto Divino','Organyca','Particular','Rosario','Santiago MDQ','Nuevo cliente…'];
const CAJAS    = ['Empresa (Ludico)','Mati','Lucas','Pipi (Andrés)','Santi'];
const CANALES  = ['Vinoteca','Distribuidor','Venta directa','Exportación','Muestra','Otro'];

const SEED_OPS = [
  { id:1, icon:'🍾', desc:'42 bot Malbec 2022 → Organyca',     monto:'$336.000 ARS', fecha:'04/03/2026', user:'Mati' },
  { id:2, icon:'🍾', desc:'18 bot Malbec 2021 → Organyca',     monto:'$144.000 ARS', fecha:'04/03/2026', user:'Mati' },
  { id:3, icon:'🍾', desc:'30 bot Cab. Franc 2022 → Organyca', monto:'$240.000 ARS', fecha:'06/12/2025', user:'Mati' },
  { id:4, icon:'🍾', desc:'30 bot Malbec 2021 → Rosario',      monto:'$240.000 ARS', fecha:'06/12/2025', user:'Mati' },
];

// ── GAS HELPER ───────────────────────────────────────────────────────────
async function gasGet(params) {
  if (!GAS_URL) return null;
  const url = GAS_URL + '?' + new URLSearchParams(params).toString();
  const r   = await fetch(url, { redirect: 'follow' });
  return r.json();
}

// ── HOOKS ────────────────────────────────────────────────────────────────
function useDolarBlue() {
  const [tc,setTc]=useState(null);const [err,setErr]=useState(false);const [date,setDate]=useState(null);
  useEffect(()=>{
    fetch('https://api.bluelytics.com.ar/v2/latest').then(r=>r.json()).then(d=>{
      const v=d?.blue?.value_sell;if(v){setTc(v);setDate(new Date(d.last_update||Date.now()).toLocaleDateString('es-AR'));}else setErr(true);
    }).catch(()=>setErr(true));
  },[]);
  return {tc,err,date};
}

function useAppData() {
  const [price,  setPrice]  = useState(null);
  const [ops,    setOps]    = useState(SEED_OPS);
  const [source, setSource] = useState('loading');

  const refresh = async () => {
    if (!GAS_URL) {
      // Sin GAS: usar storage local
      setPrice(8000); setSource('local');
      try { const r=await window.storage.get('ops'); if(r) setOps(JSON.parse(r.value)); } catch {}
      return;
    }
    try {
      const d = await gasGet({ action: 'getData' });
      if (d?.ok) {
        if (d.lastPrice) { setPrice(d.lastPrice); setSource('sheets'); }
        else             { setPrice(8000);         setSource('fallback'); }
        if (d.ops?.length) setOps(d.ops);
      }
    } catch { setPrice(8000); setSource('fallback'); }
  };

  useEffect(() => { refresh(); }, []);

  const addOp = async (op) => {
    const next = [op, ...ops].slice(0, 20);
    setOps(next);
    try { await window.storage.set('ops', JSON.stringify(next)); } catch {}
  };

  return { price, source, ops, addOp, refresh };
}

function useStockControl() {
  const [last,setLast]=useState(null);
  useEffect(()=>{ (async()=>{try{const r=await window.storage.get('stock_ctrl');if(r)setLast(JSON.parse(r.value));}catch{}})(); },[]);
  const save=async(ctrl)=>{ setLast(ctrl); try{await window.storage.set('stock_ctrl',JSON.stringify(ctrl));}catch{}; };
  return {last,save};
}

async function getPin(uid){try{const r=await window.storage.get(`pin_${uid}`);return r?.value||'1234';}catch{return'1234';}}
async function savePin(uid,pin){try{await window.storage.set(`pin_${uid}`,pin);return true;}catch{return false;}}

// ── CONTEXTO IA ──────────────────────────────────────────────────────────
function buildCtx(tc, price) {
  const total=PRODUCTOS.reduce((s,p)=>s+p.stock,0);
  const ars=price?total*price:null; const usd=ars&&tc?Math.round(ars/tc):null;
  return `Sos el asistente de Novato, bodega boutique de Mendoza. Datos al ${new Date().toLocaleDateString('es-AR')}.
STOCK: Malbec 2021:66 | Malbec 2022:3138 | Malbec 2023:3960 (sin ventas) | CF 2021:agotado | CF 2022:1158 | Chardonnay 2022:agotado. Total:${total} bot. Precio ref.:$${price||'N/A'}/bot. Valorización:${ars?'ARS $'+ars.toLocaleString('es-AR'):'—'}${usd?' ≈ USD '+usd.toLocaleString('es-AR'):''}.  TC blue:${tc?'$'+tc.toLocaleString('es-AR')+'/USD':'N/A'}.
CLIENTES ACTIVOS: Organyca (última 04/03/2026, en cierre). INACTIVOS: Angela San Rafael, Carlos De Aquín, Bahia Blanca, Rosario, Gahvino (perdido 2023, distribuidor BA), Mercadito Chacras (vinoteca, 23 op, inactivo desde 2023).
FINANZAS: Cajas socios ~$0 ARS | Empresa ≈ USD 275. CONTEXTO: 4 socios, sin clientes activos, relanzamiento comercial en enero con vuelta de Mati. Uva Valle de Uco, elaboración Otero Ramos.
Respondé directo y útil. Abreviaciones: $, ARS, USD, bot, op.`;
}

// ── UI BASE ──────────────────────────────────────────────────────────────
const F=({label,children})=>(<div style={{display:'flex',flexDirection:'column',gap:6}}>{label&&<label style={{fontSize:11,color:C.muted,letterSpacing:'0.1em',textTransform:'uppercase',fontFamily:'system-ui'}}>{label}</label>}{children}</div>);
const iS={background:C.cork,border:`1px solid ${C.border}`,borderRadius:10,padding:'12px 14px',color:C.text,fontSize:15,fontFamily:'system-ui',outline:'none',width:'100%',boxSizing:'border-box'};
const Inp=p=><input style={iS} {...p}/>;
const Sel=({children,...p})=><select style={{...iS,cursor:'pointer'}} {...p}>{children}</select>;
const Btn=({children,variant='primary',style:s,...p})=>(<button style={{background:variant==='primary'?C.gold:variant==='danger'?C.wine:C.cork,color:variant==='primary'?'#1E1914':C.text,border:`1px solid ${variant==='primary'?C.gold:variant==='danger'?C.wine:C.border}`,borderRadius:10,padding:'12px 20px',fontSize:15,fontFamily:'system-ui',fontWeight:600,cursor:'pointer',...s}} {...p}>{children}</button>);
const Card=({children,style:s})=><div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:14,padding:16,...s}}>{children}</div>;
const SL=({children})=><div style={{fontSize:10,color:C.dim,letterSpacing:'0.15em',textTransform:'uppercase',fontFamily:'system-ui',marginBottom:10}}>{children}</div>;

// ── LOGIN ────────────────────────────────────────────────────────────────
function LoginScreen({onLogin}){
  const [step,setStep]=useState('users');const [pending,setPending]=useState(null);const [pin,setPin]=useState('');const [shake,setShake]=useState(false);const [hov,setHov]=useState(null);
  const tryPin=async()=>{const s=await getPin(pending.id);if(pin===s){onLogin(pending);}else{setShake(true);setPin('');setTimeout(()=>setShake(false),500);}};
  useEffect(()=>{if(pin.length===4)tryPin();},[pin]);
  if(step==='pin') return(
    <div style={{minHeight:'100vh',background:C.cellar,display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:28}}>
      <button onClick={()=>{setStep('users');setPin('');}} style={{position:'absolute',top:20,left:20,background:'none',border:'none',color:C.muted,fontSize:22,cursor:'pointer'}}>←</button>
      <div style={{textAlign:'center',marginBottom:36}}><div style={{fontSize:44,marginBottom:10}}>{pending.emoji}</div><div style={{color:C.text,fontSize:20,fontFamily:'Georgia, serif',fontWeight:700}}>{pending.nombre}</div><div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',marginTop:4}}>Ingresá tu PIN</div></div>
      <div style={{display:'flex',gap:16,marginBottom:32,animation:shake?'shake 0.4s ease':'none'}}>
        {[0,1,2,3].map(i=><div key={i} style={{width:16,height:16,borderRadius:'50%',background:pin.length>i?C.gold:C.cork,border:`2px solid ${pin.length>i?C.gold:C.border}`,transition:'background 0.15s'}}/>)}
      </div>
      <input autoFocus type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} value={pin} onChange={e=>{const v=e.target.value.replace(/\D/g,'');if(v.length<=4)setPin(v);}} style={{position:'absolute',opacity:0,width:1,height:1}}/>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,72px)',gap:10}}>
        {['1','2','3','4','5','6','7','8','9','','0','⌫'].map((k,i)=>(
          <button key={i} onClick={()=>{if(k==='⌫')setPin(p=>p.slice(0,-1));else if(k)setPin(p=>p.length<4?p+k:p);}} style={{height:64,background:k?C.barrel:C.cellar,border:`1px solid ${k?C.border:'transparent'}`,borderRadius:12,color:C.text,fontSize:22,fontFamily:'Georgia, serif',cursor:k?'pointer':'default',opacity:k?1:0}}>{k}</button>
        ))}
      </div>
      <div style={{color:C.dim,fontSize:11,fontFamily:'system-ui',marginTop:28}}>PIN inicial: 1234</div>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}`}</style>
    </div>
  );
  return(
    <div style={{minHeight:'100vh',background:C.cellar,display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:28}}>
      <div style={{textAlign:'center',marginBottom:40}}>
        <div style={{color:C.gold,fontSize:38,fontWeight:700,letterSpacing:'0.3em',fontFamily:'Georgia, serif'}}>NOVATO</div>
        <div style={{color:C.dim,fontSize:11,letterSpacing:'0.2em',fontFamily:'system-ui',marginTop:4}}>MENDOZA · ARGENTINA</div>
        <div style={{width:48,height:1,background:`linear-gradient(90deg,transparent,${C.gold},transparent)`,margin:'18px auto'}}/>
        <div style={{color:C.muted,fontSize:13,fontFamily:'system-ui'}}>¿Quién está operando?</div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,width:'100%',maxWidth:340}}>
        {USUARIOS.map(u=><button key={u.id} onClick={()=>{setPending(u);setPin('');setStep('pin');}} onMouseEnter={()=>setHov(u.id)} onMouseLeave={()=>setHov(null)} style={{background:C.barrel,border:`1px solid ${hov===u.id?C.gold:C.border}`,borderRadius:16,padding:'24px 16px',cursor:'pointer',textAlign:'center',transition:'border-color 0.2s'}}><div style={{fontSize:30,marginBottom:10}}>{u.emoji}</div><div style={{color:C.text,fontWeight:700,fontSize:16,fontFamily:'Georgia, serif',marginBottom:4}}>{u.nombre}</div><div style={{color:C.dim,fontSize:11,fontFamily:'system-ui'}}>{u.rol}</div></button>)}
      </div>
      {!GAS_URL&&<div style={{marginTop:24,background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:12,color:'#E07080',maxWidth:340,textAlign:'center'}}>⚠ Sin conexión a Google Sheets — modo offline</div>}
    </div>
  );
}

// ── SETTINGS ─────────────────────────────────────────────────────────────
function SettingsPanel({user,onClose,onLogout,showToast}){
  const [step,setStep]=useState('menu');const [np,setNp]=useState('');const [cp,setCp]=useState('');const [saving,setSaving]=useState(false);
  const changePin=async()=>{if(np.length!==4){showToast('El PIN debe ser de 4 dígitos','error');return;}if(np!==cp){showToast('Los PINs no coinciden','error');return;}setSaving(true);const ok=await savePin(user.id,np);setSaving(false);if(ok){showToast('✓ PIN actualizado');setStep('menu');setNp('');setCp('');}else showToast('Error al guardar','error');};
  return(<>
    <div onClick={onClose} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.6)',zIndex:300}}/>
    <div style={{position:'fixed',bottom:0,left:0,right:0,background:C.barrel,border:`1px solid ${C.border}`,borderRadius:'20px 20px 0 0',padding:'8px 20px 40px',zIndex:301}}>
      <div style={{width:40,height:4,background:C.border,borderRadius:2,margin:'12px auto 20px'}}/>
      <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:24,paddingBottom:20,borderBottom:`1px solid ${C.border}`}}>
        <div style={{width:52,height:52,borderRadius:'50%',background:C.cork,border:`2px solid ${C.gold}`,display:'flex',alignItems:'center',justifyContent:'center',fontSize:26}}>{user.emoji}</div>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>{user.nombre}</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>{GAS_URL?'✓ Conectado a Google Sheets':'⚠ Sin conexión a Google Sheets'}</div></div>
      </div>
      {step==='menu'?(<div style={{display:'flex',flexDirection:'column',gap:10}}><SL>Configuración</SL>
        <button onClick={()=>setStep('pin')} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,padding:'14px 16px',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',color:C.text,fontFamily:'system-ui',fontSize:14}}><span>🔒 Cambiar PIN</span><span style={{color:C.dim}}>›</span></button>
        <button onClick={onLogout} style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:12,padding:'14px 16px',cursor:'pointer',color:'#E07080',fontFamily:'system-ui',fontSize:14,textAlign:'left'}}>↩ Cerrar sesión</button>
      </div>):(<div style={{display:'flex',flexDirection:'column',gap:14}}>
        <button onClick={()=>setStep('menu')} style={{background:'none',border:'none',color:C.muted,cursor:'pointer',fontSize:14,fontFamily:'system-ui',textAlign:'left',marginBottom:4}}>← Volver</button>
        <SL>Cambiar PIN</SL>
        <F label="Nuevo PIN (4 dígitos)"><Inp type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} placeholder="••••" value={np} onChange={e=>setNp(e.target.value.replace(/\D/g,'').slice(0,4))}/></F>
        <F label="Confirmar PIN"><Inp type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} placeholder="••••" value={cp} onChange={e=>setCp(e.target.value.replace(/\D/g,'').slice(0,4))}/></F>
        <Btn onClick={changePin} style={{opacity:saving?0.6:1}}>{saving?'Guardando…':'Guardar PIN'}</Btn>
      </div>)}
    </div>
  </>);
}

// ── STOCK CONTROL ────────────────────────────────────────────────────────
function StockControlScreen({onBack,showToast,onSaved}){
  const {last,save}=useStockControl();
  const [real,setReal]=useState(Object.fromEntries(PRODUCTOS.map(p=>[p.id,''])));
  const set=(id,v)=>setReal(p=>({...p,[id]:v.replace(/\D/g,'')}));
  const diffs=PRODUCTOS.map(p=>{const r=real[p.id]===''?null:parseInt(real[p.id]);return{...p,real:r,diff:r!==null?r-p.stock:null};}).filter(p=>p.real!==null);
  const hasDiff=diffs.some(d=>d.diff!==0);const allFilled=PRODUCTOS.every(p=>real[p.id]!=='');
  const confirm=async()=>{
    const ctrl={fecha:new Date().toLocaleDateString('es-AR'),hora:new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'}),items:diffs};
    await save(ctrl);onSaved();
    if(hasDiff){showToast(`⚠ Control guardado — diferencias en ${diffs.filter(d=>d.diff!==0).length} producto(s)`,'error');}
    else{showToast('✓ Stock coincide con lo teórico','ok');}
    onBack();
  };
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:8}}>
        <button onClick={onBack} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px 12px',color:C.muted,cursor:'pointer',fontSize:16}}>←</button>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>Control de stock real</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Ingresá el conteo físico en depósito</div></div>
      </div>
      {last&&<div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:12,color:C.muted,marginBottom:16,marginTop:8}}>Último control: <strong style={{color:C.text}}>{last.fecha} a las {last.hora}</strong></div>}
      <div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:20,marginTop:last?0:16}}>
        <div style={{display:'grid',gridTemplateColumns:'1fr 80px 80px 80px',gap:8,padding:'0 4px'}}>
          {['Producto','Teórico','Real','Dif.'].map(h=><div key={h} style={{fontSize:10,color:C.dim,letterSpacing:'0.1em',textTransform:'uppercase',fontFamily:'system-ui',textAlign:h!=='Producto'?'center':'left'}}>{h}</div>)}
        </div>
        {PRODUCTOS.map(p=>{
          const r=real[p.id]===''?null:parseInt(real[p.id]);
          const diff=r!==null?r-p.stock:null;
          const dc=diff===null?C.dim:diff<0?'#f08080':diff>0?'#7dce9b':C.muted;
          return(
            <div key={p.id} style={{background:C.barrel,border:`1px solid ${diff!==null&&diff!==0?dc+'44':C.border}`,borderRadius:12,padding:'12px 14px',display:'grid',gridTemplateColumns:'1fr 80px 80px 80px',gap:8,alignItems:'center'}}>
              <div><div style={{color:C.text,fontSize:14,fontFamily:'system-ui',fontWeight:600}}>{p.label}</div><div style={{width:40,height:3,background:p.hex,borderRadius:2,marginTop:4}}/></div>
              <div style={{textAlign:'center',color:C.muted,fontSize:15,fontFamily:'system-ui'}}>{p.stock}</div>
              <input type="number" min="0" placeholder="—" value={real[p.id]} onChange={e=>set(p.id,e.target.value)} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px',color:C.text,fontSize:15,fontFamily:'system-ui',outline:'none',width:'100%',boxSizing:'border-box',textAlign:'center'}}/>
              <div style={{textAlign:'center',fontWeight:700,fontSize:15,fontFamily:'system-ui',color:dc}}>{diff===null?'—':diff===0?'✓':diff>0?`+${diff}`:diff}</div>
            </div>
          );
        })}
      </div>
      {diffs.some(d=>d.diff!==0&&d.diff!==null)&&(
        <div style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:12,padding:'14px 16px',marginBottom:16,fontFamily:'system-ui'}}>
          <div style={{color:'#E07080',fontSize:13,fontWeight:700,marginBottom:8}}>⚠ Diferencias detectadas</div>
          {diffs.filter(d=>d.diff!==0&&d.diff!==null).map(d=>(
            <div key={d.id} style={{display:'flex',justifyContent:'space-between',fontSize:13,color:C.muted,paddingBottom:6,marginBottom:6,borderBottom:`1px solid ${C.border}`}}>
              <span>{d.label}</span><span style={{color:d.diff<0?'#f08080':'#7dce9b',fontWeight:700}}>{d.diff>0?'+':''}{d.diff} bot</span>
            </div>
          ))}
          <div style={{color:'#A06070',fontSize:12,marginTop:4,lineHeight:1.5}}>Causas posibles: retiros sin avisar, cobro en especie del depósito, movimientos sin cargar.</div>
        </div>
      )}
      <Btn onClick={confirm} style={{width:'100%',opacity:allFilled?1:0.5}} disabled={!allFilled}>{allFilled?'Confirmar control de stock':'Completá todos los campos'}</Btn>
    </div>
  );
}

// ── DASHBOARD ────────────────────────────────────────────────────────────
function DashboardScreen({onNavigate,price,source,ops}){
  const {tc,err:tcErr,date:tcDate}=useDolarBlue();
  const {last}=useStockControl();
  const total=PRODUCTOS.reduce((s,p)=>s+p.stock,0);
  const ars=price?total*price:null; const usd=ars&&tc?Math.round(ars/tc):null;
  const daysSince=last?(()=>{try{const [d,m,y]=last.fecha.split('/');return Math.floor((Date.now()-new Date(`${y}-${m}-${d}`).getTime())/86400000);}catch{return null;}})():null;
  return(
    <div style={{padding:'20px 16px',maxWidth:540,margin:'0 auto'}}>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:10,marginBottom:12}}>
        {[{v:total.toLocaleString('es-AR'),u:'botellas',l:'Stock total'},{v:usd?`USD ${usd.toLocaleString('es-AR')}`:'…',u:'valorización',l:price?`$${price.toLocaleString('es-AR')}/bot`:'cargando…'},{v:'USD 275',u:'disponibles',l:'Caja empresa'}].map(k=><Card key={k.l} style={{padding:'12px 10px'}}><div style={{color:C.gold,fontSize:17,fontWeight:700,lineHeight:1,fontFamily:'Georgia, serif'}}>{k.v}</div><div style={{color:C.dim,fontSize:10,fontFamily:'system-ui',marginTop:2}}>{k.u}</div><div style={{color:C.muted,fontSize:10,fontFamily:'system-ui',marginTop:5,borderTop:`1px solid ${C.border}`,paddingTop:5}}>{k.l}</div></Card>)}
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:6,marginBottom:18}}>
        <div style={{display:'flex',alignItems:'center',gap:8,padding:'8px 12px',background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10}}>
          <div style={{width:7,height:7,borderRadius:'50%',background:tc?C.green:tcErr?C.wine:C.dim,flexShrink:0}}/>
          <span style={{fontSize:12,fontFamily:'system-ui',color:C.muted}}>{tc?<>Dólar blue: <strong style={{color:C.text}}>${tc.toLocaleString('es-AR')}</strong> <span style={{color:C.dim}}>· {tcDate}</span></>:tcErr?<span style={{color:'#f08080'}}>TC no disponible</span>:'Obteniendo cotización…'}</span>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:8,padding:'8px 12px',background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10}}>
          <div style={{width:7,height:7,borderRadius:'50%',background:source==='sheets'?C.green:C.goldDim,flexShrink:0}}/>
          <span style={{fontSize:12,fontFamily:'system-ui',color:C.muted}}>Precio ref.: <strong style={{color:C.text}}>{price?`$${price.toLocaleString('es-AR')}/bot`:'cargando…'}</strong><span style={{color:source==='sheets'?C.green:C.goldDim}}>{source==='sheets'?' · Google Sheets':' · último dato manual'}</span></span>
        </div>
      </div>
      <SL>Stock por producto</SL>
      <Card style={{marginBottom:14}}>
        <div style={{display:'flex',gap:8,alignItems:'flex-end',height:110,marginBottom:10}}>
          {PRODUCTOS.map(p=>{const pct=p.max>0?p.stock/p.max:0;const h=Math.max(Math.round(pct*90),p.stock>0?4:2);return(<div key={p.id} style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',gap:4}}><div style={{fontSize:9,color:p.stock>0?C.muted:C.dim,fontFamily:'system-ui',minHeight:13}}>{p.stock>0?p.stock.toLocaleString('es-AR'):''}</div><div style={{width:'100%',display:'flex',flexDirection:'column',alignItems:'center'}}><div style={{width:'35%',height:13,background:C.cork,borderRadius:'3px 3px 0 0',border:`1px solid ${C.border}`,borderBottom:'none'}}/><div style={{width:'100%',height:80,background:C.cork,borderRadius:'2px 2px 6px 6px',border:`1px solid ${C.border}`,overflow:'hidden',display:'flex',alignItems:'flex-end'}}><div style={{width:'100%',height:`${h}px`,background:p.stock>0?p.hex:C.cork,transition:'height 0.8s cubic-bezier(.4,0,.2,1)'}}/></div></div></div>);})}
        </div>
        <div style={{display:'flex',gap:8}}>{PRODUCTOS.map(p=><div key={p.id} style={{flex:1,textAlign:'center'}}><div style={{fontSize:8,color:p.stock>0?C.muted:C.dim,fontFamily:'system-ui',lineHeight:1.3}}>{p.label.split(' ').map((w,i)=><span key={i}>{w}<br/></span>)}</div></div>)}</div>
      </Card>
      <button onClick={()=>onNavigate('stock')} style={{width:'100%',background:daysSince===null||daysSince>30?C.wineBg:C.barrel,border:`1px solid ${daysSince===null||daysSince>30?C.wine+'55':C.border}`,borderRadius:12,padding:'13px 16px',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:18}}>
        <div style={{textAlign:'left'}}><div style={{color:daysSince===null||daysSince>30?'#E07080':C.text,fontSize:13,fontWeight:700,fontFamily:'system-ui',marginBottom:2}}>{daysSince===null?'⚠ Control de stock pendiente':daysSince>30?`⚠ Último control hace ${daysSince} días`:`✓ Control de stock · hace ${daysSince} día${daysSince===1?'':'s'}`}</div><div style={{color:daysSince===null||daysSince>30?'#A06070':C.dim,fontSize:12,fontFamily:'system-ui'}}>{last?`${last.fecha} a las ${last.hora}`:'Nunca realizado — tocá para hacer el primer control'}</div></div>
        <span style={{color:C.gold,fontSize:18}}>›</span>
      </button>
      <SL>Últimas operaciones</SL>
      <div style={{display:'flex',flexDirection:'column',gap:8}}>
        {ops.slice(0,4).map((op,i)=>(
          <div key={op.id||i} style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:12,padding:'12px 14px',display:'flex',alignItems:'center',gap:12}}>
            <span style={{fontSize:22,flexShrink:0}}>{op.icon}</span>
            <div style={{flex:1,minWidth:0}}><div style={{color:C.text,fontSize:13,fontFamily:'system-ui',fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{op.desc}</div><div style={{color:C.dim,fontSize:11,fontFamily:'system-ui',marginTop:2}}>{op.fecha} · {op.user}</div></div>
            <div style={{color:C.muted,fontSize:12,fontFamily:'system-ui',flexShrink:0,textAlign:'right'}}>{op.monto}</div>
          </div>
        ))}
        {ops.length===0&&<div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',textAlign:'center',padding:'20px 0'}}>Sin operaciones registradas</div>}
      </div>
    </div>
  );
}

// ── NUEVA VENTA ──────────────────────────────────────────────────────────
function VentaScreen({user,onBack,showToast,addOp,price}){
  const hoy=new Date().toISOString().split('T')[0];
  const [f,setF]=useState({producto:'',botellas:'',cliente:'',clienteNuevo:'',canal:'',monto:'',moneda:'ARS',fecha:hoy,notas:''});
  const [sending,setSending]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const prod=PRODUCTOS.find(p=>p.id===f.producto);
  const sug=price&&f.botellas?price*parseInt(f.botellas||0):null;
  const submit=async()=>{
    if(!f.producto||!f.botellas||!f.cliente||(f.cliente==='Nuevo cliente…'&&!f.clienteNuevo)){showToast('Completá los campos obligatorios','error');return;}
    const cl=f.cliente==='Nuevo cliente…'?f.clienteNuevo:f.cliente;
    setSending(true);
    try {
      if(GAS_URL) await gasGet({action:'addSale',fecha:f.fecha,producto:prod?.label,botellas:f.botellas,monto:f.monto||0,moneda:f.moneda,cliente:cl,canal:f.canal||'',notas:f.notas||'',user:user.nombre});
      const montoStr=f.monto?`$${parseInt(f.monto).toLocaleString('es-AR')} ${f.moneda}`:'sin monto';
      await addOp({id:Date.now(),icon:'🍾',desc:`${f.botellas} bot ${prod?.label} → ${cl}`,monto:montoStr,fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      showToast(`✓ Venta registrada${GAS_URL?' y enviada a Sheets':''}`,'ok');
      onBack();
    } catch(e){ showToast('Error al registrar. Intentá de nuevo.','error'); }
    setSending(false);
  };
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:24}}>
        <button onClick={onBack} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px 12px',color:C.muted,cursor:'pointer',fontSize:16}}>←</button>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>Nueva venta</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Registrar movimiento de stock</div></div>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <F label="Producto *"><Sel value={f.producto} onChange={e=>set('producto',e.target.value)}><option value="">Seleccionar…</option>{PRODUCTOS.filter(p=>p.stock>0).map(p=><option key={p.id} value={p.id}>{p.label} — {p.stock.toLocaleString('es-AR')} bot</option>)}</Sel></F>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
          <F label="Botellas *"><Inp type="number" min="1" max={prod?.stock} placeholder="0" value={f.botellas} onChange={e=>set('botellas',e.target.value)}/></F>
          <F label="Fecha"><Inp type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></F>
        </div>
        {sug&&<div style={{background:C.greenBg,border:`1px solid ${C.green}55`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:13}}><span style={{color:C.muted}}>Monto sugerido: </span><strong style={{color:'#7dce9b'}}>${sug.toLocaleString('es-AR')}</strong><span style={{color:C.dim}}> (${price.toLocaleString('es-AR')}/bot · última venta)</span></div>}
        <div style={{display:'grid',gridTemplateColumns:'2fr 1fr',gap:12}}>
          <F label="Monto total *"><Inp type="number" placeholder="0" value={f.monto} onChange={e=>set('monto',e.target.value)}/></F>
          <F label="Moneda"><Sel value={f.moneda} onChange={e=>set('moneda',e.target.value)}><option value="ARS">ARS $</option><option value="USD">USD $</option></Sel></F>
        </div>
        <F label="Cliente *"><Sel value={f.cliente} onChange={e=>set('cliente',e.target.value)}><option value="">Seleccionar…</option>{CLIENTES.map(c=><option key={c} value={c}>{c}</option>)}</Sel></F>
        {f.cliente==='Nuevo cliente…'&&<F label="Nombre nuevo cliente"><Inp placeholder="Nombre o comercio" value={f.clienteNuevo} onChange={e=>set('clienteNuevo',e.target.value)}/></F>}
        <F label="Canal"><Sel value={f.canal} onChange={e=>set('canal',e.target.value)}><option value="">Seleccionar…</option>{CANALES.map(c=><option key={c} value={c}>{c}</option>)}</Sel></F>
        <F label="Notas"><Inp placeholder="Referencia, condiciones de pago, etc." value={f.notas} onChange={e=>set('notas',e.target.value)}/></F>
        <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':'Registrar venta'}</Btn></div>
      </div>
    </div>
  );
}

// ── CAJA ─────────────────────────────────────────────────────────────────
function CajaScreen({user,onBack,showToast,addOp}){
  const hoy=new Date().toISOString().split('T')[0];
  const [f,setF]=useState({tipo:'cobro',monto:'',moneda:'ARS',caja:'Empresa (Ludico)',concepto:'',fecha:hoy});
  const [sending,setSending]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const submit=async()=>{
    if(!f.monto||!f.concepto){showToast('Completá los campos obligatorios','error');return;}
    setSending(true);
    try {
      if(GAS_URL) await gasGet({action:'addMovement',fecha:f.fecha,tipo:f.tipo,monto:f.monto,moneda:f.moneda,caja:f.caja,concepto:f.concepto,user:user.nombre});
      const icon=f.tipo==='cobro'?'💵':'💸';
      await addOp({id:Date.now(),icon,desc:`${f.tipo==='cobro'?'Cobro':'Gasto'}: ${f.concepto}`,monto:`${f.moneda} ${parseInt(f.monto).toLocaleString('es-AR')}`,fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      showToast(`✓ Movimiento registrado${GAS_URL?' y enviado a Sheets':''}`,'ok');
      onBack();
    } catch(e){ showToast('Error al registrar. Intentá de nuevo.','error'); }
    setSending(false);
  };
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:24}}>
        <button onClick={onBack} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px 12px',color:C.muted,cursor:'pointer',fontSize:16}}>←</button>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>Movimiento de caja</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Registrar cobro, pago o gasto</div></div>
      </div>
      <div style={{display:'flex',marginBottom:20,border:`1px solid ${C.border}`,borderRadius:10,overflow:'hidden'}}>
        {[{k:'cobro',l:'↓ Cobro / Ingreso',bg:C.greenBg,col:'#7dce9b'},{k:'gasto',l:'↑ Gasto / Egreso',bg:C.wineBg,col:'#f08080'}].map((t,i)=>(
          <button key={t.k} onClick={()=>set('tipo',t.k)} style={{flex:1,padding:'12px 8px',background:f.tipo===t.k?t.bg:C.barrel,color:f.tipo===t.k?t.col:C.dim,border:'none',borderRight:i===0?`1px solid ${C.border}`:'none',cursor:'pointer',fontSize:13,fontFamily:'system-ui',fontWeight:f.tipo===t.k?700:400}}>{t.l}</button>
        ))}
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <div style={{display:'grid',gridTemplateColumns:'2fr 1fr',gap:12}}>
          <F label="Monto *"><Inp type="number" placeholder="0" value={f.monto} onChange={e=>set('monto',e.target.value)}/></F>
          <F label="Moneda"><Sel value={f.moneda} onChange={e=>set('moneda',e.target.value)}><option value="ARS">ARS $</option><option value="USD">USD $</option></Sel></F>
        </div>
        <F label={f.tipo==='cobro'?'Caja que recibe':'Caja de la que sale'}><Sel value={f.caja} onChange={e=>set('caja',e.target.value)}>{CAJAS.map(c=><option key={c} value={c}>{c}</option>)}</Sel></F>
        <F label="Concepto *"><Inp placeholder={f.tipo==='cobro'?'Ej: Cobro factura Organyca':'Ej: Pago etiquetas'} value={f.concepto} onChange={e=>set('concepto',e.target.value)}/></F>
        <F label="Fecha"><Inp type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></F>
        <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':'Registrar movimiento'}</Btn></div>
      </div>
    </div>
  );
}

// ── CONSULTAS IA ─────────────────────────────────────────────────────────
function ConsultasScreen({price}){
  const {tc}=useDolarBlue();
  const [msgs,setMsgs]=useState([{role:'assistant',content:'Hola 👋 Soy el asistente de Novato. Preguntame sobre stock, clientes, caja o cualquier dato de la bodega.'}]);
  const [input,setInput]=useState('');const [loading,setL]=useState(false);const ref=useRef(null);
  useEffect(()=>ref.current?.scrollIntoView({behavior:'smooth'}),[msgs,loading]);
  const QUICK=['¿Cuánto stock tenemos disponible?','¿Cuánto vale el stock en dólares?','¿Qué clientes están inactivos?','Resumime el estado de la bodega'];
  const send=async(text)=>{
    const q=(text||input).trim();if(!q||loading)return;setInput('');
    const next=[...msgs,{role:'user',content:q}];setMsgs(next);setL(true);
    try{const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({model:'claude-sonnet-4-6',max_tokens:1000,system:buildCtx(tc,price),messages:next.slice(1)})});const d=await res.json();setMsgs(p=>[...p,{role:'assistant',content:d?.content?.[0]?.text||'Sin respuesta.'}]);}
    catch{setMsgs(p=>[...p,{role:'assistant',content:'Error de conexión.'}]);}
    setL(false);
  };
  return(
    <div style={{display:'flex',flexDirection:'column',height:'calc(100vh - 130px)'}}>
      <div style={{flex:1,overflowY:'auto',padding:'16px 16px 0'}}>
        {msgs.length===1&&<div style={{marginBottom:20}}><SL>Preguntas frecuentes</SL><div style={{display:'flex',flexDirection:'column',gap:8}}>{QUICK.map(q=><button key={q} onClick={()=>send(q)} style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10,padding:'11px 14px',color:C.muted,fontSize:13,fontFamily:'system-ui',textAlign:'left',cursor:'pointer'}}>{q}</button>)}</div></div>}
        {msgs.map((m,i)=>(<div key={i} style={{marginBottom:14,display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start',alignItems:'flex-start',gap:8}}>{m.role==='assistant'&&<div style={{width:28,height:28,borderRadius:'50%',background:C.gold,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,fontWeight:700,color:'#1E1914',flexShrink:0,marginTop:2}}>N</div>}<div style={{maxWidth:'78%',background:m.role==='user'?C.wine:C.barrel,border:m.role==='assistant'?`1px solid ${C.border}`:'none',borderRadius:m.role==='user'?'16px 4px 16px 16px':'4px 16px 16px 16px',padding:'11px 14px',fontSize:14,fontFamily:'system-ui',color:C.text,lineHeight:1.6,whiteSpace:'pre-wrap'}}>{m.content}</div></div>))}
        {loading&&<div style={{display:'flex',gap:8,marginBottom:14}}><div style={{width:28,height:28,borderRadius:'50%',background:C.gold,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,fontWeight:700,color:'#1E1914',flexShrink:0}}>N</div><div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:'4px 16px 16px 16px',padding:'14px 18px',display:'flex',gap:5,alignItems:'center'}}>{[0,1,2].map(d=><div key={d} style={{width:6,height:6,borderRadius:'50%',background:C.gold,animation:`dot 1.2s ${d*0.2}s infinite`}}/>)}</div></div>}
        <div ref={ref}/>
      </div>
      <div style={{padding:'10px 16px 14px',borderTop:`1px solid ${C.border}`,background:C.barrel,display:'flex',gap:10,alignItems:'center'}}>
        <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}}} placeholder="Preguntá sobre stock, clientes, caja…" style={{flex:1,background:C.cork,border:`1px solid ${C.border}`,borderRadius:24,padding:'12px 18px',color:C.text,fontSize:14,fontFamily:'system-ui',outline:'none'}}/>
        <button onClick={()=>send()} disabled={loading||!input.trim()} style={{background:input.trim()&&!loading?C.gold:C.cork,border:'none',borderRadius:'50%',width:44,height:44,display:'flex',alignItems:'center',justifyContent:'center',cursor:input.trim()?'pointer':'default',fontSize:18,color:input.trim()?'#1E1914':C.dim,flexShrink:0}}>↑</button>
      </div>
      <style>{`@keyframes dot{0%,80%,100%{opacity:.3;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}`}</style>
    </div>
  );
}

// ── NAV ──────────────────────────────────────────────────────────────────
function Nav({screen,setScreen}){
  const tabs=[{id:'dashboard',icon:'⌂',l:'Inicio'},{id:'venta',icon:'🍾',l:'Venta'},{id:'caja',icon:'⇄',l:'Caja'},{id:'stock',icon:'📦',l:'Stock'},{id:'consultas',icon:'✦',l:'Asistente'}];
  return(<div style={{position:'fixed',bottom:0,left:0,right:0,background:C.barrel,borderTop:`1px solid ${C.border}`,display:'flex',zIndex:200,paddingBottom:'env(safe-area-inset-bottom, 0px)'}}>{tabs.map(t=><button key={t.id} onClick={()=>setScreen(t.id)} style={{flex:1,padding:'10px 2px 12px',background:'none',border:'none',cursor:'pointer',display:'flex',flexDirection:'column',alignItems:'center',gap:2}}><span style={{fontSize:t.id==='dashboard'?20:17,opacity:screen===t.id?1:0.35}}>{t.icon}</span><span style={{fontSize:9,fontFamily:'system-ui',color:screen===t.id?C.gold:C.dim,fontWeight:screen===t.id?700:400}}>{t.l}</span></button>)}</div>);
}

// ── APP ───────────────────────────────────────────────────────────────────
export default function NovatoApp(){
  const [user,setUser]=useState(null);const [screen,setScreen]=useState('dashboard');const [toast,setToast]=useState(null);const [settings,setSettings]=useState(false);const [ctrlKey,setCtrlKey]=useState(0);
  const {price,source,ops,addOp}=useAppData();
  const showToast=(msg,type='ok')=>{setToast({msg,type});setTimeout(()=>setToast(null),3500);};
  const logout=()=>{setUser(null);setSettings(false);setScreen('dashboard');};
  if(!user)return <LoginScreen onLogin={u=>{setUser(u);setScreen('dashboard');}}/>;
  return(
    <div style={{fontFamily:'Georgia, serif',background:C.cellar,minHeight:'100vh',color:C.text}}>
      {toast&&<div style={{position:'fixed',top:16,left:'50%',transform:'translateX(-50%)',zIndex:999,whiteSpace:'nowrap',background:toast.type==='ok'?C.greenBg:C.wineBg,border:`1px solid ${toast.type==='ok'?C.green:C.wine}`,color:toast.type==='ok'?'#7dce9b':'#f08080',padding:'11px 22px',borderRadius:10,fontSize:13,fontFamily:'system-ui',boxShadow:'0 4px 20px rgba(0,0,0,0.5)'}}>{toast.msg}</div>}
      {settings&&<SettingsPanel user={user} onClose={()=>setSettings(false)} onLogout={logout} showToast={showToast}/>}
      <div style={{background:C.barrel,borderBottom:`1px solid ${C.border}`,padding:'env(safe-area-inset-top, 13px) 18px 13px',display:'flex',alignItems:'center',justifyContent:'space-between',position:'sticky',top:0,zIndex:100}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}><span style={{color:C.gold,fontWeight:700,letterSpacing:'0.25em',fontSize:15,fontFamily:'Georgia, serif'}}>NOVATO</span><span style={{color:C.border}}>|</span><span style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Gestión</span></div>
        <button onClick={()=>setSettings(true)} style={{display:'flex',alignItems:'center',gap:7,background:C.cork,border:`1px solid ${C.border}`,borderRadius:20,padding:'6px 13px',color:C.muted,fontSize:12,fontFamily:'system-ui',cursor:'pointer'}}><span>{user.emoji}</span><span>{user.nombre}</span><span style={{color:C.dim,fontSize:10}}>⚙</span></button>
      </div>
      <div style={{paddingBottom:80}}>
        {screen==='dashboard'&&<DashboardScreen onNavigate={setScreen} price={price} source={source} ops={ops}/>}
        {screen==='venta'&&<VentaScreen user={user} onBack={()=>setScreen('dashboard')} showToast={showToast} addOp={addOp} price={price}/>}
        {screen==='caja'&&<CajaScreen user={user} onBack={()=>setScreen('dashboard')} showToast={showToast} addOp={addOp}/>}
        {screen==='stock'&&<StockControlScreen onBack={()=>setScreen('dashboard')} showToast={showToast} onSaved={()=>setCtrlKey(k=>k+1)}/>}
        {screen==='consultas'&&<ConsultasScreen price={price}/>}
      </div>
      <Nav screen={screen} setScreen={setScreen}/>
    </div>
  );
}
