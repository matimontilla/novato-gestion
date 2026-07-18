import { useState, useRef, useEffect } from "react";

// ── CONFIG ───────────────────────────────────────────────────────────────
// Vercel: agregar variable de entorno VITE_GAS_URL con la URL del deployment
const GAS_URL       = import.meta.env.VITE_GAS_URL       || '';
const ANTHROPIC_KEY = import.meta.env.VITE_ANTHROPIC_KEY || '';

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
// Depósitos entre los que se puede mover stock sin que medie una venta
const UBICACIONES = ['R Peña','Pipi','Lucas','Santi','Mati'];
const vacio = () => Object.fromEntries(UBICACIONES.map(u=>[u,0]));

// Stock inicial por producto y depósito (según último conteo en General_Dinamico → hoja STOCK)
const PRODUCTOS_SEED = [
  { id:'mb21', label:'Malbec 2021',     max:2964, hex:'#6B2030', stockUbic:{...vacio()} },
  { id:'mb22', label:'Malbec 2022',     max:3818, hex:'#9B2335', stockUbic:{...vacio(),'R Peña':2964,'Pipi':30,'Lucas':15} },
  { id:'mb23', label:'Malbec 2023',     max:3960, hex:'#C04050', stockUbic:{...vacio(),'R Peña':3960} },
  { id:'cf21', label:'Cab. Franc 2021', max:972,  hex:'#4A1528', stockUbic:{...vacio()} },
  { id:'cf22', label:'Cab. Franc 2022', max:2834, hex:'#7B2040', stockUbic:{...vacio(),'R Peña':894,'Pipi':18,'Lucas':15} },
  { id:'ch22', label:'Chardonnay 2022', max:2208, hex:'#C4A84F', stockUbic:{...vacio()} },
];
const totalStock = p => UBICACIONES.reduce((s,u)=>s+(p.stockUbic[u]||0),0);

// Maneja el stock por producto/depósito, persistido localmente y (si hay backend) enviado a Sheets
function useProductos(backendStock) {
  const [productos, setProductos] = useState(() => {
    try {
      const raw = localStorage.getItem('productos_stock');
      if (raw) {
        const saved = JSON.parse(raw);
        return PRODUCTOS_SEED.map(p => ({ ...p, stockUbic: { ...p.stockUbic, ...(saved[p.id]||{}) } }));
      }
    } catch {}
    return PRODUCTOS_SEED;
  });

  const persist = (list) => {
    try {
      const toSave = Object.fromEntries(list.map(p => [p.id, p.stockUbic]));
      localStorage.setItem('productos_stock', JSON.stringify(toSave));
    } catch {}
  };

  // Cuando llega el estado real desde Sheets (APP_STOCK_UBICACION), pasa a ser la fuente de verdad,
  // así todos los socios ven lo mismo en vez de quedarse con lo último guardado en este celular.
  useEffect(() => {
    if (!backendStock) return;
    setProductos(prev => {
      const next = prev.map(p => {
        const fromSheet = backendStock[p.label];
        return fromSheet ? { ...p, stockUbic: { ...vacio(), ...fromSheet } } : p;
      });
      persist(next);
      return next;
    });
  }, [backendStock]);

  // Mueve botellas de un depósito a otro para un producto, sin generar una venta
  const applyTransfer = ({ productoId, desde, hacia, cantidad }) => {
    setProductos(prev => {
      const next = prev.map(p => {
        if (p.id !== productoId) return p;
        const stockUbic = { ...p.stockUbic };
        stockUbic[desde] = (stockUbic[desde]||0) - cantidad;
        stockUbic[hacia] = (stockUbic[hacia]||0) + cantidad;
        return { ...p, stockUbic };
      });
      persist(next);
      return next;
    });
  };

  // Descuenta botellas de un depósito por una venta (no van a ningún otro lado, salen de la bodega)
  const applySale = ({ productoId, deposito, cantidad }) => {
    setProductos(prev => {
      const next = prev.map(p => {
        if (p.id !== productoId) return p;
        const stockUbic = { ...p.stockUbic };
        stockUbic[deposito] = (stockUbic[deposito]||0) - cantidad;
        return { ...p, stockUbic };
      });
      persist(next);
      return next;
    });
  };

  return { productos, applyTransfer, applySale };
}
const CAJAS    = ['Empresa (Ludico)','Mati','Lucas','Pipi (Andrés)','Santi'];
const CANALES  = ['Vinoteca','Distribuidor','Restaurant','Directo','Exportación','Otro'];

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
  const [ops,    setOps]    = useState(GAS_URL ? [] : SEED_OPS); // SEED_OPS es sólo un placeholder para modo offline sin GAS
  const [source, setSource] = useState('loading');
  const [stockUbicacion, setStockUbicacion] = useState(null); // { 'Malbec 2021': {'R Peña':0,...}, ... } | null
  const [ventasPendientes, setVentasPendientes] = useState([]); // [{referencia,cliente,producto,saldoArs,saldoUsd}]
  const [operacionesPendientes, setOperacionesPendientes] = useState([]); // ventas por cobrar + compras por pagar, combinadas
  const [comprasPendientes, setComprasPendientes] = useState([]); // [{referencia,detalle,proveedor,producto,saldoArs,saldoUsd}]
  const [clientes, setClientes] = useState([]); // [{nombre,canal,activo}] — real, desde la pestaña CLIENTES
  const [categorias, setCategorias] = useState([]); // [{detalle,tipo,total}] — categorías reales de BALANCE (Venta, Retiros, Muestra, Tapones, etc.)
  const [contactosBalance, setContactosBalance] = useState([]); // clientes Y proveedores mezclados, desde SUBDETALLE
  const [ultimoControlStock, setUltimoControlStock] = useState(undefined); // undefined=sin cargar aún; null=cargó y no hay control; objeto=hay control
  const [resumenCajas, setResumenCajas] = useState([]); // [{caja,ars,usd,crypto}] — real, desde el cuadro CAJAS al pie de CAJA

  const refresh = async () => {
    if (!GAS_URL) {
      // Sin GAS: usar storage local
      setPrice(8000); setSource('local');
      try { const r=localStorage.getItem('ops'); if(r) setOps(JSON.parse(r)); } catch {}
      return;
    }
    try {
      const d = await gasGet({ action: 'getData' });
      if (d?.ok) {
        if (d.lastPrice) { setPrice(d.lastPrice); setSource('sheets'); }
        else             { setPrice(8000);         setSource('fallback'); }
        if (d.ops) setOps(d.ops); // confiamos en la respuesta del backend siempre, incluso si viene vacía
        if (d.stockUbicacion) setStockUbicacion(d.stockUbicacion);
        if (d.ventasPendientes) setVentasPendientes(d.ventasPendientes);
        if (d.operacionesPendientes) setOperacionesPendientes(d.operacionesPendientes);
        if (d.comprasPendientes) setComprasPendientes(d.comprasPendientes);
        if (d.clientes?.length) setClientes(d.clientes);
        if (d.categorias?.length) setCategorias(d.categorias);
        if (d.contactosBalance?.length) setContactosBalance(d.contactosBalance);
        setUltimoControlStock(d.ultimoControlStock ?? null);
        if (d.resumenCajas) setResumenCajas(d.resumenCajas);
      }
    } catch { setPrice(8000); setSource('fallback'); }
  };

  useEffect(() => { refresh(); }, []);

  const addOp = async (op) => {
    const next = [op, ...ops].slice(0, 20);
    setOps(next);
    try { localStorage.setItem('ops', JSON.stringify(next)); } catch {}
  };

  return { price, source, ops, addOp, refresh, stockUbicacion, ventasPendientes, comprasPendientes, operacionesPendientes, clientes, categorias, contactosBalance, ultimoControlStock, resumenCajas };
}

function useStockControl(backendControl) {
  const [last,setLast]=useState(null);
  useEffect(()=>{
    if (!GAS_URL) {
      try{const r=localStorage.getItem('stock_ctrl');if(r)setLast(JSON.parse(r));}catch{}
      return;
    }
    // Con GAS conectado, el backend manda siempre — incluso si todavía no hay
    // ningún control cargado (backendControl===null), para no mostrar un control
    // viejo guardado en este celular de antes de que el control físico se sincronizara.
    if (backendControl !== undefined) setLast(backendControl);
  },[backendControl]);
  const save=async(ctrl)=>{ setLast(ctrl); try{localStorage.setItem('stock_ctrl',JSON.stringify(ctrl));}catch{}; };
  return {last,save};
}

async function getPin(uid){try{return localStorage.getItem(`pin_${uid}`)||'1234';}catch{return'1234';}}
async function savePin(uid,pin){try{localStorage.setItem(`pin_${uid}`,pin);return true;}catch{return false;}}

// ── CONTEXTO IA ──────────────────────────────────────────────────────────
function buildCtx(tc, price, productos) {
  const total=productos.reduce((s,p)=>s+totalStock(p),0);
  const ars=price?total*price:null; const usd=ars&&tc?Math.round(ars/tc):null;
  const stockLine=productos.map(p=>`${p.label}:${totalStock(p)||'agotado'}`).join(' | ');
  const ubicLine=productos.filter(p=>totalStock(p)>0).map(p=>`${p.label} [${UBICACIONES.filter(u=>p.stockUbic[u]>0).map(u=>`${u}:${p.stockUbic[u]}`).join(', ')}]`).join(' | ');
  return `Sos el asistente de Novato, bodega boutique de Mendoza. Datos al ${new Date().toLocaleDateString('es-AR')}.
STOCK: ${stockLine}. Total:${total} bot. Precio ref.:$${price||'N/A'}/bot. Valorización:${ars?'ARS $'+ars.toLocaleString('es-AR'):'—'}${usd?' ≈ USD '+usd.toLocaleString('es-AR'):''}.  TC blue:${tc?'$'+tc.toLocaleString('es-AR')+'/USD':'N/A'}.
STOCK POR DEPÓSITO: ${ubicLine||'sin datos'}.
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

// ── STOCK (control físico + transferencias entre depósitos) ────────────────
function StockScreen({onBack,showToast,productos,applyTransfer,user,refresh,last,save}){
  const [view,setView]=useState('control'); // 'control' | 'transfer'
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:16}}>
        <button onClick={onBack} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px 12px',color:C.muted,cursor:'pointer',fontSize:16}}>←</button>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>Stock</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Control físico y movimientos entre depósitos</div></div>
      </div>
      <div style={{display:'flex',marginBottom:20,border:`1px solid ${C.border}`,borderRadius:10,overflow:'hidden'}}>
        {[{k:'control',l:'Control físico'},{k:'transfer',l:'Mover entre depósitos'}].map((t,i)=>(
          <button key={t.k} onClick={()=>setView(t.k)} style={{flex:1,padding:'12px 8px',background:view===t.k?C.cork:C.barrel,color:view===t.k?C.gold:C.dim,border:'none',borderRight:i===0?`1px solid ${C.border}`:'none',cursor:'pointer',fontSize:13,fontFamily:'system-ui',fontWeight:view===t.k?700:400}}>{t.l}</button>
        ))}
      </div>
      {view==='control'
        ? <StockControlBody productos={productos} showToast={showToast} onBack={onBack} last={last} save={save} user={user} refresh={refresh}/>
        : <TransferForm productos={productos} applyTransfer={applyTransfer} user={user} showToast={showToast} onBack={onBack} refresh={refresh}/>}
    </div>
  );
}

function StockControlBody({productos,showToast,onBack,last,save,user,refresh}){
  const productosOrdenados=[...productos].sort((a,b)=>totalStock(b)-totalStock(a)); // orden general por stock total
  const [deposito,setDeposito]=useState(UBICACIONES[0]);
  const productosDeposito=productosOrdenados.filter(p=>(p.stockUbic[deposito]||0)>0); // solo los que tienen stock EN ESTE depósito
  const [real,setReal]=useState(Object.fromEntries(productosDeposito.map(p=>[p.id,'0'])));
  useEffect(()=>{ setReal(Object.fromEntries(productosDeposito.map(p=>[p.id,'0']))); },[deposito]); // no arrastrar números de otro depósito
  const set=(id,v)=>setReal(p=>({...p,[id]:v.replace(/\D/g,'')}));
  const diffs=productosDeposito.map(p=>{
    const teorico=p.stockUbic[deposito]||0;
    const r=real[p.id]===''?null:parseInt(real[p.id]);
    return{...p,stock:teorico,real:r,diff:r!==null?r-teorico:null};
  }).filter(p=>p.real!==null);
  const hasDiff=diffs.some(d=>d.diff!==0);const allFilled=productosDeposito.length>0&&productosDeposito.every(p=>real[p.id]!=='');
  const [sending,setSending]=useState(false);
  const confirm=async()=>{
    const ctrl={deposito,fecha:new Date().toLocaleDateString('es-AR'),hora:new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'}),items:diffs};
    setSending(true);
    try{
      if(GAS_URL){
        const items=diffs.map(d=>({label:d.label,stock:d.stock,real:d.real,diff:d.diff}));
        await gasGet({action:'addStockControl',deposito,fecha:ctrl.fecha,hora:ctrl.hora,items:JSON.stringify(items),user:user.nombre});
        await refresh();
      }
      await save(ctrl);
      if(hasDiff){showToast(`⚠ Control de ${deposito} guardado${GAS_URL?' en Sheets':''} — diferencias en ${diffs.filter(d=>d.diff!==0).length} producto(s)`,'error');}
      else{showToast(`✓ ${deposito}: stock coincide con lo teórico${GAS_URL?' (guardado en Sheets)':''}`,'ok');}
      onBack();
    }catch(e){ showToast('Error al guardar el control. Intentá de nuevo.','error'); }
    setSending(false);
  };
  return(<>
      <F label="Depósito a controlar *">
        <Sel value={deposito} onChange={e=>setDeposito(e.target.value)}>
          {UBICACIONES.map(u=><option key={u} value={u}>{u}</option>)}
        </Sel>
      </F>
      {last&&<div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:12,color:C.muted,margin:'16px 0'}}>Último control: <strong style={{color:C.text}}>{last.deposito?`${last.deposito} · `:''}{last.fecha} a las {last.hora}</strong></div>}
      {productosDeposito.length===0&&<div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',textAlign:'center',padding:'20px 0'}}>No hay stock teórico registrado en {deposito}</div>}
      {productosDeposito.length>0&&(<div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:20,marginTop:last?0:16}}>
        <div style={{display:'grid',gridTemplateColumns:'1fr 80px 80px 80px',gap:8,padding:'0 4px'}}>
          {['Producto','Teórico','Real','Dif.'].map(h=><div key={h} style={{fontSize:10,color:C.dim,letterSpacing:'0.1em',textTransform:'uppercase',fontFamily:'system-ui',textAlign:h!=='Producto'?'center':'left'}}>{h}</div>)}
        </div>
        {productosDeposito.map(p=>{
          const stock=p.stockUbic[deposito]||0;
          const r=real[p.id]===''?null:parseInt(real[p.id]);
          const diff=r!==null?r-stock:null;
          const dc=diff===null?C.dim:diff<0?'#f08080':diff>0?'#7dce9b':C.muted;
          return(
            <div key={p.id} style={{background:C.barrel,border:`1px solid ${diff!==null&&diff!==0?dc+'44':C.border}`,borderRadius:12,padding:'12px 14px',display:'grid',gridTemplateColumns:'1fr 80px 80px 80px',gap:8,alignItems:'center'}}>
              <div><div style={{color:C.text,fontSize:14,fontFamily:'system-ui',fontWeight:600}}>{p.label}</div><div style={{width:40,height:3,background:p.hex,borderRadius:2,marginTop:4}}/></div>
              <div style={{textAlign:'center',color:C.muted,fontSize:15,fontFamily:'system-ui'}}>{stock}</div>
              <input type="number" min="0" placeholder="—" value={real[p.id]} onChange={e=>set(p.id,e.target.value)} onFocus={e=>e.target.select()} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px',color:C.text,fontSize:15,fontFamily:'system-ui',outline:'none',width:'100%',boxSizing:'border-box',textAlign:'center'}}/>
              <div style={{textAlign:'center',fontWeight:700,fontSize:15,fontFamily:'system-ui',color:dc}}>{diff===null?'—':diff===0?'✓':diff>0?`+${diff}`:diff}</div>
            </div>
          );
        })}
      </div>)}
      {diffs.some(d=>d.diff!==0&&d.diff!==null)&&(
        <div style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:12,padding:'14px 16px',marginBottom:16,fontFamily:'system-ui'}}>
          <div style={{color:'#E07080',fontSize:13,fontWeight:700,marginBottom:8}}>⚠ Diferencias detectadas en {deposito}</div>
          {diffs.filter(d=>d.diff!==0&&d.diff!==null).map(d=>(
            <div key={d.id} style={{display:'flex',justifyContent:'space-between',fontSize:13,color:C.muted,paddingBottom:6,marginBottom:6,borderBottom:`1px solid ${C.border}`}}>
              <span>{d.label}</span><span style={{color:d.diff<0?'#f08080':'#7dce9b',fontWeight:700}}>{d.diff>0?'+':''}{d.diff} bot</span>
            </div>
          ))}
          <div style={{color:'#A06070',fontSize:12,marginTop:4,lineHeight:1.5}}>Causas posibles: retiros sin avisar, cobro en especie del depósito, o un movimiento a/desde {deposito} sin cargar (usá "Mover entre depósitos").</div>
        </div>
      )}
      <Btn onClick={confirm} style={{width:'100%',opacity:(allFilled&&!sending)?1:0.5}} disabled={!allFilled||sending}>{sending?'Guardando…':allFilled?`Confirmar control de ${deposito}`:'Completá todos los campos'}</Btn>
  </>);
}

// ── TRANSFERENCIA ENTRE DEPÓSITOS ───────────────────────────────────────
function TransferForm({productos,applyTransfer,user,showToast,onBack,refresh}){
  const hoy=new Date().toISOString().split('T')[0];
  const [f,setF]=useState({producto:'',cantidad:'',desde:'',hacia:'',fecha:hoy,notas:''});
  const [sending,setSending]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const prod=productos.find(p=>p.id===f.producto);
  const disponible=prod&&f.desde?(prod.stockUbic[f.desde]||0):null;
  const submit=async()=>{
    const cant=parseInt(f.cantidad||0);
    if(!f.producto||!cant||!f.desde||!f.hacia){showToast('Completá los campos obligatorios','error');return;}
    if(f.desde===f.hacia){showToast('Origen y destino no pueden ser el mismo depósito','error');return;}
    if(disponible!==null&&cant>disponible){showToast(`Sólo hay ${disponible} bot en ${f.desde}`,'error');return;}
    setSending(true);
    try{
      applyTransfer({productoId:f.producto,desde:f.desde,hacia:f.hacia,cantidad:cant});
      if(GAS_URL){
        await gasGet({action:'addTransfer',fecha:f.fecha,producto:prod?.label,cantidad:cant,desde:f.desde,hacia:f.hacia,notas:f.notas||'',user:user.nombre});
        await refresh();
      }
      showToast(`✓ ${cant} bot de ${prod?.label}: ${f.desde} → ${f.hacia}${GAS_URL?' (enviado a Sheets)':''}`,'ok');
      onBack();
    }catch(e){ showToast('Error al registrar el movimiento.','error'); }
    setSending(false);
  };
  return(
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <F label="Producto *"><Sel value={f.producto} onChange={e=>set('producto',e.target.value)}><option value="">Seleccionar…</option>{productos.filter(p=>totalStock(p)>0).map(p=><option key={p.id} value={p.id}>{p.label} — {totalStock(p)} bot</option>)}</Sel></F>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
        <F label="Desde *"><Sel value={f.desde} onChange={e=>set('desde',e.target.value)}><option value="">Seleccionar…</option>{UBICACIONES.map(u=><option key={u} value={u}>{u}{prod?` (${prod.stockUbic[u]||0})`:''}</option>)}</Sel></F>
        <F label="Hacia *"><Sel value={f.hacia} onChange={e=>set('hacia',e.target.value)}><option value="">Seleccionar…</option>{UBICACIONES.filter(u=>u!==f.desde).map(u=><option key={u} value={u}>{u}</option>)}</Sel></F>
      </div>
      <F label="Cantidad (botellas) *"><Inp type="number" min="1" max={disponible||undefined} placeholder="0" value={f.cantidad} onChange={e=>set('cantidad',e.target.value)}/></F>
      {disponible!==null&&<div style={{fontSize:12,color:C.dim,fontFamily:'system-ui'}}>Disponible en {f.desde}: {disponible} bot</div>}
      <F label="Fecha"><Inp type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></F>
      <F label="Notas"><Inp placeholder="Motivo del movimiento (opcional)" value={f.notas} onChange={e=>set('notas',e.target.value)}/></F>
      <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':'Registrar movimiento'}</Btn></div>
    </div>
  );
}

// ── DASHBOARD ────────────────────────────────────────────────────────────
function DashboardScreen({onNavigate,price,source,productos,last,operacionesPendientes,resumenCajas}){
  const {tc,err:tcErr,date:tcDate}=useDolarBlue();
  const total=productos.reduce((s,p)=>s+totalStock(p),0);
  const productosOrdenados=[...productos].filter(p=>totalStock(p)>0).sort((a,b)=>totalStock(b)-totalStock(a)); // solo con stock, mayor a menor
  const ars=price?total*price:null; const usd=ars&&tc?Math.round(ars/tc):null;
  const totalCajaArs=Math.round(resumenCajas?.reduce((s,c)=>s+c.ars,0)||0);
  const totalCajaUsd=Math.round(resumenCajas?.reduce((s,c)=>s+c.usd,0)||0);
  const daysSince=last?(()=>{try{const [d,m,y]=last.fecha.split('/');return Math.floor((Date.now()-new Date(`${y}-${m}-${d}`).getTime())/86400000);}catch{return null;}})():null;
  return(
    <div style={{padding:'20px 16px',maxWidth:540,margin:'0 auto'}}>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:10,marginBottom:12}}>
        {[{v:total.toLocaleString('es-AR'),u:'botellas',l:'Stock total'},{v:usd?`USD ${usd.toLocaleString('es-AR')}`:'…',u:'valorización',l:price?`$${price.toLocaleString('es-AR')}/bot`:'cargando…'}].map(k=><Card key={k.l} style={{padding:'12px 10px'}}><div style={{color:C.gold,fontSize:17,fontWeight:700,lineHeight:1,fontFamily:'Georgia, serif'}}>{k.v}</div><div style={{color:C.dim,fontSize:10,fontFamily:'system-ui',marginTop:2}}>{k.u}</div><div style={{color:C.muted,fontSize:10,fontFamily:'system-ui',marginTop:5,borderTop:`1px solid ${C.border}`,paddingTop:5}}>{k.l}</div></Card>)}
        <Card style={{padding:'12px 10px'}}>
          <div style={{color:C.gold,fontSize:14,fontWeight:700,lineHeight:1.3,fontFamily:'Georgia, serif'}}>${totalCajaArs.toLocaleString('es-AR')}</div>
          <div style={{color:C.gold,fontSize:12,fontWeight:700,lineHeight:1.3,fontFamily:'Georgia, serif',opacity:0.8}}>USD {totalCajaUsd.toLocaleString('es-AR')}</div>
          <div style={{color:C.muted,fontSize:10,fontFamily:'system-ui',marginTop:4,borderTop:`1px solid ${C.border}`,paddingTop:5}}>Caja empresa</div>
        </Card>
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
          {productosOrdenados.map(p=>{const stock=totalStock(p);const pct=p.max>0?stock/p.max:0;const h=Math.max(Math.round(pct*90),stock>0?4:2);return(<div key={p.id} style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',gap:4}}><div style={{fontSize:9,color:stock>0?C.muted:C.dim,fontFamily:'system-ui',minHeight:13}}>{stock>0?stock.toLocaleString('es-AR'):''}</div><div style={{width:'100%',display:'flex',flexDirection:'column',alignItems:'center'}}><div style={{width:'35%',height:13,background:C.cork,borderRadius:'3px 3px 0 0',border:`1px solid ${C.border}`,borderBottom:'none'}}/><div style={{width:'100%',height:80,background:C.cork,borderRadius:'2px 2px 6px 6px',border:`1px solid ${C.border}`,overflow:'hidden',display:'flex',alignItems:'flex-end'}}><div style={{width:'100%',height:`${h}px`,background:stock>0?p.hex:C.cork,transition:'height 0.8s cubic-bezier(.4,0,.2,1)'}}/></div></div></div>);})}
        </div>
        <div style={{display:'flex',gap:8}}>{productosOrdenados.map(p=><div key={p.id} style={{flex:1,textAlign:'center'}}><div style={{fontSize:8,color:totalStock(p)>0?C.muted:C.dim,fontFamily:'system-ui',lineHeight:1.3}}>{p.label.split(' ').map((w,i)=><span key={i}>{w}<br/></span>)}</div></div>)}</div>
      </Card>
      <SL>Stock por depósito</SL>
      <Card style={{marginBottom:14}}>
        <div style={{display:'flex',flexDirection:'column',gap:10}}>
          {productosOrdenados.map(p=>(
            <div key={p.id}>
              <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:5}}><div style={{width:8,height:8,borderRadius:2,background:p.hex,flexShrink:0}}/><span style={{color:C.text,fontSize:12,fontFamily:'system-ui',fontWeight:600}}>{p.label}</span></div>
              <div style={{display:'flex',flexWrap:'wrap',gap:6,paddingLeft:14}}>
                {UBICACIONES.filter(u=>(p.stockUbic[u]||0)>0).map(u=>(
                  <div key={u} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'3px 9px',fontSize:11,fontFamily:'system-ui',color:C.muted}}>{u} <strong style={{color:C.text}}>{p.stockUbic[u]}</strong></div>
                ))}
              </div>
            </div>
          ))}
          {productos.every(p=>totalStock(p)===0)&&<div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',textAlign:'center',padding:'8px 0'}}>Sin stock disponible</div>}
        </div>
      </Card>
      {/* Card control de stock */}
      <div style={{background:daysSince===null||daysSince>30?C.wineBg:C.barrel,border:`1px solid ${daysSince===null||daysSince>30?C.wine+'44':C.border}`,borderRadius:14,padding:'14px 16px',marginBottom:18}}>
        {/* Header */}
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom: last?.items?.filter(d=>d.diff!==0).length>0?12:0}}>
          <div>
            <div style={{color:daysSince===null||daysSince>30?'#E07080':C.text,fontSize:13,fontWeight:700,fontFamily:'system-ui',marginBottom:2}}>
              {daysSince===null?'⚠ Control de stock pendiente':daysSince>30?`⚠ Último control hace ${daysSince} días`:`✓ Control de stock · hace ${daysSince} día${daysSince===1?'':'s'}`}
            </div>
            <div style={{color:daysSince===null||daysSince>30?'#A06070':C.dim,fontSize:12,fontFamily:'system-ui'}}>
              {last?`${last.deposito?last.deposito+' · ':''}${last.fecha} a las ${last.hora}`:'Nunca realizado'}
            </div>
          </div>
          <button onClick={()=>onNavigate('stock')} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'7px 12px',color:C.gold,fontSize:12,fontFamily:'system-ui',fontWeight:700,cursor:'pointer',flexShrink:0}}>
            + Nuevo
          </button>
        </div>
        {/* Diferencias del último control */}
        {last?.items?.filter(d=>d.diff!==0).length>0&&(
          <div style={{borderTop:`1px solid ${C.border}`,paddingTop:10,display:'flex',flexDirection:'column',gap:6}}>
            <div style={{fontSize:10,color:'#E07080',letterSpacing:'0.1em',textTransform:'uppercase',fontFamily:'system-ui',marginBottom:2}}>Diferencias pendientes de resolver</div>
            {last.items.filter(d=>d.diff!==0).map(d=>(
              <div key={d.id} style={{display:'flex',alignItems:'center',gap:10}}>
                <div style={{width:3,height:28,background:d.hex||C.wine,borderRadius:2,flexShrink:0}}/>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,color:C.text,fontFamily:'system-ui',fontWeight:600}}>{d.label}</div>
                  <div style={{fontSize:11,color:C.dim,fontFamily:'system-ui'}}>Teórico: {d.stock} bot · Real: {d.real} bot</div>
                </div>
                <div style={{fontWeight:700,fontSize:15,fontFamily:'system-ui',color:d.diff<0?'#f08080':'#7dce9b',flexShrink:0}}>
                  {d.diff>0?'+':''}{d.diff} bot
                </div>
              </div>
            ))}
          </div>
        )}
        {last&&last.items?.filter(d=>d.diff!==0).length===0&&(
          <div style={{borderTop:`1px solid ${C.border}`,paddingTop:8,fontSize:12,color:C.green,fontFamily:'system-ui'}}>✓ Stock coincide con lo teórico</div>
        )}
      </div>
      <SL>Operaciones pendientes</SL>
      <div style={{display:'flex',flexDirection:'column',gap:8}}>
        {operacionesPendientes.slice(0,8).map((op,i)=>{
          const dias=op.fecha?Math.floor((Date.now()-new Date(op.fecha).getTime())/86400000):null;
          const esVenta=op.tipo==='venta';
          const vencida=esVenta&&dias!==null&&dias>60;
          return (
            <div key={op.referencia+'-'+i} style={{background:C.barrel,border:`1px solid ${vencida?'#f0808088':C.border}`,borderRadius:12,padding:'12px 14px',display:'flex',alignItems:'center',gap:12}}>
              <span style={{fontSize:22,flexShrink:0}}>{esVenta?'🍾':'📋'}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{color:C.text,fontSize:13,fontFamily:'system-ui',fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{op.contraparte||'(sin nombre)'}{op.producto?` · ${op.producto}`:''}</div>
                <div style={{color:vencida?'#f08080':C.dim,fontSize:11,fontFamily:'system-ui',marginTop:2,fontWeight:vencida?700:400}}>{vencida?'⚠ ':''}{op.detalle}{dias!==null?` · hace ${dias} día${dias===1?'':'s'}`:''}</div>
              </div>
              <div style={{fontSize:13,fontFamily:'system-ui',flexShrink:0,textAlign:'right',fontWeight:700,color:esVenta?'#7dce9b':'#f08080'}}>
                {esVenta?'+':'-'}${Math.abs(op.saldoArs).toLocaleString('es-AR')}
              </div>
            </div>
          );
        })}
        {operacionesPendientes.length===0&&<div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',textAlign:'center',padding:'20px 0'}}>Sin operaciones pendientes</div>}
        {operacionesPendientes.length>8&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',textAlign:'center',padding:'4px 0'}}>+{operacionesPendientes.length-8} más</div>}
      </div>
    </div>
  );
}

// ── NUEVA VENTA ──────────────────────────────────────────────────────────
function VentaScreen({user,onBack,showToast,addOp,price,productos,applySale,clientes,categorias,contactosBalance,refresh}){
  const hoy=new Date().toISOString().split('T')[0];
  const [f,setF]=useState({detalle:'Venta',producto:'',botellas:'',contacto:'',contactoNuevo:'',canal:'',monto:'',fecha:hoy,deposito:''});
  const [sending,setSending]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const prod=productos.find(p=>p.id===f.producto);
  const esVenta=f.detalle==='Venta';
  const catInfo=categorias?.find(c=>c.detalle===f.detalle);
  const tipoBadge=catInfo?.tipo||'Ingreso';

  const listaDetalles = categorias?.length ? categorias.map(c=>c.detalle) : ['Venta'];
  const listaContactos = Array.from(new Set([...(clientes?.map(c=>c.nombre)||[]), ...(contactosBalance||[])])).sort((a,b)=>a.localeCompare(b,'es'));
  const clienteInfo = esVenta ? clientes?.find(c=>c.nombre===f.contacto) : null;
  const setContacto = (nombre) => {
    setF(p=>({...p, contacto:nombre, canal: clientes?.find(c=>c.nombre===nombre)?.canal || p.canal }));
  };

  const sug=esVenta&&price&&f.botellas?price*parseInt(f.botellas||0):null;
  const requiereStock=parseInt(f.botellas||0)>0;
  const disponibleEnDeposito=prod&&f.deposito?(prod.stockUbic[f.deposito]||0):null;

  const submit=async()=>{
    if(!f.detalle||!f.contacto||(f.contacto==='Nuevo…'&&!f.contactoNuevo)){showToast('Completá los campos obligatorios','error');return;}
    if(requiereStock&&(!f.producto||!f.deposito)){showToast('Con botellas, indicá producto y depósito','error');return;}
    const cant=parseInt(f.botellas)||0;
    if(requiereStock&&disponibleEnDeposito!==null&&cant>disponibleEnDeposito){showToast(`Sólo hay ${disponibleEnDeposito} bot en ${f.deposito}`,'error');return;}
    const ct=f.contacto==='Nuevo…'?f.contactoNuevo:f.contacto;
    setSending(true);
    try {
      let referencia='';
      if(GAS_URL){
        const r=await gasGet({action:'addTransaccion',fecha:f.fecha,detalle:f.detalle,producto:prod?.label||'',botellas:cant,monto:f.monto||0,contacto:ct,canal:esVenta?(f.canal||''):'',deposito:f.deposito,user:user.nombre});
        referencia=r?.referencia||'';
      }
      if(requiereStock) applySale({productoId:f.producto,deposito:f.deposito,cantidad:cant});
      const montoStr=f.monto?`$${parseInt(f.monto).toLocaleString('es-AR')}`:'sin monto';
      await addOp({id:Date.now(),icon:esVenta?'🍾':'📋',desc:`${f.detalle}${cant?` · ${cant} bot`:''}${prod?` ${prod.label}`:''} → ${ct}`,monto:montoStr,fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      if(GAS_URL) await refresh();
      showToast(`✓ ${f.detalle} registrada en BALANCE${referencia?' ('+referencia+')':''}`,'ok');
      onBack();
    } catch(e){ showToast('Error al registrar. Intentá de nuevo.','error'); }
    setSending(false);
  };
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:24}}>
        <button onClick={onBack} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:8,padding:'8px 12px',color:C.muted,cursor:'pointer',fontSize:16}}>←</button>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>Nueva transacción</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Registrar en BALANCE</div></div>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <F label="Detalle *">
          <Sel value={f.detalle} onChange={e=>set('detalle',e.target.value)}>
            {listaDetalles.map(d=><option key={d} value={d}>{d}</option>)}
          </Sel>
        </F>
        {tipoBadge!=='Neutro'&&<div style={{fontSize:11,color:C.dim,fontFamily:'system-ui',marginTop:-8}}>Se registra como <strong style={{color:tipoBadge==='Ingreso'?'#7dce9b':'#f08080'}}>{tipoBadge.toLowerCase()}</strong> automáticamente.</div>}
        <F label="Cliente / Proveedor *">
          <Sel value={f.contacto} onChange={e=>setContacto(e.target.value)}>
            <option value="">Seleccionar…</option>
            {listaContactos.map(c=><option key={c} value={c}>{c}</option>)}
            <option value="Nuevo…">Nuevo…</option>
          </Sel>
        </F>
        {clienteInfo&&!clienteInfo.activo&&<div style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:10,padding:'8px 14px',fontFamily:'system-ui',fontSize:12,color:'#E07080'}}>⚠ {f.contacto} figura inactivo — última operación hace tiempo</div>}
        {f.contacto==='Nuevo…'&&<F label="Nombre"><Inp placeholder="Nombre o comercio" value={f.contactoNuevo} onChange={e=>set('contactoNuevo',e.target.value)}/></F>}
        <F label={`Producto${requiereStock?' *':' (opcional)'}`}>
          <Sel value={f.producto} onChange={e=>set('producto',e.target.value)}>
            <option value="">Seleccionar…</option>
            {productos.filter(p=>totalStock(p)>0).map(p=><option key={p.id} value={p.id}>{p.label} — {totalStock(p).toLocaleString('es-AR')} bot</option>)}
          </Sel>
        </F>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
          <F label={`Depósito${requiereStock?' *':' (opcional)'}`}>
            <Sel value={f.deposito} onChange={e=>set('deposito',e.target.value)}>
              <option value="">Seleccionar…</option>
              {UBICACIONES.filter(u=>!prod||(prod.stockUbic[u]||0)>0).map(u=><option key={u} value={u}>{u}{prod?` (${prod.stockUbic[u]||0})`:''}</option>)}
            </Sel>
          </F>
          <F label="Botellas (opcional)"><Inp type="number" min="0" max={disponibleEnDeposito||(prod?totalStock(prod):undefined)} placeholder="0" value={f.botellas} onChange={e=>set('botellas',e.target.value)}/></F>
        </div>
        <F label="Fecha"><Inp type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></F>
        {sug&&<div style={{background:C.greenBg,border:`1px solid ${C.green}55`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:13}}><span style={{color:C.muted}}>Monto sugerido: </span><strong style={{color:'#7dce9b'}}>${sug.toLocaleString('es-AR')}</strong><span style={{color:C.dim}}> (${price.toLocaleString('es-AR')}/bot · última venta)</span></div>}
        <F label="Monto (opcional, $ ARS)"><Inp type="number" placeholder="0" value={f.monto} onChange={e=>set('monto',e.target.value)}/></F>
        {esVenta&&<F label="Canal"><Sel value={f.canal} onChange={e=>set('canal',e.target.value)}><option value="">Seleccionar…</option>{CANALES.map(c=><option key={c} value={c}>{c}</option>)}</Sel></F>}
        <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':'Registrar'}</Btn></div>
      </div>
    </div>
  );
}

// ── CAJA ─────────────────────────────────────────────────────────────────
function CajaScreen({user,onBack,showToast,addOp,ventasPendientes,comprasPendientes,refresh,resumenCajas,categorias,contactosBalance,clientes}){
  const hoy=new Date().toISOString().split('T')[0];
  const [f,setF]=useState({tipo:'cobro',monto:'',caja:'Empresa (Ludico)',detalle:'',contacto:'',contactoNuevo:'',fecha:hoy,referencia:''});
  const [sending,setSending]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v,...((k==='detalle'||k==='contacto')?{referencia:''}:{})}));
  const listaDetalles = categorias?.length ? categorias.map(c=>c.detalle) : ['Cobro','Gasto'];
  const listaContactos = Array.from(new Set([...(clientes?.map(c=>c.nombre)||[]), ...(contactosBalance||[])])).sort((a,b)=>a.localeCompare(b,'es'));
  const submit=async()=>{
    if(!f.monto||!f.detalle||!f.contacto||(f.contacto==='Nuevo…'&&!f.contactoNuevo)){showToast('Completá los campos obligatorios','error');return;}
    const ct=f.contacto==='Nuevo…'?f.contactoNuevo:f.contacto;
    setSending(true);
    try {
      if(GAS_URL) await gasGet({action:'addMovement',fecha:f.fecha,tipo:f.tipo,monto:f.monto,caja:f.caja,detalle:f.detalle,contacto:ct,referencia:f.referencia||'',user:user.nombre});
      const icon=f.tipo==='cobro'?'💵':'💸';
      await addOp({id:Date.now(),icon,desc:`${f.detalle}: ${ct}`,monto:`$${parseInt(f.monto).toLocaleString('es-AR')}`,fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      if(GAS_URL) await refresh();
      showToast(`✓ Movimiento registrado${GAS_URL?' en CAJA':''}`,'ok');
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
      {resumenCajas?.length>0&&(
        <Card style={{marginBottom:20}}>
          <SL>Estado de cajas</SL>
          <div style={{display:'grid',gridTemplateColumns:'1fr 76px 76px',gap:6,padding:'0 2px 6px'}}>
            {['Caja','AR$','USD'].map((h,i)=><div key={h} style={{fontSize:9,color:C.dim,letterSpacing:'0.08em',textTransform:'uppercase',fontFamily:'system-ui',textAlign:i===0?'left':'center'}}>{h}</div>)}
          </div>
          {resumenCajas.map(c=>(
            <div key={c.caja} style={{display:'grid',gridTemplateColumns:'1fr 76px 76px',gap:6,padding:'7px 2px',borderTop:`1px solid ${C.border}`,alignItems:'center'}}>
              <span style={{color:C.text,fontSize:13,fontFamily:'system-ui',fontWeight:600}}>{c.caja}</span>
              <span style={{textAlign:'center',color:C.muted,fontSize:12,fontFamily:'system-ui'}}>{c.ars?`$${Math.round(c.ars).toLocaleString('es-AR')}`:'—'}</span>
              <span style={{textAlign:'center',color:C.muted,fontSize:12,fontFamily:'system-ui'}}>{c.usd?`$${Math.round(c.usd).toLocaleString('es-AR')}`:'—'}</span>
            </div>
          ))}
          <div style={{display:'grid',gridTemplateColumns:'1fr 76px 76px',gap:6,padding:'8px 2px 0',borderTop:`2px solid ${C.border}`,marginTop:2}}>
            <span style={{color:C.gold,fontSize:13,fontFamily:'system-ui',fontWeight:700}}>TOTAL</span>
            <span style={{textAlign:'center',color:C.gold,fontSize:12,fontFamily:'system-ui',fontWeight:700}}>${Math.round(resumenCajas.reduce((s,c)=>s+c.ars,0)).toLocaleString('es-AR')}</span>
            <span style={{textAlign:'center',color:C.gold,fontSize:12,fontFamily:'system-ui',fontWeight:700}}>${Math.round(resumenCajas.reduce((s,c)=>s+c.usd,0)).toLocaleString('es-AR')}</span>
          </div>
        </Card>
      )}
      <div style={{display:'flex',marginBottom:20,border:`1px solid ${C.border}`,borderRadius:10,overflow:'hidden'}}>
        {[{k:'cobro',l:'↓ Cobro / Ingreso',bg:C.greenBg,col:'#7dce9b'},{k:'gasto',l:'↑ Gasto / Egreso',bg:C.wineBg,col:'#f08080'}].map((t,i)=>(
          <button key={t.k} onClick={()=>{set('tipo',t.k);set('referencia','');}} style={{flex:1,padding:'12px 8px',background:f.tipo===t.k?t.bg:C.barrel,color:f.tipo===t.k?t.col:C.dim,border:'none',borderRight:i===0?`1px solid ${C.border}`:'none',cursor:'pointer',fontSize:13,fontFamily:'system-ui',fontWeight:f.tipo===t.k?700:400}}>{t.l}</button>
        ))}
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <F label="Monto *"><Inp type="number" placeholder="0" value={f.monto} onChange={e=>set('monto',e.target.value)}/></F>
        <F label={f.tipo==='cobro'?'Caja que recibe':'Caja de la que sale'}><Sel value={f.caja} onChange={e=>set('caja',e.target.value)}>{CAJAS.map(c=><option key={c} value={c}>{c}</option>)}</Sel></F>
        <F label="Detalle *">
          <Sel value={f.detalle} onChange={e=>set('detalle',e.target.value)}>
            <option value="">Seleccionar…</option>
            {listaDetalles.map(d=><option key={d} value={d}>{d}</option>)}
          </Sel>
        </F>
        <F label="Cliente / Proveedor *">
          <Sel value={f.contacto} onChange={e=>set('contacto',e.target.value)}>
            <option value="">Seleccionar…</option>
            {listaContactos.map(c=><option key={c} value={c}>{c}</option>)}
            <option value="Nuevo…">Nuevo…</option>
          </Sel>
        </F>
        {f.contacto==='Nuevo…'&&<F label="Nombre"><Inp placeholder="Nombre o comercio" value={f.contactoNuevo} onChange={e=>set('contactoNuevo',e.target.value)}/></F>}
        {f.tipo==='cobro'&&ventasPendientes?.length>0&&(
          <F label="Vincular a venta pendiente (opcional)">
            <Sel value={f.referencia} onChange={e=>set('referencia',e.target.value)}>
              <option value="">Sin vincular</option>
              {ventasPendientes.map(v=><option key={v.referencia} value={v.referencia}>{v.referencia} — {v.cliente} · debe ${v.saldoArs.toLocaleString('es-AR')}</option>)}
            </Sel>
          </F>
        )}
        {f.tipo==='gasto'&&comprasPendientes?.length>0&&(
          <F label="Vincular a compra pendiente (opcional)">
            <Sel value={f.referencia} onChange={e=>set('referencia',e.target.value)}>
              <option value="">Sin vincular</option>
              {comprasPendientes.map(c=><option key={c.referencia} value={c.referencia}>{c.referencia} — {c.detalle}{c.proveedor?` (${c.proveedor})`:''} · debe ${c.saldoArs.toLocaleString('es-AR')}</option>)}
            </Sel>
          </F>
        )}
        <F label="Fecha"><Inp type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></F>
        <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':'Registrar movimiento'}</Btn></div>
      </div>
    </div>
  );
}

// ── CONSULTAS IA ─────────────────────────────────────────────────────────
function ConsultasScreen({price,productos}){
  const {tc}=useDolarBlue();
  const [msgs,setMsgs]=useState([{role:'assistant',content:'Hola 👋 Soy el asistente de Novato. Preguntame sobre stock, clientes, caja o cualquier dato de la bodega.'}]);
  const [input,setInput]=useState('');const [loading,setL]=useState(false);const ref=useRef(null);
  useEffect(()=>ref.current?.scrollIntoView({behavior:'smooth'}),[msgs,loading]);
  const QUICK=['¿Cuánto stock tenemos disponible?','¿Cuánto vale el stock en dólares?','¿Qué clientes están inactivos?','Resumime el estado de la bodega'];
  const send=async(text)=>{
    const q=(text||input).trim();if(!q||loading)return;setInput('');
    const next=[...msgs,{role:'user',content:q}];setMsgs(next);setL(true);
    try{
      const res=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json','x-api-key':ANTHROPIC_KEY,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
        body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:1000,system:buildCtx(tc,price,productos),messages:next.slice(1)})
      });
      const d=await res.json();
      if(d?.content?.[0]?.text){
        setMsgs(p=>[...p,{role:'assistant',content:d.content[0].text}]);
      } else {
        const errMsg=d?.error?.message||d?.message||JSON.stringify(d);
        setMsgs(p=>[...p,{role:'assistant',content:'Error: '+errMsg}]);
      }
    }
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
  const tabs=[{id:'dashboard',icon:'⌂',l:'Inicio'},{id:'venta',icon:'🍾',l:'Cargar'},{id:'caja',icon:'⇄',l:'Caja'},{id:'stock',icon:'📦',l:'Stock'},{id:'consultas',icon:'✦',l:'Asistente'}];
  return(<div style={{position:'fixed',bottom:0,left:0,right:0,background:C.barrel,borderTop:`1px solid ${C.border}`,display:'flex',zIndex:200,paddingBottom:'env(safe-area-inset-bottom, 0px)'}}>{tabs.map(t=><button key={t.id} onClick={()=>setScreen(t.id)} style={{flex:1,padding:'10px 2px 12px',background:'none',border:'none',cursor:'pointer',display:'flex',flexDirection:'column',alignItems:'center',gap:2}}><span style={{fontSize:t.id==='dashboard'?20:17,opacity:screen===t.id?1:0.35}}>{t.icon}</span><span style={{fontSize:9,fontFamily:'system-ui',color:screen===t.id?C.gold:C.dim,fontWeight:screen===t.id?700:400}}>{t.l}</span></button>)}</div>);
}

// ── APP ───────────────────────────────────────────────────────────────────
export default function NovatoApp(){
  const [user,setUser]=useState(null);const [screen,setScreen]=useState('dashboard');const [toast,setToast]=useState(null);const [settings,setSettings]=useState(false);
  const {price,source,ops,addOp,stockUbicacion,ventasPendientes,comprasPendientes,operacionesPendientes,clientes,categorias,contactosBalance,ultimoControlStock,resumenCajas,refresh}=useAppData();
  const {productos,applyTransfer,applySale}=useProductos(stockUbicacion);
  const {last,save}=useStockControl(ultimoControlStock);
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
        {screen==='dashboard'&&<DashboardScreen onNavigate={setScreen} price={price} source={source} productos={productos} last={last} operacionesPendientes={operacionesPendientes} resumenCajas={resumenCajas}/>}
        {screen==='venta'&&<VentaScreen user={user} onBack={()=>setScreen('dashboard')} showToast={showToast} addOp={addOp} price={price} productos={productos} applySale={applySale} clientes={clientes} categorias={categorias} contactosBalance={contactosBalance} refresh={refresh}/>}
        {screen==='caja'&&<CajaScreen user={user} onBack={()=>setScreen('dashboard')} showToast={showToast} addOp={addOp} ventasPendientes={ventasPendientes} comprasPendientes={comprasPendientes} refresh={refresh} resumenCajas={resumenCajas} categorias={categorias} contactosBalance={contactosBalance} clientes={clientes}/>}
        {screen==='stock'&&<StockScreen onBack={()=>setScreen('dashboard')} showToast={showToast} productos={productos} applyTransfer={applyTransfer} user={user} refresh={refresh} last={last} save={save}/>}
        {screen==='consultas'&&<ConsultasScreen price={price} productos={productos}/>}
      </div>
      <Nav screen={screen} setScreen={setScreen}/>
    </div>
  );
}
