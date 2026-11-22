import { useState, useRef, useEffect } from "react";

// ── CONFIG ───────────────────────────────────────────────────────────────
// Vercel: agregar variable de entorno VITE_GAS_URL con la URL del deployment
const GAS_URL = import.meta.env.VITE_GAS_URL || '';

// ── TOKENS ───────────────────────────────────────────────────────────────
const C = {
  cellar:'#0D0B09', barrel:'#1E1914', cork:'#2E2820', border:'#3A322A',
  gold:'#C9AD55', goldDim:'#8A7438', wine:'#8C2030', wineBg:'#1A0A0D',
  text:'#F0EDE6', muted:'#B3A99E', dim:'#978D82', green:'#4A8C5C', greenBg:'#0A1A0F',
};

// ── DATOS ────────────────────────────────────────────────────────────────
const USUARIOS = [
  { id:'mati',   nombre:'Mati',   rol:'Socio' },
  { id:'andres', nombre:'Andrés', rol:'Socio' },
  { id:'lucas',  nombre:'Lucas',  rol:'Socio' },
  { id:'santi',  nombre:'Santi',  rol:'Socio' },
];
// Depósitos entre los que se puede mover stock sin que medie una venta
const UBICACIONES = ['R Peña','Pipi','Lucas','Santi','Mati'];
const vacio = () => Object.fromEntries(UBICACIONES.map(u=>[u,0]));

// Stock inicial por producto y depósito (según último conteo en General_Dinamico → hoja STOCK)
const PRODUCTOS_SEED = [
  { id:'mb21', label:'Malbec 2021',     max:2964, hex:'#6B2030', stockUbic:{...vacio()} },
  { id:'mb22', label:'Malbec 2022',     max:3818, hex:'#9B2335', stockUbic:{...vacio(),'R Peña':2964,'Pipi':30,'Lucas':15} },
  { id:'mb23', label:'Blend 2023',      max:3960, hex:'#C04050', stockUbic:{...vacio(),'R Peña':3960} },
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
// Fecha de hoy (yyyy-mm-dd) según el reloj del dispositivo. No usar toISOString(): da
// la fecha en UTC, que en Australia por la mañana es "ayer" y en Mendoza de noche "mañana".
function fechaLocalHoy(){
  const d=new Date();
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}

// ── CONEXIÓN CON EL BACKEND ──────────────────────────────────────────────
// El token de sesión lo da el servidor al validar el PIN (vale 12 h) y vive sólo en
// memoria: al cerrar o recargar la app se vuelve a pedir el PIN.
let authToken = '';
let alVencerSesion = null; // la app lo registra para volver al login si el servidor dice AUTH
const nuevoRequestId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 10);

// Si el backend contesta con error, se lanza: antes una respuesta {error} se mostraba
// como "registrado con éxito" porque nadie la miraba.
function revisarRespuesta(d) {
  if (d && d.error) {
    if (d.code === 'AUTH') { authToken = ''; if (alVencerSesion) alVencerSesion(); }
    const e = new Error(d.error); e.delServidor = true; throw e;
  }
  return d;
}
// Texto para el toast: si el error vino del servidor se muestra su mensaje (ej. "Stock insuficiente…").
const textoError = (e, base) => (e && e.delServidor ? e.message : base);

// Lecturas
async function gasGet(params) {
  if (!GAS_URL) return null;
  const q = { ...params }; if (authToken) q.token = authToken;
  const r = await fetch(GAS_URL + '?' + new URLSearchParams(q).toString(), { redirect: 'follow' });
  return revisarRespuesta(await r.json());
}

// Escrituras: POST (texto plano para no disparar el preflight de CORS) con un requestId
// que el servidor recuerda 6 h. Si la red falla a mitad de camino y se reintenta por GET,
// el servidor devuelve el resultado anterior en vez de cargar la operación dos veces.
async function gasPost(params, { soloPost = false } = {}) {
  if (!GAS_URL) return null;
  const cuerpo = { ...params, requestId: params.requestId || nuevoRequestId() };
  if (authToken) cuerpo.token = authToken;
  try {
    const r = await fetch(GAS_URL, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(cuerpo) });
    return revisarRespuesta(await r.json());
  } catch (e) {
    if (e instanceof TypeError && !soloPost) {
      const r = await fetch(GAS_URL + '?' + new URLSearchParams(cuerpo).toString(), { redirect: 'follow' });
      return revisarRespuesta(await r.json());
    }
    throw e;
  }
}

// Arma la lista de contactos para el selector, filtrada por el detalle elegido y
// ordenada por uso más reciente. contactosBalance viene como [{nombre,detalles:[{detalle,ultima}],ultimaFecha}].
// Si hay un detalle seleccionado, muestra SOLO los contactos ya usados con ese detalle
// (ordenados por su última fecha con ese detalle). Si no hay detalle elegido, muestra
// todos por uso reciente. Para un contacto nuevo para ese detalle siempre está "Nuevo…".
function contactosParaDetalle(contactosBalance, clientes, detalle) {
  const lista = Array.isArray(contactosBalance) ? contactosBalance : [];
  if (!detalle) {
    // Sin detalle: todos, por uso reciente, con los clientes de la pestaña al final si faltan.
    const orden = [...lista].sort((a, b) => (b.ultimaFecha || 0) - (a.ultimaFecha || 0)).map(c => c.nombre);
    const vistos = new Set(orden);
    (clientes || []).forEach(c => { if (c.nombre && !vistos.has(c.nombre)) { vistos.add(c.nombre); orden.push(c.nombre); } });
    return orden;
  }
  // Con detalle: sólo los que ya se usaron con ese detalle, por fecha de ese uso.
  return lista
    .map(c => {
      const match = (c.detalles || []).find(d => d.detalle === detalle);
      return match ? { nombre: c.nombre, orden: match.ultima } : null;
    })
    .filter(Boolean)
    .sort((a, b) => b.orden - a.orden)
    .map(x => x.nombre);
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

function useAppData(user) {
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
        else             { setPrice(null);         setSource('fallback'); } // sin dato real no se inventa un precio
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
    } catch (e) { if (e?.delServidor) { /* sesión vencida u otro error del servidor: lo maneja revisarRespuesta */ } setSource('fallback'); }
  };

  // Con servidor, los datos se piden recién cuando hay sesión (antes se pedían ya en la pantalla de login).
  useEffect(() => { if (!GAS_URL || user) refresh(); }, [user]);

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

// ── UI BASE ──────────────────────────────────────────────────────────────
const F=({label,children})=>(<div style={{display:'flex',flexDirection:'column',gap:6}}>{label&&<label style={{fontSize:13,color:C.muted,fontFamily:'system-ui'}}>{label}</label>}{children}</div>);
const iS={background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,padding:'12px 14px',minHeight:48,minWidth:0,maxWidth:'100%',color:C.text,fontSize:16,fontFamily:'system-ui',outline:'none',width:'100%',boxSizing:'border-box'};
// Los input de fecha de iPhone tienen un ancho propio que se desborda de su columna: se normalizan.
const Inp=p=><input style={p.type==='date'?{...iS,WebkitAppearance:'none',appearance:'none',display:'block'}:iS} {...p}/>;
const Sel=({children,...p})=><select style={{...iS,cursor:'pointer'}} {...p}>{children}</select>;
const Btn=({children,variant='primary',style:s,...p})=>(<button style={{background:variant==='primary'?C.gold:variant==='danger'?C.wine:C.cork,color:variant==='primary'?'#1A1612':C.text,border:`1px solid ${variant==='primary'?C.gold:variant==='danger'?C.wine:C.border}`,borderRadius:12,padding:'12px 20px',minHeight:52,fontSize:16,fontFamily:'system-ui',fontWeight:700,cursor:'pointer',...s}} {...p}>{children}</button>);
const Card=({children,style:s})=><div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:16,padding:16,...s}}>{children}</div>;
const SL=({children})=><div style={{fontSize:12,color:C.dim,letterSpacing:'0.12em',textTransform:'uppercase',fontFamily:'system-ui',fontWeight:600,marginBottom:10}}>{children}</div>;

// ── ÍCONOS, CHIPS Y AVATARES ─────────────────────────────────────────────
// Íconos de trazo (SVG en línea, sin dependencias). Se usan en lugar de emojis para que se vean
// igual en todos los celulares y respeten los colores de la app.
const ICONOS={
  inicio:'M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z',
  stock:'M21 8l-9-5-9 5v8l9 5 9-5zM3 8l9 5 9-5M12 13v8',
  mas:'M12 5v14M5 12h14',
  menos:'M5 12h14',
  caja:'M3 7h18v12H3zM3 11h18M16 15h2',
  datos:'M4 20V10M10 20V4M16 20v-7M22 20H2',
  botella:'M10 2h4v4l1.5 3v11a2 2 0 0 1-2 2h-3a2 2 0 0 1-2-2V9L10 6zM8.5 13h7',
  cobrar:'M12 2v20M17 6.5H9.5a3 3 0 0 0 0 6h5a3 3 0 0 1 0 6H6',
  mover:'M7 7h13l-4-4M17 17H4l4 4',
  alerta:'M12 3l10 18H2zM12 10v5M12 18v.01',
  reloj:'M12 7v5l3 2M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z',
  chevron:'M9 6l6 6-6 6',
  atras:'M15 6l-6 6 6 6',
  mic:'M12 15a3 3 0 0 0 3-3V6a3 3 0 0 0-6 0v6a3 3 0 0 0 3 3zM6 11a6 6 0 0 0 12 0M12 17v4',
  check:'M5 13l4 4L19 7',
  ajustes:'M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM19 12a7 7 0 0 0-.1-1.2l2-1.5-2-3.4-2.3.9a7 7 0 0 0-2-1.2L14.2 3h-4l-.4 2.6a7 7 0 0 0-2 1.2l-2.3-.9-2 3.4 2 1.5a7 7 0 0 0 0 2.4l-2 1.5 2 3.4 2.3-.9a7 7 0 0 0 2 1.2l.4 2.6h4l.4-2.6a7 7 0 0 0 2-1.2l2.3.9 2-3.4-2-1.5c.1-.4.1-.8.1-1.2z',
  candado:'M6 11h12v9H6zM8 11V8a4 4 0 0 1 8 0v3',
  salir:'M9 21H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h4M16 17l5-5-5-5M21 12H9',
};
function Ico({n,s=22,c='currentColor',w=1.8}){
  return(<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth={w} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{flexShrink:0}}><path d={ICONOS[n]||''}/></svg>);
}
// Botón-píldora de 44px para elegir una opción (reemplaza a los selects cuando hay pocas).
function Chip({activo,punteado,cuadrado,children,style:s,...p}){
  return(<button type="button" aria-pressed={activo?true:undefined} {...p} style={{minHeight:44,padding:'0 16px',borderRadius:cuadrado?12:22,cursor:'pointer',fontSize:14,fontFamily:'system-ui',fontWeight:activo?700:500,whiteSpace:'nowrap',
    background:activo?C.gold:'transparent',color:activo?'#1A1612':(punteado?C.gold:C.text),
    border:`1px ${punteado&&!activo?'dashed':'solid'} ${activo?C.gold:(punteado?C.goldDim:C.border)}`,...s}}>{children}</button>);
}
// Avatar: por defecto la inicial en dorado sobre fondo neutro (igual para todos). Cada usuario puede
// elegir un emoji o una foto. Se guarda sólo en ESTE celular (localStorage): no viaja al servidor ni a
// la planilla, así que no hay datos personales en Sheets, pero tampoco acompaña a otro dispositivo.
const EMOJIS_PERFIL=['🍷','🍇','🏔️','🎿','⛷️','🔥','🔪','🌿','☀️','🌙','⭐','🌵','🐶','🐱','🦊','🐻','🚀','⚽','🎸','📷','🍕','🚲','🌊','🏠'];
const claveAvatar=id=>`novato_avatar_${id}`;
function leerAvatar(id){ try{ const r=localStorage.getItem(claveAvatar(id)); return r?JSON.parse(r):null; }catch{ return null; } }
function guardarAvatar(id,v){
  try{ if(v) localStorage.setItem(claveAvatar(id),JSON.stringify(v)); else localStorage.removeItem(claveAvatar(id)); }catch{ return false; }
  window.dispatchEvent(new Event('novato-avatar')); return true;
}
// Recorta al cuadrado central y reduce a 160 px (JPEG): pesa unos pocos KB.
function reducirFoto(file,lado=160){
  return new Promise((ok,mal)=>{
    const url=URL.createObjectURL(file); const img=new Image();
    img.onload=()=>{
      const c=document.createElement('canvas'); c.width=c.height=lado;
      const m=Math.min(img.width,img.height);
      c.getContext('2d').drawImage(img,(img.width-m)/2,(img.height-m)/2,m,m,0,0,lado,lado);
      URL.revokeObjectURL(url); ok(c.toDataURL('image/jpeg',0.82));
    };
    img.onerror=()=>{ URL.revokeObjectURL(url); mal(new Error('imagen')); };
    img.src=url;
  });
}
function Avatar({u,s=36}){
  const [,refrescar]=useState(0);
  useEffect(()=>{ const f=()=>refrescar(n=>n+1); window.addEventListener('novato-avatar',f); return ()=>window.removeEventListener('novato-avatar',f); },[]);
  const a=leerAvatar(u.id);
  const base={width:s,height:s,borderRadius:'50%',background:C.cork,border:`1px solid ${C.goldDim}`,display:'inline-flex',alignItems:'center',justifyContent:'center',flexShrink:0,overflow:'hidden',boxSizing:'border-box'};
  if(a?.tipo==='foto'&&a.valor) return(<span aria-hidden="true" style={base}><img src={a.valor} alt="" style={{width:'100%',height:'100%',objectFit:'cover',display:'block'}}/></span>);
  if(a?.tipo==='emoji'&&a.valor) return(<span aria-hidden="true" style={{...base,fontSize:Math.round(s*0.52),lineHeight:1}}>{a.valor}</span>);
  return(<span aria-hidden="true" style={{...base,color:C.gold,fontFamily:'Georgia, serif',fontWeight:700,fontSize:Math.round(s*0.45)}}>{u.nombre.charAt(0)}</span>);
}

// ── LOGIN ────────────────────────────────────────────────────────────────
function LoginScreen({onLogin}){
  const [step,setStep]=useState('users');const [pending,setPending]=useState(null);const [pin,setPin]=useState('');const [shake,setShake]=useState(false);const [hov,setHov]=useState(null);
  const [checking,setChecking]=useState(false);const [msg,setMsg]=useState('');
  const fallar=(m)=>{setMsg(m||'PIN incorrecto');setShake(true);setPin('');setTimeout(()=>setShake(false),500);};
  // El PIN se valida en el servidor y éste devuelve un token de sesión. Si el servidor todavía no
  // tiene la autenticación configurada (AUTH_NOT_CONFIGURED) se usa el PIN local de antes.
  const tryPin=async()=>{
    if(checking) return;
    setChecking(true);setMsg('');
    try{
      if(GAS_URL){
        let r;
        try{ r=await gasPost({action:'login',userId:pending.id,pin},{soloPost:true}); }
        catch(e){ r={ok:false,code:'RED'}; }
        if(r?.ok){ onLogin(pending,r.token); return; }
        if(r?.code==='AUTH_NOT_CONFIGURED'){
          const s=await getPin(pending.id);
          if(pin===s){ onLogin(pending,''); return; }
          fallar(); return;
        }
        if(r?.code==='RED'){ fallar('Sin conexión con el servidor. Probá de nuevo.'); return; }
        fallar((r?.mensaje||'PIN incorrecto')+(r?.restantes!=null&&r.restantes<=2?` (te quedan ${r.restantes})`:''));
      } else {
        const s=await getPin(pending.id);
        if(pin===s) onLogin(pending,''); else fallar();
      }
    } finally { setChecking(false); }
  };
  useEffect(()=>{if(pin.length===4)tryPin();},[pin]);
  // Teclado físico (compu): sin un input oculto, para que el teléfono no abra su propio
  // teclado encima del teclado numérico de la app.
  useEffect(()=>{
    if(step!=='pin') return;
    const onKey=e=>{
      if(/^[0-9]$/.test(e.key)) setPin(p=>p.length<4?p+e.key:p);
      else if(e.key==='Backspace') setPin(p=>p.slice(0,-1));
    };
    window.addEventListener('keydown',onKey);
    return ()=>window.removeEventListener('keydown',onKey);
  },[step]);
  if(step==='pin') return(
    <div style={{minHeight:'100vh',background:C.cellar,display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:28}}>
      <button onClick={()=>{setStep('users');setPin('');}} aria-label="Volver" style={{position:'absolute',top:16,left:16,width:48,height:48,display:'flex',alignItems:'center',justifyContent:'center',background:'none',border:'none',cursor:'pointer'}}><Ico n="atras" s={26} c={C.muted} w={2}/></button>
      <div style={{textAlign:'center',marginBottom:36}}><div style={{marginBottom:12,display:'flex',justifyContent:'center'}}><Avatar u={pending} s={72}/></div><div style={{color:C.text,fontSize:20,fontFamily:'Georgia, serif',fontWeight:700}}>{pending.nombre}</div><div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',marginTop:4}}>Ingresá tu PIN</div></div>
      <div style={{display:'flex',gap:16,marginBottom:32,animation:shake?'shake 0.4s ease':'none'}}>
        {[0,1,2,3].map(i=><div key={i} style={{width:16,height:16,borderRadius:'50%',background:pin.length>i?C.gold:C.cork,border:`2px solid ${pin.length>i?C.gold:C.border}`,transition:'background 0.15s'}}/>)}
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,76px)',gap:12}}>
        {['1','2','3','4','5','6','7','8','9','','0','⌫'].map((k,i)=>(
          <button key={i} aria-label={k==='⌫'?'Borrar':(k||undefined)} onClick={()=>{if(k==='⌫')setPin(p=>p.slice(0,-1));else if(k)setPin(p=>p.length<4?p+k:p);}} style={{height:68,background:k?C.barrel:C.cellar,border:`1px solid ${k?C.border:'transparent'}`,borderRadius:12,color:C.text,fontSize:22,fontFamily:'Georgia, serif',cursor:k?'pointer':'default',opacity:k?1:0}}>{k}</button>
        ))}
      </div>
      <div style={{color:msg&&!checking?'#f08080':C.muted,fontSize:13,fontFamily:'system-ui',marginTop:28,minHeight:18,textAlign:'center'}}>{checking?'Verificando…':msg}</div>
      <style>{`@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}`}</style>
    </div>
  );
  return(
    <div style={{minHeight:'100vh',background:C.cellar,display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',padding:28}}>
      <div style={{textAlign:'center',marginBottom:40}}>
        <div style={{color:C.gold,fontSize:38,fontWeight:700,letterSpacing:'0.3em',fontFamily:'Georgia, serif'}}>NOVATO</div>
        <div style={{color:C.dim,fontSize:12,letterSpacing:'0.2em',fontFamily:'system-ui',marginTop:4}}>MENDOZA · ARGENTINA</div>
        <div style={{width:48,height:1,background:`linear-gradient(90deg,transparent,${C.gold},transparent)`,margin:'18px auto'}}/>
        <div style={{color:C.muted,fontSize:13,fontFamily:'system-ui'}}>¿Quién está operando?</div>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14,width:'100%',maxWidth:340}}>
        {USUARIOS.map(u=><button key={u.id} onClick={()=>{setPending(u);setPin('');setStep('pin');}} onMouseEnter={()=>setHov(u.id)} onMouseLeave={()=>setHov(null)} style={{background:C.barrel,border:`1px solid ${hov===u.id?C.gold:C.border}`,borderRadius:16,padding:'24px 16px',cursor:'pointer',textAlign:'center',transition:'border-color 0.2s'}}><div style={{marginBottom:10,display:'flex',justifyContent:'center'}}><Avatar u={u} s={52}/></div><div style={{color:C.text,fontWeight:700,fontSize:16,fontFamily:'Georgia, serif',marginBottom:4}}>{u.nombre}</div><div style={{color:C.muted,fontSize:12,fontFamily:'system-ui'}}>{u.rol}</div></button>)}
      </div>
      {!GAS_URL&&<div style={{marginTop:24,background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:12,color:'#E07080',maxWidth:340,textAlign:'center'}}>⚠ Sin conexión a Google Sheets — modo offline</div>}
    </div>
  );
}

// ── SETTINGS ─────────────────────────────────────────────────────────────
function SettingsPanel({user,onClose,onLogout,showToast}){
  const [step,setStep]=useState('menu');const [op,setOp]=useState('');const [np,setNp]=useState('');const [cp,setCp]=useState('');const [saving,setSaving]=useState(false);const archivo=useRef(null);
  const conServidor=!!authToken; // con sesión del servidor, el PIN se cambia en el servidor (y vale en todos tus dispositivos)
  const elegirEmoji=(e)=>{ if(guardarAvatar(user.id,{tipo:'emoji',valor:e})) showToast('Ícono actualizado'); else showToast('No se pudo guardar en este celular','error'); };
  const usarInicial=()=>{ guardarAvatar(user.id,null); showToast('Volviste a la inicial'); };
  const subirFoto=async(ev)=>{
    const f=ev.target.files?.[0]; ev.target.value='';
    if(!f) return;
    try{
      const dataUrl=await reducirFoto(f);
      if(guardarAvatar(user.id,{tipo:'foto',valor:dataUrl})) showToast('Foto actualizada'); else showToast('No se pudo guardar en este celular','error');
    }catch(e){ showToast('No pude leer esa imagen. Probá con otra.','error'); }
  };
  const changePin=async()=>{
    if(conServidor&&op.length!==4){showToast('Ingresá tu PIN actual','error');return;}
    if(np.length!==4){showToast('El PIN debe ser de 4 dígitos','error');return;}
    if(np!==cp){showToast('Los PINs no coinciden','error');return;}
    setSaving(true);
    try{
      if(conServidor){
        const r=await gasPost({action:'changePin',pinActual:op,pinNuevo:np},{soloPost:true});
        if(!r?.ok){showToast(r?.mensaje||'No se pudo cambiar el PIN','error');setSaving(false);return;}
      } else {
        const ok=await savePin(user.id,np);
        if(!ok){showToast('Error al guardar','error');setSaving(false);return;}
      }
      showToast('PIN actualizado');setStep('menu');setOp('');setNp('');setCp('');
    }catch(e){ showToast(textoError(e,'Error de conexión. Intentá de nuevo.'),'error'); }
    setSaving(false);
  };
  return(<>
    <div onClick={onClose} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.6)',zIndex:300}}/>
    <div style={{position:'fixed',bottom:0,left:0,right:0,background:C.barrel,border:`1px solid ${C.border}`,borderRadius:'20px 20px 0 0',padding:'8px 20px calc(40px + env(safe-area-inset-bottom, 0px))',zIndex:301,maxHeight:'90vh',overflowY:'auto'}}>
      <div style={{width:40,height:4,background:C.border,borderRadius:2,margin:'12px auto 20px'}}/>
      <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:24,paddingBottom:20,borderBottom:`1px solid ${C.border}`}}>
        <Avatar u={user} s={52}/>
        <div><div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif'}}>{user.nombre}</div><div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>{GAS_URL?'Conectado a Google Sheets':'Sin conexión a Google Sheets'}</div></div>
      </div>
      {step==='menu'?(<div style={{display:'flex',flexDirection:'column',gap:10}}><SL>Configuración</SL>
        <button onClick={()=>setStep('avatar')} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,padding:'14px 16px',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',color:C.text,fontFamily:'system-ui',fontSize:16,minHeight:52}}><span style={{display:'flex',alignItems:'center',gap:10}}><Ico n="botella" s={20} c={C.gold}/>Cambiar foto o ícono</span><Ico n="chevron" s={18} c={C.dim} w={2}/></button>
        <button onClick={()=>setStep('pin')} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,padding:'14px 16px',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',color:C.text,fontFamily:'system-ui',fontSize:16,minHeight:52}}><span style={{display:'flex',alignItems:'center',gap:10}}><Ico n="candado" s={20} c={C.gold}/>Cambiar PIN</span><Ico n="chevron" s={18} c={C.dim} w={2}/></button>
        <button onClick={onLogout} style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:12,padding:'14px 16px',cursor:'pointer',color:'#E07080',fontFamily:'system-ui',fontSize:16,minHeight:52,display:'flex',alignItems:'center',gap:10}}><Ico n="salir" s={20} c="#E07080"/>Cerrar sesión</button>
      </div>):step==='avatar'?(<div style={{display:'flex',flexDirection:'column',gap:14}}>
        <button onClick={()=>setStep('menu')} style={{background:'none',border:'none',color:C.muted,cursor:'pointer',fontSize:15,fontFamily:'system-ui',textAlign:'left',marginBottom:4,minHeight:44,display:'flex',alignItems:'center',gap:6}}><Ico n="atras" s={18} c={C.muted} w={2}/>Volver</button>
        <SL>Foto o ícono de sesión</SL>
        <div style={{display:'grid',gridTemplateColumns:'repeat(6,minmax(0,1fr))',gap:8}}>
          {EMOJIS_PERFIL.map(e=><button key={e} onClick={()=>elegirEmoji(e)} aria-label={`Ícono ${e}`} style={{height:48,fontSize:24,background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,cursor:'pointer',padding:0}}>{e}</button>)}
        </div>
        <input ref={archivo} type="file" accept="image/*" onChange={subirFoto} style={{display:'none'}}/>
        <Btn variant="ghost" onClick={()=>archivo.current?.click()}>Subir una foto</Btn>
        <Btn variant="ghost" onClick={usarInicial}>Usar mi inicial</Btn>
        <div style={{fontSize:13,color:C.dim,fontFamily:'system-ui',lineHeight:1.4}}>Se guarda sólo en este celular. En otro dispositivo se vuelve a elegir.</div>
      </div>):(<div style={{display:'flex',flexDirection:'column',gap:14}}>
        <button onClick={()=>setStep('menu')} style={{background:'none',border:'none',color:C.muted,cursor:'pointer',fontSize:15,fontFamily:'system-ui',textAlign:'left',marginBottom:4,minHeight:44,display:'flex',alignItems:'center',gap:6}}><Ico n="atras" s={18} c={C.muted} w={2}/>Volver</button>
        <SL>Cambiar PIN</SL>
        {conServidor&&<F label="PIN actual"><Inp type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} placeholder="••••" value={op} onChange={e=>setOp(e.target.value.replace(/\D/g,'').slice(0,4))}/></F>}
        <F label="Nuevo PIN (4 dígitos)"><Inp type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} placeholder="••••" value={np} onChange={e=>setNp(e.target.value.replace(/\D/g,'').slice(0,4))}/></F>
        <F label="Confirmar PIN"><Inp type="password" inputMode="numeric" pattern="[0-9]*" maxLength={4} placeholder="••••" value={cp} onChange={e=>setCp(e.target.value.replace(/\D/g,'').slice(0,4))}/></F>
        <Btn onClick={changePin} style={{opacity:saving?0.6:1}}>{saving?'Guardando…':'Guardar PIN'}</Btn>
      </div>)}
    </div>
  </>);
}

// ── STOCK (control físico + transferencias entre depósitos) ────────────────
function StockScreen({onBack,showToast,productos,applyTransfer,user,refresh,last,save,vistaInicial,borrador}){
  const [view,setView]=useState(vistaInicial||'control'); // 'control' | 'transfer'
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:16}}>
        <button onClick={onBack} aria-label="Volver" style={{width:44,height:44,display:'flex',alignItems:'center',justifyContent:'center',background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,cursor:'pointer'}}><Ico n="atras" s={20} c={C.muted} w={2}/></button>
        <div><h1 style={{margin:0,color:C.text,fontWeight:700,fontSize:20,fontFamily:'Georgia, serif'}}>Stock</h1><div style={{color:C.dim,fontSize:13,fontFamily:'system-ui'}}>Control físico y movimientos entre depósitos</div></div>
      </div>
      <div style={{display:'flex',marginBottom:20,border:`1px solid ${C.border}`,borderRadius:10,overflow:'hidden'}}>
        {[{k:'control',l:'Control físico'},{k:'transfer',l:'Mover entre depósitos'}].map((t,i)=>(
          <button key={t.k} onClick={()=>setView(t.k)} style={{flex:1,padding:'14px 8px',background:view===t.k?C.cork:C.barrel,color:view===t.k?C.gold:C.muted,border:'none',borderRight:i===0?`1px solid ${C.border}`:'none',cursor:'pointer',fontSize:14,fontFamily:'system-ui',fontWeight:view===t.k?700:400}}>{t.l}</button>
        ))}
      </div>
      {view==='control'
        ? <StockControlBody productos={productos} showToast={showToast} onBack={onBack} last={last} save={save} user={user} refresh={refresh}/>
        : <TransferForm productos={productos} applyTransfer={applyTransfer} user={user} showToast={showToast} onBack={onBack} refresh={refresh} borrador={borrador}/>}
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
        await gasPost({action:'addStockControl',deposito,items:JSON.stringify(items)}); // fecha, hora y teórico los fija el servidor
        await refresh();
      }
      await save(ctrl);
      if(hasDiff){showToast(`Control de ${deposito} guardado${GAS_URL?' en Sheets':''} — diferencias en ${diffs.filter(d=>d.diff!==0).length} producto(s)`,'error');}
      else{showToast(`${deposito}: stock coincide con lo teórico${GAS_URL?' (guardado en Sheets)':''}`,'ok');}
      onBack();
    }catch(e){ showToast(textoError(e,'Error al guardar el control. Intentá de nuevo.'),'error'); }
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
          {['Producto','Teórico','Real','Dif.'].map(h=><div key={h} style={{fontSize:12,color:C.dim,letterSpacing:'0.1em',textTransform:'uppercase',fontFamily:'system-ui',textAlign:h!=='Producto'?'center':'left'}}>{h}</div>)}
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
          <div style={{color:'#E07080',fontSize:14,fontWeight:700,marginBottom:8}}>Diferencias detectadas en {deposito}</div>
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
function TransferForm({productos,applyTransfer,user,showToast,onBack,refresh,borrador}){
  const hoy=fechaLocalHoy();
  // Si viene de la carga rápida, el formulario arranca completado (se revisa antes de registrar).
  const bt=borrador?.tipo==='transferencia'?borrador.transferencia:null;
  const [f,setF]=useState({producto:bt?(productos.find(p=>p.label===bt.producto)?.id||''):'',cantidad:bt?.cantidad?String(bt.cantidad):'',desde:bt?.desde||'',hacia:bt?.hacia||'',fecha:borrador?.fecha||hoy,notas:''});
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
      // El stock local se mueve recién cuando el servidor aceptó el movimiento: si falla (por ejemplo
      // por falta de stock real), el celular no queda mostrando un stock que no existe.
      if(GAS_URL) await gasPost({action:'addTransfer',fecha:f.fecha,producto:prod?.label,cantidad:cant,desde:f.desde,hacia:f.hacia,notas:f.notas||''});
      applyTransfer({productoId:f.producto,desde:f.desde,hacia:f.hacia,cantidad:cant});
      if(GAS_URL) await refresh();
      showToast(`${cant} bot de ${prod?.label}: ${f.desde} → ${f.hacia}${GAS_URL?' (enviado a Sheets)':''}`,'ok');
      onBack();
    }catch(e){ showToast(textoError(e,'Error al registrar el movimiento.'),'error'); }
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
function DashboardScreen({user,onNavigate,price,source,productos,last,operacionesPendientes,resumenCajas}){
  const {tc,err:tcErr}=useDolarBlue();
  // Detalle de una operación (se abre al tocar una operación pendiente)
  const [detalle,setDetalle]=useState(null);        // null=cerrado, {}=cargando, {...}=datos
  const [filtro,setFiltro]=useState('todas');       // todas | cobrar | pagar
  const [verTodas,setVerTodas]=useState(false);
  const abrirDetalle=async(referencia)=>{
    if(!referencia||!GAS_URL) return;
    setDetalle({cargando:true,referencia});
    try{
      const d=await gasGet({action:'getDetalleOperacion',referencia});
      setDetalle(d&&d.referencia?d:{error:true,referencia});
    }catch(e){ setDetalle({error:true,referencia}); }
  };
  const total=productos.reduce((s,p)=>s+totalStock(p),0);
  const productosOrdenados=[...productos].filter(p=>totalStock(p)>0).sort((a,b)=>totalStock(b)-totalStock(a)); // solo con stock, mayor a menor
  const ars=price?total*price:null; const usd=ars&&tc?Math.round(ars/tc):null;
  const totalCajaArs=Math.round(resumenCajas?.reduce((s,c)=>s+c.ars,0)||0);
  const totalCajaUsd=Math.round(resumenCajas?.reduce((s,c)=>s+c.usd,0)||0);
  const daysSince=last?(()=>{try{const [d,m,y]=last.fecha.split('/');return Math.floor((Date.now()-new Date(`${y}-${m}-${d}`).getTime())/86400000);}catch{return null;}})():null;
  const diffsControl=(last?.items||[]).filter(d=>d.diff!==0);

  // Operaciones pendientes con sus días, signo y si están vencidas (venta sin cobrar hace más de 60 días).
  const ops=(operacionesPendientes||[]).map(op=>{
    const dias=op.fecha?Math.floor((Date.now()-new Date(op.fecha).getTime())/86400000):null;
    const esVenta=op.tipo==='venta';
    const pendiente=esVenta?op.saldoArs:-op.saldoArs; // positivo = nos deben (venta) / debemos (compra)
    return {...op,dias,esVenta,pendiente,esCredito:pendiente<0,vencida:esVenta&&dias!==null&&dias>60};
  });
  let porCobrar=0, porPagar=0;
  ops.forEach(op=>{ if(op.esVenta) porCobrar+=op.saldoArs; else porPagar+=(-op.saldoArs); });
  const neto=porCobrar-porPagar;
  const baseBarra=Math.max(Math.abs(porCobrar)+Math.abs(porPagar),1);
  const pctCobrar=Math.round(Math.abs(porCobrar)/baseBarra*100);
  const opsFiltradas=ops.filter(op=>filtro==='todas'||(filtro==='cobrar'&&op.esVenta)||(filtro==='pagar'&&!op.esVenta));
  const opsVisibles=verTodas?opsFiltradas:opsFiltradas.slice(0,4);
  const vencidas=ops.filter(o=>o.vencida).sort((a,b)=>b.dias-a.dias).slice(0,2);
  const controlAtrasado=daysSince===null||daysSince>30;
  const hayAlertas=controlAtrasado||vencidas.length>0||diffsControl.length>0;
  const hoy=(()=>{const t=new Date().toLocaleDateString('es-AR',{weekday:'long',day:'numeric',month:'long'});return t.charAt(0).toUpperCase()+t.slice(1);})();
  const sinc=source==='sheets';

  const Titulo=({children,derecha})=>(<div style={{display:'flex',alignItems:'baseline',justifyContent:'space-between',marginBottom:10}}><h2 style={{margin:0,fontSize:12,letterSpacing:'0.12em',textTransform:'uppercase',color:C.dim,fontWeight:600,fontFamily:'system-ui'}}>{children}</h2>{derecha&&<span style={{fontSize:13,color:C.dim,fontFamily:'system-ui'}}>{derecha}</span>}</div>);
  const Alerta=({tono,icono,titulo,sub,accion,onAccion})=>{
    const t=tono==='rojo'?{bg:'#2A1216',bd:'#7A2A36',fg:'#F08A8A'}:{bg:'#2A1D0E',bd:'#6B4E1F',fg:'#E6B566'};
    return(<div style={{display:'flex',alignItems:'center',gap:12,background:t.bg,border:`1px solid ${t.bd}`,borderRadius:14,padding:'12px 14px'}}>
      <Ico n={icono} c={t.fg}/>
      <div style={{flex:1,minWidth:0,fontFamily:'system-ui'}}><div style={{fontSize:14,fontWeight:600,color:C.text}}>{titulo}</div><div style={{fontSize:13,color:t.fg,marginTop:2}}>{sub}</div></div>
      <button onClick={onAccion} style={{minHeight:44,padding:'0 14px',background:C.cork,border:`1px solid ${t.bd}`,borderRadius:10,color:t.fg,fontSize:14,fontWeight:700,cursor:'pointer',fontFamily:'system-ui'}}>{accion}</button>
    </div>);
  };

  return(
    <div style={{padding:'20px 16px',maxWidth:540,margin:'0 auto',display:'flex',flexDirection:'column',gap:20}}>
      <div style={{fontFamily:'system-ui'}}>
        <div style={{fontSize:13,color:C.dim}}>{hoy}</div>
        <h1 style={{margin:'2px 0 0',fontFamily:'Georgia, serif',fontSize:26,fontWeight:700}}>Hola, {user.nombre}</h1>
      </div>

      {/* Plata en cajas */}
      <section aria-label="Plata en cajas" style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:18,padding:18,display:'flex',flexDirection:'column',gap:14}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',fontFamily:'system-ui'}}>
          <span style={{fontSize:12,letterSpacing:'0.12em',textTransform:'uppercase',color:C.dim}}>Plata en cajas</span>
          <span style={{display:'flex',alignItems:'center',gap:6,fontSize:12,color:C.muted}}><span style={{width:7,height:7,borderRadius:'50%',background:sinc?'#7DCE9B':'#E6B566'}}/>{sinc?'Planilla al día':(source==='loading'?'Cargando…':'Sin conexión con la planilla')}</span>
        </div>
        <div style={{display:'flex',alignItems:'baseline',gap:12,flexWrap:'wrap'}}>
          <span style={{fontFamily:'Georgia, serif',fontSize:36,fontWeight:700,color:C.gold,lineHeight:1}}>${totalCajaArs.toLocaleString('es-AR')}</span>
          <span style={{fontFamily:'Georgia, serif',fontSize:20,fontWeight:700,color:C.muted}}>US$ {totalCajaUsd.toLocaleString('es-AR')}</span>
        </div>
        <div style={{height:1,background:'#2A241D'}}/>
        <div style={{display:'flex',justifyContent:'space-between',flexWrap:'wrap',gap:6,fontSize:13,color:C.muted,fontFamily:'system-ui'}}>
          <span>Dólar blue <strong style={{color:C.text}}>{tc?`$${tc.toLocaleString('es-AR')}`:(tcErr?'no disponible':'…')}</strong></span>
          <span>Precio ref. <strong style={{color:C.text}}>{price?`$${price.toLocaleString('es-AR')}/bot`:(source==='loading'?'…':'sin dato')}</strong></span>
        </div>
      </section>

      {/* Acciones rápidas */}
      <section aria-label="Acciones rápidas" style={{display:'grid',gridTemplateColumns:'repeat(3,minmax(0,1fr))',gap:10}}>
        {[{ico:'botella',l:'Vender',ir:'venta',pri:true},{ico:'cobrar',l:'Cobrar / pagar',ir:'caja'},{ico:'mover',l:'Mover stock',ir:'stock-mover'}].map(a=>(
          <button key={a.l} onClick={()=>onNavigate(a.ir)} style={{display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:8,minHeight:84,background:a.pri?C.gold:C.cork,border:a.pri?'none':`1px solid ${C.border}`,borderRadius:16,color:a.pri?'#1A1612':C.text,fontSize:14,fontWeight:a.pri?700:600,cursor:'pointer',fontFamily:'system-ui'}}>
            <Ico n={a.ico} s={24} c={a.pri?'#1A1612':C.gold}/>{a.l}
          </button>
        ))}
      </section>

      {/* Para resolver */}
      <section aria-label="Para resolver">
        <Titulo>Para resolver</Titulo>
        <div style={{display:'flex',flexDirection:'column',gap:10}}>
          {controlAtrasado&&<Alerta tono="ambar" icono="alerta" titulo={daysSince===null?'Control de stock pendiente':`Control de stock hace ${daysSince} días`} sub={last?`Último: ${last.deposito?last.deposito+' · ':''}${last.fecha}`:'Nunca realizado'} accion="Controlar" onAccion={()=>onNavigate('stock')}/>}
          {vencidas.map(op=><Alerta key={op.referencia} tono="rojo" icono="reloj" titulo={`${op.contraparte||'(sin nombre)'} lleva ${op.dias} días sin pagar`} sub={`Saldo $${Math.abs(op.pendiente).toLocaleString('es-AR')}`} accion="Ver" onAccion={()=>abrirDetalle(op.referencia)}/>)}
          {diffsControl.length>0&&(
            <div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:14,padding:'12px 14px',fontFamily:'system-ui'}}>
              <div style={{fontSize:14,fontWeight:600,color:C.text,marginBottom:8}}>Diferencias del último control{last?.deposito?` · ${last.deposito}`:''}</div>
              {diffsControl.map((d,i)=>(
                <div key={i} style={{display:'flex',alignItems:'center',gap:10,padding:'4px 0'}}>
                  <div style={{width:3,height:28,background:d.hex||C.wine,borderRadius:2,flexShrink:0}}/>
                  <div style={{flex:1}}><div style={{fontSize:14,color:C.text,fontWeight:600}}>{d.label}</div><div style={{fontSize:13,color:C.dim}}>Teórico {d.stock} · Real {d.real}</div></div>
                  <div style={{fontWeight:700,fontSize:15,color:d.diff<0?'#f08080':'#7dce9b'}}>{d.diff>0?'+':''}{d.diff} bot</div>
                </div>
              ))}
            </div>
          )}
          {!hayAlertas&&(<div style={{display:'flex',alignItems:'center',gap:10,background:C.greenBg,border:`1px solid ${C.green}55`,borderRadius:14,padding:'12px 14px',fontFamily:'system-ui',fontSize:14,color:'#7DCE9B'}}><Ico n="check" c="#7DCE9B"/>Todo en orden: control de stock al día y sin cobros vencidos.</div>)}
        </div>
      </section>

      {/* Cuentas abiertas */}
      {ops.length>0&&(
        <section aria-label="Cuentas abiertas" style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:18,padding:18,display:'flex',flexDirection:'column',gap:14}}>
          <h2 style={{margin:0,fontSize:12,letterSpacing:'0.12em',textTransform:'uppercase',color:C.dim,fontWeight:600,fontFamily:'system-ui'}}>Cuentas abiertas</h2>
          <div style={{display:'flex',height:10,borderRadius:5,overflow:'hidden',gap:3}} role="img" aria-label={`Por cobrar ${pctCobrar} por ciento, por pagar ${100-pctCobrar} por ciento`}>
            <div style={{flex:Math.max(pctCobrar,1),background:'#7DCE9B'}}/><div style={{flex:Math.max(100-pctCobrar,1),background:'#F08A8A'}}/>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(2,minmax(0,1fr))',gap:12,fontFamily:'system-ui'}}>
            <div><div style={{fontSize:13,color:C.muted}}>Por cobrar</div><div style={{fontFamily:'Georgia, serif',fontSize:20,fontWeight:700,color:'#7DCE9B'}}>+${Math.round(porCobrar).toLocaleString('es-AR')}</div></div>
            <div><div style={{fontSize:13,color:C.muted}}>Por pagar</div><div style={{fontFamily:'Georgia, serif',fontSize:20,fontWeight:700,color:'#F08A8A'}}>-${Math.round(porPagar).toLocaleString('es-AR')}</div></div>
          </div>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',paddingTop:12,borderTop:'1px solid #2A241D',fontFamily:'system-ui'}}>
            <span style={{fontSize:14,fontWeight:700,color:C.gold}}>Neto</span>
            <span style={{fontFamily:'Georgia, serif',fontSize:22,fontWeight:700,color:neto>=0?'#7DCE9B':'#F08A8A'}}>{neto>=0?'+':'-'}${Math.abs(Math.round(neto)).toLocaleString('es-AR')}</span>
          </div>
        </section>
      )}

      {/* Operaciones pendientes */}
      <section aria-label="Operaciones pendientes">
        <Titulo derecha={`${ops.length} abierta${ops.length===1?'':'s'}`}>Operaciones pendientes</Titulo>
        {ops.length>0&&(
          <div style={{display:'flex',gap:8,marginBottom:10,flexWrap:'wrap'}}>
            {[['todas','Todas'],['cobrar','Por cobrar'],['pagar','Por pagar']].map(([k,l])=><Chip key={k} activo={filtro===k} onClick={()=>{setFiltro(k);setVerTodas(false);}}>{l}</Chip>)}
          </div>
        )}
        <div style={{display:'flex',flexDirection:'column',gap:8}}>
          {opsVisibles.map((op,i)=>(
            <button key={op.referencia+'-'+i} onClick={()=>abrirDetalle(op.referencia)} style={{display:'flex',alignItems:'center',gap:12,width:'100%',minHeight:68,textAlign:'left',background:C.barrel,border:`1px solid ${op.vencida?'#7A2A36':C.border}`,borderRadius:14,padding:'12px 14px',color:C.text,cursor:'pointer',fontFamily:'system-ui'}}>
              <span style={{width:4,alignSelf:'stretch',borderRadius:2,background:op.esCredito?'#7FB2E0':(op.esVenta?'#7DCE9B':'#F08A8A'),flexShrink:0}}/>
              <span style={{flex:1,minWidth:0,display:'flex',flexDirection:'column',gap:3}}>
                <span style={{fontSize:15,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{op.contraparte||'(sin nombre)'}{op.producto?` · ${op.producto}`:''}</span>
                <span style={{fontSize:13,color:op.vencida?'#F08A8A':C.muted,fontWeight:op.vencida?600:400}}>{op.vencida?'Vencida · ':''}{op.detalle}{op.dias!==null?` · hace ${op.dias} día${op.dias===1?'':'s'}`:''}</span>
              </span>
              <span style={{textAlign:'right',whiteSpace:'nowrap'}}>
                <span style={{display:'block',fontSize:15,fontWeight:700,color:op.esCredito?'#7FB2E0':(op.esVenta?'#7DCE9B':'#F08A8A')}}>{op.esCredito?'':(op.esVenta?'+':'-')}${Math.abs(op.pendiente).toLocaleString('es-AR')}</span>
                {op.esCredito&&<span style={{display:'block',fontSize:12,color:'#7FB2E0'}}>a favor</span>}
              </span>
              <Ico n="chevron" s={18} c={C.dim} w={2}/>
            </button>
          ))}
          {ops.length===0&&<div style={{color:C.dim,fontSize:14,fontFamily:'system-ui',textAlign:'center',padding:'20px 0'}}>Sin operaciones pendientes</div>}
          {ops.length>0&&opsFiltradas.length===0&&<div style={{color:C.dim,fontSize:14,fontFamily:'system-ui',textAlign:'center',padding:'12px 0'}}>No hay operaciones en este filtro</div>}
          {opsFiltradas.length>4&&(
            <button onClick={()=>setVerTodas(v=>!v)} style={{minHeight:48,background:'none',border:`1px dashed ${C.border}`,borderRadius:12,color:C.gold,fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:'system-ui'}}>{verTodas?'Ver menos':`Ver las ${opsFiltradas.length} operaciones`}</button>
          )}
        </div>
      </section>

      {/* Stock */}
      <section aria-label="Stock" style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:18,padding:18,display:'flex',flexDirection:'column',gap:16}}>
        <div style={{display:'flex',alignItems:'baseline',justifyContent:'space-between',gap:8,flexWrap:'wrap',fontFamily:'system-ui'}}>
          <h2 style={{margin:0,fontSize:12,letterSpacing:'0.12em',textTransform:'uppercase',color:C.dim,fontWeight:600}}>Stock</h2>
          <span style={{fontSize:13,color:C.muted}}><strong style={{color:C.text,fontFamily:'Georgia, serif',fontSize:18}}>{total.toLocaleString('es-AR')}</strong> botellas{usd?` · US$ ${usd.toLocaleString('es-AR')}`:''}</span>
        </div>
        {productosOrdenados.map(p=>{
          const stock=totalStock(p); const pct=p.max>0?Math.min(stock/p.max,1):0;
          return(
            <div key={p.id} style={{display:'flex',flexDirection:'column',gap:8,fontFamily:'system-ui'}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',fontSize:14}}>
                <span style={{display:'flex',alignItems:'center',gap:8}}><span style={{width:10,height:10,borderRadius:3,background:p.hex}}/>{p.label}</span>
                <strong style={{fontFamily:'Georgia, serif',fontSize:16}}>{stock.toLocaleString('es-AR')}</strong>
              </div>
              <div style={{height:8,borderRadius:4,background:C.cork,overflow:'hidden'}}><div style={{width:`${Math.max(pct*100,2)}%`,height:'100%',background:p.hex,transition:'width 0.8s cubic-bezier(.4,0,.2,1)'}}/></div>
              <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                {UBICACIONES.filter(u=>(p.stockUbic[u]||0)>0).map(u=><span key={u} style={{fontSize:13,color:C.muted,background:C.cork,borderRadius:8,padding:'3px 10px'}}>{u} <strong style={{color:C.text}}>{p.stockUbic[u].toLocaleString('es-AR')}</strong></span>)}
              </div>
            </div>
          );
        })}
        {productosOrdenados.length===0&&<div style={{color:C.dim,fontSize:14,fontFamily:'system-ui',textAlign:'center'}}>Sin stock disponible</div>}
      </section>

      {detalle&&(
        <div onClick={()=>setDetalle(null)} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.75)',zIndex:250,display:'flex',alignItems:'flex-end',justifyContent:'center'}}>
          <div onClick={e=>e.stopPropagation()} style={{background:C.cork,border:`1px solid ${C.border}`,borderRadius:'16px 16px 0 0',padding:'18px 16px calc(28px + env(safe-area-inset-bottom, 0px))',width:'100%',maxWidth:500,maxHeight:'88vh',overflowY:'auto',WebkitOverflowScrolling:'touch'}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:14}}>
              <div style={{color:C.gold,fontSize:16,fontFamily:'Georgia, serif',fontWeight:700}}>Detalle de operación</div>
              <button onClick={()=>setDetalle(null)} aria-label="Cerrar" style={{background:'none',border:'none',color:C.muted,fontSize:26,cursor:'pointer',lineHeight:1,padding:'0 8px',minHeight:44}}>×</button>
            </div>
            {detalle.cargando&&<div style={{color:C.dim,fontSize:14,fontFamily:'system-ui',padding:'20px 0',textAlign:'center'}}>Cargando…</div>}
            {detalle.error&&<div style={{color:'#f08080',fontSize:14,fontFamily:'system-ui',padding:'20px 0',textAlign:'center'}}>No se pudo cargar el detalle.</div>}

            {!detalle.cargando&&!detalle.error&&detalle.lineas&&(
              <>
                <div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10,padding:'10px 12px',marginBottom:14,fontFamily:'system-ui'}}>
                  <div style={{color:C.text,fontSize:14,fontWeight:700}}>{detalle.cabecera?.contacto||'(sin nombre)'}</div>
                  <div style={{color:C.dim,fontSize:12,marginTop:3}}>
                    {detalle.cabecera?.detalle} · {detalle.cabecera?.fecha} · ref {detalle.referencia}
                    {detalle.cabecera?.user?` · cargó ${detalle.cabecera.user}`:''}
                  </div>
                </div>

                <div style={{color:C.muted,fontSize:12,fontFamily:'system-ui',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:6}}>Productos</div>
                {detalle.lineas.map((l,i)=>(
                  <div key={i} style={{padding:'8px 2px',borderTop:`1px solid ${C.border}`,fontFamily:'system-ui'}}>
                    <div style={{display:'flex',justifyContent:'space-between',gap:8}}>
                      <span style={{color:C.text,fontSize:13,fontWeight:600}}>{l.producto||'(sin producto)'}</span>
                      <span style={{color:C.text,fontSize:13,fontWeight:700,whiteSpace:'nowrap'}}>${Math.abs(l.montoArs).toLocaleString('es-AR')}</span>
                    </div>
                    <div style={{color:C.dim,fontSize:12,marginTop:2}}>
                      {l.botellas?`${l.botellas} bot`:'sin botellas'}
                      {l.precioUnit?` · $${l.precioUnit.toLocaleString('es-AR')}/bot`:''}
                      {l.deposito?` · desde ${l.deposito}`:''}
                      {l.montoUsd?` · US$${Math.abs(l.montoUsd).toLocaleString('es-AR')}`:''}
                    </div>
                  </div>
                ))}
                <div style={{display:'flex',justifyContent:'space-between',padding:'8px 2px',borderTop:`2px solid ${C.border}`,fontFamily:'system-ui',fontSize:13}}>
                  <span style={{color:C.muted}}>Total{detalle.totalBotellas?` (${detalle.totalBotellas} bot)`:''}</span>
                  <span style={{color:C.text,fontWeight:700}}>${Math.abs(detalle.totalArs).toLocaleString('es-AR')}</span>
                </div>

                <div style={{color:C.muted,fontSize:12,fontFamily:'system-ui',textTransform:'uppercase',letterSpacing:'0.05em',margin:'16px 0 6px'}}>
                  {detalle.totalArs>=0?'Cobros':'Pagos'} ({detalle.pagos.length})
                </div>
                {detalle.pagos.length===0&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'6px 2px'}}>Todavía no se registró ninguno.</div>}
                {detalle.pagos.map((p,i)=>(
                  <div key={i} style={{display:'flex',justifyContent:'space-between',gap:8,padding:'7px 2px',borderTop:`1px solid ${C.border}`,fontFamily:'system-ui'}}>
                    <div style={{minWidth:0}}>
                      <div style={{color:C.text,fontSize:12,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{p.fecha} · {p.caja}</div>
                      <div style={{color:C.dim,fontSize:12,marginTop:1}}>{p.producto||p.detalle}{p.user?` · ${p.user}`:''}</div>
                    </div>
                    <span style={{color:'#7dce9b',fontSize:12,fontWeight:700,whiteSpace:'nowrap'}}>${Math.abs(p.montoArs).toLocaleString('es-AR')}</span>
                  </div>
                ))}

                <div style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:10,padding:'10px 12px',marginTop:16,fontFamily:'system-ui'}}>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:C.muted,padding:'2px 0'}}>
                    <span>Facturado</span><span>${Math.abs(detalle.totalArs).toLocaleString('es-AR')}</span>
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:C.muted,padding:'2px 0'}}>
                    <span>{detalle.totalArs>=0?'Cobrado':'Pagado'}</span><span>${Math.abs(detalle.pagadoArs).toLocaleString('es-AR')}</span>
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:14,fontWeight:700,padding:'6px 0 0',borderTop:`1px solid ${C.border}`,marginTop:4}}>
                    <span style={{color:C.gold}}>Saldo pendiente</span>
                    <span style={{color:Math.abs(detalle.saldoArs)<1?'#7dce9b':'#f0c674'}}>${Math.abs(detalle.saldoArs).toLocaleString('es-AR')}</span>
                  </div>
                  {!!detalle.saldoUsd&&<div style={{color:C.dim,fontSize:12,textAlign:'right',marginTop:3}}>US${Math.abs(detalle.saldoUsd).toLocaleString('es-AR')}</div>}
                  {detalle.dias!==null&&detalle.dias!==undefined&&(
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:12,padding:'7px 0 0',borderTop:`1px solid ${C.border}`,marginTop:6}}>
                      <span style={{color:C.muted}}>{detalle.cerrada?(detalle.totalArs>=0?'Tardó en cobrarse':'Tardó en pagarse'):'Lleva abierta'}</span>
                      <span style={{color:detalle.cerrada?'#7dce9b':(detalle.dias>60?'#f08080':C.text),fontWeight:700}}>{detalle.dias} día{detalle.dias===1?'':'s'}</span>
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── CARGA RÁPIDA (texto o voz → borrador del formulario) ────────────────
// El servidor (interpretarCarga) devuelve un BORRADOR; nada se registra hasta que el usuario
// revisa el formulario y confirma. Por voz: el micrófono del propio teclado del celular anda
// siempre; además, si el navegador lo permite, se ofrece un botón propio.
function CargaRapida({onResultado}){
  const [texto,setTexto]=useState('');
  const [cargando,setCargando]=useState(false);
  const [error,setError]=useState('');
  const [avisos,setAvisos]=useState([]);
  const [escuchando,setEscuchando]=useState(false);
  const rec=useRef(null);
  // En iPhone, la app agregada a la pantalla de inicio no entrega el dictado del navegador (el micrófono
  // se enciende pero nunca llega el texto), así que ahí no se ofrece el botón propio: se usa el micrófono
  // del teclado, que anda siempre.
  const esIosInstalada=typeof navigator!=='undefined'&&(navigator.standalone===true||(/iPhone|iPad|iPod/.test(navigator.userAgent)&&window.matchMedia?.('(display-mode: standalone)').matches));
  const SR=typeof window!=='undefined'&&!esIosInstalada&&(window.SpeechRecognition||window.webkitSpeechRecognition);
  useEffect(()=>()=>{ try{rec.current?.stop();}catch{} },[]);
  if(!GAS_URL) return null;

  const dictar=()=>{
    if(!SR) return;
    if(escuchando){ try{rec.current?.stop();}catch{} return; }
    try{
      const base=texto.trim();
      const r=new SR(); r.lang='es-AR'; r.interimResults=true; r.continuous=false;
      setError('');
      // Se va mostrando lo que se entiende (parcial incluido), así un corte no pierde lo dictado.
      r.onresult=e=>{ const t=Array.from(e.results).map(x=>x[0].transcript).join(' ').trim(); if(t) setTexto(((base?base+' ':'')+t).slice(0,500)); };
      r.onend=()=>setEscuchando(false);
      r.onerror=e=>{ setEscuchando(false); setError(e?.error==='not-allowed'||e?.error==='service-not-allowed'?'No hay permiso para usar el micrófono. Escribí la frase o usá el micrófono del teclado.':(e?.error==='no-speech'?'No escuché nada. Probá de nuevo.':'No pude dictar. Escribí la frase o usá el micrófono del teclado.')); };
      rec.current=r; setEscuchando(true); r.start();
    }catch{ setEscuchando(false); setError('No pude dictar. Escribí la frase o usá el micrófono del teclado.'); }
  };
  const interpretar=async()=>{
    const t=texto.trim();
    if(!t||cargando) return;
    setCargando(true); setError(''); setAvisos([]);
    try{
      const r=await gasPost({action:'interpretarCarga',texto:t},{soloPost:true});
      if(!r?.ok){ setError(r?.mensaje||'No se pudo interpretar. Cargalo a mano.'); }
      else if(r.propuesta.tipo==='no_entiendo'){ setError('No entendí qué operación es.'); setAvisos(r.propuesta.avisos||[]); }
      else { setAvisos(r.propuesta.avisos||[]); onResultado(r.propuesta); }
    }catch(e){ setError(textoError(e,'No pude conectarme. Cargalo a mano o probá de nuevo.')); }
    setCargando(false);
  };
  return(
    <section aria-label="Carga rápida" style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:16,padding:14,display:'flex',flexDirection:'column',gap:10}}>
      <span style={{fontSize:12,letterSpacing:'0.12em',textTransform:'uppercase',color:C.gold,fontWeight:700,fontFamily:'system-ui'}}>Carga rápida</span>
      <div style={{display:'flex',gap:8,alignItems:'stretch'}}>
        <textarea value={texto} onChange={e=>setTexto(e.target.value.slice(0,500))} rows={2} aria-label="Contá qué querés cargar" placeholder="Ej: Vendí 12 Malbec 2022 a Organyca a $9.000 cada una" style={{...iS,flex:1,resize:'none',lineHeight:1.4,minHeight:56}}/>
        {SR&&<button onClick={dictar} aria-label={escuchando?'Dejar de escuchar':'Dictar'} style={{width:56,flexShrink:0,display:'flex',alignItems:'center',justifyContent:'center',background:escuchando?C.wine:C.cork,border:`1px solid ${escuchando?C.wine:C.border}`,borderRadius:12,cursor:'pointer'}}><Ico n="mic" c={escuchando?'#fff':C.gold}/></button>}
      </div>
      <Btn onClick={interpretar} disabled={!texto.trim()||cargando} style={{opacity:(!texto.trim()||cargando)?0.5:1}}>{cargando?'Interpretando…':'Completar el formulario'}</Btn>
      {!SR&&<span style={{fontSize:13,color:C.dim,fontFamily:'system-ui'}}>Para dictar, tocá el micrófono del teclado de tu celular.</span>}
      {error&&<div role="alert" style={{fontSize:14,color:'#f08080',fontFamily:'system-ui',lineHeight:1.4}}>{error}</div>}
      {avisos.length>0&&<div style={{fontSize:13,color:'#E6B566',fontFamily:'system-ui',lineHeight:1.5}}>{avisos.map((a,i)=><div key={i}>• {a}</div>)}</div>}
      {!error&&<span style={{fontSize:13,color:C.dim,fontFamily:'system-ui',lineHeight:1.4}}>Arma el formulario con lo que digas. Siempre lo revisás antes de registrar.</span>}
    </section>
  );
}

// ── NUEVA VENTA / TRANSACCIÓN ────────────────────────────────────────────
function VentaScreen({user,onBack,showToast,addOp,price,productos,applySale,clientes,categorias,contactosBalance,refresh,onIrA}){
  const hoy=fechaLocalHoy();
  const lineaVacia={producto:'',deposito:'',botellas:'',monto:''};
  const [f,setF]=useState({detalle:'Venta',contacto:'',contactoNuevo:'',canal:'',fecha:hoy,lineas:[{...lineaVacia}]});
  const [sending,setSending]=useState(false);
  const [verDetalles,setVerDetalles]=useState(false);
  const [verContactos,setVerContactos]=useState(false);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const setLinea=(idx,campo,valor)=>setF(p=>{
    const lineas=[...p.lineas];lineas[idx]={...lineas[idx],[campo]:valor};
    // Al elegir producto: si sólo hay un depósito con stock, queda elegido solo.
    if(campo==='producto'){
      const prod=productos.find(x=>x.id===valor);
      const con=prod?UBICACIONES.filter(u=>(prod.stockUbic[u]||0)>0):[];
      lineas[idx].deposito=con.length===1?con[0]:'';
    }
    return{...p,lineas};
  });
  const agregarLinea=()=>setF(p=>({...p,lineas:[...p.lineas,{...lineaVacia}]}));
  const quitarLinea=(idx)=>setF(p=>({...p,lineas:p.lineas.filter((_,i)=>i!==idx)}));

  const esVenta=f.detalle==='Venta';
  const catInfo=categorias?.find(c=>c.detalle===f.detalle);
  const tipoBadge=catInfo?.tipo||'Ingreso';

  const listaDetalles = categorias?.length ? categorias.map(c=>c.detalle) : ['Venta'];
  const listaContactos = contactosParaDetalle(contactosBalance, clientes, f.detalle);
  const clienteInfo = esVenta ? clientes?.find(c=>c.nombre===f.contacto) : null;
  const setContacto = (nombre) => {
    setF(p=>({...p, contacto:nombre, canal: clientes?.find(c=>c.nombre===nombre)?.canal || p.canal }));
  };
  // Chips: los más usados con este detalle (o el elegido, aunque no esté entre ellos)
  const chipsDetalle = listaDetalles.length<=6||verDetalles ? listaDetalles : [...new Set([f.detalle,...listaDetalles.slice(0,5)])];
  const chipsContacto = verContactos ? listaContactos : [...new Set([...(f.contacto&&f.contacto!=='Nuevo…'?[f.contacto]:[]),...listaContactos.slice(0,5)])];
  const opcionesContacto = f.contacto&&f.contacto!=='Nuevo…'&&!listaContactos.includes(f.contacto) ? [f.contacto,...listaContactos] : listaContactos;
  // "Última operación con este contacto" (dato real de BALANCE)
  const ultimaConContacto=(()=>{
    const c=(contactosBalance||[]).find(x=>x.nombre===f.contacto);
    const d=c?.detalles?.find(x=>x.detalle===f.detalle);
    if(!d||!d.ultima) return null;
    const dias=Math.floor((Date.now()-d.ultima)/86400000);
    return dias<0?null:dias;
  })();

  const totalMonto=f.lineas.reduce((s,l)=>s+(parseFloat(l.monto)||0),0);

  // Aplica un borrador de la carga rápida al formulario (el usuario lo revisa y confirma).
  const aplicarBorrador=(b)=>{
    if(b.tipo==='movimiento'){ onIrA('caja',b); showToast('Completé el movimiento de caja. Revisalo antes de registrar.','ok'); return; }
    if(b.tipo==='transferencia'){ onIrA('stock',b); showToast('Completé el movimiento de stock. Revisalo antes de registrar.','ok'); return; }
    setF(p=>{
      const lineas=b.lineas.length?b.lineas.map(l=>{
        const prod=productos.find(x=>x.label===l.producto);
        return {producto:prod?prod.id:'',deposito:l.deposito||'',botellas:l.botellas?String(l.botellas):'',monto:l.monto!==''?String(l.monto):''};
      }):p.lineas;
      const det=b.detalle&&listaDetalles.includes(b.detalle)?b.detalle:p.detalle;
      return {...p,detalle:det,
        contacto:b.contactoEsNuevo?'Nuevo…':(b.contacto||p.contacto),
        contactoNuevo:b.contactoEsNuevo?b.contacto:p.contactoNuevo,
        canal:b.canal||p.canal, fecha:b.fecha||p.fecha, lineas};
    });
    showToast('Formulario completado. Revisalo antes de registrar.','ok');
  };

  const submit=async()=>{
    if(!f.detalle||!f.contacto||(f.contacto==='Nuevo…'&&!f.contactoNuevo)){showToast('Completá los campos obligatorios','error');return;}
    for(const l of f.lineas){
      const cant=parseInt(l.botellas)||0;
      if(cant>0&&(!l.producto||!l.deposito)){showToast('Con botellas, cada línea necesita producto y depósito','error');return;}
      if(cant>0){
        const prod=productos.find(p=>p.id===l.producto);
        const disponible=prod?(prod.stockUbic[l.deposito]||0):0;
        if(cant>disponible){showToast(`Sólo hay ${disponible} bot de ${prod?.label} en ${l.deposito}`,'error');return;}
      }
    }
    const ct=f.contacto==='Nuevo…'?f.contactoNuevo:f.contacto;
    setSending(true);
    try {
      let referencia='';
      const lineasEnvio=f.lineas.map(l=>{
        const prod=productos.find(p=>p.id===l.producto);
        return {producto:prod?.label||'',deposito:l.deposito,botellas:parseInt(l.botellas)||0,monto:l.monto||0};
      });
      if(GAS_URL){
        const r=await gasPost({action:'addTransaccion',fecha:f.fecha,detalle:f.detalle,contacto:ct,canal:esVenta?(f.canal||''):'',lineas:JSON.stringify(lineasEnvio)});
        referencia=r?.referencia||'';
      }
      f.lineas.forEach(l=>{
        const cant=parseInt(l.botellas)||0;
        if(cant>0) applySale({productoId:l.producto,deposito:l.deposito,cantidad:cant});
      });
      const resumen=lineasEnvio.filter(l=>l.producto).map(l=>l.producto).join(', ')||f.detalle;
      await addOp({id:Date.now(),icon:esVenta?'venta':'otro',desc:`${f.detalle}${lineasEnvio.length>1?` (${lineasEnvio.length} productos)`:''}: ${resumen} → ${ct}`,monto:totalMonto?`$${Math.round(totalMonto).toLocaleString('es-AR')}`:'sin monto',fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      if(GAS_URL) await refresh();
      showToast(`✓ ${f.detalle} registrada en BALANCE${referencia?' ('+referencia+')':''}`,'ok');
      onBack();
    } catch(e){ showToast(textoError(e,'Error al registrar. Intentá de nuevo.'),'error'); }
    setSending(false);
  };

  const Etiqueta=({children,htmlFor})=><label htmlFor={htmlFor} style={{fontSize:13,color:C.muted,fontFamily:'system-ui'}}>{children}</label>;
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:20}}>
        <button onClick={onBack} aria-label="Volver" style={{width:44,height:44,display:'flex',alignItems:'center',justifyContent:'center',background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,cursor:'pointer'}}><Ico n="atras" s={20} c={C.muted} w={2}/></button>
        <div><h1 style={{margin:0,color:C.text,fontWeight:700,fontSize:20,fontFamily:'Georgia, serif'}}>Cargar</h1><div style={{color:C.dim,fontSize:13,fontFamily:'system-ui'}}>Se registra en BALANCE</div></div>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:20}}>
        <CargaRapida onResultado={aplicarBorrador}/>

        <section style={{display:'flex',flexDirection:'column',gap:8}}>
          <SL>Tipo de operación</SL>
          <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
            {chipsDetalle.map(d=><Chip key={d} activo={f.detalle===d} onClick={()=>set('detalle',d)}>{d}</Chip>)}
            {listaDetalles.length>6&&!verDetalles&&<Chip punteado onClick={()=>setVerDetalles(true)}>Más…</Chip>}
          </div>
          {tipoBadge!=='Neutro'&&<div style={{fontSize:13,color:C.dim,fontFamily:'system-ui'}}>Se registra como <strong style={{color:tipoBadge==='Ingreso'?'#7dce9b':'#f08080'}}>{tipoBadge.toLowerCase()}</strong> automáticamente.</div>}
        </section>

        <section style={{display:'flex',flexDirection:'column',gap:8}}>
          <SL>Cliente / proveedor</SL>
          <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
            {chipsContacto.map(c=><Chip key={c} activo={f.contacto===c} onClick={()=>setContacto(c)}>{c}</Chip>)}
            {!verContactos&&listaContactos.length>5&&<Chip punteado onClick={()=>setVerContactos(true)}>Más…</Chip>}
            <Chip punteado activo={f.contacto==='Nuevo…'} onClick={()=>set('contacto','Nuevo…')}>+ Nuevo</Chip>
          </div>
          {verContactos&&opcionesContacto.length>0&&<span style={{fontSize:13,color:C.dim,fontFamily:'system-ui'}}>Tocá uno de la lista.</span>}
          {f.contacto==='Nuevo…'&&<Inp aria-label="Nombre del cliente o proveedor" placeholder="Nombre o comercio" value={f.contactoNuevo} onChange={e=>set('contactoNuevo',e.target.value)}/>}
          {ultimaConContacto!==null&&f.contacto&&f.contacto!=='Nuevo…'&&<div style={{display:'flex',alignItems:'center',gap:8,background:C.barrel,border:`1px solid ${C.border}`,borderRadius:12,padding:'10px 12px',fontSize:13,color:C.muted,fontFamily:'system-ui'}}><Ico n="reloj" s={16} c={C.dim} w={2}/>Última “{f.detalle}” con {f.contacto}: {ultimaConContacto===0?'hoy':`hace ${ultimaConContacto} día${ultimaConContacto===1?'':'s'}`}</div>}
          {clienteInfo&&!clienteInfo.activo&&<div style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:10,padding:'10px 14px',fontFamily:'system-ui',fontSize:13,color:'#E07080'}}>{f.contacto} figura inactivo: la última operación fue hace tiempo.</div>}
        </section>

        {f.lineas.map((l,idx)=>{
          const prod=productos.find(p=>p.id===l.producto);
          const disponible=prod&&l.deposito?(prod.stockUbic[l.deposito]||0):null;
          const cant=parseInt(l.botellas)||0;
          const sugLinea=esVenta&&price&&cant?price*cant:null;
          const ajustar=(d)=>setLinea(idx,'botellas',String(Math.max(0,cant+d)));
          return(
            <section key={idx} style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:16,padding:14,display:'flex',flexDirection:'column',gap:14}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <SL>{`Producto${f.lineas.length>1?' '+(idx+1):''}`}</SL>
                {f.lineas.length>1&&<button onClick={()=>quitarLinea(idx)} style={{background:'none',border:'none',color:'#f08080',fontSize:14,cursor:'pointer',fontFamily:'system-ui',minHeight:44,padding:'0 4px'}}>Quitar</button>}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                <Etiqueta htmlFor={`prod-${idx}`}>Producto</Etiqueta>
                <Sel id={`prod-${idx}`} value={l.producto} onChange={e=>setLinea(idx,'producto',e.target.value)}>
                  <option value="">Seleccionar…</option>
                  {productos.filter(p=>totalStock(p)>0).map(p=><option key={p.id} value={p.id}>{p.label} — {totalStock(p).toLocaleString('es-AR')} bot</option>)}
                </Sel>
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                <Etiqueta>Sale del depósito</Etiqueta>
                {!prod&&<span style={{fontSize:13,color:C.dim,fontFamily:'system-ui'}}>Elegí primero el producto.</span>}
                {prod&&<div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                  {UBICACIONES.filter(u=>(prod.stockUbic[u]||0)>0).map(u=><Chip key={u} cuadrado activo={l.deposito===u} onClick={()=>setLinea(idx,'deposito',u)}>{u} · {prod.stockUbic[u].toLocaleString('es-AR')}</Chip>)}
                </div>}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                <Etiqueta htmlFor={`bot-${idx}`}>Botellas</Etiqueta>
                <div style={{display:'flex',alignItems:'center',gap:10}}>
                  <button onClick={()=>ajustar(-1)} aria-label="Restar una botella" style={{width:52,height:52,borderRadius:14,background:C.cork,border:`1px solid ${C.border}`,display:'flex',alignItems:'center',justifyContent:'center',cursor:'pointer'}}><Ico n="menos" c={C.text} w={2.2}/></button>
                  <input id={`bot-${idx}`} type="number" inputMode="numeric" min="0" placeholder="0" value={l.botellas} onChange={e=>setLinea(idx,'botellas',e.target.value)} style={{flex:1,minWidth:0,height:52,textAlign:'center',background:C.cork,border:`1px solid ${C.border}`,borderRadius:14,color:C.text,fontSize:26,fontFamily:'Georgia, serif',fontWeight:700,outline:'none'}}/>
                  <button onClick={()=>ajustar(1)} aria-label="Sumar una botella" style={{width:52,height:52,borderRadius:14,background:C.cork,border:`1px solid ${C.border}`,display:'flex',alignItems:'center',justifyContent:'center',cursor:'pointer'}}><Ico n="mas" c={C.text} w={2.2}/></button>
                </div>
                {disponible!==null&&<span style={{fontSize:13,color:cant>disponible?'#f08080':C.dim,fontFamily:'system-ui'}}>{cant>disponible?`Sólo hay ${disponible} en ${l.deposito}`:`Quedarían ${(disponible-cant).toLocaleString('es-AR')} en ${l.deposito}`}</span>}
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                <Etiqueta htmlFor={`monto-${idx}`}>Monto ($ ARS, opcional)</Etiqueta>
                <Inp id={`monto-${idx}`} type="number" inputMode="decimal" placeholder="0" value={l.monto} onChange={e=>setLinea(idx,'monto',e.target.value)}/>
                {sugLinea&&String(Math.round(sugLinea))!==String(l.monto)&&<button onClick={()=>setLinea(idx,'monto',String(Math.round(sugLinea)))} style={{alignSelf:'flex-start',minHeight:44,padding:'0 14px',borderRadius:22,background:'none',border:'1px solid #2F4A38',color:'#7DCE9B',fontSize:14,cursor:'pointer',fontFamily:'system-ui'}}>Usar sugerido: ${Math.round(sugLinea).toLocaleString('es-AR')} (${Math.round(price).toLocaleString('es-AR')} c/u)</button>}
              </div>
            </section>
          );
        })}
        <button onClick={agregarLinea} style={{minHeight:52,background:'none',border:`1px dashed ${C.border}`,borderRadius:14,color:C.gold,fontSize:15,fontWeight:600,cursor:'pointer',fontFamily:'system-ui'}}>+ Agregar otro producto</button>

        <div style={{display:'grid',gridTemplateColumns:esVenta?'minmax(0,1fr) minmax(0,1fr)':'minmax(0,1fr)',gap:12}}>
          <div style={{display:'flex',flexDirection:'column',gap:6,minWidth:0}}><Etiqueta htmlFor="fecha">Fecha</Etiqueta><Inp id="fecha" type="date" value={f.fecha} onChange={e=>set('fecha',e.target.value)}/></div>
          {esVenta&&<div style={{display:'flex',flexDirection:'column',gap:6,minWidth:0}}><Etiqueta htmlFor="canal">Canal</Etiqueta><Sel id="canal" value={f.canal} onChange={e=>set('canal',e.target.value)}><option value="">Seleccionar…</option>{CANALES.map(c=><option key={c} value={c}>{c}</option>)}</Sel></div>}
        </div>

        <section aria-label="Resumen" style={{background:C.barrel,border:`1px solid ${C.border}`,borderRadius:16,padding:16,display:'flex',flexDirection:'column',gap:10,fontFamily:'system-ui'}}>
          <SL>Vas a registrar</SL>
          {f.lineas.map((l,i)=>{const prod=productos.find(p=>p.id===l.producto);const cant=parseInt(l.botellas)||0;return(
            <div key={i} style={{display:'flex',justifyContent:'space-between',gap:12,fontSize:15,lineHeight:1.4}}>
              <span>{cant?`${cant} bot `:''}{prod?prod.label:(f.detalle)}{f.contacto&&f.contacto!=='Nuevo…'?` → ${f.contacto}`:(f.contactoNuevo?` → ${f.contactoNuevo}`:'')}{l.deposito&&<><br/><span style={{fontSize:13,color:C.dim}}>desde {l.deposito}</span></>}</span>
              <strong style={{whiteSpace:'nowrap'}}>{l.monto?`$${Math.round(parseFloat(l.monto)||0).toLocaleString('es-AR')}`:'sin monto'}</strong>
            </div>);})}
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',paddingTop:10,borderTop:`1px solid ${C.border}`}}>
            <span style={{fontSize:14,fontWeight:700,color:C.gold}}>Total</span>
            <span style={{fontFamily:'Georgia, serif',fontSize:24,fontWeight:700,color:'#7DCE9B'}}>${Math.round(totalMonto).toLocaleString('es-AR')}</span>
          </div>
        </section>

        <div style={{display:'flex',gap:10}}><Btn variant="ghost" onClick={onBack} style={{flex:1}}>Cancelar</Btn><Btn onClick={submit} style={{flex:2,opacity:sending?0.6:1}} disabled={sending}>{sending?'Registrando…':`Registrar ${esVenta?'venta':f.detalle.toLowerCase()}`}</Btn></div>
      </div>
    </div>
  );
}

// ── CAJA ─────────────────────────────────────────────────────────────────
function CajaScreen({user,onBack,showToast,addOp,ventasPendientes,comprasPendientes,refresh,resumenCajas,categorias,contactosBalance,clientes,borrador}){
  const hoy=fechaLocalHoy();
  // Si viene de la carga rápida, el formulario arranca completado (se revisa antes de registrar).
  const bm=borrador?.tipo==='movimiento'?borrador:null;
  const detallesOk=categorias?.length?categorias.map(c=>c.detalle):['Cobro','Gasto'];
  const [f,setF]=useState({
    tipo:bm?.movimiento?.tipo||'cobro', monto:bm?.movimiento?.monto?String(bm.movimiento.monto):'',
    caja:bm?.movimiento?.caja||'Empresa (Ludico)', detalle:bm&&detallesOk.includes(bm.detalle)?bm.detalle:'',
    contacto:bm?(bm.contactoEsNuevo?'Nuevo…':(bm.contacto||'')):'', contactoNuevo:bm?.contactoEsNuevo?bm.contacto:'',
    fecha:bm?.fecha||hoy, referencia:bm?.movimiento?.referencia||''});
  const [sending,setSending]=useState(false);
  // Cuadro de control: últimos movimientos de CAJA, cargados una vez y revelados de a 10.
  const [movimientos,setMovimientos]=useState(null); // null=cargando, []=vacío, [...]=datos
  const [visibles,setVisibles]=useState(10);
  const cargarMovimientos=async()=>{
    if(!GAS_URL){setMovimientos([]);return;}
    try{
      const d=await gasGet({action:'getMovimientosCaja'});
      setMovimientos(Array.isArray(d?.movimientos)?d.movimientos:[]);
    }catch(e){ setMovimientos([]); }
  };
  useEffect(()=>{cargarMovimientos();},[]);
  const set=(k,v)=>setF(p=>({...p,[k]:v,...((k==='detalle'||k==='contacto')?{referencia:''}:{})}));
  const listaDetalles = categorias?.length ? categorias.map(c=>c.detalle) : ['Cobro','Gasto'];
  const listaContactos = contactosParaDetalle(contactosBalance, clientes, f.detalle);
  const submit=async()=>{
    if(!f.monto||!f.detalle||!f.contacto||(f.contacto==='Nuevo…'&&!f.contactoNuevo)){showToast('Completá los campos obligatorios','error');return;}
    if(!(parseFloat(f.monto)>0)){showToast('El monto tiene que ser mayor a cero (cobro o gasto define el signo)','error');return;}
    const ct=f.contacto==='Nuevo…'?f.contactoNuevo:f.contacto;
    setSending(true);
    try {
      if(GAS_URL) await gasPost({action:'addMovement',fecha:f.fecha,tipo:f.tipo,monto:f.monto,caja:f.caja,detalle:f.detalle,contacto:ct,referencia:f.referencia||''});
      const icon=f.tipo;
      await addOp({id:Date.now(),icon,desc:`${f.detalle}: ${ct}`,monto:`$${Math.round(parseFloat(f.monto)).toLocaleString('es-AR')}`,fecha:new Date(f.fecha+'T12:00').toLocaleDateString('es-AR'),user:user.nombre});
      if(GAS_URL) await refresh();
      showToast(`Movimiento registrado${GAS_URL?' en CAJA':''}`,'ok');
      onBack();
    } catch(e){ showToast(textoError(e,'Error al registrar. Intentá de nuevo.'),'error'); }
    setSending(false);
  };
  return(
    <div style={{padding:'20px 16px',maxWidth:500,margin:'0 auto'}}>
      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:24}}>
        <button onClick={onBack} aria-label="Volver" style={{width:44,height:44,display:'flex',alignItems:'center',justifyContent:'center',background:C.cork,border:`1px solid ${C.border}`,borderRadius:12,cursor:'pointer'}}><Ico n="atras" s={20} c={C.muted} w={2}/></button>
        <div><h1 style={{margin:0,color:C.text,fontWeight:700,fontSize:20,fontFamily:'Georgia, serif'}}>Caja</h1><div style={{color:C.dim,fontSize:13,fontFamily:'system-ui'}}>Registrar cobro, pago o gasto</div></div>
      </div>
      {resumenCajas?.length>0&&(
        <Card style={{marginBottom:20}}>
          <SL>Estado de cajas</SL>
          <div style={{display:'grid',gridTemplateColumns:'1fr 76px 76px',gap:6,padding:'0 2px 6px'}}>
            {['Caja','AR$','USD'].map((h,i)=><div key={h} style={{fontSize:12,color:C.dim,letterSpacing:'0.08em',textTransform:'uppercase',fontFamily:'system-ui',textAlign:i===0?'left':'center'}}>{h}</div>)}
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
        {[{k:'cobro',l:'Cobro / Ingreso',bg:C.greenBg,col:'#7dce9b'},{k:'gasto',l:'Gasto / Egreso',bg:C.wineBg,col:'#f08080'}].map((t,i)=>(
          <button key={t.k} onClick={()=>{set('tipo',t.k);set('referencia','');}} style={{flex:1,padding:'14px 8px',background:f.tipo===t.k?t.bg:C.barrel,color:f.tipo===t.k?t.col:C.muted,border:'none',borderRight:i===0?`1px solid ${C.border}`:'none',cursor:'pointer',fontSize:14,fontFamily:'system-ui',fontWeight:f.tipo===t.k?700:400}}>{t.l}</button>
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

      <Card style={{marginTop:24}}>
        <SL>Últimos movimientos</SL>
        {movimientos===null&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'8px 2px'}}>Cargando…</div>}
        {movimientos!==null&&movimientos.length===0&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'8px 2px'}}>No hay movimientos para mostrar.</div>}
        {movimientos!==null&&movimientos.length>0&&(
          <>
            {movimientos.slice(0,visibles).map((m,i)=>{
              const esGasto=m.monto<0;
              return (
                <div key={i} style={{display:'grid',gridTemplateColumns:'1fr auto',gap:8,padding:'9px 2px',borderTop:`1px solid ${C.border}`,alignItems:'center'}}>
                  <div style={{minWidth:0}}>
                    <div style={{color:C.text,fontSize:13,fontFamily:'system-ui',fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                      {m.detalle||'—'}{m.contacto?` · ${m.contacto}`:''}
                    </div>
                    <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginTop:2}}>
                      {m.fecha} · {m.caja}{m.user?` · cargó ${m.user}`:''}
                    </div>
                  </div>
                  <div style={{textAlign:'right',fontSize:13,fontFamily:'system-ui',fontWeight:700,color:esGasto?'#f08080':'#7dce9b',whiteSpace:'nowrap'}}>
                    {esGasto?'-':'+'}${Math.abs(m.monto).toLocaleString('es-AR')}
                  </div>
                </div>
              );
            })}
            {visibles<movimientos.length&&(
              <button onClick={()=>setVisibles(v=>v+10)} style={{width:'100%',marginTop:10,background:'none',border:`1px dashed ${C.border}`,borderRadius:10,padding:'10px',color:C.gold,fontSize:13,fontFamily:'system-ui',cursor:'pointer'}}>
                Ver más ({movimientos.length-visibles} restantes)
              </button>
            )}
          </>
        )}
      </Card>
    </div>
  );
}

// ── GRÁFICOS (SVG a mano, sin dependencias) ─────────────────────────────
// Barras verticales. datos: [{label, valor}]. fmt: función para formatear el valor.
function BarChart({datos,fmt,color=C.gold,alto=150}){
  const limpios=(datos||[]).map(d=>({...d,valor:Number(d.valor)||0}));
  if(limpios.length===0) return <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'16px 0',textAlign:'center'}}>Sin datos</div>;
  const max=Math.max(...limpios.map(d=>Math.abs(d.valor)),1);
  const topPad=16;   // aire arriba para que la etiqueta del valor no se corte
  const botPad=30;   // aire abajo para las etiquetas del eje
  const ancho=Math.max(limpios.length*38,280);
  const paso=ancho/limpios.length;
  const anchoBarra=Math.max(6,Math.min(28,paso*0.6));
  return(
    <div style={{overflowX:'auto',paddingBottom:4}}>
      <svg width={ancho} height={topPad+alto+botPad} style={{display:'block'}}>
        {limpios.map((d,i)=>{
          const x=(i+0.5)*paso;
          const h=Math.abs(d.valor)/max*alto;
          const yBarra=topPad+alto-h;
          return(
            <g key={i}>
              <rect x={x-anchoBarra/2} y={yBarra} width={anchoBarra} height={Math.max(h,1)} rx={3} fill={d.valor<0?'#f08080':color} opacity={0.85}/>
              <text x={x} y={yBarra-4} textAnchor="middle" fontSize="11" fill={C.muted} fontFamily="system-ui">{fmt?fmt(d.valor):Math.round(d.valor)}</text>
              <text x={x} y={topPad+alto+14} textAnchor="middle" fontSize="11" fill={C.dim} fontFamily="system-ui">{d.label}</text>
              {d.sub&&<text x={x} y={topPad+alto+25} textAnchor="middle" fontSize="11" fill={C.dim} fontFamily="system-ui" opacity={0.7}>{d.sub}</text>}
            </g>
          );
        })}
        <line x1="0" y1={topPad+alto} x2={ancho} y2={topPad+alto} stroke={C.border} strokeWidth="1"/>
      </svg>
    </div>
  );
}

// Barras horizontales, mejor para rankings con nombres largos (clientes, productos).
function BarChartH({datos,fmt,color=C.gold}){
  const limpios=(datos||[]).map(d=>({...d,valor:Number(d.valor)||0}));
  if(limpios.length===0) return <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'16px 0',textAlign:'center'}}>Sin datos</div>;
  const max=Math.max(...limpios.map(d=>Math.abs(d.valor)),1);
  return(
    <div>
      {limpios.map((d,i)=>(
        <div key={i} style={{marginBottom:9}}>
          <div style={{display:'flex',justifyContent:'space-between',gap:8,marginBottom:3}}>
            <span style={{color:C.text,fontSize:11,fontFamily:'system-ui',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{d.label}</span>
            <span style={{color:C.muted,fontSize:11,fontFamily:'system-ui',fontWeight:700,whiteSpace:'nowrap'}}>{fmt?fmt(d.valor):Math.round(d.valor)}</span>
          </div>
          <div style={{height:6,background:C.barrel,borderRadius:3,overflow:'hidden'}}>
            <div style={{height:'100%',width:`${Math.abs(d.valor)/max*100}%`,background:d.valor<0?'#f08080':color,borderRadius:3}}/>
          </div>
        </div>
      ))}
    </div>
  );
}

// Barras agrupadas de a dos (ej. ingresos vs egresos por año).
function BarChartDoble({datos,fmt,colorA=C.gold,colorB='#f08080',labelA,labelB,alto=150}){
  const limpios=(datos||[]).map(d=>({...d,a:Number(d.a)||0,b:Number(d.b)||0}));
  if(limpios.length===0) return <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',padding:'16px 0',textAlign:'center'}}>Sin datos</div>;
  const max=Math.max(...limpios.flatMap(d=>[Math.abs(d.a),Math.abs(d.b)]),1);
  const topPad=10, botPad=22;
  const ancho=Math.max(limpios.length*62,280);
  const paso=ancho/limpios.length;
  const bw=Math.max(6,Math.min(20,paso*0.3));
  return(
    <div>
      <div style={{display:'flex',gap:14,marginBottom:8,fontFamily:'system-ui',fontSize:11}}>
        <span style={{color:C.muted}}><span style={{display:'inline-block',width:8,height:8,background:colorA,borderRadius:2,marginRight:4}}/>{labelA}</span>
        <span style={{color:C.muted}}><span style={{display:'inline-block',width:8,height:8,background:colorB,borderRadius:2,marginRight:4}}/>{labelB}</span>
      </div>
      <div style={{overflowX:'auto',paddingBottom:4}}>
        <svg width={ancho} height={topPad+alto+botPad} style={{display:'block'}}>
          {limpios.map((d,i)=>{
            const cx=(i+0.5)*paso;
            const ha=Math.abs(d.a)/max*alto, hb=Math.abs(d.b)/max*alto;
            return(
              <g key={i}>
                <rect x={cx-bw-2} y={topPad+alto-ha} width={bw} height={Math.max(ha,1)} rx={3} fill={colorA} opacity={0.85}/>
                <rect x={cx+2} y={topPad+alto-hb} width={bw} height={Math.max(hb,1)} rx={3} fill={colorB} opacity={0.85}/>
                <text x={cx} y={topPad+alto+15} textAnchor="middle" fontSize="11" fill={C.dim} fontFamily="system-ui">{d.label}</text>
              </g>
            );
          })}
          <line x1="0" y1={topPad+alto} x2={ancho} y2={topPad+alto} stroke={C.border} strokeWidth="1"/>
        </svg>
      </div>
    </div>
  );
}

// ── DATOS / COSTEO ───────────────────────────────────────────────────────
function DatosScreen({price,productos}){
  const [data,setData]=useState(null); // null=cargando
  const [an,setAn]=useState(null);     // analytics
  const [modoVentas,setModoVentas]=useState('ars');   // ars | usd | bot
  const [escalaVentas,setEscalaVentas]=useState('mes'); // mes | anio
  useEffect(()=>{(async()=>{
    if(!GAS_URL){setData({insumos:[],existe:false});setAn({});return;}
    // Las dos consultas salen juntas: antes la segunda esperaba a que terminara la primera.
    await Promise.all([
      gasGet({action:'getInsumos'}).then(setData).catch(()=>setData({insumos:[],existe:false,error:true})),
      gasGet({action:'getAnalytics'}).then(setAn).catch(()=>setAn({})),
    ]);
  })();},[]);

  const totalArs=data?.totalArs||0;
  const totalUsd=data?.totalUsd||0;
  // Margen contra el precio de venta actual (última venta registrada)
  const margen=price&&totalArs?price-totalArs:null;
  const margenPct=price&&totalArs?((price-totalArs)/price*100):null;
  const activos=(data?.insumos||[]).filter(i=>i.activo);
  const sinCargar=activos.filter(i=>!i.cargado).length;
  const desactualizados=activos.filter(i=>i.cargado&&i.diasDesde!==null&&i.diasDesde>60).length;

  return(
    <div style={{padding:'20px 16px 90px',maxWidth:500,margin:'0 auto'}}>
      <div style={{color:C.text,fontWeight:700,fontSize:18,fontFamily:'Georgia, serif',marginBottom:2}}>Datos</div>
      <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginBottom:20}}>Costo de reposición y análisis</div>

      {data===null&&<div style={{color:C.dim,fontSize:13,fontFamily:'system-ui',padding:'20px 0'}}>Cargando…</div>}

      {data&&!data.existe&&(
        <Card>
          <SL>Costeo de reposición</SL>
          <div style={{color:C.muted,fontSize:13,fontFamily:'system-ui',lineHeight:1.5}}>
            Todavía no existe la pestaña <strong>INSUMOS</strong> en la planilla.
            Corré <code style={{color:C.gold}}>crearHojaInsumos()</code> desde Apps Script y después cargá los precios ahí.
          </div>
        </Card>
      )}

      {data?.existe&&(
        <>
          <Card>
            <SL>Costo por botella hoy</SL>
            {(()=>{
              const todosUsd=activos.length>0&&activos.filter(i=>i.cargado).every(i=>i.moneda==='USD');
              return (
                <div style={{display:'flex',alignItems:'baseline',gap:10,marginBottom:4}}>
                  {todosUsd?(
                    <>
                      <span style={{color:C.gold,fontSize:30,fontFamily:'Georgia, serif',fontWeight:700}}>US${totalUsd.toFixed(2)}</span>
                      <span style={{color:C.dim,fontSize:14,fontFamily:'system-ui'}}>${Math.round(totalArs).toLocaleString('es-AR')}</span>
                    </>
                  ):(
                    <>
                      <span style={{color:C.gold,fontSize:30,fontFamily:'Georgia, serif',fontWeight:700}}>${Math.round(totalArs).toLocaleString('es-AR')}</span>
                      <span style={{color:C.dim,fontSize:14,fontFamily:'system-ui'}}>US${totalUsd.toFixed(2)}</span>
                    </>
                  )}
                </div>
              );
            })()}
            <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>Suma de insumos activos, convertidos al dólar de hoy</div>

            {price&&(
              <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${C.border}`,fontFamily:'system-ui'}}>
                <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:C.muted,padding:'3px 0'}}>
                  <span>Precio de venta (última)</span><span>${Math.round(price).toLocaleString('es-AR')}</span>
                </div>
                <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:C.muted,padding:'3px 0'}}>
                  <span>Costo de reposición</span><span>-${Math.round(totalArs).toLocaleString('es-AR')}</span>
                </div>
                <div style={{display:'flex',justifyContent:'space-between',fontSize:15,fontWeight:700,padding:'7px 0 0',borderTop:`1px solid ${C.border}`,marginTop:4}}>
                  <span style={{color:C.gold}}>Margen por botella</span>
                  <span style={{color:margen>=0?'#7dce9b':'#f08080'}}>
                    {margen>=0?'+':'-'}${Math.abs(Math.round(margen)).toLocaleString('es-AR')}
                    {margenPct!==null&&<span style={{fontSize:12,fontWeight:400,marginLeft:5}}>({margenPct.toFixed(0)}%)</span>}
                  </span>
                </div>
              </div>
            )}
          </Card>

          {(sinCargar>0||desactualizados>0)&&(
            <div style={{background:C.wineBg,border:`1px solid ${C.wine}55`,borderRadius:10,padding:'10px 14px',margin:'12px 0',fontFamily:'system-ui',fontSize:12,color:'#E0A070',lineHeight:1.5}}>
              {sinCargar>0&&<div>{sinCargar} insumo{sinCargar===1?'':'s'} sin precio cargado — el costo está incompleto.</div>}
              {desactualizados>0&&<div>{desactualizados} precio{desactualizados===1?'':'s'} con más de 60 días.</div>}
            </div>
          )}

          <Card style={{marginTop:12}}>
            <SL>Insumos</SL>
            {activos.length===0&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui'}}>No hay insumos activos.</div>}
            {activos.map((it,i)=>(
              <div key={i} style={{display:'flex',justifyContent:'space-between',gap:8,padding:'8px 2px',borderTop:i?`1px solid ${C.border}`:'none',fontFamily:'system-ui'}}>
                <div style={{minWidth:0}}>
                  <div style={{color:it.cargado?C.text:C.dim,fontSize:13,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{it.nombre}</div>
                  <div style={{color:C.dim,fontSize:12,marginTop:2}}>
                    {it.cargado?`${it.moneda} ${it.costo.toLocaleString('es-AR')}`:'sin precio'}
                    {it.fecha?` · ${it.fecha}`:''}
                    {it.diasDesde!==null&&it.diasDesde>60?'':''}
                    {it.proveedor?` · ${it.proveedor}`:''}
                  </div>
                </div>
                <span style={{color:it.cargado?C.text:C.dim,fontSize:13,fontWeight:700,whiteSpace:'nowrap'}}>
                  {it.cargado?`$${Math.round(it.costoArs).toLocaleString('es-AR')}`:'—'}
                </span>
              </div>
            ))}
            <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginTop:10,lineHeight:1.5}}>
              Los precios se cargan en la pestaña INSUMOS de la planilla, ya divididos por botella.
            </div>
          </Card>
        </>
      )}

      {an&&Object.keys(an).length>0&&(()=>{
        const fmtArs=v=>`$${Math.round(v/1000)}k`;
        const fmtUsd=v=>`US$${Math.round(v)}`;
        const fmtBot=v=>`${Math.round(v)}`;
        const fmtArsFull=v=>`$${Math.round(v).toLocaleString('es-AR')}`;

        // Ventas: agrupar por mes o por año según la escala elegida
        const serieVentas=(()=>{
          const base=an.ventasMes||[];
          if(escalaVentas==='anio'){
            const porAnio={};
            base.forEach(m=>{
              if(!porAnio[m.anio]) porAnio[m.anio]={label:String(m.anio),ars:0,usd:0,botellas:0};
              porAnio[m.anio].ars+=m.ars; porAnio[m.anio].usd+=m.usd; porAnio[m.anio].botellas+=m.botellas;
            });
            return Object.values(porAnio);
          }
          return base.slice(-14).map(m=>({label:m.periodo.slice(5)+'/'+m.periodo.slice(2,4),ars:m.ars,usd:m.usd,botellas:m.botellas}));
        })();
        const campo=modoVentas==='ars'?'ars':(modoVentas==='usd'?'usd':'botellas');
        const fmtSel=modoVentas==='ars'?fmtArs:(modoVentas==='usd'?fmtUsd:fmtBot);

        return(
        <>
          <Card style={{marginTop:12}}>
            <SL>Ventas</SL>
            <div style={{display:'flex',gap:6,marginBottom:12,flexWrap:'wrap'}}>
              {[['ars','$'],['usd','US$'],['bot','Botellas']].map(([k,l])=>(
                <button key={k} onClick={()=>setModoVentas(k)} style={{padding:'5px 12px',borderRadius:14,border:`1px solid ${modoVentas===k?C.gold:C.border}`,background:modoVentas===k?`${C.gold}22`:'none',color:modoVentas===k?C.gold:C.dim,fontSize:12,fontFamily:'system-ui',cursor:'pointer',fontWeight:modoVentas===k?700:400}}>{l}</button>
              ))}
              <span style={{flex:1}}/>
              {[['mes','Mes'],['anio','Año']].map(([k,l])=>(
                <button key={k} onClick={()=>setEscalaVentas(k)} style={{padding:'5px 12px',borderRadius:14,border:`1px solid ${escalaVentas===k?C.gold:C.border}`,background:escalaVentas===k?`${C.gold}22`:'none',color:escalaVentas===k?C.gold:C.dim,fontSize:12,fontFamily:'system-ui',cursor:'pointer',fontWeight:escalaVentas===k?700:400}}>{l}</button>
              ))}
            </div>
            <BarChart datos={serieVentas.map(d=>({label:d.label,valor:d[campo]}))} fmt={fmtSel}/>
          </Card>

          <Card style={{marginTop:12}}>
            <SL>Ventas por cliente</SL>
            <BarChartH datos={(an.ventasCliente||[]).slice(0,10).map(c=>({label:c.cliente,valor:modoVentas==='usd'?c.usd:(modoVentas==='bot'?c.botellas:c.ars)}))} fmt={modoVentas==='usd'?fmtUsd:(modoVentas==='bot'?fmtBot:fmtArsFull)}/>
            {(an.ventasCliente||[]).length>0&&(()=>{
              const tot=an.ventasCliente.reduce((s,c)=>s+c.ars,0);
              const top=an.ventasCliente[0];
              const pct=tot?Math.round(top.ars/tot*100):0;
              return <div style={{color:pct>40?'#E0A070':C.dim,fontSize:12,fontFamily:'system-ui',marginTop:10,lineHeight:1.5}}>
                {pct>40?'':''}{top.cliente} concentra el {pct}% de las ventas históricas.
              </div>;
            })()}
          </Card>

          <Card style={{marginTop:12}}>
            <SL>Costo unitario por producto (US$)</SL>
            <BarChartH datos={(an.costoProducto||[]).map(p=>({label:p.producto,valor:p.cuUsd}))} fmt={v=>`US$${v.toFixed(2)}`}/>
            {totalUsd>0&&<div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginTop:10,lineHeight:1.5}}>
              Costo de reposición hoy: <strong style={{color:C.gold}}>US${totalUsd.toFixed(2)}</strong> por botella. Comparalo con el histórico de cada producto para ver la erosión.
            </div>}
          </Card>

          <Card style={{marginTop:12}}>
            <SL>Flujo por año (liquidez)</SL>
            <BarChartDoble
              datos={(an.flujoAnual||[]).map(f=>({label:String(f.anio),a:f.ingresosUsd,b:f.egresosUsd}))}
              labelA="Ingresos US$" labelB="Egresos US$" colorA="#7dce9b" colorB="#f08080"/>
            <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginTop:10,lineHeight:1.5}}>
              Plata que entró y salió cada año. No mide rentabilidad: los costos de una añada se pagan antes de venderla.
            </div>
          </Card>

          <Card style={{marginTop:12}}>
            <SL>Margen por añada (US$)</SL>
            <BarChartDoble
              datos={(an.margenAnada||[]).map(m=>({label:m.anada,a:m.ventasUsd,b:m.costosUsd}))}
              labelA="Vendido" labelB="Costó" colorA="#7dce9b" colorB="#f08080"/>
            {(an.margenAnada||[]).map((m,i)=>(
              <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'6px 2px',borderTop:`1px solid ${C.border}`,fontFamily:'system-ui',fontSize:12}}>
                <span style={{color:C.muted}}>{m.anada}<span style={{color:C.dim,fontSize:12}}>{m.botellasVendidas?` · ${m.botellasVendidas} bot`:''}</span></span>
                <span style={{color:m.margenUsd>=0?'#7dce9b':'#f08080',fontWeight:700}}>{m.margenUsd>=0?'+':'-'}US${Math.abs(m.margenUsd).toLocaleString('es-AR')}</span>
              </div>
            ))}
            <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',marginTop:10,lineHeight:1.5}}>
              Cada añada: todo lo que se vendió contra todo lo que costó producirla. Las añadas nuevas todavía tienen stock por vender.
            </div>
          </Card>

          {(an.insumosAnada||[]).length>0&&(
            <Card style={{marginTop:12}}>
              <SL>Costo de insumos por añada (US$/botella)</SL>
              {(()=>{
                const porInsumo={};
                (an.insumosAnada||[]).forEach(x=>{
                  if(!porInsumo[x.insumo]) porInsumo[x.insumo]=[];
                  porInsumo[x.insumo].push(x);
                });
                const nombres=Object.keys(porInsumo).sort();
                return nombres.map(n=>(
                  <div key={n} style={{marginBottom:14}}>
                    <div style={{color:C.text,fontSize:12,fontFamily:'system-ui',fontWeight:600,marginBottom:6}}>{n}</div>
                    <BarChart datos={porInsumo[n].map(x=>({label:x.anada,valor:x.cuUsd}))} fmt={v=>`$${v.toFixed(2)}`} alto={70}/>
                  </div>
                ));
              })()}
              <div style={{color:C.dim,fontSize:12,fontFamily:'system-ui',lineHeight:1.5}}>
                Costo unitario real de cada insumo por añada. Los que más suben en dólares son candidatos a stockear.
              </div>
            </Card>
          )}
        </>
        );
      })()}
    </div>
  );
}

// ── NAV ──────────────────────────────────────────────────────────────────
// Cinco lugares; "Cargar" es el botón central, destacado, porque es lo que más se usa.
function Nav({screen,setScreen}){
  const tabs=[{id:'dashboard',ico:'inicio',l:'Inicio'},{id:'stock',ico:'stock',l:'Stock'},{id:'venta',ico:'mas',l:'Cargar',fab:true},{id:'caja',ico:'caja',l:'Caja'},{id:'datos',ico:'datos',l:'Datos'}];
  return(<nav aria-label="Principal" style={{position:'fixed',bottom:0,left:0,right:0,background:C.barrel,borderTop:`1px solid ${C.border}`,display:'flex',alignItems:'flex-end',zIndex:200,paddingBottom:'env(safe-area-inset-bottom, 0px)'}}>
    {tabs.map(t=>{
      const act=screen===t.id;
      return(<button key={t.id} onClick={()=>setScreen(t.id)} aria-current={act?'page':undefined} style={{flex:1,minHeight:60,padding:'8px 2px 8px',background:'none',border:'none',cursor:'pointer',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'flex-end',gap:3}}>
        {t.fab
          ?<span style={{width:52,height:52,marginTop:-22,borderRadius:'50%',background:C.gold,display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 4px 14px rgba(0,0,0,0.5)',border:`3px solid ${C.barrel}`}}><Ico n={t.ico} s={26} c="#1A1612" w={2.4}/></span>
          :<Ico n={t.ico} c={act?C.gold:C.dim}/>}
        <span style={{fontSize:12,fontFamily:'system-ui',color:act?C.gold:C.muted,fontWeight:act?700:500}}>{t.l}</span>
      </button>);
    })}
  </nav>);
}

// ── APP ───────────────────────────────────────────────────────────────────
export default function NovatoApp(){
  const [user,setUser]=useState(null);const [screen,setScreen]=useState('dashboard');const [borrador,setBorrador]=useState(null);const [vistaStock,setVistaStock]=useState(null);const [toast,setToast]=useState(null);const [settings,setSettings]=useState(false);
  const {price,source,ops,addOp,stockUbicacion,ventasPendientes,comprasPendientes,operacionesPendientes,clientes,categorias,contactosBalance,ultimoControlStock,resumenCajas,refresh}=useAppData(user);
  const {productos,applyTransfer,applySale}=useProductos(stockUbicacion);
  const {last,save}=useStockControl(ultimoControlStock);
  const showToast=(msg,type='ok')=>{setToast({msg,type});setTimeout(()=>setToast(null),3500);};
  // Navegación: ir a una pantalla limpia, o llegar con un borrador de la carga rápida ya cargado.
  const irA=(destino)=>{
    setBorrador(null);setVistaStock(null);
    if(destino==='stock-mover'){setVistaStock('transfer');setScreen('stock');return;}
    setScreen(destino);window.scrollTo?.(0,0);
  };
  const irConBorrador=(destino,datos)=>{
    setBorrador(datos);setVistaStock(destino==='stock'?'transfer':null);setScreen(destino);window.scrollTo?.(0,0);
  };
  const logout=()=>{authToken='';setUser(null);setSettings(false);setScreen('dashboard');};
  // Si el servidor avisa que la sesión venció, se vuelve al login.
  alVencerSesion=()=>{ setUser(null);setSettings(false);setScreen('dashboard');showToast('Tu sesión venció. Ingresá de nuevo.','error'); };
  // Al volver a la app (de otra app o del bloqueo del celular) se actualizan los datos si pasó más de 1 minuto,
  // para no operar con un stock viejo.
  const ultimoRefresh=useRef(Date.now());
  useEffect(()=>{
    if(!user||!GAS_URL) return;
    const alVolver=()=>{ if(document.visibilityState==='visible'&&Date.now()-ultimoRefresh.current>60000){ ultimoRefresh.current=Date.now(); refresh(); } };
    document.addEventListener('visibilitychange',alVolver);
    return ()=>document.removeEventListener('visibilitychange',alVolver);
  },[user]);
  if(!user)return <LoginScreen onLogin={(u,t)=>{authToken=t||'';ultimoRefresh.current=Date.now();setUser(u);setScreen('dashboard');}}/>;
  return(
    <div style={{fontFamily:'Georgia, serif',background:C.cellar,minHeight:'100vh',color:C.text}}>
      {toast&&<div style={{position:'fixed',top:'calc(env(safe-area-inset-top, 0px) + 68px)',left:'50%',transform:'translateX(-50%)',zIndex:999,maxWidth:'calc(100vw - 32px)',textAlign:'center',background:toast.type==='ok'?C.greenBg:C.wineBg,border:`1px solid ${toast.type==='ok'?C.green:C.wine}`,color:toast.type==='ok'?'#7dce9b':'#f08080',padding:'12px 20px',borderRadius:12,fontSize:14,fontFamily:'system-ui',boxShadow:'0 4px 20px rgba(0,0,0,0.5)'}}>{toast.msg}</div>}
      {settings&&<SettingsPanel user={user} onClose={()=>setSettings(false)} onLogout={logout} showToast={showToast}/>}
      <div style={{background:C.barrel,borderBottom:`1px solid ${C.border}`,padding:'calc(env(safe-area-inset-top, 0px) + 13px) 18px 13px',display:'flex',alignItems:'center',justifyContent:'space-between',position:'sticky',top:0,zIndex:100}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}><span style={{color:C.gold,fontWeight:700,letterSpacing:'0.25em',fontSize:15,fontFamily:'Georgia, serif'}}>NOVATO</span><span style={{color:C.border}}>|</span><span style={{color:C.muted,fontSize:13,fontFamily:'system-ui'}}>Gestión</span></div>
        <button onClick={()=>setSettings(true)} aria-label="Cuenta y ajustes" style={{display:'flex',alignItems:'center',gap:8,minHeight:44,background:C.cork,border:`1px solid ${C.border}`,borderRadius:22,padding:'4px 14px 4px 5px',color:C.text,fontSize:14,fontFamily:'system-ui',cursor:'pointer'}}><Avatar u={user} s={34}/><span>{user.nombre}</span><Ico n="ajustes" s={16} c={C.dim}/></button>
      </div>
      <main style={{paddingBottom:'calc(96px + env(safe-area-inset-bottom, 0px))'}}>
        {screen==='dashboard'&&<DashboardScreen user={user} onNavigate={irA} price={price} source={source} productos={productos} last={last} operacionesPendientes={operacionesPendientes} resumenCajas={resumenCajas}/>}
        {screen==='venta'&&<VentaScreen user={user} onBack={()=>irA('dashboard')} showToast={showToast} addOp={addOp} price={price} productos={productos} applySale={applySale} clientes={clientes} categorias={categorias} contactosBalance={contactosBalance} refresh={refresh} onIrA={irConBorrador}/>}
        {screen==='caja'&&<CajaScreen user={user} onBack={()=>irA('dashboard')} showToast={showToast} addOp={addOp} ventasPendientes={ventasPendientes} comprasPendientes={comprasPendientes} refresh={refresh} resumenCajas={resumenCajas} categorias={categorias} contactosBalance={contactosBalance} clientes={clientes} borrador={borrador}/>}
        {screen==='stock'&&<StockScreen onBack={()=>irA('dashboard')} showToast={showToast} productos={productos} applyTransfer={applyTransfer} user={user} refresh={refresh} last={last} save={save} vistaInicial={vistaStock} borrador={borrador}/>}
        {screen==='datos'&&<DatosScreen price={price} productos={productos}/>}
      </main>
      <Nav screen={screen} setScreen={irA}/>
    </div>
  );
}
