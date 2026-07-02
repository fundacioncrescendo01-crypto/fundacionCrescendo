import { useState, useEffect, useMemo, useRef } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, AreaChart, Area,
} from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabase";
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';

/* ─── FONTS ─────────────────────────────────────────────────────────────── */
const fl = document.createElement("link");
fl.rel = "stylesheet";
fl.href = "https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap";
document.head.appendChild(fl);



/* ─── PALETTE "CUIDADO INTEGRAL" ──────────────────────────────────────── */
const T = {
  primary: "#2C7A7B",
  primaryLt: "#4A9C9D",
  accent: "#E8A87C",
  accentLt: "#F2C4A8",
  gold: "#D4B86A",
  bg: "#F4F7F6",
  bgCard: "#FFFFFF",
  bgSoft: "#E8F0EF",
  border: "#D0DFDE",
  text: "#3D4A4A",
  textSub: "#5A6B6B",
  textMuted: "#8FA3A3",
  green: "#6BBF8A",
  red: "#E5989B",
  blue: "#A8DADC",
  teal: "#80CBC4",
  purple: "#B39DDB",
  sidebar: "#1E3D3D",
  sidebarLine: "rgba(255,255,255,0.08)",
};

const PIE = [T.primary, T.accent, T.blue, T.gold, T.purple, T.teal];

/* ─── GLOBAL STYLES ─────────────────────────────────────────────────────── */
const GS = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background: ${T.bg};
    color: ${T.text};
    line-height: 1.6;
  }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: #C5D6D5; border-radius: 4px; }
  input, select, textarea {
    font-family: 'Plus Jakarta Sans', sans-serif;
    border: 1.5px solid ${T.border};
    border-radius: 9px;
    padding: 9px 13px;
    font-size: 13px;
    width: 100%;
    background: #FAFCFB;
    color: ${T.text};
    outline: none;
    transition: border-color 0.18s, box-shadow 0.18s;
  }
  input:focus, select:focus, textarea:focus {
    border-color: ${T.primary};
    box-shadow: 0 0 0 3px ${T.primary}18;
    background: #FFFFFF;
  }
  textarea { resize: vertical; min-height: 80px; }
  label {
    font-size: 11px;
    font-weight: 700;
    color: ${T.textMuted};
    letter-spacing: 0.06em;
    text-transform: uppercase;
    display: block;
    margin-bottom: 5px;
  }
  @keyframes fadeUp {
    from { opacity:0; transform:translateY(14px); }
    to { opacity:1; transform:translateY(0); }
  }
  @keyframes spin {
    to { transform: rotate(360deg); }
  }
`;

/* ─── ROLES & PERMISSIONS ─────────────────────────────────────────────── */
const ROLES = {
  trabajador_social: { label:"Trabajador/a Social", color:"#4A9C9D", icon:"👤", level:1 },
  terapeuta:         { label:"Terapeuta",           color:"#6BBF8A", icon:"🧠", level:1 },
  administrativo:    { label:"Administrativo/a",    color:"#E8A87C", icon:"📋", level:1 },
  coordinador:       { label:"Coordinador/a",       color:"#B39DDB", icon:"🗂", level:2 },
  director:          { label:"Director/a",          color:"#D4B86A", icon:"⭐", level:3 },
  soporte:           { label:"Soporte Técnico",     color:"#E5989B", icon:"🛠", level:4 },
};

const PERMS = {
  trabajador_social: {
    dashboard:true, patients_view:true, patients_add:true, patients_edit:true, patients_del:false,
    followups_view:true, followups_add:false, followups_edit:false,
    donations_view:true, donations_add:true,
    projects_view:false, projects_add:false,
    stats:false, export:false, import:false, users:false,
  },
  terapeuta: {
    dashboard:true, patients_view:true, patients_add:false, patients_edit:false, patients_del:false,
    followups_view:true, followups_add:true, followups_edit:true,
    donations_view:false, donations_add:false,
    projects_view:false, projects_add:false,
    stats:false, export:false, import:false, users:false,
  },
  administrativo: {
    dashboard:true, patients_view:true, patients_add:false, patients_edit:false, patients_del:false,
    followups_view:true, followups_add:false, followups_edit:false,
    donations_view:true, donations_add:true,
    projects_view:true, projects_add:false,
    stats:true, export:true, import:false, users:false,
  },
  coordinador: {
    dashboard:true, patients_view:true, patients_add:true, patients_edit:true, patients_del:false,
    followups_view:true, followups_add:true, followups_edit:true,
    donations_view:true, donations_add:true,
    projects_view:true, projects_add:true,
    stats:true, export:true, import:false, users:false,
  },
  director: {
    dashboard:true, patients_view:true, patients_add:true, patients_edit:true, patients_del:true,
    followups_view:true, followups_add:true, followups_edit:true,
    donations_view:true, donations_add:true,
    projects_view:true, projects_add:true,
    stats:true, export:true, import:true, users:true,
  },
  soporte: {
    dashboard:true, patients_view:true, patients_add:true, patients_edit:true, patients_del:true,
    followups_view:true, followups_add:true, followups_edit:true,
    donations_view:true, donations_add:true,
    projects_view:true, projects_add:true,
    stats:true, export:true, import:true, users:true,
  },
};

const can = (user, perm) => user && PERMS[user.role] && PERMS[user.role][perm];

/* ─── UTILS ──────────────────────────────────────────────────────────────── */
const uid = () => Date.now() + Math.floor(Math.random()*9999);
const edad = f => { if(!f)return"–"; const d=new Date(f),h=new Date(); let a=h.getFullYear()-d.getFullYear(); if(h.getMonth()<d.getMonth()||(h.getMonth()===d.getMonth()&&h.getDate()<d.getDate()))a--; return a; };
const fmt = n => `$${(n||0).toLocaleString("es-CL")}`;
const fmtD = s => s?new Date(s+"T00:00:00").toLocaleDateString("es-CL"):"–";

/* ─── ATOMS ──────────────────────────────────────────────────────────────── */
const Btn = ({children,onClick,variant="primary",sm,icon,style,disabled})=>{
  const V={
    primary:{bg:`linear-gradient(135deg, ${T.primary}, ${T.primaryLt})`, color:"#fff", sh:`0 2px 12px ${T.primary}40`},
    accent:{bg:`linear-gradient(135deg, ${T.accent}, ${T.accentLt})`, color:"#fff", sh:`0 2px 12px ${T.accent}40`},
    danger:{bg:`linear-gradient(135deg, ${T.red}, #D88A8A)`, color:"#fff", sh:`0 2px 12px ${T.red}30`},
    ghost:{bg:"transparent", color:T.primary, border:`1.5px solid ${T.primary}`, sh:"none"},
    soft:{bg:T.bgSoft, color:T.primary, sh:"none"},
    gold:{bg:`linear-gradient(135deg, ${T.gold}, #E2C98A)`, color:"#fff", sh:`0 2px 12px ${T.gold}40`},
  };
  const v=V[variant]||V.primary;
  return <button disabled={disabled} onClick={onClick} style={{
    background:v.bg, color:v.color, border:v.border||"none", borderRadius:9,
    padding:sm?"6px 14px":"10px 20px", fontSize:sm?11.5:13, fontWeight:600,
    cursor:disabled?"not-allowed":"pointer", fontFamily:"'Plus Jakarta Sans',sans-serif",
    boxShadow:v.sh, display:"flex", alignItems:"center", gap:6,
    opacity:disabled?0.5:1, transition:"all 0.2s", whiteSpace:"nowrap", ...style
  }}
    onMouseEnter={e=>{if(!disabled){e.currentTarget.style.opacity="0.85"; e.currentTarget.style.transform="translateY(-1px)";}}}
    onMouseLeave={e=>{e.currentTarget.style.opacity="1"; e.currentTarget.style.transform="none";}}
  >{icon&&<span>{icon}</span>}{children}</button>;
};

const Badge=({text,color="#999"})=><span style={{
  background:color+"1e", color, borderRadius:20, padding:"3px 12px",
  fontSize:11, fontWeight:700, letterSpacing:"0.03em", whiteSpace:"nowrap",
  display:"inline-block"
}}>{text}</span>;


const estadoBadge = e => {
  const m = {
    Activo: T.green,
    Inactivo: T.red,
    Completado: T.blue,
    Planificado: T.gold,
    Suspendido: T.accent,
  };
  return <Badge text={e} color={m[e] || T.textMuted} />;
};

const tipoBadge = t => {
  const m = {
    Médico: T.blue,
    Terapia: T.green,
    Educativo: T.accent,
    Social: T.purple,
    Psicológico: T.teal,
  };
  return <Badge text={t} color={m[t] || T.textMuted} />;
};

const Card = ({children,style}) => <div style={{
  background:T.bgCard, borderRadius:16,
  boxShadow:"0 4px 20px rgba(0,0,0,0.04)", ...style
}}>{children}</div>;

const Modal = ({title,subtitle,onClose,children,wide}) => (
  <div style={{
    position:"fixed", inset:0,
    background:"rgba(30,61,61,0.5)",
    backdropFilter:"blur(6px)",
    display:"flex", alignItems:"center", justifyContent:"center",
    zIndex:1000, padding:20
  }}>
    <div style={{
      background:T.bgCard, borderRadius:18, width:"100%",
      maxWidth:wide?800:560, maxHeight:"90vh", overflowY:"auto",
      boxShadow:"0 30px 80px rgba(0,0,0,0.2)", animation:"fadeUp 0.25s ease"
    }}>
      <div style={{
        padding:"22px 28px 18px", borderBottom:`1px solid ${T.border}`,
        display:"flex", alignItems:"flex-start", justifyContent:"space-between"
      }}>
        <div>
          <div style={{fontFamily:"'Fraunces',serif",fontSize:19,fontWeight:700,color:T.text}}>{title}</div>
          {subtitle&&<div style={{fontSize:12,color:T.textMuted,marginTop:3}}>{subtitle}</div>}
        </div>
        <button onClick={onClose} style={{
          background:T.bgSoft, border:"none", borderRadius:8,
          width:32, height:32, cursor:"pointer", fontSize:16,
          color:T.textMuted, flexShrink:0
        }}>✕</button>
      </div>
      <div style={{padding:28}}>{children}</div>
    </div>
  </div>
);

const StatCard = ({icon,label,value,sub,color=T.primary}) => (
  <div style={{
    background:T.bgCard, borderRadius:16, padding:"22px 24px",
    boxShadow:"0 4px 20px rgba(0,0,0,0.04)",
    display:"flex", flexDirection:"column", gap:10,
    position:"relative", overflow:"hidden",
    animation:"fadeUp 0.3s ease both",
    borderBottom:`3px solid ${color}`
  }}>
    <div style={{position:"absolute", top:-10, right:-10, width:72, height:72, borderRadius:"50%", background:color+"0d"}}/>
    <div style={{background:color+"15", borderRadius:12, width:46, height:46, display:"flex", alignItems:"center", justifyContent:"center", fontSize:20}}>{icon}</div>
    <div>
      <div style={{fontSize:11,color:T.textMuted,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>{label}</div>
      <div style={{fontSize:26,fontWeight:700,color:T.text,lineHeight:1.15,fontFamily:"'Fraunces',serif",marginTop:2}}>{value}</div>
      {sub&&<div style={{fontSize:11.5,color:T.textMuted,marginTop:3}}>{sub}</div>}
    </div>
  </div>
);

const Field = ({label,children,span=2,hint}) => (
  <div style={{gridColumn:`span ${span}`}}>
    <label>{label}</label>
    {children}
    {hint&&<div style={{fontSize:11,color:T.textMuted,marginTop:4}}>{hint}</div>}
  </div>
);

/* ─── FORMS ──────────────────────────────────────────────────────────────── */
const EPAT={nombre:"",apellido:"",fecha_nac:"",diagnostico:"",condicion:"Intelectual",nivel:"Leve",telefono:"",email:"",direccion:"",tutor:"",relacionTutor:"Madre",telefonoTutor:"",fecha_ingreso:new Date().toISOString().slice(0,10),estado:"Activo",notas:""};
const PatientForm=({init,onSave,onClose})=>{
  const[f,sf]=useState(init||EPAT);
  const s=k=>e=>sf(p=>({...p,[k]:e.target.value}));
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Field label="Nombre" span={1}><input value={f.nombre} onChange={s("nombre")}/></Field>
      <Field label="Apellido" span={1}><input value={f.apellido} onChange={s("apellido")}/></Field>
      <Field label="Fecha de Nacimiento" span={1}><input type="date" value={f.fecha_nac} onChange={s("fecha_nac")}/></Field>
      <Field label="Diagnóstico" span={1}><input value={f.diagnostico} onChange={s("diagnostico")}/></Field>
      <Field label="Tipo de Condicion" span={1}><select value={f.condicion} onChange={s("condicion")}>{["Intelectual","Física","Sensorial","Cognitiva","Múltiple","Otra"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Nivel de Apoyo" span={1}><select value={f.nivel} onChange={s("nivel")}>{["Leve","Moderado","Severo"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Teléfono" span={1}><input value={f.telefono} onChange={s("telefono")}/></Field>
      <Field label="Email (opcional)" span={1}><input value={f.email} onChange={s("email")}/></Field>
      <Field label="Dirección"><input value={f.direccion} onChange={s("direccion")}/></Field>
      <div style={{gridColumn:"span 2",height:1,background:T.border,margin:"4px 0"}}/>
      <div style={{gridColumn:"span 2"}}><span style={{fontFamily:"'Fraunces',serif",fontWeight:600,fontSize:14,color:T.primary}}>Tutor / Representante Legal</span></div>
      <Field label="Nombre del Tutor" span={1}><input value={f.tutor} onChange={s("tutor")}/></Field>
      <Field label="Relación" span={1}><select value={f.relacionTutor} onChange={s("relacionTutor")}>{["Madre","Padre","Abuelo/a","Hermano/a","Tío/a","Tutor Legal","Otro"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Teléfono del Tutor" span={1}><input value={f.telefonoTutor} onChange={s("telefonoTutor")}/></Field>
      <Field label="Fecha de Ingreso" span={1}><input type="date" value={f.fecha_ingreso} onChange={s("fecha_ingreso")}/></Field>
      <Field label="Estado" span={1}><select value={f.estado} onChange={s("estado")}>{["Activo","Inactivo"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Notas / Observaciones"><textarea value={f.notas} onChange={s("notas")}/></Field>
      <div style={{gridColumn:"span 2",display:"flex",gap:10,justifyContent:"flex-end",marginTop:4}}>
        <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
        <Btn onClick={()=>onSave(f)}>Guardar Beneficiario</Btn>
      </div>
    </div>
  );
};

const EDON={donante:"",relacion:"Familiar",beneficiarioId:"",monto:"",fecha:new Date().toISOString().slice(0,10),tipo:"Transferencia",proyectoId:"",notas:""};
const DonationForm=({onSave,onClose,beneficiaries,projects})=>{
  const[f,sf]=useState(EDON);
  const s=k=>e=>sf(p=>({...p,[k]:e.target.value}));
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Field label="Nombre del Donante"><input value={f.donante} onChange={s("donante")}/></Field>
      <Field label="Tipo de Relación" span={1}><select value={f.relacion} onChange={s("relacion")}>{["Familiar","Externo","Corporativo","Anónimo"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Beneficiario Relacionado" span={1}><select value={f.beneficiarioId} onChange={s("beneficiarioId")}><option value="">Sin beneficiario específico</option>{beneficiaries.map(p=><option key={p.id} value={p.id}>{p.nombre} {p.apellido}</option>)}</select></Field>
      <Field label="Monto ($)" span={1}><input type="number" value={f.monto} onChange={s("monto")}/></Field>
      <Field label="Fecha" span={1}><input type="date" value={f.fecha} onChange={s("fecha")}/></Field>
      <Field label="Tipo de Pago" span={1}><select value={f.tipo} onChange={s("tipo")}>{["Transferencia","Efectivo","Cheque","Tarjeta","Especie"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Proyecto Destino"><select value={f.proyectoId} onChange={s("proyectoId")}><option value="">Sin proyecto específico</option>{projects.map(p=><option key={p.id} value={p.id}>{p.nombre}</option>)}</select></Field>
      <Field label="Notas" span={1}><textarea value={f.notas} onChange={s("notas")} style={{minHeight:60}}/></Field>
      <div style={{gridColumn:"span 2",display:"flex",gap:10,justifyContent:"flex-end"}}>
        <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
        <Btn variant="accent" onClick={()=>onSave({...f,monto:parseFloat(f.monto)||0,beneficiarioId:f.beneficiarioId?parseInt(f.beneficiarioId):null,proyectoId:f.proyectoId?parseInt(f.proyectoId):null})}>Registrar Donación</Btn>
      </div>
    </div>
  );
};

const EPROJ={nombre:"",descripcion:"",objetivo:"",recaudado:"0",inicio:new Date().toISOString().slice(0,10),fin:"",estado:"Planificado",beneficiarios:[]};
const ProjectForm=({init,onSave,onClose,beneficiaries})=>{
  const[f,sf]=useState(init||EPROJ);
  const s=k=>e=>sf(p=>({...p,[k]:e.target.value}));
  const tog=id=>sf(p=>({...p,beneficiarios:p.beneficiarios.includes(id)?p.beneficiarios.filter(x=>x!==id):[...p.beneficiarios,id]}));
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Field label="Nombre del Proyecto"><input value={f.nombre} onChange={s("nombre")}/></Field>
      <Field label="Descripción"><textarea value={f.descripcion} onChange={s("descripcion")} style={{minHeight:72}}/></Field>
      <Field label="Meta ($)" span={1}><input type="number" value={f.objetivo} onChange={s("objetivo")}/></Field>
      <Field label="Recaudado ($)" span={1}><input type="number" value={f.recaudado} onChange={s("recaudado")}/></Field>
      <Field label="Fecha Inicio" span={1}><input type="date" value={f.inicio} onChange={s("inicio")}/></Field>
      <Field label="Fecha Fin" span={1}><input type="date" value={f.fin} onChange={s("fin")}/></Field>
      <Field label="Estado" span={1}><select value={f.estado} onChange={s("estado")}>{["Planificado","Activo","Completado","Suspendido"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <div style={{gridColumn:"span 2"}}><label>Beneficiarios</label>
        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:6}}>
          {beneficiaries.map(p=>(
            <label key={p.id} style={{display:"flex",alignItems:"center",gap:7,textTransform:"none",letterSpacing:0,fontSize:12.5,color:T.text,cursor:"pointer",background:f.beneficiarios.includes(p.id)?T.primary+"15":T.bgSoft,borderRadius:9,padding:"7px 13px",border:`1.5px solid ${f.beneficiarios.includes(p.id)?T.primary:T.border}`,fontWeight:500}}>
              <input type="checkbox" checked={f.beneficiarios.includes(p.id)} onChange={()=>tog(p.id)} style={{width:"auto",margin:0}}/>
              {p.nombre} {p.apellido}
            </label>
          ))}
        </div>
      </div>
      <div style={{gridColumn:"span 2",display:"flex",gap:10,justifyContent:"flex-end"}}>
        <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
        <Btn onClick={()=>onSave({...f,objetivo:parseFloat(f.objetivo)||0,recaudado:parseFloat(f.recaudado)||0})}>Guardar Proyecto</Btn>
      </div>
    </div>
  );
};

const EFU={beneficiarioId:"",fecha:new Date().toISOString().slice(0,10),tipo:"Médico",descripcion:"",profesional:"",proxima:"",resultado:"Positivo"};
const FollowupForm=({init,onSave,onClose,beneficiaries})=>{
  const[f,sf]=useState(init||EFU);
  const s=k=>e=>sf(p=>({...p,[k]:e.target.value}));
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Field label="Beneficiario"><select value={f.beneficiarioId} onChange={s("beneficiarioId")}><option value="">Seleccionar beneficiario...</option>{beneficiaries.map(p=><option key={p.id} value={p.id}>{p.nombre} {p.apellido}</option>)}</select></Field>
      <Field label="Fecha" span={1}><input type="date" value={f.fecha} onChange={s("fecha")}/></Field>
      <Field label="Tipo" span={1}><select value={f.tipo} onChange={s("tipo")}>{["Médico","Terapia","Educativo","Social","Psicológico"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Profesional a Cargo" span={1}><input value={f.profesional} onChange={s("profesional")}/></Field>
      <Field label="Próxima Cita" span={1}><input type="date" value={f.proxima} onChange={s("proxima")}/></Field>
      <Field label="Resultado" span={1}><select value={f.resultado} onChange={s("resultado")}>{["Positivo","Neutral","Requiere Atención"].map(o=><option key={o}>{o}</option>)}</select></Field>
      <Field label="Descripción / Observaciones"><textarea value={f.descripcion} onChange={s("descripcion")}/></Field>
      <div style={{gridColumn:"span 2",display:"flex",gap:10,justifyContent:"flex-end"}}>
        <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
        <Btn onClick={()=>onSave({...f,beneficiarioId:parseInt(f.beneficiarioId)})}>Guardar Seguimiento</Btn>
      </div>
    </div>
  );
};

/* ─── IMPORT MODAL ───────────────────────────────────────────────────────── */
const ImportModal=({onClose,onImport})=>{
  const[tab,setTab]=useState("excel");
  const[driveFiles,setDriveFiles]=useState([]);
  const[driveLoading,setDriveLoading]=useState(false);
  const[driveMsg,setDriveMsg]=useState("");
  const[preview,setPreview]=useState(null);
  const[selectedFile,setSelectedFile]=useState(null);
  const[done,setDone]=useState(null);
  const fileRef=useRef();

  const parseXLSX=buf=>{
    const wb=XLSX.read(buf,{type:"array",cellDates:true});
    const sheets={};
    wb.SheetNames.forEach(n=>{sheets[n]=XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""});});
    return sheets;
  };

  const normalise=sheets=>{
    const res={beneficiaries:[],donations:[]};
    const ps=sheets["Beneficiarios"]||sheets["Beneficiarios"]||[];
    ps.forEach(r=>{
      if(!r["Nombre"]&&!r["nombre"])return;
      res.beneficiaries.push({
        id:uid(),
        nombre:r["Nombre"]||"",
        apellido:r["Apellido"]||"",
        fecha_nac:r["Fecha Nac."]||"",
        diagnostico:r["Diagnóstico"]||"",
        condicion:r["Condicion"]||"Otra",
        nivel:r["Nivel"]||"Leve",
        telefono:r["Teléfono"]||"",
        email:r["Email"]||"",
        direccion:r["Dirección"]||"",
        tutor:r["Tutor"]||"",
        relacionTutor:r["Relación"]||"",
        telefonoTutor:"",
        fecha_ingreso:r["Ingreso"]||new Date().toISOString().slice(0,10),
        estado:r["Estado"]||"Activo",
        notas:r["Notas"]||""
      });
    });
    const ds=sheets["Donaciones"]||sheets["donaciones"]||[];
    ds.forEach(r=>{
      if(!r["Donante"]&&!r["donante"])return;
      res.donations.push({
        id:uid(),
        donante:r["Donante"]||"",
        relacion:r["Relación"]||"Externo",
        beneficiarioId:null,
        monto:parseFloat(r["Monto ($)"]||0),
        fecha:r["Fecha"]||"",
        tipo:r["Tipo Pago"]||"Transferencia",
        proyectoId:null,
        notas:r["Notas"]||""
      });
    });
    return res;
  };

  const handleFile=async file=>{
    if(!file)return;
    setSelectedFile(file);
    const buf=await file.arrayBuffer();
    const sheets=parseXLSX(new Uint8Array(buf));
    const counts=Object.entries(sheets).map(([k,v])=>({sheet:k,rows:v.length}));
    setPreview({sheets,counts});
  };

  const searchDrive=async()=>{
    setDriveLoading(true);setDriveMsg("");setDriveFiles([]);
    try{
      const resp=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514",max_tokens:1000,
          system:"Search Google Drive for Excel files (.xlsx). Return ONLY a JSON array: [{\"id\":\"fileId\",\"name\":\"filename\",\"modifiedTime\":\"date\"}]. No markdown, no extra text.",
          messages:[{role:"user",content:"Use Google Drive MCP to list Excel (.xlsx or .xls) files. Return JSON array."}],
          mcp_servers:[{type:"url",url:"https://drivemcp.googleapis.com/mcp/v1",name:"gdrive"}],
        })
      });
      const data=await resp.json();
      const texts=data.content?.filter(c=>c.type==="text").map(c=>c.text).join("\n")||"";
      const match=texts.match(/\[[\s\S]*?\]/);
      if(match){const files=JSON.parse(match[0]);setDriveFiles(files.slice(0,10));}
      else{
        const toolResults=data.content?.filter(c=>c.type==="mcp_tool_result")||[];
        let found=[];
        toolResults.forEach(tr=>{
          try{const p=JSON.parse(tr.content?.[0]?.text||"");(p.files||[]).forEach(f=>{if(f.name?.match(/\.(xlsx|xls)$/i))found.push({id:f.id,name:f.name,modifiedTime:f.modifiedTime||""});});}catch{}
        });
        if(found.length)setDriveFiles(found);
        else setDriveMsg("No se encontraron archivos .xlsx en tu Drive. Sube un archivo exportado desde este sistema.");
      }
    }catch(e){setDriveMsg("Error al conectar con Drive: "+e.message);}
    setDriveLoading(false);
  };

  const confirmImport=()=>{
    if(!preview)return;
    const data=normalise(preview.sheets);
    onImport(data);
    setDone(data);
  };

  if(done)return(
    <Modal title="Importación Completada" onClose={onClose}>
      <div style={{textAlign:"center",padding:"24px 0"}}>
        <div style={{fontSize:52,marginBottom:14}}>✅</div>
        <div style={{fontFamily:"'Fraunces',serif",fontSize:20,fontWeight:700,marginBottom:8}}>¡Datos importados con éxito!</div>
        <div style={{color:T.textSub,fontSize:14,marginBottom:24}}>Se añadieron <strong>{done.beneficiaries.length}</strong> Beneficiarios y <strong>{done.donations.length}</strong> donaciones.</div>
        <Btn onClick={onClose}>Cerrar</Btn>
      </div>
    </Modal>
  );

  return(
    <Modal title="Importar Datos" subtitle="Desde archivo Excel o Google Drive" onClose={onClose} wide>
      <div style={{display:"flex",gap:4,background:T.bgSoft,borderRadius:11,padding:4,marginBottom:24}}>
        {[["excel","📄 Archivo Excel"],["drive","🟢 Google Drive"]].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{flex:1,padding:"9px 16px",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"'Plus Jakarta Sans',sans-serif",fontWeight:600,fontSize:13,transition:"all 0.18s",background:tab===id?T.bgCard:"transparent",color:tab===id?T.primary:T.textMuted,boxShadow:tab===id?"0 2px 8px rgba(0,0,0,0.08)":"none"}}>{label}</button>
        ))}
      </div>

      {tab==="excel"&&(
        <div>
          <div onClick={()=>fileRef.current?.click()}
            onDragOver={e=>{e.preventDefault();e.currentTarget.style.borderColor=T.primary;}}
            onDragLeave={e=>{e.currentTarget.style.borderColor=T.border;}}
            onDrop={e=>{e.preventDefault();e.currentTarget.style.borderColor=T.border;const f=e.dataTransfer.files[0];if(f)handleFile(f);}}
            style={{border:`2px dashed ${T.border}`,borderRadius:14,padding:"32px 20px",textAlign:"center",cursor:"pointer",transition:"all 0.18s",marginBottom:16}}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0])}/>
            <div style={{fontSize:36,marginBottom:10}}>📊</div>
            <div style={{fontWeight:600,fontSize:14,color:T.text,marginBottom:5}}>{selectedFile?selectedFile.name:"Arrastra tu archivo Excel aquí"}</div>
            <div style={{fontSize:12,color:T.textMuted}}>o haz clic para seleccionar · Compatible con .xlsx exportado de este sistema</div>
          </div>
          {preview&&(
            <div style={{background:T.bgSoft,borderRadius:12,padding:16,marginBottom:16}}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:10,color:T.primary}}>Vista previa del archivo</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                {preview.counts.map(({sheet,rows})=>(
                  <div key={sheet} style={{background:T.bgCard,borderRadius:9,padding:"9px 16px",border:`1px solid ${T.border}`}}>
                    <div style={{fontSize:10,color:T.textMuted,textTransform:"uppercase",fontWeight:700}}>{sheet}</div>
                    <div style={{fontSize:22,fontWeight:700,fontFamily:"'Fraunces',serif",color:T.primary}}>{rows}</div>
                    <div style={{fontSize:11,color:T.textMuted}}>registros</div>
                  </div>
                ))}
              </div>
              <div style={{fontSize:11,color:T.textMuted,marginTop:10}}>ℹ️ Se importarán hojas <strong>Beneficiarios</strong> y <strong>Donaciones</strong>. Otras hojas serán ignoradas.</div>
            </div>
          )}
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <Btn variant="ghost" onClick={onClose}>Cancelar</Btn>
            <Btn disabled={!preview} onClick={confirmImport} icon="📥">Importar Datos</Btn>
          </div>
        </div>
      )}

      {tab==="drive"&&(
        <div>
          <div style={{background:`${T.primary}10`,border:`1px solid ${T.primary}22`,borderRadius:12,padding:16,marginBottom:20}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
              <span style={{fontSize:20}}>🔗</span>
              <div>
                <div style={{fontWeight:700,fontSize:13,color:T.primary}}>Google Drive Conectado</div>
                <div style={{fontSize:11,color:T.textSub}}>Tu cuenta está vinculada a esta sesión de Claude</div>
              </div>
              <Badge text="Activo" color={T.green}/>
            </div>
            <div style={{fontSize:12,color:T.textSub}}>Busca archivos Excel en tu Drive. Exporta primero desde este sistema para garantizar compatibilidad de formato.</div>
          </div>
          <Btn onClick={searchDrive} disabled={driveLoading} icon={driveLoading?"⏳":"🔍"} style={{marginBottom:16}}>
            {driveLoading?"Buscando archivos...":"Buscar archivos Excel en Drive"}
          </Btn>
          {driveMsg&&<div style={{background:T.red+"10",border:`1px solid ${T.red}30`,borderRadius:10,padding:14,marginBottom:16,fontSize:13,color:T.red}}>{driveMsg}</div>}
          {driveFiles.length>0&&(
            <div>
              <div style={{fontSize:11,fontWeight:700,color:T.textMuted,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:10}}>Archivos encontrados en Drive</div>
              {driveFiles.map(f=>(
                <div key={f.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 16px",background:T.bgSoft,borderRadius:10,marginBottom:8,border:`1px solid ${T.border}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{fontSize:20}}>📊</span>
                    <div>
                      <div style={{fontWeight:600,fontSize:13}}>{f.name}</div>
                      {f.modifiedTime&&<div style={{fontSize:11,color:T.textMuted}}>Modificado: {new Date(f.modifiedTime).toLocaleDateString("es-CL")}</div>}
                    </div>
                  </div>
                  <Btn sm variant="soft" onClick={()=>setDriveMsg("Para importar este archivo, descárgalo como .xlsx desde Drive y usa la pestaña 'Archivo Excel'.")}>Importar</Btn>
                </div>
              ))}
            </div>
          )}
          {driveFiles.length===0&&!driveLoading&&!driveMsg&&(
            <div style={{textAlign:"center",padding:"30px 20px",color:T.textMuted}}>
              <div style={{fontSize:32,marginBottom:10}}>📂</div>
              <div style={{fontSize:13}}>Haz clic en "Buscar" para listar tus archivos Excel en Drive</div>
            </div>
          )}
        </div>
      )}
    </Modal>
  );
};

/* ─── USERS PANEL ───────────────────────────────────────────────────────── */
const UsersPanel = ({ users, onSaveUser, onToggleActive, onDeleteUser, currentUser }) => {
  const [showForm, setShowForm] = useState(false);
  const [editUser, setEditUser] = useState(null);
  const EUSER = { username:"", password:"1234", name:"", role:"administrativo", email:"", active:true };
  const [form, setForm] = useState(EUSER);
  const sf = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  const handleSave = () => {
    onSaveUser(form);
    setShowForm(false);
    setEditUser(null);
    setForm(EUSER);
  };

  const handleToggle = (id, active) => {
    onToggleActive(id, active);
  };

  return (
    <div style={{ animation:"fadeUp 0.3s ease" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:24 }}>
        <div>
          <h2 style={{ fontFamily:"'Fraunces',serif", fontSize:22, fontWeight:700 }}>Gestión de Usuarios</h2>
          <p style={{ fontSize:13, color:T.textMuted, marginTop:4 }}>Administra los perfiles y permisos del sistema · {users.filter(u=>u.active).length} usuarios activos</p>
        </div>
        <Btn onClick={()=>{setShowForm(true);setEditUser(null);setForm(EUSER);}} icon="＋">Nuevo Usuario</Btn>
      </div>

      <Card style={{ padding:24, marginBottom:20, overflow:"hidden" }}>
        <div style={{ fontFamily:"'Fraunces',serif", fontWeight:700, fontSize:15, marginBottom:16 }}>Matriz de Permisos por Rol</div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead>
              <tr style={{ background:T.bgSoft }}>
                <th style={{ padding:"10px 14px", textAlign:"left", fontWeight:700, color:T.textMuted, textTransform:"uppercase", fontSize:10.5, letterSpacing:"0.05em" }}>Módulo</th>
                {Object.entries(ROLES).map(([key,r])=>(
                  <th key={key} style={{ padding:"10px 12px", textAlign:"center", fontWeight:700, color:r.color, fontSize:10.5, whiteSpace:"nowrap" }}>
                    {r.icon} {r.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[
                ["Dashboard","dashboard"],
                ["Ver Beneficiarios","patients_view"],
                ["Agregar/Editar Beneficiarios","patients_add"],
                ["Eliminar Beneficiarios","patients_del"],
                ["Ver Seguimientos","followups_view"],
                ["Registrar Seguimientos","followups_add"],
                ["Ver Donaciones","donations_view"],
                ["Registrar Donaciones","donations_add"],
                ["Ver Proyectos","projects_view"],
                ["Gestionar Proyectos","projects_add"],
                ["Estadísticas","stats"],
                ["Exportar Excel","export"],
                ["Importar Datos","import"],
                ["Gestión de Usuarios","users"],
              ].map(([label, perm]) => (
                <tr key={perm} style={{ borderBottom:`1px solid ${T.border}` }}
                  onMouseEnter={e=>e.currentTarget.style.background=T.bgSoft}
                  onMouseLeave={e=>e.currentTarget.style.background=""}>
                  <td style={{ padding:"9px 14px", fontWeight:500, color:T.textSub }}>{label}</td>
                  {Object.keys(ROLES).map(role => (
                    <td key={role} style={{ padding:"9px 12px", textAlign:"center" }}>
                      {PERMS[role][perm]
                        ? <span style={{ color:T.green, fontSize:16 }}>✓</span>
                        : <span style={{ color:T.border, fontSize:16 }}>–</span>}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card style={{ overflow:"hidden" }}>
        <div style={{ padding:"16px 22px", borderBottom:`1px solid ${T.border}` }}>
          <div style={{ fontFamily:"'Fraunces',serif", fontWeight:700, fontSize:15 }}>Usuarios del Sistema</div>
        </div>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead>
              <tr style={{ background:T.bgSoft }}>
                {["Usuario","Nombre","Rol","Email","Estado","Acciones"].map(h=>(
                  <th key={h} style={{ padding:"11px 16px", textAlign:"left", fontSize:10.5, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.06em", whiteSpace:"nowrap" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {users.map(u => {
                const r = ROLES[u.role];
                const isSelf = u.id === currentUser.id;
                return (
                  <tr key={u.id} style={{ borderBottom:`1px solid ${T.border}` }}
                    onMouseEnter={e=>e.currentTarget.style.background=T.bgSoft}
                    onMouseLeave={e=>e.currentTarget.style.background=""}>
                    <td style={{ padding:"12px 16px", fontWeight:700, fontSize:13, fontFamily:"monospace", color:T.primary }}>{u.username}</td>
                    <td style={{ padding:"12px 16px", fontWeight:600, fontSize:13 }}>
                      {u.name} {isSelf && <span style={{ fontSize:10, color:T.gold, fontWeight:700, background:T.gold+"15", padding:"2px 7px", borderRadius:10, marginLeft:6 }}>Tú</span>}
                    </td>
                    <td style={{ padding:"12px 16px" }}>
                      <span style={{ display:"flex", alignItems:"center", gap:6, fontSize:12, fontWeight:600, color:r?.color }}>
                        <span>{r?.icon}</span>{r?.label}
                      </span>
                    </td>
                    <td style={{ padding:"12px 16px", fontSize:12, color:T.textMuted }}>{u.email}</td>
                    <td style={{ padding:"12px 16px" }}>
                      <button onClick={()=>handleToggle(u.id, !u.active)} disabled={isSelf} style={{ background:u.active?T.green+"15":T.red+"15", border:`1px solid ${u.active?T.green+"44":T.red+"44"}`, borderRadius:20, padding:"4px 12px", fontSize:11, fontWeight:700, color:u.active?T.green:T.red, cursor:isSelf?"not-allowed":"pointer" }}>
                        {u.active?"Activo":"Inactivo"}
                      </button>
                    </td>
                    <td style={{ padding:"12px 16px" }}>
                      <div style={{ display:"flex", gap:6 }}>
                        <button onClick={()=>{setEditUser(u);setForm({...u});setShowForm(true);}} style={{ background:"none", border:"none", cursor:"pointer", color:T.primary, fontSize:15, padding:"2px 5px" }}>✏️</button>
                        <button onClick={()=>onDeleteUser(u.id)} disabled={isSelf} style={{ background:"none", border:"none", cursor:isSelf?"not-allowed":"pointer", color:T.red, fontSize:15, padding:"2px 5px", opacity:isSelf?0.3:1 }}>🗑</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {showForm && (
        <Modal title={editUser?"Editar Usuario":"Nuevo Usuario"} onClose={()=>{setShowForm(false);setEditUser(null);setForm(EUSER);}}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
            <Field label="Nombre Completo"><input value={form.name} onChange={sf("name")}/></Field>
            <Field label="Nombre de Usuario" span={1}><input value={form.username} onChange={sf("username")}/></Field>
            <Field label="Contraseña" span={1} hint="Mínimo 4 caracteres"><input type="password" value={form.password} onChange={sf("password")}/></Field>
            <Field label="Email"><input value={form.email} onChange={sf("email")}/></Field>
            <Field label="Rol" span={1}>
              <select value={form.role} onChange={sf("role")}>
                {Object.entries(ROLES).map(([key,r])=><option key={key} value={key}>{r.icon} {r.label}</option>)}
              </select>
            </Field>
            <Field label="Estado" span={1}>
              <select value={form.active} onChange={e=>setForm(p=>({...p,active:e.target.value==="true"}))}>
                <option value="true">Activo</option><option value="false">Inactivo</option>
              </select>
            </Field>
            {form.role && (
              <div style={{ gridColumn:"span 2", background:T.bgSoft, borderRadius:10, padding:14 }}>
                <div style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:8 }}>Permisos del rol seleccionado</div>
                <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                  {Object.entries(PERMS[form.role]||{}).filter(([,v])=>v).map(([k])=>(
                    <span key={k} style={{ background:T.primary+"15", color:T.primary, borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:600 }}>{k}</span>
                  ))}
                </div>
              </div>
            )}
            <div style={{ gridColumn:"span 2", display:"flex", gap:10, justifyContent:"flex-end" }}>
              <Btn variant="ghost" onClick={()=>{setShowForm(false);setEditUser(null);setForm(EUSER);}}>Cancelar</Btn>
              <Btn onClick={handleSave}>Guardar Usuario</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
};

/* ─── LOGIN SCREEN ──────────────────────────────────────────────────────── */
const LoginScreen = ({ onLogin }) => {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [showPass, setShowPass] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const user = await onLogin(username, password);
      if (user) {
        // El login se maneja en App, solo recibimos el usuario
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const quickLogin = (u) => {
    setUsername(u.username);
    setPassword(u.password);
    setError("");
  };

  return (
    <div style={{
      minHeight:"100vh",
      background:`linear-gradient(135deg, #1A3A3A 0%, #2C5A5A 50%, #1A3A3A 100%)`,
      display:"flex", alignItems:"center", justifyContent:"center", padding:20,
      fontFamily:"'Plus Jakarta Sans',sans-serif"
    }}>
      <div style={{
        position:"fixed", inset:0,
        backgroundImage:"radial-gradient(circle at 20% 50%, rgba(74,156,157,0.08) 0%, transparent 50%), radial-gradient(circle at 80% 20%, rgba(232,168,124,0.06) 0%, transparent 40%)",
        pointerEvents:"none"
      }}/>
      <div style={{
        width:"100%", maxWidth:940,
        display:"grid", gridTemplateColumns:"1fr 1fr", gap:0,
        borderRadius:22, overflow:"hidden",
        boxShadow:"0 40px 100px rgba(0,0,0,0.4)",
        animation:"fadeUp 0.4s ease"
      }}>
        <div style={{
          background:`linear-gradient(160deg, #1E3D3D, #2C5A5A)`,
          padding:"52px 44px",
          display:"flex", flexDirection:"column", justifyContent:"space-between"
        }}>
          <div>
            <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:48 }}>
              <div style={{
                width:46, height:46,
                background:`linear-gradient(135deg, ${T.primaryLt}, ${T.accent})`,
                borderRadius:13,
                display:"flex", alignItems:"center", justifyContent:"center",
                fontSize:22
              }}>🌿</div>
              <div>
                <div style={{ fontFamily:"'Fraunces',serif", color:"#fff", fontSize:18, fontWeight:700 }}>Fundación Crescendo</div>
                <div style={{ color:"#80CBC4", fontSize:11, fontWeight:600, letterSpacing:"0.08em", textTransform:"uppercase" }}>Sistema de Gestión Integral</div>
              </div>
            </div>
            <div style={{ fontFamily:"'Fraunces',serif", fontSize:30, fontWeight:700, color:"#fff", lineHeight:1.3, marginBottom:16 }}>
              Bienvenido/a de vuelta
            </div>
            <div style={{ fontSize:14, color:"rgba(255,255,255,0.6)", lineHeight:1.7 }}>
              Plataforma centralizada para el seguimiento integral de beneficiarios, gestión de proyectos y donaciones.
            </div>
          </div>
          <div>
            <div style={{ fontSize:11, color:"rgba(255,255,255,0.4)", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>
              Acceso rápido (demo)
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
              {Object.entries(ROLES).map(([key, r]) => {
                const u = { username: key, password: "1234", name: r.label, role: key, email: `${key}@crescendo.cl`, active: true };
                return (
                  <button key={key} onClick={() => quickLogin(u)} style={{
                    background:"rgba(255,255,255,0.05)",
                    border:"1px solid rgba(255,255,255,0.08)",
                    borderRadius:10,
                    padding:"9px 14px",
                    display:"flex", alignItems:"center", gap:10,
                    cursor:"pointer", transition:"all 0.15s", textAlign:"left"
                  }}
                    onMouseEnter={e => { e.currentTarget.style.background="rgba(255,255,255,0.1)"; e.currentTarget.style.borderColor="rgba(255,255,255,0.2)"; }}
                    onMouseLeave={e => { e.currentTarget.style.background="rgba(255,255,255,0.05)"; e.currentTarget.style.borderColor="rgba(255,255,255,0.08)"; }}
                  >
                    <span style={{ fontSize:16 }}>{r.icon}</span>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:12, fontWeight:600, color:"#fff" }}>{r.label}</div>
                      <div style={{ fontSize:10, color:"rgba(255,255,255,0.4)" }}>{key} / 1234</div>
                    </div>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:r.color }} />
                  </button>
                );
              })}
            </div>
          </div>
        </div>
        <div style={{ background:T.bgCard, padding:"52px 44px", display:"flex", flexDirection:"column", justifyContent:"center" }}>
          <div style={{ fontFamily:"'Fraunces',serif", fontSize:24, fontWeight:700, color:T.text, marginBottom:6 }}>Iniciar Sesión</div>
          <div style={{ fontSize:13, color:T.textMuted, marginBottom:36 }}>Ingresa tus credenciales para continuar</div>
          <form onSubmit={handleSubmit}>
            <div style={{ marginBottom:18 }}>
              <label>Usuario</label>
              <div style={{ position:"relative" }}>
                <span style={{ position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", fontSize:16 }}>👤</span>
                <input value={username} onChange={e=>setUsername(e.target.value)} placeholder="tu_usuario" style={{ paddingLeft:38 }} />
              </div>
            </div>
            <div style={{ marginBottom:24 }}>
              <label>Contraseña</label>
              <div style={{ position:"relative" }}>
                <span style={{ position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", fontSize:16 }}>🔒</span>
                <input type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" style={{ paddingLeft:38, paddingRight:40 }} />
                <button onClick={()=>setShowPass(s=>!s)} style={{ position:"absolute", right:12, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", cursor:"pointer", fontSize:15, color:T.textMuted }}>
                  {showPass?"🙈":"👁"}
                </button>
              </div>
            </div>
            {error && (
              <div style={{ background:T.red+"12", border:`1px solid ${T.red}33`, borderRadius:10, padding:"11px 14px", marginBottom:18, fontSize:13, color:T.red, display:"flex", alignItems:"center", gap:8 }}>
                ⚠️ {error}
              </div>
            )}
            <button type="submit" disabled={loading || !username || !password} style={{
              width:"100%", padding:"13px",
              background:loading||!username||!password?"#aaa":`linear-gradient(135deg, ${T.primary}, ${T.primaryLt})`,
              border:"none", borderRadius:11, color:"#fff",
              fontFamily:"'Plus Jakarta Sans',sans-serif", fontSize:14, fontWeight:700,
              cursor:loading||!username||!password?"not-allowed":"pointer",
              boxShadow:`0 4px 16px ${T.primary}44`,
              transition:"all 0.2s",
              display:"flex", alignItems:"center", justifyContent:"center", gap:8
            }}>
              {loading ? <><span style={{ display:"inline-block", width:16, height:16, border:"2.5px solid rgba(255,255,255,0.3)", borderTopColor:"#fff", borderRadius:"50%", animation:"spin 0.7s linear infinite" }} /> Verificando...</> : "Ingresar al Sistema →"}
            </button>
          </form>
          <div style={{ marginTop:32, padding:"16px", background:T.bgSoft, borderRadius:12 }}>
            <div style={{ fontSize:11, color:T.textMuted, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>Contraseña demo</div>
            <div style={{ fontSize:13, color:T.textSub }}>Todos los usuarios utilizan la contraseña <strong>1234</strong>. Selecciona un usuario del panel izquierdo para acceso rápido.</div>
          </div>
        </div>
      </div>
    </div>
  );
};

/* ─── MAIN APP ───────────────────────────────────────────────────────────── */
const NAV_ALL = [
  { id:"dashboard", icon:"🏠", label:"Dashboard", perm:"dashboard" },
  { id:"beneficiaries",  icon:"👥", label:"Beneficiarios",  perm:"patients_view" },
  { id:"projects",  icon:"📁", label:"Proyectos",  perm:"projects_view" },
  { id:"donations", icon:"💚", label:"Donaciones", perm:"donations_view" },
  { id:"followups", icon:"📋", label:"Seguimientos",perm:"followups_view" },
  { id:"stats",     icon:"📊", label:"Estadísticas",perm:"stats" },
  { id:"users",     icon:"⚙️", label:"Usuarios",   perm:"users" },
];

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [view, setView] = useState("dashboard");
  const [beneficiaries, setPatients] = useState([]);
  const [projects, setProjects] = useState([]);
  const [donations, setDonations] = useState([]);
  const [followups, setFollowups] = useState([]);
  const [patModal, setPatModal] = useState(null);
  const [donModal, setDonModal] = useState(false);
  const [projModal, setProjModal] = useState(null);
  const [fuModal, setFuModal] = useState(null);
  const [detailPat, setDetailPat] = useState(null);
  const [importModal, setImportModal] = useState(false);
  const [search, setSearch] = useState("");
  const [filterDisc, setFilterDisc] = useState("Todos");
  const [filterEst, setFilterEst] = useState("Todos");
  const [filterFuPatient, setFilterFuPatient] = useState("Todos");
  const [sideOpen, setSideOpen] = useState(true);
  const [showProfile, setShowProfile] = useState(false);

  // --- Cargar datos desde Supabase ---
  const fetchData = async () => {
    setLoading(true);
    try {
      const { data: patientsData, error: patientsError } = await supabase
        .from('beneficiaries')
        .select('*')
        .order('id');
      if (patientsError) throw patientsError;
      setPatients(patientsData || []);

      const { data: projectsData, error: projectsError } = await supabase
        .from('projects')
        .select('*')
        .order('id');
      if (projectsError) throw projectsError;
      setProjects(projectsData || []);

      const { data: donationsData, error: donationsError } = await supabase
        .from('donations')
        .select('*')
        .order('fecha', { ascending: false });
      if (donationsError) throw donationsError;
      setDonations(donationsData || []);

      const { data: followupsData, error: followupsError } = await supabase
        .from('followups')
        .select('*')
        .order('fecha', { ascending: false });
      if (followupsError) throw followupsError;
      setFollowups(followupsData || []);

      if (currentUser && can(currentUser, 'users')) {
        const { data: usersData, error: usersError } = await supabase
          .from('users')
          .select('*')
          .order('id');
        if (usersError) throw usersError;
        setUsers(usersData || []);
      }
    } catch (error) {
      console.error('Error al cargar datos:', error);
      alert('Error al cargar datos: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // --- Autenticación ---
  const login = async (username, password) => {
    try {
      const { data, error } = await supabase
        .from('users')
        .select('*')
        .eq('username', username)
        .eq('password', password)
        .eq('active', true)
        .single();

      if (error || !data) {
        throw new Error('Usuario o contraseña incorrectos');
      }

      setCurrentUser(data);
      return data;
    } catch (error) {
      throw error;
    }
  };

  const logout = () => {
    setCurrentUser(null);
  };

  useEffect(() => {
    if (currentUser) {
      fetchData();
    }
  }, [currentUser]);

  // ─── CRUD Operations ──────────────────────────────────────────────────

  // Beneficiarios
  const savePat = async (patient) => {
    try {
      let result;
      if (patient.id) {
        const { data, error } = await supabase
          .from('beneficiaries')
          .update(patient)
          .eq('id', patient.id)
          .select()
          .single();
        if (error) throw error;
        result = data;
        setPatients(prev => prev.map(p => p.id === result.id ? result : p));
      } else {
        const { data, error } = await supabase
          .from('beneficiaries')
          .insert([patient])
          .select()
          .single();
        if (error) throw error;
        result = data;
        setPatients(prev => [...prev, result]);
      }
      setPatModal(null);
      return result;
    } catch (error) {
      alert('Error al guardar beneficiario: ' + error.message);
    }
  };

  const delPat = async (id) => {
    if (!window.confirm('¿Eliminar beneficiario? Esto eliminará también todos sus seguimientos y donaciones.')) return;
    try {
      const { error: fuError } = await supabase
        .from('followups')
        .delete()
        .eq('beneficiarioId', id);
      if (fuError) throw fuError;

      const { error: donError } = await supabase
        .from('donations')
        .delete()
        .eq('beneficiarioId', id);
      if (donError) throw donError;

      const { error: patError } = await supabase
        .from('beneficiaries')
        .delete()
        .eq('id', id);
      if (patError) throw patError;

      setPatients(prev => prev.filter(p => p.id !== id));
      setDetailPat(null);
      alert('Beneficiario y sus datos asociados eliminados correctamente.');
    } catch (error) {
      alert('Error al eliminar beneficiario: ' + error.message);
    }
  };

  // PROYECTOS
  const saveProj = async (project) => {
    try {
      let result;
      if (project.id) {
        const { data, error } = await supabase
          .from('projects')
          .update(project)
          .eq('id', project.id)
          .select()
          .single();
        if (error) throw error;
        result = data;
        setProjects(prev => prev.map(p => p.id === result.id ? result : p));
      } else {
        const { data, error } = await supabase
          .from('projects')
          .insert([project])
          .select()
          .single();
        if (error) throw error;
        result = data;
        setProjects(prev => [...prev, result]);
      }
      setProjModal(null);
      return result;
    } catch (error) {
      alert('Error al guardar proyecto: ' + error.message);
    }
  };

  const delProj = async (id) => {
    if (!window.confirm('¿Eliminar proyecto?')) return;
    try {
      const { error } = await supabase
        .from('projects')
        .delete()
        .eq('id', id);
      if (error) throw error;
      setProjects(prev => prev.filter(p => p.id !== id));
    } catch (error) {
      alert('Error al eliminar proyecto: ' + error.message);
    }
  };

  // DONACIONES
  const saveDon = async (donation) => {
    try {
      const { data, error } = await supabase
        .from('donations')
        .insert([donation])
        .select()
        .single();
      if (error) throw error;
      setDonations(prev => [data, ...prev]);
      setDonModal(false);
      return data;
    } catch (error) {
      alert('Error al registrar donación: ' + error.message);
    }
  };

  // SEGUIMIENTOS
  const saveFu = async (followup) => {
    try {
      let result;
      if (followup.id) {
        const { data, error } = await supabase
          .from('followups')
          .update(followup)
          .eq('id', followup.id)
          .select()
          .single();
        if (error) throw error;
        result = data;
        setFollowups(prev => prev.map(f => f.id === result.id ? result : f));
      } else {
        const { data, error } = await supabase
          .from('followups')
          .insert([followup])
          .select()
          .single();
        if (error) throw error;
        result = data;
        setFollowups(prev => [result, ...prev]);
      }
      setFuModal(null);
      return result;
    } catch (error) {
      alert('Error al guardar seguimiento: ' + error.message);
    }
  };

  const delFu = async (id) => {
    if (!window.confirm('¿Eliminar este seguimiento?')) return;
    try {
      const { error } = await supabase
        .from('followups')
        .delete()
        .eq('id', id);
      if (error) throw error;
      setFollowups(prev => prev.filter(f => f.id !== id));
    } catch (error) {
      alert('Error al eliminar seguimiento: ' + error.message);
    }
  };

  // USUARIOS
  const saveUser = async (user) => {
    try {
      let result;
      if (user.id) {
        const { data, error } = await supabase
          .from('users')
          .update(user)
          .eq('id', user.id)
          .select()
          .single();
        if (error) throw error;
        result = data;
        setUsers(prev => prev.map(u => u.id === result.id ? result : u));
      } else {
        const { data, error } = await supabase
          .from('users')
          .insert([user])
          .select()
          .single();
        if (error) throw error;
        result = data;
        setUsers(prev => [...prev, result]);
      }
      return result;
    } catch (error) {
      alert('Error al guardar usuario: ' + error.message);
    }
  };

  const toggleUserActive = async (id, active) => {
    try {
      const { error } = await supabase
        .from('users')
        .update({ active })
        .eq('id', id);
      if (error) throw error;
      setUsers(prev => prev.map(u => u.id === id ? { ...u, active } : u));
    } catch (error) {
      alert('Error al cambiar estado del usuario: ' + error.message);
    }
  };

  const deleteUser = async (id) => {
    if (!window.confirm('¿Eliminar usuario?')) return;
    try {
      const { error } = await supabase
        .from('users')
        .delete()
        .eq('id', id);
      if (error) throw error;
      setUsers(prev => prev.filter(u => u.id !== id));
    } catch (error) {
      alert('Error al eliminar usuario: ' + error.message);
    }
  };

  // IMPORTAR
  const handleImport = async ({ beneficiaries: newPatients, donations: newDonations }) => {
    try {
      if (newPatients.length) {
        const { error } = await supabase
          .from('beneficiaries')
          .insert(newPatients);
        if (error) throw error;
        const { data } = await supabase.from('beneficiaries').select('*').order('id');
        setPatients(data || []);
      }
      if (newDonations.length) {
        const { error } = await supabase
          .from('donations')
          .insert(newDonations);
        if (error) throw error;
        const { data } = await supabase.from('donations').select('*').order('fecha', { ascending: false });
        setDonations(data || []);
      }
      setImportModal(false);
    } catch (error) {
      alert('Error al importar datos: ' + error.message);
    }
  };

  // EXPORTAR EXCEL
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      beneficiaries.map(p => ({
        Nombre: `${p.nombre} ${p.apellido}`,
        "Fecha Nac.": p.fecha_nac,
        Edad: edad(p.fecha_nac),
        Diagnóstico: p.diagnostico,
        Condicion: p.condicion,
        Nivel: p.nivel,
        Tutor: p.tutor,
        Relación: p.relacionTutor,
        Teléfono: p.telefono,
        Email: p.email,
        Dirección: p.direccion,
        Ingreso: p.fecha_ingreso,
        Estado: p.estado,
        Notas: p.notas
      }))
    ), "Beneficiarios");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      donations.map(d => {
        const pac = beneficiaries.find(p => p.id === d.beneficiarioId);
        const proj = projects.find(p => p.id === d.proyectoId);
        return {
          Donante: d.donante,
          Relación: d.relacion,
          Beneficiario: pac ? `${pac.nombre} ${pac.apellido}` : "–",
          "Monto ($)": d.monto,
          Fecha: d.fecha,
          "Tipo Pago": d.tipo,
          Proyecto: proj ? proj.nombre : "–",
          Notas: d.notas
        };
      })
    ), "Donaciones");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      projects.map(p => ({
        Nombre: p.nombre,
        Descripción: p.descripcion,
        "Meta ($)": p.objetivo,
        "Recaudado ($)": p.recaudado,
        Inicio: p.inicio,
        Fin: p.fin,
        Estado: p.estado
      }))
    ), "Proyectos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      followups.map(f => {
        const pac = beneficiaries.find(p => p.id === f.beneficiarioId);
        return {
          Beneficiario: pac ? `${pac.nombre} ${pac.apellido}` : "–",
          Fecha: f.fecha,
          Tipo: f.tipo,
          Descripción: f.descripcion,
          Profesional: f.profesional,
          "Próxima Cita": f.proxima,
          Resultado: f.resultado
        };
      })
    ), "Seguimientos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { Indicador:"Total Beneficiarios", Valor:beneficiaries.length },
      { Indicador:"Beneficiarios Activos", Valor:beneficiaries.filter(p => p.estado === "Activo").length },
      { Indicador:"Total Donaciones ($)", Valor:donations.reduce((s,d) => s + d.monto, 0) },
      { Indicador:"Proyectos Activos", Valor:projects.filter(p => p.estado === "Activo").length },
      { Indicador:"Total Seguimientos", Valor:followups.length }
    ]), "Estadísticas");
    XLSX.writeFile(wb, `crescendo_${new Date().toISOString().slice(0,10)}.xlsx`);
  };
  // ─── ExportPDF ──────────────────────────────────────────────────────────
  const pdfContentRef = useRef(null);

  const renderPDFContent = () => {
    if (!detailPat) return null;
    
    const p = beneficiaries.find(x => x.id === detailPat.id) || detailPat;
    const patFu = followups.filter(f => f.beneficiarioId === p.id);
    const patDon = donations.filter(d => d.beneficiarioId === p.id);
    const patProj = projects.filter(pr => pr.beneficiarios && pr.beneficiarios.includes(p.id));

    return (
      <div ref={pdfContentRef} style={{
        padding: '40px',
        fontFamily: "'Plus Jakarta Sans', sans-serif",
        background: '#FFFFFF',
        color: '#1E3D3D',
        maxWidth: '800px',
        margin: '0 auto'
      }}>
        {/* Header */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: '16px',
          borderBottom: '3px solid #2C7A7B',
          paddingBottom: '20px',
          marginBottom: '24px'
        }}>
          <div style={{
            width: '48px',
            height: '48px',
            background: 'linear-gradient(135deg, #2C7A7B, #E8A87C)',
            borderRadius: '12px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '24px',
            color: '#fff'
          }}>🌿</div>
          <div>
            <div style={{ fontSize: '22px', fontWeight: 700, fontFamily: "'Fraunces', serif", color: '#2C7A7B' }}>
              Fundación Crescendo
            </div>
            <div style={{ fontSize: '12px', color: '#6B7A7A', fontWeight: 500, letterSpacing: '0.5px' }}>
              INFORME DE BENEFICIARIO
            </div>
          </div>
          <div style={{ marginLeft: 'auto', textAlign: 'right' }}>
            <div style={{ fontSize: '11px', color: '#6B7A7A' }}>Fecha de emisión</div>
            <div style={{ fontSize: '12px', fontWeight: 600 }}>{new Date().toLocaleDateString('es-CL')}</div>
          </div>
        </div>

        {/* Información personal */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px', marginBottom: '28px' }}>
          <div>
            <div style={{ fontSize: '26px', fontWeight: 700, fontFamily: "'Fraunces', serif", color: '#1E3D3D' }}>
              {p.nombre} {p.apellido}
            </div>
            <div style={{ fontSize: '14px', color: '#5A6B6B', marginTop: '4px' }}>
              {edad(p.fecha_nac)} años · {fmtD(p.fecha_nac)}
            </div>
            <div style={{ marginTop: '8px' }}>
              <span style={{
                background: p.estado === 'Activo' ? '#6BBF8A' : '#E5989B',
                color: '#fff',
                padding: '4px 14px',
                borderRadius: '20px',
                fontSize: '12px',
                fontWeight: 700,
                display: 'inline-block'
              }}>{p.estado}</span>
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', fontSize: '13px' }}>
            <div><span style={{ color: '#6B7A7A' }}>Diagnóstico:</span> <strong>{p.diagnostico}</strong></div>
            <div><span style={{ color: '#6B7A7A' }}>Condición:</span> <strong>{p.condicion}</strong></div>
            <div><span style={{ color: '#6B7A7A' }}>Nivel:</span> <strong>{p.nivel}</strong></div>
            <div><span style={{ color: '#6B7A7A' }}>Teléfono:</span> <strong>{p.telefono || '–'}</strong></div>
            <div style={{ gridColumn: 'span 2' }}><span style={{ color: '#6B7A7A' }}>Email:</span> <strong>{p.email || '–'}</strong></div>
            <div style={{ gridColumn: 'span 2' }}><span style={{ color: '#6B7A7A' }}>Dirección:</span> <strong>{p.direccion || '–'}</strong></div>
          </div>
        </div>

        {/* Tutor */}
        <div style={{
          background: '#F4F7F6',
          borderRadius: '12px',
          padding: '16px 20px',
          marginBottom: '24px'
        }}>
          <div style={{ fontSize: '11px', fontWeight: 700, color: '#6B7A7A', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '4px' }}>
            Tutor / Representante
          </div>
          <div style={{ fontSize: '16px', fontWeight: 600 }}>{p.tutor}</div>
          <div style={{ fontSize: '13px', color: '#5A6B6B' }}>{p.relacionTutor} · {p.telefonoTutor}</div>
          {p.notas && <div style={{ fontSize: '13px', marginTop: '6px', color: '#5A6B6B' }}>📝 {p.notas}</div>}
        </div>

        {/* Seguimientos */}
        {patFu.length > 0 && (
          <div style={{ marginBottom: '24px' }}>
            <div style={{
              fontSize: '16px',
              fontWeight: 700,
              color: '#2C7A7B',
              borderBottom: '2px solid #E8F0EF',
              paddingBottom: '8px',
              marginBottom: '12px'
            }}>
              Seguimientos ({patFu.length})
            </div>
            {patFu.sort((a, b) => b.fecha.localeCompare(a.fecha)).map(f => {
              const tc = { Médico: '#A8DADC', Terapia: '#6BBF8A', Educativo: '#E8A87C', Social: '#B39DDB', Psicológico: '#80CBC4' }[f.tipo] || '#2C7A7B';
              return (
                <div key={f.id} style={{
                  display: 'flex',
                  gap: '12px',
                  padding: '12px 0',
                  borderBottom: '1px solid #E8F0EF'
                }}>
                  <div style={{
                    width: '4px',
                    background: tc,
                    borderRadius: '4px',
                    flexShrink: 0
                  }} />
                  <div style={{ flex: 1 }}>
                    <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                      <span style={{
                        background: tc,
                        color: '#fff',
                        padding: '2px 10px',
                        borderRadius: '12px',
                        fontSize: '11px',
                        fontWeight: 700
                      }}>{f.tipo}</span>
                      <span style={{ fontSize: '12px', color: '#6B7A6B' }}>{fmtD(f.fecha)}</span>
                      <span style={{
                        background: f.resultado === 'Positivo' ? '#6BBF8A' : f.resultado === 'Neutral' ? '#D4B86A' : '#E5989B',
                        color: '#fff',
                        padding: '2px 10px',
                        borderRadius: '12px',
                        fontSize: '11px',
                        fontWeight: 700
                      }}>{f.resultado}</span>
                    </div>
                    <div style={{ fontSize: '13px', marginTop: '4px' }}>{f.descripcion}</div>
                    <div style={{ fontSize: '12px', color: '#6B7A6B', marginTop: '2px' }}>
                      {f.profesional} · Próx: {fmtD(f.proxima)}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Footer */}
        <div style={{
          marginTop: '32px',
          paddingTop: '16px',
          borderTop: '1px solid #E8F0EF',
          fontSize: '11px',
          color: '#8FA3A3',
          textAlign: 'center'
        }}>
          Documento generado desde el Sistema de Gestión Fundación Crescendo · {new Date().toLocaleString('es-CL')}
        </div>
      </div>
    );
  };
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  const exportPDF = async () => {
  if (!detailPat) return;
  
  try {
    // Esperar un momento para que el contenido oculto se renderice
    await new Promise(resolve => setTimeout(resolve, 200));
    
    const element = pdfContentRef.current;
    if (!element) {
      alert('Error: No se pudo generar el PDF');
      return;
    }
    
    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#FFFFFF',
      logging: false
    });
    
    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    
    // Si el contenido es más alto que una página, ajustamos
    if (pdfHeight > pdf.internal.pageSize.getHeight()) {
      // Para contenido largo, podemos usar varias páginas o escalar
      // Por simplicidad, escalamos para que quepa en una página
      const scaleFactor = pdf.internal.pageSize.getHeight() / pdfHeight;
      const newWidth = pdfWidth * scaleFactor;
      const newHeight = pdf.internal.pageSize.getHeight();
      pdf.addImage(imgData, 'PNG', (pdfWidth - newWidth) / 2, 0, newWidth, newHeight);
    } else {
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
    }
    
    pdf.save(`perfil_${detailPat.nombre}_${detailPat.apellido}.pdf`);
  } catch (error) {
    console.error('Error al generar PDF:', error);
    alert('Error al generar el PDF. Inténtalo de nuevo.');
  }
};

  // ─── UseMemo ──────────────────────────────────────────────────────────
  const discData = useMemo(() => {
    const c = {};
    beneficiaries.forEach(p => { c[p.condicion] = (c[p.condicion] || 0) + 1; });
    return Object.entries(c).map(([n, v]) => ({ name: n, value: v }));
  }, [beneficiaries]);

  const donByMonth = useMemo(() => {
    const c = {};
    donations.forEach(d => {
      const m = d.fecha.slice(0, 7);
      c[m] = (c[m] || 0) + d.monto;
    });
    return Object.entries(c).sort().slice(-6).map(([k, v]) => ({
      mes: k.slice(5) + "/" + k.slice(2, 4),
      monto: v
    }));
  }, [donations]);

  const fuByType = useMemo(() => {
    const c = {};
    followups.forEach(f => { c[f.tipo] = (c[f.tipo] || 0) + 1; });
    return Object.entries(c).map(([n, v]) => ({ name: n, value: v }));
  }, [followups]);

  const resultData = useMemo(() => {
    const c = { Positivo: 0, Neutral: 0, "Requiere Atención": 0 };
    followups.forEach(f => { c[f.resultado] = (c[f.resultado] || 0) + 1; });
    return Object.entries(c).map(([n, v]) => ({ name: n, value: v }));
  }, [followups]);

  // ─── Estadísticas ──────────────────────────────────────────────────────
  const totalDon = donations.reduce((s,d) => s + d.monto, 0);
  const activeP = beneficiaries.filter(p => p.estado === "Activo").length;
  const activePr = projects.filter(p => p.estado === "Activo").length;
  const now = new Date();
  const fuMonth = followups.filter(f => {
    const d = new Date(f.fecha);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  }).length;

  const nav = NAV_ALL.filter(n => can(currentUser, n.perm));
  const roleInfo = currentUser ? ROLES[currentUser.role] : null;
  const discTypes = ["Todos", ...new Set(beneficiaries.map(p => p.condicion))];

  const filtPat = beneficiaries.filter(p => {
    const s = `${p.nombre} ${p.apellido} ${p.diagnostico} ${p.tutor}`.toLowerCase().includes(search.toLowerCase());
    const d = filterDisc === "Todos" || p.condicion === filterDisc;
    const e = filterEst === "Todos" || p.estado === filterEst;
    return s && d && e;
  });

  // ─── Tooltip ───────────────────────────────────────────────────────────
  const TTip = ({ active, payload, label }) => {
    if (!active || !payload?.length) return null;
    return (
      <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px", boxShadow: "0 6px 24px rgba(0,0,0,0.1)", fontSize: 12 }}>
        {label && <div style={{ fontWeight: 700, marginBottom: 4 }}>{label}</div>}
        {payload.map((e, i) => (
          <div key={i} style={{ color: e.color || T.primary }}>
            {e.name}: <strong>{typeof e.value === "number" && e.value > 100 ? fmt(e.value) : e.value}</strong>
          </div>
        ))}
      </div>
    );
  };

  const AccessDenied = () => (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "60px 20px", animation: "fadeUp 0.3s ease" }}>
      <div style={{ fontSize: 52, marginBottom: 16 }}>🔒</div>
      <div style={{ fontFamily: "'Fraunces',serif", fontSize: 20, fontWeight: 700, marginBottom: 8, color: T.text }}>Acceso restringido</div>
      <div style={{ fontSize: 13, color: T.textMuted, textAlign: "center", maxWidth: 380 }}>
        Tu rol de <strong>{roleInfo?.label}</strong> no tiene permisos para esta sección. Contacta a un Director o al equipo de Soporte.
      </div>
    </div>
  );

  // ─── Render ────────────────────────────────────────────────────────────
  if (!currentUser) {
    return <LoginScreen onLogin={login} />;
  }

  if (loading) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', background: T.bg }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ width: 40, height: 40, border: `4px solid ${T.border}`, borderTopColor: T.primary, borderRadius: '50%', animation: 'spin 0.8s linear infinite', margin: '0 auto 16px' }} />
          <div style={{ color: T.textMuted, fontSize: 14 }}>Cargando datos...</div>
        </div>
      </div>
    );
  }

  return (
    <>
      <style>{GS}</style>
      <div style={{ display: "flex", minHeight: "100vh" }}>
        {/* SIDEBAR */}
        <aside style={{
          width: sideOpen ? 228 : 64,
          background: T.sidebar,
          flexShrink: 0,
          transition: "width 0.25s cubic-bezier(.4,0,.2,1)",
          overflow: "hidden",
          display: "flex",
          flexDirection: "column",
          position: "sticky",
          top: 0,
          height: "100vh"
        }}>
          <div style={{
            padding: sideOpen ? "22px 20px 18px" : "18px 14px",
            borderBottom: `1px solid ${T.sidebarLine}`,
            display: "flex",
            alignItems: "center",
            gap: 12
          }}>
            <div style={{
              width: 38, height: 38,
              background: `linear-gradient(135deg, ${T.primaryLt}, ${T.accent})`,
              borderRadius: 11,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 18, flexShrink: 0,
              boxShadow: `0 4px 12px ${T.accent}44`
            }}>🌿</div>
            {sideOpen && (
              <div>
                <div style={{ fontFamily: "'Fraunces',serif", color: "#fff", fontSize: 15, fontWeight: 700, lineHeight: 1.2 }}>Fundación Crescendo</div>
                <div style={{ color: "#80CBC4", fontSize: 10, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase" }}>Gestión Integral</div>
              </div>
            )}
          </div>

          {sideOpen && (
            <div style={{
              padding: "14px 16px",
              borderBottom: `1px solid ${T.sidebarLine}`,
              display: "flex",
              alignItems: "center",
              gap: 10
            }}>
              <div style={{
                width: 36, height: 36,
                borderRadius: "50%",
                background: roleInfo?.color + "33",
                border: `2px solid ${roleInfo?.color}55`,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 15, flexShrink: 0
              }}>{roleInfo?.icon}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 12.5, fontWeight: 700, color: "#fff", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser.name}</div>
                <div style={{ fontSize: 10, fontWeight: 600, color: roleInfo?.color, textTransform: "uppercase", letterSpacing: "0.05em" }}>{roleInfo?.label}</div>
              </div>
            </div>
          )}

          <nav style={{ flex: 1, padding: "10px 8px", overflowY: "auto" }}>
            {sideOpen && <div style={{ fontSize: 9, color: "#4A7A7A", fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", padding: "8px 12px 6px" }}>NAVEGACIÓN</div>}
            {nav.map(n => {
              const active = view === n.id;
              return (
                <button key={n.id} onClick={() => { setView(n.id); setDetailPat(null); }} style={{
                  display: "flex", alignItems: "center", gap: 12, width: "100%",
                  padding: sideOpen ? "10px 14px" : "12px", borderRadius: 10,
                  border: "none", background: active ? "rgba(255,255,255,0.12)" : "transparent",
                  color: active ? "#fff" : "rgba(255,255,255,0.5)", cursor: "pointer",
                  textAlign: "left", marginBottom: 2, transition: "all 0.15s",
                  fontFamily: "'Plus Jakarta Sans',sans-serif", fontSize: 13,
                  fontWeight: active ? 600 : 400,
                  justifyContent: sideOpen ? "flex-start" : "center",
                  borderLeft: active ? `3px solid ${T.accentLt}` : "3px solid transparent"
                }}
                  onMouseEnter={e => { if (!active) e.currentTarget.style.background = "rgba(255,255,255,0.08)"; }}
                  onMouseLeave={e => { if (!active) e.currentTarget.style.background = "transparent"; }}
                >
                  <span style={{ fontSize: 18, flexShrink: 0 }}>{n.icon}</span>
                  {sideOpen && n.label}
                </button>
              );
            })}
          </nav>

          {sideOpen && (
            <div style={{ padding: 16, borderTop: `1px solid ${T.sidebarLine}`, display: "flex", flexDirection: "column", gap: 8 }}>
              {can(currentUser, "import") && (
                <button onClick={() => setImportModal(true)} style={{
                  width: "100%", padding: "9px 12px",
                  background: "rgba(255,255,255,0.06)",
                  border: "1px solid rgba(255,255,255,0.12)",
                  borderRadius: 10,
                  color: "rgba(255,255,255,0.75)",
                  cursor: "pointer",
                  fontFamily: "'Plus Jakarta Sans',sans-serif",
                  fontSize: 12, fontWeight: 600,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 7
                }}
                  onMouseEnter={e => e.currentTarget.style.background = "rgba(255,255,255,0.1)"}
                  onMouseLeave={e => e.currentTarget.style.background = "rgba(255,255,255,0.06)"}
                >
                  📤 Importar datos
                </button>
              )}
              {can(currentUser, "export") && (
                <button onClick={exportExcel} style={{
                  width: "100%", padding: "9px 12px",
                  background: `${T.gold}18`,
                  border: `1px solid ${T.gold}44`,
                  borderRadius: 10,
                  color: T.gold,
                  cursor: "pointer",
                  fontFamily: "'Plus Jakarta Sans',sans-serif",
                  fontSize: 12, fontWeight: 600,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 7
                }}
                  onMouseEnter={e => e.currentTarget.style.background = `${T.gold}28`}
                  onMouseLeave={e => e.currentTarget.style.background = `${T.gold}18`}
                >
                  📥 Exportar Excel
                </button>
              )}
            </div>
          )}

          <div style={{ display: "flex", alignItems: "center", borderTop: `1px solid ${T.sidebarLine}` }}>
            <button onClick={() => setSideOpen(o => !o)} style={{
              flex: 1, padding: 14, background: "transparent", border: "none",
              color: "rgba(255,255,255,0.3)", cursor: "pointer", fontSize: 14,
              textAlign: "left"
            }}
              onMouseEnter={e => e.currentTarget.style.color = "rgba(255,255,255,0.6)"}
              onMouseLeave={e => e.currentTarget.style.color = "rgba(255,255,255,0.3)"}
            >
              {sideOpen ? "◂ Colapsar" : "▸"}
            </button>
            {sideOpen && (
              <button onClick={logout} style={{
                padding: "8px 12px", marginRight: 10,
                background: "rgba(229,152,155,0.15)",
                border: "1px solid rgba(229,152,155,0.3)",
                borderRadius: 8,
                color: "#D88A8A",
                cursor: "pointer",
                fontFamily: "'Plus Jakarta Sans',sans-serif",
                fontSize: 11, fontWeight: 600
              }}>
                Salir
              </button>
            )}
          </div>
        </aside>

        {/* MAIN CONTENT */}
        <main style={{ flex: 1, overflowY: "auto", minWidth: 0, background: T.bg }}>
          <header style={{
            background: "rgba(244,247,246,0.92)",
            backdropFilter: "blur(12px)",
            borderBottom: `1px solid ${T.border}`,
            padding: "14px 32px",
            display: "flex", alignItems: "center", justifyContent: "space-between",
            position: "sticky", top: 0, zIndex: 100
          }}>
            <div>
              <div style={{ fontFamily: "'Fraunces',serif", fontSize: 20, fontWeight: 700, color: T.text, display: "flex", alignItems: "center", gap: 8 }}>
                <span>{nav.find(n => n.id === view)?.icon}</span>
                {nav.find(n => n.id === view)?.label}
                {detailPat && view === "beneficiaries" && <span style={{ color: T.textMuted, fontWeight: 400, fontSize: 16 }}>/ {detailPat.nombre} {detailPat.apellido}</span>}
              </div>
              <div style={{ fontSize: 11.5, color: T.textMuted, marginTop: 1 }}>
                {new Date().toLocaleDateString("es-CL", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              {can(currentUser, "import") && <Btn variant="soft" onClick={() => setImportModal(true)} sm icon="📤">Importar</Btn>}
              {view === "beneficiaries" && !detailPat && can(currentUser, "patients_add") && <Btn onClick={() => setPatModal("new")} icon="＋">Nuevo Beneficiario</Btn>}
              {view === "donations" && can(currentUser, "donations_add") && <Btn variant="accent" onClick={() => setDonModal(true)} icon="＋">Registrar Donación</Btn>}
              {view === "projects" && can(currentUser, "projects_add") && <Btn onClick={() => setProjModal("new")} icon="＋">Nuevo Proyecto</Btn>}
              {view === "followups" && can(currentUser, "followups_add") && <Btn onClick={() => setFuModal("new")} icon="＋">Nuevo Seguimiento</Btn>}
              {can(currentUser, "export") && <Btn variant="gold" onClick={exportExcel} sm icon="📥">Excel</Btn>}
              {detailPat && can(currentUser, "patients_view") && (<Btn variant="accent" onClick={exportPDF} sm icon="📄">PDF</Btn>)}
              <button onClick={() => setShowProfile(p => !p)} style={{
                display: "flex", alignItems: "center", gap: 8,
                background: T.bgSoft, border: `1px solid ${T.border}`,
                borderRadius: 11, padding: "7px 12px", cursor: "pointer",
                transition: "all 0.15s"
              }}
                onMouseEnter={e => e.currentTarget.style.background = T.border}
                onMouseLeave={e => e.currentTarget.style.background = T.bgSoft}
              >
                <span style={{ fontSize: 15 }}>{roleInfo?.icon}</span>
                <div style={{ textAlign: "left" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.text }}>{currentUser.name.split(" ")[0]}</div>
                  <div style={{ fontSize: 10, color: roleInfo?.color, fontWeight: 600 }}>{roleInfo?.label}</div>
                </div>
              </button>
            </div>
          </header>

          {showProfile && (
            <>
              <div onClick={() => setShowProfile(false)} style={{ position: "fixed", inset: 0, zIndex: 199 }} />
              <div style={{
                position: "fixed", top: 66, right: 28,
                background: T.bgCard, border: `1px solid ${T.border}`,
                borderRadius: 14, padding: 20,
                boxShadow: "0 12px 40px rgba(0,0,0,0.12)",
                zIndex: 200, minWidth: 260,
                animation: "fadeUp 0.15s ease"
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
                  <div style={{
                    width: 48, height: 48, borderRadius: "50%",
                    background: roleInfo?.color + "22",
                    border: `2px solid ${roleInfo?.color}44`,
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: 22
                  }}>{roleInfo?.icon}</div>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{currentUser.name}</div>
                    <div style={{ fontSize: 12, color: T.textMuted }}>{currentUser.email}</div>
                    <Badge text={roleInfo?.label} color={roleInfo?.color} />
                  </div>
                </div>
                <div style={{ borderTop: `1px solid ${T.border}`, paddingTop: 14 }}>
                  <div style={{ fontSize: 11, color: T.textMuted, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Mis permisos activos</div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 5 }}>
                    {Object.entries(PERMS[currentUser.role] || {}).filter(([, v]) => v).map(([k]) => (
                      <span key={k} style={{ background: T.primary + "12", color: T.primary, borderRadius: 20, padding: "2px 8px", fontSize: 10, fontWeight: 600 }}>{k}</span>
                    ))}
                  </div>
                </div>
                <div style={{ marginTop: 16, display: "flex", gap: 8 }}>
                  <Btn sm variant="ghost" onClick={() => setShowProfile(false)} style={{ flex: 1 }}>Cerrar</Btn>
                  <Btn sm variant="danger" onClick={() => { setCurrentUser(null); setShowProfile(false); }} style={{ flex: 1 }}>Cerrar Sesión</Btn>
                </div>
              </div>
            </>
          )}

          <div style={{ padding: "28px 32px" }}>
            {/* --- DASHBOARD --- */}
            {view === "dashboard" && (
              <div style={{ animation: "fadeUp 0.3s ease" }}>
                <div style={{
                  background: `linear-gradient(135deg, ${roleInfo?.color}12, ${roleInfo?.color}05)`,
                  border: `1px solid ${roleInfo?.color}22`,
                  borderRadius: 13, padding: "14px 20px",
                  marginBottom: 22,
                  display: "flex", alignItems: "center", gap: 12
                }}>
                  <span style={{ fontSize: 22 }}>{roleInfo?.icon}</span>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>Hola, {currentUser.name.split(" ")[0]} 👋</div>
                    <div style={{ fontSize: 12, color: T.textSub }}>Tienes acceso como <strong style={{ color: roleInfo?.color }}>{roleInfo?.label}</strong> · {Object.values(PERMS[currentUser.role] || {}).filter(Boolean).length} permisos activos</div>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 16, marginBottom: 28 }}>
                  <StatCard icon="👥" label="Beneficiarios Activos" value={activeP} sub={`de ${beneficiaries.length} registrados`} color={T.primary} />
                  {can(currentUser, "donations_view") && <StatCard icon="💚" label="Total Donaciones" value={fmt(totalDon)} sub={`${donations.length} donaciones`} color={T.green} />}
                  {can(currentUser, "projects_view") && <StatCard icon="📁" label="Proyectos Activos" value={activePr} sub={`de ${projects.length} proyectos`} color={T.accent} />}
                  {can(currentUser, "followups_view") && <StatCard icon="📋" label="Seguimientos / Mes" value={fuMonth} sub={`${followups.length} histórico`} color={T.blue} />}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1.7fr 1fr", gap: 20, marginBottom: 20 }}>
                  {can(currentUser, "donations_view") && (
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Donaciones por Mes</div>
                      <ResponsiveContainer width="100%" height={210}>
                        <AreaChart data={donByMonth}>
                          <defs><linearGradient id="dg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={T.accent} stopOpacity={0.3} /><stop offset="95%" stopColor={T.accent} stopOpacity={0} /></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false} />
                          <XAxis dataKey="mes" tick={{ fontSize: 11, fill: T.textMuted }} axisLine={false} tickLine={false} />
                          <YAxis tick={{ fontSize: 11, fill: T.textMuted }} axisLine={false} tickLine={false} />
                          <Tooltip content={<TTip />} />
                          <Area type="monotone" dataKey="monto" name="Monto" stroke={T.accent} strokeWidth={2.5} fill="url(#dg)" dot={{ fill: T.accent, r: 4 }} />
                        </AreaChart>
                      </ResponsiveContainer>
                    </Card>
                  )}
                  <Card style={{ padding: 24 }}>
                    <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Beneficiarios por Tipo</div>
                    <ResponsiveContainer width="100%" height={210}>
                      <PieChart>
                        <Pie data={discData} cx="50%" cy="50%" innerRadius={56} outerRadius={82} paddingAngle={3} dataKey="value">
                          {discData.map((_, i) => <Cell key={i} fill={PIE[i % PIE.length]} />)}
                        </Pie>
                        <Legend iconType="circle" iconSize={7} wrapperStyle={{ fontSize: 11 }} />
                        <Tooltip content={<TTip />} />
                      </PieChart>
                    </ResponsiveContainer>
                  </Card>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                  {can(currentUser, "projects_view") && (
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Avance de Proyectos</div>
                      {projects.map(p => {
                        const pct = Math.min(100, Math.round(p.recaudado / p.objetivo * 100));
                        return (
                          <div key={p.id} style={{ marginBottom: 18 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 7 }}>
                              <div>
                                <div style={{ fontSize: 13, fontWeight: 600 }}>{p.nombre}</div>
                                <div style={{ fontSize: 11, color: T.textMuted }}>{fmt(p.recaudado)} de {fmt(p.objetivo)}</div>
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                <span style={{ fontSize: 13, fontWeight: 700, color: pct === 100 ? T.green : T.text }}>{pct}%</span>
                                {estadoBadge(p.estado)}
                              </div>
                            </div>
                            <div style={{ height: 7, background: T.bgSoft, borderRadius: 8, overflow: "hidden" }}>
                              <div style={{ height: "100%", width: pct + "%", background: pct === 100 ? `linear-gradient(90deg, ${T.green}, ${T.primaryLt})` : `linear-gradient(90deg, ${T.primary}, ${T.accentLt})`, borderRadius: 8, transition: "width 0.7s ease" }} />
                            </div>
                          </div>
                        );
                      })}
                    </Card>
                  )}
                  {can(currentUser, "followups_view") && (
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Últimos Seguimientos</div>
                      {[...followups].sort((a, b) => b.fecha.localeCompare(a.fecha)).slice(0, 5).map(f => {
                        const pac = beneficiaries.find(p => p.id === f.beneficiarioId);
                        const rc = f.resultado === "Positivo" ? T.green : f.resultado === "Neutral" ? T.gold : T.red;
                        return (
                          <div key={f.id} style={{ display: "flex", alignItems: "center", gap: 12, paddingBottom: 13, marginBottom: 13, borderBottom: `1px solid ${T.border}` }}>
                            <div style={{ width: 40, height: 40, borderRadius: 11, background: T.bgSoft, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0, border: `1px solid ${T.border}` }}>
                              {{ Médico: "🏥", Terapia: "🧠", Educativo: "📚", Social: "🤝", Psicológico: "💬" }[f.tipo] || "📋"}
                            </div>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: 13, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{pac ? `${pac.nombre} ${pac.apellido}` : "–"}</div>
                              <div style={{ fontSize: 11, color: T.textMuted }}>{fmtD(f.fecha)} · {f.profesional}</div>
                            </div>
                            <div style={{ width: 8, height: 8, borderRadius: "50%", background: rc, flexShrink: 0, boxShadow: `0 0 6px ${rc}` }} />
                          </div>
                        );
                      })}
                    </Card>
                  )}
                </div>
              </div>
            )}

            {/* --- PATIENTS LIST --- */}
            {view === "beneficiaries" && !detailPat && (
              can(currentUser, "patients_view") ? (
                <div style={{ animation: "fadeUp 0.3s ease" }}>
                  <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap", alignItems: "center" }}>
                    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍  Buscar por nombre, diagnóstico, tutor..." style={{ flex: 1, minWidth: 200, maxWidth: 380 }} />
                    <select value={filterDisc} onChange={e=>setFilterDisc(e.target.value)} style={{ width:"auto", minWidth:160 }}>{discTypes.map(o=><option key={o}>{o}</option>)}</select>
                    <select value={filterEst} onChange={e=>setFilterEst(e.target.value)} style={{ width:"auto", minWidth:130 }}>{["Todos","Activo","Inactivo"].map(o=><option key={o}>{o}</option>)}</select>
                    <span style={{ fontSize:12, color:T.textMuted, whiteSpace:"nowrap" }}>{filtPat.length} resultado{filtPat.length !== 1 ? "s" : ""}</span>
                  </div>
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))", gap:16 }}>
                    {filtPat.map((p, i) => {
                      const di = { Física:"♿", Sensorial:"👁", Cognitiva:"🧩" }[p.condicion] || "🧩";
                      return (
                        <div key={p.id} onClick={() => setDetailPat(p)} style={{
                          background: T.bgCard, borderRadius: 16, padding: 20,
                          boxShadow: "0 2px 14px rgba(0,0,0,0.055)", cursor: "pointer",
                          border: "1.5px solid transparent", transition: "all 0.18s",
                          animation: `fadeUp 0.3s ease ${i * 0.04}s both`
                        }}
                          onMouseEnter={e => { e.currentTarget.style.border = `1.5px solid ${T.primary}40`; e.currentTarget.style.boxShadow = "0 6px 24px rgba(0,0,0,0.1)"; e.currentTarget.style.transform = "translateY(-2px)"; }}
                          onMouseLeave={e => { e.currentTarget.style.border = "1.5px solid transparent"; e.currentTarget.style.boxShadow = "0 2px 14px rgba(0,0,0,0.055)"; e.currentTarget.style.transform = "none"; }}
                        >
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                              <div style={{ width: 46, height: 46, borderRadius: 13, background: `${T.primary}15`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20, border: `1px solid ${T.primary}20` }}>{di}</div>
                              <div>
                                <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15 }}>{p.nombre} {p.apellido}</div>
                                <div style={{ fontSize: 11, color: T.textMuted }}>{edad(p.fecha_nac)} años</div>
                              </div>
                            </div>
                            {estadoBadge(p.estado)}
                          </div>
                          {[["Diagnóstico", p.diagnostico], ["Condicion", `${p.condicion} · Nivel ${p.nivel}`], ["Tutor", `${p.tutor} (${p.relacionTutor})`]].map(([k, v]) => (
                            <div key={k} style={{ display: "flex", gap: 6, fontSize: 12, marginBottom: 4 }}>
                              <span style={{ color: T.primary, fontWeight: 600, flexShrink: 0 }}>{k}:</span>
                              <span style={{ color: T.textSub }}>{v}</span>
                            </div>
                          ))}
                          <div style={{ marginTop: 13, paddingTop: 12, borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: 11, color: T.textMuted }}>📅 Ingreso: {fmtD(p.fecha_ingreso)}</span>
                            <span style={{ fontSize: 11, color: T.primary, fontWeight: 600 }}>Ver ficha →</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : <AccessDenied />
            )}

            {/* --- PATIENT DETAIL --- */}
            {view === "beneficiaries" && detailPat && (() => {
              const p = beneficiaries.find(x => x.id === detailPat.id) || detailPat;
              const patFu = followups.filter(f => f.beneficiarioId === p.id);
              const patDon = donations.filter(d => d.beneficiarioId === p.id);
              const patProj = projects.filter(pr => pr.beneficiarios && pr.beneficiarios.includes(p.id));
              return (
                <div style={{ animation: "fadeUp 0.25s ease" }}>
                  <button onClick={() => setDetailPat(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.primary, fontWeight: 600, fontSize: 13, marginBottom: 22, display: "flex", alignItems: "center", gap: 6 }}>← Volver a Beneficiarios</button>
                  <div id="perfil-beneficiario" style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 20 }}>
                    <div>
                      <Card style={{ padding: 24, marginBottom: 16 }}>
                        <div style={{ textAlign: "center", marginBottom: 20 }}>
                          <div style={{ width: 72, height: 72, borderRadius: 20, background: `${T.primary}15`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 32, margin: "0 auto 12px", border: `1px solid ${T.primary}20` }}>{({ Física: "♿", Sensorial: "👁" })[p.condicion] || "🧩"}</div>
                          <div style={{ fontFamily: "'Fraunces',serif", fontSize: 20, fontWeight: 700 }}>{p.nombre} {p.apellido}</div>
                          <div style={{ color: T.textMuted, fontSize: 12, margin: "4px 0 10px" }}>{edad(p.fecha_nac)} años · {fmtD(p.fecha_nac)}</div>
                          {estadoBadge(p.estado)}
                        </div>
                        {[["Diagnóstico", p.diagnostico], ["Condicion", p.condicion], ["Nivel de Apoyo", p.nivel], ["Teléfono", p.telefono], ["Email", p.email || "–"], ["Dirección", p.direccion], ["Ingreso", fmtD(p.fecha_ingreso)]].map(([k, v]) => (
                          <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${T.border}`, fontSize: 12.5 }}>
                            <span style={{ color: T.textMuted, fontWeight: 500 }}>{k}</span>
                            <span style={{ color: T.text, fontWeight: 600, textAlign: "right", maxWidth: "58%" }}>{v}</span>
                          </div>
                        ))}
                        <div style={{ marginTop: 16, background: T.bgSoft, borderRadius: 11, padding: 14 }}>
                          <div style={{ fontSize: 10, color: T.textMuted, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>Tutor / Representante</div>
                          <div style={{ fontWeight: 700, fontSize: 13 }}>{p.tutor}</div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>{p.relacionTutor} · {p.telefonoTutor}</div>
                        </div>
                        {p.notas && <div style={{ marginTop: 12, padding: 12, background: T.accent + "10", borderRadius: 10, fontSize: 12, borderLeft: `3px solid ${T.accent}` }}>{p.notas}</div>}
                        {(can(currentUser, "patients_edit") || can(currentUser, "patients_del")) && (
                          <div style={{ 
                            marginTop: 16, 
                            display: "flex", 
                            gap: 8, 
                            flexWrap: "wrap",
                            visibility: isExporting ? "hidden" : "visible" 
                          }}>
                            {can(currentUser, "patients_edit") && <Btn sm onClick={() => setPatModal(p)} icon="✏️">Editar</Btn>}
                            {can(currentUser, "patients_del") && <Btn sm variant="danger" onClick={() => delPat(p.id)} icon="🗑">Eliminar</Btn>}
                          </div>
                        )}
                      </Card>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                      {can(currentUser, "followups_view") && (
                        <Card style={{ padding: 20 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                            <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15 }}>Seguimientos <span style={{ color: T.textMuted, fontWeight: 400, fontSize: 13 }}>({patFu.length})</span></div>
                            {can(currentUser, "followups_add") && <Btn sm onClick={() => setFuModal({ beneficiarioId: p.id })} icon="＋">Agregar</Btn>}
                          </div>
                          {patFu.length === 0 ? <div style={{ color: T.textMuted, fontSize: 13, textAlign: "center", padding: 20 }}>Sin seguimientos registrados</div> :
                            patFu.sort((a, b) => b.fecha.localeCompare(a.fecha)).map(f => (
                              <div key={f.id} style={{ display: "flex", gap: 12, paddingBottom: 13, marginBottom: 13, borderBottom: `1px solid ${T.border}` }}>
                                <div style={{ flex: 1 }}>
                                  <div style={{ display: "flex", gap: 7, alignItems: "center", marginBottom: 5, flexWrap: "wrap" }}>
                                    {tipoBadge(f.tipo)}
                                    <span style={{ fontSize: 11, color: T.textMuted }}>{fmtD(f.fecha)}</span>
                                    <Badge text={f.resultado} color={f.resultado === "Positivo" ? T.green : f.resultado === "Neutral" ? T.gold : T.red} />
                                  </div>
                                  <div style={{ fontSize: 12.5, lineHeight: 1.5 }}>{f.descripcion}</div>
                                  <div style={{ fontSize: 11, color: T.textMuted, marginTop: 4 }}>{f.profesional} · Próx: {fmtD(f.proxima)}</div>
                                </div>
                                {can(currentUser, "followups_edit") && (
                                  <button onClick={() => delFu(f.id)} style={{ background: "none", border: "none", cursor: "pointer", color: T.red, fontSize: 14, padding: "0 4px", opacity: 0.6 }}>✕</button>
                                )}
                              </div>
                            ))
                          }
                        </Card>
                      )}
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                        {can(currentUser, "projects_view") && (
                          <Card style={{ padding: 20 }}>
                            <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 14, marginBottom: 14 }}>Proyectos ({patProj.length})</div>
                            {patProj.length === 0 ? <div style={{ color: T.textMuted, fontSize: 12, textAlign: "center", padding: 12 }}>No asignado a proyectos</div> :
                              patProj.map(pr => (
                                <div key={pr.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                                  <div><div style={{ fontWeight: 600 }}>{pr.nombre}</div><div style={{ fontSize: 11, color: T.textMuted }}>{fmt(pr.recaudado)} / {fmt(pr.objetivo)}</div></div>
                                  {estadoBadge(pr.estado)}
                                </div>
                              ))
                            }
                          </Card>
                        )}
                        {can(currentUser, "donations_view") && (
                          <Card style={{ padding: 20 }}>
                            <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 14, marginBottom: 14 }}>Donaciones Familiares ({patDon.length})</div>
                            {patDon.length === 0 ? <div style={{ color: T.textMuted, fontSize: 12, textAlign: "center", padding: 12 }}>Sin donaciones</div> :
                              patDon.map(d => (
                                <div key={d.id} style={{ display: "flex", justifyContent: "space-between", padding: "9px 0", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                                  <div><div style={{ fontWeight: 600 }}>{d.donante}</div><div style={{ fontSize: 11, color: T.textMuted }}>{fmtD(d.fecha)}</div></div>
                                  <div style={{ fontWeight: 700, color: T.green }}>{fmt(d.monto)}</div>
                                </div>
                              ))
                            }
                            {patDon.length > 0 && <div style={{ textAlign: "right", fontSize: 13, fontWeight: 700, color: T.green, marginTop: 10 }}>Total: {fmt(patDon.reduce((s, d) => s + d.monto, 0))}</div>}
                          </Card>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* --- DONATIONS --- */}
            {view === "donations" && (
              can(currentUser, "donations_view") ? (
                <div style={{ animation: "fadeUp 0.3s ease" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16, marginBottom: 24 }}>
                    <StatCard icon="💰" label="Total Recaudado" value={fmt(totalDon)} color={T.green} />
                    <StatCard icon="👨‍👩‍👦" label="Donaciones Familiares" value={fmt(donations.filter(d => d.relacion === "Familiar").reduce((s, d) => s + d.monto, 0))} sub={`${donations.filter(d => d.relacion === "Familiar").length} donaciones`} color={T.primary} />
                    <StatCard icon="🏢" label="Donaciones Externas/Corp." value={fmt(donations.filter(d => d.relacion !== "Familiar").reduce((s, d) => s + d.monto, 0))} sub={`${donations.filter(d => d.relacion !== "Familiar").length} donaciones`} color={T.accent} />
                  </div>
                  <Card style={{ overflow: "hidden" }}>
                    <div style={{ padding: "16px 22px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15 }}>Registro de Donaciones</div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.textMuted }}>
                        <span style={{ width: 7, height: 7, borderRadius: "50%", background: T.textMuted, display: "inline-block" }} />
                        Solo lectura · Para modificar, contacta al administrador
                      </div>
                    </div>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: T.bgSoft }}>
                            {["Donante", "Relación", "Beneficiario", "Monto", "Fecha", "Tipo", "Proyecto", "Notas"].map(h => (
                              <th key={h} style={{ padding: "11px 16px", textAlign: "left", fontSize: 10.5, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {[...donations].sort((a, b) => b.fecha.localeCompare(a.fecha)).map(d => {
                            const pac = beneficiaries.find(p => p.id === d.beneficiarioId);
                            const proj = projects.find(p => p.id === d.proyectoId);
                            return (
                              <tr key={d.id} style={{ borderBottom: `1px solid ${T.border}` }}
                                onMouseEnter={e => e.currentTarget.style.background = T.bgSoft}
                                onMouseLeave={e => e.currentTarget.style.background = ""}>
                                <td style={{ padding: "12px 16px", fontWeight: 600, fontSize: 13 }}>{d.donante}</td>
                                <td style={{ padding: "12px 16px" }}><Badge text={d.relacion} color={d.relacion === "Familiar" ? T.primary : d.relacion === "Corporativo" ? T.blue : T.accent} /></td>
                                <td style={{ padding: "12px 16px", fontSize: 12, color: T.textMuted }}>{pac ? `${pac.nombre} ${pac.apellido}` : "–"}</td>
                                <td style={{ padding: "12px 16px", fontWeight: 700, color: T.green, fontSize: 14 }}>{fmt(d.monto)}</td>
                                <td style={{ padding: "12px 16px", fontSize: 12, color: T.textMuted, whiteSpace: "nowrap" }}>{fmtD(d.fecha)}</td>
                                <td style={{ padding: "12px 16px", fontSize: 12 }}>{d.tipo}</td>
                                <td style={{ padding: "12px 16px", fontSize: 12, color: T.textMuted, maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{proj ? proj.nombre : "–"}</td>
                                <td style={{ padding: "12px 16px", fontSize: 11, color: T.textMuted, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{d.notas || "–"}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </Card>
                </div>
              ) : <AccessDenied />
            )}

            {/* --- PROJECTS --- */}
            {view === "projects" && (
              can(currentUser, "projects_view") ? (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(330px,1fr))", gap: 20, animation: "fadeUp 0.3s ease" }}>
                  {projects.map(p => {
                    const pct = Math.min(100, Math.round(p.recaudado / p.objetivo * 100));
                    const bens = beneficiaries.filter(x => p.beneficiarios && p.beneficiarios.includes(x.id));
                    const pDon = donations.filter(d => d.proyectoId === p.id);
                    return (
                      <Card key={p.id} style={{ overflow: "hidden" }}>
                        <div style={{ height: 4, background: `linear-gradient(90deg, ${T.primary}, ${T.accent})` }} />
                        <div style={{ padding: "20px 22px 16px" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                            <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 17, flex: 1, lineHeight: 1.3 }}>{p.nombre}</div>
                            {estadoBadge(p.estado)}
                          </div>
                          <div style={{ fontSize: 12.5, color: T.textSub, lineHeight: 1.6, marginBottom: 16 }}>{p.descripcion}</div>
                          <div style={{ marginBottom: 14 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 7 }}>
                              <span style={{ color: T.textMuted }}>Recaudado</span>
                              <span style={{ fontWeight: 700, color: pct === 100 ? T.green : T.text }}>{fmt(p.recaudado)} <span style={{ color: T.textMuted, fontWeight: 400 }}>/ {fmt(p.objetivo)}</span></span>
                            </div>
                            <div style={{ height: 9, background: T.bgSoft, borderRadius: 9, overflow: "hidden" }}>
                              <div style={{ height: "100%", width: pct + "%", background: pct === 100 ? `linear-gradient(90deg, ${T.green}, ${T.primaryLt})` : `linear-gradient(90deg, ${T.primary}, ${T.accentLt})`, borderRadius: 9, transition: "width 0.7s ease" }} />
                            </div>
                            <div style={{ textAlign: "right", fontSize: 11, color: pct === 100 ? T.green : T.textMuted, marginTop: 4, fontWeight: 600 }}>{pct}% completado</div>
                          </div>
                          <div style={{ display: "flex", gap: 14, fontSize: 11.5, color: T.textMuted }}>
                            <span>📅 {fmtD(p.inicio)} – {fmtD(p.fin)}</span>
                            <span>💚 {pDon.length} donaciones · {fmt(pDon.reduce((s, d) => s + d.monto, 0))}</span>
                          </div>
                        </div>
                        {bens.length > 0 && <div style={{ padding: "12px 22px", background: T.bgSoft, borderTop: `1px solid ${T.border}` }}>
                          <div style={{ fontSize: 10, color: T.textMuted, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>Beneficiarios</div>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{bens.map(b => <Badge key={b.id} text={`${b.nombre} ${b.apellido}`} color={T.primary} />)}</div>
                        </div>}
                        {can(currentUser, "projects_add") && (
                          <div style={{ padding: "12px 22px", borderTop: `1px solid ${T.border}`, display: "flex", gap: 8 }}>
                            <Btn sm onClick={() => setProjModal(p)} icon="✏️">Editar</Btn>
                            <Btn sm variant="danger" onClick={() => delProj(p.id)}>🗑</Btn>
                          </div>
                        )}
                      </Card>
                    );
                  })}
                </div>
              ) : <AccessDenied />
            )}

            {/* --- FOLLOWUPS (MEJORADO) --- */}
            {view === "followups" && (
              can(currentUser, "followups_view") ? (
                <div style={{ animation: "fadeUp 0.3s ease" }}>
                  <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap", alignItems: "center" }}>
                    <select
                      value={filterFuPatient}
                      onChange={e => setFilterFuPatient(e.target.value)}
                      style={{ width: "auto", minWidth: 180 }}
                    >
                      <option value="Todos">Todos los Beneficiarios</option>
                      {beneficiaries.map(p => (
                        <option key={p.id} value={p.id}>{p.nombre} {p.apellido}</option>
                      ))}
                    </select>
                    <span style={{ fontSize: 12, color: T.textMuted }}>
                      {followups.filter(f => filterFuPatient === "Todos" || f.beneficiarioId === parseInt(filterFuPatient)).length} seguimientos
                    </span>
                    {can(currentUser, "followups_add") && (
                      <Btn onClick={() => setFuModal("new")} icon="＋">Nuevo Seguimiento</Btn>
                    )}
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(290px,1fr))", gap: 16 }}>
                    {[...followups]
                      .filter(f => filterFuPatient === "Todos" || f.beneficiarioId === parseInt(filterFuPatient))
                      .sort((a, b) => b.fecha.localeCompare(a.fecha))
                      .map((f, i) => {
                        const pac = beneficiaries.find(p => p.id === f.beneficiarioId);
                        const tc = { Médico: T.blue, Terapia: T.green, Educativo: T.accent, Social: T.purple, Psicológico: T.teal }[f.tipo] || T.primary;
                        const rc = f.resultado === "Positivo" ? T.green : f.resultado === "Neutral" ? T.gold : T.red;
                        return (
                          <Card key={f.id} style={{ borderTop: `3px solid ${tc}`, overflow: "hidden", animation: `fadeUp 0.3s ease ${i * 0.04}s both` }}>
                            <div style={{ padding: "18px 20px" }}>
                              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
                                <div>
                                  <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15 }}>{pac ? `${pac.nombre} ${pac.apellido}` : "–"}</div>
                                  <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>{fmtD(f.fecha)}</div>
                                </div>
                                <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 5 }}>
                                  {tipoBadge(f.tipo)}
                                  <Badge text={f.resultado} color={rc} />
                                </div>
                              </div>
                              <div style={{ fontSize: 12.5, color: T.text, lineHeight: 1.6, marginBottom: 12 }}>{f.descripcion}</div>
                              <div style={{ fontSize: 11.5, color: T.textMuted, paddingTop: 11, borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                                <span>👤 {f.profesional}</span>
                                <span>Próx: {fmtD(f.proxima)}</span>
                              </div>
                            </div>
                            {can(currentUser, "followups_edit") && (
                              <div style={{ padding: "10px 20px 14px", display: "flex", gap: 8 }}>
                                <Btn sm onClick={() => setFuModal(f)} icon="✏️">Editar</Btn>
                                <Btn sm variant="danger" onClick={() => delFu(f.id)}>🗑</Btn>
                              </div>
                            )}
                          </Card>
                        );
                      })}
                  </div>
                </div>
              ) : <AccessDenied />
            )}

            {/* --- STATS --- */}
            {view === "stats" && (
              can(currentUser, "stats") ? (
                <div style={{ animation: "fadeUp 0.3s ease" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Seguimientos por Tipo</div>
                      <ResponsiveContainer width="100%" height={220}>
                        <BarChart data={fuByType} layout="vertical">
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border} horizontal={false} />
                          <XAxis type="number" tick={{ fontSize: 11, fill: T.textMuted }} axisLine={false} tickLine={false} />
                          <YAxis type="category" dataKey="name" tick={{ fontSize: 12, fill: T.text }} width={82} axisLine={false} tickLine={false} />
                          <Tooltip content={<TTip />} />
                          <Bar dataKey="value" name="Cantidad" fill={T.primary} radius={[0, 7, 7, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </Card>
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Resultados de Seguimientos</div>
                      <ResponsiveContainer width="100%" height={220}>
                        <PieChart>
                          <Pie data={resultData} cx="50%" cy="50%" outerRadius={82} dataKey="value" label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false}>
                            {resultData.map((_, i) => <Cell key={i} fill={[T.green, T.gold, T.red][i]} />)}
                          </Pie>
                          <Tooltip content={<TTip />} />
                        </PieChart>
                      </ResponsiveContainer>
                    </Card>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 20, marginBottom: 20 }}>
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Meta vs. Recaudado por Proyecto</div>
                      <ResponsiveContainer width="100%" height={220}>
                        <BarChart data={projects.map(p => ({ name: p.nombre.split(" ").slice(0, 3).join(" "), meta: p.objetivo, recaudado: p.recaudado }))}>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false} />
                          <XAxis dataKey="name" tick={{ fontSize: 10, fill: T.textMuted }} axisLine={false} tickLine={false} />
                          <YAxis tick={{ fontSize: 11, fill: T.textMuted }} axisLine={false} tickLine={false} />
                          <Tooltip content={<TTip />} />
                          <Legend wrapperStyle={{ fontSize: 12 }} />
                          <Bar dataKey="meta" name="Meta" fill={T.border} radius={[4, 4, 0, 0]} />
                          <Bar dataKey="recaudado" name="Recaudado" fill={T.primary} radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </Card>
                    <Card style={{ padding: 24 }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15, marginBottom: 18 }}>Donaciones por Tipo de Pago</div>
                      <ResponsiveContainer width="100%" height={220}>
                        <PieChart>
                          <Pie data={(() => { const c = {}; donations.forEach(d => { c[d.tipo] = (c[d.tipo] || 0) + d.monto; }); return Object.entries(c).map(([n, v]) => ({ name: n, value: v })); })()}
                            cx="50%" cy="50%" outerRadius={82} innerRadius={42} dataKey="value">
                            {PIE.map((color, i) => <Cell key={i} fill={color} />)}
                          </Pie>
                          <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize: 11 }} />
                          <Tooltip formatter={v => fmt(v)} content={<TTip />} />
                        </PieChart>
                      </ResponsiveContainer>
                    </Card>
                  </div>
                  <Card style={{ overflow: "hidden" }}>
                    <div style={{ padding: "18px 24px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontFamily: "'Fraunces',serif", fontWeight: 700, fontSize: 15 }}>Resumen por Beneficiario</div>
                      {can(currentUser, "export") && <Btn sm variant="gold" onClick={exportExcel} icon="📥">Exportar Excel</Btn>}
                    </div>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ background: T.bgSoft }}>
                            {["Beneficiario", "Condicion", "Nivel", "Seguimientos", "Donaciones ($)", "Proyectos", "Estado"].map(h => (
                              <th key={h} style={{ padding: "11px 16px", textAlign: "left", fontSize: 10.5, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {beneficiaries.map(p => (
                            <tr key={p.id} style={{ borderBottom: `1px solid ${T.border}` }}
                              onMouseEnter={e => e.currentTarget.style.background = T.bgSoft}
                              onMouseLeave={e => e.currentTarget.style.background = ""}>
                              <td style={{ padding: "11px 16px", fontWeight: 600, fontSize: 13 }}>{p.nombre} {p.apellido}</td>
                              <td style={{ padding: "11px 16px", fontSize: 12, color: T.textSub }}>{p.condicion}</td>
                              <td style={{ padding: "11px 16px", fontSize: 12, color: T.textSub }}>{p.nivel}</td>
                              <td style={{ padding: "11px 16px", textAlign: "center", fontWeight: 700, color: T.primary }}>{followups.filter(f => f.beneficiarioId === p.id).length}</td>
                              <td style={{ padding: "11px 16px", fontWeight: 700, color: T.green }}>{fmt(donations.filter(d => d.beneficiarioId === p.id).reduce((s, d) => s + d.monto, 0))}</td>
                              <td style={{ padding: "11px 16px", textAlign: "center" }}>{projects.filter(pr => pr.beneficiarios && pr.beneficiarios.includes(p.id)).length}</td>
                              <td style={{ padding: "11px 16px" }}>{estadoBadge(p.estado)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </Card>
                </div>
              ) : <AccessDenied />
            )}

            {/* --- USERS --- */}
            {view === "users" && (
              can(currentUser, "users") ? (
                <UsersPanel
                  users={users}
                  onSaveUser={saveUser}
                  onToggleActive={toggleUserActive}
                  onDeleteUser={deleteUser}
                  currentUser={currentUser}
                />
              ) : <AccessDenied />
            )}
          </div>
        </main>
      </div>

      {/* MODALES */}
      {patModal && <Modal title={patModal === "new" ? "Registrar Nuevo Beneficiario" : "Editar Beneficiario"} subtitle="Fundación Crescendo" onClose={() => setPatModal(null)} wide>
        <PatientForm init={patModal !== "new" ? patModal : null} onSave={savePat} onClose={() => setPatModal(null)} />
      </Modal>}
      {donModal && <Modal title="Registrar Donación" subtitle="Los registros de donaciones son permanentes" onClose={() => setDonModal(false)}>
        <DonationForm onSave={saveDon} onClose={() => setDonModal(false)} beneficiaries={beneficiaries} projects={projects} />
      </Modal>}
      {projModal && <Modal title={projModal === "new" ? "Nuevo Proyecto" : "Editar Proyecto"} onClose={() => setProjModal(null)} wide>
        <ProjectForm init={projModal !== "new" ? projModal : null} onSave={saveProj} onClose={() => setProjModal(null)} beneficiaries={beneficiaries} />
      </Modal>}
      {fuModal && (
        <Modal title={fuModal?.id ? "Editar Seguimiento" : "Nuevo Seguimiento"} onClose={() => setFuModal(null)}>
          <FollowupForm
            init={fuModal?.id ? fuModal : (fuModal === "new" ? null : null)}
            onSave={saveFu}
            onClose={() => setFuModal(null)}
            beneficiaries={beneficiaries}
          />
        </Modal>
      )}
      {importModal && <ImportModal onClose={() => setImportModal(false)} onImport={handleImport} />}
      {/* ─── CONTENEDOR OCULTO PARA PDF ───────────────────────────── */}
      <div style={{ position: 'absolute', left: '-9999px', top: '-9999px', zIndex: -999, width: '800px' }}>
        {renderPDFContent()}
      </div>
    </>
  );
}