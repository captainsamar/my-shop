import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// Enter URL+Key in setup screen, or set env vars in Vercel dashboard
const SB_URL = import.meta.env?.VITE_SUPABASE_URL ||"";
const SB_KEY = import.meta.env?.VITE_SUPABASE_ANON_KEY || "";

const sb = async (method, table, body=null, qs="") => {
  const url = `${SB_URL}/rest/v1/${table}${qs?"?"+qs:""}`;
  const res = await fetch(url, {
    method,
    headers: {
      apikey: SB_KEY, Authorization: `Bearer ${SB_KEY}`,
      "Content-Type": "application/json",
      ...(["POST","PATCH"].includes(method)?{Prefer:"return=representation"}:{}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) throw new Error(await res.text());
  const t = await res.text(); return t ? JSON.parse(t) : null;
};
const sbUpsert = (table, rows) => fetch(`${SB_URL}/rest/v1/${table}`, {
  method:"POST",
  headers:{ apikey:SB_KEY, Authorization:`Bearer ${SB_KEY}`, "Content-Type":"application/json", Prefer:"resolution=merge-duplicates,return=representation" },
  body: JSON.stringify(rows),
}).then(r=>r.json());

const LS = {
  get: k => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set: (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} },
};

const fmt  = n => (n||0).toLocaleString("en-US")+" Ks";
const uid  = () => (crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)+Date.now().toString(36));
const onum = () => "ORD-"+Date.now().toString(36).toUpperCase();
const fdate = s => s ? new Date(s).toLocaleString("en-GB",{day:"2-digit",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"}) : "";
const fileToB64 = f => new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(f);});
const getImgs = p => { let i=p?.images; if(typeof i==="string"){try{i=JSON.parse(i);}catch{i=[];}} return Array.isArray(i)?i:[]; };
const isOOS  = p => !p.preorder && (p.stock||0)<=0;
const calcUnit = (p, qty=1) => {
  if(!p) return 0;
  let bd=p.bulk_discounts; if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch{bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>b.min_qty-a.min_qty);
  const bulk=bulks.find(b=>qty>=b.min_qty);
  if(bulk) return Math.round((p.price||0)*(1-bulk.discount_percent/100));
  if((p.discount_value||0)>0) return p.discount_type==="fixed"?Math.max(0,(p.price||0)-p.discount_value):Math.round((p.price||0)*(1-p.discount_value/100));
  return p.price||0;
};

// GLASS BEAUTY THEME
const G = {
  bg: "linear-gradient(135deg, #FDE8D8 0%, #F5C6B0 20%, #EAA8C0 45%, #C9A8E0 70%, #A8C0E8 100%)",
  bgFixed: true,
  primary: "#8B2252",
  primaryHover: "#6D1A40",
  accent: "#D4896A",
  accentLight: "#F5C6B0",
  dark: "#2D1B2E",
  surface: "rgba(255,255,255,0.22)",
  surface2: "rgba(255,255,255,0.12)",
  surfaceStrong: "rgba(255,255,255,0.35)",
  border: "rgba(255,255,255,0.35)",
  borderStrong: "rgba(255,255,255,0.55)",
  text: "#2D1B2E",
  textLight: "#FFFFFF",
  textSub: "#5D3060",
  muted: "rgba(93,48,96,0.6)",
  success: "#2E7D52",
  successBg: "rgba(46,125,82,0.15)",
  danger: "#B71C5A",
  dangerBg: "rgba(183,28,90,0.12)",
  warning: "#8B5E00",
  warningBg: "rgba(139,94,0,0.12)",
  blur: "blur(24px)",
  shadow: "0 8px 32px rgba(139,34,82,0.15)",
  shadowMd: "0 16px 48px rgba(139,34,82,0.2)",
  shadowLg: "0 24px 64px rgba(139,34,82,0.25)",
  navBg: "rgba(255,255,255,0.25)",
  tag: "rgba(139,34,82,0.12)",
  tagText: "#8B2252",
  blob1: "rgba(255,182,193,0.4)",
  blob2: "rgba(216,191,216,0.4)",
  blob3: "rgba(255,218,185,0.4)",
};

// DARK GLASS THEME
const DG = {
  bg: "linear-gradient(135deg, #1A0A2E 0%, #2D1B45 25%, #1A2D35 50%, #2D1A30 75%, #0A1A2E 100%)",
  bgFixed: true,
  primary: "#E88DB5",
  primaryHover: "#D4769E",
  accent: "#E8B48D",
  accentLight: "rgba(232,180,141,0.2)",
  dark: "#F0E8F5",
  surface: "rgba(255,255,255,0.08)",
  surface2: "rgba(255,255,255,0.04)",
  surfaceStrong: "rgba(255,255,255,0.15)",
  border: "rgba(255,255,255,0.12)",
  borderStrong: "rgba(255,255,255,0.25)",
  text: "#F0E8F5",
  textLight: "#FFFFFF",
  textSub: "#D4C8E0",
  muted: "rgba(212,200,224,0.6)",
  success: "#5CB87A",
  successBg: "rgba(92,184,122,0.15)",
  danger: "#E87A9E",
  dangerBg: "rgba(232,122,158,0.15)",
  warning: "#E8C47A",
  warningBg: "rgba(232,196,122,0.15)",
  blur: "blur(24px)",
  shadow: "0 8px 32px rgba(0,0,0,0.4)",
  shadowMd: "0 16px 48px rgba(0,0,0,0.5)",
  shadowLg: "0 24px 64px rgba(0,0,0,0.6)",
  navBg: "rgba(26,10,46,0.7)",
  tag: "rgba(232,141,181,0.2)",
  tagText: "#E88DB5",
  blob1: "rgba(139,34,82,0.3)",
  blob2: "rgba(100,60,140,0.3)",
  blob3: "rgba(60,100,140,0.3)",
};

const THEMES = { glass: G, dark: DG };

const SC = {
  pending:   {bg:"rgba(255,200,100,0.2)",color:"#8B5E00",border:"rgba(255,200,100,0.4)"},
  confirmed: {bg:"rgba(100,150,255,0.2)",color:"#1A3A8B",border:"rgba(100,150,255,0.4)"},
  deposited: {bg:"rgba(100,200,150,0.2)",color:"#1A6B3A",border:"rgba(100,200,150,0.4)"},
  completed: {bg:"rgba(80,200,120,0.2)",color:"#0A5A2A",border:"rgba(80,200,120,0.4)"},
  cancelled: {bg:"rgba(220,80,120,0.2)",color:"#8B1A3A",border:"rgba(220,80,120,0.4)"},
};

const TR = {
  en:{ home:"Home", cats:"Categories", cart:"Cart", orders:"Orders", admin:"Admin", search:"Search beauty products...", all:"All", addCart:"Add to Cart", buyNow:"Buy Now", preorder:"Pre-order", outOfStock:"Out of Stock", tempOOS:"Temporarily Out of Stock", inStock:"In Stock", qty:"Qty", total:"Total", checkout:"Checkout", orderSummary:"Order Summary", contactOrder:"Contact to Order", trackOrder:"Track Order", enterOrderNo:"Enter order number", track:"Track", yourName:"Your Name *", yourPhone:"Phone *", yourAddress:"Delivery Address *", note:"Special requests / note", adminPanel:"Admin Panel", password:"Password", login:"Login", products:"Products", reports:"Reports", settings:"Settings", addProduct:"Add Product", editProduct:"Edit Product", save:"Save", cancel:"Cancel", delete:"Delete", hide:"Hide", show:"Show", addStock:"Add Stock", switchPreorder:"Switch to Pre-order", productName:"Product Name (English) *", productNameMM:"Product Name (Myanmar)", category:"Category", price:"Price (Ks) *", discountType:"Discount Type", discountPct:"Percent %", discountFixed:"Fixed Amount (Ks)", discountVal:"Discount Value", stockQty:"Stock Qty", enablePreorder:"Enable Pre-order", images:"Images", videoUrl:"Video", emoji:"Emoji", description:"Description", suitableFor:"Suitable For", benefits:"Benefits", usage:"How to Use", warning:"Warning", bulkTiers:"Bulk Tiers", addTier:"+ Add Tier", gdriveTip:"Use imgbb.com for images (Google Drive blocked)", shopName:"Shop Name (EN)", shopNameMM:"Shop Name (MM)", fbLink:"Facebook Messenger URL", viberNum:"Viber Number", waNum:"WhatsApp Number", phoneNum:"Phone Number", adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme", langLabel:"Language", saveSettings:"Save Settings", pending:"Pending", confirmed:"Confirmed", deposited:"Deposited", completed:"Completed", cancelled:"Cancelled", depositPaid:"Deposit Paid", balanceDue:"Balance Due", confirmOrder:"Confirm & Deduct Stock", updateDeposit:"Update Deposit", adminNote:"Admin Note", updateStatus:"Update Status", today:"Today", week:"This Week", month:"This Month", allTime:"All Time", revenue:"Revenue", totalOrders:"Total Orders", topProducts:"Top Products", recentOrders:"Recent Orders", exportExcel:"Export Excel", noData:"No data yet", noProducts:"No products found", emptyCart:"Cart is empty", orderPlaced:"Order Sent!", orderPlacedMsg:"We'll contact you soon.", wrongPw:"Wrong password", logout:"Logout", backToShop:"Back to Shop", notifications:"Notifications", clearAll:"Clear All", stockAdded:"Stock added", stockDeducted:"Stock deducted", notifUpdated:"updated", popularItems:"Popular Items" },
  mm:{ home:"ပင်မ", cats:"အမျိုးအစား", cart:"Cart", orders:"Orders", admin:"Admin", search:"ထုတ်ကုန် ရှာပါ...", all:"အားလုံး", addCart:"Cart ထည့်", buyNow:"ချက်ချင်းဝယ်", preorder:"Pre-order", outOfStock:"ကုန်ပြီ", tempOOS:"ယာယီ ကုန်ပြီ", inStock:"Stock ရှိ", qty:"အရေ", total:"စုစုပေါင်း", checkout:"Order လုပ်မည်", orderSummary:"Order အကျဉ်း", contactOrder:"ဆက်သွယ်ပြီး Order မှာမည်", trackOrder:"Order ရှာမည်", enterOrderNo:"Order နံပါတ် ထည့်ပါ", track:"ရှာမည်", yourName:"နာမည် *", yourPhone:"ဖုန်းနံပါတ် *", yourAddress:"လိပ်စာ *", note:"မှာကြားချက်", adminPanel:"Admin Panel", password:"Password", login:"Login ဝင်မည်", products:"Products", reports:"Reports", settings:"Settings", addProduct:"Product ထည့်မည်", editProduct:"Product ပြင်မည်", save:"သိမ်းမည်", cancel:"မလုပ်တော့ပါ", delete:"ဖျက်", hide:"ဖျောက်", show:"ပြ", addStock:"Stock ထည့်မည်", switchPreorder:"Pre-order ပြောင်းမည်", productName:"Product နာမည် (English) *", productNameMM:"Product နာမည် (မြန်မာ)", category:"အမျိုးအစား", price:"ဈေးနှုန်း (Ks) *", discountType:"လျှော့ဈေး အမျိုး", discountPct:"ရာခိုင်နှုန်း %", discountFixed:"ပမာဏ (Ks)", discountVal:"လျှော့ဈေး", stockQty:"Stock အရေ", enablePreorder:"Pre-order မှာနိုင်", images:"ပုံများ", videoUrl:"Video", emoji:"Emoji", description:"ဖော်ပြချက်", suitableFor:"သင့်တော်သူ", benefits:"အကျိုးကျေးဇူး", usage:"သုံးနည်း", warning:"သတိပြုရန်", bulkTiers:"အရေ Tier လျှော့ဈေး", addTier:"+ Tier ထည့်", gdriveTip:"ပုံများအတွက် imgbb.com သုံးပါ", shopName:"ဆိုင်နာမည် (EN)", shopNameMM:"ဆိုင်နာမည် (MM)", fbLink:"Facebook Messenger URL", viberNum:"Viber နံပါတ်", waNum:"WhatsApp နံပါတ်", phoneNum:"ဖုန်းနံပါတ်", adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme", langLabel:"ဘာသာ", saveSettings:"Settings သိမ်းမည်", pending:"စောင့်ဆိုင်း", confirmed:"အတည်ပြု", deposited:"စရံပေးပြီး", completed:"ပြီးစီး", cancelled:"ပယ်ဖျက်", depositPaid:"စရံပေးပြီး", balanceDue:"ကျန်ငွေ", confirmOrder:"အတည်ပြုပြီး Stock နှုတ်မည်", updateDeposit:"စရံ Update", adminNote:"Admin မှတ်ချက်", updateStatus:"Status ပြောင်းမည်", today:"ဒီနေ့", week:"ဒီအပတ်", month:"ဒီလ", allTime:"အားလုံး", revenue:"ဝင်ငွေ", totalOrders:"Order စုစုပေါင်း", topProducts:"အရောင်းကောင်းဆုံး", recentOrders:"နောက်ဆုံး Orders", exportExcel:"Excel ထုတ်", noData:"Data မရှိသေးပါ", noProducts:"ထုတ်ကုန် မတွေ့ပါ", emptyCart:"Cart ထဲ ဘာမှ မရှိသေးပါ", orderPlaced:"Order ပေးပို့ပြီးပြီ!", orderPlacedMsg:"Admin မှ မကြာမီ ဆက်သွယ်ပါမည်။", wrongPw:"Password မှားနေသည်", logout:"Logout", backToShop:"ဆိုင်သို့ ပြန်", notifications:"အကြောင်းကြားချက်", clearAll:"အားလုံးဖျက်", stockAdded:"Stock ထည့်ပြီး", stockDeducted:"Stock နှုတ်သွားသည်", notifUpdated:"ပြောင်းလဲပြီး", popularItems:"Popular Items" },
};

let _nl=[];
const NS={items:[],add(msg){NS.items=[{id:uid(),msg,time:new Date().toISOString(),read:false},...NS.items].slice(0,30);_nl.forEach(f=>f([...NS.items]));if(typeof Notification!=="undefined"&&Notification.permission==="granted")new Notification("Shop",{body:msg});},markRead(){NS.items=NS.items.map(n=>({...n,read:true}));_nl.forEach(f=>f([...NS.items]));},clear(){NS.items=[];_nl.forEach(f=>f([]));},subscribe(fn){_nl.push(fn);return()=>{_nl=_nl.filter(f=>f!==fn);};}};
function useNotifs(){const[n,setN]=useState(NS.items);useEffect(()=>NS.subscribe(setN),[]);return n;}

// GLASS HELPER
const glassStyle = (G, extra={}) => ({
  background: G.surface,
  backdropFilter: G.blur,
  WebkitBackdropFilter: G.blur,
  border: `1px solid ${G.border}`,
  ...extra,
});

const glassCard = (G, extra={}) => ({
  ...glassStyle(G),
  borderRadius: 20,
  boxShadow: G.shadow,
  ...extra,
});

const SETUP_SQL = `-- Paste in Supabase SQL Editor
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
DROP TABLE IF EXISTS shop_settings,orders,products,categories CASCADE;
CREATE TABLE categories(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL UNIQUE,sort_order INT DEFAULT 0,created_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE products(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL,name_mm TEXT,category TEXT NOT NULL DEFAULT '',description TEXT,suitable_for TEXT,benefits TEXT,usage_info TEXT,warning TEXT,price NUMERIC NOT NULL DEFAULT 0,discount_type TEXT DEFAULT 'percent',discount_value NUMERIC DEFAULT 0,bulk_discounts JSONB DEFAULT '[]'::jsonb,stock INT NOT NULL DEFAULT 0,preorder BOOLEAN DEFAULT false,images JSONB DEFAULT '[]'::jsonb,video_url TEXT,emoji TEXT DEFAULT '🛍️',visible BOOLEAN DEFAULT true,featured BOOLEAN DEFAULT false,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE orders(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),order_number TEXT,customer_name TEXT,customer_phone TEXT,customer_address TEXT,items JSONB NOT NULL DEFAULT '[]'::jsonb,total NUMERIC NOT NULL DEFAULT 0,status TEXT DEFAULT 'pending',contact_method TEXT,customer_note TEXT,deposit_paid NUMERIC DEFAULT 0,balance_due NUMERIC DEFAULT 0,admin_note TEXT,stock_deducted BOOLEAN DEFAULT false,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE shop_settings(key TEXT PRIMARY KEY,value TEXT NOT NULL DEFAULT '');
ALTER TABLE categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE shop_settings ENABLE ROW LEVEL SECURITY;
CREATE POLICY "all" ON categories FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON products FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON orders FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON shop_settings FOR ALL USING(true)WITH CHECK(true);
INSERT INTO categories(name,sort_order)VALUES('Makeup',1),('Skin Care',2),('Hair Care',3),('Fragrance',4),('Lip Care',5),('Eye Care',6),('Others',7);
INSERT INTO shop_settings(key,value)VALUES('shop_name','Beauty Cosmetics'),('shop_name_mm','ဗျူတီ ကောမ်စမက်တစ်'),('fb_link','https://m.me/yourpage'),('viber_num','+95912345678'),('wa_num','+95912345678'),('phone_num','+95912345678'),('admin_pw','admin123'),('logo',''),('banner','');
SELECT 'Done: '||COUNT(*)||' rows' FROM shop_settings;`;

function SetupScreen({onConnect}){
  const[url,setUrl]=useState("");const[key,setKey]=useState("");const[busy,setBusy]=useState(false);const[err,setErr]=useState("");const[showSql,setShowSql]=useState(false);
  const go=async()=>{if(!url||!key){setErr("URL နှင့် Key ထည့်ပါ");return;}setBusy(true);setErr("");try{const cfg={url:url.replace(/\/+$/,""),key};const r=await fetch(`${cfg.url}/rest/v1/shop_settings?limit=1`,{headers:{apikey:cfg.key,Authorization:`Bearer ${cfg.key}`}});if(!r.ok)throw new Error(await r.text());LS.set("sb_cfg",cfg);onConnect(cfg);}catch(e){setErr("ချိတ်မရပါ: "+e.message);}setBusy(false);};
  return(
    <div style={{minHeight:"100vh",background:G.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{...glassCard(G,{padding:"40px 28px",maxWidth:460,width:"100%"})}}>
        <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:56,marginBottom:8}}>💄</div><div style={{fontSize:24,fontWeight:800,color:G.text,fontFamily:"Georgia,serif"}}>Beauty Store</div><div style={{fontSize:13,color:G.muted,marginTop:4}}>Supabase ချိတ်ဆက်မည်</div></div>
        <div style={{...glassStyle(G,{borderRadius:12,padding:"12px 14px",marginBottom:14,fontSize:12,color:G.text,lineHeight:1.9})}}>
          <b>Step 1:</b> supabase.com → New Project<br/><b>Step 2:</b> SQL Editor → paste SQL → Run<br/><b>Step 3:</b> Vercel → Env Variables → VITE_SUPABASE_URL + KEY<br/><b>Step 4:</b> URL + Key ထည့်ပြီး connect
        </div>
        <button onClick={()=>setShowSql(v=>!v)} style={{width:"100%",padding:9,borderRadius:10,...glassStyle(G,{}),cursor:"pointer",fontSize:12,color:G.muted,marginBottom:10,border:`1px solid ${G.border}`}}>{showSql?"▲ SQL ပိတ်":"▼ Setup SQL ကြည့်မည်"}</button>
        {showSql&&<pre style={{background:"rgba(0,0,0,0.4)",borderRadius:10,padding:12,fontSize:9.5,fontFamily:"monospace",color:"#7FFFD4",overflow:"auto",maxHeight:160,lineHeight:1.6,marginBottom:12,whiteSpace:"pre-wrap"}}>{SETUP_SQL}</pre>}
        {["SUPABASE PROJECT URL","ANON PUBLIC KEY"].map((lbl,i)=><div key={i}><label style={{fontSize:11,fontWeight:700,color:G.textSub,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:0.5}}>{lbl}</label><input type={i===1?"password":"text"} style={{width:"100%",padding:"12px 14px",borderRadius:10,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:G.text}} value={i===0?url:key} onChange={e=>i===0?setUrl(e.target.value):setKey(e.target.value)} placeholder={i===0?"https://xxxx.supabase.co":"eyJhbGci..."}/></div>)}
        {err&&<div style={{color:G.danger,fontSize:13,marginBottom:12,...glassStyle(G,{borderRadius:8,padding:"10px 12px",background:G.dangerBg})}}>{err}</div>}
        <button onClick={go} disabled={busy} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:busy?"rgba(139,34,82,0.4)":G.primary,color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 20px ${G.primary}60`}}>{busy?"⏳ စစ်ဆေးနေသည်...":"💄 ချိတ်ဆက်မည်"}</button>
      </div>
    </div>
  );
}


// PRODUCT IMAGE with BLOB
function PImg({p,size=120,r=16,showBlob=false,G}){
  const[err,setErr]=useState(false);
  const imgs=getImgs(p);const src=imgs[0]||"";
  const blobs=[G.blob1,G.blob2,G.blob3];
  const blob=blobs[(p?.id||"x").charCodeAt(1)%3||0];
  const bgs=["rgba(255,182,193,0.3)","rgba(216,191,216,0.3)","rgba(255,218,185,0.3)","rgba(176,224,230,0.3)","rgba(255,200,180,0.3)","rgba(200,180,255,0.3)"];
  const bg=bgs[(p?.id||"x").charCodeAt(1)%bgs.length||0];
  return(
    <div style={{position:"relative",width:size,height:size,flexShrink:0}}>
      {showBlob&&<div style={{position:"absolute",inset:"-15%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",zIndex:0,filter:"blur(4px)"}}/>}
      <div style={{position:"relative",zIndex:1,width:"100%",height:"100%",borderRadius:r,overflow:"hidden",background:bg}}>
        {src&&!err?<img src={src} alt={p?.name||""} onError={()=>setErr(true)} style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}/>
          :<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.38,background:bg}}>{p?.emoji||"🛍️"}</div>}
      </div>
    </div>
  );
}

// STATUS BADGE
function SBadge({status,t,G}){
  const c=SC[status]||SC.pending;
  const L={pending:t.pending,confirmed:t.confirmed,deposited:t.deposited,completed:t.completed,cancelled:t.cancelled};
  return<span style={{background:c.bg,color:c.color,border:`1px solid ${c.border}`,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{L[status]||status}</span>;
}

// NOTIFICATION BELL
function NotifBell({G,t}){
  const notifs=useNotifs();const[open,setOpen]=useState(false);const ref=useRef();const unread=notifs.filter(n=>!n.read).length;
  useEffect(()=>{const fn=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener("mousedown",fn);return()=>document.removeEventListener("mousedown",fn);},[]);
  return(<div ref={ref} style={{position:"relative"}}>
    <button onClick={()=>{setOpen(v=>!v);if(!open)NS.markRead();}} style={{...glassStyle(G,{borderRadius:20,padding:"7px 11px",border:`1px solid ${G.border}`}),color:G.textLight,cursor:"pointer",fontSize:15,position:"relative",background:"rgba(255,255,255,0.2)"}}>
      🔔{unread>0&&<span style={{position:"absolute",top:-4,right:-4,background:G.primary,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:700,minWidth:14,textAlign:"center"}}>{unread}</span>}
    </button>
    {open&&<div style={{position:"absolute",right:0,top:"calc(100% + 8px)",width:272,maxHeight:300,overflowY:"auto",borderRadius:16,boxShadow:G.shadowMd,zIndex:300,...glassCard(G,{})}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderBottom:`1px solid ${G.border}`}}>
        <span style={{fontWeight:700,fontSize:13,color:G.text}}>{t.notifications}</span>
        <button onClick={()=>NS.clear()} style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:G.muted}}>{t.clearAll}</button>
      </div>
      {notifs.length===0?<div style={{padding:20,textAlign:"center",color:G.muted,fontSize:13}}>—</div>
        :notifs.map(n=><div key={n.id} style={{padding:"10px 14px",borderBottom:`1px solid ${G.border}`,opacity:n.read?0.6:1}}><div style={{fontSize:12,fontWeight:600,color:G.text}}>{n.msg}</div><div style={{fontSize:10,color:G.muted,marginTop:2}}>{fdate(n.time)}</div></div>)}
    </div>}
  </div>);
}

// TOP BAR — Beauty Glass style
function TopBar({G,t,shopMM,logo,cartCount,searchQ,setSearchQ,showDrop,setShowDrop,dropRes,onHit,onCart,onAdmin,onLogo,onLangToggle,onThemeToggle,themeName}){
  return(<div style={{...glassStyle(G,{position:"sticky",top:0,zIndex:200,boxShadow:G.shadowMd,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 16px"}}>
      <button onClick={onLogo} style={{background:"none",border:"none",cursor:"pointer",padding:0,display:"flex",alignItems:"center",gap:8}}>
        {logo?<img src={logo} style={{width:34,height:34,borderRadius:10,objectFit:"cover"}} alt="logo"/>:<div style={{width:34,height:34,borderRadius:10,background:G.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>💄</div>}
        <span style={{fontWeight:800,fontSize:17,color:G.text,fontFamily:"Georgia,serif",letterSpacing:0.3}}>{shopMM}</span>
      </button>
      <div style={{flex:1}}/>
      <button onClick={onLangToggle} style={{...glassStyle(G,{borderRadius:12,padding:"5px 9px",border:`1px solid ${G.border}`}),color:G.text,cursor:"pointer",fontSize:11,fontWeight:700,background:"rgba(255,255,255,0.2)"}}>{t===TR.mm?"EN":"မြ"}</button>
      <button onClick={onThemeToggle} style={{...glassStyle(G,{borderRadius:12,padding:"5px 9px",border:`1px solid ${G.border}`}),color:G.text,cursor:"pointer",fontSize:13,background:"rgba(255,255,255,0.2)"}}>{themeName==="glass"?"🌙":"💄"}</button>
      <NotifBell G={G} t={t}/>
      <button onClick={onCart} style={{...glassStyle(G,{borderRadius:20,padding:"7px 14px",border:`1px solid ${G.border}`}),color:G.text,cursor:"pointer",display:"flex",alignItems:"center",gap:5,fontSize:14,fontWeight:700,background:"rgba(255,255,255,0.2)"}}>
        🛒{cartCount>0&&<span style={{background:G.primary,color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:11,fontWeight:800}}>{cartCount}</span>}
      </button>
      <button onClick={onAdmin} style={{...glassStyle(G,{borderRadius:20,padding:"7px 11px",border:`1px solid ${G.border}`}),color:G.text,cursor:"pointer",fontSize:13,background:"rgba(255,255,255,0.15)"}}>⚙️</button>
    </div>
    <div style={{padding:"0 16px 12px",position:"relative"}}>
      <div style={{position:"relative"}}>
        <span style={{position:"absolute",left:14,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none",color:G.muted}}>🔍</span>
        <input value={searchQ} onChange={e=>{setSearchQ(e.target.value);setShowDrop(true);}} onFocus={()=>searchQ&&setShowDrop(true)} onBlur={()=>setTimeout(()=>setShowDrop(false),200)}
          placeholder={t.search}
          style={{width:"100%",padding:"11px 16px 11px 38px",borderRadius:24,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",fontFamily:"inherit",color:G.text,background:G.surfaceStrong}}/>
      </div>
      {showDrop&&searchQ&&<div style={{position:"absolute",top:"100%",left:16,right:16,...glassCard(G,{zIndex:300,overflow:"hidden",maxHeight:280,overflowY:"auto"})}}>
        {dropRes.length===0?<div style={{padding:14,textAlign:"center",color:G.muted,fontSize:13}}>"{searchQ}" {t.noProducts}</div>
          :dropRes.map(p=><div key={p.id} onMouseDown={()=>onHit(p)} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",cursor:"pointer",borderBottom:`1px solid ${G.border}`}} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.1)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <PImg p={p} size={36} r={8} G={G}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:G.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,color:G.muted}}>{p.category} · {fmt(calcUnit(p))}</div></div>
          </div>)}
      </div>}
    </div>
  </div>);
}

// BOTTOM NAV — Glass beauty style
function BottomNav({G,t,tab,setTab,cartCount}){
  const items=[{k:"home",icon:"🏠",label:t.home},{k:"cats",icon:"✨",label:t.cats},{k:"cart",icon:"🛒",label:t.cart,badge:cartCount},{k:"track",icon:"📋",label:t.orders}];
  return(<div style={{position:"fixed",bottom:0,left:0,right:0,...glassStyle(G,{borderTop:`1px solid ${G.border}`,display:"flex",zIndex:190,boxShadow:`0 -4px 20px rgba(139,34,82,0.1)`}),background:G.navBg}}>
    {items.map(it=><button key={it.k} onClick={()=>setTab(it.k)} style={{flex:1,padding:"8px 4px 10px",border:"none",background:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:3,position:"relative",transition:"all 0.2s"}}>
      <div style={{width:38,height:38,borderRadius:19,background:tab===it.k?G.primary:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,transition:"all 0.2s",boxShadow:tab===it.k?`0 4px 12px ${G.primary}60`:undefined}}>
        {it.icon}
      </div>
      {it.badge>0&&<span style={{position:"absolute",top:4,right:"calc(50% - 18px)",background:G.primary,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:800,minWidth:14,textAlign:"center"}}>{it.badge}</span>}
      <span style={{fontSize:10,color:tab===it.k?G.primary:G.muted,fontWeight:tab===it.k?700:400}}>{it.label}</span>
    </button>)}
  </div>);
}

// PRODUCT CARD — Beauty glass style with blob
function ProductCard({G,t,p,onClick,onAdd}){
  const[added,setAdded]=useState(false);const up=calcUnit(p);const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;
  const doAdd=e=>{e.stopPropagation();if(!canBuy)return;onAdd(p);setAdded(true);setTimeout(()=>setAdded(false),1800);};
  const blobs=[G.blob1,G.blob2,G.blob3];const blob=blobs[(p?.id||"x").charCodeAt(1)%3||0];
  return(<div onClick={()=>onClick(p)} style={{...glassCard(G,{overflow:"hidden",cursor:"pointer",transition:"all 0.25s",position:"relative"})} } onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-6px)";e.currentTarget.style.boxShadow=G.shadowLg;}} onMouseLeave={e=>{e.currentTarget.style.transform="none";e.currentTarget.style.boxShadow=G.shadow;}}>
    {dv>0&&<div style={{position:"absolute",top:10,left:10,background:G.primary,color:"#fff",borderRadius:10,padding:"3px 9px",fontSize:10,fontWeight:800,zIndex:2,boxShadow:`0 2px 8px ${G.primary}60`}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</div>}
    {p.preorder&&<div style={{position:"absolute",top:10,right:10,background:"rgba(139,94,0,0.8)",backdropFilter:"blur(8px)",color:"#fff",borderRadius:10,padding:"3px 8px",fontSize:9,fontWeight:700,zIndex:2}}>Pre-order</div>}
    {oos&&!p.preorder&&<div style={{position:"absolute",top:10,right:10,background:"rgba(0,0,0,0.45)",backdropFilter:"blur(8px)",color:"#fff",borderRadius:10,padding:"3px 8px",fontSize:9,zIndex:2}}>{t.tempOOS}</div>}
    {/* Image area with blob */}
    <div style={{background:"rgba(255,255,255,0.08)",position:"relative",display:"flex",justifyContent:"center",alignItems:"center",padding:"20px 16px 12px",minHeight:150,overflow:"hidden"}}>
      <div style={{position:"absolute",inset:"-20%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",filter:"blur(8px)",opacity:0.7}}/>
      <div style={{position:"relative",zIndex:1}}><PImg p={p} size={110} r={14} G={G}/></div>
    </div>
    <div style={{padding:"10px 14px 14px"}}>
      <div style={{fontSize:10,color:G.muted,marginBottom:3,fontWeight:600,textTransform:"uppercase",letterSpacing:0.8}}>{p.category}</div>
      <div style={{fontSize:13,fontWeight:700,color:G.text,lineHeight:1.35,marginBottom:8,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",minHeight:35,fontFamily:"Georgia,serif"}}>{p.name_mm||p.name}</div>
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10}}>
        <span style={{fontWeight:800,color:G.primary,fontSize:15}}>{fmt(up)}</span>
        {dv>0&&<span style={{fontSize:11,color:G.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
      </div>
      <button onClick={doAdd} disabled={!canBuy} style={{width:"100%",padding:"9px 0",borderRadius:12,border:"none",cursor:canBuy?"pointer":"not-allowed",background:added?G.success:(oos&&!p.preorder?"rgba(0,0,0,0.15)":G.primary),color:oos&&!p.preorder?G.muted:"#fff",fontSize:12,fontWeight:700,transition:"all 0.2s",boxShadow:canBuy&&!added?`0 4px 12px ${G.primary}50`:undefined}}>
        {added?`✓ Added`:(p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:t.addCart))}
      </button>
    </div>
  </div>);
}

// HOME PAGE — Beauty style
function HomePage({G,t,products,cats,catFilter,setCatFilter,onOpen,onAdd,banner,shopMM}){
  const featured=products.filter(p=>p.featured||(p.discount_value||0)>0).slice(0,6);
  const filtered=catFilter==="all"?products:products.filter(p=>p.category===catFilter);
  const catIcons={"Makeup":"💄","Skin Care":"✨","Hair Care":"💆","Fragrance":"🌸","Lip Care":"💋","Eye Care":"👁️","Others":"🛍️","ဆံပင်":"💆","မျက်နှာ":"✨","နှုတ်ခမ်း":"💋","မျက်လုံး":"👁️","အဝတ်အထည်":"👗","အိတ်":"👜"};
  return(<div style={{paddingBottom:80}}>
    {/* Hero banner or header */}
    {banner?<div style={{margin:"10px 14px",borderRadius:20,overflow:"hidden",height:140,boxShadow:G.shadow}}><img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="banner"/></div>
    :<div style={{margin:"10px 14px",...glassCard(G,{padding:"20px 20px",textAlign:"center"})}}>
      <div style={{fontSize:11,color:G.muted,letterSpacing:2,textTransform:"uppercase",marginBottom:4}}>Welcome to</div>
      <div style={{fontSize:24,fontWeight:800,color:G.text,fontFamily:"Georgia,serif",marginBottom:4}}>{shopMM}</div>
      <div style={{fontSize:12,color:G.muted}}>✨ Premium Beauty & Cosmetics ✨</div>
    </div>}
    {/* Category circles */}
    <div style={{padding:"14px 14px 4px"}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,marginBottom:10,fontFamily:"Georgia,serif"}}>Explore Categories</div>
      <div style={{display:"flex",gap:12,overflowX:"auto",scrollbarWidth:"none",paddingBottom:4}}>
        {[{name:"all",label:t.all,icon:"🌟"},...cats.map(c=>({name:c.name,label:c.name,icon:catIcons[c.name]||"💄"}))].map(c=>(
          <div key={c.name} onClick={()=>setCatFilter(c.name)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:6,cursor:"pointer",flexShrink:0}}>
            <div style={{width:58,height:58,borderRadius:29,...glassStyle(G,{}),display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,border:catFilter===c.name?`2px solid ${G.primary}`:`1px solid ${G.border}`,background:catFilter===c.name?G.primary:"rgba(255,255,255,0.25)",boxShadow:catFilter===c.name?`0 4px 16px ${G.primary}50`:G.shadow,transition:"all 0.2s"}}>{c.icon}</div>
            <span style={{fontSize:10,fontWeight:catFilter===c.name?700:500,color:catFilter===c.name?G.primary:G.text,textAlign:"center",maxWidth:60,lineHeight:1.2}}>{c.label}</span>
          </div>
        ))}
      </div>
    </div>
    {/* Popular / Flash deals */}
    {featured.length>0&&catFilter==="all"&&<div style={{padding:"14px 14px 4px"}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,marginBottom:10,fontFamily:"Georgia,serif"}}>🔥 {t.popularItems}</div>
      <div style={{display:"flex",gap:12,overflowX:"auto",scrollbarWidth:"none"}}>
        {featured.map(p=><div key={p.id} onClick={()=>onOpen(p)} style={{flexShrink:0,width:130,...glassCard(G,{overflow:"hidden",cursor:"pointer",transition:"all 0.2s"})}}>
          <div style={{background:"rgba(255,255,255,0.08)",position:"relative",display:"flex",justifyContent:"center",padding:"12px 10px 8px",overflow:"hidden"}}>
            <div style={{position:"absolute",inset:"-20%",background:[G.blob1,G.blob2,G.blob3][(p.id||"x").charCodeAt(1)%3||0],borderRadius:"50%",filter:"blur(10px)",opacity:0.6}}/>
            <div style={{position:"relative",zIndex:1}}><PImg p={p} size={80} r={10} G={G}/></div>
          </div>
          <div style={{padding:"8px 10px 12px"}}>
            <div style={{fontSize:11,fontWeight:700,color:G.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontFamily:"Georgia,serif"}}>{p.name_mm||p.name}</div>
            <div style={{fontSize:13,fontWeight:800,color:G.primary,marginTop:3}}>{fmt(calcUnit(p))}</div>
          </div>
        </div>)}
      </div>
    </div>}
    {/* Products grid */}
    <div style={{padding:"14px 14px 0"}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,marginBottom:12,fontFamily:"Georgia,serif"}}>{catFilter==="all"?`✨ All Products`:`📦 ${catFilter}`}</div>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:G.muted}}><div style={{fontSize:52,marginBottom:10}}>💄</div><div style={{fontSize:15,color:G.text}}>{t.noProducts}</div></div>
        :<div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:14}}>
          {filtered.map(p=><ProductCard key={p.id} G={G} t={t} p={p} onClick={onOpen} onAdd={onAdd}/>)}
        </div>}
    </div>
  </div>);
}


// PRODUCT DETAIL PAGE
function ProductDetail({G,t,p,onBack,addToCart,onBuyNow}){
  const[qty,setQty]=useState(1);const[added,setAdded]=useState(false);const[imgIdx,setImgIdx]=useState(0);
  const imgs=getImgs(p);let bd=p.bulk_discounts;if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch{bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>a.min_qty-b.min_qty);
  const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;const up=calcUnit(p,qty);
  const blob=[G.blob1,G.blob2,G.blob3][(p?.id||"x").charCodeAt(1)%3||0];
  const Sec=({title,txt,warn})=>txt?<div style={{marginBottom:14}}><div style={{fontSize:12,fontWeight:700,color:warn?G.danger:G.textSub,marginBottom:5,textTransform:"uppercase",letterSpacing:0.5}}>{title}</div><div style={{fontSize:13,color:G.text,lineHeight:1.8,...glassStyle(G,{borderRadius:12,padding:"12px 14px",background:warn?G.dangerBg:G.surface2})}}>{txt}</div></div>:null;
  return(<div style={{minHeight:"100vh",paddingBottom:100}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,flex:1,color:G.text,fontFamily:"Georgia,serif",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</span>
    </div>
    {/* Image hero with blob */}
    <div style={{position:"relative",display:"flex",justifyContent:"center",alignItems:"center",padding:"30px 20px 20px",minHeight:280,overflow:"hidden"}}>
      <div style={{position:"absolute",width:"70%",height:"70%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",filter:"blur(20px)",opacity:0.7,zIndex:0}}/>
      <div style={{position:"absolute",width:"50%",height:"50%",background:G.blob2,borderRadius:"40% 60% 45% 55% / 55% 45% 55% 45%",filter:"blur(15px)",opacity:0.5,zIndex:0,transform:"translate(30%,20%)"}}/>
      <div style={{position:"relative",zIndex:1}}>
        {imgs[imgIdx]?<img src={imgs[imgIdx]} style={{width:220,height:220,objectFit:"contain",borderRadius:20,filter:"drop-shadow(0 8px 24px rgba(139,34,82,0.25))"}} alt={p.name} onError={e=>e.target.style.display="none"}/>:<PImg p={p} size={220} r={20} G={G}/>}
      </div>
    </div>
    {/* Thumbnail strip */}
    {imgs.length>1&&<div style={{display:"flex",gap:10,padding:"0 16px 12px",justifyContent:"center"}}>
      {imgs.map((img,i)=><div key={i} onClick={()=>setImgIdx(i)} style={{width:52,height:52,borderRadius:12,overflow:"hidden",border:`2px solid ${i===imgIdx?G.primary:G.border}`,cursor:"pointer",transition:"all 0.15s",boxShadow:i===imgIdx?`0 4px 12px ${G.primary}50`:undefined}}>
        <img src={img} style={{width:"100%",height:"100%",objectFit:"cover"}} alt=""/>
      </div>)}
    </div>}
    {/* Product info */}
    <div style={{padding:"0 16px"}}>
      <div style={{...glassCard(G,{padding:"18px 18px",marginBottom:12})}}>
        <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
          <span style={{background:G.tag,color:G.tagText,padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:700}}>{p.category}</span>
          {p.preorder&&<span style={{background:G.warningBg,color:G.warning,padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:700}}>Pre-order</span>}
          {(p.stock||0)>0&&!p.preorder&&<span style={{background:G.successBg,color:G.success,padding:"4px 12px",borderRadius:20,fontSize:11}}>{t.inStock}: {p.stock}</span>}
          {oos&&!p.preorder&&<span style={{background:"rgba(0,0,0,0.1)",color:G.muted,padding:"4px 12px",borderRadius:20,fontSize:11}}>⚠️ {t.tempOOS}</span>}
        </div>
        <div style={{fontSize:22,fontWeight:800,color:G.text,marginBottom:4,fontFamily:"Georgia,serif",lineHeight:1.3}}>{p.name_mm||p.name}</div>
        {p.name_mm&&<div style={{fontSize:13,color:G.muted,marginBottom:14}}>{p.name}</div>}
        <div style={{display:"flex",alignItems:"baseline",gap:12,marginBottom:6,flexWrap:"wrap"}}>
          <span style={{fontSize:28,fontWeight:800,color:G.primary}}>{fmt(up)}</span>
          {dv>0&&<span style={{fontSize:16,color:G.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
          {dv>0&&<span style={{background:G.primary,color:"#fff",borderRadius:8,padding:"3px 10px",fontSize:12,fontWeight:700,boxShadow:`0 2px 8px ${G.primary}50`}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</span>}
        </div>
      </div>
      {/* Bulk discounts */}
      {bulks.length>0&&<div style={{...glassCard(G,{padding:"14px 16px",marginBottom:12,background:"rgba(139,34,82,0.08)"})}}>
        <div style={{fontSize:12,fontWeight:700,color:G.primary,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>🎁 {t.bulkTiers}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {bulks.map((b,i)=><div key={i} style={{background:qty>=b.min_qty?G.primary:"rgba(139,34,82,0.15)",color:qty>=b.min_qty?"#fff":G.primary,borderRadius:10,padding:"5px 12px",fontSize:12,fontWeight:600,transition:"all 0.2s",border:`1px solid ${qty>=b.min_qty?G.primary:"rgba(139,34,82,0.3)"}`}}>{b.min_qty}+ → -{b.discount_percent}%</div>)}
        </div>
      </div>}
      {/* Qty */}
      {canBuy&&<div style={{...glassCard(G,{padding:"12px 16px",marginBottom:12})}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <span style={{fontSize:12,color:G.textSub,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{t.qty}</span>
          <div style={{display:"flex",alignItems:"center",gap:10,...glassStyle(G,{borderRadius:16,padding:"4px 6px"})}}>
            <button onClick={()=>setQty(q=>Math.max(1,q-1))} style={{width:34,height:34,borderRadius:17,border:`1px solid ${G.border}`,background:"rgba(255,255,255,0.3)",cursor:"pointer",fontSize:18,fontWeight:700,color:G.text,display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
            <span style={{minWidth:32,textAlign:"center",fontWeight:800,fontSize:17,color:G.text}}>{qty}</span>
            <button onClick={()=>setQty(q=>q+1)} style={{width:34,height:34,borderRadius:17,border:`1px solid ${G.border}`,background:"rgba(255,255,255,0.3)",cursor:"pointer",fontSize:18,fontWeight:700,color:G.text,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
          </div>
          <span style={{fontSize:13,color:G.muted}}>= <b style={{color:G.primary,fontSize:15}}>{fmt(up*qty)}</b></span>
        </div>
      </div>}
      {/* Info sections */}
      <div style={{...glassCard(G,{padding:"16px",marginBottom:12})}}>
        <Sec title={`📝 ${t.description}`} txt={p.description}/>
        <Sec title={`👤 ${t.suitableFor}`} txt={p.suitable_for}/>
        <Sec title={`✅ ${t.benefits}`} txt={p.benefits}/>
        <Sec title={`📋 ${t.usage}`} txt={p.usage_info}/>
        <Sec title={`⚠️ ${t.warning}`} txt={p.warning} warn/>
      </div>
      {p.video_url&&<div style={{...glassCard(G,{padding:"14px",marginBottom:12})}}><div style={{fontSize:12,fontWeight:700,color:G.textSub,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>📹 Video</div><video src={p.video_url} controls style={{width:"100%",borderRadius:12,background:"#000"}}/></div>}
    </div>
    {/* Bottom action */}
    <div style={{position:"fixed",bottom:0,left:0,right:0,...glassStyle(G,{padding:"12px 16px",boxShadow:`0 -4px 20px rgba(139,34,82,0.15)`,display:"flex",gap:12,zIndex:190})}}>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);setAdded(true);setTimeout(()=>setAdded(false),1800);}} disabled={!canBuy}
        style={{flex:1,padding:"14px 0",borderRadius:14,border:`2px solid ${canBuy?G.primary:G.border}`,background:"transparent",color:canBuy?G.primary:G.muted,cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700,transition:"all 0.2s"}}>
        {added?"✓ Added!":t.addCart}
      </button>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);onBuyNow();}} disabled={!canBuy}
        style={{flex:1,padding:"14px 0",borderRadius:14,border:"none",background:canBuy?G.primary:"rgba(0,0,0,0.15)",color:"#fff",cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700,boxShadow:canBuy?`0 4px 16px ${G.primary}60`:undefined,transition:"all 0.2s"}}>
        {p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:t.buyNow||"Buy Now")}
      </button>
    </div>
  </div>);
}

// CART PAGE
function CartPage({G,t,cart,updateQty,removeItem,total,onCheckout,onBack}){
  if(!cart.length)return(<div style={{minHeight:"100vh",paddingBottom:80}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,color:G.text,fontFamily:"Georgia,serif"}}>{t.cart}</span>
    </div>
    <div style={{textAlign:"center",padding:"100px 20px",color:G.muted}}>
      <div style={{fontSize:64,marginBottom:16}}>🛒</div>
      <div style={{fontSize:18,fontWeight:700,color:G.text,fontFamily:"Georgia,serif"}}>{t.emptyCart}</div>
    </div>
  </div>);
  return(<div style={{minHeight:"100vh",paddingBottom:100}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,color:G.text,fontFamily:"Georgia,serif"}}>{t.cart} ({cart.length})</span>
    </div>
    <div style={{padding:"12px 14px 0",display:"flex",flexDirection:"column",gap:12}}>
      {cart.map(item=>{const up=calcUnit(item.p,item.qty);return(
        <div key={item.p.id} style={{...glassCard(G,{padding:"14px",display:"flex",gap:12,alignItems:"flex-start"})}}>
          <PImg p={item.p} size={68} r={12} G={G}/>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",lineHeight:1.35,marginBottom:3}}>{item.p.name_mm||item.p.name}</div>
            {item.p.preorder&&<span style={{fontSize:10,background:G.warningBg,color:G.warning,padding:"1px 7px",borderRadius:8,fontWeight:700,display:"inline-block",marginBottom:5}}>Pre-order</span>}
            <div style={{fontSize:15,color:G.primary,fontWeight:800,marginBottom:10}}>{fmt(up)}</div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <button onClick={()=>updateQty(item.p.id,item.qty-1)} style={{width:30,height:30,borderRadius:15,...glassStyle(G,{border:`1px solid ${G.border}`}),cursor:"pointer",fontWeight:700,color:G.text,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>−</button>
              <span style={{minWidth:24,textAlign:"center",fontWeight:800,fontSize:15,color:G.text}}>{item.qty}</span>
              <button onClick={()=>updateQty(item.p.id,item.qty+1)} style={{width:30,height:30,borderRadius:15,...glassStyle(G,{border:`1px solid ${G.border}`}),cursor:"pointer",fontWeight:700,color:G.text,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>+</button>
              <span style={{marginLeft:"auto",fontSize:14,fontWeight:700,color:G.text}}>{fmt(up*item.qty)}</span>
            </div>
          </div>
          <button onClick={()=>removeItem(item.p.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:G.muted,padding:0}}>✕</button>
        </div>
      );})}
    </div>
    <div style={{position:"sticky",bottom:0,...glassStyle(G,{padding:"14px 16px",boxShadow:`0 -4px 20px rgba(139,34,82,0.12)`,display:"flex",gap:12,alignItems:"center",marginTop:14})}}>
      <div style={{flex:1}}>
        <div style={{fontSize:11,color:G.muted,textTransform:"uppercase",letterSpacing:0.5}}>{t.total}</div>
        <div style={{fontSize:24,fontWeight:800,color:G.primary}}>{fmt(total)}</div>
      </div>
      <button onClick={onCheckout} style={{padding:"14px 28px",borderRadius:14,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${G.primary}60`}}>{t.checkout} →</button>
    </div>
  </div>);
}

// CHECKOUT PAGE
function CheckoutPage({G,t,cart,total,settings,onPlaced,onBack}){
  const[name,setName]=useState("");const[phone,setPhone]=useState("");const[address,setAddress]=useState("");const[note,setNote]=useState("");const[sent,setSent]=useState(false);const[sentOrder,setSentOrder]=useState(null);
  const hasPreorder=cart.some(i=>i.p.preorder);
  const msgLines=cart.map(i=>`• ${i.p.name_mm||i.p.name}${i.p.preorder?" [PRE-ORDER]":""} x${i.qty} = ${fmt(calcUnit(i.p,i.qty)*i.qty)}`).join("\n");
  const msg=encodeURIComponent(`မင်္ဂလာပါ!\nName: ${name}\nPhone: ${phone}\nAddress: ${address}\n\nOrder:\n${msgLines}\nTotal: ${fmt(total)}${hasPreorder?"\n⚠️ Pre-order ပါဝင်":""}${note?"\nNote: "+note:""}`);
  const saveOrder=async(method)=>{
    if(sent)return;
    const newOrd={order_number:onum(),customer_name:name,customer_phone:phone,customer_address:address,customer_note:note,contact_method:method,items:cart.map(i=>({id:i.p.id,name:i.p.name,name_mm:i.p.name_mm,qty:i.qty,is_preorder:i.p.preorder,unit_price:calcUnit(i.p,i.qty),total:calcUnit(i.p,i.qty)*i.qty})),total,status:"pending",deposit_paid:0,balance_due:total,stock_deducted:false,admin_note:"",created_at:new Date().toISOString()};
    try{const rows=await sb("POST","orders",newOrd);const saved=Array.isArray(rows)?rows[0]:newOrd;setSentOrder(saved);setSent(true);onPlaced(saved);}catch(e){alert("Error: "+e.message);}
  };
  const contacts=[
    {label:"Facebook",sub:"Messenger",color:"#1877F2",emoji:"💬",key:"messenger",url:`https://m.me/${(settings.fb_link||"").replace(/^https?:\/\/m\.me\//,"")}?text=${msg}`},
    {label:"Viber",sub:settings.viber_num,color:"#7360F2",emoji:"📱",key:"viber",url:`viber://chat?number=${encodeURIComponent(settings.viber_num||"")}&text=${msg}`},
    {label:"WhatsApp",sub:settings.wa_num,color:"#25D366",emoji:"💚",key:"whatsapp",url:`https://wa.me/${(settings.wa_num||"").replace(/\D/g,"")}?text=${msg}`},
    {label:"Phone",sub:settings.phone_num,color:"#546E7A",emoji:"📞",key:"phone",url:`tel:${settings.phone_num}`},
  ];
  const inp={width:"100%",padding:"12px 14px",borderRadius:12,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:G.text,background:G.surfaceStrong};
  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,color:G.text,fontFamily:"Georgia,serif"}}>{t.checkout}</span>
    </div>
    {sent?<div style={{padding:20}}>
      <div style={{...glassCard(G,{padding:32,textAlign:"center"})}}>
        <div style={{fontSize:56,marginBottom:12}}>💄</div>
        <div style={{fontSize:20,fontWeight:800,color:G.primary,marginBottom:8,fontFamily:"Georgia,serif"}}>{t.orderPlaced}</div>
        <div style={{fontSize:13,color:G.muted,marginBottom:16}}>{t.orderPlacedMsg}</div>
        {sentOrder&&<div style={{...glassStyle(G,{borderRadius:10,padding:"10px 14px",fontSize:12,color:G.text})}}>Order #: <b>{sentOrder.order_number}</b></div>}
      </div>
    </div>
    :<div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:12}}>
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:10}}>{t.orderSummary}</div>
        {hasPreorder&&<div style={{background:G.warningBg,borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:G.warning,fontWeight:600}}>⚠️ Pre-order items — timeline ညှိနှိုင်းပါ</div>}
        {cart.map(item=><div key={item.p.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${G.border}`,fontSize:13}}><span style={{color:G.text,flex:1,paddingRight:10}}>{item.p.name_mm||item.p.name} × {item.qty}</span><span style={{fontWeight:700,color:G.text,whiteSpace:"nowrap"}}>{fmt(calcUnit(item.p,item.qty)*item.qty)}</span></div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"12px 0 0",fontSize:17,fontWeight:800}}><span style={{color:G.text}}>{t.total}</span><span style={{color:G.primary}}>{fmt(total)}</span></div>
      </div>
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>📋 Customer Info</div>
        <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder={t.yourName}/>
        <input style={inp} value={phone} onChange={e=>setPhone(e.target.value)} placeholder={t.yourPhone} type="tel"/>
        <input style={inp} value={address} onChange={e=>setAddress(e.target.value)} placeholder={t.yourAddress}/>
        <textarea value={note} onChange={e=>setNote(e.target.value)} placeholder={t.note} style={{...inp,minHeight:56,resize:"vertical",marginBottom:0}}/>
      </div>
      {(!name.trim()||!phone.trim())&&<div style={{...glassStyle(G,{borderRadius:10,padding:"10px 12px",background:G.warningBg}),fontSize:12,color:G.warning}}>⚠️ Name & Phone ထည့်ပြီးမှ order လုပ်ပါ</div>}
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>{t.contactOrder}</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {contacts.map(ct=><a key={ct.key} href={(name.trim()&&phone.trim())?ct.url:"#"} onClick={e=>{if(!name.trim()||!phone.trim()){e.preventDefault();alert("Name & Phone ထည့်ပါ");return;}saveOrder(ct.key);}} target="_blank" rel="noopener"
            style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",background:ct.color,borderRadius:14,textDecoration:"none",color:"#fff",boxShadow:`0 4px 14px ${ct.color}60`,opacity:(name.trim()&&phone.trim())?1:0.5,transition:"transform 0.15s"}}
            onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"} onMouseLeave={e=>e.currentTarget.style.transform="none"}>
            <span style={{fontSize:22}}>{ct.emoji}</span><div><div style={{fontWeight:700,fontSize:14}}>{ct.label}</div><div style={{fontSize:11,opacity:0.85}}>{ct.sub}</div></div><span style={{marginLeft:"auto",fontSize:18,opacity:0.8}}>→</span>
          </a>)}
        </div>
      </div>
    </div>}
  </div>);
}

// ORDER TRACKING
function TrackOrder({G,t}){
  const[q,setQ]=useState("");const[result,setResult]=useState(null);const[busy,setBusy]=useState(false);const[err,setErr]=useState("");
  const search=async()=>{if(!q.trim())return;setBusy(true);setErr("");setResult(null);try{const rows=await sb("GET","orders",null,`order_number=eq.${q.trim()}`);if(!rows||rows.length===0)setErr("Order မတွေ့ပါ");else setResult(rows[0]);}catch(e){setErr(e.message);}setBusy(false);};
  return(<div style={{padding:"16px 14px",paddingBottom:90}}>
    <div style={{fontSize:20,fontWeight:800,color:G.text,fontFamily:"Georgia,serif",marginBottom:4}}>📋 {t.trackOrder}</div>
    <div style={{fontSize:13,color:G.muted,marginBottom:16}}>Order number ဖြင့် status စစ်ဆေးနိုင်သည်</div>
    <div style={{display:"flex",gap:8,marginBottom:14}}>
      <input value={q} onChange={e=>setQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&search()} placeholder={t.enterOrderNo}
        style={{flex:1,padding:"12px 14px",borderRadius:12,...glassStyle(G,{}),fontSize:14,outline:"none",fontFamily:"inherit",color:G.text,minWidth:0,boxSizing:"border-box",background:G.surfaceStrong}}/>
      <button onClick={search} disabled={busy} style={{padding:"12px 18px",borderRadius:12,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0,boxShadow:`0 4px 12px ${G.primary}50`}}>{busy?"...":t.track}</button>
    </div>
    {err&&<div style={{...glassStyle(G,{borderRadius:12,padding:"12px 14px",background:G.dangerBg,marginBottom:12}),fontSize:13,color:G.danger}}>⚠️ {err}</div>}
    {result&&<div style={{...glassCard(G,{padding:20})}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}><div><div style={{fontSize:17,fontWeight:800,color:G.text,fontFamily:"Georgia,serif"}}>{result.order_number}</div><div style={{fontSize:12,color:G.muted}}>{fdate(result.created_at)}</div></div><SBadge status={result.status} t={t} G={G}/></div>
      {(result.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${G.border}`,fontSize:13}}><span style={{color:G.text}}>{item.name_mm||item.name} × {item.qty}{item.is_preorder?" [PRE]":""}</span><span style={{fontWeight:700,color:G.text}}>{fmt(item.total)}</span></div>)}
      <div style={{paddingTop:12,borderTop:`1px solid ${G.border}`,marginTop:4}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:G.text}}>{t.total}</span><span style={{color:G.primary}}>{fmt(result.total)}</span></div>
        {(result.deposit_paid||0)>0&&<><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{color:G.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:G.success,fontWeight:700}}>{fmt(result.deposit_paid)}</span></div><div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:G.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:G.danger,fontWeight:700}}>{fmt(Math.max(0,result.total-(result.deposit_paid||0)))}</span></div></>}
      </div>
      {result.admin_note&&<div style={{...glassStyle(G,{borderRadius:10,padding:"10px 12px",marginTop:12}),fontSize:12,color:G.textSub}}><b>Note:</b> {result.admin_note}</div>}
    </div>}
  </div>);
}


// ADMIN LOGIN BOX
function AdminLoginBox({G,t,adminPw,onSuccess,onBack}){
  const[pw,setPw]=useState("");const[err,setErr]=useState("");
  const check=()=>{if(pw===(adminPw||"admin123"))onSuccess();else setErr(t.wrongPw);};
  const inp={width:"100%",padding:"13px 16px",borderRadius:12,...glassStyle(G,{}),fontSize:14,outline:"none",boxSizing:"border-box",fontFamily:"inherit",color:G.text,background:G.surfaceStrong};
  return(<>
    <input type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&check()} placeholder="admin123" style={{...inp,marginBottom:err?8:16,border:`1.5px solid ${err?G.danger:G.border}`}}/>
    {err&&<div style={{color:G.danger,fontSize:12,marginBottom:14,...glassStyle(G,{borderRadius:8,padding:"8px 12px",background:G.dangerBg})}}>{err}</div>}
    <button onClick={check} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:10,boxShadow:`0 4px 16px ${G.primary}60`}}>{t.login}</button>
    <button onClick={onBack} style={{width:"100%",padding:12,borderRadius:12,border:`1.5px solid ${G.border}`,background:"transparent",color:G.muted,cursor:"pointer",fontSize:13}}>{t.backToShop}</button>
  </>);
}

// PRODUCT FORM (Admin)
function ProductForm({G,t,product,cats,onSave,onCancel}){
  const blank={name:"",name_mm:"",category:cats[0]?.name||"",description:"",suitable_for:"",benefits:"",usage_info:"",warning:"",price:"",discount_type:"percent",discount_value:"0",bulk_discounts:[],stock:"0",preorder:false,images:[],video_url:"",emoji:"💄",featured:false};
  const[form,setForm]=useState(product?{...product,discount_value:product.discount_value??0,bulk_discounts:Array.isArray(product.bulk_discounts)?product.bulk_discounts:[]}:blank);
  const[saving,setSaving]=useState(false);const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const iRef=useRef();const vRef=useRef();const[urlIn,setUrlIn]=useState("");
  const upImg=async e=>{const b64s=await Promise.all(Array.from(e.target.files||[]).map(fileToB64));set("images",[...(form.images||[]),...b64s]);e.target.value="";};
  const upVid=async e=>{const f=e.target.files?.[0];if(!f)return;set("video_url",await fileToB64(f));e.target.value="";};
  const addUrl=()=>{if(!urlIn.trim())return;set("images",[...(form.images||[]),urlIn.trim()]);setUrlIn("");};
  const handleSave=async()=>{
    if(!form.name.trim()||!form.price){alert("Name & Price required");return;}
    setSaving(true);
    try{const data={...form,price:Number(form.price),discount_value:Number(form.discount_value)||0,stock:Number(form.stock)||0,visible:product?.visible??true,updated_at:new Date().toISOString()};if(!product)data.created_at=new Date().toISOString();const rows=product?.id?await sb("PATCH","products",data,`id=eq.${product.id}`):await sb("POST","products",data);onSave(Array.isArray(rows)?rows[0]:(rows||{...data,id:uid()}));}catch(e){alert("Error: "+e.message);}
    setSaving(false);
  };
  const inp={width:"100%",padding:"10px 13px",borderRadius:10,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:G.text,background:G.surfaceStrong};
  const ta={...inp,minHeight:60,resize:"vertical"};
  const lbl={fontSize:11,fontWeight:700,color:G.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <button onClick={onCancel} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,color:G.text,fontFamily:"Georgia,serif"}}>{product?t.editProduct:t.addProduct}</span>
    </div>
    <div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
      <div style={{...glassCard(G,{padding:16})}}>
        <label style={lbl}>{t.productName}</label><input style={inp} value={form.name} onChange={e=>set("name",e.target.value)} placeholder="e.g. Argan Oil Hair Serum"/>
        <label style={lbl}>{t.productNameMM}</label><input style={inp} value={form.name_mm||""} onChange={e=>set("name_mm",e.target.value)} placeholder="optional"/>
        <label style={lbl}>{t.category}</label>
        {cats.length===0?<div style={{...glassStyle(G,{borderRadius:8,padding:"8px 12px",background:G.dangerBg}),fontSize:12,color:G.danger,marginBottom:12}}>⚠️ Add categories first</div>
          :<select style={{...inp,background:G.surfaceStrong}} value={form.category} onChange={e=>set("category",e.target.value)}>{cats.map(c=><option key={c.id} value={c.name}>{c.name}</option>)}</select>}
        <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,alignItems:"end",marginBottom:0}}>
          <div><label style={lbl}>{t.emoji}</label><input style={{...inp,marginBottom:0}} value={form.emoji||""} onChange={e=>set("emoji",e.target.value)}/></div>
          <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",paddingBottom:1,whiteSpace:"nowrap",color:G.text,fontSize:13}}><input type="checkbox" checked={form.featured||false} onChange={e=>set("featured",e.target.checked)} style={{width:15,height:15,accentColor:G.primary}}/>⭐ Featured</label>
        </div>
      </div>
      <div style={{...glassCard(G,{padding:16})}}>
        <label style={lbl}>📝 {t.description}</label><textarea style={ta} value={form.description||""} onChange={e=>set("description",e.target.value)}/>
        <label style={lbl}>👤 {t.suitableFor}</label><textarea style={{...ta,minHeight:48}} value={form.suitable_for||""} onChange={e=>set("suitable_for",e.target.value)}/>
        <label style={lbl}>✅ {t.benefits}</label><textarea style={{...ta,minHeight:48}} value={form.benefits||""} onChange={e=>set("benefits",e.target.value)}/>
        <label style={lbl}>📋 {t.usage}</label><textarea style={{...ta,minHeight:48}} value={form.usage_info||""} onChange={e=>set("usage_info",e.target.value)}/>
        <label style={lbl}>⚠️ {t.warning}</label><textarea style={{...ta,minHeight:44,marginBottom:0}} value={form.warning||""} onChange={e=>set("warning",e.target.value)}/>
      </div>
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>{t.price}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.price} onChange={e=>set("price",e.target.value)}/></div>
          <div><label style={lbl}>{t.discountType}</label><select style={{...inp,marginBottom:0,background:G.surfaceStrong}} value={form.discount_type} onChange={e=>set("discount_type",e.target.value)}><option value="percent">{t.discountPct}</option><option value="fixed">{t.discountFixed}</option></select></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={lbl}>{t.discountVal}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.discount_value} onChange={e=>set("discount_value",e.target.value)}/></div>
          <div><label style={lbl}>{t.stockQty}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.stock} onChange={e=>set("stock",e.target.value)}/></div>
        </div>
        <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",userSelect:"none",marginBottom:12,fontSize:13,color:G.text}}><input type="checkbox" checked={form.preorder||false} onChange={e=>set("preorder",e.target.checked)} style={{width:15,height:15,accentColor:G.primary}}/>{t.enablePreorder}</label>
        <div style={{fontSize:11,fontWeight:700,color:G.textSub,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>{t.bulkTiers}</div>
        {(Array.isArray(form.bulk_discounts)?form.bulk_discounts:[]).map((b,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
          <input type="number" min="2" value={b.min_qty} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,min_qty:Number(e.target.value)}:x))} style={{width:52,padding:"6px 8px",borderRadius:8,...glassStyle(G,{}),fontSize:12,outline:"none",textAlign:"center",color:G.text,background:G.surfaceStrong}}/>
          <span style={{fontSize:11,color:G.muted}}>ခု+</span>
          <input type="number" min="1" max="99" value={b.discount_percent} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,discount_percent:Number(e.target.value)}:x))} style={{width:52,padding:"6px 8px",borderRadius:8,...glassStyle(G,{}),fontSize:12,outline:"none",textAlign:"center",color:G.text,background:G.surfaceStrong}}/>
          <span style={{fontSize:11,color:G.muted}}>% off</span>
          <button onClick={()=>set("bulk_discounts",form.bulk_discounts.filter((_,j)=>j!==i))} style={{background:"none",border:"none",cursor:"pointer",color:G.danger,fontSize:15,marginLeft:"auto"}}>✕</button>
        </div>)}
        <button onClick={()=>set("bulk_discounts",[...(form.bulk_discounts||[]),{min_qty:2,discount_percent:10}])} style={{padding:"6px 12px",borderRadius:8,...glassStyle(G,{border:`1.5px dashed ${G.border}`}),cursor:"pointer",fontSize:12,color:G.primary,fontWeight:600,background:"transparent"}}>{t.addTier}</button>
      </div>
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{fontSize:11,color:G.muted,marginBottom:10,fontStyle:"italic"}}>💡 {t.gdriveTip}</div>
        <div style={{fontSize:11,fontWeight:700,color:G.textSub,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>🖼️ {t.images}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
          {(Array.isArray(form.images)?form.images:[]).map((img,i)=><div key={i} style={{position:"relative"}}><img src={img} style={{width:54,height:54,objectFit:"cover",borderRadius:10,border:`1px solid ${G.border}`,display:"block"}} onError={e=>e.target.style.opacity="0.3"}/><button onClick={()=>set("images",(form.images||[]).filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:17,height:17,borderRadius:9,border:"none",background:G.danger,color:"#fff",cursor:"pointer",fontSize:9,lineHeight:"17px",textAlign:"center",padding:0}}>✕</button></div>)}
          <button onClick={()=>iRef.current?.click()} style={{width:54,height:54,borderRadius:10,border:`2px dashed ${G.border}`,background:"rgba(255,255,255,0.1)",cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:G.muted}}>+</button>
        </div>
        <input ref={iRef} type="file" accept="image/*" multiple onChange={upImg} style={{display:"none"}}/>
        <div style={{display:"flex",gap:8,marginBottom:12}}><input style={{flex:1,padding:"9px 11px",borderRadius:9,...glassStyle(G,{}),fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,color:G.text,background:G.surfaceStrong}} value={urlIn} onChange={e=>setUrlIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addUrl()} placeholder="imgbb.com URL"/><button onClick={addUrl} style={{padding:"9px 12px",borderRadius:9,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,flexShrink:0}}>Add</button></div>
        <div style={{fontSize:11,fontWeight:700,color:G.textSub,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>🎬 {t.videoUrl}</div>
        <div style={{display:"flex",gap:8}}><input style={{flex:1,padding:"9px 11px",borderRadius:9,...glassStyle(G,{}),fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,color:G.text,background:G.surfaceStrong}} value={form.video_url||""} onChange={e=>set("video_url",e.target.value)} placeholder="Video URL (mp4...)"/><button onClick={()=>vRef.current?.click()} style={{padding:"9px 11px",borderRadius:9,...glassStyle(G,{border:`1.5px solid ${G.border}`}),color:G.text,cursor:"pointer",fontSize:11,flexShrink:0}}>Upload</button></div>
        <input ref={vRef} type="file" accept="video/*" onChange={upVid} style={{display:"none"}}/>
      </div>
      <div style={{display:"flex",gap:10}}><button onClick={handleSave} disabled={saving} style={{flex:1,padding:14,borderRadius:12,border:"none",background:saving?"rgba(0,0,0,0.2)":G.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:saving?undefined:`0 4px 14px ${G.primary}50`}}>{saving?"⏳ Saving...":(product?t.save:t.addProduct)}</button><button onClick={onCancel} style={{flex:1,padding:14,borderRadius:12,border:`1.5px solid ${G.border}`,background:"transparent",color:G.muted,cursor:"pointer",fontSize:14}}>{t.cancel}</button></div>
    </div>
  </div>);
}

// ORDER MANAGER (Admin)
function OrderMgr({G,t,products,setProducts}){
  const[orders,setOrders]=useState([]);const[sel,setSel]=useState(null);const[dep,setDep]=useState("");const[noteVal,setNoteVal]=useState("");const[busy,setBusy]=useState(false);const[filter,setFilter]=useState("all");
  const load=useCallback(async()=>{try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){console.error(e);}},[]);
  useEffect(()=>{load();},[load]);
  const updateOrder=async(id,fields)=>{await sb("PATCH","orders",{...fields,updated_at:new Date().toISOString()},`id=eq.${id}`);setOrders(ords=>ords.map(o=>o.id===id?{...o,...fields}:o));if(sel?.id===id)setSel(o=>({...o,...fields}));};
  const confirmAndDeduct=async()=>{
    if(!sel||sel.stock_deducted)return;if(!window.confirm("Order confirm ပြီး stock နှုတ်မည်?"))return;
    setBusy(true);try{const updatedProds=[...products];for(const item of(sel.items||[])){if(item.is_preorder)continue;const pidx=updatedProds.findIndex(p=>p.id===item.id);if(pidx===-1)continue;const p=updatedProds[pidx];const newStock=Math.max(0,(p.stock||0)-(item.qty||0));await sb("PATCH","products",{stock:newStock,updated_at:new Date().toISOString()},`id=eq.${p.id}`);updatedProds[pidx]={...p,stock:newStock};NS.add(`${t.stockDeducted}: ${p.name_mm||p.name} -${item.qty}`);if(newStock===0)NS.add(`⚠️ ${p.name_mm||p.name} — ${t.tempOOS}`);}setProducts(updatedProds);await updateOrder(sel.id,{status:"confirmed",stock_deducted:true});}catch(e){alert("Error: "+e.message);}setBusy(false);
  };
  const statuses=["pending","confirmed","deposited","completed","cancelled"];
  const filtered=filter==="all"?orders:orders.filter(o=>o.status===filter);
  const gInp={width:"100%",padding:"10px 12px",borderRadius:10,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10,fontFamily:"inherit",color:G.text,background:G.surfaceStrong};
  if(sel){const dp=Number(sel.deposit_paid||0);const bal=Math.max(0,sel.total-dp);return(<div style={{paddingBottom:80}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <button onClick={()=>{setSel(null);load();}} style={{background:"none",border:"none",cursor:"pointer",color:G.text,fontSize:22,padding:0}}>←</button>
      <span style={{fontWeight:700,fontSize:15,color:G.text,fontFamily:"Georgia,serif"}}>Order Detail</span>
    </div>
    <div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}><div><div style={{fontSize:15,fontWeight:800,color:G.text,fontFamily:"Georgia,serif"}}>{sel.order_number}</div><div style={{fontSize:12,color:G.muted}}>{fdate(sel.created_at)}</div></div><SBadge status={sel.status} t={t} G={G}/></div>
        <div style={{...glassStyle(G,{borderRadius:10,padding:"10px 12px",marginBottom:10,fontSize:12,background:G.surface2})}}>
          <div style={{color:G.text}}><b>Name:</b> {sel.customer_name} | <b>Phone:</b> {sel.customer_phone}</div>
          {sel.customer_address&&<div style={{color:G.textSub,marginTop:2}}><b>Address:</b> {sel.customer_address}</div>}
          {sel.contact_method&&<div style={{color:G.muted,marginTop:2}}><b>Via:</b> {sel.contact_method}</div>}
        </div>
        {(sel.items||[]).some(i=>i.is_preorder)&&<div style={{...glassStyle(G,{borderRadius:8,padding:"8px 12px",marginBottom:10,background:G.warningBg}),fontSize:11,color:G.warning,fontWeight:700}}>⚠️ Pre-order items ပါဝင်</div>}
        {(sel.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${G.border}`,fontSize:13}}><div><span style={{color:G.text,fontWeight:600}}>{item.name_mm||item.name}</span>{item.is_preorder&&<span style={{fontSize:9,background:G.warningBg,color:G.warning,padding:"1px 5px",borderRadius:5,fontWeight:700,marginLeft:4}}>PRE</span>}<span style={{color:G.muted}}> × {item.qty}</span></div><span style={{fontWeight:700,color:G.text}}>{fmt(item.total)}</span></div>)}
        <div style={{paddingTop:12,borderTop:`1px solid ${G.border}`,marginTop:4}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:G.text}}>{t.total}</span><span style={{color:G.primary}}>{fmt(sel.total)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{color:G.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:G.success,fontWeight:700}}>{fmt(dp)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:G.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:G.danger,fontWeight:700}}>{fmt(bal)}</span></div>
        </div>
        {sel.customer_note&&<div style={{...glassStyle(G,{borderRadius:8,padding:"8px 12px",marginTop:10,background:G.surface2}),fontSize:12,color:G.textSub}}><b>Note:</b> {sel.customer_note}</div>}
      </div>
      {!sel.stock_deducted?<button onClick={confirmAndDeduct} disabled={busy} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:busy?"rgba(0,0,0,0.2)":G.success,color:"#fff",cursor:busy?"wait":"pointer",fontSize:14,fontWeight:700}}>{busy?"⏳ Processing...":t.confirmOrder}</button>:<div style={{...glassStyle(G,{borderRadius:10,padding:"10px 14px",background:G.successBg}),fontSize:13,color:G.success,fontWeight:600,textAlign:"center"}}>✅ Stock already deducted</div>}
      <div style={{...glassCard(G,{padding:14})}}><div style={{fontSize:12,fontWeight:700,color:G.textSub,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.updateStatus}</div><div style={{display:"flex",flexWrap:"wrap",gap:8}}>{statuses.map(s=><button key={s} onClick={()=>updateOrder(sel.id,{status:s})} style={{padding:"5px 10px",borderRadius:20,...glassStyle(G,{border:`1.5px solid ${sel.status===s?G.primary:G.border}`}),background:sel.status===s?G.primary:"transparent",color:sel.status===s?"#fff":G.text,cursor:"pointer"}}><SBadge status={s} t={t} G={G}/></button>)}</div></div>
      <div style={{...glassCard(G,{padding:14})}}><div style={{fontSize:12,fontWeight:700,color:G.textSub,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.updateDeposit}</div><div style={{display:"flex",gap:8}}><input type="number" min="0" value={dep} onChange={e=>setDep(e.target.value)} placeholder="0 Ks" style={{flex:1,padding:"10px 12px",borderRadius:10,...glassStyle(G,{}),fontSize:14,outline:"none",color:G.text,fontFamily:"inherit",minWidth:0,background:G.surfaceStrong}}/><button onClick={async()=>{const d=Number(dep);await updateOrder(sel.id,{deposit_paid:d,balance_due:Math.max(0,sel.total-d)});setDep("");}} style={{padding:"10px 16px",borderRadius:10,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0,boxShadow:`0 4px 12px ${G.primary}50`}}>OK</button></div></div>
      <div style={{...glassCard(G,{padding:14})}}><div style={{fontSize:12,fontWeight:700,color:G.textSub,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.adminNote}</div><textarea value={noteVal||sel.admin_note||""} onChange={e=>setNoteVal(e.target.value)} style={{...gInp,minHeight:72,resize:"vertical"}}/><button onClick={()=>updateOrder(sel.id,{admin_note:noteVal||sel.admin_note||""})} style={{width:"100%",padding:10,borderRadius:10,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700,boxShadow:`0 3px 10px ${G.primary}50`}}>{t.save}</button></div>
    </div>
  </div>);}

  return(<div style={{paddingBottom:80}}>
    <div style={{display:"flex",gap:6,padding:"10px 12px",overflowX:"auto",scrollbarWidth:"none",...glassStyle(G,{borderBottom:`1px solid ${G.border}`})}}>
      {["all",...statuses].map(s=><button key={s} onClick={()=>setFilter(s)} style={{flexShrink:0,padding:"6px 12px",borderRadius:20,...glassStyle(G,{border:`1.5px solid ${filter===s?G.primary:G.border}`}),background:filter===s?G.primary:"transparent",color:filter===s?"#fff":G.text,cursor:"pointer",fontSize:12,fontWeight:600}}>{s==="all"?t.all:<SBadge status={s} t={t} G={G}/>}</button>)}
    </div>
    <div style={{padding:"10px 14px",display:"flex",flexDirection:"column",gap:10}}>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:G.muted}}><div style={{fontSize:40,marginBottom:8}}>📦</div><div>{t.noData}</div></div>
        :filtered.map(o=><div key={o.id} onClick={()=>{setSel(o);setNoteVal(o.admin_note||"");setDep("");}} style={{...glassCard(G,{padding:"14px",cursor:"pointer",transition:"all 0.15s"})}} onMouseEnter={e=>e.currentTarget.style.transform="translateX(3px)"} onMouseLeave={e=>e.currentTarget.style.transform="none"}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}><div><div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif"}}>{o.order_number}</div><div style={{fontSize:11,color:G.muted}}>{fdate(o.created_at)} · {o.customer_name} · {o.contact_method}</div></div><SBadge status={o.status} t={t} G={G}/></div>
          {(o.items||[]).some(i=>i.is_preorder)&&<div style={{fontSize:10,color:G.warning,fontWeight:700,marginBottom:4}}>⚠️ PREORDER</div>}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:12,color:G.muted}}>{(o.items||[]).length} items{o.stock_deducted?" · ✅":""}</div><div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:800,color:G.primary}}>{fmt(o.total)}</div>{(o.deposit_paid||0)>0&&<div style={{fontSize:10,color:G.success}}>Dep: {fmt(o.deposit_paid)}</div>}</div></div>
        </div>)}
    </div>
  </div>);
}

// STOCK MANAGER (Admin)
function StockMgr({G,t,products,setProducts}){
  const[adding,setAdding]=useState({});const[vals,setVals]=useState({});
  const doAdd=async(p)=>{const n=Number(vals[p.id]||0);if(!n||n<=0){alert("Valid qty ထည့်ပါ");return;}const ns=(p.stock||0)+n;try{await sb("PATCH","products",{stock:ns,preorder:false,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,stock:ns,preorder:false}:x));NS.add(`${t.stockAdded}: ${p.name_mm||p.name} +${n}`);setAdding(v=>({...v,[p.id]:false}));setVals(v=>({...v,[p.id]:""}));}catch(e){alert(e.message);}};
  const togglePO=async(p)=>{const np=!p.preorder;try{await sb("PATCH","products",{preorder:np,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,preorder:np}:x));NS.add(`${p.name_mm||p.name} — ${np?"Pre-order ON":"Pre-order OFF"}`);}catch(e){alert(e.message);}};
  const oos_=products.filter(p=>isOOS(p));const low=products.filter(p=>(p.stock||0)<=5&&!isOOS(p));
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    {oos_.length>0&&<div style={{...glassStyle(G,{borderRadius:12,padding:"12px 14px",marginBottom:10,background:G.dangerBg,border:`1px solid ${G.danger}30`})}}>
      <div style={{fontSize:13,fontWeight:700,color:G.danger,marginBottom:4}}>⚠️ Out of Stock ({oos_.length})</div>
      {oos_.map(p=><div key={p.id} style={{fontSize:12,color:G.danger}}>• {p.name_mm||p.name}{p.preorder?" [Pre-order]":""}</div>)}
    </div>}
    {low.length>0&&<div style={{...glassStyle(G,{borderRadius:12,padding:"12px 14px",marginBottom:10,background:G.warningBg,border:`1px solid ${G.warning}30`})}}>
      <div style={{fontSize:13,fontWeight:700,color:G.warning,marginBottom:4}}>🔶 Low Stock (≤5)</div>
      {low.map(p=><div key={p.id} style={{fontSize:12,color:G.warning}}>• {p.name_mm||p.name}: {p.stock}</div>)}
    </div>}
    <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:10}}>📦 All Stock</div>
    {products.map(p=>{const oos__=isOOS(p);return(<div key={p.id} style={{...glassCard(G,{padding:"14px",marginBottom:10})}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><PImg p={p} size={42} r={10} G={G}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:G.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontFamily:"Georgia,serif"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,fontWeight:700,marginTop:1,color:oos__?G.danger:(p.stock||0)<=5?G.warning:G.success}}>{oos__?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`}</div></div></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {adding[p.id]?<div style={{display:"flex",gap:6,flex:1}}><input type="number" min="1" value={vals[p.id]||""} onChange={e=>setVals(v=>({...v,[p.id]:e.target.value}))} placeholder="qty" style={{width:70,padding:"7px 10px",borderRadius:8,...glassStyle(G,{}),fontSize:13,outline:"none",color:G.text,fontFamily:"inherit",background:G.surfaceStrong}}/><button onClick={()=>doAdd(p)} style={{padding:"7px 14px",borderRadius:8,border:"none",background:G.success,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add</button><button onClick={()=>setAdding(v=>({...v,[p.id]:false}))} style={{padding:"7px 10px",borderRadius:8,...glassStyle(G,{border:`1px solid ${G.border}`}),background:"transparent",color:G.muted,cursor:"pointer",fontSize:12}}>✕</button></div>
          :<button onClick={()=>setAdding(v=>({...v,[p.id]:true}))} style={{padding:"7px 14px",borderRadius:8,...glassStyle(G,{border:`1.5px solid ${G.success}`}),background:`${G.success}20`,color:G.success,cursor:"pointer",fontSize:12,fontWeight:700}}>➕ {t.addStock}</button>}
        <button onClick={()=>togglePO(p)} style={{padding:"7px 12px",borderRadius:8,...glassStyle(G,{border:`1.5px solid ${p.preorder?G.warning:G.border}`}),background:p.preorder?G.warningBg:"transparent",color:p.preorder?G.warning:G.muted,cursor:"pointer",fontSize:11,fontWeight:600}}>{p.preorder?"✅ Pre-order ON":`📋 ${t.switchPreorder}`}</button>
      </div>
    </div>);})}
  </div>);
}

// REPORTS (Admin)
function AdminReports({G,t}){
  const[orders,setOrders]=useState([]);const[period,setPeriod]=useState("week");const[loading,setLoading]=useState(true);
  useEffect(()=>{(async()=>{setLoading(true);try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){}setLoading(false);})();},[]);
  const filtered=useMemo(()=>{const now=new Date();return orders.filter(o=>{const d=new Date(o.created_at);if(period==="today")return d.toDateString()===now.toDateString();if(period==="week"){const w=new Date(now);w.setDate(w.getDate()-7);return d>=w;}if(period==="month"){const m=new Date(now);m.setDate(m.getDate()-30);return d>=m;}return true;});},[orders,period]);
  const active=filtered.filter(o=>o.status!=="cancelled");const rev=active.reduce((s,o)=>s+Number(o.total),0);const dep=active.reduce((s,o)=>s+(Number(o.deposit_paid)||0),0);
  const tops=useMemo(()=>{const m={};active.forEach(o=>(o.items||[]).forEach(item=>{const k=item.name||"?";if(!m[k])m[k]={name:item.name_mm||item.name||k,qty:0,rev:0};m[k].qty+=(item.qty||0);m[k].rev+=item.total||0;}));return Object.values(m).sort((a,b)=>b.rev-a.rev).slice(0,6);},[active]);
  const exportExcel=()=>{const wb=XLSX.utils.book_new();const oRows=[["Order#","Date","Status","Customer","Phone","Address","Items","Total","Deposit","Balance","Contact","Note","Admin Note"]];filtered.forEach(o=>oRows.push([o.order_number,fdate(o.created_at),o.status,o.customer_name,o.customer_phone,o.customer_address,(o.items||[]).map(i=>`${i.name_mm||i.name}x${i.qty}${i.is_preorder?"[PRE]":""}`).join(", "),o.total,o.deposit_paid||0,Math.max(0,o.total-(o.deposit_paid||0)),o.contact_method,o.customer_note||"",o.admin_note||""]));XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(oRows),"Orders");const sRows=[["Metric","Value"],["Period",period],["Total Revenue",rev],["Deposit Received",dep],["Balance Due",rev-dep],["Total Orders",filtered.length],["Cancelled",filtered.length-active.length]];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),"Summary");if(tops.length>0){const pRows=[["Product","Qty","Revenue"],...tops.map(p=>[p.name,p.qty,p.rev])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(pRows),"Top Products");}XLSX.writeFile(wb,`BeautyReport_${period}_${new Date().toISOString().slice(0,10)}.xlsx`);};
  const periods=[{k:"today",l:t.today},{k:"week",l:t.week},{k:"month",l:t.month},{k:"year",l:t.allTime}];
  const Stat=({emoji,label,value,color})=><div style={{...glassCard(G,{padding:14,flex:1,minWidth:80})}}><div style={{fontSize:18,marginBottom:2}}>{emoji}</div><div style={{fontSize:11,color:G.muted,fontWeight:600,textTransform:"uppercase",letterSpacing:0.3}}>{label}</div><div style={{fontSize:14,fontWeight:800,color:color||G.text,marginTop:2}}>{value}</div></div>;
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    <div style={{display:"flex",gap:6,marginBottom:14,...glassCard(G,{padding:6})}}>
      {periods.map(({k,l})=><button key={k} onClick={()=>setPeriod(k)} style={{flex:1,padding:"8px 4px",borderRadius:10,border:"none",background:period===k?G.primary:"transparent",color:period===k?"#fff":G.muted,cursor:"pointer",fontSize:12,fontWeight:period===k?700:400,transition:"all 0.15s"}}>{l}</button>)}
    </div>
    {loading?<div style={{textAlign:"center",padding:40,color:G.muted}}>Loading...</div>:<>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}><Stat emoji="💰" label={t.revenue} value={fmt(rev)} color={G.primary}/><Stat emoji="📦" label={t.totalOrders} value={filtered.length}/><Stat emoji="✅" label="Deposit" value={fmt(dep)} color={G.success}/><Stat emoji="⏳" label="Balance" value={fmt(rev-dep)} color={G.danger}/></div>
      {tops.length>0&&<div style={{...glassCard(G,{padding:16,marginBottom:12})}}>
        <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>🏆 {t.topProducts}</div>
        {tops.map((p,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:i<tops.length-1?`1px solid ${G.border}`:"none"}}><span style={{width:22,height:22,borderRadius:11,background:[G.primary,"rgba(192,192,192,0.8)","rgba(205,127,50,0.8)","rgba(160,160,160,0.5)","rgba(160,160,160,0.5)","rgba(160,160,160,0.5)"][i],display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{i+1}</span><span style={{flex:1,fontSize:12,fontWeight:600,color:G.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span><span style={{fontSize:11,color:G.muted,marginRight:6}}>{p.qty}ခု</span><span style={{fontSize:12,fontWeight:700,color:G.primary,whiteSpace:"nowrap"}}>{fmt(p.rev)}</span></div>)}
      </div>}
      <div style={{...glassCard(G,{padding:16})}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif"}}>{t.recentOrders}</div><button onClick={exportExcel} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 14px",borderRadius:10,border:"none",background:"#21A366",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700,boxShadow:"0 3px 10px rgba(33,163,102,0.4)"}}>📊 {t.exportExcel}</button></div>
        {filtered.slice(0,10).map(o=><div key={o.id} style={{padding:"9px 0",borderBottom:`1px solid ${G.border}`,fontSize:12}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}><span style={{fontWeight:700,color:G.text}}>{o.order_number} · {o.customer_name}</span><SBadge status={o.status} t={t} G={G}/></div><div style={{display:"flex",justifyContent:"space-between",color:G.muted}}><span>{fdate(o.created_at)}</span><span style={{fontWeight:700,color:G.primary}}>{fmt(o.total)}</span></div></div>)}
        {filtered.length===0&&<div style={{textAlign:"center",color:G.muted,padding:20,fontSize:13}}>{t.noData}</div>}
      </div>
    </>}
  </div>);
}


// ADMIN SETTINGS
function AdminSettings({G,t,settings,onSave,onTheme,onLang,themeName,lang,onDisconnect}){
  const[form,setForm]=useState({...settings});const[logo,setLogo]=useState(settings.logo||"");const[banner,setBanner]=useState(settings.banner||"");const[saving,setSaving]=useState(false);
  const logoRef=useRef();const bannerRef=useRef();const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const upLogo=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setLogo(b);set("logo",b);e.target.value="";};
  const upBanner=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setBanner(b);set("banner",b);e.target.value="";};
  const doSave=async()=>{setSaving(true);try{const data={...form,logo,banner};const rows=Object.entries(data).map(([key,value])=>({key,value:String(value||"")}));await sbUpsert("shop_settings",rows);onSave(data);alert("✅ Saved!");}catch(e){alert("Error: "+e.message);}setSaving(false);};
  const inp={width:"100%",padding:"10px 13px",borderRadius:10,...glassStyle(G,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:G.text,background:G.surfaceStrong};
  const lbl={fontSize:11,fontWeight:700,color:G.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    <div style={{...glassCard(G,{padding:16,marginBottom:12})}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>🎨 {t.themeLabel} & {t.langLabel}</div>
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        {Object.entries(THEMES).map(([k])=><button key={k} onClick={()=>onTheme(k)} style={{flex:1,padding:10,borderRadius:10,...glassStyle(G,{border:`2px solid ${themeName===k?G.primary:G.border}`}),background:themeName===k?G.primary:"transparent",color:themeName===k?"#fff":G.text,cursor:"pointer",fontSize:12,fontWeight:700,transition:"all 0.15s"}}>{k==="glass"?"💄 Glass":"🌙 Dark"}</button>)}
      </div>
      <div style={{display:"flex",gap:8}}>
        {["mm","en"].map(l=><button key={l} onClick={()=>onLang(l)} style={{flex:1,padding:10,borderRadius:10,...glassStyle(G,{border:`2px solid ${lang===l?G.primary:G.border}`}),background:lang===l?G.primary:"transparent",color:lang===l?"#fff":G.text,cursor:"pointer",fontSize:12,fontWeight:700,transition:"all 0.15s"}}>{l==="mm"?"🇲🇲 မြန်မာ":"🇬🇧 English"}</button>)}
      </div>
    </div>
    <div style={{...glassCard(G,{padding:16,marginBottom:12})}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>🖼️ Logo & Banner</div>
      <div style={{display:"flex",gap:14,marginBottom:12}}>
        <div style={{textAlign:"center"}}>
          <div style={{width:72,height:72,borderRadius:16,...glassStyle(G,{border:`2px dashed ${G.border}`}),display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",cursor:"pointer"}} onClick={()=>logoRef.current?.click()}>
            {logo?<img src={logo} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:28}}>💄</span>}
          </div>
          <div style={{fontSize:10,color:G.muted,marginTop:4}}>Logo</div>
        </div>
        <div style={{flex:1}}>
          <div style={{height:72,borderRadius:16,...glassStyle(G,{border:`2px dashed ${G.border}`}),display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",cursor:"pointer"}} onClick={()=>bannerRef.current?.click()}>
            {banner?<img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:12,color:G.muted}}>+ Banner Image</span>}
          </div>
        </div>
      </div>
      <input ref={logoRef} type="file" accept="image/*" onChange={upLogo} style={{display:"none"}}/>
      <input ref={bannerRef} type="file" accept="image/*" onChange={upBanner} style={{display:"none"}}/>
    </div>
    <div style={{...glassCard(G,{padding:16,marginBottom:12})}}>
      <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:12}}>🏪 Shop Info</div>
      <label style={lbl}>{t.shopName}</label><input style={inp} value={form.shop_name||""} onChange={e=>set("shop_name",e.target.value)}/>
      <label style={lbl}>{t.shopNameMM}</label><input style={inp} value={form.shop_name_mm||""} onChange={e=>set("shop_name_mm",e.target.value)}/>
      <label style={lbl}>{t.fbLink}</label><input style={inp} value={form.fb_link||""} onChange={e=>set("fb_link",e.target.value)} placeholder="https://m.me/yourpage"/>
      <label style={lbl}>{t.viberNum}</label><input style={inp} value={form.viber_num||""} onChange={e=>set("viber_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.waNum}</label><input style={inp} value={form.wa_num||""} onChange={e=>set("wa_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.phoneNum}</label><input style={inp} value={form.phone_num||""} onChange={e=>set("phone_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.adminPw}</label><input style={{...inp,marginBottom:0}} type="password" value={form.admin_pw||""} onChange={e=>set("admin_pw",e.target.value)}/>
    </div>
    <button onClick={doSave} disabled={saving} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:saving?"rgba(0,0,0,0.2)":G.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700,marginBottom:10,boxShadow:saving?undefined:`0 4px 16px ${G.primary}60`}}>{saving?"⏳ Saving...":t.saveSettings}</button>
    <button onClick={onDisconnect} style={{width:"100%",padding:12,borderRadius:12,...glassStyle(G,{border:`1.5px solid ${G.danger}`}),background:G.dangerBg,color:G.danger,cursor:"pointer",fontSize:13,fontWeight:600}}>🔌 Disconnect Supabase</button>
  </div>);
}

// ADMIN PANEL
function AdminPanel({G,t,products,cats,settings,onProdChange,onCatsChange,onSettingsChange,onTheme,onLang,themeName,lang,onBack,onDisconnect}){
  const[tab,setTab]=useState("products");const[showForm,setShowForm]=useState(false);const[editing,setEditing]=useState(null);const[newCat,setNewCat]=useState("");
  const tabs=[{k:"products",l:"📦 "+t.products},{k:"stock",l:"📊 Stock"},{k:"orders",l:"🛍️ "+t.orders},{k:"reports",l:"📈 "+t.reports},{k:"settings",l:"⚙️ "+t.settings}];
  const delProd=async p=>{if(!window.confirm(`Delete "${p.name_mm||p.name}"?`))return;try{await sb("DELETE","products",null,`id=eq.${p.id}`);onProdChange(products.filter(x=>x.id!==p.id));}catch(e){alert(e.message);}};
  const toggleVis=async p=>{try{await sb("PATCH","products",{visible:!p.visible},`id=eq.${p.id}`);onProdChange(products.map(x=>x.id===p.id?{...x,visible:!x.visible}:x));NS.add(`${p.name_mm||p.name} — ${!p.visible?"visible":"hidden"}`);}catch(e){alert(e.message);}};
  const addCat=async name=>{if(!name||cats.find(c=>c.name===name))return;try{const rows=await sb("POST","categories",{name,sort_order:cats.length+1});onCatsChange([...cats,...(Array.isArray(rows)?rows:[rows])]);}catch(e){alert(e.message);}};
  const delCat=async cat=>{if(!window.confirm(`Delete "${cat.name}"?`))return;try{await sb("DELETE","categories",null,`id=eq.${cat.id}`);onCatsChange(cats.filter(c=>c.id!==cat.id));}catch(e){alert(e.message);}};

  if(showForm)return<ProductForm G={G} t={t} product={editing} cats={cats} onSave={saved=>{if(editing)onProdChange(products.map(x=>x.id===saved.id?saved:x));else onProdChange([saved,...products]);NS.add(editing?`${t.notifUpdated}: ${saved.name_mm||saved.name}`:`✨ New: ${saved.name_mm||saved.name}`);setShowForm(false);setEditing(null);}} onCancel={()=>{setShowForm(false);setEditing(null);}}/>;

  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...glassStyle(G,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100})}}>
      <span style={{fontWeight:800,fontSize:16,color:G.text,fontFamily:"Georgia,serif",flex:1}}>💄 {t.adminPanel}</span>
      <button onClick={onBack} style={{...glassStyle(G,{borderRadius:20,padding:"6px 14px",border:`1px solid ${G.border}`}),color:G.text,cursor:"pointer",fontSize:12,fontWeight:700,background:"rgba(255,255,255,0.2)"}}>{t.logout} ←</button>
    </div>
    <div style={{display:"flex",...glassStyle(G,{borderBottom:`1px solid ${G.border}`}),overflowX:"auto",scrollbarWidth:"none"}}>
      {tabs.map(tab_=><button key={tab_.k} onClick={()=>setTab(tab_.k)} style={{flex:"none",padding:"12px 12px",border:"none",background:"none",cursor:"pointer",fontSize:11,fontWeight:tab===tab_.k?700:500,color:tab===tab_.k?G.primary:G.muted,borderBottom:`2.5px solid ${tab===tab_.k?G.primary:"transparent"}`,whiteSpace:"nowrap",transition:"all 0.15s"}}>{tab_.l}</button>)}
    </div>
    {tab==="products"&&<div style={{padding:"12px 14px",paddingBottom:80}}>
      <button onClick={()=>{setEditing(null);setShowForm(true);}} style={{width:"100%",padding:13,borderRadius:14,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:14,boxShadow:`0 4px 16px ${G.primary}60`}}>✨ {t.addProduct}</button>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {products.map(p=><div key={p.id} style={{...glassCard(G,{padding:"12px",display:"flex",gap:10,alignItems:"center",opacity:p.visible?1:0.5})}}>
          <PImg p={p} size={50} r={10} G={G}/>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div>
            <div style={{fontSize:11,color:G.muted,marginTop:1}}>{p.category} · {fmt(calcUnit(p))}</div>
            <div style={{fontSize:10,marginTop:1,color:isOOS(p)?G.danger:(p.stock||0)<=5?G.warning:G.success,fontWeight:600}}>{isOOS(p)?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`} · {p.visible?"🟢":"⚫"}</div>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
            <button onClick={()=>{setEditing(p);setShowForm(true);}} style={{padding:"5px 9px",borderRadius:8,...glassStyle(G,{border:`1px solid ${G.border}`}),cursor:"pointer",fontSize:12,color:G.text,background:"rgba(255,255,255,0.2)"}}>✏️</button>
            <button onClick={()=>toggleVis(p)} style={{padding:"5px 9px",borderRadius:8,...glassStyle(G,{border:`1px solid ${G.border}`}),cursor:"pointer",fontSize:12,color:G.text,background:"rgba(255,255,255,0.2)"}}>{p.visible?"🙈":"👁️"}</button>
            <button onClick={()=>delProd(p)} style={{padding:"5px 9px",borderRadius:8,...glassStyle(G,{border:`1px solid ${G.danger}50`,background:G.dangerBg}),cursor:"pointer",fontSize:12,color:G.danger}}>🗑️</button>
          </div>
        </div>)}
        {products.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:G.muted,fontSize:14}}>{t.noProducts}</div>}
      </div>
      <div style={{...glassCard(G,{padding:14,marginTop:14})}}>
        <div style={{fontSize:13,fontWeight:700,color:G.text,fontFamily:"Georgia,serif",marginBottom:10}}>🏷️ Categories</div>
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          <input value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>e.key==="Enter"&&newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} placeholder="New category"
            style={{flex:1,padding:"9px 12px",borderRadius:10,...glassStyle(G,{}),fontSize:13,outline:"none",fontFamily:"inherit",color:G.text,minWidth:0,boxSizing:"border-box",background:G.surfaceStrong}}/>
          <button onClick={()=>newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} style={{padding:"9px 16px",borderRadius:10,border:"none",background:G.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700,flexShrink:0}}>+</button>
        </div>
        {cats.map(cat=><div key={cat.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 12px",...glassStyle(G,{borderRadius:10,border:`1px solid ${G.border}`}),marginBottom:6,background:"rgba(255,255,255,0.1)"}}>
          <span style={{fontSize:13,fontWeight:600,color:G.text}}>{cat.name}</span>
          <button onClick={()=>delCat(cat)} style={{background:"none",...glassStyle(G,{border:`1px solid ${G.danger}50`,borderRadius:8,background:G.dangerBg}),cursor:"pointer",color:G.danger,fontSize:11,padding:"3px 9px"}}>🗑️</button>
        </div>)}
      </div>
    </div>}
    {tab==="stock"&&<StockMgr G={G} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="orders"&&<OrderMgr G={G} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="reports"&&<AdminReports G={G} t={t}/>}
    {tab==="settings"&&<AdminSettings G={G} t={t} settings={settings} onSave={s=>{onSettingsChange(s);}} onTheme={onTheme} onLang={onLang} themeName={themeName} lang={lang} onDisconnect={onDisconnect}/>}
  </div>);
}

// MAIN APP
export default function App(){
  const[cfg,setCfg]=useState(null);const[ready,setReady]=useState(false);
  const[products,setProducts]=useState([]);const[cats,setCats]=useState([]);const[settings,setSettings]=useState({});
  const[cart,setCart]=useState([]);const[tab,setTab]=useState("home");const[catFilter,setCatFilter]=useState("all");
  const[searchQ,setSearchQ]=useState("");const[showDrop,setShowDrop]=useState(false);
  const[selProd,setSelProd]=useState(null);const[page,setPage]=useState("home");const[adminIn,setAdminIn]=useState(false);
  const[themeName,setThemeName]=useState(LS.get("theme")||"glass");
  const[lang,setLang]=useState(LS.get("lang")||"mm");
  const G_=THEMES[themeName]||THEMES.glass;
  const t=TR[lang]||TR.mm;

  useEffect(()=>{(async()=>{let initCfg=null;if(SB_URL&&SB_KEY){initCfg={url:SB_URL,key:SB_KEY};}else{initCfg=LS.get("sb_cfg");}if(initCfg){setCfg(initCfg);}else{setReady(true);}})();},[]);
  useEffect(()=>{if(!cfg)return;(async()=>{setReady(false);try{const[prods,catsData,setsData]=await Promise.all([sb("GET","products",null,"order=created_at.desc&visible=eq.true"),sb("GET","categories",null,"order=sort_order.asc"),sb("GET","shop_settings")]);setProducts(prods||[]);setCats(catsData||[]);const smap={};(setsData||[]).forEach(s=>{smap[s.key]=s.value;});setSettings(smap);if(smap.theme)setThemeName(smap.theme);if(smap.language)setLang(smap.language);}catch(e){console.error("Load error:",e);}setReady(true);})();},[cfg]);

  const loadAllProducts=useCallback(async()=>{if(!cfg)return;try{setProducts(await sb("GET","products",null,"order=created_at.desc")||[]);}catch(e){}},[cfg]);
  useEffect(()=>{if(adminIn)loadAllProducts();},[adminIn]);

  const searchRes=useMemo(()=>{if(!searchQ.trim())return[];const q=searchQ.toLowerCase();return products.filter(p=>p.visible!==false&&((p.name||"").toLowerCase().includes(q)||(p.name_mm||"").includes(searchQ)||(p.description||"").includes(searchQ)||(p.category||"").includes(searchQ))).slice(0,7);},[searchQ,products]);
  const visProds=useMemo(()=>products.filter(p=>p.visible!==false&&(catFilter==="all"||p.category===catFilter)),[products,catFilter]);
  const cartCount=cart.reduce((s,i)=>s+i.qty,0);const cartTotal=cart.reduce((s,i)=>s+calcUnit(i.p,i.qty)*i.qty,0);
  const addToCart=(p,qty=1)=>setCart(prev=>{const ex=prev.find(i=>i.p.id===p.id);return ex?prev.map(i=>i.p.id===p.id?{...i,qty:i.qty+qty}:i):[...prev,{p,qty}];});
  const removeItem=id=>setCart(c=>c.filter(i=>i.p.id!==id));
  const updateQty=(id,qty)=>{if(qty<=0)removeItem(id);else setCart(c=>c.map(i=>i.p.id===id?{...i,qty}:i));};
  const openProd=p=>{setSelProd(p);setPage("product");setShowDrop(false);setSearchQ("");};
  const changeTheme=k=>{setThemeName(k);LS.set("theme",k);};
  const changeLang=l=>{setLang(l);LS.set("lang",l);};
  const onSettingsChange=s=>{setSettings(s);if(s.theme)changeTheme(s.theme);if(s.language)changeLang(s.language);};

  const bgStyle={background:G_.bg,backgroundAttachment:"fixed",minHeight:"100vh",fontFamily:"'Segoe UI','Myanmar Text',Helvetica,sans-serif"};

  if(!ready&&!cfg)return<SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;
  if(!cfg)return<SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;
  if(!ready)return(<div style={{...bgStyle,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
    <div style={{fontSize:56,animation:"pulse 1.5s infinite"}}>💄</div>
    <div style={{color:G_.text,fontSize:15,fontWeight:700,fontFamily:"Georgia,serif"}}>Beauty Store</div>
    <div style={{color:G_.muted,fontSize:13}}>Loading...</div>
  </div>);

  if(tab==="admin"){
    if(!adminIn)return(<div style={{...bgStyle,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{...glassCard(G_,{padding:"40px 28px",maxWidth:360,width:"100%"})}}>
        <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:52,marginBottom:8}}>💄</div><div style={{fontSize:20,fontWeight:800,color:G_.text,fontFamily:"Georgia,serif"}}>{t.adminPanel}</div></div>
        <AdminLoginBox G={G_} t={t} adminPw={settings.admin_pw} onSuccess={()=>setAdminIn(true)} onBack={()=>setTab("home")}/>
      </div>
    </div>);
    return(<div style={bgStyle}><AdminPanel G={G_} t={t} products={products} cats={cats} settings={settings} onProdChange={setProducts} onCatsChange={setCats} onSettingsChange={onSettingsChange} onTheme={changeTheme} onLang={changeLang} themeName={themeName} lang={lang} onBack={()=>{setTab("home");setAdminIn(false);setProducts(ps=>ps.filter(p=>p.visible!==false));}} onDisconnect={()=>{LS.set("sb_cfg",null);setCfg(null);setAdminIn(false);setTab("home");}}/></div>);
  }

  if(page==="product"&&selProd)return(<div style={bgStyle}><ProductDetail G={G_} t={t} p={selProd} onBack={()=>setPage("home")} addToCart={addToCart} onBuyNow={()=>setPage("checkout")}/></div>);
  if(page==="checkout")return(<div style={bgStyle}><CheckoutPage G={G_} t={t} cart={cart} total={cartTotal} settings={settings} onPlaced={()=>{setCart([]);setPage("home");}} onBack={()=>setPage("home")}/></div>);

  const shopMM=settings.shop_name_mm||settings.shop_name||"Beauty Store";
  return(<div style={bgStyle}>
    <TopBar G={G_} t={t} shopMM={shopMM} logo={settings.logo||""} cartCount={cartCount} searchQ={searchQ} setSearchQ={setSearchQ} showDrop={showDrop} setShowDrop={setShowDrop} dropRes={searchRes} onHit={openProd} onCart={()=>setTab("cart")} onAdmin={()=>setTab("admin")} onLogo={()=>{setTab("home");setPage("home");}} onLangToggle={()=>changeLang(lang==="mm"?"en":"mm")} onThemeToggle={()=>changeTheme(themeName==="glass"?"dark":"glass")} themeName={themeName}/>
    <div style={{paddingBottom:80}}>
      {tab==="home"&&<HomePage G={G_} t={t} products={visProds} cats={cats} catFilter={catFilter} setCatFilter={setCatFilter} onOpen={openProd} onAdd={addToCart} banner={settings.banner||""} shopMM={shopMM}/>}
      {tab==="cart"&&<CartPage G={G_} t={t} cart={cart} updateQty={updateQty} removeItem={removeItem} total={cartTotal} onCheckout={()=>setPage("checkout")} onBack={()=>setTab("home")}/>}
      {tab==="track"&&<TrackOrder G={G_} t={t}/>}
      {tab==="cats"&&<div style={{padding:"14px 14px"}}>
        <div style={{fontSize:16,fontWeight:800,color:G_.text,fontFamily:"Georgia,serif",marginBottom:14}}>✨ {t.cats}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
          {cats.map(c=>{const icons={"Makeup":"💄","Skin Care":"✨","Hair Care":"💆","Fragrance":"🌸","Lip Care":"💋","Eye Care":"👁️","Others":"🛍️","ဆံပင်":"💆","မျက်နှာ":"✨","နှုတ်ခမ်း":"💋","မျက်လုံး":"👁️"};return(<div key={c.id} onClick={()=>{setCatFilter(c.name);setTab("home");}} style={{...glassCard(G_,{padding:"18px 10px",textAlign:"center",cursor:"pointer",transition:"all 0.2s"})}}>
            <div style={{fontSize:30,marginBottom:6}}>{icons[c.name]||"💄"}</div>
            <div style={{fontSize:11,fontWeight:600,color:G_.text,lineHeight:1.3}}>{c.name}</div>
          </div>);})}
        </div>
      </div>}
    </div>
    <BottomNav G={G_} t={t} tab={tab} setTab={t2=>{setTab(t2);setPage("home");}} cartCount={cartCount}/>
  </div>);
}
