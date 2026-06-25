import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// CONFIG — set in Vercel: VITE_SUPABASE_URL + VITE_SUPABASE_ANON_KEY
const SB_URL = import.meta.env?.VITE_SUPABASE_URL || "";
const SB_KEY = import.meta.env?.VITE_SUPABASE_ANON_KEY || "";

// SUPABASE REST
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

// HELPERS
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

// THEMES
const THEMES = {
  light:{ primary:"#EE4D2D", primaryD:"#D94528", accent:"#FFB347", bg:"#F5F5F5", surface:"#FFFFFF", surface2:"#F9F9F9", border:"rgba(0,0,0,0.09)", text:"#212121", textSub:"#555", muted:"#999", success:"#27AE60", danger:"#E74C3C", warning:"#F39C12", shadow:"0 2px 12px rgba(0,0,0,0.07)", shadowMd:"0 4px 24px rgba(0,0,0,0.12)", navBg:"#FFFFFF", tag:"#FFF0ED", tagText:"#C0392B" },
  dark: { primary:"#FF6140", primaryD:"#EE4D2D", accent:"#FFB347", bg:"#121212", surface:"#1E1E1E", surface2:"#2A2A2A", border:"rgba(255,255,255,0.1)", text:"#EFEFEF", textSub:"#AAA", muted:"#777", success:"#2ECC71", danger:"#E74C3C", warning:"#F1C40F", shadow:"0 2px 12px rgba(0,0,0,0.5)", shadowMd:"0 4px 24px rgba(0,0,0,0.6)", navBg:"#1E1E1E", tag:"rgba(238,77,45,0.2)", tagText:"#FF8A70" },
};

// i18n
const TR = {
  en:{ home:"Home", cats:"Categories", cart:"Cart", orders:"Orders", admin:"Admin", search:"Search products...", all:"All", addCart:"Add to Cart", buyNow:"Buy Now", preorder:"Pre-order", outOfStock:"Out of Stock", tempOOS:"Temporarily Out of Stock", inStock:"In Stock", qty:"Qty", total:"Total", checkout:"Checkout", placeOrder:"Place Order", orderSummary:"Order Summary", contactOrder:"Contact to Order", trackOrder:"Track Order", enterOrderNo:"Enter order number", track:"Track", yourName:"Your Name *", yourPhone:"Phone *", yourAddress:"Delivery Address *", note:"Special requests / note", adminPanel:"Admin Panel", password:"Password", login:"Login", products:"Products", reports:"Reports", settings:"Settings", addProduct:"Add Product", editProduct:"Edit Product", save:"Save", cancel:"Cancel", delete:"Delete", hide:"Hide", show:"Show", addStock:"Add Stock", switchPreorder:"Switch to Pre-order", productName:"Product Name (English) *", productNameMM:"Product Name (Myanmar)", category:"Category", price:"Price (Ks) *", discountType:"Discount Type", discountPct:"Percent %", discountFixed:"Fixed Amount (Ks)", discountVal:"Discount Value", stockQty:"Stock Qty", enablePreorder:"Enable Pre-order", images:"Images", videoUrl:"Video", emoji:"Emoji", description:"Description", suitableFor:"Suitable For", benefits:"Benefits", usage:"How to Use", warning:"Warning", bulkTiers:"Bulk Tiers", addTier:"+ Add Tier", gdriveTip:"Use imgbb.com for images (Google Drive blocked by browsers)", shopName:"Shop Name (EN)", shopNameMM:"Shop Name (MM)", fbLink:"Facebook Messenger URL", viberNum:"Viber Number", waNum:"WhatsApp Number", phoneNum:"Phone Number", adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme", langLabel:"Language", saveSettings:"Save Settings", pending:"Pending", confirmed:"Confirmed", deposited:"Deposited", completed:"Completed", cancelled:"Cancelled", depositPaid:"Deposit Paid", balanceDue:"Balance Due", confirmOrder:"Confirm & Deduct Stock", updateDeposit:"Update Deposit", adminNote:"Admin Note", updateStatus:"Update Status", today:"Today", week:"This Week", month:"This Month", allTime:"All Time", revenue:"Revenue", totalOrders:"Total Orders", avgOrder:"Avg Order", topProducts:"Top Products", recentOrders:"Recent Orders", exportExcel:"Export Excel", noData:"No data yet", noProducts:"No products found", emptyCart:"Cart is empty", orderPlaced:"Order Sent!", orderPlacedMsg:"Admin will contact you soon.", wrongPw:"Wrong password", logout:"Logout", backToShop:"Back to Shop", notifications:"Notifications", clearAll:"Clear All", stockAdded:"Stock added", stockDeducted:"Stock deducted", notifUpdated:"updated", connRequired:"Supabase required" },
  mm:{ home:"ပင်မ", cats:"အမျိုးအစား", cart:"Cart", orders:"Orders", admin:"Admin", search:"ထုတ်ကုန် ရှာပါ...", all:"အားလုံး", addCart:"Cart ထည့်", buyNow:"ချက်ချင်းဝယ်", preorder:"Pre-order", outOfStock:"ကုန်ပြီ", tempOOS:"ယာယီ ကုန်ပြီ", inStock:"Stock ရှိ", qty:"အရေ", total:"စုစုပေါင်း", checkout:"Order လုပ်မည်", placeOrder:"Order ပေးပို့မည်", orderSummary:"Order အကျဉ်း", contactOrder:"ဆက်သွယ်ပြီး Order မှာမည်", trackOrder:"Order ရှာမည်", enterOrderNo:"Order နံပါတ် ထည့်ပါ", track:"ရှာမည်", yourName:"နာမည် *", yourPhone:"ဖုန်းနံပါတ် *", yourAddress:"လိပ်စာ *", note:"မှာကြားချက်", adminPanel:"Admin Panel", password:"Password", login:"Login ဝင်မည်", products:"Products", reports:"Reports", settings:"Settings", addProduct:"Product ထည့်မည်", editProduct:"Product ပြင်မည်", save:"သိမ်းမည်", cancel:"မလုပ်တော့ပါ", delete:"ဖျက်", hide:"ဖျောက်", show:"ပြ", addStock:"Stock ထည့်မည်", switchPreorder:"Pre-order ပြောင်းမည်", productName:"Product နာမည် (English) *", productNameMM:"Product နာမည် (မြန်မာ)", category:"အမျိုးအစား", price:"ဈေးနှုန်း (Ks) *", discountType:"လျှော့ဈေး အမျိုး", discountPct:"ရာခိုင်နှုန်း %", discountFixed:"ပမာဏ (Ks)", discountVal:"လျှော့ဈေး", stockQty:"Stock အရေ", enablePreorder:"Pre-order မှာနိုင်", images:"ပုံများ", videoUrl:"Video", emoji:"Emoji", description:"ဖော်ပြချက်", suitableFor:"သင့်တော်သူ", benefits:"အကျိုးကျေးဇူး", usage:"သုံးနည်း", warning:"သတိပြုရန်", bulkTiers:"အရေ Tier လျှော့ဈေး", addTier:"+ Tier ထည့်", gdriveTip:"ပုံများအတွက် imgbb.com သုံးပါ (Google Drive browser ကပိတ်ထားသည်)", shopName:"ဆိုင်နာမည် (EN)", shopNameMM:"ဆိုင်နာမည် (MM)", fbLink:"Facebook Messenger URL", viberNum:"Viber နံပါတ်", waNum:"WhatsApp နံပါတ်", phoneNum:"ဖုန်းနံပါတ်", adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme", langLabel:"ဘာသာ", saveSettings:"Settings သိမ်းမည်", pending:"စောင့်ဆိုင်း", confirmed:"အတည်ပြု", deposited:"စရံပေးပြီး", completed:"ပြီးစီး", cancelled:"ပယ်ဖျက်", depositPaid:"စရံပေးပြီး", balanceDue:"ကျန်ငွေ", confirmOrder:"အတည်ပြုပြီး Stock နှုတ်မည်", updateDeposit:"စရံ Update", adminNote:"Admin မှတ်ချက်", updateStatus:"Status ပြောင်းမည်", today:"ဒီနေ့", week:"ဒီအပတ်", month:"ဒီလ", allTime:"အားလုံး", revenue:"ဝင်ငွေ", totalOrders:"Order စုစုပေါင်း", avgOrder:"ပျမ်းမျှ", topProducts:"အရောင်းကောင်းဆုံး", recentOrders:"နောက်ဆုံး Orders", exportExcel:"Excel ထုတ်", noData:"Data မရှိသေးပါ", noProducts:"ထုတ်ကုန် မတွေ့ပါ", emptyCart:"Cart ထဲ ဘာမှ မရှိသေးပါ", orderPlaced:"Order ပေးပို့ပြီးပြီ!", orderPlacedMsg:"Admin မှ မကြာမီ ဆက်သွယ်ပါမည်။", wrongPw:"Password မှားနေသည်", logout:"Logout", backToShop:"ဆိုင်သို့ ပြန်", notifications:"အကြောင်းကြားချက်", clearAll:"အားလုံးဖျက်", stockAdded:"Stock ထည့်ပြီး", stockDeducted:"Order confirm ပြီး stock နှုတ်သွားသည်", notifUpdated:"ပြောင်းလဲပြီး", connRequired:"Supabase ချိတ်ဆက်ရန် လိုသည်" },
};

const SC = { pending:{bg:"#FFF8E1",color:"#F57F17",border:"#FFE082"}, confirmed:{bg:"#E3F2FD",color:"#1565C0",border:"#90CAF9"}, deposited:{bg:"#E8F5E9",color:"#2E7D32",border:"#A5D6A7"}, completed:{bg:"#E8F5E9",color:"#1B5E20",border:"#66BB6A"}, cancelled:{bg:"#FFEBEE",color:"#B71C1C",border:"#EF9A9A"} };

// NOTIFICATION STORE
let _nl=[];
const NS={items:[],add(msg){NS.items=[{id:uid(),msg,time:new Date().toISOString(),read:false},...NS.items].slice(0,30);_nl.forEach(f=>f([...NS.items]));if(typeof Notification!=="undefined"&&Notification.permission==="granted")new Notification("Shop",{body:msg});},markRead(){NS.items=NS.items.map(n=>({...n,read:true}));_nl.forEach(f=>f([...NS.items]));},clear(){NS.items=[];_nl.forEach(f=>f([]));},subscribe(fn){_nl.push(fn);return()=>{_nl=_nl.filter(f=>f!==fn);};}};
function useNotifs(){const[n,setN]=useState(NS.items);useEffect(()=>NS.subscribe(setN),[]);return n;}

// SETUP SQL
const SETUP_SQL = `-- Paste in Supabase SQL Editor and Run
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
DROP TABLE IF EXISTS shop_settings,orders,products,categories CASCADE;

CREATE TABLE categories(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL UNIQUE,sort_order INT DEFAULT 0,created_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE products(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL,name_mm TEXT,category TEXT NOT NULL DEFAULT \'\',description TEXT,suitable_for TEXT,benefits TEXT,usage_info TEXT,warning TEXT,price NUMERIC NOT NULL DEFAULT 0,discount_type TEXT DEFAULT \'percent\',discount_value NUMERIC DEFAULT 0,bulk_discounts JSONB DEFAULT \'[]\'::jsonb,stock INT NOT NULL DEFAULT 0,preorder BOOLEAN DEFAULT false,images JSONB DEFAULT \'[]\'::jsonb,video_url TEXT,emoji TEXT DEFAULT \'🛍️\',visible BOOLEAN DEFAULT true,featured BOOLEAN DEFAULT false,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE orders(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),order_number TEXT,customer_name TEXT,customer_phone TEXT,customer_address TEXT,items JSONB NOT NULL DEFAULT \'[]\'::jsonb,total NUMERIC NOT NULL DEFAULT 0,status TEXT DEFAULT \'pending\',contact_method TEXT,customer_note TEXT,deposit_paid NUMERIC DEFAULT 0,balance_due NUMERIC DEFAULT 0,admin_note TEXT,stock_deducted BOOLEAN DEFAULT false,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE shop_settings(key TEXT PRIMARY KEY,value TEXT NOT NULL DEFAULT \'\');

ALTER TABLE categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE shop_settings ENABLE ROW LEVEL SECURITY;

CREATE POLICY "all" ON categories FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON products FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON orders FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON shop_settings FOR ALL USING(true)WITH CHECK(true);

INSERT INTO categories(name,sort_order)VALUES(\'Hair Care\',1),(\'Face\',2),(\'Lips\',3),(\'Eyes\',4),(\'Skin Care\',5),(\'Accessories\',6),(\'Others\',7);
INSERT INTO shop_settings(key,value)VALUES(\'shop_name\',\'Beauty Store MM\'),(\'shop_name_mm\',\'ဗျူတီ စတိုး\'),(\'fb_link\',\'https://m.me/yourpage\'),(\'viber_num\',\'+95912345678\'),(\'wa_num\',\'+95912345678\'),(\'phone_num\',\'+95912345678\'),(\'admin_pw\',\'admin123\'),(\'logo\',\'\'),(\'banner\',\'\');
SELECT \'Done! \' || COUNT(*) || \' settings\' FROM shop_settings;`;


// SETUP SCREEN
function SetupScreen({onConnect}){
  const[url,setUrl]=useState("");const[key,setKey]=useState("");const[busy,setBusy]=useState(false);const[err,setErr]=useState("");const[showSql,setShowSql]=useState(false);
  const go=async()=>{if(!url||!key){setErr("URL နှင့် Key ထည့်ပါ");return;}setBusy(true);setErr("");try{const cfg={url:url.replace(/\/+$/,""),key};const r=await fetch(`${cfg.url}/rest/v1/shop_settings?limit=1`,{headers:{apikey:cfg.key,Authorization:`Bearer ${cfg.key}`}});if(!r.ok)throw new Error(await r.text());LS.set("sb_cfg",cfg);onConnect(cfg);}catch(e){setErr("ချိတ်မရပါ: "+e.message);}setBusy(false);};
  const s={width:"100%",padding:"12px 14px",borderRadius:10,border:"1.5px solid #ddd",fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:14,fontFamily:"inherit"};
  return(<div style={{minHeight:"100vh",background:"#F5F5F5",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}><div style={{background:"#fff",borderRadius:20,padding:"36px 28px",maxWidth:460,width:"100%",boxShadow:"0 8px 40px rgba(0,0,0,0.1)"}}>
    <div style={{textAlign:"center",marginBottom:22}}><div style={{fontSize:52,marginBottom:8}}>🛍️</div><div style={{fontSize:22,fontWeight:700}}>Production Setup</div><div style={{fontSize:13,color:"#999",marginTop:4}}>Supabase ချိတ်ဆက်မည်</div></div>
    <div style={{background:"#E8F5E9",borderRadius:10,padding:"12px 14px",marginBottom:14,fontSize:12,color:"#2E7D32",lineHeight:1.9}}><b>Step 1:</b> supabase.com → New Project<br/><b>Step 2:</b> SQL Editor → paste SQL → Run<br/><b>Step 3:</b> Vercel → Environment Variables → VITE_SUPABASE_URL + VITE_SUPABASE_ANON_KEY<br/><b>Step 4:</b> Project URL + anon key ထည့်ပြီး connect</div>
    <button onClick={()=>setShowSql(v=>!v)} style={{width:"100%",padding:9,borderRadius:9,border:"1.5px solid #ddd",background:"#fff",cursor:"pointer",fontSize:12,color:"#999",marginBottom:10}}>{showSql?"▲ SQL ပိတ်":"▼ Setup SQL ကြည့်မည်"}</button>
    {showSql&&<pre style={{background:"#1A1A2E",borderRadius:10,padding:12,fontSize:9.5,fontFamily:"monospace",color:"#7FFFD4",overflow:"auto",maxHeight:180,lineHeight:1.6,marginBottom:12,whiteSpace:"pre-wrap"}}>{SETUP_SQL}</pre>}
    <label style={{fontSize:12,fontWeight:700,color:"#555",display:"block",marginBottom:4}}>SUPABASE PROJECT URL</label>
    <input style={s} value={url} onChange={e=>setUrl(e.target.value)} placeholder="https://xxxx.supabase.co"/>
    <label style={{fontSize:12,fontWeight:700,color:"#555",display:"block",marginBottom:4}}>ANON PUBLIC KEY</label>
    <input style={s} type="password" value={key} onChange={e=>setKey(e.target.value)} placeholder="eyJhbGci..."/>
    {err&&<div style={{color:"#E74C3C",fontSize:13,marginBottom:12,background:"#FFEBEE",padding:"10px 12px",borderRadius:8}}>{err}</div>}
    <button onClick={go} disabled={busy} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:busy?"#ccc":"#EE4D2D",color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700}}>{busy?"⏳ စစ်ဆေးနေသည်...":"🔗 ချိတ်ဆက်မည်"}</button>
  </div></div>);
}

// PRODUCT IMAGE
function PImg({p,size=120,r=12}){
  const[errs,setErrs]=useState({});const imgs=getImgs(p);const src=imgs[0]||"";
  const bgs=["#FFE0D5","#D5E8FF","#FFE5FF","#D5FFE8","#FFEDD5","#E5D5FF"];const bg=bgs[(p?.id||"x").charCodeAt(1)%bgs.length||0];
  if(src&&!errs[0])return<img src={src} alt={p?.name||""} onError={()=>setErrs(e=>({...e,0:true}))} style={{width:size,height:size,objectFit:"cover",borderRadius:r,flexShrink:0,display:"block"}}/>;
  return<div style={{width:size,height:size,borderRadius:r,background:bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.36,flexShrink:0}}>{p?.emoji||"🛍️"}</div>;
}

// STATUS BADGE
function SBadge({status,t}){
  const c=SC[status]||SC.pending;
  const L={pending:t.pending,confirmed:t.confirmed,deposited:t.deposited,completed:t.completed,cancelled:t.cancelled};
  return<span style={{background:c.bg,color:c.color,border:`1px solid ${c.border}`,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{L[status]||status}</span>;
}

// NOTIFICATION BELL
function NotifBell({T,t}){
  const notifs=useNotifs();const[open,setOpen]=useState(false);const ref=useRef();const unread=notifs.filter(n=>!n.read).length;
  useEffect(()=>{const fn=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener("mousedown",fn);return()=>document.removeEventListener("mousedown",fn);},[]);
  return(<div ref={ref} style={{position:"relative"}}>
    <button onClick={()=>{setOpen(v=>!v);if(!open)NS.markRead();}} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:20,padding:"6px 10px",color:"#fff",cursor:"pointer",fontSize:16,position:"relative"}}>
      🔔{unread>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#fff",color:"#EE4D2D",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:700,minWidth:14,textAlign:"center"}}>{unread}</span>}
    </button>
    {open&&<div style={{position:"absolute",right:0,top:"calc(100% + 8px)",width:272,maxHeight:300,overflowY:"auto",borderRadius:12,boxShadow:T.shadowMd,zIndex:300,background:T.surface,border:`1px solid ${T.border}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderBottom:`1px solid ${T.border}`}}>
        <span style={{fontWeight:700,fontSize:13,color:T.text}}>{t.notifications}</span>
        <button onClick={()=>NS.clear()} style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:T.muted}}>{t.clearAll}</button>
      </div>
      {notifs.length===0?<div style={{padding:20,textAlign:"center",color:T.muted,fontSize:13}}>—</div>
        :notifs.map(n=><div key={n.id} style={{padding:"10px 14px",borderBottom:`1px solid ${T.border}`,opacity:n.read?0.6:1}}><div style={{fontSize:12,fontWeight:600,color:T.text}}>{n.msg}</div><div style={{fontSize:10,color:T.muted,marginTop:2}}>{fdate(n.time)}</div></div>)}
    </div>}
  </div>);
}

// TOP BAR
function TopBar({T,t,shopMM,logo,cartCount,searchQ,setSearchQ,showDrop,setShowDrop,dropRes,onHit,onCart,onAdmin,onLogo,onLangToggle,onThemeToggle,themeName}){
  return(<div style={{background:T.primary,position:"sticky",top:0,zIndex:200,boxShadow:"0 2px 8px rgba(0,0,0,0.15)"}}>
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px"}}>
      <button onClick={onLogo} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",padding:0,display:"flex",alignItems:"center",gap:6}}>
        {logo?<img src={logo} style={{width:30,height:30,borderRadius:8,objectFit:"cover"}} alt="logo"/>:<span style={{fontSize:20}}>🛍️</span>}
        <span style={{fontWeight:800,fontSize:15,color:"#fff",letterSpacing:0.2}}>{shopMM}</span>
      </button>
      <div style={{flex:1}}/>
      <button onClick={onLangToggle} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:12,padding:"5px 9px",color:"#fff",cursor:"pointer",fontSize:11,fontWeight:700}}>{t===TR.mm?"EN":"မြ"}</button>
      <button onClick={onThemeToggle} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:12,padding:"5px 9px",color:"#fff",cursor:"pointer",fontSize:13}}>{themeName==="light"?"🌙":"☀️"}</button>
      <NotifBell T={T} t={t}/>
      <button onClick={onCart} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:20,padding:"6px 12px",color:"#fff",cursor:"pointer",display:"flex",alignItems:"center",gap:4,fontSize:14,fontWeight:700}}>
        🛒{cartCount>0&&<span style={{background:"#fff",color:T.primary,borderRadius:10,padding:"1px 6px",fontSize:11,fontWeight:800}}>{cartCount}</span>}
      </button>
      <button onClick={onAdmin} style={{background:"rgba(255,255,255,0.15)",border:"none",borderRadius:20,padding:"6px 11px",color:"#fff",cursor:"pointer",fontSize:13}}>⚙️</button>
    </div>
    <div style={{padding:"0 12px 10px",position:"relative"}}>
      <div style={{position:"relative"}}>
        <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none"}}>🔍</span>
        <input value={searchQ} onChange={e=>{setSearchQ(e.target.value);setShowDrop(true);}} onFocus={()=>searchQ&&setShowDrop(true)} onBlur={()=>setTimeout(()=>setShowDrop(false),200)}
          placeholder={t.search} style={{width:"100%",padding:"9px 14px 9px 36px",borderRadius:20,border:"none",fontSize:13,outline:"none",boxSizing:"border-box",background:"rgba(255,255,255,0.95)",fontFamily:"inherit"}}/>
      </div>
      {showDrop&&searchQ&&<div style={{position:"absolute",top:"100%",left:12,right:12,background:T.surface,borderRadius:12,boxShadow:T.shadowMd,zIndex:300,overflow:"hidden",maxHeight:280,overflowY:"auto",border:`1px solid ${T.border}`}}>
        {dropRes.length===0?<div style={{padding:14,textAlign:"center",color:T.muted,fontSize:13}}>"{searchQ}" {t.noProducts}</div>
          :dropRes.map(p=><div key={p.id} onMouseDown={()=>onHit(p)} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",cursor:"pointer",borderBottom:`1px solid ${T.border}`}} onMouseEnter={e=>e.currentTarget.style.background=T.surface2} onMouseLeave={e=>e.currentTarget.style.background=T.surface}>
            <PImg p={p} size={36} r={6}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.category} · {fmt(calcUnit(p))}</div></div>
          </div>)}
      </div>}
    </div>
  </div>);
}

// BOTTOM NAV
function BottomNav({T,t,tab,setTab,cartCount}){
  const items=[{k:"home",icon:"🏠",label:t.home},{k:"cats",icon:"📦",label:t.cats},{k:"cart",icon:"🛒",label:t.cart,badge:cartCount},{k:"track",icon:"📋",label:t.orders}];
  return(<div style={{position:"fixed",bottom:0,left:0,right:0,background:T.navBg,borderTop:`1px solid ${T.border}`,display:"flex",zIndex:190,boxShadow:"0 -2px 12px rgba(0,0,0,0.08)"}}>
    {items.map(it=><button key={it.k} onClick={()=>setTab(it.k)} style={{flex:1,padding:"8px 4px 6px",border:"none",background:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,position:"relative"}}>
      <span style={{fontSize:20,opacity:tab===it.k?1:0.45}}>{it.icon}</span>
      {it.badge>0&&<span style={{position:"absolute",top:4,right:"calc(50% - 16px)",background:T.primary,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:800,minWidth:14,textAlign:"center"}}>{it.badge}</span>}
      <span style={{fontSize:10,color:tab===it.k?T.primary:T.muted,fontWeight:tab===it.k?700:400}}>{it.label}</span>
    </button>)}
  </div>);
}

// HOME PAGE
function HomePage({T,t,products,cats,catFilter,setCatFilter,onOpen,onAdd,banner}){
  const featured=products.filter(p=>p.featured||(p.discount_value||0)>0).slice(0,6);
  const filtered=catFilter==="all"?products:products.filter(p=>p.category===catFilter);
  return(<div style={{paddingBottom:72}}>
    {banner&&<div style={{margin:"8px 10px",borderRadius:12,overflow:"hidden",height:110}}><img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="banner"/></div>}
    <div style={{overflowX:"auto",scrollbarWidth:"none",padding:"10px 10px",display:"flex",gap:8,background:T.surface,borderBottom:`1px solid ${T.border}`}}>
      {[{name:"all",label:t.all},...cats.map(c=>({name:c.name,label:c.name}))].map(c=><button key={c.name} onClick={()=>setCatFilter(c.name)} style={{flexShrink:0,padding:"7px 16px",borderRadius:20,border:`1.5px solid ${catFilter===c.name?T.primary:T.border}`,background:catFilter===c.name?T.primary:T.surface,color:catFilter===c.name?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:catFilter===c.name?700:400,transition:"all 0.15s"}}>{c.label}</button>)}
    </div>
    {featured.length>0&&catFilter==="all"&&<div style={{marginTop:10}}>
      <div style={{padding:"0 10px 8px",fontSize:14,fontWeight:800,color:T.text}}>🔥 Flash Deals</div>
      <div style={{overflowX:"auto",scrollbarWidth:"none",padding:"0 10px",display:"flex",gap:10}}>
        {featured.map(p=><div key={p.id} onClick={()=>onOpen(p)} style={{flexShrink:0,width:120,background:T.surface,borderRadius:12,overflow:"hidden",cursor:"pointer",boxShadow:T.shadow}}>
          <div style={{background:T.surface2,display:"flex",justifyContent:"center",padding:"8px 8px 0"}}><PImg p={p} size={84} r={8}/></div>
          <div style={{padding:"7px 8px 10px"}}><div style={{fontSize:11,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:12,fontWeight:800,color:T.primary,marginTop:2}}>{fmt(calcUnit(p))}</div></div>
        </div>)}
      </div>
    </div>}
    <div style={{padding:"12px 10px 0"}}>
      <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10,paddingLeft:2}}>{catFilter==="all"?"🛍️ All Products":`📦 ${catFilter}`}</div>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}><div style={{fontSize:48,marginBottom:10}}>🛍️</div><div style={{fontSize:15,color:T.text}}>{t.noProducts}</div></div>
        :<div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10}}>
          {filtered.map(p=><ProductCard key={p.id} T={T} t={t} p={p} onClick={onOpen} onAdd={onAdd}/>)}
        </div>}
    </div>
  </div>);
}

// PRODUCT CARD
function ProductCard({T,t,p,onClick,onAdd}){
  const[added,setAdded]=useState(false);const up=calcUnit(p);const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;
  const doAdd=e=>{e.stopPropagation();if(!canBuy)return;onAdd(p);setAdded(true);setTimeout(()=>setAdded(false),1800);};
  return(<div onClick={()=>onClick(p)} style={{background:T.surface,borderRadius:12,overflow:"hidden",cursor:"pointer",boxShadow:T.shadow,transition:"transform 0.15s",position:"relative"}} onMouseEnter={e=>e.currentTarget.style.transform="translateY(-3px)"} onMouseLeave={e=>e.currentTarget.style.transform="none"}>
    {dv>0&&<div style={{position:"absolute",top:6,left:6,background:T.primary,color:"#fff",borderRadius:6,padding:"2px 7px",fontSize:10,fontWeight:800,zIndex:1}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</div>}
    {p.preorder&&<div style={{position:"absolute",top:dv>0?26:6,left:6,background:"#FF6F00",color:"#fff",borderRadius:6,padding:"2px 6px",fontSize:9,fontWeight:700,zIndex:1}}>Pre-order</div>}
    {oos&&!p.preorder&&<div style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.55)",color:"#fff",borderRadius:6,padding:"2px 6px",fontSize:9,zIndex:1}}>{t.tempOOS}</div>}
    <div style={{background:T.surface2,display:"flex",justifyContent:"center",padding:"12px 12px 0"}}><PImg p={p} size={108} r={8}/></div>
    <div style={{padding:"8px 10px 11px"}}>
      <div style={{fontSize:10,color:T.muted,marginBottom:2,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{p.category}</div>
      <div style={{fontSize:12,fontWeight:600,color:T.text,lineHeight:1.35,marginBottom:6,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",minHeight:32}}>{p.name_mm||p.name}</div>
      <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:800,color:T.primary,fontSize:14}}>{fmt(up)}</span>
        {dv>0&&<span style={{fontSize:11,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
      </div>
      <button onClick={doAdd} disabled={!canBuy} style={{width:"100%",padding:"7px 0",borderRadius:8,border:"none",cursor:canBuy?"pointer":"not-allowed",background:added?T.success:(oos&&!p.preorder?T.surface2:T.primary),color:oos&&!p.preorder?T.muted:"#fff",fontSize:11,fontWeight:700,transition:"background 0.2s"}}>
        {added?"✓":(p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:`🛒 ${t.addCart}`))}
      </button>
    </div>
  </div>);
}


// PRODUCT DETAIL
function ProductDetail({T,t,p,onBack,addToCart,onBuyNow}){
  const[qty,setQty]=useState(1);const[added,setAdded]=useState(false);const[imgIdx,setImgIdx]=useState(0);
  const imgs=getImgs(p);let bd=p.bulk_discounts;if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch{bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>a.min_qty-b.min_qty);
  const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;const up=calcUnit(p,qty);
  const Sec=({title,txt,warn})=>txt?<div style={{marginBottom:14}}><div style={{fontSize:13,fontWeight:700,color:warn?T.danger:T.text,marginBottom:4}}>{title}</div><div style={{fontSize:13,color:warn?"#8B0000":T.textSub,lineHeight:1.8,background:warn?"#FFEBEE":T.surface2,borderRadius:10,padding:"12px 14px"}}>{txt}</div></div>:null;
  return(<div style={{background:T.bg,minHeight:"100vh",paddingBottom:80}}>
    <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</span></div>
    <div style={{background:T.surface2}}>
      <div style={{display:"flex",justifyContent:"center",padding:"16px",height:240,alignItems:"center",overflow:"hidden"}}>
        {imgs[imgIdx]?<img src={imgs[imgIdx]} style={{maxWidth:"100%",maxHeight:240,objectFit:"contain",borderRadius:12}} alt={p.name} onError={e=>e.target.style.display="none"}/>:<PImg p={p} size={200} r={12}/>}
      </div>
      {imgs.length>1&&<div style={{display:"flex",gap:8,padding:"8px 14px",overflowX:"auto",scrollbarWidth:"none"}}>
        {imgs.map((img,i)=><div key={i} onClick={()=>setImgIdx(i)} style={{width:50,height:50,borderRadius:8,overflow:"hidden",border:`2px solid ${i===imgIdx?T.primary:T.border}`,cursor:"pointer",flexShrink:0}}><img src={img} style={{width:"100%",height:"100%",objectFit:"cover"}} alt=""/></div>)}
      </div>}
    </div>
    <div style={{background:T.surface,padding:"14px 16px",marginBottom:6}}>
      <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:6,flexWrap:"wrap"}}>
        <span style={{fontSize:24,fontWeight:800,color:T.primary}}>{fmt(up)}</span>
        {dv>0&&<span style={{fontSize:14,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
        {dv>0&&<span style={{background:T.primary,color:"#fff",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</span>}
      </div>
      <div style={{fontSize:18,fontWeight:700,color:T.text,marginBottom:4}}>{p.name_mm||p.name}</div>
      {p.name_mm&&<div style={{fontSize:13,color:T.muted,marginBottom:8}}>{p.name}</div>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <span style={{background:T.tag,color:T.tagText,padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:600}}>{p.category}</span>
        {p.preorder&&<span style={{background:"#FFF8E1",color:"#E65100",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>Pre-order</span>}
        {!oos&&!p.preorder&&<span style={{background:"#E8F5E9",color:"#2E7D32",padding:"3px 10px",borderRadius:20,fontSize:12}}>{t.inStock}: {p.stock}</span>}
        {oos&&!p.preorder&&<span style={{background:"#FFEBEE",color:"#B71C1C",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>⚠️ {t.tempOOS}</span>}
      </div>
    </div>
    {bulks.length>0&&<div style={{background:T.surface,padding:"12px 16px",marginBottom:6}}><div style={{fontSize:12,fontWeight:700,color:"#6A1B9A",marginBottom:8}}>🎁 {t.bulkTiers}</div><div style={{display:"flex",flexWrap:"wrap",gap:8}}>{bulks.map((b,i)=><div key={i} style={{background:qty>=b.min_qty?"#6A1B9A":"#F3E5F5",color:qty>=b.min_qty?"#fff":"#6A1B9A",borderRadius:8,padding:"4px 10px",fontSize:12,fontWeight:600,transition:"all 0.2s"}}>{b.min_qty}+ → -{b.discount_percent}%</div>)}</div></div>}
    {canBuy&&<div style={{background:T.surface,padding:"12px 16px",marginBottom:6}}><div style={{display:"flex",alignItems:"center",gap:14}}><span style={{fontSize:13,color:T.textSub}}>{t.qty}:</span><div style={{display:"flex",alignItems:"center",gap:8,border:`1px solid ${T.border}`,borderRadius:8,overflow:"hidden"}}><button onClick={()=>setQty(q=>Math.max(1,q-1))} style={{width:36,height:36,border:"none",background:T.surface2,cursor:"pointer",fontSize:18,fontWeight:700,color:T.text}}>−</button><span style={{minWidth:32,textAlign:"center",fontWeight:700,fontSize:16,color:T.text}}>{qty}</span><button onClick={()=>setQty(q=>q+1)} style={{width:36,height:36,border:"none",background:T.surface2,cursor:"pointer",fontSize:18,fontWeight:700,color:T.text}}>+</button></div><span style={{fontSize:13,color:T.muted}}>Total: <b style={{color:T.primary}}>{fmt(up*qty)}</b></span></div></div>}
    <div style={{background:T.surface,padding:"14px 16px",marginBottom:6}}><Sec title={`📝 ${t.description}`} txt={p.description}/><Sec title={`👤 ${t.suitableFor}`} txt={p.suitable_for}/><Sec title={`✅ ${t.benefits}`} txt={p.benefits}/><Sec title={`📋 ${t.usage}`} txt={p.usage_info}/><Sec title={`⚠️ ${t.warning}`} txt={p.warning} warn/></div>
    {p.video_url&&<div style={{background:T.surface,padding:"14px 16px",marginBottom:6}}><div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:8}}>📹 Video</div><video src={p.video_url} controls style={{width:"100%",borderRadius:10,background:"#000"}}/></div>}
    <div style={{position:"fixed",bottom:0,left:0,right:0,background:T.surface,padding:"10px 14px",boxShadow:"0 -2px 12px rgba(0,0,0,0.1)",display:"flex",gap:10,zIndex:190}}>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);setAdded(true);setTimeout(()=>setAdded(false),1800);}} disabled={!canBuy} style={{flex:1,padding:"13px 0",borderRadius:10,border:`2px solid ${canBuy?T.primary:T.border}`,background:"transparent",color:canBuy?T.primary:T.muted,cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700}}>{added?"✓ Added!":t.addCart}</button>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);onBuyNow();}} disabled={!canBuy} style={{flex:1,padding:"13px 0",borderRadius:10,border:"none",background:canBuy?T.primary:"#ccc",color:"#fff",cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700,boxShadow:canBuy?`0 4px 14px ${T.primary}50`:undefined}}>{p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:t.buyNow)}</button>
    </div>
  </div>);
}

// CART PAGE
function CartPage({T,t,cart,updateQty,removeItem,total,onCheckout,onBack}){
  if(!cart.length)return(<div style={{background:T.bg,minHeight:"100vh"}}><div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15}}>{t.cart}</span></div><div style={{textAlign:"center",padding:"80px 20px",color:T.muted}}><div style={{fontSize:60,marginBottom:14}}>🛒</div><div style={{fontSize:16,fontWeight:600,color:T.text}}>{t.emptyCart}</div></div></div>);
  return(<div style={{background:T.bg,minHeight:"100vh",paddingBottom:100}}>
    <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15}}>{t.cart} ({cart.length})</span></div>
    <div style={{padding:"10px 12px 0",display:"flex",flexDirection:"column",gap:10}}>
      {cart.map(item=>{const up=calcUnit(item.p,item.qty);return(<div key={item.p.id} style={{background:T.surface,borderRadius:12,padding:"12px",display:"flex",gap:12,alignItems:"flex-start"}}>
        <PImg p={item.p} size={68} r={8}/>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:600,color:T.text,lineHeight:1.35,marginBottom:2}}>{item.p.name_mm||item.p.name}</div>
          {item.p.preorder&&<span style={{fontSize:10,background:"#FFF8E1",color:"#E65100",padding:"1px 6px",borderRadius:6,fontWeight:700,display:"inline-block",marginBottom:4}}>Pre-order</span>}
          <div style={{fontSize:14,color:T.primary,fontWeight:800,marginBottom:8}}>{fmt(up)}</div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <button onClick={()=>updateQty(item.p.id,item.qty-1)} style={{width:28,height:28,borderRadius:14,border:`1px solid ${T.border}`,background:T.surface2,cursor:"pointer",fontWeight:700,color:T.text,display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
            <span style={{minWidth:24,textAlign:"center",fontWeight:700,fontSize:14,color:T.text}}>{item.qty}</span>
            <button onClick={()=>updateQty(item.p.id,item.qty+1)} style={{width:28,height:28,borderRadius:14,border:`1px solid ${T.border}`,background:T.surface2,cursor:"pointer",fontWeight:700,color:T.text,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
            <span style={{marginLeft:"auto",fontSize:14,fontWeight:700,color:T.text}}>{fmt(up*item.qty)}</span>
          </div>
        </div>
        <button onClick={()=>removeItem(item.p.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:T.muted,padding:0}}>✕</button>
      </div>);})}
    </div>
    <div style={{position:"sticky",bottom:0,background:T.surface,padding:"12px 16px",boxShadow:"0 -2px 12px rgba(0,0,0,0.08)",display:"flex",gap:12,alignItems:"center",marginTop:10}}>
      <div style={{flex:1}}><div style={{fontSize:11,color:T.muted}}>{t.total}</div><div style={{fontSize:22,fontWeight:800,color:T.primary}}>{fmt(total)}</div></div>
      <button onClick={onCheckout} style={{padding:"13px 28px",borderRadius:12,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 14px ${T.primary}50`}}>{t.checkout} →</button>
    </div>
  </div>);
}

// CHECKOUT PAGE
function CheckoutPage({T,t,cart,total,settings,onPlaced,onBack}){
  const[name,setName]=useState("");const[phone,setPhone]=useState("");const[address,setAddress]=useState("");const[note,setNote]=useState("");const[sent,setSent]=useState(false);const[sentOrder,setSentOrder]=useState(null);
  const hasPreorder=cart.some(i=>i.p.preorder);
  const msgLines=cart.map(i=>`• ${i.p.name_mm||i.p.name}${i.p.preorder?" [PRE-ORDER]":""} x${i.qty} = ${fmt(calcUnit(i.p,i.qty)*i.qty)}`).join("\n");
  const msg=encodeURIComponent(`မင်္ဂလာပါ!\nName: ${name}\nPhone: ${phone}\nAddress: ${address}\n\nOrder:\n${msgLines}\nTotal: ${fmt(total)}${hasPreorder?"\n⚠️ Pre-order ပါဝင်":""}${note?"\nNote: "+note:""}`);
  const saveOrder=async(method)=>{
    if(sent)return;
    const newOrd={order_number:onum(),customer_name:name,customer_phone:phone,customer_address:address,customer_note:note,contact_method:method,items:cart.map(i=>({id:i.p.id,name:i.p.name,name_mm:i.p.name_mm,qty:i.qty,is_preorder:i.p.preorder,unit_price:calcUnit(i.p,i.qty),total:calcUnit(i.p,i.qty)*i.qty})),total,status:"pending",deposit_paid:0,balance_due:total,stock_deducted:false,admin_note:"",created_at:new Date().toISOString()};
    try{const rows=await sb("POST","orders",newOrd);const saved=Array.isArray(rows)?rows[0]:newOrd;setSentOrder(saved);setSent(true);onPlaced(saved);}catch(e){alert("Error saving order: "+e.message);}
  };
  const contacts=[
    {label:"Facebook",sub:"Messenger",color:"#1877F2",emoji:"💬",key:"messenger",url:`https://m.me/${(settings.fb_link||"").replace(/^https?:\/\/m\.me\//,"")}?text=${msg}`},
    {label:"Viber",sub:settings.viber_num,color:"#7360F2",emoji:"📱",key:"viber",url:`viber://chat?number=${encodeURIComponent(settings.viber_num||"")}&text=${msg}`},
    {label:"WhatsApp",sub:settings.wa_num,color:"#25D366",emoji:"💚",key:"whatsapp",url:`https://wa.me/${(settings.wa_num||"").replace(/\D/g,"")}?text=${msg}`},
    {label:"Phone",sub:settings.phone_num,color:"#546E7A",emoji:"📞",key:"phone",url:`tel:${settings.phone_num}`},
  ];
  const inp={width:"100%",padding:"11px 13px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10,fontFamily:"inherit",background:T.surface2,color:T.text};
  return(<div style={{background:T.bg,minHeight:"100vh",paddingBottom:40}}>
    <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15}}>{t.checkout}</span></div>
    {sent?<div style={{padding:20}}><div style={{background:T.surface,borderRadius:16,padding:28,textAlign:"center",boxShadow:T.shadow}}><div style={{fontSize:52,marginBottom:12}}>✅</div><div style={{fontSize:18,fontWeight:700,color:T.success,marginBottom:8}}>{t.orderPlaced}</div><div style={{fontSize:13,color:T.muted,marginBottom:16}}>{t.orderPlacedMsg}</div>{sentOrder&&<div style={{background:T.surface2,borderRadius:10,padding:"10px 14px",fontSize:12,color:T.text}}><b>Order #:</b> {sentOrder.order_number}</div>}</div></div>
    :<div style={{padding:12,display:"flex",flexDirection:"column",gap:10}}>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:10}}>{t.orderSummary}</div>
        {hasPreorder&&<div style={{background:"#FFF8E1",border:"1px solid #FFE082",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12,color:"#E65100",fontWeight:600}}>⚠️ Pre-order items — Admin နှင့် timeline ညှိနှိုင်းပါ</div>}
        {cart.map(item=><div key={item.p.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`,fontSize:13}}><span style={{color:T.text,flex:1,paddingRight:10}}>{item.p.name_mm||item.p.name} × {item.qty}{item.p.preorder?" [PRE]":""}</span><span style={{fontWeight:700,color:T.text,whiteSpace:"nowrap"}}>{fmt(calcUnit(item.p,item.qty)*item.qty)}</span></div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"12px 0 0",fontSize:16,fontWeight:800}}><span style={{color:T.text}}>{t.total}</span><span style={{color:T.primary}}>{fmt(total)}</span></div>
      </div>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:10}}>📋 Customer Info</div>
        <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder={t.yourName}/>
        <input style={inp} value={phone} onChange={e=>setPhone(e.target.value)} placeholder={t.yourPhone} type="tel"/>
        <input style={inp} value={address} onChange={e=>setAddress(e.target.value)} placeholder={t.yourAddress}/>
        <textarea value={note} onChange={e=>setNote(e.target.value)} placeholder={t.note} style={{...inp,minHeight:56,resize:"vertical",marginBottom:0}}/>
      </div>
      {(!name.trim()||!phone.trim())&&<div style={{background:"#FFF3E0",border:"1px solid #FFB74D",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#E65100"}}>⚠️ Name & Phone ထည့်ပြီးမှ order လုပ်ပါ</div>}
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:12}}>{t.contactOrder}</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {contacts.map(ct=><a key={ct.key} href={(name.trim()&&phone.trim())?ct.url:"#"} onClick={e=>{if(!name.trim()||!phone.trim()){e.preventDefault();alert("Name & Phone ထည့်ပါ");return;}saveOrder(ct.key);}} target="_blank" rel="noopener" style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",background:ct.color,borderRadius:12,textDecoration:"none",color:"#fff",boxShadow:`0 3px 12px ${ct.color}60`,opacity:(name.trim()&&phone.trim())?1:0.5}}>
            <span style={{fontSize:22}}>{ct.emoji}</span><div><div style={{fontWeight:700,fontSize:14}}>{ct.label}</div><div style={{fontSize:11,opacity:0.85}}>{ct.sub}</div></div><span style={{marginLeft:"auto",fontSize:18,opacity:0.8}}>→</span>
          </a>)}
        </div>
      </div>
    </div>}
  </div>);
}

// ORDER TRACKING
function TrackOrder({T,t}){
  const[q,setQ]=useState("");const[result,setResult]=useState(null);const[busy,setBusy]=useState(false);const[err,setErr]=useState("");
  const search=async()=>{if(!q.trim())return;setBusy(true);setErr("");setResult(null);try{const rows=await sb("GET","orders",null,`order_number=eq.${q.trim()}`);if(!rows||rows.length===0)setErr("Order မတွေ့ပါ");else setResult(rows[0]);}catch(e){setErr(e.message);}setBusy(false);};
  return(<div style={{background:T.bg,minHeight:"100vh",padding:"16px 14px",paddingBottom:80}}>
    <div style={{fontSize:18,fontWeight:800,color:T.text,marginBottom:4}}>📋 {t.trackOrder}</div>
    <div style={{fontSize:13,color:T.muted,marginBottom:14}}>Order number ဖြင့် status စစ်ဆေးနိုင်သည်</div>
    <div style={{display:"flex",gap:8,marginBottom:14}}>
      <input value={q} onChange={e=>setQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&search()} placeholder={t.enterOrderNo} style={{flex:1,padding:"12px 14px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",fontFamily:"inherit",background:T.surface,color:T.text,minWidth:0,boxSizing:"border-box"}}/>
      <button onClick={search} disabled={busy} style={{padding:"12px 16px",borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0}}>{busy?"...":t.track}</button>
    </div>
    {err&&<div style={{background:"#FFEBEE",border:"1px solid #EF9A9A",borderRadius:10,padding:"12px 14px",fontSize:13,color:"#B71C1C",marginBottom:12}}>⚠️ {err}</div>}
    {result&&<div style={{background:T.surface,borderRadius:14,padding:18,boxShadow:T.shadow}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}><div><div style={{fontSize:16,fontWeight:800,color:T.text}}>{result.order_number}</div><div style={{fontSize:12,color:T.muted}}>{fdate(result.created_at)}</div></div><SBadge status={result.status} t={t}/></div>
      {(result.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`,fontSize:13}}><span style={{color:T.text}}>{item.name_mm||item.name} × {item.qty}{item.is_preorder?" [PRE]":""}</span><span style={{fontWeight:700}}>{fmt(item.total)}</span></div>)}
      <div style={{paddingTop:12,borderTop:`2px solid ${T.border}`,marginTop:4}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:15,fontWeight:800,marginBottom:8}}><span style={{color:T.text}}>{t.total}</span><span style={{color:T.primary}}>{fmt(result.total)}</span></div>
        {(result.deposit_paid||0)>0&&<><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{color:T.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:T.success,fontWeight:700}}>{fmt(result.deposit_paid)}</span></div><div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:T.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:T.danger,fontWeight:700}}>{fmt(Math.max(0,result.total-(result.deposit_paid||0)))}</span></div></>}
      </div>
      {result.admin_note&&<div style={{background:T.surface2,borderRadius:8,padding:"10px 12px",marginTop:10,fontSize:12,color:T.textSub}}><b>Note:</b> {result.admin_note}</div>}
    </div>}
  </div>);
}


// ADMIN ORDER MANAGER
function OrderMgr({T,t,products,setProducts}){
  const[orders,setOrders]=useState([]);const[sel,setSel]=useState(null);const[dep,setDep]=useState("");const[noteVal,setNoteVal]=useState("");const[busy,setBusy]=useState(false);const[filter,setFilter]=useState("all");
  const load=useCallback(async()=>{try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){console.error(e);}},[]);
  useEffect(()=>{load();},[load]);
  const updateOrder=async(id,fields)=>{await sb("PATCH","orders",{...fields,updated_at:new Date().toISOString()},`id=eq.${id}`);setOrders(ords=>ords.map(o=>o.id===id?{...o,...fields}:o));if(sel?.id===id)setSel(o=>({...o,...fields}));};

  const confirmAndDeduct=async()=>{
    if(!sel||sel.stock_deducted)return;
    if(!window.confirm("Order confirm ပြီး stock နှုတ်မည်?"))return;
    setBusy(true);
    try{
      const updatedProds=[...products];
      for(const item of(sel.items||[])){
        if(item.is_preorder)continue;
        const pidx=updatedProds.findIndex(p=>p.id===item.id);if(pidx===-1)continue;
        const p=updatedProds[pidx];const newStock=Math.max(0,(p.stock||0)-(item.qty||0));
        await sb("PATCH","products",{stock:newStock,updated_at:new Date().toISOString()},`id=eq.${p.id}`);
        updatedProds[pidx]={...p,stock:newStock};
        NS.add(`${t.stockDeducted}: ${p.name_mm||p.name} -${item.qty} (Stock: ${newStock})`);
        if(newStock===0)NS.add(`⚠️ ${p.name_mm||p.name} — ${t.tempOOS}`);
      }
      setProducts(updatedProds);
      await updateOrder(sel.id,{status:"confirmed",stock_deducted:true});
    }catch(e){alert("Error: "+e.message);}
    setBusy(false);
  };

  const statuses=["pending","confirmed","deposited","completed","cancelled"];
  const filtered=filter==="all"?orders:orders.filter(o=>o.status===filter);

  if(sel){
    const dp=Number(sel.deposit_paid||0);const bal=Math.max(0,sel.total-dp);
    return(<div style={{paddingBottom:80}}>
      <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={()=>{setSel(null);load();}} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15}}>Order Detail</span></div>
      <div style={{padding:12,display:"flex",flexDirection:"column",gap:10}}>
        <div style={{background:T.surface,borderRadius:12,padding:16}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}><div><div style={{fontSize:15,fontWeight:800,color:T.text}}>{sel.order_number}</div><div style={{fontSize:12,color:T.muted}}>{fdate(sel.created_at)}</div></div><SBadge status={sel.status} t={t}/></div>
          <div style={{background:T.surface2,borderRadius:8,padding:"10px 12px",marginBottom:10,fontSize:12}}><div style={{color:T.text}}><b>Name:</b> {sel.customer_name} | <b>Phone:</b> {sel.customer_phone}</div>{sel.customer_address&&<div style={{color:T.textSub,marginTop:2}}><b>Address:</b> {sel.customer_address}</div>}{sel.contact_method&&<div style={{color:T.muted,marginTop:2}}><b>Via:</b> {sel.contact_method}</div>}</div>
          {(sel.items||[]).some(i=>i.is_preorder)&&<div style={{background:"#FFF8E1",border:"1px solid #FFE082",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:11,color:"#E65100",fontWeight:700}}>⚠️ Pre-order items ပါဝင်</div>}
          {(sel.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${T.border}`,fontSize:13}}><div><span style={{color:T.text,fontWeight:600}}>{item.name_mm||item.name}</span>{item.is_preorder&&<span style={{fontSize:9,background:"#FFF8E1",color:"#E65100",padding:"1px 5px",borderRadius:5,fontWeight:700,marginLeft:4}}>PRE</span>}<span style={{color:T.muted}}> × {item.qty}</span></div><span style={{fontWeight:700,color:T.text}}>{fmt(item.total)}</span></div>)}
          <div style={{paddingTop:12,borderTop:`2px solid ${T.border}`,marginTop:4}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:T.text}}>{t.total}</span><span style={{color:T.primary}}>{fmt(sel.total)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}><span style={{color:T.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:T.success,fontWeight:700}}>{fmt(dp)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:T.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:T.danger,fontWeight:700}}>{fmt(bal)}</span></div>
          </div>
          {sel.customer_note&&<div style={{background:T.surface2,borderRadius:8,padding:"8px 12px",marginTop:10,fontSize:12,color:T.textSub}}><b>Customer note:</b> {sel.customer_note}</div>}
        </div>
        {!sel.stock_deducted?<button onClick={confirmAndDeduct} disabled={busy} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:busy?"#ccc":T.success,color:"#fff",cursor:busy?"wait":"pointer",fontSize:14,fontWeight:700}}>{busy?"⏳ Processing...":t.confirmOrder}</button>:<div style={{background:"#E8F5E9",borderRadius:10,padding:"10px 14px",fontSize:13,color:"#2E7D32",fontWeight:600,textAlign:"center"}}>✅ Stock already deducted</div>}
        <div style={{background:T.surface,borderRadius:12,padding:14}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.updateStatus}</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>{statuses.map(s=><button key={s} onClick={()=>updateOrder(sel.id,{status:s})} style={{padding:"5px 10px",borderRadius:20,border:`1.5px solid ${sel.status===s?T.primary:T.border}`,background:sel.status===s?T.primary:"transparent",color:sel.status===s?"#fff":T.text,cursor:"pointer"}}><SBadge status={s} t={t}/></button>)}</div>
        </div>
        <div style={{background:T.surface,borderRadius:12,padding:14}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.updateDeposit}</div>
          <div style={{display:"flex",gap:8}}><input type="number" min="0" value={dep} onChange={e=>setDep(e.target.value)} placeholder="0 Ks" style={{flex:1,padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",background:T.surface2,color:T.text,fontFamily:"inherit",minWidth:0}}/><button onClick={async()=>{const d=Number(dep);await updateOrder(sel.id,{deposit_paid:d,balance_due:Math.max(0,sel.total-d)});setDep("");}} style={{padding:"10px 16px",borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0}}>OK</button></div>
        </div>
        <div style={{background:T.surface,borderRadius:12,padding:14}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.adminNote}</div>
          <textarea value={noteVal||sel.admin_note||""} onChange={e=>setNoteVal(e.target.value)} style={{width:"100%",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",resize:"vertical",minHeight:72,boxSizing:"border-box",fontFamily:"inherit",background:T.surface2,color:T.text,marginBottom:10}}/>
          <button onClick={()=>updateOrder(sel.id,{admin_note:noteVal||sel.admin_note||""})} style={{width:"100%",padding:10,borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>{t.save}</button>
        </div>
      </div>
    </div>);
  }

  return(<div style={{paddingBottom:80}}>
    <div style={{display:"flex",gap:6,padding:"10px 10px",overflowX:"auto",scrollbarWidth:"none",background:T.surface,borderBottom:`1px solid ${T.border}`}}>
      {["all",...statuses].map(s=><button key={s} onClick={()=>setFilter(s)} style={{flexShrink:0,padding:"6px 12px",borderRadius:20,border:`1.5px solid ${filter===s?T.primary:T.border}`,background:filter===s?T.primary:"transparent",color:filter===s?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:600}}>{s==="all"?t.all:<SBadge status={s} t={t}/>}</button>)}
    </div>
    <div style={{padding:"10px 12px",display:"flex",flexDirection:"column",gap:8}}>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}><div style={{fontSize:40,marginBottom:8}}>📦</div><div>{t.noData}</div></div>
        :filtered.map(o=><div key={o.id} onClick={()=>{setSel(o);setNoteVal(o.admin_note||"");setDep("");}} style={{background:T.surface,borderRadius:12,padding:"12px 14px",cursor:"pointer",boxShadow:T.shadow,transition:"transform 0.1s"}} onMouseEnter={e=>e.currentTarget.style.transform="translateX(3px)"} onMouseLeave={e=>e.currentTarget.style.transform="none"}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}><div><div style={{fontSize:13,fontWeight:700,color:T.text}}>{o.order_number}</div><div style={{fontSize:11,color:T.muted}}>{fdate(o.created_at)} · {o.customer_name} · {o.contact_method}</div></div><SBadge status={o.status} t={t}/></div>
          {(o.items||[]).some(i=>i.is_preorder)&&<div style={{fontSize:10,color:"#E65100",fontWeight:700,marginBottom:4}}>⚠️ PREORDER</div>}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:12,color:T.muted}}>{(o.items||[]).length} items{o.stock_deducted?" · ✅ Stock OK":""}</div><div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:800,color:T.primary}}>{fmt(o.total)}</div>{(o.deposit_paid||0)>0&&<div style={{fontSize:10,color:T.success}}>Dep: {fmt(o.deposit_paid)} · Bal: {fmt(Math.max(0,o.total-(o.deposit_paid||0)))}</div>}</div></div>
        </div>)}
    </div>
  </div>);
}

// STOCK MANAGER
function StockMgr({T,t,products,setProducts}){
  const[adding,setAdding]=useState({});const[vals,setVals]=useState({});
  const doAdd=async(p)=>{const n=Number(vals[p.id]||0);if(!n||n<=0){alert("Valid qty ထည့်ပါ");return;}const ns=(p.stock||0)+n;try{await sb("PATCH","products",{stock:ns,preorder:false,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,stock:ns,preorder:false}:x));NS.add(`${t.stockAdded}: ${p.name_mm||p.name} +${n} (Total: ${ns})`);setAdding(v=>({...v,[p.id]:false}));setVals(v=>({...v,[p.id]:""}));}catch(e){alert("Error: "+e.message);}};
  const togglePO=async(p)=>{const np=!p.preorder;try{await sb("PATCH","products",{preorder:np,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,preorder:np}:x));NS.add(`${p.name_mm||p.name} — ${np?"Pre-order ON":"Pre-order OFF"}`);}catch(e){alert(e.message);}};
  const oos_=products.filter(p=>isOOS(p));const low=products.filter(p=>(p.stock||0)<=5&&!isOOS(p));
  return(<div style={{padding:12,paddingBottom:80}}>
    {oos_.length>0&&<div style={{background:"#FFEBEE",border:"1px solid #EF9A9A",borderRadius:10,padding:"10px 14px",marginBottom:10}}><div style={{fontSize:13,fontWeight:700,color:"#B71C1C",marginBottom:4}}>⚠️ Out of Stock ({oos_.length})</div>{oos_.map(p=><div key={p.id} style={{fontSize:12,color:"#B71C1C"}}>• {p.name_mm||p.name}{p.preorder?" [Pre-order]":""}</div>)}</div>}
    {low.length>0&&<div style={{background:"#FFF8E1",border:"1px solid #FFE082",borderRadius:10,padding:"10px 14px",marginBottom:10}}><div style={{fontSize:13,fontWeight:700,color:"#E65100",marginBottom:4}}>🔶 Low Stock (≤5)</div>{low.map(p=><div key={p.id} style={{fontSize:12,color:"#E65100"}}>• {p.name_mm||p.name}: {p.stock}</div>)}</div>}
    <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:10}}>📦 All Stock</div>
    {products.map(p=>{const oos__=isOOS(p);return(<div key={p.id} style={{background:T.surface,borderRadius:10,padding:"12px 14px",marginBottom:8,boxShadow:T.shadow}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}><PImg p={p} size={40} r={8}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,fontWeight:700,marginTop:1,color:oos__?T.danger:(p.stock||0)<=5?T.warning:T.success}}>{oos__?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`}</div></div></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {adding[p.id]?<div style={{display:"flex",gap:6,flex:1}}><input type="number" min="1" value={vals[p.id]||""} onChange={e=>setVals(v=>({...v,[p.id]:e.target.value}))} placeholder="qty" style={{width:70,padding:"7px 10px",borderRadius:8,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",background:T.surface2,color:T.text,fontFamily:"inherit"}}/><button onClick={()=>doAdd(p)} style={{padding:"7px 14px",borderRadius:8,border:"none",background:T.success,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add</button><button onClick={()=>setAdding(v=>({...v,[p.id]:false}))} style={{padding:"7px 10px",borderRadius:8,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:12}}>✕</button></div>
          :<button onClick={()=>setAdding(v=>({...v,[p.id]:true}))} style={{padding:"7px 14px",borderRadius:8,border:`1.5px solid ${T.success}`,background:"transparent",color:T.success,cursor:"pointer",fontSize:12,fontWeight:700}}>➕ {t.addStock}</button>}
        <button onClick={()=>togglePO(p)} style={{padding:"7px 12px",borderRadius:8,border:`1.5px solid ${p.preorder?T.warning:T.border}`,background:p.preorder?"#FFF8E1":"transparent",color:p.preorder?"#E65100":T.muted,cursor:"pointer",fontSize:11,fontWeight:600}}>{p.preorder?"✅ Pre-order ON":`📋 ${t.switchPreorder}`}</button>
      </div>
    </div>);})}
  </div>);
}

// REPORTS
function AdminReports({T,t}){
  const[orders,setOrders]=useState([]);const[period,setPeriod]=useState("week");const[loading,setLoading]=useState(true);
  useEffect(()=>{(async()=>{setLoading(true);try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){}setLoading(false);})();},[]);
  const filtered=useMemo(()=>{const now=new Date();return orders.filter(o=>{const d=new Date(o.created_at);if(period==="today")return d.toDateString()===now.toDateString();if(period==="week"){const w=new Date(now);w.setDate(w.getDate()-7);return d>=w;}if(period==="month"){const m=new Date(now);m.setDate(m.getDate()-30);return d>=m;}return true;});},[orders,period]);
  const active=filtered.filter(o=>o.status!=="cancelled");
  const rev=active.reduce((s,o)=>s+Number(o.total),0);const dep=active.reduce((s,o)=>s+(Number(o.deposit_paid)||0),0);
  const tops=useMemo(()=>{const m={};active.forEach(o=>(o.items||[]).forEach(item=>{const k=item.name||"?";if(!m[k])m[k]={name:item.name_mm||item.name||k,qty:0,rev:0};m[k].qty+=(item.qty||0);m[k].rev+=item.total||0;}));return Object.values(m).sort((a,b)=>b.rev-a.rev).slice(0,6);},[active]);
  const exportExcel=()=>{
    const wb=XLSX.utils.book_new();
    const oRows=[["Order#","Date","Status","Customer","Phone","Address","Items","Total","Deposit","Balance","Contact","Customer Note","Admin Note"]];
    filtered.forEach(o=>oRows.push([o.order_number,fdate(o.created_at),o.status,o.customer_name,o.customer_phone,o.customer_address,(o.items||[]).map(i=>`${i.name_mm||i.name}x${i.qty}${i.is_preorder?"[PRE]":""}`).join(", "),o.total,o.deposit_paid||0,Math.max(0,o.total-(o.deposit_paid||0)),o.contact_method,o.customer_note||"",o.admin_note||""]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(oRows),"Orders");
    const sRows=[["Metric","Value"],["Period",period],["Total Revenue",rev],["Deposit Received",dep],["Balance Due",rev-dep],["Total Orders",filtered.length],["Active Orders",active.length],["Cancelled",filtered.length-active.length]];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),"Summary");
    if(tops.length>0){const pRows=[["Product","Qty Sold","Revenue"],...tops.map(p=>[p.name,p.qty,p.rev])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(pRows),"Top Products");}
    XLSX.writeFile(wb,`Report_${period}_${new Date().toISOString().slice(0,10)}.xlsx`);
  };
  const periods=[{k:"today",l:t.today},{k:"week",l:t.week},{k:"month",l:t.month},{k:"year",l:t.allTime}];
  const Stat=({emoji,label,value,color})=><div style={{background:T.surface,borderRadius:12,padding:14,flex:1,minWidth:80}}><div style={{fontSize:18,marginBottom:2}}>{emoji}</div><div style={{fontSize:11,color:T.muted,fontWeight:600}}>{label}</div><div style={{fontSize:14,fontWeight:800,color:color||T.text,marginTop:2}}>{value}</div></div>;
  return(<div style={{padding:12,paddingBottom:80}}>
    <div style={{display:"flex",gap:6,marginBottom:14,background:T.surface,borderRadius:10,padding:5}}>{periods.map(({k,l})=><button key={k} onClick={()=>setPeriod(k)} style={{flex:1,padding:"8px 4px",borderRadius:8,border:"none",background:period===k?T.primary:"transparent",color:period===k?"#fff":T.muted,cursor:"pointer",fontSize:12,fontWeight:period===k?700:400}}>{l}</button>)}</div>
    {loading?<div style={{textAlign:"center",padding:40,color:T.muted}}>Loading...</div>:<>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}><Stat emoji="💰" label={t.revenue} value={fmt(rev)} color={T.primary}/><Stat emoji="📦" label={t.totalOrders} value={filtered.length}/><Stat emoji="✅" label="Deposit" value={fmt(dep)} color={T.success}/><Stat emoji="⏳" label="Balance" value={fmt(rev-dep)} color={T.danger}/></div>
      {tops.length>0&&<div style={{background:T.surface,borderRadius:12,padding:14,marginBottom:12}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>🏆 {t.topProducts}</div>
        {tops.map((p,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 0",borderBottom:i<tops.length-1?`1px solid ${T.border}`:"none"}}><span style={{width:20,height:20,borderRadius:10,background:["#FFB347","#C0C0C0","#CD7F32","#AAA","#AAA","#AAA"][i],display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{i+1}</span><span style={{flex:1,fontSize:12,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span><span style={{fontSize:11,color:T.muted,marginRight:4}}>{p.qty}ခု</span><span style={{fontSize:12,fontWeight:700,color:T.primary,whiteSpace:"nowrap"}}>{fmt(p.rev)}</span></div>)}
      </div>}
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontSize:13,fontWeight:700,color:T.text}}>{t.recentOrders}</div><button onClick={exportExcel} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 12px",borderRadius:8,border:"none",background:"#21A366",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>📊 {t.exportExcel}</button></div>
        {filtered.slice(0,10).map(o=><div key={o.id} style={{padding:"9px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}><span style={{fontWeight:700,color:T.text}}>{o.order_number} · {o.customer_name}</span><SBadge status={o.status} t={t}/></div><div style={{display:"flex",justifyContent:"space-between",color:T.muted}}><span>{fdate(o.created_at)} · {o.contact_method}</span><span style={{fontWeight:700,color:T.primary}}>{fmt(o.total)}</span></div></div>)}
        {filtered.length===0&&<div style={{textAlign:"center",color:T.muted,padding:20,fontSize:13}}>{t.noData}</div>}
      </div>
    </>}
  </div>);
}

// PRODUCT FORM
function ProductForm({T,t,product,cats,onSave,onCancel}){
  const blank={name:"",name_mm:"",category:cats[0]?.name||"",description:"",suitable_for:"",benefits:"",usage_info:"",warning:"",price:"",discount_type:"percent",discount_value:"0",bulk_discounts:[],stock:"0",preorder:false,images:[],video_url:"",emoji:"🛍️",featured:false};
  const[form,setForm]=useState(product?{...product,discount_value:product.discount_value??0,bulk_discounts:Array.isArray(product.bulk_discounts)?product.bulk_discounts:[]}:blank);
  const[saving,setSaving]=useState(false);const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const iRef=useRef();const vRef=useRef();const[urlIn,setUrlIn]=useState("");
  const upImg=async e=>{const b64s=await Promise.all(Array.from(e.target.files||[]).map(fileToB64));set("images",[...(form.images||[]),...b64s]);e.target.value="";};
  const upVid=async e=>{const f=e.target.files?.[0];if(!f)return;set("video_url",await fileToB64(f));e.target.value="";};
  const addUrl=()=>{if(!urlIn.trim())return;set("images",[...(form.images||[]),urlIn.trim()]);setUrlIn("");};
  const handleSave=async()=>{
    if(!form.name.trim()||!form.price){alert("Name & Price required");return;}
    setSaving(true);
    try{
      const data={...form,price:Number(form.price),discount_value:Number(form.discount_value)||0,stock:Number(form.stock)||0,visible:product?.visible??true,updated_at:new Date().toISOString()};
      if(!product)data.created_at=new Date().toISOString();
      const rows=product?.id?await sb("PATCH","products",data,`id=eq.${product.id}`):await sb("POST","products",data);
      onSave(Array.isArray(rows)?rows[0]:(rows||{...data,id:uid()}));
    }catch(e){alert("Error: "+e.message);}
    setSaving(false);
  };
  const inp={width:"100%",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",background:T.surface2,color:T.text};
  const ta={...inp,minHeight:60,resize:"vertical"};
  const lbl={fontSize:11,fontWeight:700,color:T.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{background:T.bg,minHeight:"100vh",paddingBottom:40}}>
    <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><button onClick={onCancel} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:20,padding:0}}>←</button><span style={{color:"#fff",fontWeight:700,fontSize:15}}>{product?t.editProduct:t.addProduct}</span></div>
    <div style={{padding:12,display:"flex",flexDirection:"column",gap:10}}>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <label style={lbl}>{t.productName}</label><input style={inp} value={form.name} onChange={e=>set("name",e.target.value)} placeholder="e.g. Argan Oil Hair Serum"/>
        <label style={lbl}>{t.productNameMM}</label><input style={inp} value={form.name_mm||""} onChange={e=>set("name_mm",e.target.value)} placeholder="optional"/>
        <label style={lbl}>{t.category}</label>
        {cats.length===0?<div style={{padding:"8px 12px",borderRadius:8,background:"#FFEBEE",fontSize:12,color:T.danger,marginBottom:12}}>⚠️ Add categories first</div>:<select style={{...inp,background:T.surface2}} value={form.category} onChange={e=>set("category",e.target.value)}>{cats.map(c=><option key={c.id} value={c.name}>{c.name}</option>)}</select>}
        <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,alignItems:"end",marginBottom:12}}><div><label style={lbl}>{t.emoji}</label><input style={{...inp,marginBottom:0}} value={form.emoji||""} onChange={e=>set("emoji",e.target.value)}/></div><label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",paddingBottom:1,whiteSpace:"nowrap"}}><input type="checkbox" checked={form.featured||false} onChange={e=>set("featured",e.target.checked)} style={{width:15,height:15,accentColor:T.primary}}/>⭐ Featured</label></div>
      </div>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <label style={lbl}>📝 {t.description}</label><textarea style={ta} value={form.description||""} onChange={e=>set("description",e.target.value)}/>
        <label style={lbl}>👤 {t.suitableFor}</label><textarea style={{...ta,minHeight:48}} value={form.suitable_for||""} onChange={e=>set("suitable_for",e.target.value)}/>
        <label style={lbl}>✅ {t.benefits}</label><textarea style={{...ta,minHeight:48}} value={form.benefits||""} onChange={e=>set("benefits",e.target.value)}/>
        <label style={lbl}>📋 {t.usage}</label><textarea style={{...ta,minHeight:48}} value={form.usage_info||""} onChange={e=>set("usage_info",e.target.value)}/>
        <label style={lbl}>⚠️ {t.warning}</label><textarea style={{...ta,minHeight:44,marginBottom:0}} value={form.warning||""} onChange={e=>set("warning",e.target.value)}/>
      </div>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>{t.price}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.price} onChange={e=>set("price",e.target.value)}/></div><div><label style={lbl}>{t.discountType}</label><select style={{...inp,marginBottom:0,background:T.surface2}} value={form.discount_type} onChange={e=>set("discount_type",e.target.value)}><option value="percent">{t.discountPct}</option><option value="fixed">{t.discountFixed}</option></select></div></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>{t.discountVal} ({form.discount_type==="fixed"?"Ks":"%"})</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.discount_value} onChange={e=>set("discount_value",e.target.value)}/></div><div><label style={lbl}>{t.stockQty}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.stock} onChange={e=>set("stock",e.target.value)}/></div></div>
        <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",userSelect:"none",marginBottom:12,fontSize:13}}><input type="checkbox" checked={form.preorder||false} onChange={e=>set("preorder",e.target.checked)} style={{width:15,height:15,accentColor:T.primary}}/>{t.enablePreorder}</label>
        <div style={{fontSize:11,fontWeight:700,color:T.textSub,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>{t.bulkTiers}</div>
        {(Array.isArray(form.bulk_discounts)?form.bulk_discounts:[]).map((b,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}><input type="number" min="2" value={b.min_qty} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,min_qty:Number(e.target.value)}:x))} style={{width:52,padding:"6px 8px",borderRadius:8,border:`1.5px solid ${T.border}`,fontSize:12,outline:"none",textAlign:"center",background:T.surface2,color:T.text}}/><span style={{fontSize:11,color:T.muted}}>ခု+</span><input type="number" min="1" max="99" value={b.discount_percent} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,discount_percent:Number(e.target.value)}:x))} style={{width:52,padding:"6px 8px",borderRadius:8,border:`1.5px solid ${T.border}`,fontSize:12,outline:"none",textAlign:"center",background:T.surface2,color:T.text}}/><span style={{fontSize:11,color:T.muted}}>% off</span><button onClick={()=>set("bulk_discounts",form.bulk_discounts.filter((_,j)=>j!==i))} style={{background:"none",border:"none",cursor:"pointer",color:T.danger,fontSize:15,marginLeft:"auto"}}>✕</button></div>)}
        <button onClick={()=>set("bulk_discounts",[...(form.bulk_discounts||[]),{min_qty:2,discount_percent:10}])} style={{padding:"6px 12px",borderRadius:8,border:`1.5px dashed ${T.border}`,background:"transparent",cursor:"pointer",fontSize:12,color:T.primary,fontWeight:600}}>{t.addTier}</button>
      </div>
      <div style={{background:T.surface,borderRadius:12,padding:14}}>
        <div style={{fontSize:11,color:T.muted,marginBottom:8,fontStyle:"italic"}}>💡 {t.gdriveTip}</div>
        <div style={{fontSize:11,fontWeight:700,color:T.textSub,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>🖼️ {t.images}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
          {(Array.isArray(form.images)?form.images:[]).map((img,i)=><div key={i} style={{position:"relative"}}><img src={img} style={{width:54,height:54,objectFit:"cover",borderRadius:8,border:`1px solid ${T.border}`,display:"block"}} onError={e=>e.target.style.opacity="0.3"}/><button onClick={()=>set("images",(form.images||[]).filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:17,height:17,borderRadius:9,border:"none",background:T.danger,color:"#fff",cursor:"pointer",fontSize:9,lineHeight:"17px",textAlign:"center",padding:0}}>✕</button></div>)}
          <button onClick={()=>iRef.current?.click()} style={{width:54,height:54,borderRadius:8,border:`2px dashed ${T.border}`,background:T.surface2,cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted}}>+</button>
        </div>
        <input ref={iRef} type="file" accept="image/*" multiple onChange={upImg} style={{display:"none"}}/>
        <div style={{display:"flex",gap:8,marginBottom:12}}><input style={{flex:1,padding:"9px 11px",borderRadius:9,border:`1.5px solid ${T.border}`,fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,background:T.surface2,color:T.text}} value={urlIn} onChange={e=>setUrlIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addUrl()} placeholder="imgbb.com URL"/><button onClick={addUrl} style={{padding:"9px 12px",borderRadius:9,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,flexShrink:0}}>Add</button></div>
        <div style={{fontSize:11,fontWeight:700,color:T.textSub,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>🎬 {t.videoUrl}</div>
        <div style={{display:"flex",gap:8}}><input style={{flex:1,padding:"9px 11px",borderRadius:9,border:`1.5px solid ${T.border}`,fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,background:T.surface2,color:T.text}} value={form.video_url||""} onChange={e=>set("video_url",e.target.value)} placeholder="Video URL (mp4...)"/><button onClick={()=>vRef.current?.click()} style={{padding:"9px 11px",borderRadius:9,border:`1.5px solid ${T.border}`,background:T.surface2,color:T.text,cursor:"pointer",fontSize:11,flexShrink:0}}>Upload</button></div>
        <input ref={vRef} type="file" accept="video/*" onChange={upVid} style={{display:"none"}}/>
      </div>
      <div style={{display:"flex",gap:10}}><button onClick={handleSave} disabled={saving} style={{flex:1,padding:14,borderRadius:12,border:"none",background:saving?"#ccc":T.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700}}>{saving?"⏳ Saving...":(product?t.save:t.addProduct)}</button><button onClick={onCancel} style={{flex:1,padding:14,borderRadius:12,border:`1.5px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:14}}>{t.cancel}</button></div>
    </div>
  </div>);
}

// ADMIN SETTINGS
function AdminSettings({T,t,settings,onSave,onTheme,onLang,themeName,lang,onDisconnect}){
  const[form,setForm]=useState({...settings});const[logo,setLogo]=useState(settings.logo||"");const[banner,setBanner]=useState(settings.banner||"");const[saving,setSaving]=useState(false);
  const logoRef=useRef();const bannerRef=useRef();const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const upLogo=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setLogo(b);set("logo",b);e.target.value="";};
  const upBanner=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setBanner(b);set("banner",b);e.target.value="";};
  const doSave=async()=>{setSaving(true);try{const data={...form,logo,banner};const rows=Object.entries(data).map(([key,value])=>({key,value:String(value||"")}));await sbUpsert("shop_settings",rows);onSave(data);alert("✅ Saved!");}catch(e){alert("Error: "+e.message);}setSaving(false);};
  const inp={width:"100%",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",background:T.surface2,color:T.text};
  const lbl={fontSize:11,fontWeight:700,color:T.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{padding:12,paddingBottom:80}}>
    <div style={{background:T.surface,borderRadius:12,padding:14,marginBottom:10}}>
      <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>🎨 {t.themeLabel} & {t.langLabel}</div>
      <div style={{display:"flex",gap:8,marginBottom:10}}>{Object.keys(THEMES).map(k=><button key={k} onClick={()=>onTheme(k)} style={{flex:1,padding:10,borderRadius:10,border:`2px solid ${themeName===k?T.primary:T.border}`,background:themeName===k?T.primary:"transparent",color:themeName===k?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:700}}>{k==="light"?"☀️ Light":"🌙 Dark"}</button>)}</div>
      <div style={{display:"flex",gap:8}}>{["mm","en"].map(l=><button key={l} onClick={()=>onLang(l)} style={{flex:1,padding:10,borderRadius:10,border:`2px solid ${lang===l?T.primary:T.border}`,background:lang===l?T.primary:"transparent",color:lang===l?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:700}}>{l==="mm"?"🇲🇲 မြန်မာ":"🇬🇧 English"}</button>)}</div>
    </div>
    <div style={{background:T.surface,borderRadius:12,padding:14,marginBottom:10}}>
      <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>🖼️ Logo & Banner</div>
      <div style={{display:"flex",gap:12,marginBottom:10}}>
        <div style={{textAlign:"center"}}><div style={{width:68,height:68,borderRadius:12,border:`2px dashed ${T.border}`,display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",background:T.surface2,cursor:"pointer"}} onClick={()=>logoRef.current?.click()}>{logo?<img src={logo} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:26}}>🛍️</span>}</div><div style={{fontSize:10,color:T.muted,marginTop:3}}>Logo</div></div>
        <div style={{flex:1}}><div style={{height:68,borderRadius:12,border:`2px dashed ${T.border}`,display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",background:T.surface2,cursor:"pointer"}} onClick={()=>bannerRef.current?.click()}>{banner?<img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:12,color:T.muted}}>+ Banner Image</span>}</div></div>
      </div>
      <input ref={logoRef} type="file" accept="image/*" onChange={upLogo} style={{display:"none"}}/>
      <input ref={bannerRef} type="file" accept="image/*" onChange={upBanner} style={{display:"none"}}/>
    </div>
    <div style={{background:T.surface,borderRadius:12,padding:14,marginBottom:10}}>
      <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:12}}>🏪 Shop Info</div>
      <label style={lbl}>{t.shopName}</label><input style={inp} value={form.shop_name||""} onChange={e=>set("shop_name",e.target.value)}/>
      <label style={lbl}>{t.shopNameMM}</label><input style={inp} value={form.shop_name_mm||""} onChange={e=>set("shop_name_mm",e.target.value)}/>
      <label style={lbl}>{t.fbLink}</label><input style={inp} value={form.fb_link||""} onChange={e=>set("fb_link",e.target.value)} placeholder="https://m.me/yourpage"/>
      <label style={lbl}>{t.viberNum}</label><input style={inp} value={form.viber_num||""} onChange={e=>set("viber_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.waNum}</label><input style={inp} value={form.wa_num||""} onChange={e=>set("wa_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.phoneNum}</label><input style={inp} value={form.phone_num||""} onChange={e=>set("phone_num",e.target.value)} placeholder="+95912345678"/>
      <label style={lbl}>{t.adminPw}</label><input style={{...inp,marginBottom:0}} type="password" value={form.admin_pw||""} onChange={e=>set("admin_pw",e.target.value)}/>
    </div>
    <button onClick={doSave} disabled={saving} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:saving?"#ccc":T.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700,marginBottom:10}}>{saving?"⏳ Saving...":t.saveSettings}</button>
    <button onClick={onDisconnect} style={{width:"100%",padding:12,borderRadius:12,border:`1.5px solid ${T.danger}`,background:"transparent",color:T.danger,cursor:"pointer",fontSize:13,fontWeight:600}}>🔌 Disconnect Supabase</button>
  </div>);
}


// ADMIN PANEL
function AdminPanel({T,t,cfg,products,cats,settings,onProdChange,onCatsChange,onSettingsChange,onTheme,onLang,themeName,lang,onBack,onDisconnect}){
  const[tab,setTab]=useState("products");const[showForm,setShowForm]=useState(false);const[editing,setEditing]=useState(null);const[newCat,setNewCat]=useState("");
  const tabs=[{k:"products",l:`📦 ${t.products}`},{k:"stock",l:"📊 Stock"},{k:"orders",l:`🛍️ ${t.orders}`},{k:"reports",l:`📈 ${t.reports}`},{k:"settings",l:`⚙️ ${t.settings}`}];
  const delProd=async p=>{if(!window.confirm(`Delete "${p.name_mm||p.name}"?`))return;try{await sb("DELETE","products",null,`id=eq.${p.id}`);onProdChange(products.filter(x=>x.id!==p.id));}catch(e){alert(e.message);}};
  const toggleVis=async p=>{try{await sb("PATCH","products",{visible:!p.visible},`id=eq.${p.id}`);onProdChange(products.map(x=>x.id===p.id?{...x,visible:!x.visible}:x));NS.add(`${p.name_mm||p.name} — ${!p.visible?"visible":"hidden"}`);}catch(e){alert(e.message);}};
  const addCat=async name=>{if(!name||cats.find(c=>c.name===name)){return;}try{const rows=await sb("POST","categories",{name,sort_order:cats.length+1});onCatsChange([...cats,...(Array.isArray(rows)?rows:[rows])]);}catch(e){alert(e.message);}};
  const delCat=async cat=>{if(!window.confirm(`Delete "${cat.name}"?`))return;try{await sb("DELETE","categories",null,`id=eq.${cat.id}`);onCatsChange(cats.filter(c=>c.id!==cat.id));}catch(e){alert(e.message);}};

  if(showForm)return<ProductForm T={T} t={t} product={editing} cats={cats} onSave={saved=>{if(editing)onProdChange(products.map(x=>x.id===saved.id?saved:x));else onProdChange([saved,...products]);NS.add(editing?`${t.notifUpdated}: ${saved.name_mm||saved.name}`:`New: ${saved.name_mm||saved.name}`);setShowForm(false);setEditing(null);}} onCancel={()=>{setShowForm(false);setEditing(null);}}/>;

  return(<div style={{background:T.bg,minHeight:"100vh",paddingBottom:40}}>
    <div style={{background:T.primary,padding:"10px 14px",display:"flex",alignItems:"center",gap:8}}><span style={{color:"#fff",fontWeight:800,fontSize:16,flex:1}}>⚙️ {t.adminPanel}</span><button onClick={onBack} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:16,padding:"6px 12px",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>{t.logout} ←</button></div>
    <div style={{display:"flex",background:T.surface,borderBottom:`1px solid ${T.border}`,overflowX:"auto",scrollbarWidth:"none"}}>
      {tabs.map(tab_=><button key={tab_.k} onClick={()=>setTab(tab_.k)} style={{flex:"none",padding:"12px 12px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:tab===tab_.k?700:500,color:tab===tab_.k?T.primary:T.muted,borderBottom:`2.5px solid ${tab===tab_.k?T.primary:"transparent"}`,whiteSpace:"nowrap"}}>{tab_.l}</button>)}
    </div>
    {tab==="products"&&<div style={{padding:12,paddingBottom:80}}>
      <button onClick={()=>{setEditing(null);setShowForm(true);}} style={{width:"100%",padding:13,borderRadius:12,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:12}}>➕ {t.addProduct}</button>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {products.map(p=><div key={p.id} style={{background:T.surface,borderRadius:10,padding:"11px 12px",display:"flex",gap:10,alignItems:"center",opacity:p.visible?1:0.5}}>
          <PImg p={p} size={50} r={8}/>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,color:T.muted,marginTop:1}}>{p.category} · {fmt(calcUnit(p))}</div><div style={{fontSize:10,marginTop:1,color:isOOS(p)?T.danger:(p.stock||0)<=5?T.warning:T.success,fontWeight:600}}>{isOOS(p)?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`} · {p.visible?"🟢":"⚫"}</div></div>
          <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
            <button onClick={()=>{setEditing(p);setShowForm(true);}} style={{padding:"5px 9px",borderRadius:7,border:`1px solid ${T.border}`,background:"transparent",cursor:"pointer",fontSize:12,color:T.text}}>✏️</button>
            <button onClick={()=>toggleVis(p)} style={{padding:"5px 9px",borderRadius:7,border:`1px solid ${T.border}`,background:"transparent",cursor:"pointer",fontSize:12,color:T.text}}>{p.visible?"🙈":"👁️"}</button>
            <button onClick={()=>delProd(p)} style={{padding:"5px 9px",borderRadius:7,border:`1px solid ${T.danger}50`,background:"#FFEBEE",cursor:"pointer",fontSize:12,color:T.danger}}>🗑️</button>
          </div>
        </div>)}
        {products.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:T.muted,fontSize:14}}>{t.noProducts}</div>}
      </div>
      <div style={{background:T.surface,borderRadius:12,padding:14,marginTop:14}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>🏷️ Categories</div>
        <div style={{display:"flex",gap:8,marginBottom:10}}><input value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>e.key==="Enter"&&newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} placeholder="New category" style={{flex:1,padding:"9px 12px",borderRadius:9,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",fontFamily:"inherit",background:T.surface2,color:T.text,minWidth:0,boxSizing:"border-box"}}/><button onClick={()=>newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} style={{padding:"9px 14px",borderRadius:9,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700,flexShrink:0}}>+</button></div>
        {cats.map(cat=><div key={cat.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",background:T.surface2,borderRadius:8,marginBottom:6}}><span style={{fontSize:13,fontWeight:600,color:T.text}}>{cat.name}</span><button onClick={()=>delCat(cat)} style={{background:"none",border:`1px solid ${T.danger}50`,borderRadius:7,cursor:"pointer",color:T.danger,fontSize:11,padding:"3px 9px"}}>🗑️</button></div>)}
      </div>
    </div>}
    {tab==="stock"&&<StockMgr T={T} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="orders"&&<OrderMgr T={T} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="reports"&&<AdminReports T={T} t={t}/>}
    {tab==="settings"&&<AdminSettings T={T} t={t} settings={settings} onSave={s=>{onSettingsChange(s);}} onTheme={onTheme} onLang={onLang} themeName={themeName} lang={lang} onDisconnect={onDisconnect}/>}
  </div>);
}

// MAIN APP
export default function App(){
  const[cfg,setCfg]=useState(null);const[ready,setReady]=useState(false);const[products,setProducts]=useState([]);const[cats,setCats]=useState([]);const[settings,setSettings]=useState({});
  const[cart,setCart]=useState([]);const[tab,setTab]=useState("home");const[catFilter,setCatFilter]=useState("all");const[searchQ,setSearchQ]=useState("");const[showDrop,setShowDrop]=useState(false);
  const[selProd,setSelProd]=useState(null);const[page,setPage]=useState("home");const[adminIn,setAdminIn]=useState(false);
  const[themeName,setThemeName]=useState(LS.get("theme")||"light");const[lang,setLang]=useState(LS.get("lang")||"mm");
  const T=THEMES[themeName]||THEMES.light;const t=TR[lang]||TR.mm;

  useEffect(()=>{
    (async()=>{
      let initCfg=null;
      if(SB_URL&&SB_KEY){initCfg={url:SB_URL,key:SB_KEY};}
      else{initCfg=LS.get("sb_cfg");}
      if(initCfg){setCfg(initCfg);}else{setReady(true);}
    })();
  },[]);

  useEffect(()=>{
    if(!cfg)return;
    (async()=>{
      setReady(false);
      try{
        const[prods,catsData,setsData]=await Promise.all([
          sb("GET","products",null,"order=created_at.desc&visible=eq.true"),
          sb("GET","categories",null,"order=sort_order.asc"),
          sb("GET","shop_settings"),
        ]);
        setProducts(prods||[]);setCats(catsData||[]);
        const smap={};(setsData||[]).forEach(s=>{smap[s.key]=s.value;});setSettings(smap);
        if(smap.theme)setThemeName(smap.theme);if(smap.language)setLang(smap.language);
      }catch(e){console.error("Load error:",e);}
      setReady(true);
    })();
  },[cfg]);

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

  if(!ready&&!cfg)return<SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;
  if(!cfg)return<SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;
  if(!ready)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:T.bg,flexDirection:"column",gap:12}}><div style={{fontSize:52}}>🛍️</div><div style={{color:T.muted,fontSize:15,fontWeight:600}}>Loading...</div></div>);

  if(tab==="admin"){
    if(!adminIn)return(<div style={{background:T.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}><div style={{background:T.surface,borderRadius:20,padding:"36px 28px",maxWidth:360,width:"100%",boxShadow:T.shadowMd}}><div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:52}}>🔐</div><div style={{fontSize:20,fontWeight:800,color:T.text}}>{t.adminPanel}</div></div><AdminLoginBox T={T} t={t} adminPw={settings.admin_pw} onSuccess={()=>setAdminIn(true)} onBack={()=>setTab("home")}/></div></div>);
    return<AdminPanel T={T} t={t} cfg={cfg} products={products} cats={cats} settings={settings} onProdChange={setProducts} onCatsChange={setCats} onSettingsChange={onSettingsChange} onTheme={changeTheme} onLang={changeLang} themeName={themeName} lang={lang} onBack={()=>{setTab("home");setAdminIn(false);setProducts(ps=>ps.filter(p=>p.visible!==false));}} onDisconnect={()=>{LS.set("sb_cfg",null);setCfg(null);setAdminIn(false);setTab("home");}}/>;
  }

  if(page==="product"&&selProd)return<ProductDetail T={T} t={t} p={selProd} onBack={()=>setPage("home")} addToCart={addToCart} onBuyNow={()=>setPage("checkout")}/>;
  if(page==="checkout")return<CheckoutPage T={T} t={t} cart={cart} total={cartTotal} settings={settings} onPlaced={()=>{setCart([]);setPage("home");}} onBack={()=>setPage("home")}/>;

  const shopMM=settings.shop_name_mm||settings.shop_name||"Shop";
  return(<div style={{background:T.bg,minHeight:"100vh",fontFamily:"'Segoe UI','Myanmar Text',Helvetica,sans-serif"}}>
    <TopBar T={T} t={t} shopMM={shopMM} logo={settings.logo||""} cartCount={cartCount} searchQ={searchQ} setSearchQ={setSearchQ} showDrop={showDrop} setShowDrop={setShowDrop} dropRes={searchRes} onHit={openProd} onCart={()=>setTab("cart")} onAdmin={()=>setTab("admin")} onLogo={()=>{setTab("home");setPage("home");}} onLangToggle={()=>changeLang(lang==="mm"?"en":"mm")} onThemeToggle={()=>changeTheme(themeName==="light"?"dark":"light")} themeName={themeName}/>
    <div style={{paddingBottom:72}}>
      {tab==="home"&&<HomePage T={T} t={t} products={visProds} cats={cats} catFilter={catFilter} setCatFilter={setCatFilter} onOpen={openProd} onAdd={addToCart} banner={settings.banner||""}/>}
      {tab==="cart"&&<CartPage T={T} t={t} cart={cart} updateQty={updateQty} removeItem={removeItem} total={cartTotal} onCheckout={()=>setPage("checkout")} onBack={()=>setTab("home")}/>}
      {tab==="track"&&<TrackOrder T={T} t={t}/>}
      {tab==="cats"&&<div style={{padding:12}}><div style={{fontSize:15,fontWeight:700,color:T.text,marginBottom:12}}>{t.cats}</div><div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>{cats.map(c=><div key={c.id} onClick={()=>{setCatFilter(c.name);setTab("home");}} style={{background:T.surface,borderRadius:12,padding:"16px 8px",textAlign:"center",cursor:"pointer",boxShadow:T.shadow}}><div style={{fontSize:26,marginBottom:4}}>📦</div><div style={{fontSize:12,fontWeight:600,color:T.text}}>{c.name}</div></div>)}</div></div>}
    </div>
    <BottomNav T={T} t={t} tab={tab} setTab={t2=>{setTab(t2);setPage("home");}} cartCount={cartCount}/>
  </div>);
}

function AdminLoginBox({T,t,adminPw,onSuccess,onBack}){
  const[pw,setPw]=useState("");const[err,setErr]=useState("");
  const check=()=>{if(pw===(adminPw||"admin123"))onSuccess();else setErr(t.wrongPw);};
  return(<><input type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&check()} placeholder="admin123" style={{width:"100%",padding:"12px 14px",borderRadius:10,border:`1.5px solid ${err?"#EF9A9A":T.border}`,fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:err?8:14,fontFamily:"inherit",background:T.surface2,color:T.text}}/>{err&&<div style={{color:T.danger,fontSize:12,marginBottom:12,background:"#FFEBEE",padding:"8px 12px",borderRadius:8}}>{err}</div>}<button onClick={check} style={{width:"100%",padding:13,borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:10}}>{t.login}</button><button onClick={onBack} style={{width:"100%",padding:11,borderRadius:10,border:`1.5px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:13}}>{t.backToShop}</button></>);
}
