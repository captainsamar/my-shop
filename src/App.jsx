import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════
//  THEMES
// ═══════════════════════════════════════════════════════════════
const THEMES = {
  light: {
    name:"Light",
    primary:"#FF5733", primaryHover:"#E84D2B", accent:"#FFB347",
    bg:"#FFF8F5", surface:"#FFFFFF", surface2:"#F7F3F0",
    border:"rgba(0,0,0,0.08)", borderStrong:"rgba(0,0,0,0.15)",
    text:"#1A202C", textSub:"#4A5568", muted:"#718096",
    success:"#38A169", danger:"#E53E3E", info:"#3182CE",
    warning:"#D69E2E", purple:"#805AD5",
    shadow:"0 4px 20px rgba(0,0,0,0.08)",
    shadowMd:"0 8px 32px rgba(0,0,0,0.12)",
    tag:"#FFF0EA", tagText:"#C44B2A",
    glass:false,
  },
  dark: {
    name:"Dark",
    primary:"#FF6B4A", primaryHover:"#FF5733", accent:"#FFB347",
    bg:"#0F1117", surface:"#1A1D27", surface2:"#252836",
    border:"rgba(255,255,255,0.08)", borderStrong:"rgba(255,255,255,0.15)",
    text:"#F0F2F5", textSub:"#A0AEC0", muted:"#718096",
    success:"#48BB78", danger:"#FC8181", info:"#63B3ED",
    warning:"#F6E05E", purple:"#B794F4",
    shadow:"0 4px 20px rgba(0,0,0,0.4)",
    shadowMd:"0 8px 32px rgba(0,0,0,0.5)",
    tag:"rgba(255,87,51,0.15)", tagText:"#FF8C6B",
    glass:false,
  },
  glass: {
    name:"Glass",
    primary:"#FF5733", primaryHover:"#E84D2B", accent:"#FFB347",
    bg:"linear-gradient(135deg,#667eea 0%,#764ba2 50%,#f093fb 100%)",
    surface:"rgba(255,255,255,0.15)", surface2:"rgba(255,255,255,0.08)",
    border:"rgba(255,255,255,0.25)", borderStrong:"rgba(255,255,255,0.4)",
    text:"#FFFFFF", textSub:"rgba(255,255,255,0.8)", muted:"rgba(255,255,255,0.6)",
    success:"#68D391", danger:"#FC8181", info:"#90CDF4",
    warning:"#F6E05E", purple:"#D6BCFA",
    shadow:"0 8px 32px rgba(31,38,135,0.37)",
    shadowMd:"0 16px 48px rgba(31,38,135,0.5)",
    tag:"rgba(255,255,255,0.15)", tagText:"#FFFFFF",
    glass:true,
  },
};

// ═══════════════════════════════════════════════════════════════
//  i18n
// ═══════════════════════════════════════════════════════════════
const TR = {
  en: {
    shopName:"Beauty Store MM", search:"Search products (Myanmar / English)",
    allCat:"All", addCart:"Add to Cart", addedCart:"Added!",
    preorder:"Pre-order", outOfStock:"Out of Stock", cart:"Cart",
    viewCart:"View Cart", orderNow:"Order Now",
    orderSummary:"Order Summary", total:"Total",
    contactOrder:"Contact to Place Order",
    contactMsg:"Click to send your order and arrange payment",
    note:"Note (delivery address, special requests...)",
    stock:"Stock", category:"Category",
    description:"Description", suitableFor:"Suitable For",
    benefits:"Benefits", usage:"Usage", warning:"Warning",
    bulkDiscount:"Bulk Discount", qty:"Qty",
    noProducts:"No products found", emptyCart:"Your cart is empty",
    addProducts:"Add some products first",
    adminPanel:"Admin Panel", password:"Password",
    login:"Login", back:"Back to Shop",
    products:"Products", categories:"Categories",
    reports:"Reports", settings:"Settings",
    addProduct:"+ Add Product", editProduct:"Edit Product",
    saveProduct:"Save Product", cancel:"Cancel",
    delete:"Delete", hide:"Hide", show:"Show",
    productName:"Product Name (English)*", productNameMM:"Product Name (Myanmar)",
    price:"Price (Ks)*", discountType:"Discount Type",
    discountValue:"Discount Value", stockQty:"Stock Qty",
    preorderEnable:"Enable Pre-order", images:"Images",
    videoUrl:"Video URL", emoji:"Emoji",
    shopSettings:"Shop Settings", contactLinks:"Contact Links",
    adminPassword:"Admin Password", save:"Save",
    logoUpload:"Shop Logo",
    // Order management
    orders:"Orders", orderList:"Order List",
    orderNo:"Order #", orderDate:"Date",
    orderStatus:"Status", orderItems:"Items",
    orderTotal:"Total", depositPaid:"Deposit Paid",
    balanceDue:"Balance Due", confirmOrder:"Confirm Order",
    pendingStatus:"Pending", confirmedStatus:"Confirmed",
    depositedStatus:"Deposited", completedStatus:"Completed",
    cancelledStatus:"Cancelled",
    updateDeposit:"Update Deposit", adminNote:"Admin Note",
    isPreorder:"Preorder Item", contactMethod:"Contact Via",
    customerNote:"Customer Note", updateStatus:"Update Status",
    // Reports
    today:"Today", thisWeek:"This Week", thisMonth:"This Month", allTime:"All Time",
    revenue:"Revenue", totalOrders:"Total Orders", avgOrder:"Avg Order",
    topProducts:"Top Products", recentOrders:"Recent Orders",
    exportExcel:"Export Excel", noData:"No data yet",
    // Notification
    notif:"Notifications", clearAll:"Clear All",
    productUpdated:"Product Updated",
    newProductAdded:"New product added",
    theme:"Theme", language:"Language",
    wrongPassword:"Wrong password",
    logout:"Logout",
    percent:"Percent %", fixedKs:"Fixed Amount (Ks)",
    addTier:"+ Add Tier", bulkTiers:"Quantity Tiers",
    enterMin:"Min Qty", discount:"Discount",
    gdriveTip:"Google Drive: paste share link directly",
    addImageUrl:"Add Image URL", uploadImage:"Upload",
    uploadVideo:"Upload Video",
  },
  mm: {
    shopName:"ဗျူတီ စတိုး MM", search:"Product ရှာပါ (မြန်မာ / English)",
    allCat:"အားလုံး", addCart:"Cart ထည့်", addedCart:"ထည့်ပြီး!",
    preorder:"Pre-order", outOfStock:"ကုန်ပြီ", cart:"Cart",
    viewCart:"Cart ကြည့်မည်", orderNow:"Order မှာမည်",
    orderSummary:"Order အကျဉ်းချုပ်", total:"စုစုပေါင်း",
    contactOrder:"ဆက်သွယ်ပြီး Order ပေးပို့မည်",
    contactMsg:"ကြိုက်နှစ်သက်ရာ နည်းလမ်းဖြင့် order ပေးပို့နိုင်သည်",
    note:"Note (delivery address, မှာကြားချက်...)",
    stock:"Stock", category:"အမျိုးအစား",
    description:"ဖော်ပြချက်", suitableFor:"သင့်တော်သူ",
    benefits:"အကျိုးကျေးဇူး", usage:"သုံးနည်း", warning:"သတိပြုရန်",
    bulkDiscount:"အရေအတွက်လျှော့ဈေး", qty:"အရေအတွက်",
    noProducts:"ထုတ်ကုန် မတွေ့ပါ", emptyCart:"Cart ထဲ ဘာမှ မရှိသေးပါ",
    addProducts:"ထုတ်ကုန်များ ထည့်ပါ",
    adminPanel:"Admin Panel", password:"Password",
    login:"Login ဝင်ရောက်မည်", back:"ဆိုင်သို့ ပြန်သွားမည်",
    products:"Products", categories:"Categories",
    reports:"Reports", settings:"Settings",
    addProduct:"+ Product အသစ် ထည့်မည်", editProduct:"Product ပြင်မည်",
    saveProduct:"သိမ်းမည်", cancel:"ဖျက်သိမ်းမည်",
    delete:"ဖျက်", hide:"ဖျောက်", show:"ပြ",
    productName:"Product နာမည် (English)*", productNameMM:"Product နာမည် (မြန်မာ)",
    price:"ဈေးနှုန်း (Ks)*", discountType:"လျှော့ဈေး အမျိုးအစား",
    discountValue:"လျှော့ဈေး", stockQty:"Stock အရေအတွက်",
    preorderEnable:"Pre-order မှာနိုင်သည်", images:"ပုံများ",
    videoUrl:"Video URL", emoji:"Emoji",
    shopSettings:"ဆိုင်အချက်အလက်", contactLinks:"ဆက်သွယ်ရေး Links",
    adminPassword:"Admin Password", save:"သိမ်းမည်",
    logoUpload:"ဆိုင် Logo",
    orders:"Orders", orderList:"Order စာရင်း",
    orderNo:"Order အမှတ်", orderDate:"ရက်စွဲ",
    orderStatus:"အခြေအနေ", orderItems:"Items",
    orderTotal:"စုစုပေါင်း", depositPaid:"စရံပေးပြီး",
    balanceDue:"ကျန်ငွေ", confirmOrder:"Order အတည်ပြုမည်",
    pendingStatus:"စောင့်ဆိုင်းနေ", confirmedStatus:"အတည်ပြုပြီး",
    depositedStatus:"စရံပေးပြီး", completedStatus:"ပြီးစီး",
    cancelledStatus:"ပယ်ဖျက်ပြီး",
    updateDeposit:"စရံ Update လုပ်မည်", adminNote:"Admin မှတ်ချက်",
    isPreorder:"Preorder ပစ္စည်း", contactMethod:"ဆက်သွယ်နည်း",
    customerNote:"Customer မှာကြားချက်", updateStatus:"Status ပြောင်းမည်",
    today:"ဒီနေ့", thisWeek:"ဒီအပတ်", thisMonth:"ဒီလ", allTime:"အားလုံး",
    revenue:"ဝင်ငွေ", totalOrders:"Order စုစုပေါင်း", avgOrder:"ပျမ်းမျှ Order",
    topProducts:"အရောင်းကောင်းဆုံး", recentOrders:"နောက်ဆုံး Orders",
    exportExcel:"Excel ထုတ်မည်", noData:"Data မရှိသေးပါ",
    notif:"အကြောင်းကြားချက်", clearAll:"အားလုံးဖျက်",
    productUpdated:"Product ပြောင်းလဲမှုရှိသည်",
    newProductAdded:"Product အသစ် ထည့်ထားသည်",
    theme:"Theme", language:"ဘာသာစကား",
    wrongPassword:"Password မှားနေသည်",
    logout:"Logout",
    percent:"ရာခိုင်နှုန်း %", fixedKs:"ပမာဏ (Ks)",
    addTier:"+ Tier ထည့်မည်", bulkTiers:"အရေအတွက် Tiers",
    enterMin:"အနည်းဆုံး Qty", discount:"လျှော့",
    gdriveTip:"Google Drive: share link ကို တိုက်ရိုက် ထည့်နိုင်သည်",
    addImageUrl:"Image URL ထည့်မည်", uploadImage:"Upload",
    uploadVideo:"Video Upload",
  }
};

// ═══════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════
const ST = {
  get: async k => { try { const r=await window.storage.get(k); return r?JSON.parse(r.value):null; } catch(e){return null;} },
  set: async (k,v) => { try { await window.storage.set(k,JSON.stringify(v)); } catch(e){} },
};

// Convert Google Drive share link → direct viewable URL
const fixMediaUrl = url => {
  if (!url) return url;
  // Google Drive file share
  let m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=view&id=${m[1]}`;
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=view&id=${m[1]}`;
  // Google Drive open link
  m = url.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=view&id=${m[1]}`;
  return url;
};

const calcUnit = (p, qty=1) => {
  if(!p) return 0;
  let bd=p.bulk_discounts; if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch(e){bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>b.min_qty-a.min_qty);
  const bulk=bulks.find(b=>qty>=b.min_qty);
  if(bulk) return Math.round((p.price||0)*(1-bulk.discount_percent/100));
  if((p.discount_value||0)>0)
    return p.discount_type==="fixed" ? Math.max(0,(p.price||0)-p.discount_value) : Math.round((p.price||0)*(1-p.discount_value/100));
  return p.price||0;
};
const fmt = n => (n||0).toLocaleString("en-US")+" Ks";
const uid = () => Math.random().toString(36).slice(2)+Date.now().toString(36);
const onum = () => "ORD-"+Date.now().toString(36).toUpperCase();
const fdate = s => s ? new Date(s).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"}) : "";
const fileToB64 = f => new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(f);});

const STATUS_COLORS = {
  pending:   {bg:"#FFFBEB",text:"#92600A",border:"#F6E05E"},
  confirmed: {bg:"#EBF8FF",text:"#2C5282",border:"#90CDF4"},
  deposited: {bg:"#F0FFF4",text:"#276749",border:"#9AE6B4"},
  completed: {bg:"#F0FFF4",text:"#276749",border:"#68D391"},
  cancelled: {bg:"#FFF5F5",text:"#742A2A",border:"#FEB2B2"},
};

const SEED_CATS=[
  {id:"c1",name:"Hair Care",sort_order:1},{id:"c2",name:"Face",sort_order:2},
  {id:"c3",name:"Lips",sort_order:3},{id:"c4",name:"Eyes",sort_order:4},
  {id:"c5",name:"Accessories",sort_order:5},{id:"c6",name:"Others",sort_order:6},
];
const SEED_SETTINGS={
  shop_name:"Beauty Store MM",shop_name_mm:"ဗျူတီ စတိုး",
  fb_link:"https://m.me/yourpage",viber_num:"+95912345678",
  wa_num:"+95912345678",phone_num:"+95912345678",admin_pw:"admin123",
  logo:"",theme:"light",language:"mm",
};
const SEED_PRODUCTS=[
  {id:"p1",name:"Argan Oil Hair Serum",name_mm:"အာဂန်ဆီ ဆံပင်ဆီ",category:"Hair Care",
   description:"Premium argan oil serum for silky hair.",suitable_for:"All hair types",
   benefits:"Smooths frizz, adds shine",usage_info:"Apply 2-3 drops to damp hair",warning:"",
   price:18000,discount_type:"percent",discount_value:15,bulk_discounts:[{min_qty:3,discount_percent:20}],
   stock:30,preorder:false,images:["https://images.unsplash.com/photo-1629470791566-3a8f2b05e3c5?w=400&h=400&fit=crop"],
   video_url:"",emoji:"💆",visible:true,created_at:new Date().toISOString()},
  {id:"p2",name:"Rose Lip Gloss Set",name_mm:"နှင်းဆီ နှုတ်ခမ်းဆေး စုံ",category:"Lips",
   description:"3-piece rose lip gloss set.",suitable_for:"All skin types",
   benefits:"Moisturizes lips",usage_info:"Apply directly to lips",warning:"",
   price:12000,discount_type:"percent",discount_value:20,bulk_discounts:[{min_qty:2,discount_percent:25}],
   stock:0,preorder:true,images:["https://images.unsplash.com/photo-1586495777744-4e6232bf2e53?w=400&h=400&fit=crop"],
   video_url:"",emoji:"💄",visible:true,created_at:new Date().toISOString()},
];

// ═══════════════════════════════════════════════════════════════
//  GLASS/THEME STYLES
// ═══════════════════════════════════════════════════════════════
const gls = (T) => T.glass ? {
  background: T.surface,
  backdropFilter: "blur(20px)",
  WebkitBackdropFilter: "blur(20px)",
  border: `1px solid ${T.border}`,
} : { background: T.surface };

const glsCard = (T, extra={}) => ({
  ...gls(T),
  borderRadius:16,
  boxShadow: T.shadow,
  border: `1px solid ${T.border}`,
  ...extra,
});

// ═══════════════════════════════════════════════════════════════
//  NOTIFICATION STORE
// ═══════════════════════════════════════════════════════════════
let _notifListeners = [];
const NotifStore = {
  items: [],
  add(msg, type="info") {
    const n = {id:uid(), msg, type, time:new Date().toISOString(), read:false};
    NotifStore.items = [n, ...NotifStore.items].slice(0,50);
    _notifListeners.forEach(fn=>fn([...NotifStore.items]));
    // Browser notification (best effort)
    if(typeof Notification!=="undefined" && Notification.permission==="granted")
      new Notification("Shop Update", {body:msg, icon:""});
  },
  subscribe(fn) { _notifListeners.push(fn); return ()=>{ _notifListeners=_notifListeners.filter(f=>f!==fn); }; },
  markRead() { NotifStore.items=NotifStore.items.map(n=>({...n,read:true})); _notifListeners.forEach(fn=>fn([...NotifStore.items])); },
  clear() { NotifStore.items=[]; _notifListeners.forEach(fn=>fn([])); },
};

function useNotifs() {
  const [notifs,setNotifs]=useState(NotifStore.items);
  useEffect(()=>NotifStore.subscribe(setNotifs),[]);
  return notifs;
}

// ═══════════════════════════════════════════════════════════════
//  COMPONENTS
// ═══════════════════════════════════════════════════════════════

function PImg({p,size=120,r=12}) {
  const [idx,setIdx]=useState(0);
  const [err,setErr]=useState({});
  let imgs=p.images; if(typeof imgs==="string"){try{imgs=JSON.parse(imgs);}catch(e){imgs=[];}} imgs=Array.isArray(imgs)?imgs:[];
  const fixedImgs=imgs.map(fixMediaUrl);
  const src=fixedImgs[idx]||"";
  const bgs=["#FFE0D5","#D5E8FF","#FFE5FF","#D5FFE8","#FFEDD5","#E5D5FF"];
  const bg=bgs[(p.id||"").charCodeAt(1)%bgs.length||0];
  if(src&&!err[idx]) return <img src={src} alt={p.name||""} onError={()=>setErr(e=>({...e,[idx]:true}))}
    style={{width:size,height:size,objectFit:"cover",borderRadius:r,flexShrink:0,display:"block"}}/>;
  return <div style={{width:size,height:size,borderRadius:r,background:bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.36,flexShrink:0}}>{p.emoji||"🛍️"}</div>;
}

function StatusBadge({status,t}) {
  const sc=STATUS_COLORS[status]||STATUS_COLORS.pending;
  const labels={pending:t.pendingStatus,confirmed:t.confirmedStatus,deposited:t.depositedStatus,completed:t.completedStatus,cancelled:t.cancelledStatus};
  return <span style={{background:sc.bg,color:sc.text,border:`1px solid ${sc.border}`,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{labels[status]||status}</span>;
}

function NotifBell({T,t}) {
  const notifs=useNotifs();
  const [open,setOpen]=useState(false);
  const unread=notifs.filter(n=>!n.read).length;
  const ref=useRef();
  useEffect(()=>{const fn=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener("mousedown",fn);return()=>document.removeEventListener("mousedown",fn);},[]);
  return (
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={()=>{setOpen(v=>!v);if(!open)NotifStore.markRead();}}
        style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:20,padding:"7px 11px",color:"white",cursor:"pointer",position:"relative",fontSize:16,display:"flex",alignItems:"center",gap:4}}>
        🔔{unread>0&&<span style={{position:"absolute",top:-4,right:-4,background:T.danger,color:"white",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:700,minWidth:14,textAlign:"center"}}>{unread}</span>}
      </button>
      {open&&(
        <div style={{position:"absolute",right:0,top:"calc(100% + 8px)",width:280,maxHeight:340,overflowY:"auto",borderRadius:14,boxShadow:T.shadowMd,zIndex:300,...gls(T),border:`1px solid ${T.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 14px",borderBottom:`1px solid ${T.border}`}}>
            <span style={{fontWeight:700,fontSize:13,color:T.text}}>{t.notif}</span>
            <button onClick={()=>NotifStore.clear()} style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:T.muted}}>{t.clearAll}</button>
          </div>
          {notifs.length===0
            ? <div style={{padding:20,textAlign:"center",color:T.muted,fontSize:13}}>No notifications</div>
            : notifs.map(n=>(
              <div key={n.id} style={{padding:"10px 14px",borderBottom:`1px solid ${T.border}`,opacity:n.read?0.6:1}}>
                <div style={{fontSize:12,fontWeight:600,color:T.text,marginBottom:2}}>{n.msg}</div>
                <div style={{fontSize:10,color:T.muted}}>{fdate(n.time)}</div>
              </div>
            ))}
        </div>
      )}
    </div>
  );
}

function Header({T,t,shopMM,logo,cartCount,searchQ,setSearchQ,showDrop,setShowDrop,dropRes,onHit,onCart,onAdmin,onLogo,page,onBack,onLangToggle,onThemeCycle,curThemeName}) {
  return (
    <div style={{background:T.glass?"rgba(255,87,51,0.85)":T.primary,position:"sticky",top:0,zIndex:100,backdropFilter:T.glass?"blur(20px)":undefined,boxShadow:T.shadowMd}}>
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 14px"}}>
        <button onClick={page!=="shop"?onBack:onLogo} style={{background:"none",border:"none",cursor:"pointer",color:"white",fontSize:page!=="shop"?22:16,padding:0,lineHeight:1,display:"flex",alignItems:"center",gap:4}}>
          {page!=="shop"?"←":(logo?<img src={logo} style={{width:32,height:32,borderRadius:8,objectFit:"cover"}}/>:"🛍️")}
        </button>
        <div style={{flex:1,fontWeight:800,fontSize:16,color:"white",letterSpacing:0.3,textShadow:"0 1px 4px rgba(0,0,0,0.15)"}}>{shopMM}</div>
        <button onClick={onLangToggle} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:12,padding:"5px 9px",color:"white",cursor:"pointer",fontSize:11,fontWeight:700}}>
          {t===TR.mm?"EN":"မြ"}
        </button>
        <button onClick={onThemeCycle} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:12,padding:"5px 9px",color:"white",cursor:"pointer",fontSize:11,fontWeight:700}}>
          {curThemeName==="light"?"🌙":curThemeName==="dark"?"🪩":"☀️"}
        </button>
        <NotifBell T={T} t={t}/>
        <button onClick={onCart} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:20,padding:"7px 12px",color:"white",cursor:"pointer",display:"flex",alignItems:"center",gap:5,fontSize:14,fontWeight:700}}>
          🛒{cartCount>0&&<span style={{background:T.accent,color:"#333",borderRadius:10,padding:"1px 6px",fontSize:11,fontWeight:800,minWidth:18,textAlign:"center"}}>{cartCount}</span>}
        </button>
        <button onClick={onAdmin} style={{background:"rgba(255,255,255,0.15)",border:"none",borderRadius:20,padding:"7px 11px",color:"white",cursor:"pointer",fontSize:13}}>⚙️</button>
      </div>
      <div style={{padding:"0 14px 11px",position:"relative"}}>
        <div style={{position:"relative"}}>
          <span style={{position:"absolute",left:14,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none"}}>🔍</span>
          <input value={searchQ} onChange={e=>{setSearchQ(e.target.value);setShowDrop(true);}}
            onFocus={()=>searchQ&&setShowDrop(true)} onBlur={()=>setTimeout(()=>setShowDrop(false),200)}
            placeholder={t.search}
            style={{width:"100%",padding:"10px 16px 10px 38px",borderRadius:24,border:"none",fontSize:13,outline:"none",boxSizing:"border-box",background:"rgba(255,255,255,0.95)",fontFamily:"inherit",boxShadow:"0 2px 8px rgba(0,0,0,0.15)"}}/>
        </div>
        {showDrop&&searchQ&&(
          <div style={{position:"absolute",top:"100%",left:14,right:14,borderRadius:12,boxShadow:T.shadowMd,zIndex:200,overflow:"hidden",maxHeight:300,overflowY:"auto",...gls(T)}}>
            {dropRes.length===0
              ? <div style={{padding:16,textAlign:"center",color:T.muted,fontSize:13}}>"{searchQ}" {t.noProducts}</div>
              : dropRes.map(p=>(
                <div key={p.id} onMouseDown={()=>onHit(p)}
                  style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",cursor:"pointer",borderBottom:`1px solid ${T.border}`}}>
                  <PImg p={p} size={40} r={8}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div>
                    <div style={{fontSize:11,color:T.muted}}>{p.category} · {fmt(calcUnit(p))}</div>
                  </div>
                </div>
              ))}
          </div>
        )}
      </div>
    </div>
  );
}

function ProductCard({T,t,p,onClick,onAdd}) {
  const [added,setAdded]=useState(false);
  const up=calcUnit(p);
  const canBuy=(p.stock||0)>0||p.preorder;
  const dv=p.discount_value||0;
  let bd=p.bulk_discounts; if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch(e){bd=[];}}
  const hasBulk=Array.isArray(bd)&&bd.length>0;
  const doAdd=e=>{e.stopPropagation();if(!canBuy)return;onAdd(p);setAdded(true);setTimeout(()=>setAdded(false),1800);};
  return (
    <div onClick={()=>onClick(p)}
      style={{...glsCard(T),overflow:"hidden",cursor:"pointer",transition:"all 0.2s",position:"relative"}}
      onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-4px)";e.currentTarget.style.boxShadow=T.shadowMd;}}
      onMouseLeave={e=>{e.currentTarget.style.transform="none";e.currentTarget.style.boxShadow=T.shadow;}}>
      {dv>0&&<div style={{position:"absolute",top:8,left:8,background:T.primary,color:"white",borderRadius:8,padding:"3px 8px",fontSize:10,fontWeight:800,zIndex:1,boxShadow:"0 2px 6px rgba(0,0,0,0.2)"}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</div>}
      {hasBulk&&<div style={{position:"absolute",top:dv>0?30:8,left:8,background:T.purple,color:"white",borderRadius:8,padding:"2px 6px",fontSize:9,fontWeight:800,zIndex:1}}>🎁 Bulk</div>}
      {p.preorder&&<div style={{position:"absolute",top:8,right:8,background:T.accent,color:"#333",borderRadius:8,padding:"2px 7px",fontSize:9,fontWeight:700,zIndex:1}}>Pre-order</div>}
      {!p.preorder&&(p.stock||0)===0&&<div style={{position:"absolute",top:8,right:8,background:T.muted,color:"white",borderRadius:8,padding:"2px 7px",fontSize:9,zIndex:1}}>{t.outOfStock}</div>}
      <div style={{background:T.glass?"rgba(255,255,255,0.1)":T.bg,display:"flex",justifyContent:"center",padding:"14px 14px 0"}}>
        <PImg p={p} size={112} r={12}/>
      </div>
      <div style={{padding:"10px 12px 13px"}}>
        <div style={{fontSize:10,color:T.muted,marginBottom:2,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{p.category}</div>
        <div style={{fontSize:13,fontWeight:700,color:T.text,lineHeight:1.35,marginBottom:7,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",minHeight:35}}>{p.name_mm||p.name}</div>
        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:10,flexWrap:"wrap"}}>
          <span style={{fontWeight:800,color:T.primary,fontSize:15}}>{fmt(up)}</span>
          {dv>0&&<span style={{fontSize:11,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
        </div>
        <button onClick={doAdd} disabled={!canBuy}
          style={{width:"100%",padding:"8px 0",borderRadius:10,border:"none",cursor:canBuy?"pointer":"not-allowed",
            background:added?T.success:(!canBuy?T.surface2:T.primary),
            color:!canBuy?T.muted:"white",fontSize:12,fontWeight:700,transition:"all 0.2s",
            boxShadow:canBuy&&!added?`0 3px 10px ${T.primary}40`:undefined}}>
          {added?`✓ ${t.addedCart}`:(p.preorder?`📋 ${t.preorder}`:((p.stock||0)===0?t.outOfStock:`🛒 ${t.addCart}`))}
        </button>
      </div>
    </div>
  );
}

function ShopPage({T,t,products,cats,catFilter,setCat,onOpen,onAdd}) {
  return (
    <div style={{paddingBottom:60}}>
      <div style={{display:"flex",gap:8,padding:"12px 14px",overflowX:"auto",background:T.surface,borderBottom:`1px solid ${T.border}`,scrollbarWidth:"none"}}>
        {[t.allCat,...cats.map(c=>c.name)].map(cat=>(
          <button key={cat} onClick={()=>setCat(cat)}
            style={{padding:"7px 16px",borderRadius:20,flexShrink:0,
              border:`1.5px solid ${catFilter===cat?T.primary:T.border}`,
              background:catFilter===cat?T.primary:T.glass?"rgba(255,255,255,0.1)":T.surface,
              color:catFilter===cat?"white":T.text,cursor:"pointer",fontSize:12,fontWeight:catFilter===cat?700:500,
              transition:"all 0.15s",boxShadow:catFilter===cat?`0 3px 10px ${T.primary}40`:undefined}}>
            {cat}
          </button>
        ))}
      </div>
      {products.length===0
        ? <div style={{textAlign:"center",padding:"80px 20px",color:T.muted}}><div style={{fontSize:52,marginBottom:12}}>🛍️</div><div style={{fontSize:16,fontWeight:600,color:T.text}}>{t.noProducts}</div></div>
        : <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12,padding:"14px 12px"}}>
            {products.map(p=><ProductCard key={p.id} T={T} t={t} p={p} onClick={onOpen} onAdd={onAdd}/>)}
          </div>}
    </div>
  );
}

function ProductPage({T,t,p,addToCart,onCartClick,cartCount}) {
  const [qty,setQty]=useState(1);
  const [added,setAdded]=useState(false);
  const [imgIdx,setImgIdx]=useState(0);
  const canBuy=(p.stock||0)>0||p.preorder;
  let imgs=p.images; if(typeof imgs==="string"){try{imgs=JSON.parse(imgs);}catch(e){imgs=[];}} imgs=Array.isArray(imgs)?imgs:[];
  const fixedImgs=imgs.map(fixMediaUrl);
  let bd=p.bulk_discounts; if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch(e){bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>a.min_qty-b.min_qty);
  const dv=p.discount_value||0;
  const Section=({title,txt,warn})=>txt?(<div style={{marginBottom:14}}>
    <div style={{fontSize:13,fontWeight:700,color:warn?T.danger:T.text,marginBottom:4}}>{title}</div>
    <div style={{fontSize:13,color:warn?"#742A2A":T.textSub,lineHeight:1.8,...glsCard(T,{padding:"12px 14px",background:warn?"#FFF5F5":T.surface2})}}>{txt}</div>
  </div>):null;
  return (
    <div style={{background:T.bg,minHeight:"100vh",paddingBottom:100}}>
      <div style={{background:T.glass?"rgba(255,255,255,0.05)":T.surface2,display:"flex",flexDirection:"column",alignItems:"center",padding:20}}>
        {fixedImgs.length>0&&!({[imgIdx]:true}[imgIdx]&&"")
          ? <img src={fixedImgs[imgIdx]||""} style={{width:220,height:220,objectFit:"cover",borderRadius:18,boxShadow:T.shadowMd}} onError={e=>e.target.style.display="none"}/>
          : <PImg p={p} size={220} r={18}/>}
        {fixedImgs.length>1&&(
          <div style={{display:"flex",gap:8,marginTop:12}}>
            {fixedImgs.map((img,i)=>(
              <div key={i} onClick={()=>setImgIdx(i)} style={{width:48,height:48,borderRadius:10,overflow:"hidden",border:`2px solid ${i===imgIdx?T.primary:T.border}`,cursor:"pointer",transition:"border 0.15s"}}>
                <img src={img} style={{width:"100%",height:"100%",objectFit:"cover"}}/>
              </div>
            ))}
          </div>
        )}
      </div>
      <div style={{padding:"20px 16px"}}>
        <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <span style={{background:T.tag,color:T.tagText,padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700}}>{p.category}</span>
          {p.preorder&&<span style={{background:"#FFFBEB",color:"#92600A",padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700}}>{t.preorder}</span>}
          {(p.stock||0)>0&&!p.preorder&&<span style={{background:"#F0FFF4",color:"#276749",padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:600}}>{t.stock}: {p.stock}</span>}
          {!p.preorder&&(p.stock||0)===0&&<span style={{background:T.surface2,color:T.muted,padding:"4px 12px",borderRadius:20,fontSize:12}}>{t.outOfStock}</span>}
        </div>
        <div style={{fontSize:22,fontWeight:800,color:T.text,marginBottom:3,lineHeight:1.3}}>{p.name_mm||p.name}</div>
        {p.name_mm&&<div style={{fontSize:14,color:T.muted,marginBottom:16}}>{p.name}</div>}
        <div style={{display:"flex",alignItems:"baseline",gap:12,marginBottom:bulks.length>0?10:20,flexWrap:"wrap"}}>
          <span style={{fontSize:28,fontWeight:800,color:T.primary}}>{fmt(calcUnit(p,qty))}</span>
          {dv>0&&<><span style={{fontSize:16,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>
            <span style={{background:T.primary,color:"white",borderRadius:8,padding:"3px 9px",fontSize:12,fontWeight:700}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</span></>}
        </div>
        {bulks.length>0&&(
          <div style={{...glsCard(T,{padding:"12px 14px",marginBottom:20,background:T.glass?"rgba(128,90,213,0.15)":"#F5F3FF",border:`1px solid ${T.purple}30`})}}>
            <div style={{fontSize:12,fontWeight:700,color:T.purple,marginBottom:8}}>🎁 {t.bulkDiscount}</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
              {bulks.map((b,i)=><div key={i} style={{background:qty>=b.min_qty?T.purple:"transparent",color:qty>=b.min_qty?"white":T.purple,border:`1.5px solid ${T.purple}`,borderRadius:8,padding:"4px 10px",fontSize:12,fontWeight:600,transition:"all 0.2s"}}>{b.min_qty}+ → -{b.discount_percent}%</div>)}
            </div>
          </div>
        )}
        {canBuy&&(
          <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:20}}>
            <div style={{display:"flex",alignItems:"center",gap:6,...glsCard(T,{borderRadius:28,padding:"5px 6px",display:"flex",alignItems:"center",gap:6})}}>
              <button onClick={()=>setQty(q=>Math.max(1,q-1))} style={{width:34,height:34,borderRadius:20,border:"none",background:T.surface2,cursor:"pointer",fontSize:18,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.text}}>−</button>
              <span style={{minWidth:30,textAlign:"center",fontWeight:800,fontSize:16,color:T.text}}>{qty}</span>
              <button onClick={()=>setQty(q=>q+1)} style={{width:34,height:34,borderRadius:20,border:"none",background:T.surface2,cursor:"pointer",fontSize:18,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.text}}>+</button>
            </div>
            <span style={{fontSize:13,color:T.muted}}>Total: <strong style={{color:T.primary,fontSize:15}}>{fmt(calcUnit(p,qty)*qty)}</strong></span>
          </div>
        )}
        <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);setAdded(true);setTimeout(()=>setAdded(false),2000);}} disabled={!canBuy}
          style={{width:"100%",padding:15,borderRadius:13,border:"none",cursor:canBuy?"pointer":"not-allowed",
            background:added?T.success:(!canBuy?T.surface2:T.primary),color:!canBuy?T.muted:"white",
            fontSize:16,fontWeight:700,marginBottom:12,transition:"all 0.25s",
            boxShadow:canBuy&&!added?`0 4px 16px ${T.primary}50`:undefined}}>
          {added?`✓ Added!`:(!canBuy?t.outOfStock:`🛒 ${t.addCart} (${qty})`)}
        </button>
        <button onClick={onCartClick} style={{width:"100%",padding:15,borderRadius:13,border:`2px solid ${T.primary}`,cursor:"pointer",background:"transparent",color:T.primary,fontSize:16,fontWeight:700}}>
          {t.viewCart} ({cartCount})
        </button>
        <div style={{marginTop:24}}>
          <Section title={`📝 ${t.description}`} txt={p.description}/>
          <Section title={`👤 ${t.suitableFor}`} txt={p.suitable_for}/>
          <Section title={`✅ ${t.benefits}`} txt={p.benefits}/>
          <Section title={`📋 ${t.usage}`} txt={p.usage_info}/>
          <Section title={`⚠️ ${t.warning}`} txt={p.warning} warn/>
        </div>
        {p.video_url&&<div style={{marginTop:8}}><div style={{fontSize:13,fontWeight:700,marginBottom:8,color:T.text}}>📹 Video</div><video src={fixMediaUrl(p.video_url)} controls style={{width:"100%",borderRadius:14,background:"#000"}}/></div>}
      </div>
    </div>
  );
}

function CartPage({T,t,cart,updateQty,removeItem,total,onOrder}) {
  if(!cart.length) return (
    <div style={{textAlign:"center",padding:"80px 20px",color:T.muted}}>
      <div style={{fontSize:60,marginBottom:16}}>🛒</div>
      <div style={{fontSize:18,fontWeight:700,marginBottom:6,color:T.text}}>{t.emptyCart}</div>
      <div style={{fontSize:14}}>{t.addProducts}</div>
    </div>
  );
  return (
    <div style={{background:T.bg,minHeight:"100vh",paddingBottom:100}}>
      <div style={{padding:"14px 14px 0",display:"flex",flexDirection:"column",gap:12}}>
        {cart.map(item=>{
          const up=calcUnit(item.p,item.qty);
          return (
            <div key={item.p.id} style={{...glsCard(T,{padding:13,display:"flex",gap:12,alignItems:"flex-start"})}}>
              <PImg p={item.p} size={68} r={10}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:2,lineHeight:1.35}}>{item.p.name_mm||item.p.name}</div>
                {item.p.preorder&&<span style={{fontSize:10,background:"#FFFBEB",color:"#92600A",padding:"1px 7px",borderRadius:8,fontWeight:700,display:"inline-block",marginBottom:4}}>{t.preorder}</span>}
                <div style={{fontSize:14,color:T.primary,fontWeight:800,marginBottom:10}}>{fmt(up)}</div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <button onClick={()=>updateQty(item.p.id,item.qty-1)} style={{width:30,height:30,borderRadius:15,border:`1px solid ${T.border}`,background:T.surface2,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.text,fontSize:16}}>−</button>
                  <span style={{minWidth:26,textAlign:"center",fontWeight:800,fontSize:15,color:T.text}}>{item.qty}</span>
                  <button onClick={()=>updateQty(item.p.id,item.qty+1)} style={{width:30,height:30,borderRadius:15,border:`1px solid ${T.border}`,background:T.surface2,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.text,fontSize:16}}>+</button>
                  <span style={{marginLeft:"auto",fontSize:14,fontWeight:700,color:T.text}}>{fmt(up*item.qty)}</span>
                </div>
              </div>
              <button onClick={()=>removeItem(item.p.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:T.muted,padding:0,lineHeight:1}}>✕</button>
            </div>
          );
        })}
      </div>
      <div style={{position:"sticky",bottom:0,...gls(T),padding:"14px 16px",boxShadow:`0 -4px 20px rgba(0,0,0,0.15)`,display:"flex",gap:12,alignItems:"center",marginTop:14,border:`1px solid ${T.border}`}}>
        <div style={{flex:1}}>
          <div style={{fontSize:12,color:T.muted}}>{t.total}</div>
          <div style={{fontSize:22,fontWeight:800,color:T.primary}}>{fmt(total)}</div>
        </div>
        <button onClick={onOrder} style={{padding:"14px 28px",borderRadius:13,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`}}>
          {t.orderNow} →
        </button>
      </div>
    </div>
  );
}

function OrderPage({T,t,cart,total,settings,onPlaced}) {
  const [note,setNote]=useState("");
  const [sent,setSent]=useState(false);
  const msgLines=cart.map(i=>`• ${i.p.name_mm||i.p.name}${i.p.preorder?" [PRE-ORDER]":""} x${i.qty} = ${fmt(calcUnit(i.p,i.qty)*i.qty)}`).join("\n");
  const hasPreorder=cart.some(i=>i.p.preorder);
  const msg=encodeURIComponent(`မင်္ဂလာပါ!\nOrder မှာချင်ပါသည်:\n${msgLines}\nစုစုပေါင်း: ${fmt(total)}${hasPreorder?"\n⚠️ Pre-order item(s) ပါဝင်သည်":""}${note?"\nNote: "+note:""}`);
  const saveAndPlace=async method=>{
    if(sent)return;
    const newOrd={id:uid(),order_number:onum(),
      items:cart.map(i=>({id:i.p.id,name:i.p.name,name_mm:i.p.name_mm,qty:i.qty,is_preorder:i.p.preorder,unit_price:calcUnit(i.p,i.qty),total:calcUnit(i.p,i.qty)*i.qty})),
      total,status:"pending",contact_method:method,customer_note:note,
      deposit_paid:0,balance_due:total,admin_note:"",created_at:new Date().toISOString()};
    const orders=(await ST.get("orders"))||[];
    await ST.set("orders",[newOrd,...orders]);
    setSent(true); onPlaced(newOrd);
  };
  const contacts=[
    {label:"Facebook",sub:"Messenger",color:"#1877F2",emoji:"💬",key:"messenger",url:`https://m.me/${(settings.fb_link||"").replace(/^https?:\/\/m\.me\//,"")}?text=${msg}`},
    {label:"Viber",sub:settings.viber_num,color:"#7360F2",emoji:"📱",key:"viber",url:`viber://chat?number=${encodeURIComponent(settings.viber_num||"")}&text=${msg}`},
    {label:"WhatsApp",sub:settings.wa_num,color:"#25D366",emoji:"💚",key:"whatsapp",url:`https://wa.me/${(settings.wa_num||"").replace(/\D/g,"")}?text=${msg}`},
    {label:"Phone",sub:settings.phone_num,color:"#4A5568",emoji:"📞",key:"phone",url:`tel:${settings.phone_num}`},
  ];
  return (
    <div style={{padding:16,background:T.bg,minHeight:"100vh",paddingBottom:40}}>
      <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
        <div style={{fontSize:16,fontWeight:800,marginBottom:14,color:T.text}}>📋 {t.orderSummary}</div>
        {hasPreorder&&<div style={{background:"#FFFBEB",border:"1px solid #F6E05E",borderRadius:10,padding:"10px 12px",marginBottom:12,fontSize:12,color:"#744210",fontWeight:600}}>
          ⚠️ Pre-order item(s) ပါဝင်သည် — Admin နှင့် timeline ညှိနှိုင်းပါ
        </div>}
        {cart.map(item=>(
          <div key={item.p.id} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${T.border}`,fontSize:13}}>
            <div style={{flex:1,paddingRight:12}}>
              <span style={{color:T.text,fontWeight:600}}>{item.p.name_mm||item.p.name}</span>
              {item.p.preorder&&<span style={{fontSize:10,background:"#FFFBEB",color:"#92600A",padding:"1px 5px",borderRadius:6,fontWeight:700,marginLeft:6}}>PRE</span>}
              <span style={{color:T.muted}}> × {item.qty}</span>
            </div>
            <span style={{fontWeight:700,color:T.text,whiteSpace:"nowrap"}}>{fmt(calcUnit(item.p,item.qty)*item.qty)}</span>
          </div>
        ))}
        <div style={{display:"flex",justifyContent:"space-between",padding:"14px 0 0",fontSize:18,fontWeight:800}}>
          <span style={{color:T.text}}>{t.total}</span><span style={{color:T.primary}}>{fmt(total)}</span>
        </div>
        <textarea value={note} onChange={e=>setNote(e.target.value)} placeholder={t.note}
          style={{width:"100%",marginTop:12,padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",resize:"vertical",minHeight:64,boxSizing:"border-box",fontFamily:"inherit",background:T.surface2,color:T.text}}/>
      </div>
      {sent
        ? <div style={{...glsCard(T,{padding:24,textAlign:"center"})}}>
            <div style={{fontSize:48,marginBottom:12}}>✅</div>
            <div style={{fontSize:16,fontWeight:700,color:T.success,marginBottom:4}}>Order ပေးပို့ပြီးပြီ!</div>
            <div style={{fontSize:13,color:T.muted}}>Admin မှ မကြာမီ confirm ပြန်ပါမည်</div>
          </div>
        : <div style={{...glsCard(T,{padding:20})}}>
            <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:4}}>📞 {t.contactOrder}</div>
            <div style={{fontSize:12,color:T.muted,marginBottom:16,lineHeight:1.7}}>{t.contactMsg}</div>
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              {contacts.map(ct=>(
                <a key={ct.key} href={ct.url} target="_blank" rel="noopener" onClick={()=>saveAndPlace(ct.key)}
                  style={{display:"flex",alignItems:"center",gap:14,padding:"15px 18px",background:ct.color,borderRadius:14,textDecoration:"none",color:"white",boxShadow:`0 4px 14px ${ct.color}60`,transition:"transform 0.15s"}}
                  onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"}
                  onMouseLeave={e=>e.currentTarget.style.transform="none"}>
                  <span style={{fontSize:24}}>{ct.emoji}</span>
                  <div><div style={{fontWeight:700,fontSize:14}}>{ct.label}</div><div style={{fontSize:12,opacity:0.85}}>{ct.sub}</div></div>
                  <span style={{marginLeft:"auto",fontSize:20,opacity:0.8}}>→</span>
                </a>
              ))}
            </div>
          </div>}
    </div>
  );
}

// ─── Admin Login ───────────────────────────────────────────────────────────────
function AdminLogin({T,t,adminPw,onSuccess,onBack}) {
  const [pw,setPw]=useState(""); const [err,setErr]=useState("");
  const check=()=>{if(pw===(adminPw||"admin123"))onSuccess();else setErr(t.wrongPassword);};
  return (
    <div style={{display:"flex",justifyContent:"center",alignItems:"center",minHeight:420,padding:24,background:T.bg}}>
      <div style={{...glsCard(T,{padding:"40px 28px",maxWidth:360,width:"100%"})}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{fontSize:56,marginBottom:10}}>🔐</div>
          <div style={{fontSize:22,fontWeight:800,color:T.text}}>{t.adminPanel}</div>
          <div style={{fontSize:13,color:T.muted,marginTop:5}}>Enter admin password</div>
        </div>
        <input type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&check()}
          placeholder="admin123"
          style={{width:"100%",padding:"13px 16px",borderRadius:12,border:`1.5px solid ${err?"#FC8181":T.border}`,fontSize:15,outline:"none",boxSizing:"border-box",marginBottom:err?8:16,fontFamily:"inherit",background:T.surface2,color:T.text}}/>
        {err&&<div style={{color:T.danger,fontSize:13,marginBottom:14,padding:"8px 12px",background:"#FFF5F5",borderRadius:8}}>{err}</div>}
        <button onClick={check} style={{width:"100%",padding:14,borderRadius:12,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:15,fontWeight:700,marginBottom:12,boxShadow:`0 4px 14px ${T.primary}50`}}>{t.login}</button>
        <button onClick={onBack} style={{width:"100%",padding:14,borderRadius:12,border:`1.5px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:14}}>{t.back}</button>
      </div>
    </div>
  );
}

// ─── Media Uploader ─────────────────────────────────────────────────────────────
function MediaUploader({T,t,images:imgsProp=[],videoUrl="",onImgs,onVid}) {
  const iRef=useRef(),vRef=useRef();
  const [urlIn,setUrlIn]=useState("");
  const images=Array.isArray(imgsProp)?imgsProp:[];
  const addUrl=()=>{ if(!urlIn.trim())return; onImgs([...images,urlIn.trim()]); setUrlIn(""); };
  const upImg=async e=>{const b64s=await Promise.all(Array.from(e.target.files||[]).map(fileToB64));onImgs([...images,...b64s]);e.target.value="";};
  const upVid=async e=>{const f=e.target.files?.[0];if(!f)return;onVid(await fileToB64(f));e.target.value="";};
  const sinp={flex:1,padding:"9px 12px",borderRadius:9,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,background:T.surface2,color:T.text};
  return (
    <div>
      <div style={{fontSize:12,color:T.muted,marginBottom:8,fontStyle:"italic"}}>💡 {t.gdriveTip}</div>
      <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:8}}>🖼️ {t.images}</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
        {images.map((img,i)=>(
          <div key={i} style={{position:"relative"}}>
            <img src={fixMediaUrl(img)} style={{width:60,height:60,objectFit:"cover",borderRadius:8,border:`1px solid ${T.border}`,display:"block"}} onError={e=>e.target.style.opacity="0.3"}/>
            <button onClick={()=>onImgs(images.filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:18,height:18,borderRadius:9,border:"none",background:T.danger,color:"white",cursor:"pointer",fontSize:10,lineHeight:"18px",textAlign:"center",padding:0}}>✕</button>
          </div>
        ))}
        <button onClick={()=>iRef.current?.click()} style={{width:60,height:60,borderRadius:8,border:`2px dashed ${T.border}`,background:T.surface2,cursor:"pointer",fontSize:22,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted}}>+</button>
      </div>
      <input ref={iRef} type="file" accept="image/*" multiple onChange={upImg} style={{display:"none"}}/>
      <div style={{display:"flex",gap:8,marginBottom:16}}>
        <input style={sinp} value={urlIn} onChange={e=>setUrlIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addUrl()} placeholder="Image URL / Google Drive link"/>
        <button onClick={addUrl} style={{padding:"9px 14px",borderRadius:9,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:13,fontWeight:600,flexShrink:0}}>{t.uploadImage}</button>
      </div>
      <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:8}}>🎬 Video</div>
      <div style={{display:"flex",gap:8}}>
        <input style={sinp} value={videoUrl} onChange={e=>onVid(e.target.value)} placeholder="Video URL / Google Drive link"/>
        <button onClick={()=>vRef.current?.click()} style={{padding:"9px 14px",borderRadius:9,border:`1.5px solid ${T.border}`,background:T.surface2,color:T.text,cursor:"pointer",fontSize:13,fontWeight:600,flexShrink:0}}>{t.uploadVideo}</button>
      </div>
      <input ref={vRef} type="file" accept="video/*" onChange={upVid} style={{display:"none"}}/>
    </div>
  );
}

// ─── BulkEditor ─────────────────────────────────────────────────────────────────
function BulkEditor({T,t,bulks=[],onChange}) {
  const add=()=>onChange([...bulks,{min_qty:2,discount_percent:10}]);
  const upd=(i,k,v)=>onChange(bulks.map((b,j)=>j===i?{...b,[k]:Number(v)}:b));
  const del=i=>onChange(bulks.filter((_,j)=>j!==i));
  return (
    <div style={{marginBottom:16}}>
      <div style={{fontSize:13,fontWeight:600,color:T.text,marginBottom:8}}>🎁 {t.bulkTiers}</div>
      {bulks.map((b,i)=>(
        <div key={i} style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
          <input type="number" min="2" value={b.min_qty} onChange={e=>upd(i,"min_qty",e.target.value)}
            style={{width:56,padding:"7px 8px",borderRadius:8,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",textAlign:"center",background:T.surface2,color:T.text}}/>
          <span style={{fontSize:12,color:T.muted}}>ခု+</span>
          <input type="number" min="1" max="99" value={b.discount_percent} onChange={e=>upd(i,"discount_percent",e.target.value)}
            style={{width:56,padding:"7px 8px",borderRadius:8,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",textAlign:"center",background:T.surface2,color:T.text}}/>
          <span style={{fontSize:12,color:T.muted}}>% {t.discount}</span>
          <button onClick={()=>del(i)} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",color:T.danger,fontSize:16}}>✕</button>
        </div>
      ))}
      <button onClick={add} style={{padding:"7px 14px",borderRadius:9,border:`1.5px dashed ${T.border}`,background:"transparent",cursor:"pointer",fontSize:13,color:T.primary,fontWeight:600}}>{t.addTier}</button>
    </div>
  );
}

// ─── Product Form ───────────────────────────────────────────────────────────────
function ProductForm({T,t,product,cats,onSave,onCancel}) {
  const blank={name:"",name_mm:"",category:cats[0]?.name||"",description:"",suitable_for:"",benefits:"",usage_info:"",warning:"",price:"",discount_type:"percent",discount_value:"0",bulk_discounts:[],stock:"0",preorder:false,images:[],video_url:"",emoji:"🛍️"};
  const [form,setForm]=useState(product?{...product,discount_value:product.discount_value??0,bulk_discounts:Array.isArray(product.bulk_discounts)?product.bulk_discounts:[]}:blank);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const handleSave=()=>{
    if(!form.name.trim()||!form.price){alert(`${t.productName.replace("*","")} & ${t.price.replace("*","")} required`);return;}
    onSave({...form,id:product?.id||uid(),price:Number(form.price),discount_value:Number(form.discount_value)||0,stock:Number(form.stock)||0,visible:product?.visible??true,created_at:product?.created_at||new Date().toISOString()});
  };
  const inp={width:"100%",padding:"10px 13px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:14,fontFamily:"inherit",background:T.surface2,color:T.text};
  const ta={...inp,minHeight:68,resize:"vertical"};
  const lbl={fontSize:12,fontWeight:700,color:T.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return (
    <div style={{...glsCard(T,{padding:20})}}>
      <div style={{fontSize:16,fontWeight:800,marginBottom:18,color:T.text}}>{product?`✏️ ${t.editProduct}`:`➕ ${t.addProduct}`}</div>
      <label style={lbl}>{t.productName}</label>
      <input style={inp} value={form.name} onChange={e=>set("name",e.target.value)} placeholder="e.g. Argan Oil Hair Serum"/>
      <label style={lbl}>{t.productNameMM}</label>
      <input style={inp} value={form.name_mm||""} onChange={e=>set("name_mm",e.target.value)} placeholder="e.g. အာဂန်ဆီ ဆံပင်ဆီ (optional)"/>
      <label style={lbl}>{t.category}</label>
      {cats.length===0
        ? <div style={{padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.danger}50`,background:"#FFF5F5",fontSize:13,color:T.danger,marginBottom:14}}>⚠️ Add categories first</div>
        : <select style={{...inp,background:T.surface2}} value={form.category} onChange={e=>set("category",e.target.value)}>
            {cats.map(c=><option key={c.id} value={c.name}>{c.name}</option>)}
          </select>}
      <label style={lbl}>Emoji</label>
      <input style={inp} value={form.emoji||""} onChange={e=>set("emoji",e.target.value)} placeholder="🛍️"/>
      <label style={lbl}>📝 {t.description}</label><textarea style={ta} value={form.description||""} onChange={e=>set("description",e.target.value)}/>
      <label style={lbl}>👤 {t.suitableFor}</label><textarea style={{...ta,minHeight:52}} value={form.suitable_for||""} onChange={e=>set("suitable_for",e.target.value)}/>
      <label style={lbl}>✅ {t.benefits}</label><textarea style={{...ta,minHeight:52}} value={form.benefits||""} onChange={e=>set("benefits",e.target.value)}/>
      <label style={lbl}>📋 {t.usage}</label><textarea style={{...ta,minHeight:52}} value={form.usage_info||""} onChange={e=>set("usage_info",e.target.value)}/>
      <label style={lbl}>⚠️ {t.warning}</label><textarea style={{...ta,minHeight:48}} value={form.warning||""} onChange={e=>set("warning",e.target.value)}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
        <div><label style={lbl}>{t.price}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.price} onChange={e=>set("price",e.target.value)}/></div>
        <div><label style={lbl}>{t.discountType}</label>
          <select style={{...inp,marginBottom:0,background:T.surface2}} value={form.discount_type} onChange={e=>set("discount_type",e.target.value)}>
            <option value="percent">{t.percent}</option>
            <option value="fixed">{t.fixedKs}</option>
          </select>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
        <div><label style={lbl}>{t.discountValue} ({form.discount_type==="fixed"?"Ks":"%"})</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.discount_value} onChange={e=>set("discount_value",e.target.value)}/></div>
        <div><label style={lbl}>{t.stockQty}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.stock} onChange={e=>set("stock",e.target.value)}/></div>
      </div>
      <label style={{...lbl,display:"flex",alignItems:"center",gap:8,cursor:"pointer",userSelect:"none",marginBottom:16,textTransform:"none"}}>
        <input type="checkbox" checked={form.preorder||false} onChange={e=>set("preorder",e.target.checked)} style={{width:16,height:16,accentColor:T.primary}}/>
        {t.preorderEnable}
      </label>
      <BulkEditor T={T} t={t} bulks={Array.isArray(form.bulk_discounts)?form.bulk_discounts:[]} onChange={v=>set("bulk_discounts",v)}/>
      <div style={{marginBottom:20}}><MediaUploader T={T} t={t} images={Array.isArray(form.images)?form.images:[]} videoUrl={form.video_url||""} onImgs={v=>set("images",v)} onVid={v=>set("video_url",v)}/></div>
      <div style={{display:"flex",gap:12}}>
        <button onClick={handleSave} style={{flex:1,padding:14,borderRadius:12,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 14px ${T.primary}50`}}>{product?`📝 ${t.saveProduct}`:`➕ ${t.addProduct}`}</button>
        <button onClick={onCancel} style={{flex:1,padding:14,borderRadius:12,border:`1.5px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:14}}>{t.cancel}</button>
      </div>
    </div>
  );
}

// ─── Order Manager ──────────────────────────────────────────────────────────────
function OrderManager({T,t,orders,setOrders}) {
  const [sel,setSel]=useState(null);
  const [depInput,setDepInput]=useState("");
  const [noteInput,setNoteInput]=useState("");

  const updateOrder=(id,fields)=>{
    const updated=orders.map(o=>o.id===id?{...o,...fields}:o);
    setOrders(updated); ST.set("orders",updated);
    if(sel?.id===id) setSel(o=>({...o,...fields}));
  };

  const statuses=["pending","confirmed","deposited","completed","cancelled"];

  if(sel) {
    const dep=Number(sel.deposit_paid||0);
    const bal=sel.total-dep;
    return (
      <div style={{padding:16,paddingBottom:40}}>
        <button onClick={()=>setSel(null)} style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:13,fontWeight:700,marginBottom:16,padding:0}}>← Back to Orders</button>
        <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div><div style={{fontSize:16,fontWeight:800,color:T.text}}>{sel.order_number}</div><div style={{fontSize:12,color:T.muted}}>{fdate(sel.created_at)}</div></div>
            <StatusBadge status={sel.status} t={t}/>
          </div>
          {sel.items?.some(i=>i.is_preorder)&&<div style={{background:"#FFFBEB",border:"1px solid #F6E05E",borderRadius:10,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#744210",fontWeight:600}}>⚠️ {t.isPreorder}</div>}
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:8}}>{t.orderItems}</div>
          {(sel.items||[]).map((item,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`,fontSize:13}}>
              <div><span style={{color:T.text,fontWeight:600}}>{item.name_mm||item.name}</span>{item.is_preorder&&<span style={{fontSize:10,background:"#FFFBEB",color:"#92600A",padding:"1px 5px",borderRadius:6,fontWeight:700,marginLeft:4}}>PRE</span>}<span style={{color:T.muted}}> × {item.qty}</span></div>
              <span style={{fontWeight:700,color:T.text}}>{fmt(item.total)}</span>
            </div>
          ))}
          <div style={{marginTop:12,padding:"12px 0",borderTop:`2px solid ${T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:T.text}}>{t.orderTotal}</span><span style={{color:T.primary}}>{fmt(sel.total)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:4}}><span style={{color:T.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:T.success,fontWeight:700}}>{fmt(dep)}</span></div>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:13}}><span style={{color:T.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:T.danger,fontWeight:700}}>{fmt(Math.max(0,bal))}</span></div>
          </div>
        </div>
        {sel.contact_method&&<div style={{...glsCard(T,{padding:"10px 14px",marginBottom:14,display:"flex",gap:8,alignItems:"center"})}}>
          <span style={{fontSize:12,color:T.muted}}>{t.contactMethod}:</span><span style={{fontSize:13,fontWeight:600,color:T.text,textTransform:"capitalize"}}>{sel.contact_method}</span>
        </div>}
        {sel.customer_note&&<div style={{...glsCard(T,{padding:"10px 14px",marginBottom:14})}}>
          <div style={{fontSize:11,color:T.muted,marginBottom:4}}>{t.customerNote}</div>
          <div style={{fontSize:13,color:T.text}}>{sel.customer_note}</div>
        </div>}
        <div style={{...glsCard(T,{padding:16,marginBottom:14})}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.updateStatus}</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
            {statuses.map(s=><button key={s} onClick={()=>updateOrder(sel.id,{status:s})}
              style={{padding:"6px 12px",borderRadius:20,border:`1.5px solid ${sel.status===s?T.primary:T.border}`,background:sel.status===s?T.primary:"transparent",color:sel.status===s?"white":T.text,cursor:"pointer",fontSize:12,fontWeight:600}}>
              <StatusBadge status={s} t={t}/>
            </button>)}
          </div>
        </div>
        <div style={{...glsCard(T,{padding:16,marginBottom:14})}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.updateDeposit}</div>
          <div style={{display:"flex",gap:8}}>
            <input type="number" min="0" value={depInput} onChange={e=>setDepInput(e.target.value)} placeholder="0"
              style={{flex:1,padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",background:T.surface2,color:T.text,fontFamily:"inherit",minWidth:0}}/>
            <button onClick={()=>{const d=Number(depInput);updateOrder(sel.id,{deposit_paid:d,balance_due:Math.max(0,sel.total-d)});setDepInput("");}}
              style={{padding:"10px 16px",borderRadius:10,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0}}>Update</button>
          </div>
        </div>
        <div style={{...glsCard(T,{padding:16})}}>
          <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10}}>{t.adminNote}</div>
          <textarea value={noteInput||sel.admin_note||""} onChange={e=>setNoteInput(e.target.value)}
            style={{width:"100%",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:13,outline:"none",resize:"vertical",minHeight:80,boxSizing:"border-box",fontFamily:"inherit",background:T.surface2,color:T.text,marginBottom:10}}/>
          <button onClick={()=>{updateOrder(sel.id,{admin_note:noteInput||sel.admin_note||""});}}
            style={{width:"100%",padding:10,borderRadius:10,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:14,fontWeight:700}}>Save Note</button>
        </div>
      </div>
    );
  }

  const statusCount={};orders.forEach(o=>{statusCount[o.status]=(statusCount[o.status]||0)+1;});
  return (
    <div style={{padding:16,paddingBottom:40}}>
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
        {Object.entries(statusCount).map(([s,n])=>(
          <div key={s} style={{...glsCard(T,{padding:"8px 14px",display:"flex",gap:8,alignItems:"center"})}}>
            <StatusBadge status={s} t={t}/><span style={{fontSize:13,fontWeight:700,color:T.text}}>{n}</span>
          </div>
        ))}
      </div>
      {orders.length===0
        ? <div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}><div style={{fontSize:48,marginBottom:12}}>📦</div><div style={{fontSize:15}}>No orders yet</div></div>
        : orders.map(o=>(
          <div key={o.id} onClick={()=>{setSel(o);setDepInput("");setNoteInput(o.admin_note||"");}}
            style={{...glsCard(T,{padding:"13px 14px",marginBottom:10,cursor:"pointer",transition:"all 0.15s"})}}
            onMouseEnter={e=>e.currentTarget.style.transform="translateX(3px)"}
            onMouseLeave={e=>e.currentTarget.style.transform="none"}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:T.text}}>{o.order_number}</div>
                <div style={{fontSize:11,color:T.muted}}>{fdate(o.created_at)} · {o.contact_method}</div>
              </div>
              <StatusBadge status={o.status} t={t}/>
            </div>
            {o.items?.some(i=>i.is_preorder)&&<div style={{fontSize:10,color:"#92600A",fontWeight:700,marginBottom:4}}>⚠️ PREORDER</div>}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:4}}>
              <div style={{fontSize:12,color:T.muted}}>{(o.items||[]).length} items</div>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:14,fontWeight:800,color:T.primary}}>{fmt(o.total)}</div>
                {(o.deposit_paid||0)>0&&<div style={{fontSize:11,color:T.success}}>Dep: {fmt(o.deposit_paid)} · Bal: {fmt(Math.max(0,o.total-(o.deposit_paid||0)))}</div>}
              </div>
            </div>
          </div>
        ))}
    </div>
  );
}

// ─── Reports ────────────────────────────────────────────────────────────────────
function Reports({T,t,orders}) {
  const [period,setPeriod]=useState("week");
  const filtered=useMemo(()=>{
    const now=new Date();
    return orders.filter(o=>{
      const d=new Date(o.created_at);
      if(period==="today") return d.toDateString()===now.toDateString();
      if(period==="week"){const w=new Date(now);w.setDate(w.getDate()-7);return d>=w;}
      if(period==="month"){const m=new Date(now);m.setDate(m.getDate()-30);return d>=m;}
      return true;
    });
  },[orders,period]);

  const totalRev=filtered.filter(o=>o.status!=="cancelled").reduce((s,o)=>s+Number(o.total),0);
  const totalDep=filtered.filter(o=>o.status!=="cancelled").reduce((s,o)=>s+(Number(o.deposit_paid)||0),0);
  const totalBal=totalRev-totalDep;
  const topProds=useMemo(()=>{
    const map={};
    filtered.filter(o=>o.status!=="cancelled").forEach(o=>(o.items||[]).forEach(item=>{
      const k=item.name||"?";
      if(!map[k]) map[k]={name:item.name_mm||item.name||k,qty:0,revenue:0};
      map[k].qty+=(item.qty||0); map[k].revenue+=item.total||((item.unit_price||0)*(item.qty||0));
    }));
    return Object.values(map).sort((a,b)=>b.revenue-a.revenue).slice(0,5);
  },[filtered]);

  const exportExcel=()=>{
    const wb=XLSX.utils.book_new();
    // Orders sheet
    const ordRows=[["Order #","Date","Status","Items","Total","Deposit","Balance","Contact","Customer Note","Admin Note"]];
    filtered.forEach(o=>ordRows.push([o.order_number,fdate(o.created_at),o.status,(o.items||[]).map(i=>`${i.name_mm||i.name}x${i.qty}`).join(", "),o.total,o.deposit_paid||0,Math.max(0,o.total-(o.deposit_paid||0)),o.contact_method,o.customer_note||"",o.admin_note||""]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(ordRows),"Orders");
    // Summary sheet
    const sumRows=[["Metric","Value"],["Period",period],["Total Revenue",totalRev],["Total Deposit Received",totalDep],["Total Balance Due",totalBal],["Total Orders",filtered.length],["Completed",filtered.filter(o=>o.status==="completed").length],["Pending",filtered.filter(o=>o.status==="pending").length],["Cancelled",filtered.filter(o=>o.status==="cancelled").length]];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sumRows),"Summary");
    // Top products sheet
    const prodRows=[["Product","Qty Sold","Revenue"],...topProds.map(p=>[p.name,p.qty,p.revenue])];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(prodRows),"Top Products");
    XLSX.writeFile(wb,`ShopReport_${period}_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const periods=[{k:"today",l:t.today},{k:"week",l:t.thisWeek},{k:"month",l:t.thisMonth},{k:"year",l:t.allTime}];
  const Stat=({emoji,label,value,sub,color})=>(
    <div style={{...glsCard(T,{padding:14,flex:1,minWidth:80})}}>
      <div style={{fontSize:18,marginBottom:3}}>{emoji}</div>
      <div style={{fontSize:11,color:T.muted,marginBottom:2,fontWeight:600}}>{label}</div>
      <div style={{fontSize:14,fontWeight:800,color:color||T.text}}>{value}</div>
      {sub&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>{sub}</div>}
    </div>
  );

  return (
    <div style={{padding:16,paddingBottom:40}}>
      <div style={{display:"flex",gap:6,marginBottom:16,...glsCard(T,{padding:6})}}>
        {periods.map(({k,l})=>(
          <button key={k} onClick={()=>setPeriod(k)}
            style={{flex:1,padding:"8px 4px",borderRadius:9,border:"none",background:period===k?T.primary:"transparent",color:period===k?"white":T.muted,cursor:"pointer",fontSize:12,fontWeight:period===k?700:500,transition:"all 0.15s"}}>
            {l}
          </button>
        ))}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <Stat emoji="💰" label={t.revenue} value={fmt(totalRev)} color={T.primary}/>
        <Stat emoji="📦" label={t.totalOrders} value={filtered.length}/>
        <Stat emoji="✅" label="Deposited" value={fmt(totalDep)} color={T.success}/>
        <Stat emoji="⏳" label="Balance Due" value={fmt(totalBal)} color={T.danger}/>
      </div>
      {topProds.length>0&&(
        <div style={{...glsCard(T,{padding:16,marginBottom:14})}}>
          <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:12}}>🏆 {t.topProducts}</div>
          {topProds.map((p,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:i<topProds.length-1?`1px solid ${T.border}`:"none"}}>
              <span style={{width:22,height:22,borderRadius:11,background:["#FFB347","#C0C0C0","#CD7F32","#A0AEC0","#A0AEC0"][i],display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"white",flexShrink:0}}>{i+1}</span>
              <span style={{flex:1,fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span>
              <span style={{fontSize:11,color:T.muted,marginRight:6}}>{p.qty}ခု</span>
              <span style={{fontSize:13,fontWeight:700,color:T.primary,whiteSpace:"nowrap"}}>{fmt(p.revenue)}</span>
            </div>
          ))}
        </div>
      )}
      <div style={{...glsCard(T,{padding:16,marginBottom:14})}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:14,fontWeight:700,color:T.text}}>{t.recentOrders}</div>
          <button onClick={exportExcel} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",borderRadius:10,border:"none",background:"#21A366",color:"white",cursor:"pointer",fontSize:12,fontWeight:700,boxShadow:"0 3px 10px rgba(33,163,102,0.4)"}}>
            📊 {t.exportExcel}
          </button>
        </div>
        {filtered.slice(0,10).map(o=>(
          <div key={o.id} style={{padding:"10px 0",borderBottom:`1px solid ${T.border}`,fontSize:12}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
              <span style={{fontWeight:700,color:T.text}}>{o.order_number}</span>
              <StatusBadge status={o.status} t={t}/>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",color:T.muted}}>
              <span>{fdate(o.created_at)} · {o.contact_method}</span>
              <div style={{textAlign:"right"}}><span style={{fontWeight:700,color:T.primary}}>{fmt(o.total)}</span>{(o.deposit_paid||0)>0&&<div style={{fontSize:10,color:T.success}}>Dep:{fmt(o.deposit_paid)}</div>}</div>
            </div>
          </div>
        ))}
        {filtered.length===0&&<div style={{textAlign:"center",color:T.muted,padding:20,fontSize:13}}>{t.noData}</div>}
      </div>
    </div>
  );
}

// ─── Admin Settings ──────────────────────────────────────────────────────────────
function AdminSettings({T,t,settings,onSave,onThemeChange,onLangChange,curTheme,curLang}) {
  const [form,setForm]=useState({...settings});
  const [logoPreview,setLogoPreview]=useState(settings.logo||"");
  const logoRef=useRef();
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const upLogo=async e=>{const f=e.target.files?.[0];if(!f)return;const b64=await fileToB64(f);setLogoPreview(b64);set("logo",b64);e.target.value="";};
  const inp={width:"100%",padding:"11px 14px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:14,fontFamily:"inherit",background:T.surface2,color:T.text};
  const lbl={fontSize:12,fontWeight:700,color:T.textSub,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return (
    <div style={{padding:16,paddingBottom:40}}>
      <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>🎨 {t.theme} & {t.language}</div>
        <div style={{display:"flex",gap:8,marginBottom:14}}>
          {Object.keys(THEMES).map(k=>(
            <button key={k} onClick={()=>onThemeChange(k)}
              style={{flex:1,padding:10,borderRadius:10,border:`2px solid ${curTheme===k?T.primary:T.border}`,background:curTheme===k?T.primary:"transparent",color:curTheme===k?"white":T.text,cursor:"pointer",fontSize:12,fontWeight:700,transition:"all 0.15s"}}>
              {k==="light"?"☀️ Light":k==="dark"?"🌙 Dark":"🪩 Glass"}
            </button>
          ))}
        </div>
        <div style={{display:"flex",gap:8}}>
          {["mm","en"].map(l=>(
            <button key={l} onClick={()=>onLangChange(l)}
              style={{flex:1,padding:10,borderRadius:10,border:`2px solid ${curLang===l?T.primary:T.border}`,background:curLang===l?T.primary:"transparent",color:curLang===l?"white":T.text,cursor:"pointer",fontSize:13,fontWeight:700,transition:"all 0.15s"}}>
              {l==="mm"?"🇲🇲 မြန်မာ":"🇬🇧 English"}
            </button>
          ))}
        </div>
      </div>
      <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>🖼️ {t.logoUpload}</div>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}>
          <div style={{width:72,height:72,borderRadius:14,border:`2px dashed ${T.border}`,display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",background:T.surface2,cursor:"pointer",flexShrink:0}} onClick={()=>logoRef.current?.click()}>
            {logoPreview?<img src={logoPreview} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:28}}>🛍️</span>}
          </div>
          <div>
            <button onClick={()=>logoRef.current?.click()} style={{padding:"9px 16px",borderRadius:10,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:13,fontWeight:700,display:"block",marginBottom:6}}>Upload Logo</button>
            {logoPreview&&<button onClick={()=>{setLogoPreview("");set("logo","");}} style={{padding:"6px 12px",borderRadius:8,border:`1px solid ${T.border}`,background:"transparent",color:T.danger,cursor:"pointer",fontSize:12}}>Remove</button>}
          </div>
        </div>
        <input ref={logoRef} type="file" accept="image/*" onChange={upLogo} style={{display:"none"}}/>
        <label style={lbl}>{t.shopSettings} (English)</label><input style={inp} value={form.shop_name||""} onChange={e=>set("shop_name",e.target.value)}/>
        <label style={lbl}>{t.shopSettings} (Myanmar)</label><input style={inp} value={form.shop_name_mm||""} onChange={e=>set("shop_name_mm",e.target.value)}/>
      </div>
      <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>📞 {t.contactLinks}</div>
        <label style={lbl}>Facebook Messenger URL</label><input style={inp} value={form.fb_link||""} onChange={e=>set("fb_link",e.target.value)} placeholder="https://m.me/yourpage"/>
        <label style={lbl}>Viber Number (+95...)</label><input style={inp} value={form.viber_num||""} onChange={e=>set("viber_num",e.target.value)} placeholder="+95912345678"/>
        <label style={lbl}>WhatsApp Number (+95...)</label><input style={inp} value={form.wa_num||""} onChange={e=>set("wa_num",e.target.value)} placeholder="+95912345678"/>
        <label style={lbl}>Phone Number</label><input style={inp} value={form.phone_num||""} onChange={e=>set("phone_num",e.target.value)} placeholder="+95912345678"/>
      </div>
      <div style={{...glsCard(T,{padding:20,marginBottom:14})}}>
        <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:14}}>🔐 {t.adminPassword}</div>
        <input style={inp} type="password" value={form.admin_pw||""} onChange={e=>set("admin_pw",e.target.value)}/>
      </div>
      <button onClick={()=>onSave({...form,logo:logoPreview})}
        style={{width:"100%",padding:14,borderRadius:12,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 14px ${T.primary}50`}}>
        💾 {t.save}
      </button>
    </div>
  );
}

// ─── Admin Panel ─────────────────────────────────────────────────────────────────
function AdminPanel({T,t,products,cats,settings,orders,onProdChange,onCatsChange,onSettingsChange,onBack,onThemeChange,onLangChange,curTheme,curLang}) {
  const [tab,setTab]=useState("products");
  const [showForm,setShowForm]=useState(false);
  const [editing,setEditing]=useState(null);
  const [orders_,setOrders_]=useState(orders);
  useEffect(()=>setOrders_(orders),[orders]);

  const tabs=[{k:"products",l:`📦 ${t.products}`},{k:"orders",l:`🛍️ ${t.orders}`},{k:"cats",l:`🏷️ ${t.categories}`},{k:"reports",l:`📊 ${t.reports}`},{k:"settings",l:`⚙️ ${t.settings}`}];

  const delProd=p=>{if(!window.confirm(`Delete "${p.name_mm||p.name}"?`))return;onProdChange(products.filter(x=>x.id!==p.id));};
  const toggleVis=p=>{const u=products.map(x=>x.id===p.id?{...x,visible:!x.visible}:x);onProdChange(u);NotifStore.add(`${p.name_mm||p.name} — ${p.visible?"hidden":"visible"}`);};
  const addCat=(name)=>{if(cats.find(c=>c.name===name))return;onCatsChange([...cats,{id:uid(),name,sort_order:cats.length+1}]);};
  const delCat=(cat)=>{if(!window.confirm(`Delete "${cat.name}"?`))return;onCatsChange(cats.filter(c=>c.id!==cat.id));};

  if(showForm) return (
    <div style={{background:T.bg,minHeight:"100vh",padding:16}}>
      <button onClick={()=>{setShowForm(false);setEditing(null);}} style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:13,fontWeight:700,marginBottom:16,padding:0}}>← {t.cancel}</button>
      <ProductForm T={T} t={t} product={editing} cats={cats}
        onSave={saved=>{
          const isNew=!editing;
          if(editing) onProdChange(products.map(x=>x.id===saved.id?saved:x));
          else onProdChange([saved,...products]);
          NotifStore.add(isNew?`${t.newProductAdded}: ${saved.name_mm||saved.name}`:`${t.productUpdated}: ${saved.name_mm||saved.name}`);
          // Request browser notification permission
          if(typeof Notification!=="undefined"&&Notification.permission==="default") Notification.requestPermission();
          setShowForm(false);setEditing(null);
        }}
        onCancel={()=>{setShowForm(false);setEditing(null);}}/>
    </div>
  );

  return (
    <div style={{background:T.bg,minHeight:"100vh",paddingBottom:40}}>
      <div style={{display:"flex",background:T.surface,...(T.glass?{backdropFilter:"blur(20px)",WebkitBackdropFilter:"blur(20px)"}:{}),borderBottom:`1px solid ${T.border}`,overflowX:"auto",scrollbarWidth:"none"}}>
        {tabs.map(tab_=>(
          <button key={tab_.k} onClick={()=>setTab(tab_.k)}
            style={{flex:"none",padding:"13px 12px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:tab===tab_.k?700:500,color:tab===tab_.k?T.primary:T.muted,borderBottom:`2.5px solid ${tab===tab_.k?T.primary:"transparent"}`,whiteSpace:"nowrap",transition:"all 0.15s"}}>
            {tab_.l}
          </button>
        ))}
      </div>

      {tab==="products"&&(
        <div style={{padding:16}}>
          <button onClick={()=>{setEditing(null);setShowForm(true);}}
            style={{width:"100%",padding:13,borderRadius:12,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:15,fontWeight:700,marginBottom:16,boxShadow:`0 4px 14px ${T.primary}50`}}>
            ➕ {t.addProduct}
          </button>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {products.map(p=>(
              <div key={p.id} style={{...glsCard(T,{padding:12,display:"flex",gap:12,alignItems:"center",opacity:p.visible?1:0.5})}}>
                <PImg p={p} size={56} r={10}/>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:700,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div>
                  <div style={{fontSize:11,color:T.muted,marginTop:2}}>{p.category} · {fmt(calcUnit(p))}</div>
                  <div style={{fontSize:10,marginTop:2,color:p.visible?T.success:T.muted,fontWeight:600}}>{p.visible?"🟢 Visible":"⚫ Hidden"} · Stock {p.stock||0}{p.preorder?" · Pre-order":""}</div>
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <button onClick={()=>{setEditing(p);setShowForm(true);}} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${T.border}`,background:"transparent",cursor:"pointer",fontSize:13,color:T.text}}>✏️</button>
                  <button onClick={()=>toggleVis(p)} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${T.border}`,background:"transparent",cursor:"pointer",fontSize:13,color:T.text}}>{p.visible?"🙈":"👁️"}</button>
                  <button onClick={()=>delProd(p)} style={{padding:"6px 10px",borderRadius:8,border:`1px solid ${T.danger}50`,background:"#FFF5F5",cursor:"pointer",fontSize:13,color:T.danger}}>🗑️</button>
                </div>
              </div>
            ))}
            {products.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:T.muted,fontSize:14}}>{t.noProducts}</div>}
          </div>
        </div>
      )}

      {tab==="orders"&&<OrderManager T={T} t={t} orders={orders_} setOrders={o=>{setOrders_(o);}}/>}

      {tab==="cats"&&(
        <div style={{padding:16}}>
          <div style={{...glsCard(T,{padding:20})}}>
            <div style={{fontSize:14,fontWeight:700,color:T.text,marginBottom:16}}>🏷️ {t.categories}</div>
            <CatInput T={T} t={t} onAdd={addCat}/>
            {cats.map(cat=>(
              <div key={cat.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"11px 14px",...glsCard(T,{marginBottom:8,borderRadius:10,background:T.surface2})}}>
                <span style={{fontSize:14,fontWeight:600,color:T.text}}>{cat.name}</span>
                <button onClick={()=>delCat(cat)} style={{background:"none",border:`1px solid ${T.danger}50`,borderRadius:8,cursor:"pointer",color:T.danger,fontSize:12,padding:"4px 10px"}}>🗑️</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {tab==="reports"&&<Reports T={T} t={t} orders={orders_}/>}
      {tab==="settings"&&<AdminSettings T={T} t={t} settings={settings} onSave={onSettingsChange} onThemeChange={onThemeChange} onLangChange={onLangChange} curTheme={curTheme} curLang={curLang}/>}

      <div style={{padding:"0 16px"}}>
        <button onClick={onBack} style={{width:"100%",padding:13,borderRadius:12,border:`1.5px solid ${T.border}`,background:"transparent",color:T.muted,cursor:"pointer",fontSize:14,marginTop:12}}>← {t.logout}</button>
      </div>
    </div>
  );
}

function CatInput({T,t,onAdd}) {
  const [v,setV]=useState("");
  return (
    <div style={{display:"flex",gap:8,marginBottom:16}}>
      <input value={v} onChange={e=>setV(e.target.value)} onKeyDown={e=>e.key==="Enter"&&v.trim()&&(onAdd(v.trim()),setV(""))}
        placeholder="New category name"
        style={{flex:1,padding:"11px 14px",borderRadius:10,border:`1.5px solid ${T.border}`,fontSize:14,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,background:T.surface2,color:T.text}}/>
      <button onClick={()=>v.trim()&&(onAdd(v.trim()),setV(""))} style={{padding:"11px 18px",borderRadius:10,border:"none",background:T.primary,color:"white",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0}}>+</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  MAIN APP
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [ready,setReady]=useState(false);
  const [products,setProducts]=useState([]);
  const [cats,setCats]=useState([]);
  const [settings,setSettings]=useState(SEED_SETTINGS);
  const [orders,setOrders]=useState([]);
  const [cart,setCart]=useState([]);
  const [page,setPage]=useState("shop");
  const [selProd,setSelProd]=useState(null);
  const [catFilter,setCatFilter]=useState("all");
  const [searchQ,setSearchQ]=useState("");
  const [showDrop,setShowDrop]=useState(false);
  const [adminIn,setAdminIn]=useState(false);
  const [themeName,setThemeName]=useState("light");
  const [lang,setLang]=useState("mm");

  const T=THEMES[themeName]||THEMES.light;
  const t=TR[lang]||TR.mm;

  useEffect(()=>{
    (async()=>{
      const p=await ST.get("products"); const c=await ST.get("categories");
      const s=await ST.get("settings"); const o=await ST.get("orders");
      setProducts(p||SEED_PRODUCTS); setCats(c||SEED_CATS);
      const sett=s||SEED_SETTINGS; setSettings(sett);
      setOrders(o||[]); 
      setThemeName(sett.theme||"light"); setLang(sett.language||"mm");
      setReady(true);
    })();
  },[]);

  useEffect(()=>{ if(ready){ST.set("products",products);} },[products,ready]);
  useEffect(()=>{ if(ready){ST.set("categories",cats);} },[cats,ready]);
  useEffect(()=>{ if(ready){ST.set("settings",settings);} },[settings,ready]);
  useEffect(()=>{ if(ready){ST.set("orders",orders);} },[orders,ready]);

  const changeTheme=k=>{setThemeName(k);setSettings(s=>({...s,theme:k}));};
  const changeLang=l=>{setLang(l);setSettings(s=>({...s,language:l}));};
  const cyclTheme=()=>{const ks=Object.keys(THEMES);const i=ks.indexOf(themeName);changeTheme(ks[(i+1)%ks.length]);};

  const allCatLabel = t.allCat;
  const catFilterKey = catFilter;

  const searchRes=useMemo(()=>{
    if(!searchQ.trim()) return [];
    const q=searchQ.toLowerCase();
    return products.filter(p=>p.visible&&(
      (p.name||"").toLowerCase().includes(q)||(p.name_mm||"").includes(searchQ)||
      (p.description||"").toLowerCase().includes(q)||(p.description||"").includes(searchQ)||
      (p.category||"").includes(searchQ)||(p.suitable_for||"").includes(searchQ)||(p.benefits||"").includes(searchQ)
    )).slice(0,7);
  },[searchQ,products]);

  const visProds=useMemo(()=>
    products.filter(p=>p.visible&&(catFilterKey==="all"||p.category===catFilterKey))
  ,[products,catFilterKey]);

  const cartCount=cart.reduce((s,i)=>s+i.qty,0);
  const cartTotal=cart.reduce((s,i)=>s+calcUnit(i.p,i.qty)*i.qty,0);
  const addToCart=(p,qty=1)=>setCart(prev=>{const ex=prev.find(i=>i.p.id===p.id);return ex?prev.map(i=>i.p.id===p.id?{...i,qty:i.qty+qty}:i):[...prev,{p,qty}];});
  const removeItem=id=>setCart(c=>c.filter(i=>i.p.id!==id));
  const updateQty=(id,qty)=>{if(qty<=0)removeItem(id);else setCart(c=>c.map(i=>i.p.id===id?{...i,qty}:i));};
  const openProd=p=>{setSelProd(p);setPage("product");setShowDrop(false);setSearchQ("");};
  const goBack=()=>{ if(page==="product")setPage("shop"); else if(page==="cart")setPage("shop"); else if(page==="order")setPage("cart"); else if(page==="admin"){setPage("shop");setAdminIn(false);} };

  const bgStyle = T.glass ? {background:T.bg,backgroundAttachment:"fixed"} : {background:T.bg};

  if(!ready) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:"#FFF8F5",flexDirection:"column",gap:12}}>
      <div style={{fontSize:56,animation:"pulse 1s infinite"}}>🛍️</div>
      <div style={{color:"#718096",fontSize:15,fontWeight:600}}>Loading...</div>
    </div>
  );

  return (
    <div style={{fontFamily:"'Segoe UI','Myanmar Text',Helvetica,sans-serif",minHeight:"100vh",...bgStyle}}>
      <Header T={T} t={t} shopMM={settings.shop_name_mm||settings.shop_name||"ဆိုင်"}
        logo={settings.logo||""} cartCount={cartCount} searchQ={searchQ} setSearchQ={setSearchQ}
        showDrop={showDrop} setShowDrop={setShowDrop} dropRes={searchRes}
        onHit={openProd} onCart={()=>setPage("cart")} onAdmin={()=>setPage("admin")}
        onLogo={()=>setPage("shop")} onBack={goBack} page={page}
        onLangToggle={()=>changeLang(lang==="mm"?"en":"mm")}
        onThemeCycle={cyclTheme} curThemeName={themeName}/>

      {page==="shop"&&<ShopPage T={T} t={t} products={visProds} cats={cats} catFilter={catFilterKey==="all"?t.allCat:catFilterKey}
        setCat={c=>setCatFilter(c===t.allCat?"all":c)} onOpen={openProd} onAdd={addToCart}/>}
      {page==="product"&&selProd&&<ProductPage T={T} t={t} p={selProd} addToCart={addToCart} onCartClick={()=>setPage("cart")} cartCount={cartCount}/>}
      {page==="cart"&&<CartPage T={T} t={t} cart={cart} updateQty={updateQty} removeItem={removeItem} total={cartTotal} onOrder={()=>setPage("order")}/>}
      {page==="order"&&<OrderPage T={T} t={t} cart={cart} total={cartTotal} settings={settings}
        onPlaced={newOrd=>{setOrders(prev=>[newOrd,...prev]);setCart([]);}}/>}
      {page==="admin"&&!adminIn&&<AdminLogin T={T} t={t} adminPw={settings.admin_pw} onSuccess={()=>setAdminIn(true)} onBack={()=>setPage("shop")}/>}
      {page==="admin"&&adminIn&&(
        <AdminPanel T={T} t={t} products={products} cats={cats} settings={settings} orders={orders}
          onProdChange={p=>{setProducts(p);}} onCatsChange={setCats}
          onSettingsChange={s=>{setSettings(s);setThemeName(s.theme||"light");setLang(s.language||"mm");alert("✅ Saved!");}}
          onBack={()=>{setPage("shop");setAdminIn(false);}}
          onThemeChange={changeTheme} onLangChange={changeLang}
          curTheme={themeName} curLang={lang}/>
      )}
    </div>
  );
}