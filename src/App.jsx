import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── Supabase ──────────────────────────────────────────────────
const SB_URL = "";  // Set in Vercel env: VITE_SUPABASE_URL
const SB_KEY = "";  // Set in Vercel: VITE_SUPABASE_ANON_KEY

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

const sbAuth = async (path, body) => {
  const res = await fetch(`${SB_URL}/auth/v1/${path}`, {
    method: "POST",
    headers: { apikey: SB_KEY, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return res.json();
};

const sbUpsert = (table, rows) => fetch(`${SB_URL}/rest/v1/${table}`, {
  method:"POST",
  headers:{ apikey:SB_KEY, Authorization:`Bearer ${SB_KEY}`, "Content-Type":"application/json", Prefer:"resolution=merge-duplicates,return=representation" },
  body: JSON.stringify(rows),
}).then(r=>r.json());

const LS = {
  get: k => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set: (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} },
  del: k => { try { localStorage.removeItem(k); } catch {} },
};

// ── Helpers ───────────────────────────────────────────────────
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

// ── Default Logo (Shwe Twin BKK) ──────────────────────────────
const DEFAULT_LOGO = "data:image/png;base64,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";

// ── iPhone Glass Theme ────────────────────────────────────────
const THEMES = {
  light: {
    // iPhone frosted glass - light mode
    bg: "linear-gradient(160deg,#fce4ec 0%,#f8bbd0 20%,#fce4ec 40%,#fff9c4 60%,#fce4ec 80%,#f8bbd0 100%)",
    surface:       "rgba(255,255,255,0.72)",
    surface2:      "rgba(255,255,255,0.45)",
    surfaceStrong: "rgba(255,255,255,0.88)",
    navBg:         "rgba(255,255,255,0.82)",
    headerBg:      "rgba(252,232,236,0.85)",
    blur:          "blur(40px) saturate(180%)",
    border:        "rgba(0,0,0,0.07)",
    borderStrong:  "rgba(0,0,0,0.12)",
    text:    "#1d1d1f",
    textSub: "#3d3d42",
    muted:   "#8e8e93",
    primary:      "#E91E8C",
    primaryHover: "#C2185B",
    accent:  "#F5C518",
    success: "#34C759",
    danger:  "#FF3B30",
    warning: "#FF9500",
    shadow:   "0 2px 20px rgba(233,30,140,0.10)",
    shadowMd: "0 8px 40px rgba(233,30,140,0.15)",
    shadowLg: "0 16px 60px rgba(233,30,140,0.20)",
    tag:     "rgba(233,30,140,0.10)",
    tagText: "#C2185B",
    blob1: "rgba(255,182,209,0.50)",
    blob2: "rgba(245,197,24,0.30)",
    blob3: "rgba(255,145,200,0.35)",
    successBg: "rgba(52,199,89,0.12)",
    dangerBg:  "rgba(255,59,48,0.10)",
    warningBg: "rgba(255,149,0,0.12)",
    dark: false,
  },
  dark: {
    // iPhone frosted glass - dark mode
    bg: "linear-gradient(160deg,#1a0a1e 0%,#2d1030 25%,#1a1a2e 50%,#2d1530 75%,#0a0a1e 100%)",
    surface:       "rgba(44,44,46,0.78)",
    surface2:      "rgba(44,44,46,0.50)",
    surfaceStrong: "rgba(58,58,60,0.90)",
    navBg:         "rgba(22,22,24,0.85)",
    headerBg:      "rgba(28,10,32,0.88)",
    blur:          "blur(40px) saturate(180%)",
    border:        "rgba(255,255,255,0.10)",
    borderStrong:  "rgba(255,255,255,0.18)",
    text:    "#f5f5f7",
    textSub: "#ebebf0",
    muted:   "#8e8e93",
    primary:      "#FF375F",
    primaryHover: "#E91E8C",
    accent:  "#F5C518",
    success: "#30D158",
    danger:  "#FF453A",
    warning: "#FF9F0A",
    shadow:   "0 2px 20px rgba(0,0,0,0.50)",
    shadowMd: "0 8px 40px rgba(0,0,0,0.60)",
    shadowLg: "0 16px 60px rgba(0,0,0,0.70)",
    tag:     "rgba(255,55,95,0.20)",
    tagText: "#FF6B9A",
    blob1: "rgba(233,30,140,0.25)",
    blob2: "rgba(100,40,120,0.25)",
    blob3: "rgba(60,20,80,0.30)",
    successBg: "rgba(48,209,88,0.15)",
    dangerBg:  "rgba(255,69,58,0.15)",
    warningBg: "rgba(255,159,10,0.15)",
    dark: true,
  },
};

// ── Status colors ─────────────────────────────────────────────
const SC = {
  pending:   {bg:"rgba(255,149,0,0.15)",  color:"#CC7700",border:"rgba(255,149,0,0.4)"},
  confirmed: {bg:"rgba(0,122,255,0.12)",  color:"#0055CC",border:"rgba(0,122,255,0.35)"},
  deposited: {bg:"rgba(52,199,89,0.12)",  color:"#1A8C3A",border:"rgba(52,199,89,0.35)"},
  completed: {bg:"rgba(52,199,89,0.15)",  color:"#0D6B2A",border:"rgba(52,199,89,0.5)"},
  cancelled: {bg:"rgba(255,59,48,0.12)",  color:"#CC1100",border:"rgba(255,59,48,0.35)"},
};

// ── i18n ──────────────────────────────────────────────────────
const TR = {
  en:{
    appName:"Shwe Twin BKK",
    home:"Home", cats:"Categories", cart:"Cart", orders:"Orders", profile:"Profile",
    search:"Search Shwe Twin products...", all:"All",
    addCart:"Add to Cart", buyNow:"Buy Now", preorder:"Pre-order",
    outOfStock:"Out of Stock", tempOOS:"Temporarily Out of Stock", inStock:"In Stock",
    qty:"Qty", total:"Total", checkout:"Checkout",
    orderSummary:"Order Summary", contactOrder:"Contact to Order",
    trackOrder:"Track Order", enterOrderNo:"Enter order number", track:"Track",
    yourName:"Your Name *", yourPhone:"Phone *", yourAddress:"Delivery Address *",
    note:"Note / Special requests",
    signIn:"Sign In", signUp:"Sign Up", signOut:"Sign Out",
    continueGoogle:"Continue with Google", continueFacebook:"Continue with Facebook",
    continuePhone:"Continue with Phone", continueEmail:"Continue with Email",
    orSignIn:"or sign in with", phoneNumber:"Phone Number",
    sendOTP:"Send OTP", enterOTP:"Enter OTP Code", verifyOTP:"Verify",
    email:"Email", password:"Password", forgotPw:"Forgot Password?",
    noAccount:"Don't have an account?", hasAccount:"Already have an account?",
    guestCheckout:"Continue as Guest", myOrders:"My Orders",
    adminPanel:"Admin Panel", login:"Login", products:"Products",
    reports:"Reports", settings:"Settings", addProduct:"Add Product",
    editProduct:"Edit Product", save:"Save", cancel:"Cancel",
    delete:"Delete", hide:"Hide", show:"Show", addStock:"Add Stock",
    switchPreorder:"Switch to Pre-order",
    productName:"Product Name (English) *", productNameMM:"Product Name (Myanmar)",
    category:"Category", price:"Price (Ks) *", discountType:"Discount Type",
    discountPct:"Percent %", discountFixed:"Fixed Amount (Ks)",
    discountVal:"Discount Value", stockQty:"Stock Qty",
    enablePreorder:"Enable Pre-order", images:"Images", videoUrl:"Video",
    emoji:"Emoji", description:"Description", suitableFor:"Suitable For",
    benefits:"Benefits", usage:"How to Use", warning:"Warning",
    bulkTiers:"Bulk Discount Tiers", addTier:"+ Add Tier",
    gdriveTip:"Use imgbb.com for images",
    shopName:"Shop Name (EN)", shopNameMM:"Shop Name (MM)",
    fbLink:"Facebook Messenger URL", viberNum:"Viber Number",
    waNum:"WhatsApp Number", phoneNum:"Phone Number",
    adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme",
    langLabel:"Language", saveSettings:"Save Settings",
    pending:"Pending", confirmed:"Confirmed", deposited:"Deposited",
    completed:"Completed", cancelled:"Cancelled",
    depositPaid:"Deposit Paid", balanceDue:"Balance Due",
    confirmOrder:"Confirm & Deduct Stock", updateDeposit:"Update Deposit",
    adminNote:"Admin Note", updateStatus:"Update Status",
    today:"Today", week:"This Week", month:"This Month", allTime:"All Time",
    revenue:"Revenue", totalOrders:"Total Orders",
    topProducts:"Top Products", recentOrders:"Recent Orders",
    exportExcel:"Export Excel", noData:"No data yet",
    noProducts:"No products found", emptyCart:"Cart is empty",
    orderPlaced:"Order Sent!", orderPlacedMsg:"We will contact you soon.",
    wrongPw:"Wrong password", logout:"Logout", backToShop:"Back to Shop",
    notifications:"Notifications", clearAll:"Clear All",
    stockAdded:"Stock added", stockDeducted:"Stock deducted", notifUpdated:"updated",
    popularItems:"Popular Items", newArrivals:"New Arrivals",
    welcomeBack:"Welcome Back", joinUs:"Join Shwe Twin",
    authSubtitle:"Bangkok direct shopping, delivered to you",
  },
  mm:{
    appName:"ရွှေတွင် BKK",
    home:"ပင်မ", cats:"အမျိုးအစား", cart:"Cart", orders:"Orders", profile:"ကျွန်ုပ်",
    search:"ရွှေတွင် ထုတ်ကုန် ရှာပါ...", all:"အားလုံး",
    addCart:"Cart ထည့်", buyNow:"ချက်ချင်းဝယ်", preorder:"Pre-order",
    outOfStock:"ကုန်ပြီ", tempOOS:"ယာယီ ကုန်ပြီ", inStock:"Stock ရှိ",
    qty:"အရေ", total:"စုစုပေါင်း", checkout:"Order လုပ်မည်",
    orderSummary:"Order အကျဉ်း", contactOrder:"ဆက်သွယ်ပြီး Order မှာမည်",
    trackOrder:"Order ရှာမည်", enterOrderNo:"Order နံပါတ် ထည့်ပါ", track:"ရှာမည်",
    yourName:"နာမည် *", yourPhone:"ဖုန်းနံပါတ် *", yourAddress:"လိပ်စာ *",
    note:"မှာကြားချက်",
    signIn:"Login ဝင်မည်", signUp:"အကောင့် ဖန်တီးမည်", signOut:"Logout",
    continueGoogle:"Google ဖြင့် ဝင်မည်", continueFacebook:"Facebook ဖြင့် ဝင်မည်",
    continuePhone:"ဖုန်းနံပါတ် ဖြင့် ဝင်မည်", continueEmail:"Email ဖြင့် ဝင်မည်",
    orSignIn:"သို့မဟုတ် ဤနည်းဖြင့်", phoneNumber:"ဖုန်းနံပါတ်",
    sendOTP:"OTP ပို့မည်", enterOTP:"OTP Code ထည့်ပါ", verifyOTP:"အတည်ပြုမည်",
    email:"Email", password:"Password", forgotPw:"Password မေ့နေသလား?",
    noAccount:"အကောင့် မရှိသေးပါ?", hasAccount:"အကောင့် ရှိပြီးသား?",
    guestCheckout:"Guest အဖြစ် ဆက်သွားမည်", myOrders:"ကျွန်ုပ် Orders",
    adminPanel:"Admin Panel", login:"Login ဝင်မည်", products:"Products",
    reports:"Reports", settings:"Settings", addProduct:"Product ထည့်မည်",
    editProduct:"Product ပြင်မည်", save:"သိမ်းမည်", cancel:"မလုပ်တော့ပါ",
    delete:"ဖျက်", hide:"ဖျောက်", show:"ပြ", addStock:"Stock ထည့်မည်",
    switchPreorder:"Pre-order ပြောင်းမည်",
    productName:"Product နာမည် (English) *", productNameMM:"Product နာမည် (မြန်မာ)",
    category:"အမျိုးအစား", price:"ဈေးနှုန်း (Ks) *", discountType:"လျှော့ဈေး အမျိုး",
    discountPct:"ရာခိုင်နှုန်း %", discountFixed:"ပမာဏ (Ks)",
    discountVal:"လျှော့ဈေး", stockQty:"Stock အရေ",
    enablePreorder:"Pre-order မှာနိုင်", images:"ပုံများ", videoUrl:"Video",
    emoji:"Emoji", description:"ဖော်ပြချက်", suitableFor:"သင့်တော်သူ",
    benefits:"အကျိုးကျေးဇူး", usage:"သုံးနည်း", warning:"သတိပြုရန်",
    bulkTiers:"အရေ Tier လျှော့ဈေး", addTier:"+ Tier ထည့်",
    gdriveTip:"ပုံများအတွက် imgbb.com သုံးပါ",
    shopName:"ဆိုင်နာမည် (EN)", shopNameMM:"ဆိုင်နာမည် (MM)",
    fbLink:"Facebook Messenger URL", viberNum:"Viber နံပါတ်",
    waNum:"WhatsApp နံပါတ်", phoneNum:"ဖုန်းနံပါတ်",
    adminPw:"Admin Password", logo:"Logo", themeLabel:"Theme",
    langLabel:"ဘာသာ", saveSettings:"Settings သိမ်းမည်",
    pending:"စောင့်ဆိုင်း", confirmed:"အတည်ပြု",
    deposited:"စရံပေးပြီး", completed:"ပြီးစီး", cancelled:"ပယ်ဖျက်",
    depositPaid:"စရံပေးပြီး", balanceDue:"ကျန်ငွေ",
    confirmOrder:"အတည်ပြုပြီး Stock နှုတ်မည်", updateDeposit:"စရံ Update",
    adminNote:"Admin မှတ်ချက်", updateStatus:"Status ပြောင်းမည်",
    today:"ဒီနေ့", week:"ဒီအပတ်", month:"ဒီလ", allTime:"အားလုံး",
    revenue:"ဝင်ငွေ", totalOrders:"Order စုစုပေါင်း",
    topProducts:"အရောင်းကောင်းဆုံး", recentOrders:"နောက်ဆုံး Orders",
    exportExcel:"Excel ထုတ်", noData:"Data မရှိသေးပါ",
    noProducts:"ထုတ်ကုန် မတွေ့ပါ", emptyCart:"Cart ထဲ ဘာမှ မရှိသေးပါ",
    orderPlaced:"Order ပေးပို့ပြီးပြီ!", orderPlacedMsg:"Admin မှ မကြာမီ ဆက်သွယ်ပါမည်။",
    wrongPw:"Password မှားနေသည်", logout:"Logout", backToShop:"ဆိုင်သို့ ပြန်",
    notifications:"အကြောင်းကြားချက်", clearAll:"အားလုံးဖျက်",
    stockAdded:"Stock ထည့်ပြီး", stockDeducted:"Stock နှုတ်သွားသည်",
    notifUpdated:"ပြောင်းလဲပြီး", popularItems:"Popular Items",
    newArrivals:"အသစ်ရောက်သည်",
    welcomeBack:"ကြိုဆိုပါသည်", joinUs:"ရွှေတွင်တွင် ပါဝင်ပါ",
    authSubtitle:"ဘန်ကောက် တိုက်ရိုက် ပစ္စည်းများ မှာယူနိုင်သည်",
  },
};

// ── Notification store ────────────────────────────────────────
let _nl=[];
const NS={items:[],add(msg){NS.items=[{id:uid(),msg,time:new Date().toISOString(),read:false},...NS.items].slice(0,30);_nl.forEach(f=>f([...NS.items]));},markRead(){NS.items=NS.items.map(n=>({...n,read:true}));_nl.forEach(f=>f([...NS.items]));},clear(){NS.items=[];_nl.forEach(f=>f([]));},subscribe(fn){_nl.push(fn);return()=>{_nl=_nl.filter(f=>f!==fn);};}};
function useNotifs(){const[n,setN]=useState(NS.items);useEffect(()=>NS.subscribe(setN),[]);return n;}

// ── iPhone glass helpers ──────────────────────────────────────
const ios = (T, extra={}) => ({
  background: T.surface,
  backdropFilter: T.blur,
  WebkitBackdropFilter: T.blur,
  border: `1px solid ${T.border}`,
  ...extra,
});
const iosCard = (T, extra={}) => ({
  ...ios(T),
  borderRadius: 20,
  boxShadow: T.shadow,
  ...extra,
});

// ── Setup SQL ─────────────────────────────────────────────────
const SETUP_SQL = `-- Run in Supabase SQL Editor
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
DROP TABLE IF EXISTS shop_settings,orders,products,categories CASCADE;
CREATE TABLE categories(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL UNIQUE,sort_order INT DEFAULT 0,created_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE products(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),name TEXT NOT NULL,name_mm TEXT,category TEXT NOT NULL DEFAULT '',description TEXT,suitable_for TEXT,benefits TEXT,usage_info TEXT,warning TEXT,price NUMERIC NOT NULL DEFAULT 0,discount_type TEXT DEFAULT 'percent',discount_value NUMERIC DEFAULT 0,bulk_discounts JSONB DEFAULT '[]',stock INT NOT NULL DEFAULT 0,preorder BOOLEAN DEFAULT false,images JSONB DEFAULT '[]',video_url TEXT,emoji TEXT DEFAULT '🛍️',visible BOOLEAN DEFAULT true,featured BOOLEAN DEFAULT false,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE orders(id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),order_number TEXT,customer_name TEXT,customer_phone TEXT,customer_address TEXT,items JSONB NOT NULL DEFAULT '[]',total NUMERIC NOT NULL DEFAULT 0,status TEXT DEFAULT 'pending',contact_method TEXT,customer_note TEXT,deposit_paid NUMERIC DEFAULT 0,balance_due NUMERIC DEFAULT 0,admin_note TEXT,stock_deducted BOOLEAN DEFAULT false,user_id TEXT,created_at TIMESTAMPTZ DEFAULT NOW(),updated_at TIMESTAMPTZ DEFAULT NOW());
CREATE TABLE shop_settings(key TEXT PRIMARY KEY,value TEXT NOT NULL DEFAULT '');
ALTER TABLE categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE shop_settings ENABLE ROW LEVEL SECURITY;
CREATE POLICY "all" ON categories FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON products FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON orders FOR ALL USING(true)WITH CHECK(true);
CREATE POLICY "all" ON shop_settings FOR ALL USING(true)WITH CHECK(true);
INSERT INTO categories(name,sort_order)VALUES('Thai Skincare',1),('Korean Beauty',2),('Makeup',3),('Hair Care',4),('Fashion',5),('Accessories',6),('Others',7);
INSERT INTO shop_settings(key,value)VALUES('shop_name','Shwe Twin BKK Direct'),('shop_name_mm','ရွှေတွင် BKK Direct'),('fb_link','https://m.me/yourpage'),('viber_num','+95912345678'),('wa_num','+95912345678'),('phone_num','+95912345678'),('admin_pw','admin123'),('logo',''),('banner','');
SELECT 'Done' FROM shop_settings LIMIT 1;`;

// ── Sample products ───────────────────────────────────────────
const SAMPLE_PRODUCTS = [
  {id:"sp1",name:"Snail White Body Lotion",name_mm:"ကျည်ပိုး ဖြူဖြူ Body Lotion",category:"Thai Skincare",description:"Popular Thai whitening body lotion with snail extract. ထိုင်းနိုင်ငံ လူကြိုက်များသော ဖြူဆေး ကိုယ်လိမ်းခရင်မ်။",suitable_for:"All skin types",benefits:"Brightens skin, moisturizes, anti-aging",usage_info:"Apply evenly after shower. ရေချိုးပြီးနောက် တစ်နှစ် တစ်ကြိမ် လိမ်းပါ။",warning:"",price:15000,discount_type:"percent",discount_value:15,bulk_discounts:[{min_qty:3,discount_percent:20}],stock:50,preorder:false,images:["https://images.unsplash.com/photo-1556228578-8c89e6adf883?w=400&h=400&fit=crop"],video_url:"",emoji:"🐌",visible:true,featured:true,created_at:new Date().toISOString()},
  {id:"sp2",name:"SOME BY MI AHA BHA PHA Toner",name_mm:"SOME BY MI ချဉ်ဆေး Toner",category:"Korean Beauty",description:"Korean cult-favorite exfoliating toner. ကိုရီးယား လူကြိုက်များသော dead cell ဖယ်ရှားပေးသော toner။",suitable_for:"Oily & combination skin",benefits:"Clears acne, refines pores, brightens",usage_info:"Apply with cotton pad after cleansing",warning:"Avoid eye area. Sun protection required.",price:28000,discount_type:"percent",discount_value:10,bulk_discounts:[],stock:30,preorder:false,images:["https://images.unsplash.com/photo-1620916566398-39f1143ab7be?w=400&h=400&fit=crop"],video_url:"",emoji:"✨",visible:true,featured:true,created_at:new Date().toISOString()},
  {id:"sp3",name:"Mistine Sunscreen SPF50",name_mm:"Mistine နေရောင်ကာ SPF50",category:"Thai Skincare",description:"Thailand bestselling sunscreen SPF50+. ထိုင်းနိုင်ငံ No.1 နေရောင်ကာ ခရင်မ်။ လျင်မြန်သော absorption ကြောင့် sticky မဟုတ်ပါ။",suitable_for:"All skin types, daily use",benefits:"UV protection, lightweight, non-greasy",usage_info:"Apply 15 mins before sun exposure",warning:"",price:8000,discount_type:"fixed",discount_value:1000,bulk_discounts:[{min_qty:2,discount_percent:10},{min_qty:5,discount_percent:18}],stock:100,preorder:false,images:["https://images.unsplash.com/photo-1556228720-195a672e8a03?w=400&h=400&fit=crop"],video_url:"",emoji:"☀️",visible:true,featured:false,created_at:new Date().toISOString()},
  {id:"sp4",name:"Laneige Lip Sleeping Mask",name_mm:"Laneige နှုတ်ခမ်းဆေး Mask",category:"Korean Beauty",description:"Overnight lip treatment mask. အိပ်နေစဉ် နှုတ်ခမ်းကို အာဟာရပေးသော mask။ Berry scent ပါသည်။",suitable_for:"Dry & chapped lips",benefits:"Deep moisturize, softens, plumps lips",usage_info:"Apply before sleep, wash off in morning",warning:"",price:22000,discount_type:"percent",discount_value:0,bulk_discounts:[],stock:25,preorder:false,images:["https://images.unsplash.com/photo-1586495777744-4e6232bf2e53?w=400&h=400&fit=crop"],video_url:"",emoji:"💋",visible:true,featured:true,created_at:new Date().toISOString()},
  {id:"sp5",name:"Cute Press CC Cushion",name_mm:"Cute Press CC Cushion",category:"Makeup",description:"Thai makeup brand CC cushion with SPF30. ထိုင်း makeup brand Cute Press ၏ CC cushion။ Natural glow ပေးသည်။",suitable_for:"Medium coverage, natural look",benefits:"SPF30, long-lasting, natural finish",usage_info:"Pat onto face using cushion applicator",warning:"",price:12000,discount_type:"percent",discount_value:20,bulk_discounts:[{min_qty:2,discount_percent:25}],stock:0,preorder:true,images:["https://images.unsplash.com/photo-1596462502278-27bfdc403348?w=400&h=400&fit=crop"],video_url:"",emoji:"💄",visible:true,featured:false,created_at:new Date().toISOString()},
  {id:"sp6",name:"Ellips Hair Vitamin",name_mm:"Ellips ဆံပင်ဗီတာမင်",category:"Hair Care",description:"Thailand popular hair serum capsules. ထိုင်းနိုင်ငံ လူကြိုက်များသော ဆံပင်ဆီ capsule များ။ ဆံပင်ကြမ်းမှု ပြေပျောက်စေသည်။",suitable_for:"Dry, damaged, frizzy hair",benefits:"Smooths frizz, adds shine, repairs damage",usage_info:"Break capsule and apply to damp or dry hair ends",warning:"",price:6000,discount_type:"percent",discount_value:0,bulk_discounts:[{min_qty:3,discount_percent:15}],stock:80,preorder:false,images:["https://images.unsplash.com/photo-1629470791566-3a8f2b05e3c5?w=400&h=400&fit=crop"],video_url:"",emoji:"💆",visible:true,featured:false,created_at:new Date().toISOString()},
];

const SAMPLE_CATS = [
  {id:"sc1",name:"Thai Skincare",sort_order:1},
  {id:"sc2",name:"Korean Beauty",sort_order:2},
  {id:"sc3",name:"Makeup",sort_order:3},
  {id:"sc4",name:"Hair Care",sort_order:4},
  {id:"sc5",name:"Fashion",sort_order:5},
  {id:"sc6",name:"Accessories",sort_order:6},
  {id:"sc7",name:"Others",sort_order:7},
];

const SAMPLE_SETTINGS = {
  shop_name:"Shwe Twin BKK Direct",
  shop_name_mm:"ရွှေတွင် BKK Direct",
  fb_link:"https://m.me/yourpage",
  viber_num:"+95912345678",
  wa_num:"+95912345678",
  phone_num:"+95912345678",
  admin_pw:"admin123",
  logo:"",banner:"",
};


// ── Setup Screen ──────────────────────────────────────────────
function SetupScreen({onConnect}){
  const[url,setUrl]=useState("");const[key,setKey]=useState("");const[busy,setBusy]=useState(false);const[err,setErr]=useState("");const[showSql,setShowSql]=useState(false);
  const T=THEMES.light;
  const go=async()=>{if(!url||!key){setErr("URL နှင့် Key ထည့်ပါ");return;}setBusy(true);setErr("");try{const cfg={url:url.replace(/\/+$/,""),key};const r=await fetch(`${cfg.url}/rest/v1/shop_settings?limit=1`,{headers:{apikey:cfg.key,Authorization:`Bearer ${cfg.key}`}});if(!r.ok)throw new Error(await r.text());LS.set("sb_cfg",cfg);onConnect(cfg);}catch(e){setErr("ချိတ်မရပါ: "+e.message);}setBusy(false);};
  return(<div style={{minHeight:"100vh",background:T.bg,backgroundAttachment:"fixed",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{...iosCard(T,{padding:"40px 28px",maxWidth:460,width:"100%"})}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <img src={DEFAULT_LOGO} style={{width:90,height:90,borderRadius:45,marginBottom:12,boxShadow:T.shadowMd}} alt="logo"/>
        <div style={{fontSize:22,fontWeight:800,color:T.text,letterSpacing:-0.5}}>Shwe Twin BKK Direct</div>
        <div style={{fontSize:13,color:T.muted,marginTop:4}}>Supabase ချိတ်ဆက်မည်</div>
      </div>
      <div style={{...ios(T,{borderRadius:12,padding:"12px 14px",marginBottom:14,fontSize:12,color:T.text,lineHeight:1.9})}}>
        <b>Step 1:</b> supabase.com → New Project<br/><b>Step 2:</b> SQL Editor → paste SQL → Run<br/><b>Step 3:</b> Vercel Env → VITE_SUPABASE_URL + KEY<br/><b>Step 4:</b> URL + Key ထည့်မည်
      </div>
      <button onClick={()=>setShowSql(v=>!v)} style={{width:"100%",padding:9,borderRadius:10,...ios(T,{}),cursor:"pointer",fontSize:12,color:T.muted,marginBottom:10}}>{showSql?"▲ SQL ပိတ်":"▼ Setup SQL ကြည့်မည်"}</button>
      {showSql&&<pre style={{background:"rgba(0,0,0,0.85)",borderRadius:10,padding:12,fontSize:9.5,fontFamily:"monospace",color:"#7FFFD4",overflow:"auto",maxHeight:160,lineHeight:1.6,marginBottom:12,whiteSpace:"pre-wrap"}}>{SETUP_SQL}</pre>}
      {["SUPABASE PROJECT URL","ANON PUBLIC KEY"].map((lbl,i)=><div key={i}>
        <label style={{fontSize:11,fontWeight:600,color:T.muted,display:"block",marginBottom:4,textTransform:"uppercase",letterSpacing:0.5}}>{lbl}</label>
        <input type={i===1?"password":"text"} style={{width:"100%",padding:"12px 14px",borderRadius:12,...ios(T,{}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:T.text,background:T.surfaceStrong}}
          value={i===0?url:key} onChange={e=>i===0?setUrl(e.target.value):setKey(e.target.value)} placeholder={i===0?"https://xxxx.supabase.co":"eyJhbGci..."}/>
      </div>)}
      {err&&<div style={{color:T.danger,fontSize:13,marginBottom:12,...ios(T,{borderRadius:8,padding:"10px 12px",background:T.dangerBg})}}>{err}</div>}
      <button onClick={go} disabled={busy} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:busy?"rgba(233,30,140,0.4)":T.primary,color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 20px ${T.primary}50`}}>
        {busy?"⏳ စစ်ဆေးနေသည်...":"🔗 ချိတ်ဆက်မည်"}
      </button>
    </div>
  </div>);
}

// ── Auth Screen (Sign in / Sign up) ───────────────────────────
function AuthScreen({T,t,onAuth,onGuest,cfg}){
  const[mode,setMode]=useState("main"); // main|phone|email
  const[phone,setPhone]=useState("+95");
  const[otp,setOtp]=useState("");
  const[otpSent,setOtpSent]=useState(false);
  const[email,setEmail]=useState("");
  const[pw,setPw]=useState("");
  const[isSignUp,setIsSignUp]=useState(false);
  const[busy,setBusy]=useState(false);
  const[err,setErr]=useState("");

  const sendOTP=async()=>{
    if(!phone.trim()||phone.length<8){setErr("ဖုန်းနံပါတ် မှန်မှန်ထည့်ပါ");return;}
    setBusy(true);setErr("");
    try{
      const res=await sbAuth("otp",{phone,create_user:true});
      if(res.error)throw new Error(res.error.message||"OTP မပို့နိုင်ပါ");
      setOtpSent(true);
    }catch(e){setErr(e.message);}
    setBusy(false);
  };

  const verifyOTP=async()=>{
    if(!otp.trim()){setErr("OTP ထည့်ပါ");return;}
    setBusy(true);setErr("");
    try{
      const res=await sbAuth("verify",{phone,token:otp,type:"sms"});
      if(res.error)throw new Error(res.error.message||"OTP မမှန်ကန်ပါ");
      LS.set("user",{phone,name:phone,provider:"phone",access_token:res.access_token});
      onAuth({phone,name:phone,provider:"phone"});
    }catch(e){setErr(e.message);}
    setBusy(false);
  };

  const signInEmail=async()=>{
    if(!email||!pw){setErr("Email နှင့် Password ထည့်ပါ");return;}
    setBusy(true);setErr("");
    try{
      const path=isSignUp?"signup":"token?grant_type=password";
      const res=await sbAuth(path,{email,password:pw});
      if(res.error)throw new Error(res.error.message);
      const user={email,name:email.split("@")[0],provider:"email",access_token:res.access_token};
      LS.set("user",user);onAuth(user);
    }catch(e){setErr(e.message);}
    setBusy(false);
  };

  const signInGoogle=()=>{
    const redirect=encodeURIComponent(window.location.origin);
    window.location.href=`${cfg?.url||SB_URL}/auth/v1/authorize?provider=google&redirect_to=${redirect}`;
  };

  const signInFacebook=()=>{
    const redirect=encodeURIComponent(window.location.origin);
    window.location.href=`${cfg?.url||SB_URL}/auth/v1/authorize?provider=facebook&redirect_to=${redirect}`;
  };

  const inp={width:"100%",padding:"13px 16px",borderRadius:14,...ios(T,{}),fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:T.text,background:T.surfaceStrong};
  const BigBtn=({bg,color,icon,label,onClick,sub})=>(
    <button onClick={onClick} style={{width:"100%",padding:"14px 18px",borderRadius:16,border:`1px solid ${T.border}`,background:bg||T.surface,color:color||T.text,cursor:"pointer",fontSize:14,fontWeight:600,display:"flex",alignItems:"center",gap:12,marginBottom:10,...(bg?{boxShadow:`0 4px 16px ${bg}50`}:{...ios(T,{})})}}>
      <span style={{fontSize:22,flexShrink:0}}>{icon}</span>
      <div style={{textAlign:"left",flex:1}}><div>{label}</div>{sub&&<div style={{fontSize:11,opacity:0.7,marginTop:1}}>{sub}</div>}</div>
      <span style={{opacity:0.4,fontSize:16}}>→</span>
    </button>
  );

  return(<div style={{minHeight:"100vh",background:T.bg,backgroundAttachment:"fixed",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{...iosCard(T,{padding:"32px 24px",maxWidth:400,width:"100%"})}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <img src={DEFAULT_LOGO} style={{width:80,height:80,borderRadius:40,marginBottom:10,boxShadow:T.shadowMd}} alt="logo"/>
        <div style={{fontSize:20,fontWeight:800,color:T.text,letterSpacing:-0.5}}>{t.welcomeBack}</div>
        <div style={{fontSize:13,color:T.muted,marginTop:4}}>{t.authSubtitle}</div>
      </div>

      {mode==="main"&&<>
        <BigBtn bg={T.primary} color="#fff" icon="📱" label={t.continuePhone} onClick={()=>{setMode("phone");setErr("");}}/>
        <BigBtn bg="#DB4437" color="#fff" icon="G" label={t.continueGoogle} sub="Google Account ဖြင့် ဝင်မည်" onClick={signInGoogle}/>
        <BigBtn bg="#1877F2" color="#fff" icon="f" label={t.continueFacebook} sub="Facebook ဖြင့် ဝင်မည်" onClick={signInFacebook}/>
        <BigBtn icon="✉️" label={t.continueEmail} onClick={()=>{setMode("email");setErr("");}}/>
        <div style={{textAlign:"center",marginTop:8}}>
          <button onClick={onGuest} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,color:T.muted,textDecoration:"underline"}}>{t.guestCheckout}</button>
        </div>
      </>}

      {mode==="phone"&&<>
        <button onClick={()=>{setMode("main");setOtpSent(false);setErr("");}} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:13,fontWeight:600,marginBottom:16,padding:0}}>← Back</button>
        <div style={{fontSize:16,fontWeight:700,color:T.text,marginBottom:16}}>{t.continuePhone}</div>
        {!otpSent?<>
          <label style={{fontSize:12,fontWeight:600,color:T.muted,display:"block",marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>{t.phoneNumber}</label>
          <input style={inp} value={phone} onChange={e=>setPhone(e.target.value)} placeholder="+959xxxxxxxxx" type="tel"/>
          {err&&<div style={{color:T.danger,fontSize:12,marginBottom:12,...ios(T,{borderRadius:8,padding:"8px 12px",background:T.dangerBg})}}>{err}</div>}
          <button onClick={sendOTP} disabled={busy} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`}}>{busy?"⏳ ...":t.sendOTP}</button>
        </>:<>
          <div style={{fontSize:13,color:T.muted,marginBottom:12}}>OTP ကို <b>{phone}</b> သို့ ပို့ပြီးပြီ</div>
          <label style={{fontSize:12,fontWeight:600,color:T.muted,display:"block",marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>{t.enterOTP}</label>
          <input style={{...inp,textAlign:"center",fontSize:22,letterSpacing:8}} value={otp} onChange={e=>setOtp(e.target.value)} placeholder="000000" maxLength={6} type="tel"/>
          {err&&<div style={{color:T.danger,fontSize:12,marginBottom:12,...ios(T,{borderRadius:8,padding:"8px 12px",background:T.dangerBg})}}>{err}</div>}
          <button onClick={verifyOTP} disabled={busy} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`}}>{busy?"⏳ ...":t.verifyOTP}</button>
          <button onClick={()=>setOtpSent(false)} style={{width:"100%",padding:10,borderRadius:12,...ios(T,{}),border:"none",cursor:"pointer",fontSize:13,color:T.muted,marginTop:8}}>← {t.sendOTP} ပြန်ပို့မည်</button>
        </>}
      </>}

      {mode==="email"&&<>
        <button onClick={()=>{setMode("main");setErr("");}} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:13,fontWeight:600,marginBottom:16,padding:0}}>← Back</button>
        <div style={{fontSize:16,fontWeight:700,color:T.text,marginBottom:16}}>{isSignUp?t.signUp:t.signIn}</div>
        <label style={{fontSize:12,fontWeight:600,color:T.muted,display:"block",marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>{t.email}</label>
        <input style={inp} value={email} onChange={e=>setEmail(e.target.value)} placeholder="email@example.com" type="email"/>
        <label style={{fontSize:12,fontWeight:600,color:T.muted,display:"block",marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>{t.password}</label>
        <input style={inp} value={pw} onChange={e=>setPw(e.target.value)} placeholder="••••••••" type="password"/>
        {err&&<div style={{color:T.danger,fontSize:12,marginBottom:12,...ios(T,{borderRadius:8,padding:"8px 12px",background:T.dangerBg})}}>{err}</div>}
        <button onClick={signInEmail} disabled={busy} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:busy?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`,marginBottom:10}}>{busy?"⏳ ...":(isSignUp?t.signUp:t.signIn)}</button>
        <button onClick={()=>setIsSignUp(v=>!v)} style={{width:"100%",padding:10,borderRadius:12,...ios(T,{}),border:"none",cursor:"pointer",fontSize:13,color:T.primary}}>{isSignUp?t.hasAccount:t.noAccount}</button>
      </>}
    </div>
  </div>);
}

// ── PImg with iPhone glass blob ───────────────────────────────
function PImg({p,size=120,r=16,showBlob=false,T}){
  const[err,setErr]=useState(false);
  const imgs=getImgs(p);const src=imgs[0]||"";
  const blobs=[T.blob1,T.blob2,T.blob3];
  const blob=blobs[(p?.id||"x").charCodeAt(1)%3||0];
  const bgs=["rgba(255,182,209,0.25)","rgba(255,218,185,0.25)","rgba(216,191,216,0.25)","rgba(176,224,230,0.20)","rgba(255,228,196,0.25)","rgba(221,160,221,0.20)"];
  const bg=bgs[(p?.id||"x").charCodeAt(1)%bgs.length||0];
  return(<div style={{position:"relative",width:size,height:size,flexShrink:0}}>
    {showBlob&&<div style={{position:"absolute",inset:"-18%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",zIndex:0,filter:"blur(8px)"}}/>}
    <div style={{position:"relative",zIndex:1,width:"100%",height:"100%",borderRadius:r,overflow:"hidden",background:bg}}>
      {src&&!err?<img src={src} alt={p?.name||""} onError={()=>setErr(true)} style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}/>
        :<div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.38}}>{p?.emoji||"🛍️"}</div>}
    </div>
  </div>);
}

// ── Status Badge ──────────────────────────────────────────────
function SBadge({status,t}){
  const c=SC[status]||SC.pending;
  const L={pending:t.pending,confirmed:t.confirmed,deposited:t.deposited,completed:t.completed,cancelled:t.cancelled};
  return<span style={{background:c.bg,color:c.color,border:`1px solid ${c.border}`,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{L[status]||status}</span>;
}

// ── Notification Bell ─────────────────────────────────────────
function NotifBell({T,t}){
  const notifs=useNotifs();const[open,setOpen]=useState(false);const ref=useRef();const unread=notifs.filter(n=>!n.read).length;
  useEffect(()=>{const fn=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};document.addEventListener("mousedown",fn);return()=>document.removeEventListener("mousedown",fn);},[]);
  return(<div ref={ref} style={{position:"relative"}}>
    <button onClick={()=>{setOpen(v=>!v);if(!open)NS.markRead();}} style={{...ios(T,{borderRadius:20,padding:"7px 10px"}),color:T.text,cursor:"pointer",fontSize:15,position:"relative"}}>
      🔔{unread>0&&<span style={{position:"absolute",top:-4,right:-4,background:T.primary,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:700}}>{unread}</span>}
    </button>
    {open&&<div style={{position:"absolute",right:0,top:"calc(100% + 8px)",width:272,maxHeight:300,overflowY:"auto",borderRadius:16,boxShadow:T.shadowMd,zIndex:300,...iosCard(T,{})}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderBottom:`1px solid ${T.border}`}}>
        <span style={{fontWeight:700,fontSize:13,color:T.text}}>{t.notifications}</span>
        <button onClick={()=>NS.clear()} style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:T.muted}}>{t.clearAll}</button>
      </div>
      {notifs.length===0?<div style={{padding:20,textAlign:"center",color:T.muted,fontSize:13}}>—</div>
        :notifs.map(n=><div key={n.id} style={{padding:"10px 14px",borderBottom:`1px solid ${T.border}`,opacity:n.read?0.6:1}}><div style={{fontSize:12,fontWeight:600,color:T.text}}>{n.msg}</div><div style={{fontSize:10,color:T.muted,marginTop:2}}>{fdate(n.time)}</div></div>)}
    </div>}
  </div>);
}

// ── Top Bar ───────────────────────────────────────────────────
function TopBar({T,t,shopMM,logo,cartCount,searchQ,setSearchQ,showDrop,setShowDrop,dropRes,onHit,onCart,onAdmin,onLogo,onThemeToggle,themeName}){
  return(<div style={{...ios(T,{position:"sticky",top:0,zIndex:200,boxShadow:T.shadowMd,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none",background:T.headerBg})}}>
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 16px"}}>
      <button onClick={onLogo} style={{background:"none",border:"none",cursor:"pointer",padding:0,display:"flex",alignItems:"center",gap:8}}>
        <img src={logo||DEFAULT_LOGO} style={{width:36,height:36,borderRadius:18,objectFit:"cover",boxShadow:T.shadow}} alt="logo"/>
        <div>
          <div style={{fontWeight:800,fontSize:15,color:T.text,letterSpacing:-0.3,lineHeight:1.1}}>{shopMM||"Shwe Twin"}</div>
          <div style={{fontSize:9,color:T.muted,letterSpacing:0.5,textTransform:"uppercase"}}>BKK Direct</div>
        </div>
      </button>
      <div style={{flex:1}}/>
      <NotifBell T={T} t={t}/>
      <button onClick={onThemeToggle} style={{...ios(T,{borderRadius:20,padding:"6px 10px"}),color:T.text,cursor:"pointer",fontSize:13}}>{themeName==="light"?"🌙":"☀️"}</button>
      <button onClick={onCart} style={{...ios(T,{borderRadius:20,padding:"7px 13px"}),color:T.text,cursor:"pointer",display:"flex",alignItems:"center",gap:4,fontSize:14,fontWeight:700}}>
        🛒{cartCount>0&&<span style={{background:T.primary,color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:11,fontWeight:800}}>{cartCount}</span>}
      </button>
      <button onClick={onAdmin} style={{...ios(T,{borderRadius:20,padding:"6px 11px"}),color:T.text,cursor:"pointer",fontSize:13}}>⚙️</button>
    </div>
    <div style={{padding:"0 16px 12px",position:"relative"}}>
      <div style={{position:"relative"}}>
        <span style={{position:"absolute",left:14,top:"50%",transform:"translateY(-50%)",fontSize:14,pointerEvents:"none",color:T.muted}}>🔍</span>
        <input value={searchQ} onChange={e=>{setSearchQ(e.target.value);setShowDrop(true);}} onFocus={()=>searchQ&&setShowDrop(true)} onBlur={()=>setTimeout(()=>setShowDrop(false),200)}
          placeholder={t.search}
          style={{width:"100%",padding:"11px 16px 11px 38px",borderRadius:24,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",boxSizing:"border-box",fontFamily:"inherit",color:T.text}}/>
      </div>
      {showDrop&&searchQ&&<div style={{position:"absolute",top:"100%",left:16,right:16,...iosCard(T,{zIndex:300,overflow:"hidden",maxHeight:280,overflowY:"auto"})}}>
        {dropRes.length===0?<div style={{padding:14,textAlign:"center",color:T.muted,fontSize:13}}>"{searchQ}" {t.noProducts}</div>
          :dropRes.map(p=><div key={p.id} onMouseDown={()=>onHit(p)} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",cursor:"pointer",borderBottom:`1px solid ${T.border}`}} onMouseEnter={e=>e.currentTarget.style.background="rgba(0,0,0,0.03)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <PImg p={p} size={36} r={8} T={T}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name_mm||p.name}</div><div style={{fontSize:11,color:T.muted}}>{p.category} · {fmt(calcUnit(p))}</div></div>
          </div>)}
      </div>}
    </div>
  </div>);
}

// ── Bottom Nav (iOS style) ────────────────────────────────────
function BottomNav({T,t,tab,setTab,cartCount,user}){
  const items=[{k:"home",icon:"🏠",label:t.home},{k:"cats",icon:"✨",label:t.cats},{k:"cart",icon:"🛒",label:t.cart,badge:cartCount},{k:"track",icon:"📋",label:t.orders},{k:"profile",icon:user?"👤":"🔐",label:t.profile}];
  return(<div style={{position:"fixed",bottom:0,left:0,right:0,background:T.navBg,backdropFilter:T.blur,WebkitBackdropFilter:T.blur,borderTop:`0.5px solid ${T.border}`,display:"flex",zIndex:190,paddingBottom:"env(safe-area-inset-bottom,0px)"}}>
    {items.map(it=><button key={it.k} onClick={()=>setTab(it.k)} style={{flex:1,padding:"8px 2px 6px",border:"none",background:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,position:"relative"}}>
      <div style={{width:34,height:34,borderRadius:10,background:tab===it.k?T.primary:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,transition:"all 0.2s"}}>
        {it.icon}
      </div>
      {it.badge>0&&<span style={{position:"absolute",top:4,right:"calc(50% - 18px)",background:T.primary,color:"#fff",borderRadius:10,padding:"1px 5px",fontSize:9,fontWeight:800}}>{it.badge}</span>}
      <span style={{fontSize:9.5,color:tab===it.k?T.primary:T.muted,fontWeight:tab===it.k?700:400}}>{it.label}</span>
    </button>)}
  </div>);
}


// ── Product Card ──────────────────────────────────────────────
function ProductCard({T,t,p,onClick,onAdd}){
  const[added,setAdded]=useState(false);const up=calcUnit(p);const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;
  const blobs=[T.blob1,T.blob2,T.blob3];const blob=blobs[(p?.id||"x").charCodeAt(1)%3||0];
  const doAdd=e=>{e.stopPropagation();if(!canBuy)return;onAdd(p);setAdded(true);setTimeout(()=>setAdded(false),1800);};
  return(<div onClick={()=>onClick(p)} style={{...iosCard(T,{overflow:"hidden",cursor:"pointer",transition:"all 0.2s",position:"relative"})}}>
    {dv>0&&<div style={{position:"absolute",top:8,left:8,background:T.primary,color:"#fff",borderRadius:8,padding:"3px 8px",fontSize:10,fontWeight:800,zIndex:2,boxShadow:`0 2px 8px ${T.primary}50`}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</div>}
    {p.preorder&&<div style={{position:"absolute",top:8,right:8,background:"rgba(255,149,0,0.85)",backdropFilter:"blur(10px)",color:"#fff",borderRadius:8,padding:"2px 7px",fontSize:9,fontWeight:700,zIndex:2}}>Pre-order</div>}
    {oos&&!p.preorder&&<div style={{position:"absolute",top:8,right:8,background:"rgba(0,0,0,0.45)",backdropFilter:"blur(10px)",color:"#fff",borderRadius:8,padding:"2px 7px",fontSize:9,zIndex:2}}>{t.tempOOS}</div>}
    <div style={{position:"relative",display:"flex",justifyContent:"center",alignItems:"center",padding:"18px 14px 10px",minHeight:148,overflow:"hidden"}}>
      <div style={{position:"absolute",inset:"-15%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",filter:"blur(10px)",opacity:0.65}}/>
      <div style={{position:"relative",zIndex:1,filter:"drop-shadow(0 6px 16px rgba(0,0,0,0.12))"}}><PImg p={p} size={108} r={14} T={T}/></div>
    </div>
    <div style={{padding:"8px 12px 13px"}}>
      <div style={{fontSize:10,color:T.muted,marginBottom:2,fontWeight:600,textTransform:"uppercase",letterSpacing:0.6}}>{p.category}</div>
      <div style={{fontSize:12.5,fontWeight:700,color:T.text,lineHeight:1.35,marginBottom:7,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",minHeight:34}}>{p.name_mm||p.name}</div>
      <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:10,flexWrap:"wrap"}}>
        <span style={{fontWeight:800,color:T.primary,fontSize:14}}>{fmt(up)}</span>
        {dv>0&&<span style={{fontSize:11,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
      </div>
      <button onClick={doAdd} disabled={!canBuy} style={{width:"100%",padding:"8px 0",borderRadius:12,border:"none",cursor:canBuy?"pointer":"not-allowed",background:added?T.success:(oos&&!p.preorder?"rgba(0,0,0,0.08)":T.primary),color:oos&&!p.preorder?T.muted:"#fff",fontSize:11.5,fontWeight:700,transition:"all 0.2s",boxShadow:canBuy&&!added?`0 3px 12px ${T.primary}40`:undefined}}>
        {added?"✓":(p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:`🛒 ${t.addCart}`))}
      </button>
    </div>
  </div>);
}

// ── Home Page ─────────────────────────────────────────────────
function HomePage({T,t,products,cats,catFilter,setCatFilter,onOpen,onAdd,banner,shopMM}){
  const featured=products.filter(p=>p.featured||(p.discount_value||0)>0).slice(0,8);
  const newArrivals=products.slice(0,4);
  const filtered=catFilter==="all"?products:products.filter(p=>p.category===catFilter);
  const catIcons={"Thai Skincare":"🇹🇭","Korean Beauty":"🇰🇷","Makeup":"💄","Hair Care":"💆","Fashion":"👗","Accessories":"💎","Others":"🛍️"};
  return(<div style={{paddingBottom:80}}>
    {banner?<div style={{margin:"10px 14px",borderRadius:20,overflow:"hidden",height:130,boxShadow:T.shadowMd}}><img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="banner"/></div>
    :<div style={{margin:"10px 14px",...iosCard(T,{padding:"18px 20px",textAlign:"center",background:`linear-gradient(135deg,${T.primary}15,${T.accent}15)`})}}>
      <img src={DEFAULT_LOGO} style={{width:56,height:56,borderRadius:28,marginBottom:8,boxShadow:T.shadow}} alt="logo"/>
      <div style={{fontSize:18,fontWeight:800,color:T.text,letterSpacing:-0.5}}>{shopMM||"Shwe Twin BKK Direct"}</div>
      <div style={{fontSize:11,color:T.muted,marginTop:3,letterSpacing:0.3}}>✈️ Bangkok Direct · Fast Delivery</div>
    </div>}
    {/* Categories */}
    <div style={{padding:"14px 14px 6px"}}>
      <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:10,letterSpacing:-0.3}}>Categories</div>
      <div style={{display:"flex",gap:10,overflowX:"auto",scrollbarWidth:"none",paddingBottom:4}}>
        {[{name:"all",label:t.all,icon:"⭐"},...cats.map(c=>({name:c.name,label:c.name,icon:catIcons[c.name]||"🛍️"}))].map(c=>(
          <div key={c.name} onClick={()=>setCatFilter(c.name)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:5,cursor:"pointer",flexShrink:0,transition:"all 0.15s"}}>
            <div style={{width:56,height:56,borderRadius:28,...ios(T,{}),display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,border:catFilter===c.name?`2px solid ${T.primary}`:`1px solid ${T.border}`,background:catFilter===c.name?T.primary:"rgba(255,255,255,0.6)",boxShadow:catFilter===c.name?`0 4px 16px ${T.primary}50`:T.shadow,transition:"all 0.2s"}}>
              <span style={{filter:catFilter===c.name?"brightness(10)":"none"}}>{c.icon}</span>
            </div>
            <span style={{fontSize:9.5,fontWeight:catFilter===c.name?700:500,color:catFilter===c.name?T.primary:T.muted,textAlign:"center",maxWidth:56,lineHeight:1.2}}>{c.label}</span>
          </div>
        ))}
      </div>
    </div>
    {/* Flash deals */}
    {featured.length>0&&catFilter==="all"&&<div style={{padding:"12px 14px 4px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontSize:14,fontWeight:800,color:T.text,letterSpacing:-0.3}}>🔥 Flash Deals</div><span style={{fontSize:11,color:T.primary,fontWeight:600}}>See all →</span></div>
      <div style={{display:"flex",gap:10,overflowX:"auto",scrollbarWidth:"none"}}>
        {featured.map(p=>{
          const blob=[T.blob1,T.blob2,T.blob3][(p.id||"x").charCodeAt(1)%3||0];
          return(<div key={p.id} onClick={()=>onOpen(p)} style={{flexShrink:0,width:126,...iosCard(T,{overflow:"hidden",cursor:"pointer",transition:"all 0.2s"})}}>
            <div style={{position:"relative",display:"flex",justifyContent:"center",padding:"10px 8px 6px",overflow:"hidden"}}>
              <div style={{position:"absolute",inset:"-15%",background:blob,borderRadius:"50%",filter:"blur(10px)",opacity:0.6}}/>
              <div style={{position:"relative",zIndex:1}}><PImg p={p} size={82} r={10} T={T}/></div>
            </div>
            <div style={{padding:"6px 10px 11px"}}>
              <div style={{fontSize:11,fontWeight:700,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",letterSpacing:-0.2}}>{p.name_mm||p.name}</div>
              {(p.discount_value||0)>0&&<div style={{fontSize:10,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</div>}
              <div style={{fontSize:13,fontWeight:800,color:T.primary}}>{fmt(calcUnit(p))}</div>
            </div>
          </div>);
        })}
      </div>
    </div>}
    {/* Products grid */}
    <div style={{padding:"12px 14px 0"}}>
      <div style={{fontSize:14,fontWeight:800,color:T.text,marginBottom:10,letterSpacing:-0.3}}>{catFilter==="all"?"All Products":`📦 ${catFilter}`}</div>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}><div style={{fontSize:52,marginBottom:10}}>🛍️</div><div style={{fontSize:15,color:T.text,fontWeight:600}}>{t.noProducts}</div></div>
        :<div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12}}>
          {filtered.map(p=><ProductCard key={p.id} T={T} t={t} p={p} onClick={onOpen} onAdd={onAdd}/>)}
        </div>}
    </div>
  </div>);
}

// ── Product Detail ────────────────────────────────────────────
function ProductDetail({T,t,p,onBack,addToCart,onBuyNow}){
  const[qty,setQty]=useState(1);const[added,setAdded]=useState(false);const[imgIdx,setImgIdx]=useState(0);
  const imgs=getImgs(p);let bd=p.bulk_discounts;if(typeof bd==="string"){try{bd=JSON.parse(bd);}catch{bd=[];}}
  const bulks=(Array.isArray(bd)?bd:[]).filter(b=>b.min_qty>0).sort((a,b)=>a.min_qty-b.min_qty);
  const dv=p.discount_value||0;const oos=isOOS(p);const canBuy=!oos||p.preorder;const up=calcUnit(p,qty);
  const blob=[T.blob1,T.blob2,T.blob3][(p?.id||"x").charCodeAt(1)%3||0];
  const Sec=({title,txt,warn})=>txt?<div style={{marginBottom:12}}><div style={{fontSize:11,fontWeight:700,color:warn?T.danger:T.muted,marginBottom:4,textTransform:"uppercase",letterSpacing:0.5}}>{title}</div><div style={{fontSize:13,color:T.text,lineHeight:1.8,...ios(T,{borderRadius:12,padding:"11px 14px",background:warn?T.dangerBg:T.surface2})}}>{txt}</div></div>:null;
  return(<div style={{minHeight:"100vh",paddingBottom:100}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0,fontWeight:300}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",letterSpacing:-0.3}}>{p.name_mm||p.name}</span>
    </div>
    {/* Hero image with blob */}
    <div style={{position:"relative",display:"flex",justifyContent:"center",alignItems:"center",padding:"28px 20px 16px",minHeight:270,overflow:"hidden"}}>
      <div style={{position:"absolute",width:"75%",height:"75%",background:blob,borderRadius:"60% 40% 55% 45% / 45% 55% 45% 55%",filter:"blur(24px)",opacity:0.6}}/>
      <div style={{position:"absolute",width:"45%",height:"45%",background:T.blob2,borderRadius:"40% 60% 45% 55%",filter:"blur(16px)",opacity:0.4,transform:"translate(35%,25%)"}}/>
      <div style={{position:"relative",zIndex:1,filter:"drop-shadow(0 12px 32px rgba(0,0,0,0.18))"}}>
        {imgs[imgIdx]?<img src={imgs[imgIdx]} style={{width:220,height:220,objectFit:"contain",borderRadius:20}} alt={p.name} onError={e=>e.target.style.display="none"}/>:<PImg p={p} size={220} r={20} T={T}/>}
      </div>
    </div>
    {imgs.length>1&&<div style={{display:"flex",gap:8,padding:"0 16px 12px",justifyContent:"center"}}>
      {imgs.map((img,i)=><div key={i} onClick={()=>setImgIdx(i)} style={{width:48,height:48,borderRadius:10,overflow:"hidden",border:`2px solid ${i===imgIdx?T.primary:T.border}`,cursor:"pointer",transition:"all 0.15s"}}><img src={img} style={{width:"100%",height:"100%",objectFit:"cover"}} alt=""/></div>)}
    </div>}
    <div style={{padding:"0 14px"}}>
      <div style={{...iosCard(T,{padding:"16px",marginBottom:10})}}>
        <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
          <span style={{background:T.tag,color:T.tagText,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700}}>{p.category}</span>
          {p.preorder&&<span style={{background:T.warningBg,color:T.warning,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700}}>Pre-order</span>}
          {(p.stock||0)>0&&!p.preorder&&<span style={{background:T.successBg,color:T.success,padding:"3px 10px",borderRadius:20,fontSize:11}}>{t.inStock}: {p.stock}</span>}
          {oos&&!p.preorder&&<span style={{background:"rgba(0,0,0,0.06)",color:T.muted,padding:"3px 10px",borderRadius:20,fontSize:11}}>⚠️ {t.tempOOS}</span>}
        </div>
        <div style={{fontSize:20,fontWeight:800,color:T.text,marginBottom:3,letterSpacing:-0.5,lineHeight:1.25}}>{p.name_mm||p.name}</div>
        {p.name_mm&&<div style={{fontSize:13,color:T.muted,marginBottom:12,letterSpacing:-0.2}}>{p.name}</div>}
        <div style={{display:"flex",alignItems:"baseline",gap:10,flexWrap:"wrap"}}>
          <span style={{fontSize:26,fontWeight:800,color:T.primary,letterSpacing:-0.5}}>{fmt(up)}</span>
          {dv>0&&<span style={{fontSize:15,color:T.muted,textDecoration:"line-through"}}>{fmt(p.price)}</span>}
          {dv>0&&<span style={{background:T.primary,color:"#fff",borderRadius:8,padding:"3px 9px",fontSize:11,fontWeight:700,boxShadow:`0 2px 8px ${T.primary}50`}}>{p.discount_type==="fixed"?`-${fmt(dv)}`:`-${dv}%`}</span>}
        </div>
      </div>
      {bulks.length>0&&<div style={{...iosCard(T,{padding:"12px 14px",marginBottom:10,background:T.tag})}}>
        <div style={{fontSize:11,fontWeight:700,color:T.primary,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>🎁 {t.bulkTiers}</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>{bulks.map((b,i)=><div key={i} style={{background:qty>=b.min_qty?T.primary:"rgba(233,30,140,0.1)",color:qty>=b.min_qty?"#fff":T.primary,borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:600,transition:"all 0.2s",border:`1px solid ${qty>=b.min_qty?T.primary:"rgba(233,30,140,0.2)"}`}}>{b.min_qty}+ → -{b.discount_percent}%</div>)}</div>
      </div>}
      {canBuy&&<div style={{...iosCard(T,{padding:"12px 16px",marginBottom:10})}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <span style={{fontSize:11,color:T.muted,fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{t.qty}</span>
          <div style={{display:"flex",alignItems:"center",gap:0,...ios(T,{borderRadius:14,overflow:"hidden"})}}>
            <button onClick={()=>setQty(q=>Math.max(1,q-1))} style={{width:38,height:38,border:"none",...ios(T,{borderRadius:0}),cursor:"pointer",fontSize:20,fontWeight:300,color:T.primary}}>−</button>
            <span style={{minWidth:36,textAlign:"center",fontWeight:800,fontSize:17,color:T.text,borderLeft:`0.5px solid ${T.border}`,borderRight:`0.5px solid ${T.border}`,height:38,display:"flex",alignItems:"center",justifyContent:"center"}}>{qty}</span>
            <button onClick={()=>setQty(q=>q+1)} style={{width:38,height:38,border:"none",...ios(T,{borderRadius:0}),cursor:"pointer",fontSize:20,fontWeight:300,color:T.primary}}>+</button>
          </div>
          <span style={{fontSize:13,color:T.muted}}>= <b style={{color:T.primary,fontSize:14}}>{fmt(up*qty)}</b></span>
        </div>
      </div>}
      <div style={{...iosCard(T,{padding:"16px",marginBottom:10})}}>
        <Sec title={`📝 ${t.description}`} txt={p.description}/>
        <Sec title={`👤 ${t.suitableFor}`} txt={p.suitable_for}/>
        <Sec title={`✅ ${t.benefits}`} txt={p.benefits}/>
        <Sec title={`📋 ${t.usage}`} txt={p.usage_info}/>
        <Sec title={`⚠️ ${t.warning}`} txt={p.warning} warn/>
      </div>
      {p.video_url&&<div style={{...iosCard(T,{padding:"14px",marginBottom:10})}}><div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>📹 Video</div><video src={p.video_url} controls style={{width:"100%",borderRadius:12,background:"#000"}}/></div>}
    </div>
    <div style={{position:"fixed",bottom:0,left:0,right:0,...ios(T,{padding:"12px 16px",boxShadow:`0 -1px 0 ${T.border},0 -8px 24px rgba(0,0,0,0.08)`,display:"flex",gap:10,zIndex:190,background:T.navBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderBottom:"none"}),paddingBottom:"calc(12px + env(safe-area-inset-bottom,0px))"}}>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);setAdded(true);setTimeout(()=>setAdded(false),1800);}} disabled={!canBuy}
        style={{flex:1,padding:"14px 0",borderRadius:14,...ios(T,{border:`2px solid ${canBuy?T.primary:T.border}`}),background:"transparent",color:canBuy?T.primary:T.muted,cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700,transition:"all 0.2s"}}>
        {added?"✓ Added!":t.addCart}
      </button>
      <button onClick={()=>{if(!canBuy)return;addToCart(p,qty);onBuyNow();}} disabled={!canBuy}
        style={{flex:1,padding:"14px 0",borderRadius:14,border:"none",background:canBuy?T.primary:"rgba(0,0,0,0.1)",color:"#fff",cursor:canBuy?"pointer":"not-allowed",fontSize:14,fontWeight:700,boxShadow:canBuy?`0 4px 16px ${T.primary}50`:undefined,transition:"all 0.2s"}}>
        {p.preorder?`📋 ${t.preorder}`:(oos?t.tempOOS:t.buyNow||"Buy Now")}
      </button>
    </div>
  </div>);
}

// ── Cart Page ─────────────────────────────────────────────────
function CartPage({T,t,cart,updateQty,removeItem,total,onCheckout,onBack}){
  if(!cart.length)return(<div style={{minHeight:"100vh"}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,color:T.text,letterSpacing:-0.3}}>{t.cart}</span>
    </div>
    <div style={{textAlign:"center",padding:"100px 20px",color:T.muted}}><div style={{fontSize:64,marginBottom:14}}>🛒</div><div style={{fontSize:17,fontWeight:700,color:T.text,letterSpacing:-0.3}}>{t.emptyCart}</div></div>
  </div>);
  return(<div style={{minHeight:"100vh",paddingBottom:100}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,color:T.text,letterSpacing:-0.3}}>{t.cart} ({cart.length})</span>
    </div>
    <div style={{padding:"12px 14px 0",display:"flex",flexDirection:"column",gap:10}}>
      {cart.map(item=>{const up=calcUnit(item.p,item.qty);return(
        <div key={item.p.id} style={{...iosCard(T,{padding:"13px",display:"flex",gap:12,alignItems:"flex-start"})}}>
          <PImg p={item.p} size={66} r={12} T={T}/>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:13,fontWeight:700,color:T.text,lineHeight:1.3,marginBottom:2,letterSpacing:-0.2}}>{item.p.name_mm||item.p.name}</div>
            {item.p.preorder&&<span style={{fontSize:9,background:T.warningBg,color:T.warning,padding:"1px 6px",borderRadius:8,fontWeight:700,display:"inline-block",marginBottom:4}}>Pre-order</span>}
            <div style={{fontSize:14,color:T.primary,fontWeight:800,marginBottom:9}}>{fmt(up)}</div>
            <div style={{display:"flex",alignItems:"center",gap:0,...ios(T,{borderRadius:10,overflow:"hidden",display:"inline-flex"})}}>
              <button onClick={()=>updateQty(item.p.id,item.qty-1)} style={{width:30,height:30,border:"none",...ios(T,{borderRadius:0}),cursor:"pointer",fontWeight:300,color:T.primary,fontSize:18}}>−</button>
              <span style={{minWidth:28,textAlign:"center",fontWeight:800,fontSize:14,color:T.text,borderLeft:`0.5px solid ${T.border}`,borderRight:`0.5px solid ${T.border}`,height:30,display:"flex",alignItems:"center",justifyContent:"center"}}>{item.qty}</span>
              <button onClick={()=>updateQty(item.p.id,item.qty+1)} style={{width:30,height:30,border:"none",...ios(T,{borderRadius:0}),cursor:"pointer",fontWeight:300,color:T.primary,fontSize:18}}>+</button>
            </div>
            <span style={{marginLeft:8,fontSize:13,fontWeight:700,color:T.text}}>{fmt(up*item.qty)}</span>
          </div>
          <button onClick={()=>removeItem(item.p.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:T.muted,padding:0,lineHeight:1}}>✕</button>
        </div>
      );})}
    </div>
    <div style={{position:"sticky",bottom:0,...ios(T,{padding:"13px 16px",boxShadow:`0 -1px 0 ${T.border}`,display:"flex",gap:12,alignItems:"center",marginTop:12,background:T.navBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderBottom:"none"}),paddingBottom:"calc(13px + env(safe-area-inset-bottom,0px))"}}>
      <div style={{flex:1}}><div style={{fontSize:10,color:T.muted,textTransform:"uppercase",letterSpacing:0.5}}>{t.total}</div><div style={{fontSize:22,fontWeight:800,color:T.primary,letterSpacing:-0.5}}>{fmt(total)}</div></div>
      <button onClick={onCheckout} style={{padding:"14px 28px",borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`,letterSpacing:-0.3}}>{t.checkout} →</button>
    </div>
  </div>);
}

// ── Checkout ──────────────────────────────────────────────────
function CheckoutPage({T,t,cart,total,settings,onPlaced,onBack,user}){
  const[name,setName]=useState(user?.name||"");const[phone,setPhone]=useState(user?.phone||"");const[address,setAddress]=useState("");const[note,setNote]=useState("");const[sent,setSent]=useState(false);const[sentOrder,setSentOrder]=useState(null);
  const hasPreorder=cart.some(i=>i.p.preorder);
  const msgLines=cart.map(i=>`• ${i.p.name_mm||i.p.name}${i.p.preorder?" [PRE]":""} x${i.qty} = ${fmt(calcUnit(i.p,i.qty)*i.qty)}`).join("\n");
  const msg=encodeURIComponent(`မင်္ဂလာပါ!\nName: ${name}\nPhone: ${phone}\nAddress: ${address}\n\nOrder:\n${msgLines}\nTotal: ${fmt(total)}${hasPreorder?"\n⚠️ Pre-order ပါဝင်":""}${note?"\nNote: "+note:""}`);
  const saveOrder=async method=>{
    if(sent)return;
    const newOrd={order_number:onum(),customer_name:name,customer_phone:phone,customer_address:address,customer_note:note,contact_method:method,items:cart.map(i=>({id:i.p.id,name:i.p.name,name_mm:i.p.name_mm,qty:i.qty,is_preorder:i.p.preorder,unit_price:calcUnit(i.p,i.qty),total:calcUnit(i.p,i.qty)*i.qty})),total,status:"pending",deposit_paid:0,balance_due:total,stock_deducted:false,admin_note:"",user_id:user?.id||"",created_at:new Date().toISOString()};
    try{const rows=await sb("POST","orders",newOrd);const saved=Array.isArray(rows)?rows[0]:newOrd;setSentOrder(saved);setSent(true);onPlaced(saved);}catch(e){alert("Error: "+e.message);}
  };
  const contacts=[
    {label:"Facebook",sub:"Messenger",color:"#1877F2",emoji:"💬",key:"messenger",url:`https://m.me/${(settings.fb_link||"").replace(/^https?:\/\/m\.me\//,"")}?text=${msg}`},
    {label:"Viber",sub:settings.viber_num,color:"#7360F2",emoji:"📱",key:"viber",url:`viber://chat?number=${encodeURIComponent(settings.viber_num||"")}&text=${msg}`},
    {label:"WhatsApp",sub:settings.wa_num,color:"#25D366",emoji:"💚",key:"whatsapp",url:`https://wa.me/${(settings.wa_num||"").replace(/\D/g,"")}?text=${msg}`},
    {label:"Phone",sub:settings.phone_num,color:"#546E7A",emoji:"📞",key:"phone",url:`tel:${settings.phone_num}`},
  ];
  const inp={width:"100%",padding:"12px 14px",borderRadius:12,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10,fontFamily:"inherit",color:T.text};
  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,color:T.text,letterSpacing:-0.3}}>{t.checkout}</span>
    </div>
    {sent?<div style={{padding:20}}><div style={{...iosCard(T,{padding:32,textAlign:"center"})}}>
      <div style={{fontSize:56,marginBottom:12}}>✅</div>
      <div style={{fontSize:20,fontWeight:800,color:T.primary,marginBottom:8,letterSpacing:-0.5}}>{t.orderPlaced}</div>
      <div style={{fontSize:13,color:T.muted,marginBottom:16}}>{t.orderPlacedMsg}</div>
      {sentOrder&&<div style={{...ios(T,{borderRadius:12,padding:"10px 14px",background:T.surface2}),fontSize:12,color:T.text}}>Order #: <b>{sentOrder.order_number}</b></div>}
    </div></div>
    :<div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
      <div style={{...iosCard(T,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10,letterSpacing:-0.3}}>{t.orderSummary}</div>
        {hasPreorder&&<div style={{background:T.warningBg,borderRadius:10,padding:"8px 12px",marginBottom:10,fontSize:12,color:T.warning,fontWeight:600}}>⚠️ Pre-order items ပါဝင်</div>}
        {cart.map(item=><div key={item.p.id} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`0.5px solid ${T.border}`,fontSize:13}}><span style={{color:T.text,flex:1,paddingRight:10}}>{item.p.name_mm||item.p.name} × {item.qty}</span><span style={{fontWeight:700,color:T.text,whiteSpace:"nowrap"}}>{fmt(calcUnit(item.p,item.qty)*item.qty)}</span></div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"12px 0 0",fontSize:17,fontWeight:800}}><span style={{color:T.text,letterSpacing:-0.3}}>{t.total}</span><span style={{color:T.primary,letterSpacing:-0.5}}>{fmt(total)}</span></div>
      </div>
      <div style={{...iosCard(T,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:10,letterSpacing:-0.3}}>📋 Customer Info</div>
        <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder={t.yourName}/>
        <input style={inp} value={phone} onChange={e=>setPhone(e.target.value)} placeholder={t.yourPhone} type="tel"/>
        <input style={inp} value={address} onChange={e=>setAddress(e.target.value)} placeholder={t.yourAddress}/>
        <textarea value={note} onChange={e=>setNote(e.target.value)} placeholder={t.note} style={{...inp,minHeight:56,resize:"vertical",marginBottom:0}}/>
      </div>
      {(!name.trim()||!phone.trim())&&<div style={{...ios(T,{borderRadius:12,padding:"10px 14px",background:T.warningBg}),fontSize:12,color:T.warning}}>⚠️ Name & Phone ထည့်ပြီးမှ order လုပ်ပါ</div>}
      <div style={{...iosCard(T,{padding:16})}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,marginBottom:12,letterSpacing:-0.3}}>{t.contactOrder}</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {contacts.map(ct=><a key={ct.key} href={(name.trim()&&phone.trim())?ct.url:"#"} onClick={e=>{if(!name.trim()||!phone.trim()){e.preventDefault();alert("Name & Phone ထည့်ပါ");return;}saveOrder(ct.key);}} target="_blank" rel="noopener"
            style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",background:ct.color,borderRadius:14,textDecoration:"none",color:"#fff",boxShadow:`0 4px 14px ${ct.color}55`,opacity:(name.trim()&&phone.trim())?1:0.5,transition:"all 0.15s"}}>
            <span style={{fontSize:22}}>{ct.emoji}</span><div><div style={{fontWeight:700,fontSize:14}}>{ct.label}</div><div style={{fontSize:11,opacity:0.85}}>{ct.sub}</div></div><span style={{marginLeft:"auto",fontSize:18,opacity:0.7}}>→</span>
          </a>)}
        </div>
      </div>
    </div>}
  </div>);
}

// ── Track Order ───────────────────────────────────────────────
function TrackOrder({T,t}){
  const[q,setQ]=useState("");const[result,setResult]=useState(null);const[busy,setBusy]=useState(false);const[err,setErr]=useState("");
  const search=async()=>{if(!q.trim())return;setBusy(true);setErr("");setResult(null);try{const rows=await sb("GET","orders",null,`order_number=eq.${q.trim()}`);if(!rows||rows.length===0)setErr("Order မတွေ့ပါ");else setResult(rows[0]);}catch(e){setErr(e.message);}setBusy(false);};
  return(<div style={{padding:"16px 14px",paddingBottom:90}}>
    <div style={{fontSize:18,fontWeight:800,color:T.text,letterSpacing:-0.5,marginBottom:4}}>📋 {t.trackOrder}</div>
    <div style={{fontSize:13,color:T.muted,marginBottom:14}}>Order number ဖြင့် status စစ်ဆေးနိုင်သည်</div>
    <div style={{display:"flex",gap:8,marginBottom:14}}>
      <input value={q} onChange={e=>setQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&search()} placeholder={t.enterOrderNo}
        style={{flex:1,padding:"12px 14px",borderRadius:14,...ios(T,{background:T.surfaceStrong}),fontSize:14,outline:"none",fontFamily:"inherit",color:T.text,minWidth:0,boxSizing:"border-box"}}/>
      <button onClick={search} disabled={busy} style={{padding:"12px 18px",borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0,boxShadow:`0 4px 12px ${T.primary}50`}}>{busy?"...":t.track}</button>
    </div>
    {err&&<div style={{...ios(T,{borderRadius:12,padding:"12px 14px",background:T.dangerBg,marginBottom:12}),fontSize:13,color:T.danger}}>⚠️ {err}</div>}
    {result&&<div style={{...iosCard(T,{padding:20})}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}><div><div style={{fontSize:16,fontWeight:800,color:T.text,letterSpacing:-0.5}}>{result.order_number}</div><div style={{fontSize:12,color:T.muted}}>{fdate(result.created_at)}</div></div><SBadge status={result.status} t={t}/></div>
      {(result.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`0.5px solid ${T.border}`,fontSize:13}}><span style={{color:T.text}}>{item.name_mm||item.name} × {item.qty}</span><span style={{fontWeight:700,color:T.text}}>{fmt(item.total)}</span></div>)}
      <div style={{paddingTop:12,borderTop:`0.5px solid ${T.border}`,marginTop:4}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:T.text,letterSpacing:-0.3}}>{t.total}</span><span style={{color:T.primary,letterSpacing:-0.5}}>{fmt(result.total)}</span></div>
        {(result.deposit_paid||0)>0&&<><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{color:T.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:T.success,fontWeight:700}}>{fmt(result.deposit_paid)}</span></div><div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:T.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:T.danger,fontWeight:700}}>{fmt(Math.max(0,result.total-(result.deposit_paid||0)))}</span></div></>}
      </div>
      {result.admin_note&&<div style={{...ios(T,{borderRadius:10,padding:"10px 12px",marginTop:12,background:T.surface2}),fontSize:12,color:T.textSub}}><b>Note:</b> {result.admin_note}</div>}
    </div>}
  </div>);
}

// ── Profile Page ──────────────────────────────────────────────
function ProfilePage({T,t,user,onSignIn,onSignOut}){
  return(<div style={{padding:"16px 14px",paddingBottom:90}}>
    <div style={{fontSize:18,fontWeight:800,color:T.text,letterSpacing:-0.5,marginBottom:16}}>👤 {t.profile}</div>
    {user?<>
      <div style={{...iosCard(T,{padding:20,marginBottom:14,textAlign:"center"})}}>
        <div style={{width:64,height:64,borderRadius:32,background:T.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,margin:"0 auto 10px",boxShadow:`0 4px 16px ${T.primary}50`}}>👤</div>
        <div style={{fontSize:17,fontWeight:700,color:T.text,letterSpacing:-0.3}}>{user.name||user.email||user.phone||"User"}</div>
        <div style={{fontSize:12,color:T.muted,marginTop:3}}>{user.provider==="phone"?`📱 ${user.phone}`:user.provider==="google"?"G Google Account":user.email||""}</div>
      </div>
      <div style={{...iosCard(T,{marginBottom:14})}}>
        {[{icon:"📦",label:t.myOrders},{icon:"📋",label:t.trackOrder},{icon:"⚙️",label:t.settings}].map((item,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:12,padding:"14px 16px",borderBottom:i<2?`0.5px solid ${T.border}`:"none",cursor:"pointer"}} onClick={()=>{}}>
          <span style={{fontSize:20}}>{item.icon}</span><span style={{flex:1,fontSize:14,fontWeight:600,color:T.text,letterSpacing:-0.2}}>{item.label}</span><span style={{color:T.muted,fontSize:16}}>›</span>
        </div>)}
      </div>
      <button onClick={onSignOut} style={{width:"100%",padding:14,borderRadius:14,...ios(T,{border:`1px solid ${T.danger}50`,background:T.dangerBg}),color:T.danger,cursor:"pointer",fontSize:14,fontWeight:600}}>{t.signOut}</button>
    </>:<>
      <div style={{...iosCard(T,{padding:28,textAlign:"center",marginBottom:20})}}>
        <img src={DEFAULT_LOGO} style={{width:72,height:72,borderRadius:36,marginBottom:12,boxShadow:T.shadowMd}} alt="logo"/>
        <div style={{fontSize:16,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:4}}>{t.joinUs}</div>
        <div style={{fontSize:13,color:T.muted,marginBottom:20}}>{t.authSubtitle}</div>
        <button onClick={onSignIn} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,boxShadow:`0 4px 16px ${T.primary}50`}}>{t.signIn} / {t.signUp}</button>
      </div>
    </>}
  </div>);
}


// ── Admin components (compact glass iOS style) ────────────────
function AdminLoginBox({T,t,adminPw,onSuccess,onBack}){
  const[pw,setPw]=useState("");const[err,setErr]=useState("");
  const check=()=>{if(pw===(adminPw||"admin123"))onSuccess();else setErr(t.wrongPw);};
  return(<>
    <input type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&check()} placeholder="admin123"
      style={{width:"100%",padding:"13px 16px",borderRadius:14,...ios(T,{background:T.surfaceStrong,border:`1.5px solid ${err?T.danger:T.border}`}),fontSize:14,outline:"none",boxSizing:"border-box",marginBottom:err?8:16,fontFamily:"inherit",color:T.text}}/>
    {err&&<div style={{color:T.danger,fontSize:12,marginBottom:14,...ios(T,{borderRadius:10,padding:"8px 12px",background:T.dangerBg})}}>{err}</div>}
    <button onClick={check} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:10,boxShadow:`0 4px 16px ${T.primary}50`}}>{t.login}</button>
    <button onClick={onBack} style={{width:"100%",padding:12,borderRadius:14,...ios(T,{border:`1px solid ${T.border}`}),background:"transparent",color:T.muted,cursor:"pointer",fontSize:13}}>{t.backToShop}</button>
  </>);
}

function SBadge2({status,t}){const c=SC[status]||SC.pending;const L={pending:t.pending,confirmed:t.confirmed,deposited:t.deposited,completed:t.completed,cancelled:t.cancelled};return<span style={{background:c.bg,color:c.color,border:`1px solid ${c.border}`,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{L[status]||status}</span>;}

function OrderMgr({T,t,products,setProducts}){
  const[orders,setOrders]=useState([]);const[sel,setSel]=useState(null);const[dep,setDep]=useState("");const[noteVal,setNoteVal]=useState("");const[busy,setBusy]=useState(false);const[filter,setFilter]=useState("all");
  const load=useCallback(async()=>{try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){}},[]);
  useEffect(()=>{load();},[load]);
  const upd=async(id,fields)=>{await sb("PATCH","orders",{...fields,updated_at:new Date().toISOString()},`id=eq.${id}`);setOrders(ords=>ords.map(o=>o.id===id?{...o,...fields}:o));if(sel?.id===id)setSel(o=>({...o,...fields}));};
  const confirm_=async()=>{
    if(!sel||sel.stock_deducted)return;if(!window.confirm("Stock နှုတ်မည်?"))return;
    setBusy(true);try{const up=[...products];for(const item of(sel.items||[])){if(item.is_preorder)continue;const i=up.findIndex(p=>p.id===item.id);if(i<0)continue;const ns=Math.max(0,(up[i].stock||0)-(item.qty||0));await sb("PATCH","products",{stock:ns,updated_at:new Date().toISOString()},`id=eq.${up[i].id}`);up[i]={...up[i],stock:ns};NS.add(`Stock: ${up[i].name_mm||up[i].name} -${item.qty}→${ns}`);if(ns===0)NS.add(`⚠️ ${up[i].name_mm||up[i].name} ${t.tempOOS}`);}setProducts(up);await upd(sel.id,{status:"confirmed",stock_deducted:true});}catch(e){alert(e.message);}setBusy(false);
  };
  const statuses=["pending","confirmed","deposited","completed","cancelled"];
  const filtered=filter==="all"?orders:orders.filter(o=>o.status===filter);
  const inp_={width:"100%",padding:"10px 12px",borderRadius:12,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:10,fontFamily:"inherit",color:T.text};
  if(sel){const dp=Number(sel.deposit_paid||0);const bal=Math.max(0,sel.total-dp);return(<div style={{paddingBottom:80}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={()=>{setSel(null);load();}} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,color:T.text,letterSpacing:-0.3}}>Order Detail</span>
    </div>
    <div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
      <div style={{...iosCard(T,{padding:16})}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}><div><div style={{fontSize:15,fontWeight:800,color:T.text,letterSpacing:-0.5}}>{sel.order_number}</div><div style={{fontSize:12,color:T.muted}}>{fdate(sel.created_at)}</div></div><SBadge2 status={sel.status} t={t}/></div>
        <div style={{...ios(T,{borderRadius:10,padding:"10px 12px",marginBottom:10,background:T.surface2}),fontSize:12}}>
          <div style={{color:T.text}}><b>Name:</b> {sel.customer_name} | <b>Phone:</b> {sel.customer_phone}</div>
          {sel.customer_address&&<div style={{color:T.textSub,marginTop:2}}><b>Address:</b> {sel.customer_address}</div>}
          {sel.contact_method&&<div style={{color:T.muted,marginTop:2}}><b>Via:</b> {sel.contact_method}</div>}
        </div>
        {(sel.items||[]).some(i=>i.is_preorder)&&<div style={{...ios(T,{borderRadius:8,padding:"7px 12px",marginBottom:10,background:T.warningBg}),fontSize:11,color:T.warning,fontWeight:700}}>⚠️ Pre-order items ပါဝင်</div>}
        {(sel.items||[]).map((item,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`0.5px solid ${T.border}`,fontSize:13}}><div><span style={{color:T.text,fontWeight:600}}>{item.name_mm||item.name}</span>{item.is_preorder&&<span style={{fontSize:9,background:T.warningBg,color:T.warning,padding:"1px 5px",borderRadius:5,fontWeight:700,marginLeft:4}}>PRE</span>}<span style={{color:T.muted}}> × {item.qty}</span></div><span style={{fontWeight:700,color:T.text}}>{fmt(item.total)}</span></div>)}
        <div style={{paddingTop:12,borderTop:`0.5px solid ${T.border}`,marginTop:4}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:16,fontWeight:800,marginBottom:8}}><span style={{color:T.text,letterSpacing:-0.3}}>{t.total}</span><span style={{color:T.primary,letterSpacing:-0.5}}>{fmt(sel.total)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{color:T.success,fontWeight:600}}>✅ {t.depositPaid}</span><span style={{color:T.success,fontWeight:700}}>{fmt(dp)}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:T.danger,fontWeight:600}}>⏳ {t.balanceDue}</span><span style={{color:T.danger,fontWeight:700}}>{fmt(bal)}</span></div>
        </div>
        {sel.customer_note&&<div style={{...ios(T,{borderRadius:8,padding:"8px 12px",marginTop:10,background:T.surface2}),fontSize:12,color:T.textSub}}><b>Note:</b> {sel.customer_note}</div>}
      </div>
      {!sel.stock_deducted?<button onClick={confirm_} disabled={busy} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:busy?"rgba(0,0,0,0.15)":T.success,color:"#fff",cursor:busy?"wait":"pointer",fontSize:14,fontWeight:700}}>{busy?"⏳ ...":t.confirmOrder}</button>:<div style={{...ios(T,{borderRadius:12,padding:"10px 14px",background:T.successBg}),fontSize:13,color:T.success,fontWeight:600,textAlign:"center"}}>✅ Stock deducted</div>}
      <div style={{...iosCard(T,{padding:14})}}><div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.updateStatus}</div><div style={{display:"flex",flexWrap:"wrap",gap:8}}>{statuses.map(s=><button key={s} onClick={()=>upd(sel.id,{status:s})} style={{padding:"5px 10px",borderRadius:20,...ios(T,{border:`1.5px solid ${sel.status===s?T.primary:T.border}`}),background:sel.status===s?T.primary:"transparent",color:sel.status===s?"#fff":T.text,cursor:"pointer"}}><SBadge2 status={s} t={t}/></button>)}</div></div>
      <div style={{...iosCard(T,{padding:14})}}><div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.updateDeposit}</div><div style={{display:"flex",gap:8}}><input type="number" min="0" value={dep} onChange={e=>setDep(e.target.value)} placeholder="0 Ks" style={{flex:1,padding:"10px 12px",borderRadius:12,...ios(T,{background:T.surfaceStrong}),fontSize:14,outline:"none",color:T.text,fontFamily:"inherit",minWidth:0}}/><button onClick={async()=>{const d=Number(dep);await upd(sel.id,{deposit_paid:d,balance_due:Math.max(0,sel.total-d)});setDep("");}} style={{padding:"10px 16px",borderRadius:12,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,flexShrink:0}}>OK</button></div></div>
      <div style={{...iosCard(T,{padding:14})}}><div style={{fontSize:11,fontWeight:700,color:T.muted,marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>{t.adminNote}</div><textarea value={noteVal||sel.admin_note||""} onChange={e=>setNoteVal(e.target.value)} style={{...inp_,minHeight:68,resize:"vertical"}}/><button onClick={()=>upd(sel.id,{admin_note:noteVal||sel.admin_note||""})} style={{width:"100%",padding:10,borderRadius:12,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>{t.save}</button></div>
    </div>
  </div>);}
  return(<div style={{paddingBottom:80}}>
    <div style={{display:"flex",gap:6,padding:"10px 12px",overflowX:"auto",scrollbarWidth:"none",...ios(T,{borderBottom:`0.5px solid ${T.border}`,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      {["all",...statuses].map(s=><button key={s} onClick={()=>setFilter(s)} style={{flexShrink:0,padding:"6px 12px",borderRadius:20,...ios(T,{border:`1.5px solid ${filter===s?T.primary:T.border}`}),background:filter===s?T.primary:"transparent",color:filter===s?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:600}}>{s==="all"?t.all:<SBadge2 status={s} t={t}/>}</button>)}
    </div>
    <div style={{padding:"10px 14px",display:"flex",flexDirection:"column",gap:8}}>
      {filtered.length===0?<div style={{textAlign:"center",padding:"60px 20px",color:T.muted}}><div style={{fontSize:40,marginBottom:8}}>📦</div><div>{t.noData}</div></div>
        :filtered.map(o=><div key={o.id} onClick={()=>{setSel(o);setNoteVal(o.admin_note||"");setDep("");}} style={{...iosCard(T,{padding:"13px 14px",cursor:"pointer",transition:"all 0.15s"})}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:5}}><div><div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3}}>{o.order_number}</div><div style={{fontSize:11,color:T.muted}}>{fdate(o.created_at)} · {o.customer_name} · {o.contact_method}</div></div><SBadge2 status={o.status} t={t}/></div>
          {(o.items||[]).some(i=>i.is_preorder)&&<div style={{fontSize:10,color:T.warning,fontWeight:700,marginBottom:3}}>⚠️ PREORDER</div>}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:12,color:T.muted}}>{(o.items||[]).length} items{o.stock_deducted?" · ✅":""}</div><div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:800,color:T.primary,letterSpacing:-0.3}}>{fmt(o.total)}</div>{(o.deposit_paid||0)>0&&<div style={{fontSize:10,color:T.success}}>Dep: {fmt(o.deposit_paid)}</div>}</div></div>
        </div>)}
    </div>
  </div>);
}

function StockMgr({T,t,products,setProducts}){
  const[adding,setAdding]=useState({});const[vals,setVals]=useState({});
  const doAdd=async p=>{const n=Number(vals[p.id]||0);if(!n||n<=0){alert("qty ထည့်ပါ");return;}const ns=(p.stock||0)+n;try{await sb("PATCH","products",{stock:ns,preorder:false,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,stock:ns,preorder:false}:x));NS.add(`${t.stockAdded}: ${p.name_mm||p.name} +${n}`);setAdding(v=>({...v,[p.id]:false}));setVals(v=>({...v,[p.id]:""}));}catch(e){alert(e.message);}};
  const togPO=async p=>{const np=!p.preorder;try{await sb("PATCH","products",{preorder:np,updated_at:new Date().toISOString()},`id=eq.${p.id}`);setProducts(prev=>prev.map(x=>x.id===p.id?{...x,preorder:np}:x));}catch(e){alert(e.message);}};
  const oos_=products.filter(p=>isOOS(p));const low=products.filter(p=>(p.stock||0)<=5&&!isOOS(p));
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    {oos_.length>0&&<div style={{...ios(T,{borderRadius:14,padding:"12px 14px",marginBottom:10,background:T.dangerBg}),border:`1px solid ${T.danger}30`}}><div style={{fontSize:13,fontWeight:700,color:T.danger,marginBottom:4}}>⚠️ Out of Stock ({oos_.length})</div>{oos_.map(p=><div key={p.id} style={{fontSize:12,color:T.danger}}>• {p.name_mm||p.name}{p.preorder?" [Pre-order]":""}</div>)}</div>}
    {low.length>0&&<div style={{...ios(T,{borderRadius:14,padding:"12px 14px",marginBottom:10,background:T.warningBg}),border:`1px solid ${T.warning}30`}}><div style={{fontSize:13,fontWeight:700,color:T.warning,marginBottom:4}}>🔶 Low Stock (≤5)</div>{low.map(p=><div key={p.id} style={{fontSize:12,color:T.warning}}>• {p.name_mm||p.name}: {p.stock}</div>)}</div>}
    {products.map(p=>{const oos__=isOOS(p);return(<div key={p.id} style={{...iosCard(T,{padding:"12px 14px",marginBottom:10})}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><PImg p={p} size={42} r={10} T={T}/><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:700,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",letterSpacing:-0.2}}>{p.name_mm||p.name}</div><div style={{fontSize:11,fontWeight:700,marginTop:1,color:oos__?T.danger:(p.stock||0)<=5?T.warning:T.success}}>{oos__?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`}</div></div></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {adding[p.id]?<div style={{display:"flex",gap:6,flex:1}}><input type="number" min="1" value={vals[p.id]||""} onChange={e=>setVals(v=>({...v,[p.id]:e.target.value}))} placeholder="qty" style={{width:64,padding:"7px 10px",borderRadius:10,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",color:T.text,fontFamily:"inherit"}}/><button onClick={()=>doAdd(p)} style={{padding:"7px 14px",borderRadius:10,border:"none",background:T.success,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700}}>+ Add</button><button onClick={()=>setAdding(v=>({...v,[p.id]:false}))} style={{padding:"7px 10px",borderRadius:10,...ios(T,{}),border:"none",color:T.muted,cursor:"pointer",fontSize:12}}>✕</button></div>
          :<button onClick={()=>setAdding(v=>({...v,[p.id]:true}))} style={{padding:"7px 14px",borderRadius:10,...ios(T,{border:`1.5px solid ${T.success}`}),background:`${T.success}15`,color:T.success,cursor:"pointer",fontSize:12,fontWeight:700}}>➕ {t.addStock}</button>}
        <button onClick={()=>togPO(p)} style={{padding:"7px 12px",borderRadius:10,...ios(T,{border:`1.5px solid ${p.preorder?T.warning:T.border}`}),background:p.preorder?T.warningBg:"transparent",color:p.preorder?T.warning:T.muted,cursor:"pointer",fontSize:11,fontWeight:600}}>{p.preorder?"✅ Pre-order ON":`📋 ${t.switchPreorder}`}</button>
      </div>
    </div>);})}
  </div>);
}

function AdminReports({T,t}){
  const[orders,setOrders]=useState([]);const[period,setPeriod]=useState("week");const[loading,setLoading]=useState(true);
  useEffect(()=>{(async()=>{setLoading(true);try{setOrders(await sb("GET","orders",null,"order=created_at.desc")||[]);}catch(e){}setLoading(false);})();},[]);
  const filtered=useMemo(()=>{const now=new Date();return orders.filter(o=>{const d=new Date(o.created_at);if(period==="today")return d.toDateString()===now.toDateString();if(period==="week"){const w=new Date(now);w.setDate(w.getDate()-7);return d>=w;}if(period==="month"){const m=new Date(now);m.setDate(m.getDate()-30);return d>=m;}return true;});},[orders,period]);
  const active=filtered.filter(o=>o.status!=="cancelled");const rev=active.reduce((s,o)=>s+Number(o.total),0);const dep=active.reduce((s,o)=>s+(Number(o.deposit_paid)||0),0);
  const tops=useMemo(()=>{const m={};active.forEach(o=>(o.items||[]).forEach(item=>{const k=item.name||"?";if(!m[k])m[k]={name:item.name_mm||item.name||k,qty:0,rev:0};m[k].qty+=(item.qty||0);m[k].rev+=item.total||0;}));return Object.values(m).sort((a,b)=>b.rev-a.rev).slice(0,6);},[active]);
  const exportExcel=()=>{const wb=XLSX.utils.book_new();const oRows=[["Order#","Date","Status","Customer","Phone","Address","Items","Total","Deposit","Balance","Contact","Note","Admin Note"]];filtered.forEach(o=>oRows.push([o.order_number,fdate(o.created_at),o.status,o.customer_name,o.customer_phone,o.customer_address,(o.items||[]).map(i=>`${i.name_mm||i.name}x${i.qty}${i.is_preorder?"[PRE]":""}`).join(", "),o.total,o.deposit_paid||0,Math.max(0,o.total-(o.deposit_paid||0)),o.contact_method,o.customer_note||"",o.admin_note||""]));XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(oRows),"Orders");const sRows=[["Metric","Value"],["Period",period],["Revenue",rev],["Deposit",dep],["Balance Due",rev-dep],["Orders",filtered.length]];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sRows),"Summary");if(tops.length>0)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Product","Qty","Revenue"],...tops.map(p=>[p.name,p.qty,p.rev])]),"Top Products");XLSX.writeFile(wb,`ShweTwin_${period}_${new Date().toISOString().slice(0,10)}.xlsx`);};
  const periods=[{k:"today",l:t.today},{k:"week",l:t.week},{k:"month",l:t.month},{k:"year",l:t.allTime}];
  const Stat=({emoji,label,value,color})=><div style={{...iosCard(T,{padding:14,flex:1,minWidth:80})}}><div style={{fontSize:18,marginBottom:2}}>{emoji}</div><div style={{fontSize:10,color:T.muted,fontWeight:600,textTransform:"uppercase",letterSpacing:0.3}}>{label}</div><div style={{fontSize:14,fontWeight:800,color:color||T.text,marginTop:2,letterSpacing:-0.3}}>{value}</div></div>;
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    <div style={{display:"flex",gap:6,marginBottom:14,...iosCard(T,{padding:5})}}>
      {periods.map(({k,l})=><button key={k} onClick={()=>setPeriod(k)} style={{flex:1,padding:"8px 4px",borderRadius:10,border:"none",background:period===k?T.primary:"transparent",color:period===k?"#fff":T.muted,cursor:"pointer",fontSize:12,fontWeight:period===k?700:400,transition:"all 0.15s"}}>{l}</button>)}
    </div>
    {loading?<div style={{textAlign:"center",padding:40,color:T.muted}}>Loading...</div>:<>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}><Stat emoji="💰" label={t.revenue} value={fmt(rev)} color={T.primary}/><Stat emoji="📦" label={t.totalOrders} value={filtered.length}/><Stat emoji="✅" label="Deposit" value={fmt(dep)} color={T.success}/><Stat emoji="⏳" label="Balance" value={fmt(rev-dep)} color={T.danger}/></div>
      {tops.length>0&&<div style={{...iosCard(T,{padding:16,marginBottom:12})}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:12}}>🏆 {t.topProducts}</div>
        {tops.map((p,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"7px 0",borderBottom:i<tops.length-1?`0.5px solid ${T.border}`:"none"}}><span style={{width:22,height:22,borderRadius:11,background:[T.primary,"rgba(192,192,192,0.8)","rgba(205,127,50,0.8)","rgba(160,160,160,0.5)","rgba(160,160,160,0.5)","rgba(160,160,160,0.5)"][i],display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#fff",flexShrink:0}}>{i+1}</span><span style={{flex:1,fontSize:12,fontWeight:600,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</span><span style={{fontSize:11,color:T.muted,marginRight:6}}>{p.qty}ခု</span><span style={{fontSize:12,fontWeight:700,color:T.primary,whiteSpace:"nowrap"}}>{fmt(p.rev)}</span></div>)}
      </div>}
      <div style={{...iosCard(T,{padding:16})}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3}}>{t.recentOrders}</div><button onClick={exportExcel} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 14px",borderRadius:10,border:"none",background:"#21A366",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:700,boxShadow:"0 3px 10px rgba(33,163,102,0.4)"}}>📊 {t.exportExcel}</button></div>
        {filtered.slice(0,10).map(o=><div key={o.id} style={{padding:"9px 0",borderBottom:`0.5px solid ${T.border}`,fontSize:12}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}><span style={{fontWeight:700,color:T.text,letterSpacing:-0.2}}>{o.order_number} · {o.customer_name}</span><SBadge2 status={o.status} t={t}/></div><div style={{display:"flex",justifyContent:"space-between",color:T.muted}}><span>{fdate(o.created_at)}</span><span style={{fontWeight:700,color:T.primary,letterSpacing:-0.3}}>{fmt(o.total)}</span></div></div>)}
        {filtered.length===0&&<div style={{textAlign:"center",color:T.muted,padding:20,fontSize:13}}>{t.noData}</div>}
      </div>
    </>}
  </div>);
}

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
    try{const data={...form,price:Number(form.price),discount_value:Number(form.discount_value)||0,stock:Number(form.stock)||0,visible:product?.visible??true,updated_at:new Date().toISOString()};if(!product)data.created_at=new Date().toISOString();const rows=product?.id?await sb("PATCH","products",data,`id=eq.${product.id}`):await sb("POST","products",data);onSave(Array.isArray(rows)?rows[0]:(rows||{...data,id:uid()}));}catch(e){alert("Error: "+e.message);}
    setSaving(false);
  };
  const inp={width:"100%",padding:"10px 13px",borderRadius:12,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:T.text};
  const ta={...inp,minHeight:60,resize:"vertical"};
  const lbl={fontSize:11,fontWeight:600,color:T.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <button onClick={onCancel} style={{background:"none",border:"none",cursor:"pointer",color:T.primary,fontSize:22,padding:0}}>‹</button>
      <span style={{fontWeight:700,fontSize:15,color:T.text,letterSpacing:-0.3}}>{product?t.editProduct:t.addProduct}</span>
    </div>
    <div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
      <div style={{...iosCard(T,{padding:16})}}><label style={lbl}>{t.productName}</label><input style={inp} value={form.name} onChange={e=>set("name",e.target.value)} placeholder="e.g. Snail White Body Lotion"/><label style={lbl}>{t.productNameMM}</label><input style={inp} value={form.name_mm||""} onChange={e=>set("name_mm",e.target.value)} placeholder="optional"/><label style={lbl}>{t.category}</label>{cats.length===0?<div style={{...ios(T,{borderRadius:10,padding:"8px 12px",background:T.dangerBg}),fontSize:12,color:T.danger,marginBottom:12}}>Categories မရှိသေးပါ</div>:<select style={{...inp,background:T.surfaceStrong}} value={form.category} onChange={e=>set("category",e.target.value)}>{cats.map(c=><option key={c.id} value={c.name}>{c.name}</option>)}</select>}<div style={{display:"flex",gap:10,alignItems:"end"}}><div style={{flex:1}}><label style={lbl}>{t.emoji}</label><input style={{...inp,marginBottom:0}} value={form.emoji||""} onChange={e=>set("emoji",e.target.value)}/></div><label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",paddingBottom:1,color:T.text,fontSize:13,whiteSpace:"nowrap"}}><input type="checkbox" checked={form.featured||false} onChange={e=>set("featured",e.target.checked)} style={{width:15,height:15,accentColor:T.primary}}/>⭐ Featured</label></div></div>
      <div style={{...iosCard(T,{padding:16})}}><label style={lbl}>📝 {t.description}</label><textarea style={ta} value={form.description||""} onChange={e=>set("description",e.target.value)}/><label style={lbl}>👤 {t.suitableFor}</label><textarea style={{...ta,minHeight:48}} value={form.suitable_for||""} onChange={e=>set("suitable_for",e.target.value)}/><label style={lbl}>✅ {t.benefits}</label><textarea style={{...ta,minHeight:48}} value={form.benefits||""} onChange={e=>set("benefits",e.target.value)}/><label style={lbl}>📋 {t.usage}</label><textarea style={{...ta,minHeight:48}} value={form.usage_info||""} onChange={e=>set("usage_info",e.target.value)}/><label style={lbl}>⚠️ {t.warning}</label><textarea style={{...ta,minHeight:44,marginBottom:0}} value={form.warning||""} onChange={e=>set("warning",e.target.value)}/></div>
      <div style={{...iosCard(T,{padding:16})}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>{t.price}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.price} onChange={e=>set("price",e.target.value)}/></div><div><label style={lbl}>{t.discountType}</label><select style={{...inp,marginBottom:0,background:T.surfaceStrong}} value={form.discount_type} onChange={e=>set("discount_type",e.target.value)}><option value="percent">{t.discountPct}</option><option value="fixed">{t.discountFixed}</option></select></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={lbl}>{t.discountVal}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.discount_value} onChange={e=>set("discount_value",e.target.value)}/></div><div><label style={lbl}>{t.stockQty}</label><input style={{...inp,marginBottom:0}} type="number" min="0" value={form.stock} onChange={e=>set("stock",e.target.value)}/></div></div><label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",userSelect:"none",marginBottom:12,fontSize:13,color:T.text}}><input type="checkbox" checked={form.preorder||false} onChange={e=>set("preorder",e.target.checked)} style={{width:15,height:15,accentColor:T.primary}}/>{t.enablePreorder}</label><div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>{t.bulkTiers}</div>{(Array.isArray(form.bulk_discounts)?form.bulk_discounts:[]).map((b,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}><input type="number" min="2" value={b.min_qty} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,min_qty:Number(e.target.value)}:x))} style={{width:50,padding:"6px 8px",borderRadius:8,...ios(T,{background:T.surfaceStrong}),fontSize:12,outline:"none",textAlign:"center",color:T.text}}/><span style={{fontSize:11,color:T.muted}}>ခု+</span><input type="number" min="1" max="99" value={b.discount_percent} onChange={e=>set("bulk_discounts",form.bulk_discounts.map((x,j)=>j===i?{...x,discount_percent:Number(e.target.value)}:x))} style={{width:50,padding:"6px 8px",borderRadius:8,...ios(T,{background:T.surfaceStrong}),fontSize:12,outline:"none",textAlign:"center",color:T.text}}/><span style={{fontSize:11,color:T.muted}}>% off</span><button onClick={()=>set("bulk_discounts",form.bulk_discounts.filter((_,j)=>j!==i))} style={{background:"none",border:"none",cursor:"pointer",color:T.danger,fontSize:15,marginLeft:"auto"}}>✕</button></div>)}<button onClick={()=>set("bulk_discounts",[...(form.bulk_discounts||[]),{min_qty:2,discount_percent:10}])} style={{padding:"6px 12px",borderRadius:8,...ios(T,{border:`1.5px dashed ${T.border}`}),cursor:"pointer",fontSize:12,color:T.primary,fontWeight:600,background:"transparent"}}>{t.addTier}</button></div>
      <div style={{...iosCard(T,{padding:16})}}><div style={{fontSize:11,color:T.muted,marginBottom:8,fontStyle:"italic"}}>💡 {t.gdriveTip}</div><div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>🖼️ {t.images}</div><div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>{(Array.isArray(form.images)?form.images:[]).map((img,i)=><div key={i} style={{position:"relative"}}><img src={img} style={{width:52,height:52,objectFit:"cover",borderRadius:10,border:`1px solid ${T.border}`,display:"block"}} onError={e=>e.target.style.opacity="0.3"}/><button onClick={()=>set("images",(form.images||[]).filter((_,j)=>j!==i))} style={{position:"absolute",top:-5,right:-5,width:17,height:17,borderRadius:9,border:"none",background:T.danger,color:"#fff",cursor:"pointer",fontSize:9,lineHeight:"17px",textAlign:"center",padding:0}}>✕</button></div>)}<button onClick={()=>iRef.current?.click()} style={{width:52,height:52,borderRadius:10,...ios(T,{border:`2px dashed ${T.border}`}),cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted}}>+</button></div><input ref={iRef} type="file" accept="image/*" multiple onChange={upImg} style={{display:"none"}}/><div style={{display:"flex",gap:8,marginBottom:12}}><input style={{flex:1,padding:"9px 11px",borderRadius:10,...ios(T,{background:T.surfaceStrong}),fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,color:T.text}} value={urlIn} onChange={e=>setUrlIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addUrl()} placeholder="imgbb.com URL"/><button onClick={addUrl} style={{padding:"9px 12px",borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600,flexShrink:0}}>Add</button></div><div style={{fontSize:11,fontWeight:600,color:T.muted,marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>🎬 Video</div><div style={{display:"flex",gap:8}}><input style={{flex:1,padding:"9px 11px",borderRadius:10,...ios(T,{background:T.surfaceStrong}),fontSize:12,outline:"none",fontFamily:"inherit",boxSizing:"border-box",minWidth:0,color:T.text}} value={form.video_url||""} onChange={e=>set("video_url",e.target.value)} placeholder="Video URL"/><button onClick={()=>vRef.current?.click()} style={{padding:"9px 11px",borderRadius:10,...ios(T,{border:`1px solid ${T.border}`}),color:T.text,cursor:"pointer",fontSize:11,flexShrink:0}}>Upload</button></div><input ref={vRef} type="file" accept="video/*" onChange={upVid} style={{display:"none"}}/></div>
      <div style={{display:"flex",gap:10}}><button onClick={handleSave} disabled={saving} style={{flex:1,padding:14,borderRadius:14,border:"none",background:saving?"rgba(0,0,0,0.15)":T.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700,boxShadow:saving?undefined:`0 4px 14px ${T.primary}50`}}>{saving?"⏳ Saving...":(product?t.save:t.addProduct)}</button><button onClick={onCancel} style={{flex:1,padding:14,borderRadius:14,...ios(T,{border:`1px solid ${T.border}`}),background:"transparent",color:T.muted,cursor:"pointer",fontSize:14}}>{t.cancel}</button></div>
    </div>
  </div>);
}

function AdminSettings({T,t,settings,onSave,onTheme,onLang,themeName,lang,onDisconnect}){
  const[form,setForm]=useState({...settings});const[logo,setLogo]=useState(settings.logo||"");const[banner,setBanner]=useState(settings.banner||"");const[saving,setSaving]=useState(false);
  const logoRef=useRef();const bannerRef=useRef();const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const upLogo=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setLogo(b);set("logo",b);e.target.value="";};
  const upBanner=async e=>{const f=e.target.files?.[0];if(!f)return;const b=await fileToB64(f);setBanner(b);set("banner",b);e.target.value="";};
  const doSave=async()=>{setSaving(true);try{const data={...form,logo,banner};await sbUpsert("shop_settings",Object.entries(data).map(([key,value])=>({key,value:String(value||"")})));onSave(data);alert("✅ Saved!");}catch(e){alert("Error: "+e.message);}setSaving(false);};
  const inp={width:"100%",padding:"11px 13px",borderRadius:12,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",boxSizing:"border-box",marginBottom:12,fontFamily:"inherit",color:T.text};
  const lbl={fontSize:11,fontWeight:600,color:T.muted,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:0.5};
  return(<div style={{padding:"12px 14px",paddingBottom:80}}>
    <div style={{...iosCard(T,{padding:16,marginBottom:12})}}><div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:12}}>🎨 {t.themeLabel}</div><div style={{display:"flex",gap:8,marginBottom:10}}>{Object.entries(THEMES).map(([k])=><button key={k} onClick={()=>onTheme(k)} style={{flex:1,padding:10,borderRadius:12,...ios(T,{border:`2px solid ${themeName===k?T.primary:T.border}`}),background:themeName===k?T.primary:"transparent",color:themeName===k?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:700}}>{k==="light"?"☀️ Light":"🌙 Dark"}</button>)}</div><div style={{display:"flex",gap:8}}>{["mm","en"].map(l=><button key={l} onClick={()=>onLang(l)} style={{flex:1,padding:10,borderRadius:12,...ios(T,{border:`2px solid ${lang===l?T.primary:T.border}`}),background:lang===l?T.primary:"transparent",color:lang===l?"#fff":T.text,cursor:"pointer",fontSize:12,fontWeight:700}}>{l==="mm"?"🇲🇲 မြန်မာ":"🇬🇧 English"}</button>)}</div></div>
    <div style={{...iosCard(T,{padding:16,marginBottom:12})}}><div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:12}}>🖼️ Logo & Banner</div><div style={{display:"flex",gap:12,marginBottom:10}}><div style={{textAlign:"center"}}><div style={{width:68,height:68,borderRadius:34,...ios(T,{border:`2px dashed ${T.border}`}),display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",cursor:"pointer"}} onClick={()=>logoRef.current?.click()}>{logo?<img src={logo} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<img src={DEFAULT_LOGO} style={{width:"100%",height:"100%",objectFit:"cover"}}/>}</div><div style={{fontSize:10,color:T.muted,marginTop:3}}>Logo</div></div><div style={{flex:1}}><div style={{height:68,borderRadius:14,...ios(T,{border:`2px dashed ${T.border}`}),display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",cursor:"pointer"}} onClick={()=>bannerRef.current?.click()}>{banner?<img src={banner} style={{width:"100%",height:"100%",objectFit:"cover"}}/>:<span style={{fontSize:12,color:T.muted}}>+ Banner Image</span>}</div></div></div><input ref={logoRef} type="file" accept="image/*" onChange={upLogo} style={{display:"none"}}/><input ref={bannerRef} type="file" accept="image/*" onChange={upBanner} style={{display:"none"}}/></div>
    <div style={{...iosCard(T,{padding:16,marginBottom:12})}}><div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:12}}>🏪 Shop Info</div><label style={lbl}>{t.shopName}</label><input style={inp} value={form.shop_name||""} onChange={e=>set("shop_name",e.target.value)}/><label style={lbl}>{t.shopNameMM}</label><input style={inp} value={form.shop_name_mm||""} onChange={e=>set("shop_name_mm",e.target.value)}/><label style={lbl}>{t.fbLink}</label><input style={inp} value={form.fb_link||""} onChange={e=>set("fb_link",e.target.value)} placeholder="https://m.me/yourpage"/><label style={lbl}>{t.viberNum}</label><input style={inp} value={form.viber_num||""} onChange={e=>set("viber_num",e.target.value)} placeholder="+95912345678"/><label style={lbl}>{t.waNum}</label><input style={inp} value={form.wa_num||""} onChange={e=>set("wa_num",e.target.value)} placeholder="+95912345678"/><label style={lbl}>{t.phoneNum}</label><input style={inp} value={form.phone_num||""} onChange={e=>set("phone_num",e.target.value)} placeholder="+95912345678"/><label style={lbl}>{t.adminPw}</label><input style={{...inp,marginBottom:0}} type="password" value={form.admin_pw||""} onChange={e=>set("admin_pw",e.target.value)}/></div>
    <button onClick={doSave} disabled={saving} style={{width:"100%",padding:14,borderRadius:14,border:"none",background:saving?"rgba(0,0,0,0.15)":T.primary,color:"#fff",cursor:saving?"wait":"pointer",fontSize:15,fontWeight:700,marginBottom:10,boxShadow:saving?undefined:`0 4px 16px ${T.primary}50`}}>{saving?"⏳ Saving...":t.saveSettings}</button>
    <button onClick={onDisconnect} style={{width:"100%",padding:12,borderRadius:14,...ios(T,{border:`1px solid ${T.danger}50`,background:T.dangerBg}),color:T.danger,cursor:"pointer",fontSize:13,fontWeight:600}}>🔌 Disconnect Supabase</button>
  </div>);
}

function AdminPanel({T,t,products,cats,settings,onProdChange,onCatsChange,onSettingsChange,onTheme,onLang,themeName,lang,onBack,onDisconnect}){
  const[tab,setTab]=useState("products");const[showForm,setShowForm]=useState(false);const[editing,setEditing]=useState(null);const[newCat,setNewCat]=useState("");
  const tabs=[{k:"products",l:"📦"},{k:"stock",l:"📊"},{k:"orders",l:"🛍️"},{k:"reports",l:"📈"},{k:"settings",l:"⚙️"}];
  const delProd=async p=>{if(!window.confirm(`Delete "${p.name_mm||p.name}"?`))return;try{await sb("DELETE","products",null,`id=eq.${p.id}`);onProdChange(products.filter(x=>x.id!==p.id));}catch(e){alert(e.message);}};
  const togVis=async p=>{try{await sb("PATCH","products",{visible:!p.visible},`id=eq.${p.id}`);onProdChange(products.map(x=>x.id===p.id?{...x,visible:!x.visible}:x));NS.add(`${p.name_mm||p.name} ${!p.visible?"visible":"hidden"}`);}catch(e){alert(e.message);}};
  const addCat=async name=>{if(!name||cats.find(c=>c.name===name))return;try{const rows=await sb("POST","categories",{name,sort_order:cats.length+1});onCatsChange([...cats,...(Array.isArray(rows)?rows:[rows])]);}catch(e){alert(e.message);}};
  const delCat=async cat=>{if(!window.confirm(`Delete "${cat.name}"?`))return;try{await sb("DELETE","categories",null,`id=eq.${cat.id}`);onCatsChange(cats.filter(c=>c.id!==cat.id));}catch(e){alert(e.message);}};
  if(showForm)return<ProductForm T={T} t={t} product={editing} cats={cats} onSave={saved=>{if(editing)onProdChange(products.map(x=>x.id===saved.id?saved:x));else onProdChange([saved,...products]);NS.add(editing?`${t.notifUpdated}: ${saved.name_mm||saved.name}`:`New: ${saved.name_mm||saved.name}`);setShowForm(false);setEditing(null);}} onCancel={()=>{setShowForm(false);setEditing(null);}}/>;
  return(<div style={{minHeight:"100vh",paddingBottom:40}}>
    <div style={{...ios(T,{padding:"10px 16px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100,background:T.headerBg,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      <img src={DEFAULT_LOGO} style={{width:30,height:30,borderRadius:15,objectFit:"cover"}} alt="logo"/>
      <span style={{fontWeight:800,fontSize:15,color:T.text,letterSpacing:-0.3,flex:1}}>Admin Panel</span>
      <button onClick={onBack} style={{...ios(T,{borderRadius:20,padding:"6px 14px"}),color:T.primary,cursor:"pointer",fontSize:12,fontWeight:700,border:`1px solid ${T.primary}`,background:"transparent"}}>{t.logout} ‹</button>
    </div>
    <div style={{display:"flex",...ios(T,{borderBottom:`0.5px solid ${T.border}`,borderRadius:0,borderLeft:"none",borderRight:"none",borderTop:"none"})}}>
      {tabs.map(tab_=><button key={tab_.k} onClick={()=>setTab(tab_.k)} style={{flex:1,padding:"12px 4px",border:"none",background:"none",cursor:"pointer",fontSize:18,borderBottom:`2.5px solid ${tab===tab_.k?T.primary:"transparent"}`,transition:"all 0.15s",opacity:tab===tab_.k?1:0.5}}>{tab_.l}</button>)}
    </div>
    {tab==="products"&&<div style={{padding:"12px 14px",paddingBottom:80}}>
      <button onClick={()=>{setEditing(null);setShowForm(true);}} style={{width:"100%",padding:13,borderRadius:14,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700,marginBottom:12,boxShadow:`0 4px 14px ${T.primary}50`}}>+ {t.addProduct}</button>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {products.map(p=><div key={p.id} style={{...iosCard(T,{padding:"11px 12px",display:"flex",gap:10,alignItems:"center",opacity:p.visible?1:0.45})}}>
          <PImg p={p} size={50} r={10} T={T}/>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:700,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",letterSpacing:-0.2}}>{p.name_mm||p.name}</div><div style={{fontSize:11,color:T.muted,marginTop:1}}>{p.category} · {fmt(calcUnit(p))}</div><div style={{fontSize:10,marginTop:1,color:isOOS(p)?T.danger:(p.stock||0)<=5?T.warning:T.success,fontWeight:600}}>{isOOS(p)?(p.preorder?"📋 Pre-order":"⚠️ "+t.tempOOS):`Stock: ${p.stock}`} · {p.visible?"🟢":"⚫"}</div></div>
          <div style={{display:"flex",flexDirection:"column",gap:5,flexShrink:0}}>
            <button onClick={()=>{setEditing(p);setShowForm(true);}} style={{padding:"5px 9px",borderRadius:8,...ios(T,{border:`0.5px solid ${T.border}`}),cursor:"pointer",fontSize:12,color:T.text}}>✏️</button>
            <button onClick={()=>togVis(p)} style={{padding:"5px 9px",borderRadius:8,...ios(T,{border:`0.5px solid ${T.border}`}),cursor:"pointer",fontSize:12,color:T.text}}>{p.visible?"🙈":"👁️"}</button>
            <button onClick={()=>delProd(p)} style={{padding:"5px 9px",borderRadius:8,...ios(T,{border:`0.5px solid ${T.danger}50`,background:T.dangerBg}),cursor:"pointer",fontSize:12,color:T.danger}}>🗑️</button>
          </div>
        </div>)}
        {products.length===0&&<div style={{textAlign:"center",padding:"40px 20px",color:T.muted,fontSize:14}}>{t.noProducts}</div>}
      </div>
      <div style={{...iosCard(T,{padding:14,marginTop:14})}}>
        <div style={{fontSize:13,fontWeight:700,color:T.text,letterSpacing:-0.3,marginBottom:10}}>🏷️ Categories</div>
        <div style={{display:"flex",gap:8,marginBottom:10}}><input value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>e.key==="Enter"&&newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} placeholder="New category" style={{flex:1,padding:"9px 12px",borderRadius:10,...ios(T,{background:T.surfaceStrong}),fontSize:13,outline:"none",fontFamily:"inherit",color:T.text,minWidth:0,boxSizing:"border-box"}}/><button onClick={()=>newCat.trim()&&(addCat(newCat.trim()),setNewCat(""))} style={{padding:"9px 16px",borderRadius:10,border:"none",background:T.primary,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700,flexShrink:0}}>+</button></div>
        {cats.map(cat=><div key={cat.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 12px",...ios(T,{borderRadius:12,border:`0.5px solid ${T.border}`}),marginBottom:6}}><span style={{fontSize:13,fontWeight:600,color:T.text}}>{cat.name}</span><button onClick={()=>delCat(cat)} style={{background:"none",...ios(T,{border:`0.5px solid ${T.danger}50`,borderRadius:8,background:T.dangerBg}),cursor:"pointer",color:T.danger,fontSize:11,padding:"3px 9px"}}>🗑️</button></div>)}
      </div>
    </div>}
    {tab==="stock"&&<StockMgr T={T} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="orders"&&<OrderMgr T={T} t={t} products={products} setProducts={onProdChange}/>}
    {tab==="reports"&&<AdminReports T={T} t={t}/>}
    {tab==="settings"&&<AdminSettings T={T} t={t} settings={settings} onSave={s=>onSettingsChange(s)} onTheme={onTheme} onLang={onLang} themeName={themeName} lang={lang} onDisconnect={onDisconnect}/>}
  </div>);
}


// ── Main App ──────────────────────────────────────────────────
export default function App(){
  const[cfg,setCfg]=useState(null);
  const[ready,setReady]=useState(false);
  const[products,setProducts]=useState([]);
  const[cats,setCats]=useState([]);
  const[settings,setSettings]=useState({});
  const[cart,setCart]=useState([]);
  const[tab,setTab]=useState("home");
  const[catFilter,setCatFilter]=useState("all");
  const[searchQ,setSearchQ]=useState("");
  const[showDrop,setShowDrop]=useState(false);
  const[selProd,setSelProd]=useState(null);
  const[page,setPage]=useState("home"); // home|product|checkout
  const[adminIn,setAdminIn]=useState(false);
  const[themeName,setThemeName]=useState(LS.get("theme")||"light");
  const[lang,setLang]=useState(LS.get("lang")||"mm");
  const[user,setUser]=useState(LS.get("user")||null);
  const[showAuth,setShowAuth]=useState(false);

  const T=THEMES[themeName]||THEMES.light;
  const t=TR[lang]||TR.mm;

  // Handle OAuth callback (Google/Facebook redirect)
  useEffect(()=>{
    const hash=window.location.hash;
    if(hash.includes("access_token")){
      const params=new URLSearchParams(hash.replace("#","?"));
      const access_token=params.get("access_token");
      const u={name:"User",provider:"oauth",access_token};
      setUser(u);LS.set("user",u);
      window.history.replaceState(null,"",window.location.pathname);
    }
  },[]);

  // Init config
  useEffect(()=>{
    (async()=>{
      let initCfg=null;
      if(SB_URL&&SB_KEY) initCfg={url:SB_URL,key:SB_KEY};
      else initCfg=LS.get("sb_cfg");
      if(initCfg) setCfg(initCfg);
      else setReady(true);
    })();
  },[]);

  // Fetch data
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
        setProducts(prods||[]);
        setCats(catsData||[]);
        const smap={};(setsData||[]).forEach(s=>{smap[s.key]=s.value;});
        setSettings(smap);
        if(smap.theme)setThemeName(smap.theme);
        if(smap.language)setLang(smap.language);
      }catch(e){
        // If Supabase connection fails, use sample data
        console.warn("Supabase not connected, using sample data");
        setProducts(SAMPLE_PRODUCTS);
        setCats(SAMPLE_CATS);
        setSettings(SAMPLE_SETTINGS);
      }
      setReady(true);
    })();
  },[cfg]);

  const loadAllProds=useCallback(async()=>{
    if(!cfg)return;
    try{setProducts(await sb("GET","products",null,"order=created_at.desc")||[]);}catch(e){}
  },[cfg]);
  useEffect(()=>{if(adminIn)loadAllProds();},[adminIn]);

  const searchRes=useMemo(()=>{
    if(!searchQ.trim())return[];
    const q=searchQ.toLowerCase();
    return products.filter(p=>p.visible!==false&&(
      (p.name||"").toLowerCase().includes(q)||(p.name_mm||"").includes(searchQ)||
      (p.description||"").includes(searchQ)||(p.category||"").includes(searchQ)
    )).slice(0,7);
  },[searchQ,products]);

  const visProds=useMemo(()=>products.filter(p=>p.visible!==false&&(catFilter==="all"||p.category===catFilter)),[products,catFilter]);
  const cartCount=cart.reduce((s,i)=>s+i.qty,0);
  const cartTotal=cart.reduce((s,i)=>s+calcUnit(i.p,i.qty)*i.qty,0);
  const addToCart=(p,qty=1)=>setCart(prev=>{const ex=prev.find(i=>i.p.id===p.id);return ex?prev.map(i=>i.p.id===p.id?{...i,qty:i.qty+qty}:i):[...prev,{p,qty}];});
  const removeItem=id=>setCart(c=>c.filter(i=>i.p.id!==id));
  const updateQty=(id,qty)=>{if(qty<=0)removeItem(id);else setCart(c=>c.map(i=>i.p.id===id?{...i,qty}:i));};
  const openProd=p=>{setSelProd(p);setPage("product");setShowDrop(false);setSearchQ("");};
  const changeTheme=k=>{setThemeName(k);LS.set("theme",k);};
  const changeLang=l=>{setLang(l);LS.set("lang",l);};
  const onSettingsChange=s=>{setSettings(s);if(s.theme)changeTheme(s.theme);if(s.language)changeLang(s.language);};

  const bg={background:T.bg,backgroundAttachment:"fixed",minHeight:"100vh",fontFamily:"-apple-system,'SF Pro Text','Segoe UI','Myanmar Text',sans-serif"};

  // Show setup if no config
  if(!ready&&!cfg) return <SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;
  if(!cfg) return <SetupScreen onConnect={c=>{LS.set("sb_cfg",c);setCfg(c);}}/>;

  // Loading
  if(!ready) return(
    <div style={{...bg,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <img src={DEFAULT_LOGO} style={{width:72,height:72,borderRadius:36,boxShadow:T.shadowMd,animation:"pulse 1.5s ease-in-out infinite"}} alt="logo"/>
      <div style={{color:T.text,fontSize:16,fontWeight:700,letterSpacing:-0.3}}>Shwe Twin BKK</div>
      <div style={{color:T.muted,fontSize:13}}>Loading...</div>
    </div>
  );

  // Auth screen
  if(showAuth) return(
    <div style={bg}>
      <AuthScreen T={T} t={t} cfg={cfg}
        onAuth={u=>{setUser(u);setShowAuth(false);}}
        onGuest={()=>setShowAuth(false)}/>
    </div>
  );

  // Admin panel
  if(tab==="admin"){
    if(!adminIn) return(
      <div style={{...bg,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
        <div style={{...iosCard(T,{padding:"40px 28px",maxWidth:360,width:"100%"})}}>
          <div style={{textAlign:"center",marginBottom:24}}>
            <img src={DEFAULT_LOGO} style={{width:72,height:72,borderRadius:36,marginBottom:10,boxShadow:T.shadowMd}} alt="logo"/>
            <div style={{fontSize:20,fontWeight:800,color:T.text,letterSpacing:-0.5}}>{t.adminPanel}</div>
          </div>
          <AdminLoginBox T={T} t={t} adminPw={settings.admin_pw} onSuccess={()=>setAdminIn(true)} onBack={()=>setTab("home")}/>
        </div>
      </div>
    );
    return(
      <div style={bg}>
        <AdminPanel T={T} t={t} products={products} cats={cats} settings={settings}
          onProdChange={setProducts} onCatsChange={setCats} onSettingsChange={onSettingsChange}
          onTheme={changeTheme} onLang={changeLang} themeName={themeName} lang={lang}
          onBack={()=>{setTab("home");setAdminIn(false);setProducts(ps=>ps.filter(p=>p.visible!==false));}}
          onDisconnect={()=>{LS.del("sb_cfg");setCfg(null);setAdminIn(false);setTab("home");}}/>
      </div>
    );
  }

  // Product detail
  if(page==="product"&&selProd) return(
    <div style={bg}>
      <ProductDetail T={T} t={t} p={selProd} onBack={()=>setPage("home")} addToCart={addToCart} onBuyNow={()=>setPage("checkout")}/>
    </div>
  );

  // Checkout
  if(page==="checkout") return(
    <div style={bg}>
      <CheckoutPage T={T} t={t} cart={cart} total={cartTotal} settings={settings} user={user}
        onPlaced={()=>{setCart([]);setPage("home");}} onBack={()=>setPage("home")}/>
    </div>
  );

  // Main shop
  const shopMM=settings.shop_name_mm||settings.shop_name||"Shwe Twin BKK";
  return(
    <div style={bg}>
      <TopBar T={T} t={t} shopMM={shopMM} logo={settings.logo||""}
        cartCount={cartCount} searchQ={searchQ} setSearchQ={setSearchQ}
        showDrop={showDrop} setShowDrop={setShowDrop} dropRes={searchRes}
        onHit={openProd} onCart={()=>setTab("cart")} onAdmin={()=>setTab("admin")}
        onLogo={()=>{setTab("home");setPage("home");}}
        onThemeToggle={()=>changeTheme(themeName==="light"?"dark":"light")}
        themeName={themeName}/>
      <div style={{paddingBottom:80}}>
        {tab==="home"&&<HomePage T={T} t={t} products={visProds} cats={cats} catFilter={catFilter} setCatFilter={setCatFilter} onOpen={openProd} onAdd={addToCart} banner={settings.banner||""} shopMM={shopMM}/>}
        {tab==="cart"&&<CartPage T={T} t={t} cart={cart} updateQty={updateQty} removeItem={removeItem} total={cartTotal} onCheckout={()=>setPage("checkout")} onBack={()=>setTab("home")}/>}
        {tab==="track"&&<TrackOrder T={T} t={t}/>}
        {tab==="cats"&&<div style={{padding:"14px 14px",paddingBottom:80}}>
          <div style={{fontSize:18,fontWeight:800,color:T.text,letterSpacing:-0.5,marginBottom:14}}>✨ Categories</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
            {cats.map(c=>{const icons={"Thai Skincare":"🇹🇭","Korean Beauty":"🇰🇷","Makeup":"💄","Hair Care":"💆","Fashion":"👗","Accessories":"💎","Others":"🛍️"};return(
              <div key={c.id} onClick={()=>{setCatFilter(c.name);setTab("home");}} style={{...iosCard(T,{padding:"16px 8px",textAlign:"center",cursor:"pointer",transition:"all 0.2s"})}}
                onMouseEnter={e=>e.currentTarget.style.transform="scale(1.02)"} onMouseLeave={e=>e.currentTarget.style.transform="none"}>
                <div style={{fontSize:28,marginBottom:6}}>{icons[c.name]||"💄"}</div>
                <div style={{fontSize:11,fontWeight:600,color:T.text,lineHeight:1.3,letterSpacing:-0.2}}>{c.name}</div>
              </div>
            );})}
          </div>
        </div>}
        {tab==="profile"&&<ProfilePage T={T} t={t} user={user} onSignIn={()=>setShowAuth(true)} onSignOut={()=>{setUser(null);LS.del("user");}}/>}
      </div>
      <BottomNav T={T} t={t} tab={tab} setTab={t2=>{setTab(t2);setPage("home");}} cartCount={cartCount} user={user}/>
    </div>
  );
}
