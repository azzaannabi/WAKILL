"use strict";
/* Test de bout en bout : faux Claude + faux Meta/WhatsApp, vrai serveur. Lancer : node test/e2e.js */
const http=require("node:http"),crypto=require("node:crypto"),{spawn}=require("node:child_process"),fs=require("node:fs"),os=require("node:os"),path=require("node:path");
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
let fails=0;const ok=(c,m)=>{console.log((c?"  ✓ ":"  ✗ ")+m);if(!c)fails++;};

const graphCalls=[],claudeCalls=[];
function mock(handler){return new Promise(r=>{const s=http.createServer(async(q,a)=>{const b=[];for await(const c of q)b.push(c);handler(q,a,Buffer.concat(b));});s.listen(0,()=>r(s));});}
(async()=>{
  const claudeSrv=await mock((q,a,raw)=>{
    const body=JSON.parse(raw.toString());claudeCalls.push(body);
    const last=body.messages[body.messages.length-1];
    const send=o=>{a.writeHead(200,{"content-type":"application/json"});a.end(JSON.stringify(o));};
    if(Array.isArray(last.content)&&last.content[0]?.type==="tool_result")return send({stop_reason:"end_turn",content:[{type:"text",text:"تمام ✅ طلبيتك تسجلت"}]});
    const t=String(last.content);
    if(t.includes("ORDER_FULL"))return send({stop_reason:"tool_use",content:[{type:"tool_use",id:"tu1",name:"creer_commande",input:{nom:"Test Client",telephone:"22 333 444",adresse:"Rue 1",ville:"Hammam-Lif",gouvernorat:"",quartier:"Centre",produit_id:"p1",quantite:2}}]});
    if(t.includes("RECLAM"))return send({stop_reason:"tool_use",content:[{type:"tool_use",id:"tu2",name:"signaler_reclamation",input:{motif:"Colis non reçu"}}]});
    if(t.startsWith("OUI"))return send({stop_reason:"tool_use",content:[{type:"tool_use",id:"tu3",name:"repondre_confirmation",input:{decision:"oui"}}]});
    send({stop_reason:"end_turn",content:[{type:"text",text:"عسلامة 🌷"}]});
  });
  const graphSrv=await mock((q,a,raw)=>{
    graphCalls.push({url:q.url,ct:q.headers["content-type"]||"",auth:q.headers.authorization||"",body:/json/.test(q.headers["content-type"]||"")?JSON.parse(raw.toString()||"{}"):{len:raw.length}});
    a.writeHead(200,{"content-type":"application/json"});a.end(JSON.stringify(q.url.includes("/media")?{id:"media1"}:{ok:true}));
  });
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),"wakil-"));const PORT=4599,SECRET="sec123";
  const srv=spawn("node",[path.join(__dirname,"..","server.js")],{env:{...process.env,NODE_ENV:"test",PORT,DATA_DIR:dir,DASH_PASSWORD:"pw",
    ANTHROPIC_BASE:`http://127.0.0.1:${claudeSrv.address().port}`,GRAPH_BASE:`http://127.0.0.1:${graphSrv.address().port}`,
    DEBOUNCE_MS:"0",EXPORT_HOUR:"20"},stdio:["ignore","pipe","pipe"]});
  srv.stdout.on("data",d=>process.env.V&&process.stdout.write(d));srv.stderr.on("data",d=>process.stderr.write(d));
  await sleep(700);
  const B="http://127.0.0.1:"+PORT,AUTH="Basic "+Buffer.from("u:pw").toString("base64");
  const api=async(m,p,b)=>{const r=await fetch(B+p,{method:m,headers:{authorization:AUTH,"content-type":"application/json"},body:b?JSON.stringify(b):undefined});
    return r.headers.get("content-type")?.includes("json")?r.json():Buffer.from(await r.arrayBuffer());};

  console.log("Intégrations depuis l'appli (sans aucune variable d'environnement)");
  let st0=await api("GET","/api/state");
  ok(st0.cfg.claude===false&&st0.cfg.meta===false&&st0.cfg.whatsapp===false,"rien n'est connecté au démarrage, sans clés dans l'environnement");
  ok(st0.cfg.webhookUrl.startsWith("http://127.0.0.1"),"l'URL du webhook est détectée automatiquement depuis la requête, sans PUBLIC_URL");
  await api("POST","/api/sync",{secrets:{anthropicKey:"k",pageToken:"pt",appSecret:SECRET,verifyToken:"vt",waToken:"wt",waPhoneId:"111",ownerWa:"21699000111"}});
  st0=await api("GET","/api/state");
  ok(st0.cfg.claude&&st0.cfg.meta&&st0.cfg.whatsapp&&st0.cfg.ownerWa,"les trois connexions passent à « connecté » juste après avoir collé les jetons dans Intégrations");
  ok(st0.cfg.secrets.pageToken==="pt"&&st0.cfg.secrets.anthropicKey==="k","les valeurs collées sont renvoyées pour rester visibles dans le formulaire");

  const hook=async(obj,sig)=>{const raw=JSON.stringify(obj);return fetch(B+"/webhook/meta",{method:"POST",headers:{"content-type":"application/json","x-hub-signature-256":sig??("sha256="+crypto.createHmac("sha256",SECRET).update(raw).digest("hex"))},body:raw});};
  const ig=(id,text,mid)=>({object:"instagram",entry:[{id:"page",messaging:[{sender:{id},recipient:{id:"page"},message:{mid,text}}]}]});
  const dms=()=>graphCalls.filter(c=>c.url.includes("/me/messages"));
  const waMsgs=()=>graphCalls.filter(c=>c.url.includes("/111/messages"));

  console.log("Webhook");
  let r=await fetch(`${B}/webhook/meta?hub.mode=subscribe&hub.verify_token=vt&hub.challenge=CH42`);ok((await r.text())==="CH42","vérification Meta renvoie le challenge");
  r=await fetch(`${B}/webhook/meta?hub.mode=subscribe&hub.verify_token=bad&hub.challenge=x`);ok(r.status===403,"mauvais jeton refusé");
  r=await hook(ig("U1","salam","m0"),"sha256=deadbeef");ok(r.status===403,"signature invalide refusée");
  r=await fetch(B+"/api/state");ok(r.status===401,"tableau de bord protégé par mot de passe");

  console.log("Mode brouillon");
  await hook(ig("U1","salam bkeddech?","m1"));await sleep(400);
  let st=await api("GET","/api/state");let c1=st.convs.find(c=>c.channel==="instagram");
  ok(dms().length===0,"aucun DM envoyé en mode brouillon");ok(c1&&c1.draft,"un brouillon est en attente de validation");
  await hook(ig("U1","salam bkeddech?","m1"));await sleep(200);ok((await api("GET","/api/conv/"+c1.id)).turns.filter(t=>t.r==="client").length===1,"message en double ignoré (même identifiant)");
  const conv=await api("GET","/api/conv/"+c1.id),dr=conv.turns.find(t=>t.draft);
  await api("POST",`/api/conv/${c1.id}/draft/${dr.id}/send`,{text:"عسلامة 🌷 (corrigé)"});
  ok(dms().length===1&&dms()[0].body.message.text.includes("corrigé")&&dms()[0].body.recipient.id==="U1","brouillon corrigé puis envoyé au bon destinataire");

  console.log("Mode automatique et commande");
  await api("POST","/api/sync",{settings:{...st.settings,mode:"auto",autoStatus:true,produits:[{id:"p1",nom:"Seven Green",prix:49,desc:"",actif:true}],boutique:"Atgha"}});
  await hook(ig("U2","ORDER_FULL","m2"));await sleep(500);
  st=await api("GET","/api/state");const o=st.orders[0];
  ok(o&&o.status==="attente"&&o.tel==="22333444","commande créée, téléphone normalisé");
  ok(o.montant===98,"montant calculé par le code (2 x 49)");ok(o.gouvernorat==="Ben Arous","gouvernorat déduit de la ville (Hammam-Lif)");
  ok(dms().length===2&&dms()[1].body.message.text.includes("تسجلت"),"réponse de l'agent envoyée en DM");
  ok(claudeCalls.at(-1).messages.at(-1).content[0].type==="tool_result","résultat de l'outil renvoyé à Claude");
  ok(claudeCalls[0].system.includes("Seven Green")||claudeCalls.at(-1).system.includes("Seven Green"),"catalogue injecté dans le prompt");

  console.log("Réclamation");
  await hook(ig("U3","RECLAM colis","m3"));await sleep(500);
  st=await api("GET","/api/state");const tk=st.tickets[0],c3=st.convs.find(c=>c.name.includes("U3")||c.id===tk.convId);
  ok(tk&&tk.kind==="reclamation"&&tk.convId,"ticket de réclamation créé");ok(c3.paused,"agent mis en pause sur cette conversation");
  const alertMsg=waMsgs().find(w=>w.body.text?.body?.includes("Réclamation"));
  ok(alertMsg&&alertMsg.body.to==="21699000111"&&alertMsg.body.text.body.includes(`${st0.cfg.webhookUrl.replace("/webhook/meta","")}/#/c/${tk.convId}`),"alerte WhatsApp au propriétaire avec lien direct vers la conversation (URL détectée automatiquement)");
  const before=claudeCalls.length;await hook(ig("U3","encore moi","m4"));await sleep(300);ok(claudeCalls.length===before,"l'agent ne répond plus tant que tu as la main");

  console.log("Minuteurs");
  const t0=Date.now();await api("POST","/api/_dev/timers",{now:t0+21*36e5});
  st=await api("GET","/api/state");const o2=st.orders.find(x=>x.id===o.id);
  ok(o2.confirmAskedAt,"demande de confirmation envoyée après 20 h");
  ok(dms().at(-1).body.message.text.includes("نعم"),"la demande de confirmation part en DM (fenêtre de 24 h ouverte)");
  await hook(ig("U2","OUI","m5"));await sleep(500);st=await api("GET","/api/state");
  ok(claudeCalls.at(-2).system.includes("COMMANDES EN ATTENTE DE CONFIRMATION"),"le prompt contient la commande à confirmer");
  ok(st.orders.find(x=>x.id===o.id).status==="confirmee","réponse « oui » : commande confirmée");

  await hook(ig("U4","ORDER_FULL","m6"));await sleep(500);
  const o3=(await api("GET","/api/state")).orders[0];
  await api("POST","/api/_dev/timers",{now:Date.now()+21*36e5});
  await api("POST","/api/_dev/timers",{now:Date.now()+21*36e5+25*36e5});
  st=await api("GET","/api/state");const o3b=st.orders.find(x=>x.id===o3.id);
  ok(o3b.status==="sans_reponse","sans réponse après 24 h, sans relance");
  ok(waMsgs().some(w=>w.body.text?.body?.includes("Sans réponse")&&w.body.text.body.includes("/#/o/")),"alerte WhatsApp « sans réponse » avec lien");

  console.log("Statuts depuis le tableau de bord");
  const n0=dms().length;
  await api("POST","/api/sync",{orders:[{id:o.id,patch:{status:"emballee"},evAdd:[{t:Date.now(),x:"Statut : Emballée"}]}]});await sleep(300);
  ok(dms().length===n0+1&&dms().at(-1).body.message.text.includes("تجهزت"),"message « emballée » envoyé automatiquement");

  console.log("Export Excel");
  const fname=(await fetch(B+"/api/export",{method:"POST",headers:{authorization:AUTH,"content-type":"application/json"},body:"{}"}));
  const buf=Buffer.from(await fname.arrayBuffer());fs.writeFileSync("/tmp/wakil-export.xlsx",buf);
  ok(fname.headers.get("content-disposition")?.includes("colis_"),"fichier colis_… téléchargé");
  ok((await api("POST","/api/export",{})).error==="Rien à exporter.","une commande déjà exportée n'est pas re-exportée");
  await api("POST","/api/sync",{orders:[{id:o3.id,patch:{status:"confirmee",confirmedAt:Date.now(),gouvernorat:"Tunis"},evAdd:[]}]});
  await api("POST","/api/_dev/timers",{now:new Date(new Date().setHours(21,0,0,0)).getTime()});
  const media=graphCalls.find(c=>c.url.includes("/111/media"));const doc=waMsgs().find(w=>w.body.type==="document");
  ok(media&&doc&&doc.body.document.filename.endsWith(".xlsx")&&doc.body.to==="21699000111","export du soir envoyé en document WhatsApp");

  console.log("Apprentissage");
  await api("POST","/api/sync",{settings:{...(await api("GET","/api/state")).settings,lecons:[{id:"l1",situation:"قداش هذا؟",reponse:"عسلامة 🌷 بـ 49 دينار، وتخلص كي توصلك",t:1}]}});
  await hook(ig("U7","قداش هذا؟","m9"));await sleep(500);
  ok(claudeCalls.at(-1).system.includes("عسلامة 🌷 بـ 49 دينار، وتخلص كي توصلك")&&claudeCalls.at(-1).system.includes("CE QUE LE PROPRIÉTAIRE T'A APPRIS"),"la leçon apprise est injectée dans le prompt de l'agent");
  ok(claudeCalls.at(-1).system.includes("DÉRIJA TUNISIENNE"),"la consigne dérija tunisienne est dans le prompt");

  console.log("Simulateur");
  const sim=await api("POST","/api/test/send",{text:"ORDER_FULL",canal:"TikTok"});
  ok(sim.reply&&sim.events.some(e=>e.type==="order"),"simulateur : réponse et commande de test");
  ok((await api("GET","/api/state")).orders[0].origine==="test","la commande de test est marquée « test »");

  srv.kill();claudeSrv.close();graphSrv.close();
  console.log(fails?`\n${fails} échec(s)`:"\nTout est OK");process.exit(fails?1:0);
})();
