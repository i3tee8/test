const pptx = require('pptxgenjs');
const p = new pptx();
p.layout = 'LAYOUT_WIDE';            // 13.333 x 7.5
const W = 13.333, H = 7.5;

const NAVY='000050', DEEP='000033', BLUE='1061FF', CYAN='4BF5FF',
      WHITE='FFFFFF', INK='101540', BODY='4D5680', MUTE='8F9EDA',
      LINE='DDE8FF', SURF='F4F7FF', ORANGE='FF593C', GREEN='00E66E';
const HEAD='Arial', BODYF='Calibri';

const sh = () => ({ type:'outer', color:'101540', blur:14, offset:3, angle:90, opacity:0.10 });

function dark(){ const s=p.addSlide(); s.background={color:NAVY}; return s; }
function light(){ const s=p.addSlide(); s.background={color:WHITE}; return s; }

// motif: cyan-outlined rounded square holding a number
function chip(s,x,y,n,onDark){
  s.addShape(p.ShapeType.roundRect,{x,y,w:0.42,h:0.42,rectRadius:0.1,
    fill:{color:onDark?NAVY:SURF}, line:{color:CYAN,width:1.25}});
  s.addText(String(n),{x,y,w:0.42,h:0.42,align:'center',valign:'middle',isTextBox:true,margin:0,
    fontFace:HEAD,fontSize:13,bold:true,color:onDark?CYAN:BLUE});
}
function title(s,t,onDark,y=0.52){
  s.addText(t,{x:0.7,y,w:W-1.7,h:1.30,isTextBox:true,margin:0,valign:'top',
    fontFace:HEAD,fontSize:33,bold:true,color:onDark?WHITE:INK,lineSpacing:38});
}
function kicker(s,t,onDark,y=0.30){
  s.addText(t.toUpperCase(),{x:0.7,y,w:W-1.7,h:0.3,isTextBox:true,margin:0,
    fontFace:HEAD,fontSize:11,bold:true,charSpacing:2.2,color:onDark?CYAN:BLUE});
}
function card(s,x,y,w,h,fill){
  s.addShape(p.ShapeType.roundRect,{x,y,w,h,rectRadius:0.06,
    fill:{color:fill||SURF}, line:{color:LINE,width:1}, shadow:sh()});
}

/* ---------------- 1 TITLE ---------------- */
{
 const s=dark();
 s.addImage({path:'hero.png',x:0,y:0,w:W,h:H,sizing:{type:'cover',w:W,h:H}});
 s.addShape(p.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:NAVY,transparency:32}});
 s.addText('INTERNAL AUDIT EXCELLENCE',{x:1.0,y:2.35,w:10,h:0.32,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:12,bold:true,charSpacing:2.6,color:CYAN});
 s.addText('Revamping the\nInternal Audit platform',{x:1.0,y:2.85,w:10.6,h:1.9,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:44,bold:true,color:WHITE,lineSpacing:50});
 s.addText('From a departmental noticeboard to the place the audit function actually runs from.',
   {x:1.0,y:4.85,w:9.6,h:0.5,isTextBox:true,margin:0,fontFace:BODYF,fontSize:17,color:'B9C6EF'});
 s.addText('Proposal to management  ·  September 2026',{x:1.0,y:6.35,w:8,h:0.3,isTextBox:true,margin:0,
   fontFace:BODYF,fontSize:12,color:MUTE});
 s.addNotes('Framing: this is not a redesign for its own sake. The page exists; almost nobody uses it. We are proposing to change what it is for.');
}

/* ---------------- 2 CURRENT STATE ---------------- */
{
 const s=light();
 kicker(s,'Where we are today',false);
 title(s,'The page works. It just is not what anyone needs.',false);
 const items=[
  ['Opens with culture','Birthdays, awards and event photos occupy the first screen. The audit manual is below all of it.'],
  ['Tools are missing','TeamMate+ has no link on the home page. Auditors reach it from a bookmark, or they ask someone.'],
  ['Nothing is live','Every block is static text. There is no reason to return once you have read it.'],
  ['Closed to the business','A private group site. The departments we audit cannot open it at all.']];
 items.forEach((it,i)=>{
   const x = 0.7 + (i%2)*6.35, y = 2.05 + Math.floor(i/2)*2.2;
   card(s,x,y,5.95,1.95);
   chip(s,x+0.4,y+0.38,i+1,false);
   s.addText(it[0],{x:x+1.0,y:y+0.34,w:4.7,h:0.38,isTextBox:true,margin:0,
     fontFace:HEAD,fontSize:17,bold:true,color:INK});
   s.addText(it[1],{x:x+1.0,y:y+0.82,w:4.65,h:0.95,isTextBox:true,margin:0,
     fontFace:BODYF,fontSize:13.5,color:BODY,lineSpacing:19});
 });
 s.addNotes('Be fair to the current page: it was built for a purpose and it serves culture well. The problem is that it was never built to carry the work.');
}

/* ---------------- 3 COME FOR vs SHOWS ---------------- */
{
 const s=light();
 kicker(s,'The diagnosis',false);
 title(s,'What people come for, and what the page shows.',false);
 s.addShape(p.ShapeType.roundRect,{x:0.7,y:2.1,w:5.85,h:4.35,rectRadius:0.06,
   fill:{color:SURF},line:{color:LINE,width:1}});
 s.addShape(p.ShapeType.roundRect,{x:6.85,y:2.1,w:5.78,h:4.35,rectRadius:0.06,
   fill:{color:NAVY},line:{color:NAVY,width:1}});
 s.addText('WHAT THEY COME FOR',{x:1.1,y:2.42,w:5,h:0.3,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:BLUE});
 s.addText('WHAT THE PAGE LEADS WITH',{x:7.25,y:2.42,w:5,h:0.3,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:CYAN});
 const left=['TeamMate+','The audit manual','A workpaper template','Analytics support','Whether an issue is closed','What an audit involves'];
 const right=['Birthdays this month','Awareness month photos','The mission statement','Event calendar','Recycle bin link','Holiday calendar'];
 left.forEach((t,i)=>s.addText(t,{x:1.1,y:2.95+i*0.55,w:5.1,h:0.42,isTextBox:true,margin:0,
   fontFace:BODYF,fontSize:15,bold:true,color:INK}));
 right.forEach((t,i)=>s.addText(t,{x:7.25,y:2.95+i*0.55,w:5.1,h:0.42,isTextBox:true,margin:0,
   fontFace:BODYF,fontSize:15,color:'B9C6EF'}));
 s.addText('Neither list is wrong. They are simply in the wrong order.',
   {x:0.7,y:6.72,w:11.9,h:0.38,isTextBox:true,margin:0,fontFace:BODYF,fontSize:14,italic:true,color:BODY});
 s.addNotes('The left column is what an auditor opens a browser to find. The right column is what the home page currently offers first.');
}

/* ---------------- 4 THREE SHIFTS ---------------- */
{
 const s=dark();
 kicker(s,'The approach',true);
 title(s,'Three shifts, not a repaint.',true);
 const sh3=[['Work first, culture second','The tools and the methodology take the top of the page. Culture keeps its place, below them, done properly.'],
   ['Static becomes live','A Power BI dashboard and self-updating document lists. The page shows this week’s position, so it earns a weekly visit.'],
   ['Open it to the business','A path written for the departments we audit. It is what makes the site worth linking to from outside Internal Audit.']];
 sh3.forEach((it,i)=>{
   const x=0.7+i*4.06;
   s.addShape(p.ShapeType.roundRect,{x,y:2.25,w:3.8,h:3.7,rectRadius:0.06,
     fill:{color:DEEP},line:{color:'1C2470',width:1}});
   chip(s,x+0.38,2.62,i+1,true);
   s.addText(it[0],{x:x+0.38,y:3.3,w:3.1,h:0.9,isTextBox:true,margin:0,
     fontFace:HEAD,fontSize:19,bold:true,color:WHITE});
   s.addText(it[1],{x:x+0.38,y:4.3,w:3.08,h:1.5,isTextBox:true,margin:0,
     fontFace:BODYF,fontSize:13.5,color:'B9C6EF',lineSpacing:19});
 });
 s.addText('All three are built with SharePoint web parts we already have. No new licence, no developer.',
   {x:0.7,y:6.35,w:11.9,h:0.4,isTextBox:true,margin:0,fontFace:BODYF,fontSize:14,color:CYAN});
}

/* ---------------- 5 NEW HOME PAGE ---------------- */
{
 const s=light();
 kicker(s,'The proposal',false);
 title(s,'The new home page.',false);
 s.addImage({path:'deck-hero.png',x:0.7,y:2.08,w:11.93,h:4.05});
 s.addText('The first screen now answers: where do I go, and how is the plan tracking.',
   {x:0.7,y:6.35,w:11.9,h:0.4,isTextBox:true,margin:0,fontFace:BODYF,fontSize:14,color:BODY});
}

/* ---------------- 6 ONE CLICK ---------------- */
{
 const s=light();
 kicker(s,'Access',false);
 title(s,'Six destinations, one click, nothing competing.',false);
 s.addImage({path:'deck-goto.png',x:0.7,y:2.10,w:11.93,h:2.70});
 const pts=[['TeamMate+ leads','It is the most-used system in the department and had no link at all.'],
            ['One line each','Saying what you get, not what the thing is called.'],
            ['Order is the design','The bottom two get reviewed after a month of usage data.']];
 pts.forEach((t,i)=>{
   const x=0.7+i*4.06;
   s.addText(t[0],{x,y:5.05,w:3.8,h:0.34,isTextBox:true,margin:0,fontFace:HEAD,fontSize:15,bold:true,color:INK});
   s.addText(t[1],{x,y:5.45,w:3.8,h:0.9,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13,color:BODY,lineSpacing:18});
 });
}

/* ---------------- 7 THREE AUDIENCES ---------------- */
{
 const s=light();
 kicker(s,'Reach',false);
 title(s,'Three ways in — including one for the business.',false);
 s.addImage({path:'deck-start.png',x:0.7,y:2.10,w:11.93,h:3.42});
 s.addShape(p.ShapeType.roundRect,{x:0.7,y:5.68,w:11.93,h:1.1,rectRadius:0.06,
   fill:{color:SURF},line:{color:LINE,width:1}});
 s.addText([{text:'The third card is the one that changes our standing. ',options:{bold:true,color:INK}},
            {text:'It explains an audit to the departments we audit — what we ask for, how findings get agreed, what closes an issue. Nothing like it exists today.',options:{color:BODY}}],
   {x:1.05,y:5.9,w:11.2,h:0.7,isTextBox:true,margin:0,fontFace:BODYF,fontSize:14,lineSpacing:19});
}

/* ---------------- 8 DASHBOARD ---------------- */
{
 const s=dark();
 kicker(s,'The dashboard',true);
 title(s,'The page stops being a leaflet.',true);
 s.addImage({path:'deck-dash.png',x:0.7,y:2.10,w:7.55,h:2.62});
 const rows=[['Plan delivery','engagements reported against the approved plan'],
             ['Fieldwork in flight','and which are past their planned end date'],
             ['Issue ageing','open actions bucketed at 30, 90 and 180 days'],
             ['Days to close','trend against the 45-day target']];
 rows.forEach((r,i)=>{
   const y=2.12+i*0.72;
   s.addText(r[0],{x:8.6,y,w:4.1,h:0.3,isTextBox:true,margin:0,fontFace:HEAD,fontSize:14.5,bold:true,color:CYAN});
   s.addText(r[1],{x:8.6,y:y+0.30,w:4.05,h:0.4,isTextBox:true,margin:0,fontFace:BODYF,fontSize:12.5,color:'B9C6EF'});
 });
 s.addShape(p.ShapeType.roundRect,{x:0.7,y:5.0,w:11.93,h:1.5,rectRadius:0.06,
   fill:{color:DEEP},line:{color:'1C2470',width:1}});
 s.addText('Built on Power BI over two SharePoint lists — Engagements and Issues.',
   {x:1.05,y:5.22,w:11.2,h:0.34,isTextBox:true,margin:0,fontFace:HEAD,fontSize:15,bold:true,color:WHITE});
 s.addText('Those two lists are also the first data model of the in-house platform we want later, so this work is not thrown away — the migration becomes an export rather than a re-keying exercise.',
   {x:1.05,y:5.6,w:11.2,h:0.75,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13.5,color:'B9C6EF',lineSpacing:19});
 s.addNotes('Figures on the screenshot are illustrative. Real values come from the Engagements and Issues lists once populated.');
}

/* ---------------- 9 CULTURE ---------------- */
{
 const s=light();
 kicker(s,'Department culture',false);
 title(s,'Culture is kept — and given a better setting.',false);
 s.addImage({path:'deck-culture.png',x:0.7,y:2.10,w:7.4,h:3.25});
 const c=[['Nothing is removed','SPOC awards, Awareness Month, the Guest Auditor programme, birthdays and the team calendar all stay.'],
          ['Presented properly','A gallery with consistent crops and captions, instead of a carousel competing with the mission statement.'],
          ['Placed deliberately','Below the working content, so the page earns a daily visit and culture is what people find when they stay.']];
 c.forEach((t,i)=>{
   const y=2.10+i*1.16;
   s.addText(t[0],{x:8.45,y,w:4.3,h:0.3,isTextBox:true,margin:0,fontFace:HEAD,fontSize:14.5,bold:true,color:INK});
   s.addText(t[1],{x:8.45,y:y+0.32,w:4.25,h:0.82,isTextBox:true,margin:0,fontFace:BODYF,fontSize:12.5,color:BODY,lineSpacing:17});
 });
 s.addText('Images shown are placeholder brand art. Final build uses the department’s own photographs.',
   {x:0.7,y:5.45,w:7.4,h:0.34,isTextBox:true,margin:0,fontFace:BODYF,fontSize:11.5,italic:true,color:'6E79A3'});
}

/* ---------------- 10 WHY VISITS RISE ---------------- */
{
 const s=light();
 kicker(s,'Expected outcome',false);
 title(s,'Why this gets visited, when the old page did not.',false);
 const m=[['It holds the tools','An auditor who needs TeamMate+ or a template now has a reason to start here rather than at a bookmark.'],
          ['The numbers move','Plan delivery and issue ageing change every week. Static pages are read once; live ones are checked.'],
          ['It updates itself','Highlighted content queries the library, so publishing a document refreshes the page with no one editing it.'],
          ['It has an audience outside','The “being audited” path gives other departments a reason to open and share the link.']];
 m.forEach((t,i)=>{
   const x=0.7+(i%2)*6.35, y=2.05+Math.floor(i/2)*1.85;
   card(s,x,y,5.95,1.6);
   chip(s,x+0.4,y+0.32,i+1,false);
   s.addText(t[0],{x:x+1.0,y:y+0.28,w:4.7,h:0.34,isTextBox:true,margin:0,fontFace:HEAD,fontSize:16,bold:true,color:INK});
   s.addText(t[1],{x:x+1.0,y:y+0.68,w:4.65,h:0.8,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13,color:BODY,lineSpacing:18});
 });
 s.addShape(p.ShapeType.roundRect,{x:0.7,y:5.82,w:11.93,h:0.95,rectRadius:0.06,
   fill:{color:'FFF4E5'},line:{color:'F0D5AE',width:1}});
 s.addText([{text:'How we will know:  ',options:{bold:true,color:'8A5A12'}},
   {text:'we baseline visits and unique visitors from SharePoint Site usage the week before launch, and review at 30 and 90 days. No target is claimed until that baseline exists.',options:{color:'8A5A12'}}],
   {x:1.05,y:6.02,w:11.2,h:0.6,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13.5,lineSpacing:18});
}

/* ---------------- 11 PLAN ---------------- */
{
 const s=light();
 kicker(s,'Delivery',false);
 title(s,'Three weeks, no developer, no new licence.',false);
 const ph=[['Week 1','Foundation','Theme applied, page skeleton built, navigation rewritten to seven destinations.'],
           ['Week 2','Content','Icon set, the three audience paths written, methodology index and request form live.'],
           ['Week 3','Live data','Engagements and Issues lists built, Power BI report pinned, then publish and baseline.']];
 ph.forEach((t,i)=>{
   const x=0.7+i*4.06;
   s.addShape(p.ShapeType.roundRect,{x,y:2.2,w:3.8,h:3.25,rectRadius:0.06,
     fill:{color:i===2?NAVY:SURF},line:{color:i===2?NAVY:LINE,width:1},shadow:sh()});
   s.addText(t[0].toUpperCase(),{x:x+0.38,y:2.5,w:3,h:0.3,isTextBox:true,margin:0,
     fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:i===2?CYAN:BLUE});
   s.addText(t[1],{x:x+0.38,y:2.88,w:3.1,h:0.45,isTextBox:true,margin:0,
     fontFace:HEAD,fontSize:21,bold:true,color:i===2?WHITE:INK});
   s.addText(t[2],{x:x+0.38,y:3.48,w:3.08,h:1.6,isTextBox:true,margin:0,
     fontFace:BODYF,fontSize:13.5,color:i===2?'B9C6EF':BODY,lineSpacing:19});
 });
 s.addText('Roughly seven working days of effort inside the IA Excellence team.',
   {x:0.7,y:5.75,w:11.9,h:0.4,isTextBox:true,margin:0,fontFace:BODYF,fontSize:14,color:BODY});
}

/* ---------------- 12 THE ASK ---------------- */
{
 const s=dark();
 s.addImage({path:'banner-navy-hatch.png',x:0,y:0,w:W,h:H,sizing:{type:'cover',w:W,h:H}});
 s.addShape(p.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:NAVY,transparency:18}});
 kicker(s,'What we need',true);
 title(s,'Four decisions, then we build.',true);
 const asks=[['Open the site to the organisation','Read access for all staff, working files moved to a closed library.'],
   ['Power BI licences','For everyone expected to view the dashboard, or it shows them a sign-in prompt.'],
   ['Three named owners','For methodology, TeamMate+ access and analytics requests. Not a shared mailbox.'],
   ['Agreement on the audited-party page','The content is ours to write; the tone is a management call.']];
 asks.forEach((a,i)=>{
   const y=2.2+i*1.08;
   chip(s,0.72,y+0.06,i+1,true);
   s.addText(a[0],{x:1.4,y:y,w:5.6,h:0.36,isTextBox:true,margin:0,fontFace:HEAD,fontSize:17,bold:true,color:WHITE});
   s.addText(a[1],{x:1.4,y:y+0.38,w:6.0,h:0.55,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13.5,color:'B9C6EF',lineSpacing:18});
 });
 s.addShape(p.ShapeType.roundRect,{x:8.1,y:2.2,w:4.5,h:4.3,rectRadius:0.06,
   fill:{color:DEEP},line:{color:'1C2470',width:1}});
 s.addText('And after this',{x:8.5,y:2.55,w:3.8,h:0.34,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:CYAN});
 s.addText('The in-house platform',{x:8.5,y:2.95,w:3.8,h:0.85,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:24,bold:true,color:WHITE});
 s.addText('This build is the prototype of its data model, not a detour from it. The Engagements and Issues lists carry across; the SharePoint page becomes the front door while the platform is built behind it.',
   {x:8.5,y:3.95,w:3.75,h:2.2,isTextBox:true,margin:0,fontFace:BODYF,fontSize:13.5,color:'B9C6EF',lineSpacing:19});
}

p.writeFile({fileName:'IA-Platform-Revamp.pptx'}).then(f=>console.log('wrote',f));
