const pptx=require('pptxgenjs');
const p=new pptx(); p.layout='LAYOUT_WIDE';
const W=13.333,H=7.5;
const NAVY='000050',DEEP='000033',BLUE='1061FF',CYAN='4BF5FF',WHITE='FFFFFF',
      INK='101540',BODY='4D5680',MUTE='8F9EDA',LINE='DDE8FF',SURF='F4F7FF',
      GREY='8A8A8A',GREYBG='F2F2F2',ORANGE='FF593C';
const HEAD='Arial',BODYF='Calibri';
const sh=()=>({type:'outer',color:'101540',blur:16,offset:3,angle:90,opacity:0.12});

/* ============ SLIDE 1 — look and feel ============ */
{
 const s=p.addSlide(); s.background={color:WHITE};
 s.addText('LOOK AND FEEL',{x:0.6,y:0.3,w:8,h:0.3,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2.4,color:BLUE});
 s.addText('What it is today, and what it becomes.',{x:0.6,y:0.62,w:12.1,h:0.62,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:31,bold:true,color:INK});

 // BEFORE
 s.addShape(p.ShapeType.roundRect,{x:0.6,y:1.48,w:6.0,h:5.35,rectRadius:0.04,
   fill:{color:GREYBG},line:{color:'E0E0E0',width:1}});
 s.addText('TODAY',{x:0.95,y:1.72,w:2,h:0.28,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:'767676'});
 s.addImage({path:'cmp-before-crop.png',x:0.95,y:2.10,w:2.55,h:2.86,sizing:{type:'crop',w:2.55,h:2.86}});
 const bef=[['A noticeboard','Opens with photo carousel, birthdays and the mission statement.'],
            ['Default SharePoint','Stock theme, stock icons, mixed image crops, no brand colour.'],
            ['Read once','Every block static. Links repeat. Nothing changes week to week.'],
            ['Tools absent','No TeamMate+, no manual, no dashboard on the first screen.']];
 bef.forEach((t,i)=>{
   const y=2.10+i*0.82;
   s.addText(t[0],{x:3.72,y,w:2.7,h:0.26,isTextBox:true,margin:0,fontFace:HEAD,fontSize:12.5,bold:true,color:'3B3B3B'});
   s.addText(t[1],{x:3.72,y:y+0.27,w:2.62,h:0.56,isTextBox:true,margin:0,fontFace:BODYF,fontSize:11,color:'767676',lineSpacing:14});
 });

 // arrow
 s.addShape(p.ShapeType.rightArrow,{x:6.72,y:4.0,w:0.5,h:0.42,fill:{color:CYAN},line:{color:CYAN,width:1}});

 // AFTER
 s.addShape(p.ShapeType.roundRect,{x:7.32,y:1.48,w:5.4,h:5.35,rectRadius:0.04,
   fill:{color:NAVY},line:{color:NAVY,width:1},shadow:sh()});
 s.addText('PROPOSED',{x:7.62,y:1.72,w:2.4,h:0.28,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2,color:CYAN});
 s.addImage({path:'cmp-after-crop.png',x:7.62,y:2.10,w:2.42,h:2.72,sizing:{type:'crop',w:2.42,h:2.72}});
 const aft=[['A working surface','Tools, methodology and live plan status on the first screen.'],
            ['On brand','Navy, blue and cyan applied through the site theme and card backgrounds.'],
            ['Changes weekly','Dashboard, issue ageing and a self-updating document list.'],
            ['Open to the business','A path written for the departments we audit.']];
 aft.forEach((t,i)=>{
   const y=2.10+i*0.82;
   s.addText(t[0],{x:10.26,y,w:2.3,h:0.26,isTextBox:true,margin:0,fontFace:HEAD,fontSize:12.5,bold:true,color:CYAN});
   s.addText(t[1],{x:10.26,y:y+0.27,w:2.26,h:0.56,isTextBox:true,margin:0,fontFace:BODYF,fontSize:11,color:'B9C6EF',lineSpacing:14});
 });
 s.addText('Same SharePoint. Same web parts. Same licences. The change is what the page is for.',
   {x:0.6,y:6.95,w:12.1,h:0.3,isTextBox:true,margin:0,fontFace:BODYF,fontSize:12.5,italic:true,color:BODY});
 s.addNotes('Left is the structure of the page as it stands. Right is the proposal, built from web parts already in our Toolbox.');
}

/* ============ SLIDE 2 — features ============ */
{
 const s=p.addSlide(); s.background={color:WHITE};
 s.addText('WHAT THE PLATFORM CARRIES',{x:0.6,y:0.3,w:8,h:0.3,isTextBox:true,margin:0,
   fontFace:HEAD,fontSize:11,bold:true,charSpacing:2.4,color:BLUE});
 s.addText('Everything kept. Five added. One new layer.',
   {x:0.6,y:0.62,w:12.1,h:0.62,isTextBox:true,margin:0,fontFace:HEAD,fontSize:29,bold:true,color:INK});

 const NEW='NEW', OLD='KEPT';
 const cols=[
  {t:'Get the work done',c:BLUE,items:[
    ['Quick links — TM+, Manuals, Shared folder',NEW],
    ['Assurance dashboard (Power BI)',NEW],
    ['ATHEER bot — ask the manual',NEW],
    ['Audit Analytics & Innovation requests',OLD],
    ['IA Document Library',OLD],
    ['IA Vacation Plan',OLD],
    ['Mentorship Programme',OLD]]},
  {t:'People & culture',c:'8F4FD8',items:[
    ['Our team page — who to ask',NEW],
    ['Breakfast fund board',NEW],
    ['News & announcements',OLD],
    ['Team birthday calendar',OLD],
    ['IdeIA initiative form',OLD],
    ['Recognition & SPOC awards',OLD]]},
  {t:'Gamification layer',c:'C27803',items:[
    ['Question of the week',NEW],
    ['Points, streaks & weekly leaderboard',NEW],
    ['CPE badges at 10 / 25 / 40 hours',NEW],
    ['Mentor badge for hours given',NEW],
    ['IdeIA points for ideas shortlisted',NEW],
    ['First-30-days onboarding progress',NEW]]}];

 cols.forEach((col,ci)=>{
   const x=0.6+ci*4.11;
   s.addShape(p.ShapeType.roundRect,{x,y:1.52,w:3.9,h:3.95,rectRadius:0.05,
     fill:{color:SURF},line:{color:LINE,width:1},shadow:sh()});
   s.addShape(p.ShapeType.roundRect,{x:x+0.28,y:1.78,w:0.30,h:0.30,rectRadius:0.07,
     fill:{color:col.c},line:{color:col.c,width:1}});
   s.addText(col.t,{x:x+0.70,y:1.76,w:3.0,h:0.34,isTextBox:true,margin:0,
     fontFace:HEAD,fontSize:15,bold:true,color:INK});
   col.items.forEach((it,i)=>{
     const y=2.24+i*0.46;
     s.addText(it[0],{x:x+0.28,y,w:2.72,h:0.42,isTextBox:true,margin:0,
       fontFace:BODYF,fontSize:11.5,color:it[1]===NEW?INK:BODY,bold:it[1]===NEW,lineSpacing:14});
     s.addShape(p.ShapeType.roundRect,{x:x+3.06,y:y+0.04,w:0.58,h:0.22,rectRadius:0.1,
       fill:{color:it[1]===NEW?NAVY:'E6EAF5'},line:{color:it[1]===NEW?NAVY:'E6EAF5',width:1}});
     s.addText(it[1],{x:x+3.06,y:y+0.03,w:0.58,h:0.23,isTextBox:true,margin:0,align:'center',valign:'middle',
       fontFace:HEAD,fontSize:7.5,bold:true,charSpacing:0.6,color:it[1]===NEW?CYAN:'7A86A8'});
   });
 });

 // bottom band: proof + principle
 s.addShape(p.ShapeType.roundRect,{x:0.6,y:5.62,w:12.12,h:1.42,rectRadius:0.05,
   fill:{color:NAVY},line:{color:NAVY,width:1}});
 s.addImage({path:'qow.png',x:0.82,y:5.76,w:1.72,h:1.15});
 s.addText('Gamify participation, never audit outcomes.',
   {x:2.80,y:5.82,w:9.7,h:0.32,isTextBox:true,margin:0,fontFace:HEAD,fontSize:15,bold:true,color:CYAN});
 s.addText('Points go to answering the weekly question, logging CPE, mentoring and submitting ideas. Nothing is scored on findings raised or issues closed — scoring those would distort the judgement the department exists to exercise.',
   {x:2.80,y:6.18,w:9.7,h:0.72,isTextBox:true,margin:0,fontFace:BODYF,fontSize:12.5,color:'B9C6EF',lineSpacing:17});
 s.addNotes('Five genuinely new items: quick links to the systems, the assurance dashboard, ATHEER, the team page and the breakfast fund board. Everything else on the current site is carried across, not dropped. The gamification layer is new in full, and is deliberately limited to participation.');
}

p.writeFile({fileName:'IA-Revamp-2slides.pptx'}).then(f=>console.log('wrote',f));
