const d = require('docx');
const {Document,Packer,Paragraph,TextRun,HeadingLevel,AlignmentType,Table,TableRow,TableCell,
  WidthType,ShadingType,BorderStyle,PageBreak,LevelFormat,convertInchesToTwip} = d;
const fs=require('fs');

const INK="14232A", INK2="405058", INK3="6B7C84";
const TEAL="0D6B6B", SLATE="3D5A6C", OCHRE="8A6410", DIM="77878E", FLAG="8C3A22";
const RULE="DCE4E4", SOFT="F1F5F5";
const SERIF="Georgia", SANS="Segoe UI", MONO="Consolas";
const NONE={style:BorderStyle.NONE,size:0,color:"FFFFFF"};
const noB={top:NONE,bottom:NONE,left:NONE,right:NONE};
const W=9360; // content width for Letter with 1.1" margins

const el=[];
const P=(o)=>{el.push(new Paragraph(o));};

// ---------- helpers ----------
function rule(color=RULE,size=6,space=200){
  return new Paragraph({spacing:{before:space,after:space},
    border:{bottom:{style:BorderStyle.SINGLE,size:size,color:color}}});
}
function label(text,color){
  return new Paragraph({spacing:{before:0,after:60},
    children:[new TextRun({text:text.toUpperCase(),font:MONO,size:15,bold:true,
      color:color||INK3,characterSpacing:26})]});
}
function body(text,opts={}){
  return new Paragraph({spacing:{after:opts.after??100,line:300},
    children:[new TextRun({text,font:SERIF,size:opts.size??21,color:opts.color??INK2,
      italics:opts.italics||false})]});
}
function rich(runs,opts={}){
  return new Paragraph({spacing:{after:opts.after??100,line:300},children:runs});
}
function t(text,o={}){return new TextRun({text,font:o.font||SERIF,size:o.size||21,
  color:o.color||INK2,bold:!!o.bold,italics:!!o.italics});}

// callout: shaded single-cell table
function callout(runs,fill,bar){
  return new Table({
    columnWidths:[W],
    width:{size:W,type:WidthType.DXA},
    borders:{top:NONE,bottom:NONE,right:NONE,
      left:{style:BorderStyle.SINGLE,size:18,color:bar},
      insideHorizontal:NONE,insideVertical:NONE},
    rows:[new TableRow({children:[new TableCell({
      width:{size:W,type:WidthType.DXA},
      shading:{type:ShadingType.CLEAR,fill:fill,color:"auto"},
      margins:{top:170,bottom:170,left:220,right:220},
      children:[new Paragraph({spacing:{after:0,line:300},children:runs})]
    })]})]
  });
}

// a cost-type box, rendered as a 2-col table (number | content)
function box(num,code,name,test,tagsLabel,tags,note,tint){
  const NW=760, CW=W-NW;
  const content=[];
  content.push(new Paragraph({spacing:{after:90},children:[
    new TextRun({text:code,font:MONO,size:20,bold:true,color:tint}),
    new TextRun({text:"   ",font:MONO,size:20}),
    new TextRun({text:name,font:SANS,size:24,bold:true,color:INK}),
  ]}));
  content.push(new Paragraph({spacing:{after:tags?110:60,line:300},children:test}));
  if(tags){
    content.push(label(tagsLabel));
    content.push(new Paragraph({spacing:{after:note?110:0,line:290},
      children:[new TextRun({text:tags.join("  ·  "),font:SANS,size:18,color:INK2})]}));
  }
  if(note){
    content.push(new Paragraph({spacing:{after:0,line:290},children:note}));
  }
  return new Table({
    columnWidths:[NW,CW],
    width:{size:W,type:WidthType.DXA},
    borders:{
      top:{style:BorderStyle.SINGLE,size:4,color:RULE},
      bottom:{style:BorderStyle.SINGLE,size:4,color:RULE},
      left:{style:BorderStyle.SINGLE,size:4,color:RULE},
      right:{style:BorderStyle.SINGLE,size:4,color:RULE},
      insideVertical:{style:BorderStyle.SINGLE,size:4,color:"EBF0F0"},
      insideHorizontal:NONE},
    rows:[new TableRow({cantSplit:true,children:[
      new TableCell({width:{size:NW,type:WidthType.DXA},
        shading:{type:ShadingType.CLEAR,fill:SOFT,color:"auto"},
        margins:{top:170,bottom:170,left:0,right:0},
        children:[new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:0},
          children:[new TextRun({text:String(num).padStart(2,"0"),font:MONO,size:28,
            bold:true,color:tint})]})]}),
      new TableCell({width:{size:CW,type:WidthType.DXA},
        margins:{top:170,bottom:170,left:240,right:200},
        children:content})
    ]})]
  });
}
const gap=(a=110)=>new Paragraph({spacing:{after:a},children:[]});

function section(title,range,sub,brk){
  el.push(new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:brk?0:340,after:0},
    keepNext:true,keepLines:true,pageBreakBefore:!!brk,
    children:[new TextRun({text:title,font:SANS,size:30,bold:true,color:INK})]}));
  if(range) el.push(new Paragraph({spacing:{after:60},keepNext:true,
    children:[new TextRun({text:range,font:MONO,size:17,bold:true,color:INK3})]}));
  if(sub) el.push(new Paragraph({spacing:{after:120,line:300},keepNext:true,
    children:[t(sub)]}));
  el.push(new Paragraph({spacing:{before:60,after:60},keepNext:true,
    border:{bottom:{style:BorderStyle.SINGLE,size:10,color:INK}}}));
}

// ================= MASTHEAD =================
el.push(new Paragraph({spacing:{after:80},children:[
  new TextRun({text:"DEPARTMENT OF HEALTH · ABU DHABI CLINICAL COST DATA COLLECTION",
    font:MONO,size:15,bold:true,color:TEAL,characterSpacing:26})]}));
el.push(new Paragraph({heading:HeadingLevel.TITLE,spacing:{after:120},children:[
  new TextRun({text:"The Fifteen Boxes",font:SANS,size:68,bold:true,color:INK})]}));
el.push(new Paragraph({spacing:{after:140,line:320},children:[
  t("Every dirham the clinic spends has to land in one of fifteen boxes. This is what each box means, in plain words, with the things we actually buy.",{size:24,color:INK2})]}));
el.push(new Paragraph({spacing:{after:60},children:[
  new TextRun({text:"Vision Dental Clinic L.L.C     ·     Licence MF4159     ·     DOH Cost Type list",
    font:MONO,size:16,color:INK3})]}));
el.push(rule(INK,16,180));

el.push(callout([
  t("One expense goes in one box. Never two.",{bold:true,color:INK,size:22}),
  new TextRun({break:1}),
  t("If you can put an invoice in two boxes, it belongs in the one that describes what you bought, not who used it. When you genuinely cannot decide, ask Finance rather than guessing — a wrong box is harder to find later than an unanswered question.",{size:21}),
],"E4EFEE",TEAL));
el.push(gap(260));

// ================= PEOPLE =================
section("The people we pay","BOXES 1–4",
  "Every salary in the clinic sits in one of these four. The only question is what kind of job the person does — not which department pays them.");

el.push(box(1,"SWNurs","Nurse salaries",
  [t("Is this person a "),t("registered nurse",{bold:true,color:INK}),t("?")],
  "Who goes here",["Registered nurses"],
  [t("Watch out: ",{bold:true,color:FLAG}),t("only people who actually hold a nursing licence. Someone we call a “nurse” in day-to-day speech but who is really a chair-side helper belongs in box 3 or box 4.")],
  TEAL)); el.push(gap());

el.push(box(2,"SWDoc","Dentists’ salaries",
  [t("Is this person a "),t("dentist",{bold:true,color:INK}),t("?")],
  "Who goes here",["Full-time dentists","Part-time dentists","Commissions paid to dentists"],
  [t("Don’t forget commission. ",{bold:true,color:FLAG}),t("If a dentist earns a share of what they bill, that share belongs here too — not in Goods & Services. DOH calls this “Department #1”, which for us simply means Dental; a hospital would have many.")],
  TEAL)); el.push(gap());

el.push(box(3,"SWAllied","Clinical support salaries",
  [t("Do they do "),t("clinical work",{bold:true,color:INK}),t(", but they’re not a dentist and not a nurse?")],
  "Who goes here",["Dental assistants","Dental technicians","Radiology assistants","CSSD technician"],
  [t("CSSD is the sterilisation department — the person who cleans and sterilises our instruments.",{color:INK3,size:20})],
  TEAL)); el.push(gap());

el.push(box(4,"SWNonClin","Admin & support salaries",
  [t("Does this person "),t("never treat a patient",{bold:true,color:INK}),t("?")],
  "Who goes here",["Reception","Cashier","Insurance coordinator","Clinic manager","PRO","IT administrator","Finance manager","Marketing executive","Cleaner","Housekeeping","Driver"],
  [t("Easy way to remember: ",{bold:true,color:FLAG}),t("if the clinic were a shop instead of a clinic, these jobs would still exist.")],
  TEAL)); el.push(gap());

// ================= PATIENTS =================
section("What we use on patients","BOXES 5–10",
  "Everything that reaches the patient — made for them, taken of them, given to them, or fitted into them.",true);

el.push(box(5,"Laboratory","Lab work",
  [t("Was it "),t("made or tested in a lab",{bold:true,color:INK}),t("?")],
  "What goes here",["Dental lab","Haematology lab","Biochemistry lab","Microbiology lab"],
  [t("For us this is overwhelmingly dental lab — crowns, bridges, dentures, aligners — whether made in-house or sent out. The medical labs are on the DOH list because hospitals use them.",{color:INK3,size:20})],
  SLATE)); el.push(gap());

el.push(box(6,"Imaging","X-rays & scans",
  [t("Is it a "),t("picture of the inside of the patient",{bold:true,color:INK}),t("?")],
  "What goes here",["X-rays","OPG","CBCT","CT scan","Ultrasound","MRI","Mammography"],
  [t("Includes servicing and maintaining the imaging machines.",{color:INK3,size:20})],
  SLATE)); el.push(gap());

el.push(box(7,"Pharmacy","Medicines",
  [t("Is it a "),t("drug",{bold:true,color:INK}),t("?")],
  "What goes here",["Medicine cost","Pharmacy supplies"],null,SLATE)); el.push(gap());

el.push(box(8,"Prostheses","Things fitted into the patient",
  [t("Does the patient "),t("walk out with it in their mouth",{bold:true,color:INK}),t("?")],
  "What goes here",["Dental implant screws","Abutments","Crown caps","Braces","Retainers"],
  null,SLATE)); el.push(gap());

el.push(box(9,"MS","Clinical consumables",
  [t("Is it "),t("single-use, clinical, and binned",{bold:true,color:INK}),t(" at the end of the appointment?")],
  "What goes here",["Gloves","Masks","Gowns","Caps","Shoe covers","Syringes","Needles","Drapes","Gauze","Cotton","Tapes","Blades","Sutures","Stitches","Antiseptics","IV sets","Oxygen masks","Test strips"],
  [t("“MS” stands for Medical & Surgical Supplies.",{color:INK3,size:20})],
  SLATE)); el.push(gap());

el.push(box(10,"Hotel","Feeding & transporting patients",
  [t("Did we "),t("feed the patient, or drive them",{bold:true,color:INK}),t("?")],
  "What goes here for us",["Patient nutrition","Juice & refreshments","Food","Patient car transport"],
  [t("The name is misleading. ",{bold:true,color:FLAG}),t("DOH wrote this box for hospitals, so their wording says bedsheets, gowns and TV. We have no beds, so that half doesn’t apply — but the food and patient-transport half very much does. Keep this explanation with the audit file: a reviewer will ask why a clinic with no beds has a “Hotel” figure.")],
  SLATE)); el.push(gap());

// ================= BUSINESS =================
section("Running the clinic","BOXES 11–14",
  "The cost of keeping the doors open, whether or not a single patient walks in.",true);

el.push(box(11,"GS","Goods & services",
  [t("Is it an "),t("office or admin running cost",{bold:true,color:INK}),t(" — not clinical, not a salary?")],
  "What goes here",["Paper","Printer","Stationery","Uniforms","Furniture","IT hardware","Maintenance & repairs","Pest control","Tasneef audits","Consulting","Software subscriptions","Newspaper","Courier","Petrol","Vehicle renewal","Vehicle insurance","Vehicle maintenance"],
  null,OCHRE)); el.push(gap());

el.push(box(12,"Depreciation","Wear and tear",
  [t("Is this an asset "),t("losing value",{bold:true,color:INK}),t(" rather than money leaving the bank?")],
  "What goes here",["Medical equipment","IT equipment","Furniture","Vehicles"],
  [t("We bought the dental chair once, but it wears out a little every year. That yearly slice is what goes here — no cash actually moves this month.",{color:INK3,size:20})],
  OCHRE)); el.push(gap());

el.push(box(13,"OHF","Facility overheads",
  [t("Would we "),t("still pay it on a day with zero patients",{bold:true,color:INK}),t("?")],
  "What goes here",["Electricity","Water","Internet","Building rent","Employee insurance renewal","Malpractice insurance renewal"],
  [t("OHF = Overhead Cost, Healthcare Facility. The cost of having a building at all.",{color:INK3,size:20})],
  OCHRE)); el.push(gap());

el.push(box(14,"OHC","Head office & marketing",
  [t("Is it "),t("promoting or supporting the business",{bold:true,color:INK}),t(" from above clinic level?")],
  "What goes here",["Marketing campaigns","Website","Advertising","Training"],
  [t("OHC = Overhead Cost, Corporate Office.",{color:INK3,size:20})],
  OCHRE)); el.push(gap(40));

// ================= NA =================
section("The one that doesn’t apply to us","BOX 15",null);
el.push(box(15,"NABMs","Government mandate adjustment",
  [t("Government-funded insurance adjustments. "),t("Not applicable to us — leave at zero.",{bold:true,color:INK})],
  null,null,
  [t("The full name is Negative Adjustment for Budgeted Mandate Services. It exists for facilities that receive government funding outside normal activity billing. We don’t, so this box stays empty every month.",{color:INK3,size:20})],
  DIM));

// ================= PAIRS =================
section("The four that get confused",null,
  "These are the mix-ups that actually happen. Learn these four and the rest look after themselves.");

const pairs=[
 ["Prostheses or MS?","A crown stays in the patient’s mouth. The glove worn while fitting it goes in the bin.","Crown → Prostheses    ·    Glove → MS"],
 ["GS or Depreciation?","Furniture and IT hardware appear on both DOH lists. Our house rule: if we expense it in the year we buy it, it’s GS. If we capitalise it and write it down over several years, only the yearly slice goes to Depreciation.","Never both — that double-counts the money"],
 ["The clinic van, twice?","Yes, and that’s correct. Petrol, insurance and servicing are running costs. The van losing value is wear and tear. Two different facts about the same vehicle.","Running → GS    ·    Losing value → Depreciation"],
 ["OHF or OHC?","OHF is the building — rent, electricity, water, insurance. OHC is everything that promotes or supports the business from above the clinic.","Building → OHF    ·    Brand → OHC"],
];
pairs.forEach(([h,p,v],pi)=>{
  el.push(new Table({
    columnWidths:[W],width:{size:W,type:WidthType.DXA},
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:RULE},
      bottom:{style:BorderStyle.SINGLE,size:4,color:RULE},
      left:{style:BorderStyle.SINGLE,size:4,color:RULE},
      right:{style:BorderStyle.SINGLE,size:4,color:RULE},
      insideHorizontal:NONE,insideVertical:NONE},
    rows:[new TableRow({cantSplit:true,children:[new TableCell({
      width:{size:W,type:WidthType.DXA},
      margins:{top:190,bottom:190,left:240,right:240},
      children:[
        new Paragraph({spacing:{after:80},children:[new TextRun({text:h,font:SANS,size:21,bold:true,color:INK})]}),
        new Paragraph({spacing:{after:90,line:300},children:[t(p)]}),
        new Paragraph({spacing:{after:0},children:[new TextRun({text:v,font:MONO,size:17,bold:true,color:TEAL})]}),
      ]})]})]
  }));
  if(pi<pairs.length-1) el.push(gap(120));
});

// ================= TABLE =================
section("All fifteen, at a glance",null,"Print this page and keep it by the desk.",true);

const rowsData=[
 ["1","SWNurs","Nurse salaries","Are they a registered nurse?"],
 ["2","SWDoc","Dentist salaries","Are they a dentist? Include commission."],
 ["3","SWAllied","Clinical support salaries","Clinical, but not a dentist or a nurse."],
 ["4","SWNonClin","Admin & support salaries","Never treats a patient."],
 ["5","Laboratory","Lab work","Made or tested in a lab. Mostly dental lab."],
 ["6","Imaging","X-rays & scans","A picture of the inside of the patient."],
 ["7","Pharmacy","Medicines","Is it a drug?"],
 ["8","Prostheses","Fitted items","Patient keeps it in their mouth."],
 ["9","MS","Clinical consumables","Single-use, then binned."],
 ["10","Hotel","Patient food & transport","Feeding a patient, or driving them."],
 ["11","GS","Goods & services","Office and admin running costs."],
 ["12","Depreciation","Wear and tear","Equipment losing value, not cash out."],
 ["13","OHF","Facility overheads","Paid even with zero patients."],
 ["14","OHC","Head office & marketing","Promoting or supporting the business."],
 ["15","NABMs","Government mandate adjustment","Not applicable — leave at zero."],
];
const CW=[620,1900,2740,4100];
const hdr=new TableRow({tableHeader:true,children:
  ["#","DOH CODE","PLAIN NAME","THE ONE-LINE TEST"].map((h,i)=>new TableCell({
    width:{size:CW[i],type:WidthType.DXA},
    shading:{type:ShadingType.CLEAR,fill:"EBF0F0",color:"auto"},
    margins:{top:105,bottom:105,left:150,right:150},
    children:[new Paragraph({spacing:{after:0},children:[
      new TextRun({text:h,font:MONO,size:15,bold:true,color:INK3,characterSpacing:22})]})]}))});
const trs=[hdr];
rowsData.forEach(([n,c,nm,ts],ix)=>{
  const fill = ix%2 ? "F7FAFA" : "FFFFFF";
  trs.push(new TableRow({cantSplit:true,children:[
    [n,MONO,INK3,false],[c,MONO,INK,true],[nm,SANS,INK,true],[ts,SERIF,INK2,false]
  ].map(([txt,fnt,col,bd],i)=>new TableCell({
    width:{size:CW[i],type:WidthType.DXA},
    shading:{type:ShadingType.CLEAR,fill:fill,color:"auto"},
    margins:{top:88,bottom:88,left:150,right:150},
    children:[new Paragraph({spacing:{after:0,line:250},children:[
      new TextRun({text:txt,font:fnt,size:fnt===SERIF?20:18,bold:bd,color:col})]})]}))}));
});
el.push(new Table({columnWidths:CW,width:{size:W,type:WidthType.DXA},
  borders:{top:{style:BorderStyle.SINGLE,size:4,color:RULE},
    bottom:{style:BorderStyle.SINGLE,size:4,color:RULE},
    left:{style:BorderStyle.SINGLE,size:4,color:RULE},
    right:{style:BorderStyle.SINGLE,size:4,color:RULE},
    insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:"EBF0F0"},
    insideVertical:{style:BorderStyle.SINGLE,size:2,color:"EBF0F0"}},
  rows:trs}));

// ================= FOOTER =================
el.push(rule(RULE,4,170));
el.push(rich([t("Source. ",{bold:true,color:INK2,size:19}),
  t("DOH Cost Type list (Jan 2024 template), with clinic-specific examples supplied by the Finance Director. Read alongside the Abu Dhabi Clinical Costing Standard, reference DOH/ST/HPS/CC/V1/2024, Appendix 1.",{size:19,color:INK3})],{after:70}));
el.push(rich([t("Note. ",{bold:true,color:INK2,size:19}),
  t("DOH’s Appendix 1 defines fourteen cost types. NABMs is the fifteenth and does not apply to us.",{size:19,color:INK3})],{after:70}));
el.push(rich([t("This is a filing guide only. ",{bold:true,color:INK2,size:19}),
  t("It tells you where an expense belongs. It does not calculate anything.",{size:19,color:INK3})],{after:0}));

const doc=new Document({
  creator:"Vision Dental Clinic L.L.C",
  title:"The Fifteen Boxes",
  description:"Plain-English guide to the fifteen DOH cost types",
  styles:{default:{document:{run:{font:SERIF,size:21,color:INK2}}}},
  sections:[{
    properties:{page:{size:{width:12240,height:15840},
      margin:{top:1440,bottom:1440,left:1440,right:1440}}},
    children:el
  }]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("The_Fifteen_Boxes.docx",b);console.log("written",b.length);});
