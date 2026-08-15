
const APP = {
  NAME: 'JAGSONS SMART SITE',
  TZ: 'Asia/Kolkata',
  SHEETS: {
    SETTINGS:'Settings', USERS:'Users', PROJECTS:'Projects', DPR:'DPR', PHOTOS:'Site_Photos',
    ATTENDANCE:'Attendance', MATERIAL:'Material', MACHINERY:'Machinery', VEHICLES:'Vehicles',
    PLANT:'Plant', MEASUREMENTS:'Measurements', RABILLS:'RA_Bills', PURCHASE:'Purchase',
    QUALITY:'Quality', SAFETY:'Safety', DOCUMENTS:'Documents', EXPENSES:'Expenses',
    CONTRACTORS:'Contractors', APPROVALS:'Approvals', AUDIT:'Audit'
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1');
}
function setupJagsonsSmartSite(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const defs={};
  defs[APP.SHEETS.SETTINGS]=['Key','Value'];
  defs[APP.SHEETS.USERS]=['UserID','Name','Email','Mobile','Role','PIN','Active','CreatedAt'];
  defs[APP.SHEETS.PROJECTS]=['ProjectID','ProjectName','ProjectCode','Client','Department','TenderNo','WorkOrderNo','ContractValue','StartDate','CompletionDate','ProjectManager','SiteEngineer','Location','Latitude','Longitude','RadiusM','ProjectType','Status','PhysicalProgress','FinancialProgress','CreatedBy','CreatedAt','UpdatedAt'];
  defs[APP.SHEETS.DPR]=['DPRID','Date','ProjectID','ProjectName','Location','ChainageFrom','ChainageTo','LengthM','WidthM','ThicknessMM','AreaSQM','Quantity','Unit','BOQItem','WorkDescription','Weather','Contractor','Engineer','Supervisor','Skilled','Unskilled','Operators','Drivers','Supervisors','Machinery','MaterialSummary','Remarks','PhotoURL','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.PHOTOS]=['PhotoID','Date','ProjectID','ProjectName','WorkType','Chainage','Stage','Latitude','Longitude','PhotoURL','Caption','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.ATTENDANCE]=['AttendanceID','Date','Time','StaffName','Mobile','Designation','ProjectID','ProjectName','Status','Punch','Latitude','Longitude','AccuracyM','DistanceFromSiteM','WithinRadius','PhotoURL','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.MATERIAL]=['TxnID','Date','Type','ProjectID','ProjectName','Material','Qty','Unit','Supplier','Vehicle','Challan','Invoice','Source','Destination','SourceDPRID','PhotoURL','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.MACHINERY]=['LogID','Date','ProjectID','ProjectName','Machine','AssetNo','Operator','StartReading','EndReading','TotalHours','DieselL','WorkDone','Breakdown','MaintenanceRequired','PhotoURL','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.VEHICLES]=['TripID','Date','ProjectID','ProjectName','VehicleNo','Driver','TripStart','TripEnd','StartKM','EndKM','DistanceKM','DieselL','Trips','From','To','Latitude','Longitude','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.PLANT]=['ProdID','Date','PlantType','ProjectID','ProjectName','MixGrade','ProductionQty','Unit','Bitumen','Cement','Aggregate','Sand','CrusherSand','Admixture','LDO','Diesel','StartTime','EndTime','Temperature','DispatchQty','Destination','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.MEASUREMENTS]=['MBID','Date','ProjectID','ProjectName','MBNo','BOQItem','Description','ChainageFrom','ChainageTo','MeasurementType','Length','Width','HeightThickness','Nos','Quantity','Unit','Engineer','Remarks','PhotoURL','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.RABILLS]=['RAID','Date','ProjectID','ProjectName','RANo','BOQItem','BOQQty','PreviousQty','CurrentQty','TotalQty','BalanceQty','Rate','GrossAmount','GST','Retention','SecurityDeposit','AdvanceRecovery','MaterialRecovery','OtherRecovery','NetAmount','Status','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.PURCHASE]=['PurchaseID','Date','ProjectID','ProjectName','Stage','MaterialService','Qty','Unit','Vendor','QuotationRef','Amount','RequestedBy','ApprovedBy','Status','PORef','GRNRef','InvoiceRef','Remarks','CreatedAt'];
  defs[APP.SHEETS.QUALITY]=['QualityID','Date','ProjectID','ProjectName','TestType','LocationChainage','RequiredValue','ActualValue','Result','LabReportURL','PhotoURL','Remarks','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.SAFETY]=['SafetyID','Date','ProjectID','ProjectName','Type','Location','Description','Severity','ActionTaken','PhotoURL','Status','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.DOCUMENTS]=['DocumentID','Date','ProjectID','ProjectName','DocumentType','DocumentNo','Title','Revision','FileURL','Remarks','UploadedBy','CreatedAt'];
  defs[APP.SHEETS.EXPENSES]=['ExpenseID','Date','ProjectID','ProjectName','Category','Description','Amount','ReceiptURL','PaidBy','CreatedBy','CreatedAt'];
  defs[APP.SHEETS.CONTRACTORS]=['ContractorID','ContractorName','WorkType','GST','PAN','Contact','ProjectID','WorkOrder','Rate','Quantity','BilledAmount','PaidAmount','Balance','Status','CreatedAt'];
  defs[APP.SHEETS.APPROVALS]=['ApprovalID','Date','Module','RecordID','ProjectID','RequestedBy','Approver','Status','Comment','ActionDate'];
  defs[APP.SHEETS.AUDIT]=['AuditID','Timestamp','User','Action','Module','RecordID','Details'];

  Object.keys(defs).forEach(n=>{
    let sh=ss.getSheetByName(n); if(!sh) sh=ss.insertSheet(n);
    if(sh.getLastRow()===0){
      sh.getRange(1,1,1,defs[n].length).setValues([defs[n]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,defs[n].length).setBackground('#0B2B5C').setFontColor('#fff').setFontWeight('bold');
      sh.autoResizeColumns(1,defs[n].length);
    }
  });
  const settings=sheet_(APP.SHEETS.SETTINGS);
  if(settings.getLastRow()<=1) settings.getRange(2,1,5,2).setValues([
    ['Company','Jagsons Buildcon Ltd.'],['AppName','JAGSONS SMART SITE'],['DefaultRadiusM','300'],['Timezone',APP.TZ],['Version','COMPLETE-1.0']
  ]);
  if(sheet_(APP.SHEETS.USERS).getLastRow()<=1){
    sheet_(APP.SHEETS.USERS).appendRow([uid_('USR'),'Admin',Session.getActiveUser().getEmail()||'','', 'SUPER_ADMIN','1234',true,new Date()]);
  }
  return {ok:true,message:'Complete JAGSONS SMART SITE database created.'};
}

function bootApp(){
  return {user:currentUser_(),projects:getProjects(),dashboard:getDashboard()};
}
function currentUser_(){
  const email=Session.getActiveUser().getEmail()||'';
  const r=rows_(APP.SHEETS.USERS);
  const u=r.find(x=>String(x.Email||'').toLowerCase()===email.toLowerCase() && truthy_(x.Active)) || r.find(x=>truthy_(x.Active));
  return u?{id:u.UserID,name:u.Name,email:u.Email,role:u.Role}:{id:'GUEST',name:email||'User',email,role:'SITE_ENGINEER'};
}
function getProjects(){
  return rows_(APP.SHEETS.PROJECTS).map(x=>({
    id:x.ProjectID,name:x.ProjectName,code:x.ProjectCode,client:x.Client,department:x.Department,
    tenderNo:x.TenderNo,workOrderNo:x.WorkOrderNo,contractValue:num_(x.ContractValue),
    startDate:x.StartDate,completionDate:x.CompletionDate,projectManager:x.ProjectManager,
    siteEngineer:x.SiteEngineer,location:x.Location,lat:num_(x.Latitude),lng:num_(x.Longitude),
    radius:num_(x.RadiusM)||300,projectType:x.ProjectType,status:x.Status||'ACTIVE',
    physicalProgress:num_(x.PhysicalProgress),financialProgress:num_(x.FinancialProgress)
  }));
}

function saveProject(p){
  require_(p && String(p.name||'').trim(),'Project name required');
  require_(p && String(p.code||'').trim(),'Project code required');
  const sh=sheet_(APP.SHEETS.PROJECTS),u=currentUser_(),id=p.id||uid_('PRJ'),all=rows_(APP.SHEETS.PROJECTS);
  const codeKey=String(p.code||'').trim().toLowerCase(),nameKey=String(p.name||'').trim().toLowerCase();
  require_(!all.find(x=>String(x.ProjectCode||'').trim().toLowerCase()===codeKey && String(x.ProjectID)!==String(id)),'Project Code already exists');
  require_(!all.find(x=>String(x.ProjectName||'').trim().toLowerCase()===nameKey && String(x.ProjectID)!==String(id)),'Project Name already exists');
  const rowno=findRowById_(sh,id);
  const createdAt=rowno?sh.getRange(rowno,22).getValue():new Date();
  const row=[
    id,String(p.name||'').trim(),String(p.code||'').trim(),p.client||'',p.department||'',p.tenderNo||'',p.workOrderNo||'',
    num_(p.contractValue),p.startDate||'',p.completionDate||'',p.projectManager||'',p.siteEngineer||'',p.location||'',
    num_(p.lat),num_(p.lng),num_(p.radius)||300,p.projectType||'Other Civil Works',p.status||'ACTIVE',
    num_(p.physicalProgress),num_(p.financialProgress),u.name,createdAt,new Date()
  ];
  if(rowno) sh.getRange(rowno,1,1,row.length).setValues([row]); else sh.appendRow(row);
  audit_(rowno?'UPDATE':'SAVE','PROJECTS',id,p.name);
  return {ok:true,id,message:rowno?'Project updated successfully':'Project saved successfully'};
}

function deleteProject(projectId){
  require_(projectId,'Project ID required');
  const sh=sheet_(APP.SHEETS.PROJECTS),row=findRowById_(sh,projectId);
  require_(row,'Project not found');
  const name=sh.getRange(row,2).getValue();
  sh.deleteRow(row);
  audit_('DELETE','PROJECTS',projectId,name);
  return {ok:true,message:'Project deleted successfully'};
}

function saveDPR(d){
  require_(d.projectId&&d.date&&d.workItem,'Project, Date and Work Item required');
  const p=projectById_(d.projectId),u=currentUser_(),id=uid_('DPR'),L=num_(d.length),W=num_(d.width),T=num_(d.thicknessMM);
  const area=L*W, qty=(d.unit==='Cum'||d.unit==='CUM')?area*(T/1000): (d.unit==='Sqm'||d.unit==='SQM')?area:num_(d.quantity)||area*(T/1000);
  const mats=(d.materials||[]).filter(m=>m.material&&num_(m.qty)>0);
  const matSummary=mats.map(m=>`${m.material} ${m.qty} ${m.unit}`).join(', ');
  sheet_(APP.SHEETS.DPR).appendRow([id,d.date,d.projectId,p.name,d.location||'',d.chainageFrom||'',d.chainageTo||'',L,W,T,area,qty,d.unit||'Cum',d.boqItem||d.workItem,d.workDescription||d.workItem,d.weather||'',d.contractor||'',d.engineer||u.name,d.supervisor||'',num_(d.skilled),num_(d.unskilled),num_(d.operators),num_(d.drivers),num_(d.supervisors),d.machinery||'',matSummary,d.remarks||'',d.photoURL||'',u.name,new Date()]);
  mats.forEach(m=>sheet_(APP.SHEETS.MATERIAL).appendRow([uid_('MAT'),d.date,'CONSUMPTION',d.projectId,p.name,m.material,num_(m.qty),m.unit||'','','',id,'','DPR','SITE',id,'',u.name,new Date()]));
  audit_('SAVE','DPR',id,p.name); return {ok:true,id,area,quantity:qty,message:'DPR saved + Material auto-linked'};
}

function saveAttendance(a){
  require_(a.projectId&&a.staffName,'Project and Staff required');
  const p=projectById_(a.projectId),u=currentUser_(),id=uid_('ATT');
  const dist=(isFinite(Number(a.lat))&&isFinite(Number(a.lng))&&isFinite(Number(p.lat))&&isFinite(Number(p.lng)))?haversine_(Number(a.lat),Number(a.lng),Number(p.lat),Number(p.lng)):0;
  const within=dist<=(Number(p.radius)||300);
  sheet_(APP.SHEETS.ATTENDANCE).appendRow([id,dateStr_(new Date()),timeStr_(new Date()),a.staffName,a.mobile||'',a.designation||'',a.projectId,p.name,a.status||'Present',a.punch||'IN',num_(a.lat),num_(a.lng),num_(a.accuracy),dist,within?'YES':'NO',a.photoURL||'',u.name,new Date()]);
  audit_('SAVE','ATTENDANCE',id,a.staffName); return {ok:true,id,distance:dist,within};
}

function saveMaterial(m){ genericSave_('MATERIAL',m); return {ok:true}; }
function saveMachinery(m){ genericSave_('MACHINERY',m); return {ok:true}; }
function saveVehicle(m){ genericSave_('VEHICLES',m); return {ok:true}; }
function savePlant(m){ genericSave_('PLANT',m); return {ok:true}; }
function saveMeasurement(m){
  require_(m.projectId,'Project required'); const p=projectById_(m.projectId),u=currentUser_(),id=uid_('MB');
  const L=num_(m.length),W=num_(m.width),H=num_(m.height),N=num_(m.nos)||1, type=m.measurementType||'VOLUME';
  let q=num_(m.quantity);
  if(type==='AREA') q=L*W*N;
  if(type==='VOLUME') q=L*W*H*N;
  if(type==='LENGTH') q=L*N;
  if(type==='NOS') q=N;
  sheet_(APP.SHEETS.MEASUREMENTS).appendRow([id,m.date||dateStr_(new Date()),m.projectId,p.name,m.mbNo||'',m.boqItem||'',m.description||'',m.chainageFrom||'',m.chainageTo||'',type,L,W,H,N,q,m.unit||'',m.engineer||u.name,m.remarks||'',m.photoURL||'',u.name,new Date()]);
  audit_('SAVE','MEASUREMENT',id,p.name); return {ok:true,id,quantity:q};
}
function saveRABill(m){ genericSave_('RABILLS',m); return {ok:true}; }
function savePurchase(m){ genericSave_('PURCHASE',m); return {ok:true}; }
function saveQuality(m){ genericSave_('QUALITY',m); return {ok:true}; }
function saveSafety(m){ genericSave_('SAFETY',m); return {ok:true}; }
function saveDocument(m){ genericSave_('DOCUMENTS',m); return {ok:true}; }
function saveExpense(m){ genericSave_('EXPENSES',m); return {ok:true}; }

function genericSave_(module,m){
  const map={
    MATERIAL:APP.SHEETS.MATERIAL,MACHINERY:APP.SHEETS.MACHINERY,VEHICLES:APP.SHEETS.VEHICLES,PLANT:APP.SHEETS.PLANT,
    RABILLS:APP.SHEETS.RABILLS,PURCHASE:APP.SHEETS.PURCHASE,QUALITY:APP.SHEETS.QUALITY,SAFETY:APP.SHEETS.SAFETY,
    DOCUMENTS:APP.SHEETS.DOCUMENTS,EXPENSES:APP.SHEETS.EXPENSES
  };
  const sh=sheet_(map[module]),p=m.projectId?projectById_(m.projectId):{name:''},u=currentUser_(),id=uid_(module.slice(0,3));
  if(module==='MATERIAL') sh.appendRow([id,m.date||dateStr_(new Date()),m.type||'INWARD',m.projectId||'',p.name,m.material||'',num_(m.qty),m.unit||'',m.supplier||'',m.vehicle||'',m.challan||'',m.invoice||'',m.source||'',m.destination||'','',m.photoURL||'',u.name,new Date()]);
  if(module==='MACHINERY') sh.appendRow([id,m.date,p.id,p.name,m.machine||'',m.assetNo||'',m.operator||'',num_(m.startReading),num_(m.endReading),num_(m.totalHours)||Math.max(0,num_(m.endReading)-num_(m.startReading)),num_(m.diesel),m.workDone||'',m.breakdown||'',m.maintenance||'',m.photoURL||'',u.name,new Date()]);
  if(module==='VEHICLES') sh.appendRow([id,m.date,p.id,p.name,m.vehicleNo||'',m.driver||'',m.tripStart||'',m.tripEnd||'',num_(m.startKM),num_(m.endKM),Math.max(0,num_(m.endKM)-num_(m.startKM)),num_(m.diesel),num_(m.trips),m.from||'',m.to||'',num_(m.lat),num_(m.lng),u.name,new Date()]);
  if(module==='PLANT') sh.appendRow([id,m.date,m.plantType||'ASPHALT',p.id,p.name,m.mixGrade||'',num_(m.productionQty),m.unit||'MT',num_(m.bitumen),num_(m.cement),num_(m.aggregate),num_(m.sand),num_(m.crusherSand),num_(m.admixture),num_(m.ldo),num_(m.diesel),m.startTime||'',m.endTime||'',num_(m.temperature),num_(m.dispatchQty),m.destination||'',u.name,new Date()]);
  if(module==='RABILLS'){const gross=num_(m.currentQty)*num_(m.rate);const rec=num_(m.retention)+num_(m.securityDeposit)+num_(m.advanceRecovery)+num_(m.materialRecovery)+num_(m.otherRecovery);sh.appendRow([id,m.date,p.id,p.name,m.raNo||'',m.boqItem||'',num_(m.boqQty),num_(m.previousQty),num_(m.currentQty),num_(m.previousQty)+num_(m.currentQty),Math.max(0,num_(m.boqQty)-num_(m.previousQty)-num_(m.currentQty)),num_(m.rate),gross,num_(m.gst),num_(m.retention),num_(m.securityDeposit),num_(m.advanceRecovery),num_(m.materialRecovery),num_(m.otherRecovery),gross+num_(m.gst)-rec,m.status||'DRAFT',u.name,new Date()]);}
  if(module==='PURCHASE') sh.appendRow([id,m.date,p.id,p.name,m.stage||'REQUEST',m.item||'',num_(m.qty),m.unit||'',m.vendor||'',m.quotationRef||'',num_(m.amount),u.name,m.approvedBy||'',m.status||'PENDING',m.poRef||'',m.grnRef||'',m.invoiceRef||'',m.remarks||'',new Date()]);
  if(module==='QUALITY') sh.appendRow([id,m.date,p.id,p.name,m.testType||'',m.location||'',m.requiredValue||'',m.actualValue||'',m.result||'',m.labReportURL||'',m.photoURL||'',m.remarks||'',u.name,new Date()]);
  if(module==='SAFETY') sh.appendRow([id,m.date,p.id,p.name,m.type||'',m.location||'',m.description||'',m.severity||'LOW',m.actionTaken||'',m.photoURL||'',m.status||'OPEN',u.name,new Date()]);
  if(module==='DOCUMENTS') sh.appendRow([id,m.date,p.id,p.name,m.documentType||'',m.documentNo||'',m.title||'',m.revision||'',m.fileURL||'',m.remarks||'',u.name,new Date()]);
  if(module==='EXPENSES') sh.appendRow([id,m.date,p.id,p.name,m.category||'',m.description||'',num_(m.amount),m.receiptURL||'',m.paidBy||'',u.name,new Date()]);
  audit_('SAVE',module,id,p.name);
}

function getDashboard(){
  const today=dateStr_(new Date());
  const material=count_(APP.SHEETS.MATERIAL), machinery=count_(APP.SHEETS.MACHINERY),
        vehicles=count_(APP.SHEETS.VEHICLES), plant=count_(APP.SHEETS.PLANT),
        measurements=count_(APP.SHEETS.MEASUREMENTS), raBills=count_(APP.SHEETS.RABILLS);
  return {
    projects:getProjects().length,
    dprToday:rows_(APP.SHEETS.DPR).filter(x=>x.Date===today).length,
    attendanceToday:rows_(APP.SHEETS.ATTENDANCE).filter(x=>x.Date===today).length,
    material, machinery, vehicles, plant, measurements, raBills,
    purchase:count_(APP.SHEETS.PURCHASE), quality:count_(APP.SHEETS.QUALITY),
    safety:count_(APP.SHEETS.SAFETY), documents:count_(APP.SHEETS.DOCUMENTS),

    // legacy dashboard aliases
    materialTxns:material,
    machineryLogs:machinery,
    vehicleLogs:vehicles,
    plantLogs:plant,
    mbCount:measurements
  };
}
function getRecent(module,limit){
  const map={DPR:APP.SHEETS.DPR,ATTENDANCE:APP.SHEETS.ATTENDANCE,MATERIAL:APP.SHEETS.MATERIAL,MACHINERY:APP.SHEETS.MACHINERY,VEHICLES:APP.SHEETS.VEHICLES,PLANT:APP.SHEETS.PLANT,MEASUREMENTS:APP.SHEETS.MEASUREMENTS,RABILLS:APP.SHEETS.RABILLS,PURCHASE:APP.SHEETS.PURCHASE,QUALITY:APP.SHEETS.QUALITY,SAFETY:APP.SHEETS.SAFETY,DOCUMENTS:APP.SHEETS.DOCUMENTS};
  return rows_(map[module]).slice(-(limit||100)).reverse();
}
function uploadFile(dataUrl,folder,fileName){
  const m=String(dataUrl||'').match(/^data:([^;]+);base64,(.+)$/); require_(m,'Invalid file');
  const root=getRootFolder_(),sub=getOrCreateFolder_(root,folder||'General');
  const f=sub.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],fileName||('file_'+Date.now())));
  return {ok:true,url:f.getUrl(),id:f.getId()};
}
function getRootFolder_(){
  const ps=PropertiesService.getScriptProperties(); let id=ps.getProperty('JSS_ROOT_FOLDER'),f;
  if(id){try{f=DriveApp.getFolderById(id)}catch(e){}}
  if(!f){f=DriveApp.createFolder('JAGSONS_SMART_SITE_FILES');ps.setProperty('JSS_ROOT_FOLDER',f.getId())}
  return f;
}
function getOrCreateFolder_(p,n){const it=p.getFoldersByName(n);return it.hasNext()?it.next():p.createFolder(n)}
function projectById_(id){const p=getProjects().find(x=>String(x.id)===String(id));require_(p,'Project not found');return p}
function count_(n){return Math.max(0,sheet_(n).getLastRow()-1)}
function sheet_(n){const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);if(!s)throw new Error('Missing sheet '+n+'; run setupJagsonsSmartSite()');return s}
function rows_(n){const sh=sheet_(n),lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)return[];const a=sh.getRange(1,1,lr,lc).getDisplayValues(),h=a.shift();return a.map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]])))}
function findRowById_(sh,id){if(sh.getLastRow()<2)return 0;const v=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();const i=v.findIndex(x=>String(x)===String(id));return i<0?0:i+2}
function uid_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase()}
function num_(v){const n=Number(v);return isFinite(n)?n:0}
function require_(c,m){if(!c)throw new Error(m)}
function truthy_(v){return v===true||String(v).toLowerCase()==='true'||String(v)==='1'}
function dateStr_(d){return Utilities.formatDate(new Date(d),APP.TZ,'yyyy-MM-dd')}
function timeStr_(d){return Utilities.formatDate(new Date(d),APP.TZ,'HH:mm:ss')}
function haversine_(a,b,c,d){const R=6371000,r=x=>x*Math.PI/180,x=r(c-a),y=r(d-b),q=Math.sin(x/2)**2+Math.cos(r(a))*Math.cos(r(c))*Math.sin(y/2)**2;return 2*R*Math.atan2(Math.sqrt(q),Math.sqrt(1-q))}
function audit_(a,m,id,d){try{sheet_(APP.SHEETS.AUDIT).appendRow([uid_('AUD'),new Date(),currentUser_().name,a,m,id,d||''])}catch(e){}}
