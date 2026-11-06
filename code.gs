const SHEET_ID = '1rcQyuYoouYej6On-NJrMS58ggw-mznUDD6WEgfoNQ2o'; 
const LOGS_TAB = 'Sheet1'; 
const MASTERLIST_TAB = 'Masterlist';
const PRODUCTIVITY_TAB = 'Productivity';
const RISK_TAB = 'RiskTrends'; 
const AUDIT_TAB = 'AuditLog'; 
const ADMIN_TAB = 'Admin'; 
const ADMIN_EMAIL = 'raymond.tumanday@iopex.com';

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  
  const props = PropertiesService.getScriptProperties();
  template.bgUrl = props.getProperty('bgUrl') || '';
  template.bgBlur = props.getProperty('bgBlur') || '16';
  template.bgOpacity = props.getProperty('bgOpacity') || '0.7';
  
  if (e.parameter.action === 'review' && e.parameter.id) {
    template.isReview = true;
    template.reviewId = e.parameter.id;
  } else {
    template.isReview = false;
    template.reviewId = "";
  }
  
  return template.evaluate()
      .setTitle('Newsbreak Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  const props = PropertiesService.getScriptProperties();
  const settings = {
    bgUrl: props.getProperty('bgUrl') || '',
    bgBlur: props.getProperty('bgBlur') || '16',
    bgOpacity: props.getProperty('bgOpacity') || '0.7'
  };
  
  const activeEmail = Session.getActiveUser().getEmail();
  const activeEmailLower = activeEmail.toLowerCase().trim();
  
  const isSuperAdmin = (activeEmailLower === ADMIN_EMAIL.toLowerCase().trim());
  let isRegularAdmin = false;
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const adminSheet = ss.getSheetByName(ADMIN_TAB);
    if (adminSheet) {
      const adminData = adminSheet.getRange('A:A').getValues();
      const adminEmails = adminData.map(row => row[0].toString().toLowerCase().trim());
      if (adminEmails.includes(activeEmailLower)) {
        isRegularAdmin = true;
      }
    }
  } catch (e) {
    console.error("Could not read Admin tab.");
  }

  let accessLevel = 'agent';
  if (isSuperAdmin) accessLevel = 'superadmin';
  else if (isRegularAdmin) accessLevel = 'admin';
  
  let name = activeEmail.replace('@iopex.com', '').replace(/\./g, ' ');
  name = name.replace(/\b\w/g, char => char.toUpperCase());
  const moderator = { email: activeEmail, name: name };

  const employees = getEmployeeData();
  const report = getComplianceReport();
  const announcements = JSON.parse(props.getProperty('announcements') || '[]');
  
  return { settings, employees, report, moderator, accessLevel, announcements };
}

function timeToSeconds(timeStr) {
  if (!timeStr) return 0;
  let str = timeStr.toString().trim();
  let parts = str.split(':');
  if (parts.length === 3) {
    return parseInt(parts[0], 10) * 3600 + parseInt(parts[1], 10) * 60 + parseInt(parts[2], 10);
  } else if (parts.length === 2) {
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  return parseFloat(str) || 0;
}

function getExpandedDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    let prodSheet = ss.getSheetByName(PRODUCTIVITY_TAB);
    let prodRows = [];
    
    if (prodSheet) {
      const rawData = prodSheet.getDataRange().getDisplayValues(); 
      if (rawData.length > 1) {
        for (let i = 1; i < rawData.length; i++) {
          let r = rawData[i];
          
          let dateStr = r[0] ? r[0].toString() : '';
          let email = r[1] ? r[1].toString() : ''; 
          let role = r[2] ? r[2].toString() : '';  
          
          let target = parseFloat(r[7] ? r[7].toString().replace(/,/g, '') : '0') || 0;      
          let pph = parseFloat(r[8] ? r[8].toString().replace(/,/g, '') : '0') || 0; 
          let prod = parseFloat(r[9] ? r[9].toString().replace(/,/g, '') : '0') || 0; 
          let revTimeSec = timeToSeconds(r[10]);   
          
          let attainRaw = parseFloat(r[11] ? r[11].toString().replace(/,/g, '').replace(/%/g, '') : '0') || 0;  
          let attain = attainRaw < 2 ? attainRaw * 100 : attainRaw; 
          let acc = parseFloat(r[19] ? r[19].toString().replace(/,/g, '').replace(/%/g, '') : '0') || 0;        
          
          if (email) {
            prodRows.push({
              date: dateStr, email: email, role: role, prod: prod, pph: pph, 
              target: target, attain: attain, acc: acc, revTime: revTimeSec
            });
          }
        }
      }
    }
    
    let riskSheet = ss.getSheetByName(RISK_TAB);
    const riskData = riskSheet ? riskSheet.getDataRange().getDisplayValues() : [];
    const trainingData = "Pending integration with external Google Doc. Add your Google Doc parsing logic here once the document is finalized.";

    return { prodRows, riskData, trainingData };
  } catch (error) {
    throw new Error("Failed to load dashboard data: " + error.message);
  }
}

function getEmployeeData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(MASTERLIST_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(MASTERLIST_TAB);
      sheet.appendRow(['Employee Email', 'Employee Number', 'Schedule', 'Restday', 'Name']); 
    }
    const data = sheet.getDataRange().getValues();
    const props = PropertiesService.getScriptProperties();
    const agentImages = JSON.parse(props.getProperty('agentImages') || '{}');

    const employees = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        let email = data[i][0].toString();
        let name = data[i][4] ? data[i][4].toString() : email.split('@')[0];
        employees.push({ 
          email: email, empNumber: data[i][1].toString(),
          schedule: data[i][2] ? data[i][2].toString() : 'N/A',
          restday: data[i][3] ? data[i][3].toString() : 'N/A',
          name: name, imageUrl: agentImages[email] || 'https://via.placeholder.com/150?text=No+Image'
        });
      }
    }
    return employees;
  } catch (error) { throw new Error("Backend Error: " + error.message); }
}

function logAudit(action, details) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let auditSheet = ss.getSheetByName(AUDIT_TAB);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(AUDIT_TAB);
    auditSheet.appendRow(['Timestamp', 'User Email', 'Action', 'Details']);
  }
  auditSheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]);
}

function saveAnnouncement(imageUrl, caption) {
  if (Session.getActiveUser().getEmail().toLowerCase() !== ADMIN_EMAIL.toLowerCase()) throw new Error("Unauthorized.");
  const props = PropertiesService.getScriptProperties();
  let arr = JSON.parse(props.getProperty('announcements') || '[]');
  arr.push({ url: imageUrl, caption: caption || '' });
  props.setProperty('announcements', JSON.stringify(arr));
  logAudit("Added Announcement", `Caption: ${caption}`);
  return arr;
}

function deleteAnnouncement(index) {
  if (Session.getActiveUser().getEmail().toLowerCase() !== ADMIN_EMAIL.toLowerCase()) throw new Error("Unauthorized.");
  const props = PropertiesService.getScriptProperties();
  let arr = JSON.parse(props.getProperty('announcements') || '[]');
  if (index >= 0 && index < arr.length) {
    arr.splice(index, 1);
    props.setProperty('announcements', JSON.stringify(arr));
    logAudit("Deleted Announcement", `Index: ${index}`);
    return "Announcement deleted successfully.";
  }
  throw new Error("Invalid index.");
}

function saveAgentImage(email, imageUrl) {
  if (Session.getActiveUser().getEmail().toLowerCase() !== ADMIN_EMAIL.toLowerCase()) throw new Error("Unauthorized.");
  if (!imageUrl) throw new Error("URL cannot be empty.");
  const props = PropertiesService.getScriptProperties();
  let agentImages = JSON.parse(props.getProperty('agentImages') || '{}');
  agentImages[email] = imageUrl;
  props.setProperty('agentImages', JSON.stringify(agentImages));
  logAudit("Updated Agent Image", `Email: ${email}`);
  return "Agent image updated successfully!";
}

function saveBackgroundUrl(url) {
  if (Session.getActiveUser().getEmail().toLowerCase() !== ADMIN_EMAIL.toLowerCase()) throw new Error("Unauthorized.");
  if (!url) throw new Error("URL cannot be empty.");
  PropertiesService.getScriptProperties().setProperty('bgUrl', url);
  logAudit("Updated Background URL", `URL: ${url}`);
  return "Background URL saved successfully.";
}

function saveAdminSettings(blur, opacity) {
  if (Session.getActiveUser().getEmail().toLowerCase() !== ADMIN_EMAIL.toLowerCase()) throw new Error("Unauthorized.");
  const props = PropertiesService.getScriptProperties();
  props.setProperty('bgBlur', blur);
  props.setProperty('bgOpacity', opacity);
  logAudit("Updated Admin Visual Settings", `Blur: ${blur}, Opacity: ${opacity}`);
  return "Settings saved successfully.";
}

// ----------------------------------------------------
// CORE COACHING
// ----------------------------------------------------
function processCoaching(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_TAB);
    const masterData = ss.getSheetByName(MASTERLIST_TAB).getDataRange().getValues();
    
    const activeEmail = Session.getActiveUser().getEmail();
    const activeName = activeEmail.replace('@iopex.com', '').replace(/\./g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    
    let agentEmail = "";
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][1].toString().trim() === data.empNumber.toString().trim()) {
        agentEmail = masterData[i][0].toString().trim();
        break;
      }
    }
    
    let attachmentUrl = "";
    if (data.fileContent && data.fileName) {
      let folder;
      const folders = DriveApp.getFoldersByName("Newsbreak Coaching Attachments");
      if (folders.hasNext()) folder = folders.next();
      else folder = DriveApp.createFolder("Newsbreak Coaching Attachments");
      const blob = Utilities.newBlob(Utilities.base64Decode(data.fileContent), data.fileMimeType, data.fileName);
      attachmentUrl = folder.createFile(blob).getUrl();
    }

    const id = Utilities.getUuid();
    data.rca3 = data.rca; 
    
    const threadId = sendCoachingEmail(data, agentEmail, activeEmail, activeName, id, attachmentUrl);
    
    try {
      const cStart = new Date(data.coachingDateTime);
      const cEnd = new Date(cStart.getTime() + 30 * 60000); 
      CalendarApp.getDefaultCalendar().createEvent(`Coaching Session: ${data.behavior}`, cStart, cEnd, {
          guests: `${agentEmail},${data.managerEmail}`, sendInvites: true,
          description: `Initial coaching log session for ${data.behavior}.\nModerator: ${activeName}\nAgent: ${agentEmail}\nAttachment: ${attachmentUrl}`
      });
      const fStart = new Date(data.timeline);
      const fEnd = new Date(fStart.getTime() + 30 * 60000); 
      CalendarApp.getDefaultCalendar().createEvent(`Coaching Follow-Up: ${data.behavior}`, fStart, fEnd, {
          guests: `${agentEmail},${data.managerEmail}`, sendInvites: true,
          description: `Scheduled follow-up review for ${data.behavior}.\nModerator: ${activeName}\nAgent: ${agentEmail}`
      });
    } catch(calErr) {}
    
    logSheet.appendRow([
      id, new Date(), agentEmail, activeEmail, data.managerEmail, data.behavior, 
      "", "", data.rca, data.baseline, data.target, data.actionPlan, 
      data.support, data.timeline, "Pending Agent Review", threadId, data.coachingDateTime, "", attachmentUrl
    ]);
    
    logAudit("Coaching Logged", `ID: ${id}, Agent: ${agentEmail}, Behavior: ${data.behavior}`);
    return "Action Plan Logged & Emails Sent!";
  } catch (error) { throw new Error("Failed to process: " + error.message); }
}

function sendCoachingEmail(data, agentEmail, moderatorEmail, moderatorName, id, attachmentUrl) {
  const subject = `Action Required: Newsbreak Coaching Assessment - ${data.behavior}`;
  const webAppUrl = ScriptApp.getService().getUrl(); 
  const ackUrl = `${webAppUrl}?action=review&id=${id}`;
  
  const cDateRaw = new Date(data.coachingDateTime);
  const formattedCDate = `${cDateRaw.toLocaleDateString()} at ${cDateRaw.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
  let attachmentHtml = attachmentUrl ? `<tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6;"><b>Attachment</b></td><td style="padding: 10px; border: 1px solid #ffcccc;"><a href="${attachmentUrl}" target="_blank">View Evidence File</a></td></tr>` : '';

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; background-color: #fdfdfd; border: 1px solid #ffcccc; border-radius: 10px; max-width: 600px;">
      <h2 style="color: #d93838; text-align: center;">Newsbreak Coaching & RCA Details</h2>
      <p>Hello,</p>
      <p>A new coaching session has been documented by <b>${moderatorName}</b>.</p>
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
        <tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6; width: 35%;"><b>Coaching Took Place</b></td><td style="padding: 10px; border: 1px solid #ffcccc;">${formattedCDate}</td></tr>
        <tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6;"><b>Behavior/Opportunity</b></td><td style="padding: 10px; border: 1px solid #ffcccc;">${data.behavior}</td></tr>
        <tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6;"><b>RCA (Root Cause)</b></td><td style="padding: 10px; border: 1px solid #ffcccc;">${data.rca3}</td></tr>
        <tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6;"><b>Target Goal</b></td><td style="padding: 10px; border: 1px solid #ffcccc;">${data.target}</td></tr>
        <tr><td style="padding: 10px; border: 1px solid #ffcccc; background:#ffe6e6;"><b>Action Plan</b></td><td style="padding: 10px; border: 1px solid #ffcccc;">${data.actionPlan}</td></tr>
        ${attachmentHtml}
      </table>
      <div style="text-align: center; margin-top: 30px;">
        <a href="${ackUrl}" style="background-color: #d93838; color: #ffffff; padding: 14px 28px; text-decoration: none; font-weight: bold; border-radius: 8px; display: inline-block;">Review & Sign Coaching Plan</a>
      </div>
    </div>
  `;
  const draft = GmailApp.createDraft(agentEmail, subject, "Please view this email in an HTML compatible client.", {
    htmlBody: htmlBody, cc: `${moderatorEmail},${data.managerEmail}`, name: "Newsbreak Coach"
  });
  return draft.send().getThread().getId();
}

// ----------------------------------------------------
// SCORECARD COACHING
// ----------------------------------------------------
function processScorecard(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_TAB);
    
    const activeEmail = Session.getActiveUser().getEmail();
    const activeName = activeEmail.replace('@iopex.com', '').replace(/\./g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    const id = Utilities.getUuid();
    
    const subject = `${data.agentName} - ${data.monthYear} End of Month Performance`;
    const webAppUrl = ScriptApp.getService().getUrl(); 
    const ackUrl = `${webAppUrl}?action=review&id=${id}`;

    const htmlBody = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; background: #f8fafc;">
        <div style="max-width: 600px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
          <div style="background: linear-gradient(135deg, #0f172a 0%, #334155 100%); padding: 30px; text-align: center; color: white;">
            <h2 style="margin: 0; font-size: 24px; font-weight: 600;">Performance Scorecard</h2>
            <p style="margin: 10px 0 0 0; color: #cbd5e1;">${data.monthYear}</p>
          </div>
          <div style="padding: 30px;">
            <p style="color: #475569; font-size: 16px;">Hello <b>${data.agentName}</b>,</p>
            <p style="color: #475569;">Your performance metrics for this period have been finalized. Please review your achievements below.</p>
            
            <div style="background: #f1f5f9; border-radius: 8px; padding: 20px; margin: 25px 0;">
              <table style="width: 100%; border-collapse: collapse;">
                <tr><td style="padding: 12px 0; border-bottom: 1px solid #cbd5e1; color: #64748b; font-weight: 600;">Volume Handled</td><td style="text-align: right; font-weight: bold; color: #0f172a;">${data.prod}</td></tr>
                <tr><td style="padding: 12px 0; border-bottom: 1px solid #cbd5e1; color: #64748b; font-weight: 600;">Accuracy</td><td style="text-align: right; font-weight: bold; color: #10b981;">${data.acc}%</td></tr>
                <tr><td style="padding: 12px 0; color: #64748b; font-weight: 600;">Avg Review Time</td><td style="text-align: right; font-weight: bold; color: #0f172a;">${data.revTime}</td></tr>
              </table>
            </div>
            
            <div style="background: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px; margin-bottom: 25px;">
              <h4 style="margin: 0 0 10px 0; color: #b45309;">Moderator Notes</h4>
              <p style="margin: 0; color: #78350f; font-size: 14px;">${data.notes || 'Keep up the good work!'}</p>
            </div>
            
            <div style="text-align: center; margin-top: 35px;">
              <a href="${ackUrl}" style="background-color: #0f172a; color: #ffffff; padding: 14px 32px; text-decoration: none; font-weight: 600; border-radius: 6px; display: inline-block;">Acknowledge Scorecard</a>
            </div>
          </div>
        </div>
      </div>
    `;

    const draft = GmailApp.createDraft(data.agentEmail, subject, "Please view this email in an HTML compatible client.", {
      htmlBody: htmlBody, cc: `${activeEmail},${data.managerEmail}`, name: "Newsbreak Scorecard"
    });
    const threadId = draft.send().getThread().getId();

    const compiledActionPlan = `Scorecard Logged for ${data.monthYear}.\nAcc: ${data.acc}%, Vol: ${data.prod}, AHT: ${data.revTime}\nNotes: ${data.notes}`;

    logSheet.appendRow([
      id, new Date(), data.agentEmail, activeEmail, data.managerEmail, "Scorecard Coaching", 
      "", "", "N/A", "N/A", "N/A", compiledActionPlan, 
      "N/A", "", "Pending Agent Review", threadId, "", "", ""
    ]);
    
    logAudit("Scorecard Dispatched", `Agent: ${data.agentEmail}, Month: ${data.monthYear}`);
    return "Scorecard Sent Successfully!";
  } catch (error) { throw new Error("Failed to process scorecard: " + error.message); }
}

// ----------------------------------------------------
// GOAL SETTING
// ----------------------------------------------------
function processGoalSetting(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_TAB);
    
    const activeEmail = Session.getActiveUser().getEmail();
    const activeName = activeEmail.replace('@iopex.com', '').replace(/\./g, ' ').replace(/\b\w/g, c => c.toUpperCase());
    const id = Utilities.getUuid();
    
    const subject = `Target Goals Set: ${data.agentName} - ${data.monthYear}`;
    const webAppUrl = ScriptApp.getService().getUrl(); 
    const ackUrl = `${webAppUrl}?action=review&id=${id}`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; background-color: #fdfdfd; border: 1px solid #cce5ff; border-radius: 10px; max-width: 600px;">
        <h2 style="color: #004085; text-align: center;">Newsbreak Goal Setting: ${data.monthYear}</h2>
        <p>Hello <b>${data.agentName}</b>,</p>
        <p>Your performance targets for the upcoming period have been defined by <b>${activeName}</b>.</p>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff; width: 40%;"><b>Accuracy Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.accTarget}</td></tr>
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff;"><b>Productivity Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.prodTarget}</td></tr>
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff;"><b>Avg Review Time Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.revTarget}</td></tr>
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff;"><b>Attendance Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.attTarget}</td></tr>
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff;"><b>PKT Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.pktTarget}</td></tr>
          <tr><td style="padding: 10px; border: 1px solid #b8daff; background:#cce5ff;"><b>Client Escalation Target</b></td><td style="padding: 10px; border: 1px solid #b8daff;">${data.escTarget}</td></tr>
        </table>
        <div style="background: #e2e3e5; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
          <h4 style="margin: 0 0 5px 0;">Additional Notes & Other Metrics</h4>
          <p style="margin: 0; font-size: 14px;">${data.notes}</p>
        </div>
        <div style="text-align: center; margin-top: 30px;">
          <a href="${ackUrl}" style="background-color: #0069d9; color: #ffffff; padding: 14px 28px; text-decoration: none; font-weight: bold; border-radius: 8px; display: inline-block;">Acknowledge Goals</a>
        </div>
      </div>
    `;

    const draft = GmailApp.createDraft(data.agentEmail, subject, "Please view this email in an HTML compatible client.", {
      htmlBody: htmlBody, cc: `${activeEmail},${data.managerEmail}`, name: "Newsbreak Target Setting"
    });
    const threadId = draft.send().getThread().getId();

    const compiledActionPlan = `Goal Setting: ${data.monthYear}.\nAcc: ${data.accTarget}, Prod: ${data.prodTarget}, AHT: ${data.revTarget}, Att: ${data.attTarget}, PKT: ${data.pktTarget}, Esc: ${data.escTarget}\nNotes: ${data.notes}`;

    logSheet.appendRow([
      id, new Date(), data.agentEmail, activeEmail, data.managerEmail, "Goal Setting", 
      "", "", "N/A", "N/A", "N/A", compiledActionPlan, 
      "N/A", "", "Pending Agent Review", threadId, "", "", ""
    ]);
    
    logAudit("Goal Setting Dispatched", `Agent: ${data.agentEmail}, Month: ${data.monthYear}`);
    return "Goals Set and Dispatched Successfully!";
  } catch (error) { throw new Error("Failed to process goals: " + error.message); }
}

function getAgentHistory(empNumber) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const masterData = ss.getSheetByName(MASTERLIST_TAB).getDataRange().getValues();
  let targetEmail = "";
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][1].toString().trim() === empNumber.toString().trim()) {
      targetEmail = masterData[i][0].toString().trim();
      break;
    }
  }
  if (!targetEmail) return [];
  const data = ss.getSheetByName(LOGS_TAB).getDataRange().getValues();
  const history = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() === targetEmail) {
      history.push({
        date: Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "MM/dd/yyyy"),
        mod: data[i][3], 
        behavior: data[i][5], 
        status: data[i][14] 
      });
    }
  }
  return history;
}

function getComplianceReport() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName(LOGS_TAB);
    const masterSheet = ss.getSheetByName(MASTERLIST_TAB);
    if (!logSheet || !masterSheet) return { compliant: [], nonCompliant: [], uncoachedAgents: [] };
    
    const logData = logSheet.getDataRange().getValues();
    const masterData = masterSheet.getDataRange().getValues();
    const now = new Date();
    const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
    
    const allModerators = new Set();
    const currentWeekCounts = {};
    const coachedAgents = new Set();

    for (let i = 1; i < logData.length; i++) {
      const timestamp = new Date(logData[i][1]);
      const agentEmail = logData[i][2];
      const modEmail = logData[i][3]; 
      
      if (!modEmail) continue;
      allModerators.add(modEmail);
      
      if (timestamp >= startOfWeek) {
        currentWeekCounts[modEmail] = (currentWeekCounts[modEmail] || 0) + 1;
        coachedAgents.add(agentEmail);
      }
    }
    
    const compliant = [];
    const nonCompliant = [];
    allModerators.forEach(mod => {
      const count = currentWeekCounts[mod] || 0;
      if (count >= 8) compliant.push({ name: mod, count: count });
      else nonCompliant.push({ name: mod, count: count });
    });
    
    const uncoachedAgents = [];
    for (let i = 1; i < masterData.length; i++) {
      const email = masterData[i][0];
      const name = masterData[i][4] || email;
      if (email && !coachedAgents.has(email)) {
        uncoachedAgents.push(name);
      }
    }
    return { compliant, nonCompliant, uncoachedAgents };
  } catch (error) { throw new Error("Failed to load report."); }
}

function getCoachingDetails(id) {
  const data = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOGS_TAB).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return {
        date: Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "MMM dd, yyyy"),
        moderator: data[i][3],
        behavior: data[i][5],
        actionPlan: data[i][11],
        status: data[i][14]
      };
    }
  }
  throw new Error("Coaching record not found.");
}

function processAgentReview(id, status, comments) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOGS_TAB);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (data[i][14] === "Acknowledged" || data[i][14] === "Disputed") throw new Error("This coaching plan has already been processed.");
      sheet.getRange(i + 1, 15).setValue(status);
      sheet.getRange(i + 1, 18).setValue(comments); 
      logAudit("Agent Reviewed Coaching", `ID: ${id}, Status: ${status}, Comments: ${comments}`);
      const threadId = data[i][15]; 
      if (threadId) {
        try {
          const color = status === "Acknowledged" ? "#4CAF50" : "#d93838";
          const icon = status === "Acknowledged" ? "✅" : "⚠️";
          GmailApp.getThreadById(threadId).reply(`System Notification: The agent has responded to the action plan. Status: ${status}`, {
            htmlBody: `<div style='padding: 10px; border-left: 4px solid ${color}; background-color: #f9f9f9;'><p>${icon} <em>System Auto-Reply:</em> The agent has marked this coaching as <b>${status}</b>.</p><p><b>Agent Comments:</b> ${comments}</p></div>`
          });
        } catch(e) { console.error(e); }
      }
      return `Successfully marked as ${status}.`;
    }
  }
  throw new Error("Record not found.");
}
