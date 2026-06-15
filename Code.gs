/***** ================= Insights Genie – Full Restore + Resilience Logic ================= *****/

function doGet(e) {
  if (e && e.parameter && e.parameter.view) {
    try {
      const fileId = e.parameter.view;
      const file = DriveApp.getFileById(fileId);
      const htmlContent = file.getBlob().getDataAsString();

      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle('Insights Genie Infographic')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput(
        "<h3>Error: Infographic not found or access denied.</h3>" +
        "<p>Please ensure the file exists and you have authorized Drive access.</p>"
      );
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Insights Genie')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** ================= CONFIG ================= *****/
const GEMINI_KEY_PROP = 'GEMINI_API_KEY';
const GEMINI_MODEL = 'gemini-3.1-pro-preview';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const DRIVE_FOLDER_ID = '1bB_2w_Dv36bIN3Kxr7WFNOcocicSz5A9';
const INFO_URL_PROP = 'Insights_Geine_LAST_INFO_URL';
const LOGGING_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZfsNLTHFNitPQLxU1s0PHlR3y6z_DPtluldcZTEjyeU/edit';

/***** ================= CORE: Generate AI Insights (ROUTER) ================= *****/
function generateInsightsFromSheet(sheetUrl, targetaudience) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    let prompt = "";

    if (targetaudience === "Smart Phones") {
      if (typeof getSmartPhonesPrompt !== "function") throw new Error("Missing Prompt_SP.gs file.");
      prompt = getSmartPhonesPrompt(sheetId, targetaudience);
    } 
    else if (targetaudience === "EdTech") {
      if (typeof getEdTechPrompt !== "function") throw new Error("Missing Prompt_EdTech.gs file.");
      prompt = getEdTechPrompt(sheetId, targetaudience);
    } 
    else {
      throw new Error(`Specific routing for "${targetaudience}" is not configured. Please use the CUSTOM option.`);
    }

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= GEMINI TEXT ================= *****/
function callGeminiText_(prompt) {
  const key = PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROP);
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }]}],
    generationConfig: { temperature: 0.3, maxOutputTokens: 60000 }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 60000
  });
  const parsed = JSON.parse(resp.getContentText() || '{}');
  return parsed.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

/***** ================= GENERATE INFOGRAPHIC HTML (100% POLLINATIONS PIPELINE) ================= *****/
function generateCustomInfographicHtmlFromInsights(targetaudience, reportText) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const key = scriptProps.getProperty(GEMINI_KEY_PROP);
    const pollinationsKey = scriptProps.getProperty('POLLINATIONS_API_KEY'); 

    if (!pollinationsKey) {
       throw new Error("Missing POLLINATIONS_API_KEY. Please add your Pollinations API Key to Script Properties to troubleshoot image loading.");
    }

    const jsonPrompt = getImagenPromptsJSON(targetaudience, reportText);
    const textUrl = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${key}`;
    const jsonPayload = {
      contents: [{ role: 'user', parts: [{ text: jsonPrompt }]}],
      generationConfig: { temperature: 0.2, maxOutputTokens: 60000 }
    };
    
    const jsonResp = UrlFetchApp.fetch(textUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(jsonPayload), muteHttpExceptions: true
    });
    
    const parsedJsonResp = JSON.parse(jsonResp.getContentText() || '{}');
    if (parsedJsonResp.error) {
      throw new Error(`API Error (Prompt Generation): ${parsedJsonResp.error.message}`);
    }
    
    let promptArray = [];
    try {
      let rawJson = parsedJsonResp.candidates[0].content.parts[0].text;
      rawJson = rawJson.replace(/```json/gi, '').replace(/```/g, '').trim();
      promptArray = JSON.parse(rawJson);
    } catch (e) {
      throw new Error("Failed to parse JSON image prompts. Please check the AI's JSON output format.");
    }

    if (promptArray.length !== 12) {
       promptArray = promptArray.slice(0, 12); 
    }

    let finalImageUrls = promptArray.map(prompt => {
      const safePrompt = encodeURIComponent(prompt);
      return `https://gen.pollinations.ai/image/${safePrompt}?width=800&height=600&nologo=true&model=flux&key=${pollinationsKey}`;
    });

    const finalHtmlPrompt = getCustomInfographicPrompt(targetaudience, reportText);
    
    const htmlPayload = {
      contents: [{ role: 'user', parts: [{ text: finalHtmlPrompt }]}],
      generationConfig: { temperature: 0.2, maxOutputTokens: 60000 }
    };

    const htmlResp = UrlFetchApp.fetch(textUrl, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(htmlPayload), muteHttpExceptions: true, timeout: 60000
    });

    const parsedHtmlResp = JSON.parse(htmlResp.getContentText() || '{}');
    if (parsedHtmlResp.error) {
      throw new Error(`API Error (HTML Generation): ${parsedHtmlResp.error.message}`);
    }

    let finalHtml = parsedHtmlResp.candidates[0].content.parts[0].text;
    finalHtml = finalHtml.replace(/```html/gi, '').replace(/```/g, '').trim();

    for (let i = 0; i < finalImageUrls.length; i++) {
       const placeholderToFind = `IMG_PLACEHOLDER_${i + 1}`;
       finalHtml = finalHtml.replace(placeholderToFind, finalImageUrls[i]);
    }

    const hostingResult = saveHtmlAndGetUrl_(targetaudience, finalHtml);
    setInfographicURL(hostingResult.url);

    // Context Clean: Dummy auto-publish pipeline has been stripped from here to rely on live user frontend interactions.

    return { ok: true, html: finalHtml, url: hostingResult.url };

  } catch (err) {
    return { ok: false, error: "Backend Pipeline Error: " + err.message };
  }
}

/***** ================= HELPER: HOST ON DRIVE ================= *****/
function saveHtmlAndGetUrl_(label, htmlBody) {
  const fullHtml = `<!DOCTYPE html><html><head><script src="https://cdn.tailwindcss.com"></script></head><body style="background:#f8fafc; padding:24px;"><div style="max-width:1152px; margin:0 auto; background:white; border-radius:16px; padding:32px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);">${htmlBody}</div></body></html>`;
  
  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  
  const fileName = `Infographic-${label}-${Date.now()}.html`;
  const file = folder.createFile(fileName, fullHtml, MimeType.HTML);
  
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      console.warn("Could not set sharing permissions. File remains private to the creator.");
    }
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  const finalUrl = scriptUrl ? `${scriptUrl}?view=${file.getId()}` : file.getUrl();

  return { url: finalUrl, fileId: file.getId() };
}

/***** ================= PUBLISH & EMAIL LOGIC ================= *****/
function publishReportAndLog(audience, reportText, emailsCsv, person, client, rev, sheetUrl, rowIndex, isFinalSend) {
  try {
    const recipients = (emailsCsv || '').split(',').map(x => x.trim()).filter(Boolean);
    const userProps = PropertiesService.getUserProperties();
    const infoUrl = userProps.getProperty(INFO_URL_PROP) || '';
    
    let doc;
    let logSheet;

    if (LOGGING_SHEET_URL) {
      logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
    }

    // Reuse pre-existing document link within this workflow iteration block if index matches
    if (rowIndex && rowIndex > 0 && logSheet) {
      const targetDocUrl = logSheet.getRange(rowIndex, 5).getValue();
      if (targetDocUrl && targetDocUrl.includes("docs.google.com/document")) {
        try {
          doc = DocumentApp.openByUrl(targetDocUrl);
          doc.getBody().clear();
        } catch(docErr) {
          doc = null;
        }
      }
    }

    if (!doc) {
      doc = DocumentApp.create(`Insights Genie - ${audience} - ${new Date().toDateString()}`);
    }

    const body = doc.getBody();
    body.appendParagraph(`AI Strategy Insights — ${audience}`);
    body.appendParagraph(`Strategist: ${person} | Client: ${client}`).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
    body.appendParagraph(reportText || '');

    if (infoUrl && infoUrl.startsWith('http')) {
      body.appendParagraph('Infographic').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const p = body.appendParagraph('📊 Open Infographic (Interactive)');
      p.setLinkUrl(infoUrl);
      p.setBold(true);
      p.setFontSize(13);
      p.setForegroundColor('#1a73e8');
    }
    doc.saveAndClose();

    const file = DriveApp.getFileById(doc.getId());
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    recipients.forEach(r => { try { file.addViewer(r); } catch (err) {} });

    // Synchronization logic implementation to ensure matching index logic updates
    try {
      if (logSheet) {
        if (rowIndex && rowIndex > 0) {
          logSheet.getRange(rowIndex, 4).setValue(audience);
          logSheet.getRange(rowIndex, 5).setValue(doc.getUrl());
          logSheet.getRange(rowIndex, 6).setValue(infoUrl);
          logSheet.getRange(rowIndex, 7).setValue(recipients.join(', '));
          logSheet.getRange(rowIndex, 9).setValue(rev); // Ensure Revenue is updated in Column I (9th Column)
        } else {
          logSheet.appendRow([
            new Date(), 
            person, 
            client, 
            audience, 
            doc.getUrl(), 
            infoUrl, 
            recipients.join(', '), 
            Session.getActiveUser().getEmail(),
            rev // Append Revenue as the 9th item (Column I)
          ]);
          rowIndex = logSheet.getLastRow();
        }
      }
    } catch (logErr) {
      console.log("Logging framework error safely bypassed: " + logErr);
    }

    // FIX: Default isFinalSend to true if it wasn't explicitly passed from the frontend
    if (isFinalSend === undefined) {
      isFinalSend = true;
    }

    // Only fire off automated customer notification emails on explicit Final button submissions
    if (isFinalSend && recipients.length > 0) {
      const emailSubject = `Strategic Insights Report: ${audience} | Prepared for ${client}`;
      
      const emailBody = `Dear ${person},\n\n` +
        `Please find the strategic AI insights report for ${audience}, prepared specifically for ${client}.\n\n` +
        `This document contains a comprehensive analysis of the latest search trends, competitive dynamics, and actionable marketing recommendations based on consumer intent.\n\n` +
        `📄 Access the Strategy Document here: ${doc.getUrl()}\n` +
        (sheetUrl ? `📊 Access the Source Data Sheet here: ${sheetUrl}\n\n` : `\n`) +
        `Best regards,\n` +
        `Team Genie`;

      MailApp.sendEmail({
        to: recipients.join(','),
        subject: emailSubject,
        body: emailBody,
        name: "Insights Genie",
        noReply: true
      });
    }

    return { ok: true, docUrl: doc.getUrl(), url: infoUrl, rowIndex: rowIndex };
  } catch (err) {
    return { ok: false, error: "Publishing Error: " + err.message };
  }
}

/***** ================= UTILS ================= *****/
function extractSheetId_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

function setInfographicHTML(html) { return true; }
function setInfographicURL(url) { 
  PropertiesService.getUserProperties().setProperty(INFO_URL_PROP, url);
  return true; 
}

/***** ================= USER CONTEXT ================= *****/
function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail();
  let fullName = "";
  let photoUrl = "";

  try {
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    if (user && user.name) {
      fullName = user.name.fullName || user.name.givenName;
    }
  } catch (e) {
    console.log("Admin Directory lookup failed: " + e.message);
  }

  if (!fullName) {
    const handle = email.split('@')[0];
    fullName = handle.split('.').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
  }

  return { email: email, name: fullName, photoUrl: photoUrl };
}

function authorizeScript() {
  DriveApp.getRootFolder(); 
  GmailApp.getInboxThreads(0, 1);
  if (false) { MailApp.sendEmail("test@google.com", "test", "test"); }

  const tempDoc = DocumentApp.create("Insights Genie Permission Test");
  tempDoc.saveAndClose();
  DriveApp.getFileById(tempDoc.getId()).setTrashed(true);
}

/***** ================= CORE: Generate Custom Miscellaneous Insights ================= *****/
function generateMiscellaneousInsights(sheetUrl, targetaudience, customPrompt) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    let miscSheet = ss.getSheetByName("Miscellaneous") || ss.getSheetByName("Misllenious");
    if (!miscSheet) throw new Error('Could not find a tab named "Miscellaneous" in the provided sheet.');

    const data = miscSheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the Miscellaneous tab.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    const prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nINSTRUCTIONS: ${customPrompt}\nDATA:\n${csvString}`;
    return { ok: true, reportText: callGeminiText_(prompt) };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/***** ================= CORE: Generate Custom Insights ================= *****/
function generateCustomInsightsFromSheet(sheetUrl, targetaudience) {
  try {
    const sheetId = extractSheetId_(sheetUrl);
    if (!sheetId) throw new Error('Invalid Custom Sheet URL');

    const ss = SpreadsheetApp.openById(sheetId);
    
    const sheet = ss.getSheets()[0]; 
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error('No data found in the custom sheet.');

    const csvString = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

    let prompt;
    if (typeof getCustomPrompt === "function") {
        prompt = getCustomPrompt(targetaudience, csvString);
    } else {
        prompt = `ACT AS A SENIOR DATA ANALYST AND MARKET RESEARCHER.\nTOPIC: ${targetaudience}\nAnalyze the following data and provide a highly structured, strategic insight report.\nDATA:\n${csvString}`;
    }

    const text = callGeminiText_(prompt);
    return { ok: true, reportText: text };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
/***** ================= EARLY LOGGING (CREDENTIALS) ================= *****/
// Note: Added the 'rev' parameter so revenue is captured correctly right from the start.
function logProjectCredentials(person, client, audience, rev) {
  try {
    if (!LOGGING_SHEET_URL) return { ok: false, error: "No logging URL" };
    
    const logSheet = SpreadsheetApp.openByUrl(LOGGING_SHEET_URL).getSheets()[0];
    const email = Session.getActiveUser().getEmail();
    
    logSheet.appendRow([
      new Date(),
      person,
      client,
      audience,
      "Pending Publish...",
      "Pending...",
      "",
      email,
      rev || "" // Append Revenue as the 9th item (Column I) early on
    ]);
    
    return { ok: true, rowIndex: logSheet.getLastRow() };
  } catch (err) {
    console.log("Early logging failed: " + err);
    return { ok: false, error: err.message };
  }
}
