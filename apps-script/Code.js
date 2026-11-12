/**
 * Resume Evolver: AI Resume Tailoring System
 * Google Apps Script Backend (Code.js)
 * 
 * Live inside Google Apps Script as a Web App.
 * Uses OpenRouter to tailor resume data from Google Sheets into a Google Docs template.
 */

// ==========================================
// 1. TEMPLATE PLACEHOLDER TAGS CONFIGURATION
// ==========================================
const TEMPLATE_TAGS = {
  SUMMARY: '{{SUMMARY}}',
  SKILLS: '{{SKILLS}}',
  COMPANY: '{{COMPANY}}',
  JOB_TITLE: '{{JOB_TITLE}}',
  
  SKILL_GROUP_TITLE: '{{SKILL_GROUP_%d_TITLE}}',
  SKILL_GROUP_LIST: '{{SKILL_GROUP_%d_SKILLS}}',
  
  EXP_ROLE: '{{EXP_%d_ROLE}}',
  EXP_COMPANY: '{{EXP_%d_COMPANY}}',
  EXP_DATES: '{{EXP_%d_DATES}}',
  EXP_LOCATION: '{{EXP_%d_LOCATION}}',
  EXP_BULLET: '{{EXP_%d_BULLET_%b}}',
  
  PROJECT_TITLE: '{{PROJECT_%d_TITLE}}',
  PROJECT_TECH: '{{PROJECT_%d_TECH}}',
  PROJECT_LINK: '{{PROJECT_%d_LINK}}',
  PROJECT_BULLET: '{{PROJECT_%d_BULLET_%b}}',
  
  CERT_TITLE: '{{CERT_%d_TITLE}}',
  CERT_ORG: '{{CERT_%d_ORG}}',
  CERT_DATE: '{{CERT_%d_DATE}}',
  
  POR_ROLE: '{{POR_%d_ROLE}}',
  POR_ORG: '{{POR_%d_ORG}}',
  POR_DATES: '{{POR_%d_DATES}}',
  POR_BULLET: '{{POR_%d_BULLET_%b}}'
};

// ==========================================
// 2. CONFIGURATION & CREDENTIALS
// ==========================================
const CONFIG = {
  OPENROUTER_API_KEY: 'YOUR_KEY',
  OPENROUTER_MODEL: 'google/gemini-2.5-flash',
  
  SPREADSHEET_ID: 'YOUR_DATABASE_SPREADSHEET_ID',
  TEMPLATE_DOC_ID: 'YOUR_TEMPLATE_DOCUMENT_ID',
  OUTPUT_FOLDER_ID: 'YOUR_OUTPUT_FOLDER_ID',
  
  PROMPT_DOC_ID: 'YOUR_PROMPT_DOCUMENT_ID',
  CAREER_DB_SHEET_NAME: 'Database',
  
  MAX_SKILL_GROUPS: 8,
  MAX_EXPERIENCES: 3,
  MAX_EXPERIENCE_BULLETS: 3,
  MAX_PROJECTS: 3,
  MAX_PROJECT_BULLETS: 3,
  MAX_CERTIFICATIONS: 4,
  MAX_RESPONSIBILITIES: 3,
  MAX_POR_BULLETS: 2,
};

// ==========================================
// ENTRY POINT (WEB APP GET/POST ROUTING)
// ==========================================

/**
 * Serves the Index.html loading template when accessed via GET.
 * Reads the 'd' query parameter and injects it into the template.
 */
function doGet(e) {
  const payload = (e && e.parameter && e.parameter.d) ? e.parameter.d : '';
  
  return renderIndex(payload);
}

function renderIndex(payload) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.payload = payload || '';
  
  return template.evaluate()
    .setTitle('Resume Revolver: Tailoring Resume')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Server-side entry point executed asynchronously from Index.html.
 */
function runTailoringProcessFromServer(jobDescription, company, jobTitle) {
  try {
    if (!jobDescription || jobDescription.trim() === '') {
      throw new Error('Job description is empty or missing.');
    }
    
    // 1. Fetch Career Database from Google Sheets
    const careerDb = readCareerDatabase();
    if (careerDb.length === 0) {
      throw new Error(`Career Database sheet "${CONFIG.CAREER_DB_SHEET_NAME}" is empty or has no rows.`);
    }

    // 2. Fetch System Prompt dynamically from Google Doc (or fallback)
    const systemPrompt = getSystemPrompt();

    // 3. Verify Template Doc and Output Folder access BEFORE consuming AI tokens
    verifyDocAndFolderAccess();
    
    // 4. Query OpenRouter AI
    const aiOutput = validateAiOutput(callOpenRouter(jobDescription, careerDb, systemPrompt), careerDb);

    // 5. Create & Fill Tailored Google Doc
    const docUrl = generateTailoredResume(aiOutput, jobTitle, company);
    
    return docUrl;
  } catch (error) {
    console.error("Server-side tailoring failed: " + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * Fallback Web App POST handler.
 */
function doPost(e) {
  try {
    // Browser form submissions use this route so job data stays in the request body.
    if (e && e.parameter && e.parameter.d) {
      return renderIndex(e.parameter.d);
    }

    let postData;
    if (e && e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No POST request payload received.');
    }
    
    const jobDescription = postData.jobDescription;
    const company = postData.company || '';
    const jobTitle = postData.jobTitle || '';
    
    // 1. Fetch Career Database
    const careerDb = readCareerDatabase();

    // 2. Fetch System Prompt
    const systemPrompt = getSystemPrompt();

    // 3. Verify access before consuming tokens
    verifyDocAndFolderAccess();

    // 4. Query OpenRouter AI
    const aiOutput = validateAiOutput(callOpenRouter(jobDescription, careerDb, systemPrompt), careerDb);

    // 5. Create & Fill Tailored Google Doc
    const docUrl = generateTailoredResume(aiOutput, jobTitle, company);
    
    const result = { success: true, documentUrl: docUrl, error: '' };
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const result = { success: false, documentUrl: '', error: error.toString() };
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function verifyDocAndFolderAccess() {
  const templateDocId = getSetting('TEMPLATE_DOC_ID', CONFIG.TEMPLATE_DOC_ID);
  const outputFolderId = getSetting('OUTPUT_FOLDER_ID', CONFIG.OUTPUT_FOLDER_ID);
  
  if (!templateDocId || templateDocId === 'YOUR_TEMPLATE_DOC_ID_HERE') {
    throw new Error('Template Google Doc ID is not configured.');
  }
  if (!outputFolderId || outputFolderId === 'YOUR_OUTPUT_FOLDER_ID_HERE') {
    throw new Error('Output Google Drive folder ID is not configured.');
  }

  try {
    DriveApp.getFileById(templateDocId);
  } catch (e) {
    throw new Error(`Template Google Doc (ID: "${templateDocId}") cannot be accessed: ${e.toString()}`);
  }

  try {
    DriveApp.getFolderById(outputFolderId);
  } catch (e) {
    throw new Error(`Output Drive Folder (ID: "${outputFolderId}") cannot be accessed: ${e.toString()}`);
  }
}

// ==========================================
// GOOGLE SHEETS HANDLER
// ==========================================

function readCareerDatabase() {
  const ssId = getSetting('SPREADSHEET_ID', CONFIG.SPREADSHEET_ID);
  const sheetName = getSetting('CAREER_DB_SHEET_NAME', CONFIG.CAREER_DB_SHEET_NAME);
  
  if (!ssId || ssId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Google Sheet SPREADSHEET_ID is not configured.');
  }
  
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet tab named "${sheetName}" was not found in spreadsheet ID "${ssId}".`);
  }
  
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  if (values.length <= 1) {
    return [];
  }
  
  const headers = values[0];
  const data = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowData = {};
    let hasData = false;
    
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j].toString().trim();
      const val = row[j];
      if (header) {
        const strVal = val !== null && val !== undefined ? val.toString().trim() : '';
        rowData[header] = strVal;
        if (strVal !== '') {
          hasData = true;
        }
      }
    }
    
    if (hasData) {
      data.push(rowData);
    }
  }
  
  return data;
}

// ==========================================
// AI OUTPUT VALIDATION
// ==========================================

function validateAiOutput(output, careerDb = []) {
  if (!output || typeof output !== 'object' || Array.isArray(output)) {
    throw new Error('AI output must be a JSON object.');
  }

  const asString = value => value === null || value === undefined ? '' : String(value).trim();
  
  // Build lookup map from careerDb for Duplicate Group codes by item name/company/title
  const dbGroupMap = {};
  if (Array.isArray(careerDb)) {
    careerDb.forEach(row => {
      let dupGroup = '';
      for (const key in row) {
        if (key.toLowerCase().replace(/[^a-z0-9]/g, '') === 'duplicategroup') {
          dupGroup = asString(row[key]);
          break;
        }
      }
      if (dupGroup) {
        const name = asString(row.Name || row.name || row.Company || row.company || row.Role || row.role);
        if (name) {
          dbGroupMap[name.toLowerCase()] = dupGroup;
        }
      }
    });
  }

  const getItemDupGroup = (item, primaryKey, secondaryKey) => {
    let group = asString(item.duplicate_group || item['Duplicate Group'] || item['duplicateGroup']);
    if (!group) {
      const keyVal = asString(item[primaryKey] || item[secondaryKey]);
      if (keyVal && dbGroupMap[keyVal.toLowerCase()]) {
        group = dbGroupMap[keyVal.toLowerCase()];
      }
    }
    return group;
  };

  const seenDuplicateGroups = new Set();

  const filterAndLimitObjects = (value, max, primaryKey, secondaryKey) => {
    if (!Array.isArray(value)) return [];
    const validObjects = value.filter(item => item && typeof item === 'object' && !Array.isArray(item));
    const result = [];
    
    for (const item of validObjects) {
      const groupCode = getItemDupGroup(item, primaryKey, secondaryKey);
      if (groupCode !== '') {
        if (seenDuplicateGroups.has(groupCode.toLowerCase())) {
          console.warn(`Filtering out duplicate item in group "${groupCode}": ${item[primaryKey] || item[secondaryKey]}`);
          continue;
        }
        seenDuplicateGroups.add(groupCode.toLowerCase());
      }
      result.push(item);
      if (result.length >= max) break;
    }
    return result;
  };

  const strings = (value, max) => Array.isArray(value)
    ? value.map(asString).filter(Boolean).slice(0, max)
    : [];
  const copyFields = (item, fields) => {
    fields.forEach(field => {
      item[field] = asString(item[field]);
    });
    return item;
  };

  output.summary = asString(output.summary);
  output.skills = asString(output.skills);

  output.skill_groups = (Array.isArray(output.skill_groups)
    ? output.skill_groups.filter(item => item && typeof item === 'object' && !Array.isArray(item)).slice(0, CONFIG.MAX_SKILL_GROUPS)
    : []
  ).map(item => copyFields(item, ['title', 'skills']));

  output.experiences = filterAndLimitObjects(output.experiences, CONFIG.MAX_EXPERIENCES, 'company', 'role').map(item => {
    copyFields(item, ['company', 'role', 'location', 'dates', 'link', 'duplicate_group']);
    item.bullets = strings(item.bullets, CONFIG.MAX_EXPERIENCE_BULLETS);
    return item;
  });

  output.projects = filterAndLimitObjects(output.projects, CONFIG.MAX_PROJECTS, 'title', 'tech_stack').map(item => {
    copyFields(item, ['title', 'tech_stack', 'link', 'duplicate_group']);
    item.bullets = strings(item.bullets, CONFIG.MAX_PROJECT_BULLETS);
    return item;
  });

  output.certifications = filterAndLimitObjects(output.certifications, CONFIG.MAX_CERTIFICATIONS, 'title', 'org').map(item =>
    copyFields(item, ['title', 'org', 'date', 'link', 'duplicate_group'])
  );

  output.responsibilities = filterAndLimitObjects(output.responsibilities, CONFIG.MAX_RESPONSIBILITIES, 'role', 'org').map(item => {
    copyFields(item, ['role', 'org', 'dates', 'link', 'duplicate_group']);
    item.bullets = strings(item.bullets, CONFIG.MAX_POR_BULLETS);
    return item;
  });

  return output;
}

// ==========================================
// OPENROUTER API CLIENT
// ==========================================

const DEFAULT_SYSTEM_PROMPT = `You are a world-class technical recruiter and resume tailoring system. 
Your task is to analyze a job description (JD) and a user's Career Database (containing work experience, projects, skills, etc.).
You will select the most relevant entries and write tailored resume bullet points.

DATABASE SPECIFICATION:
- Column "Item": Contains the record type. Map 'prj' or 'Project' to projects, 'Exp' or 'Work Experience' to work experiences, 'Crt' or 'Certification' to certifications, 'skills' or 'Skill Group' to skills, and 'por' or 'Position of Responsibility' to responsibilities.
- Column "Name": The name of the project, role, certification, skill category, or position.
- Column "Description": Contextual summary of the project/experience.
- Column "Technology used": Software, libraries, tools, and technical stacks (Column D).
- Column "Skills applicable": Broad concepts, design patterns, architectures, and soft skills (Column E).
- Column "Link": URL associated with the item (GitHub, project site, company home page, or credential verification). Always pass this value through if available!
- Column "Role / Context": Your position, team context, or project role (Column G).
- Column "Duplicate Group": Code for mutually exclusive entries (e.g. 'A', 'SEC_ROLE'). If 2 entries have the same code, only select AT MOST ONE that best fits the JD.
- Other columns provide granular details such as Start/End dates, location, key achievements/metrics, and links.

CRITICAL INSTRUCTIONS:
1. FACTUAL ACCURACY: Do NOT invent, assume, or expand upon achievements, impact numbers, credentials, dates, companies, or technologies. You must only use the raw facts provided in the Career Database. You can reword, restructure, and emphasize relevant aspects, but fabricating data is strictly forbidden.
2. ATS OPTIMIZATION: Incorporate key terms, tools, and methodologies from the Job Description into the tailored bullet points.
3. MEASURABLE RESULTS: Keep and highlight metrics (e.g., percentages, scale, dollar values) from the database using active verbs.
4. JSON FORMAT: You MUST return a strict JSON object that conforms EXACTLY to the structure specified. Do not include markdown codeblocks (like \`\`\`json ... \`\`\`) in your output. Return only the raw JSON.
5. ESCAPING QUOTES: If you include quote marks inside any string value (such as job titles, project names, or bullet points), you MUST use single quotes (') instead of double quotes (\") to avoid breaking the JSON format. For example, write "Completed projects: 'Securing a Business Network'" instead of using double quotes inside the string. Never output unescaped double quotes inside a JSON string value.
6. DUPLICATE GROUP MUTUAL EXCLUSIVITY: Column "Duplicate Group" contains group identifiers (e.g. 'A', 'SEC_ROLE'). If multiple entries in the database share the same non-empty Duplicate Group code, they represent mutually exclusive variations of the same role or project. You MUST select AT MOST ONE item from each Duplicate Group code (the one that best matches the Job Description). Never include multiple items with the same Duplicate Group code.
7. BULLET STRUCTURE & LOGICAL BREAKDOWN: For work experiences, projects, and positions of responsibility, you MUST break down raw achievements and metrics cleanly and logically into distinct, separate bullet points up to the requested count (maxExperienceBullets, maxProjectBullets, maxPorBullets). Do NOT cram multiple distinct accomplishments into a single bullet if open bullet slots remain.
8. DISTINCT SKILL GROUPS (NO REPETITION): Each skill group in 'skill_groups' MUST contain unique, distinct skills relevant to its group title (e.g. 'DevOps & Security' should contain security/infrastructure tools like Docker, Nginx, WireGuard, while 'Databases & Vector' should contain database/vector tools like PostgreSQL, Redis, Pinecone). NEVER repeat the exact same list of skills across multiple skill groups.`;

function getSystemPrompt() {
  const promptDocId = getSetting('PROMPT_DOC_ID', CONFIG.PROMPT_DOC_ID);
  if (promptDocId && promptDocId.trim() !== '' && promptDocId !== 'YOUR_PROMPT_DOC_ID_HERE') {
    try {
      const doc = DocumentApp.openById(promptDocId.trim());
      const text = doc.getBody().getText().trim();
      if (text && text.length > 50) {
        return text;
      }
    } catch (e) {
      console.warn("Could not read System Prompt from Google Doc ID (" + promptDocId + "), using default prompt: " + e.toString());
    }
  }
  return DEFAULT_SYSTEM_PROMPT;
}

function callOpenRouter(jobDesc, careerDb, promptOverride) {
  const apiKey = getSetting('OPENROUTER_API_KEY', CONFIG.OPENROUTER_API_KEY);
  const model = getSetting('OPENROUTER_MODEL', CONFIG.OPENROUTER_MODEL);
  
  if (!apiKey || apiKey === 'YOUR_OPENROUTER_API_KEY') {
    throw new Error('OpenRouter API Key is not set.');
  }

  const systemPrompt = promptOverride || getSystemPrompt();

  const userPrompt = {
    jobDescription: jobDesc,
    careerDatabase: careerDb,
    maxSkillGroups: CONFIG.MAX_SKILL_GROUPS,
    maxExperiences: CONFIG.MAX_EXPERIENCES,
    maxProjects: CONFIG.MAX_PROJECTS,
    maxCertifications: CONFIG.MAX_CERTIFICATIONS,
    maxResponsibilities: CONFIG.MAX_RESPONSIBILITIES,
    maxExperienceBullets: CONFIG.MAX_EXPERIENCE_BULLETS,
    maxProjectBullets: CONFIG.MAX_PROJECT_BULLETS,
    maxPorBullets: CONFIG.MAX_POR_BULLETS,
    instructions: "Select the top experiences, projects, certifications, and positions of responsibility matching the JD. For projects, populate 'tech_stack' by prioritizing and combining relevant entries from 'Skills applicable' (Column E), 'Technology used' (Column D), and 'Role / Context' (Column G) that best align with the JD requirements. Cleanly and logically break down candidate achievements into distinct bullet points for all open bullet slots (e.g. up to maxExperienceBullets/maxProjectBullets). Rewrite bullets to align with the JD, and generate a tailored summary. Select and optimize technical skill groups from 'skills' rows. Preserve original 'Link' values. Ensure mutual exclusivity for items sharing the same Duplicate Group code.",
    requiredJsonSchema: {
      summary: "A highly tailored 2-3 sentence professional summary highlighting the core value proposition matching the job requirements.",
      skills: "A single comma-separated string of ALL relevant tools and technologies (fallback general skills).",
      skill_groups: [
        {
          title: "Skill group title (e.g. Languages, DevOps & Security, Databases & Vector, GenAI & ML) mapped from sheet or optimized",
          skills: "Comma-separated list of relevant skills/tools in this group matching the JD (e.g. C++ (proficient), Python, Bash)"
        }
      ],
      experiences: [
        {
          company: "Company name (mapped from Database 'Name' or 'Company')",
          role: "Job title / role (mapped from Database 'Role' or 'Name')",
          location: "Location",
          dates: "Duration (e.g., Jan 2024 - Present)",
          link: "Company URL / Portfolio URL from database 'Link' column (if available)",
          duplicate_group: "Duplicate Group code from database (if available)",
          bullets: ["Tailored resume bullet 1 starting with a strong action verb", "Tailored resume bullet 2", "Tailored resume bullet 3"]
        }
      ],
      projects: [
        {
          title: "Project name (mapped from Database 'Name')",
          tech_stack: "Combined list of applicable skills, technologies, and role context from Database Columns E ('Skills applicable'), D ('Technology used'), and G ('Role / Context') matching the JD (e.g. 'Lead Developer | Python, System Architecture')",
          link: "Project link / URL from database 'Link' column (if available)",
          duplicate_group: "Duplicate Group code from database (if available)",
          bullets: ["Tailored resume bullet 1 starting with a strong action verb", "Tailored resume bullet 2 covering implementation/architecture", "Tailored resume bullet 3 covering metrics/impact"]
        }
      ],
      certifications: [
        {
          title: "Certification title (mapped from Database 'Name')",
          org: "Issuing organization (mapped from 'Description' or 'Company')",
          date: "Date issued (mapped from 'Start Date' or 'End Date')",
          duplicate_group: "Duplicate Group code from database (if available)",
          link: "Verification link from database 'Link' column (if available)"
        }
      ],
      responsibilities: [
        {
          role: "Position of Responsibility role (mapped from Database 'Name' or 'Role')",
          org: "Organization (mapped from 'Company' or 'Description')",
          dates: "Duration (e.g. Sep 2024 - Dec 2024)",
          link: "Verification / Info link from database 'Link' column (if available)",
          duplicate_group: "Duplicate Group code from database (if available)",
          bullets: ["Tailored bullet 1", "Tailored bullet 2"]
        }
      ]
    }
  };

  const payload = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: JSON.stringify(userPrompt) }
    ],
    response_format: { type: "json_object" },
    temperature: 0.2,
    max_tokens: 4000
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://github.com/hisiddharth/REvolver',
      'X-Title': 'Resume Revolver'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxAttempts = 2;
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode !== 200) {
        const retryable = responseCode === 429 || responseCode >= 500;
        const apiError = new Error(`OpenRouter API call failed (HTTP ${responseCode}): ${responseText}`);
        if (!retryable || attempt === maxAttempts) throw apiError;
        lastError = apiError;
      } else {
        const resultObj = JSON.parse(responseText);
        if (!resultObj.choices || resultObj.choices.length === 0 || !resultObj.choices[0].message) {
          throw new Error('No text returned in OpenRouter API response choices.');
        }
        const contentText = resultObj.choices[0].message.content;
        try {
          let jsonStr = contentText.trim();
          if (jsonStr.startsWith('```')) {
            jsonStr = jsonStr.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
          }
          return JSON.parse(jsonStr);
        } catch (parseError) {
          const firstBrace = contentText.indexOf('{');
          const lastBrace = contentText.lastIndexOf('}');
          if (firstBrace !== -1 && lastBrace > firstBrace) {
            try {
              return JSON.parse(contentText.substring(firstBrace, lastBrace + 1));
            } catch (e) {
              // ignore inner catch
            }
          }
          console.warn(`OpenRouter returned invalid JSON on attempt ${attempt}.`);
          lastError = new Error('AI did not return valid JSON: ' + parseError.toString());
          if (attempt === maxAttempts) throw lastError;
        }
      }
    } catch (error) {
      lastError = error;
      const retryable = /timed out|timeout|temporary|service unavailable|rate limit|invalid JSON|HTTP (429|5\d\d)/i.test(error.toString());
      if (attempt === maxAttempts || !retryable) throw error;
    }
    Utilities.sleep(1000 * attempt);
  }
  throw lastError || new Error('OpenRouter request failed.');
}


// ==========================================
// GOOGLE DOCUMENT GENERATION & PLACEHOLDERS
// ==========================================

function generateTailoredResume(aiOutput, jobTitle, company) {
  const templateDocId = getSetting('TEMPLATE_DOC_ID', CONFIG.TEMPLATE_DOC_ID);
  const outputFolderId = getSetting('OUTPUT_FOLDER_ID', CONFIG.OUTPUT_FOLDER_ID);
  
  if (!templateDocId || templateDocId === 'YOUR_TEMPLATE_DOC_ID_HERE') {
    throw new Error('Template Google Doc ID is not configured.');
  }
  if (!outputFolderId || outputFolderId === 'YOUR_OUTPUT_FOLDER_ID_HERE') {
    throw new Error('Output Google Drive folder ID is not configured.');
  }

  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMMdd_h:mm_a");
  const cleanCompany = (company || "TailoredCompany").replace(/[^a-zA-Z0-9]/g, "_");
  const cleanJobTitle = (jobTitle || "TailoredRole").replace(/[^a-zA-Z0-9]/g, "_");
  const newFileName = `${formattedDate}_${cleanCompany}_${cleanJobTitle}`;

  const templateFile = DriveApp.getFileById(templateDocId);
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const copiedFile = templateFile.makeCopy(newFileName, outputFolder);
  const copiedDocId = copiedFile.getId();

  try {
    const doc = DocumentApp.openById(copiedDocId);
  const body = doc.getBody();

  processPlaceholder(body, TEMPLATE_TAGS.SUMMARY, aiOutput.summary || '', false);
  processPlaceholder(body, TEMPLATE_TAGS.SKILLS, aiOutput.skills || '', false);
  processPlaceholder(body, TEMPLATE_TAGS.COMPANY, company || '', false);
  processPlaceholder(body, TEMPLATE_TAGS.JOB_TITLE, jobTitle || '', false);

  const skillGroups = aiOutput.skill_groups || [];
  for (let i = 1; i <= CONFIG.MAX_SKILL_GROUPS; i++) {
    const group = skillGroups[i - 1];
    const exists = !!group;
    
    const titleTag = getTag(TEMPLATE_TAGS.SKILL_GROUP_TITLE, i);
    const listTag = getTag(TEMPLATE_TAGS.SKILL_GROUP_LIST, i);
    
    processPlaceholder(body, titleTag, exists ? group.title : '', !exists);
    processPlaceholder(body, listTag, exists ? group.skills : '', !exists);
  }

  const experiences = aiOutput.experiences || [];
  for (let i = 1; i <= CONFIG.MAX_EXPERIENCES; i++) {
    const exp = experiences[i - 1];
    const exists = !!exp;
    
    const roleTag = getTag(TEMPLATE_TAGS.EXP_ROLE, i);
    const companyTag = getTag(TEMPLATE_TAGS.EXP_COMPANY, i);
    const datesTag = getTag(TEMPLATE_TAGS.EXP_DATES, i);
    const locationTag = getTag(TEMPLATE_TAGS.EXP_LOCATION, i);
    
    const expLink = exists ? exp.link : '';
    processPlaceholder(body, roleTag, exists ? exp.role : '', !exists);
    processPlaceholder(body, companyTag, exists ? exp.company : '', !exists, expLink);
    processPlaceholder(body, datesTag, exists ? exp.dates : '', !exists);
    processPlaceholder(body, locationTag, exists ? exp.location : '', !exists);
    
    for (let j = 1; j <= CONFIG.MAX_EXPERIENCE_BULLETS; j++) {
      const bulletVal = exists && exp.bullets ? exp.bullets[j - 1] : '';
      const bulletTag = getTag(TEMPLATE_TAGS.EXP_BULLET, i, j);
      processPlaceholder(body, bulletTag, bulletVal, true);
    }
  }

  const projects = aiOutput.projects || [];
  for (let i = 1; i <= CONFIG.MAX_PROJECTS; i++) {
    const proj = projects[i - 1];
    const exists = !!proj;
    
    const titleTag = getTag(TEMPLATE_TAGS.PROJECT_TITLE, i);
    const techTag = getTag(TEMPLATE_TAGS.PROJECT_TECH, i);
    const linkTag = getTag(TEMPLATE_TAGS.PROJECT_LINK, i);
    
    const projLink = exists ? proj.link : '';
    processPlaceholder(body, titleTag, exists ? proj.title : '', !exists, projLink);
    processPlaceholder(body, techTag, exists ? proj.tech_stack : '', !exists);
    processPlaceholder(body, linkTag, exists ? proj.link : '', !exists);
    
    for (let j = 1; j <= CONFIG.MAX_PROJECT_BULLETS; j++) {
      const bulletVal = exists && proj.bullets ? proj.bullets[j - 1] : '';
      const bulletTag = getTag(TEMPLATE_TAGS.PROJECT_BULLET, i, j);
      processPlaceholder(body, bulletTag, bulletVal, true);
    }
  }

  const certifications = aiOutput.certifications || [];
  for (let i = 1; i <= CONFIG.MAX_CERTIFICATIONS; i++) {
    const cert = certifications[i - 1];
    const exists = !!cert;
    
    const titleTag = getTag(TEMPLATE_TAGS.CERT_TITLE, i);
    const orgTag = getTag(TEMPLATE_TAGS.CERT_ORG, i);
    const dateTag = getTag(TEMPLATE_TAGS.CERT_DATE, i);
    
    const certLink = exists ? cert.link : '';
    processPlaceholder(body, titleTag, exists ? cert.title : '', !exists, certLink);
    processPlaceholder(body, orgTag, exists ? cert.org : '', !exists);
    processPlaceholder(body, dateTag, exists ? cert.date : '', !exists);
  }

  const responsibilities = aiOutput.responsibilities || [];
  for (let i = 1; i <= CONFIG.MAX_RESPONSIBILITIES; i++) {
    const por = responsibilities[i - 1];
    const exists = !!por;
    
    const roleTag = getTag(TEMPLATE_TAGS.POR_ROLE, i);
    const orgTag = getTag(TEMPLATE_TAGS.POR_ORG, i);
    const datesTag = getTag(TEMPLATE_TAGS.POR_DATES, i);
    
    const porLink = exists ? por.link : '';
    processPlaceholder(body, roleTag, exists ? por.role : '', !exists);
    processPlaceholder(body, orgTag, exists ? por.org : '', !exists, porLink);
    processPlaceholder(body, datesTag, exists ? por.dates : '', !exists);
    
    for (let j = 1; j <= CONFIG.MAX_POR_BULLETS; j++) {
      const bulletVal = exists && por.bullets ? por.bullets[j - 1] : '';
      const bulletTag = getTag(TEMPLATE_TAGS.POR_BULLET, i, j);
      processPlaceholder(body, bulletTag, bulletVal, true);
    }
  }


    doc.saveAndClose();
    return copiedFile.getUrl();
  } catch (error) {
    try {
      copiedFile.setTrashed(true);
    } catch (cleanupError) {
      console.warn("Could not remove failed resume copy: " + cleanupError.toString());
    }
    throw error;
  }
}

// ==========================================
// HELPERS
// ==========================================

function escapeRegex(str) {
  if (!str) return '';
  let escaped = str.toString().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // If placeholder ends with \}\}, make the second closing brace optional so it matches both {{TAG}} and {{TAG} (typo tolerance)
  if (escaped.endsWith('\\}\\}')) {
    escaped = escaped.slice(0, -4) + '\\}\\}?';
  }
  return escaped;
}

function getTag(pattern, index, bulletIndex) {
  if (!pattern) return '';
  let tag = pattern.replace('%d', index);
  if (bulletIndex !== undefined) {
    tag = tag.replace('%b', bulletIndex);
  }
  return tag;
}

function processPlaceholder(body, placeholder, value, deleteParagraph, linkUrl) {
  if (value === null || value === undefined) {
    value = '';
  }
  
  const strValue = value.toString().trim();
  const searchPattern = escapeRegex(placeholder);
  
  if (strValue === '') {
    if (deleteParagraph) {
      removeParagraphsContaining(body, placeholder);
    } else {
      let rangeElement = body.findText(searchPattern);
      if (!rangeElement) {
        // Safe fallback replace text using a space if it is the only text
        try {
          body.replaceText(searchPattern, ' ');
        } catch (e) {
          console.warn("Fallback replaceText failed: " + e.toString());
        }
        return;
      }
      
      while (rangeElement) {
        const textElement = rangeElement.getElement().asText();
        const startOffset = rangeElement.getStartOffset();
        const endOffsetInclusive = rangeElement.getEndOffsetInclusive();
        
        // If the placeholder is the entire text of the element, replace with a single space " " to avoid exception.
        if (startOffset === 0 && endOffsetInclusive === textElement.getText().length - 1) {
          textElement.setText(' ');
        } else {
          textElement.deleteText(startOffset, endOffsetInclusive);
        }
        
        rangeElement = body.findText(searchPattern);
      }
    }
  } else {
    let rangeElement = body.findText(searchPattern);
    
    if (!rangeElement) {
      body.replaceText(searchPattern, strValue);
      return;
    }
    
    while (rangeElement) {
      const textElement = rangeElement.getElement().asText();
      const startOffset = rangeElement.getStartOffset();
      const endOffsetInclusive = rangeElement.getEndOffsetInclusive();
      
      // 1. Capture the original formatting attributes of the placeholder text
      const attributes = textElement.getAttributes(startOffset);
      
      // 2. Insert the new text before the placeholder
      textElement.insertText(startOffset, strValue);
      
      // 3. Explicitly apply the captured attributes to the new text to prevent style leakage
      if (attributes) {
        textElement.setAttributes(startOffset, startOffset + strValue.length - 1, attributes);
      }
      
      // 4. Delete the original placeholder text (which has shifted to the right)
      const shift = strValue.length;
      textElement.deleteText(startOffset + shift, endOffsetInclusive + shift);
      
      // 5. Apply the hyperlink if a valid URL is provided
      if (linkUrl && linkUrl.trim() !== '') {
        const link = linkUrl.trim();
        textElement.setLinkUrl(startOffset, startOffset + strValue.length - 1, link);
      }
      
      // Find the next occurrence
      rangeElement = body.findText(searchPattern);
    }
  }
}

function removeParagraphsContaining(body, placeholder) {
  const searchPattern = escapeRegex(placeholder);
  let rangeElement = body.findText(searchPattern);
  while (rangeElement) {
    const textElement = rangeElement.getElement();
    let paragraph = textElement.getParent();
    
    // Walk up until we hit a PARAGRAPH or LIST_ITEM, or hit top container
    while (paragraph && 
           paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH && 
           paragraph.getType() !== DocumentApp.ElementType.LIST_ITEM &&
           paragraph.getParent()) {
      paragraph = paragraph.getParent();
    }
    
    if (!paragraph) {
      try {
        body.replaceText(searchPattern, ' ');
      } catch (e) {
        console.warn("Fallback replaceText failed: " + e.toString());
      }
      rangeElement = body.findText(searchPattern);
      continue;
    }
    
    const parentContainer = paragraph.getParent();
    const isOnlyChildInContainer = parentContainer && 
      typeof parentContainer.getNumChildren === 'function' && 
      parentContainer.getNumChildren() <= 1;

    const parentType = paragraph.getType();
    if (parentType === DocumentApp.ElementType.PARAGRAPH || parentType === DocumentApp.ElementType.LIST_ITEM) {
      if (parentType === DocumentApp.ElementType.LIST_ITEM) {
        try {
          paragraph.setListId(null);
        } catch (e) {
          console.warn("Failed to clear list ID: " + e.toString());
        }
      }
      
      if (!isOnlyChildInContainer) {
        paragraph.removeFromParent();
      } else {
        paragraph.setText(' ');
      }
    } else {
      try {
        body.replaceText(searchPattern, ' ');
      } catch (e) {
        console.warn("Fallback replaceText failed: " + e.toString());
      }
    }
    
    rangeElement = body.findText(searchPattern);
  }
}

function getSetting(key, defaultValue) {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty(key);
    return prop !== null && prop !== undefined ? prop : defaultValue;
  } catch (e) {
    console.warn(`Error reading script property for key "${key}": ` + e.toString());
    return defaultValue;
  }
}
