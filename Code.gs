/**
 * Google Apps Script - Huiswerk & Toetsen Volgsysteem
 * Rollen: Docent / Leerling
 * @OnlyCurrentDoc
 */

const SESSION_TTL = 21600; // 6 uur

function onOpen(e) {
  // Wordt automatisch uitgevoerd wanneer de spreadsheet wordt geopend.
  // Let op: voer deze functie niet handmatig uit vanuit de Apps Script-editor;
  // open/herlaad in plaats daarvan de gekoppelde Google Spreadsheet.
  try {
    SpreadsheetApp.getUi()
      .createMenu('📚 Huiswerk & Toetsen App')
      .addItem('🚀 Open App', 'showHomeworkDialog')
      .addSeparator()
      .addItem('⚙️ Tabbladen Initialiseren / Resetten', 'setupSheet')
      .addToUi();
  } catch (err) {
    console.log('onOpen: ' + err.message);
  }
}

// Nodig wanneer Apps Script een installable trigger gebruikt.
function onInstall(e) {
  onOpen(e);
}

function showHomeworkDialog() {
  setupSheet();
  const html = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📚 Huiswerk & Toetsen App');
}

/* =========================
   SHEETS / DATASTRUCTUUR
   ========================= */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, 'Huiswerk Overzicht',
    ['ID','Datum Ingevoerd','Docent','Klas','Leerling','Vak','Omschrijving','Inleverdatum','Status','LeerlingUsername']);

  ensureSheet_(ss, 'Docenten',
    ['Gebruikersnaam','Wachtwoord','Volledige Naam']);

  ensureSheet_(ss, 'Leerlingen',
    ['Gebruikersnaam','Wachtwoord','Volledige Naam']);

  ensureSheet_(ss, 'Klassen & Leerlingen',
    ['Docent','Klas','Leerlingnaam','LeerlingUsername']);

  ensureSheet_(ss, 'Toetsen Overzicht',
    ['Toets ID','Datum Ingevoerd','Docent','Klas','Vak','Toetsnaam','Toetsdatum','Vragen JSON','Resultaten JSON']);

  ensureSheet_(ss, 'Toets Toestemming',
    ['Toets ID','LeerlingUsername','Toestemming']);

  // Upgrade van de oude 3-koloms Klassen-sheet.
  const classSheet = ss.getSheetByName('Klassen & Leerlingen');
  if (classSheet.getLastColumn() < 4) {
    classSheet.getRange(1,4).setValue('LeerlingUsername');
  }

  // Upgrade van de oude 9-koloms Huiswerk-sheet.
  const hwSheet = ss.getSheetByName('Huiswerk Overzicht');
  if (hwSheet.getLastColumn() < 10) {
    hwSheet.getRange(1,10).setValue('LeerlingUsername');
  }

  SpreadsheetApp.flush();
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const currentHeaders = sheet.getRange(1,1,1,headers.length).getValues()[0];
  let needsHeaders = sheet.getLastRow() === 0;
  if (!needsHeaders) {
    for (let i = 0; i < headers.length; i++) {
      if (String(currentHeaders[i] || '').trim() !== headers[i]) {
        needsHeaders = true;
        break;
      }
    }
  }

  if (needsHeaders) {
    sheet.getRange(1,1,1,headers.length)
      .setValues([headers])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

/* =========================
   SESSIES / AUTHENTICATIE
   ========================= */

function createSession_(role, username, name) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify({
    role: role,
    username: username,
    name: name,
    created: Date.now()
  }), SESSION_TTL);
  return token;
}

function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(String(token));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function requireRole_(token, role) {
  const session = getSession_(token);
  if (!session || session.role !== role) {
    throw new Error('Je sessie is ongeldig of je hebt geen toestemming voor deze actie.');
  }
  return session;
}

function logoutAccount(token) {
  if (token) CacheService.getScriptCache().remove(String(token));
  return {success:true};
}

function registerAccount(role, username, password, fullName) {
  try {
    role = String(role || '').toLowerCase();
    const cleanUser = String(username || '').trim();
    const cleanPass = String(password || '').trim();
    const cleanName = String(fullName || '').trim();

    if (!['teacher','student'].includes(role)) {
      return {success:false, message:'Kies eerst Docent of Leerling.'};
    }
    if (!cleanUser || !cleanPass || !cleanName) {
      return {success:false, message:'Vul alle velden in.'};
    }
    if (cleanUser.length < 3) {
      return {success:false, message:'De gebruikersnaam moet minimaal 3 tekens hebben.'};
    }
    if (cleanPass.length < 4) {
      return {success:false, message:'Het wachtwoord moet minimaal 4 tekens hebben.'};
    }

    setupSheet();

    // Gebruikersnamen zijn globaal uniek tussen leerlingen en docenten.
    const teacherSheet = getSheet_('Docenten');
    const studentSheet = getSheet_('Leerlingen');

    if (usernameExists_(teacherSheet, cleanUser) || usernameExists_(studentSheet, cleanUser)) {
      return {success:false, message:'Deze gebruikersnaam is al in gebruik.'};
    }

    const sheet = role === 'teacher' ? teacherSheet : studentSheet;
    sheet.appendRow([cleanUser, cleanPass, cleanName]);
    SpreadsheetApp.flush();

    return {
      success:true,
      message: role === 'teacher'
        ? 'Docentenaccount aangemaakt. Je kunt nu inloggen.'
        : 'Leerlingaccount aangemaakt. Je kunt nu inloggen.'
    };
  } catch (err) {
    return {success:false, message:'Fout bij registreren: ' + err.message};
  }
}

function usernameExists_(sheet, username) {
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  const target = String(username).trim().toLowerCase();
  return data.some(r => String(r[0]).trim().toLowerCase() === target);
}

function loginAccount(role, username, password) {
  try {
    role = String(role || '').toLowerCase();
    const cleanUser = String(username || '').trim();
    const cleanPass = String(password || '').trim();

    if (!['teacher','student'].includes(role)) {
      return {success:false, message:'Ongeldige rol.'};
    }

    const sheet = getSheet_(role === 'teacher' ? 'Docenten' : 'Leerlingen');
    if (sheet.getLastRow() <= 1) {
      return {
        success:false,
        message: role === 'teacher'
          ? 'Er zijn nog geen docentenaccounts.'
          : 'Er zijn nog geen leerlingaccounts. Maak eerst een leerlingaccount aan.'
      };
    }

    const data = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
    for (let i = 0; i < data.length; i++) {
      const user = String(data[i][0]).trim();
      const pass = String(data[i][1]).trim();
      const name = String(data[i][2]).trim();

      if (user.toLowerCase() === cleanUser.toLowerCase() && pass === cleanPass) {
        const token = createSession_(role, user, name);
        return {
          success:true,
          token:token,
          role:role,
          username:user,
          name:name
        };
      }
    }

    return {success:false, message:'Ongeldige gebruikersnaam of wachtwoord.'};
  } catch (err) {
    return {success:false, message:'Fout bij inloggen: ' + err.message};
  }
}

// Compatibiliteit met de oude functies.
function registerTeacher(username, password, fullName) {
  return registerAccount('teacher', username, password, fullName);
}
function loginTeacher(username, password) {
  return loginAccount('teacher', username, password);
}

/* =========================
   DOCENT - KLASSEN
   ========================= */

function getClasses(token) {
  const session = requireRole_(token, 'teacher');
  const sheet = getSheet_('Klassen & Leerlingen');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = Math.min(4, sheet.getLastColumn());
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues();
  const names = new Set([String(session.name||'').trim().toLowerCase(),String(session.username||'').trim().toLowerCase()]);
  const result = new Set();
  rows.forEach(r => {
    const owner=String(r[0]||'').trim().toLowerCase(), klas=String(r[1]||'').trim();
    if(klas && names.has(owner)) result.add(klas);
  });
  return Array.from(result).sort((a,b)=>a.localeCompare(b,'nl'));
}

function addClass(token, className, studentUsernames) {
  try {
    const session = requireRole_(token, 'teacher');
    const cleanClass = String(className || '').trim();
    const students = Array.isArray(studentUsernames) ? studentUsernames : [];

    if (!cleanClass) return {success:false, message:'Vul een klasnaam in.'};

    const sheet = getSheet_('Klassen & Leerlingen');
    const existing = sheet.getLastRow() > 1
      ? sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues()
      : [];

    const ownerNames = new Set([
      String(session.name || '').trim().toLowerCase(),
      String(session.username || '').trim().toLowerCase()
    ]);

    if (existing.some(r =>
      ownerNames.has(String(r[0] || '').trim().toLowerCase()) &&
      String(r[1] || '').trim().toLowerCase() === cleanClass.toLowerCase()
    )) {
      return {success:false, message:'Deze klas bestaat al bij jouw account.'};
    }

    const validStudents = getStudentAccounts_();
    const validMap = {};
    validStudents.forEach(s => validMap[s.username.toLowerCase()] = s);

    const unique = [];
    const seen = {};
    students.forEach(u => {
      const key = String(u || '').trim().toLowerCase();
      if (key && validMap[key] && !seen[key]) {
        seen[key] = true;
        unique.push(validMap[key]);
      }
    });

    // Een klas zonder leden mag ook aangemaakt worden.
    if (unique.length === 0) {
      sheet.appendRow([session.name, cleanClass, '', '']);
    } else {
      const rows = unique.map(s => [session.name, cleanClass, s.name, s.username]);
      sheet.getRange(sheet.getLastRow()+1,1,rows.length,4).setValues(rows);
    }

    SpreadsheetApp.flush();
    return {
      success:true,
      message:`Klas "${cleanClass}" is aangemaakt met ${unique.length} leerling(en).`
    };
  } catch (err) {
    return {success:false, message:'Fout bij aanmaken klas: ' + err.message};
  }
}

function openClass(token, className) {
  const session = requireRole_(token, 'teacher');
  const wanted = String(className || '').trim();
  if (!wanted) return {success:false, message:'Geen klas geselecteerd.'};

  const sheet = getSheet_('Klassen & Leerlingen');
  if (sheet.getLastRow() <= 1) return {success:false, message:'Er zijn geen klassen opgeslagen.'};
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();

  // Accepteer zowel het huidige docentveld (volledige naam) als oude data
  // waarin per ongeluk de gebruikersnaam is opgeslagen.
  const ownerNames = new Set([
    String(session.name || '').trim().toLowerCase(),
    String(session.username || '').trim().toLowerCase()
  ]);

  const row = rows.find(r =>
    ownerNames.has(String(r[0] || '').trim().toLowerCase()) &&
    String(r[1] || '').trim().toLowerCase() === wanted.toLowerCase()
  );

  if (!row) return {success:false, message:'Deze klas is niet gekoppeld aan jouw docentaccount.'};

  const realClass = String(row[1]).trim();
  return {success:true, className:realClass, students:getClassMembers_(session.name, realClass)};
}

function getStudentAccounts(token) {
  requireRole_(token, 'teacher');
  return getStudentAccounts_().map(s => ({
    username:s.username,
    name:s.name
  }));
}

function getStudentAccounts_() {
  const sheet = getSheet_('Leerlingen');
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();

  return data
    .map(r => ({
      username:String(r[0]).trim(),
      name:String(r[2]).trim()
    }))
    .filter(s => s.username && s.name)
    .sort((a,b) => a.name.localeCompare(b.name));
}

function getStudentsInClass(token, selectedClass) {
  const session = requireRole_(token, 'teacher');
  return getClassMembers_(session.name, selectedClass).map(s => ({
    username:s.username,
    name:s.name
  }));
}

function getClassMembers_(teacherName, className) {
  const sheet = getSheet_('Klassen & Leerlingen');
  if (sheet.getLastRow() <= 1) return [];

  // Ondersteun zowel volledige docentnaam als gebruikersnaam in oude klasdata.
  const teacher = String(teacherName || '').trim();
  const ownerNames = new Set([teacher.toLowerCase()]);
  const teacherSheet = getSheet_('Docenten');
  if (teacherSheet.getLastRow() > 1) {
    const teachers = teacherSheet.getRange(2,1,teacherSheet.getLastRow()-1,3).getValues();
    teachers.forEach(r => {
      const user = String(r[0] || '').trim();
      const name = String(r[2] || '').trim();
      if (name.toLowerCase() === teacher.toLowerCase() && user) ownerNames.add(user.toLowerCase());
      if (user.toLowerCase() === teacher.toLowerCase() && name) ownerNames.add(name.toLowerCase());
    });
  }

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  const accounts = getStudentAccounts_();
  const byName = {};
  const byUser = {};
  accounts.forEach(s => {
    byName[s.name.toLowerCase()] = s;
    byUser[s.username.toLowerCase()] = s;
  });

  const result = [];
  const seen = {};

  data.forEach(row => {
    const docent = String(row[0]).trim();
    const klas = String(row[1]).trim();
    let name = String(row[2]).trim();
    let username = String(row[3]).trim();

    if (!ownerNames.has(docent.toLowerCase()) || klas !== String(className).trim()) return;

    // Alleen echte leerlingaccounts tellen als klaslid.
    // Oude/vervuilde rijen (bijv. een docent in de leerlingkolommen) worden genegeerd.
    const account = username ? byUser[username.toLowerCase()] : byName[name.toLowerCase()];
    if (!account) return;

    name = account.name;
    username = account.username;

    if (!seen[username.toLowerCase()]) {
      seen[username.toLowerCase()] = true;
      result.push({username:username, name:name});
    }
  });

  return result.sort((a,b) => a.name.localeCompare(b.name));
}

// Oude knop wordt niet meer gebruikt voor nieuwe accounts, maar blijft werken voor bestaande data.
function addStudentToClass(token, className, studentUsername) {
  try {
    const session = requireRole_(token, 'teacher');
    const members = getClassMembers_(session.name, className);
    const accounts = getStudentAccounts_();
    const account = accounts.find(s => s.username.toLowerCase() === String(studentUsername).trim().toLowerCase());

    if (!account) return {success:false, message:'Leerlingaccount niet gevonden.'};
    if (members.some(s => s.username.toLowerCase() === account.username.toLowerCase())) {
      return {success:false, message:'Deze leerling zit al in de klas.'};
    }

    getSheet_('Klassen & Leerlingen').appendRow([session.name, String(className).trim(), account.name, account.username]);
    SpreadsheetApp.flush();
    return {success:true, message:`${account.name} is toegevoegd aan de klas.`};
  } catch (err) {
    return {success:false, message:'Fout bij toevoegen leerling: ' + err.message};
  }
}

/* =========================
   DOCENT - HUISWERK
   ========================= */

function saveHomeworkForClass(token, data) {
  try {
    const session = requireRole_(token, 'teacher');
    data = data || {};

    const className = String(data.selectedClass || '').trim();
    const subject = String(data.subject || '').trim();
    const description = String(data.description || '').trim();
    const dueDate = String(data.dueDate || '').trim();
    const usernames = Array.isArray(data.students) ? data.students : [];

    if (!className || !subject || !description || !dueDate) {
      return {success:false, message:'Vul alle huiswerkvelden in.'};
    }

    const members = getClassMembers_(session.name, className);
    const allowed = {};
    members.forEach(s => allowed[s.username.toLowerCase()] = s);

    const selected = [];
    const seen = {};
    usernames.forEach(u => {
      const account = allowed[String(u).trim().toLowerCase()];
      if (account && !seen[account.username.toLowerCase()]) {
        seen[account.username.toLowerCase()] = true;
        selected.push(account);
      }
    });

    if (!selected.length) {
      return {success:false, message:'Selecteer minimaal één leerling uit deze klas.'};
    }

    const sheet = getSheet_('Huiswerk Overzicht');
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

    const rows = selected.map(student => [
      makeId_('HW'),
      timestamp,
      session.name,
      className,
      student.name,
      subject,
      description,
      dueDate,
      'Nog niet af',
      student.username
    ]);

    sheet.getRange(sheet.getLastRow()+1,1,rows.length,10).setValues(rows);
    SpreadsheetApp.flush();

    return {success:true, message:`Huiswerk toegewezen aan ${selected.length} leerling(en).`};
  } catch (err) {
    return {success:false, message:'Fout bij opslaan huiswerk: ' + err.message};
  }
}

function getClassHomework(token, selectedClass) {
  const session = requireRole_(token, 'teacher');
  return getHomeworkRows_().filter(item =>
    item.klas === String(selectedClass).trim() &&
    item.docent === session.name
  );
}

function getStudentHomework(token) {
  const session = requireRole_(token, 'student');
  return getHomeworkRows_().filter(item =>
    item.leerlingUsername.toLowerCase() === session.username.toLowerCase()
  );
}

function getHomeworkRows_() {
  const sheet = getSheet_('Huiswerk Overzicht');
  if (sheet.getLastRow() <= 1) return [];

  const lastCol = Math.max(sheet.getLastColumn(), 10);
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,lastCol).getValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';

  return data.map(row => {
    let due = String(row[7] || '');
    let rawDue = String(row[7] || '');
    if (row[7] instanceof Date) {
      due = Utilities.formatDate(row[7], tz, 'dd-MM-yyyy');
      rawDue = Utilities.formatDate(row[7], tz, 'yyyy-MM-dd');
    }

    return {
      id:String(row[0]),
      datum:row[1] instanceof Date ? Utilities.formatDate(row[1],tz,'dd-MM-yyyy') : String(row[1]),
      docent:String(row[2]).trim(),
      klas:String(row[3]).trim(),
      leerling:String(row[4]).trim(),
      vak:String(row[5]).trim(),
      omschrijving:String(row[6]).trim(),
      inleverdatum:due,
      inleverdatumRaw:rawDue,
      status:String(row[8]).trim() || 'Nog niet af',
      leerlingUsername:String(row[9] || '').trim()
    };
  }).filter(x => x.id);
}

function updateHomework(token, id, subject, description, dueDate) {
  try {
    const session = requireRole_(token, 'teacher');
    const sheet = getSheet_('Huiswerk Overzicht');
    const data = sheet.getLastRow() > 1
      ? sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues()
      : [];

    for (let i=0; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).trim() !== session.name) {
          return {success:false, message:'Je mag alleen je eigen huiswerk aanpassen.'};
        }
        sheet.getRange(i+2,6,1,3).setValues([[subject,description,dueDate]]);
        SpreadsheetApp.flush();
        return {success:true, message:'Huiswerk aangepast.'};
      }
    }
    return {success:false, message:'Huiswerk ID niet gevonden.'};
  } catch (err) {
    return {success:false, message:'Fout bij aanpassen: ' + err.message};
  }
}

function deleteHomework(token, id) {
  try {
    const session = requireRole_(token, 'teacher');
    const sheet = getSheet_('Huiswerk Overzicht');
    if (sheet.getLastRow() <= 1) return {success:false, message:'Geen huiswerk gevonden.'};

    const data = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
    for (let i=0; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).trim() !== session.name) {
          return {success:false, message:'Je mag alleen je eigen huiswerk verwijderen.'};
        }
        sheet.deleteRow(i+2);
        SpreadsheetApp.flush();
        return {success:true, message:'Huiswerk verwijderd.'};
      }
    }
    return {success:false, message:'Huiswerk ID niet gevonden.'};
  } catch (err) {
    return {success:false, message:'Fout bij verwijderen: ' + err.message};
  }
}

function toggleHomeworkStatus(token, id, currentStatus) {
  try {
    const session = requireRole_(token, 'teacher');
    const sheet = getSheet_('Huiswerk Overzicht');
    if (sheet.getLastRow() <= 1) return {success:false, message:'Geen huiswerk gevonden.'};

    const data = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
    for (let i=0; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).trim() !== session.name) {
          return {success:false, message:'Je mag alleen huiswerk van je eigen account afvinken.'};
        }
        const newStatus = String(currentStatus) === 'Af' ? 'Nog niet af' : 'Af';
        sheet.getRange(i+2,9).setValue(newStatus);
        SpreadsheetApp.flush();
        return {success:true, newStatus:newStatus};
      }
    }
    return {success:false, message:'Huiswerk ID niet gevonden.'};
  } catch (err) {
    return {success:false, message:'Fout bij status wijzigen: ' + err.message};
  }
}

/* =========================
   DOCENT - TOETSEN
   ========================= */

function saveTest(token, data) {
  try {
    const session = requireRole_(token, 'teacher');
    data = data || {};

    const sheet = getSheet_('Toetsen Overzicht');
    const className = String(data.selectedClass || '').trim();
    const testName = String(data.testName || '').trim();
    const subject = String(data.subject || '').trim();
    const testDate = String(data.testDate || '').trim();
    const questions = Array.isArray(data.questions) ? data.questions : [];
    const results = data.results || {};
    const allowedMembers = getClassMembers_(session.name, className);

    if (!className || !testName || !subject || !testDate) {
      return {success:false, message:'Vul alle toetsgegevens in.'};
    }
    if (!questions.length) {
      return {success:false, message:'Voeg minimaal één vraag toe.'};
    }

    // Alleen leerlingen uit de klas mogen resultaten krijgen.
    const allowed = {};
    allowedMembers.forEach(s => allowed[s.username.toLowerCase()] = s);

    const cleanResults = {};
    Object.keys(results).forEach(username => {
      const account = allowed[String(username).toLowerCase()];
      if (!account) return;

      cleanResults[account.username] = {};
      const source = results[username] || {};
      questions.forEach((q, idx) => {
        const qid = String(q.id || ('q_' + idx));
        const value = source[qid];
        cleanResults[account.username][qid] =
          value === '' || value === null || value === undefined ? '' : Number(value);
      });
    });

    const questionsJson = JSON.stringify(questions);
    const resultsJson = JSON.stringify(cleanResults);
    let testId = String(data.id || '').trim();

    if (testId) {
      const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues()
        : [];

      for (let i=0; i<rows.length; i++) {
        if (String(rows[i][0]) === testId) {
          if (String(rows[i][2]).trim() !== session.name) {
            return {success:false, message:'Je mag alleen je eigen toetsen bewerken.'};
          }

          sheet.getRange(i+2,5,1,5).setValues([[
            subject,testName,testDate,questionsJson,resultsJson
          ]]);

          saveTestPermissions_(testId, session.name, className, data.permissions || []);
          SpreadsheetApp.flush();
          return {success:true, message:'Toets, cijfers en toestemmingen bijgewerkt.'};
        }
      }
      return {success:false, message:'Toets niet gevonden.'};
    }

    testId = makeId_('TOETS');
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

    sheet.appendRow([
      testId,timestamp,session.name,className,subject,testName,testDate,questionsJson,resultsJson
    ]);

    saveTestPermissions_(testId, session.name, className, data.permissions || []);
    SpreadsheetApp.flush();

    return {success:true, message:'Toets en resultaten opgeslagen.'};
  } catch (err) {
    return {success:false, message:'Fout bij opslaan toets: ' + err.message};
  }
}

function getClassTests(token, selectedClass) {
  const session = requireRole_(token, 'teacher');
  return getTeacherTests_(session.name, selectedClass);
}

function getTeacherTests_(teacherName, selectedClass) {
  const sheet = getSheet_('Toetsen Overzicht');
  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';

  return data.map(row => {
    if (String(row[2]).trim() !== String(teacherName).trim() ||
        String(row[3]).trim() !== String(selectedClass).trim()) return null;

    let date = String(row[6] || '');
    let rawDate = date;
    if (row[6] instanceof Date) {
      date = Utilities.formatDate(row[6],tz,'dd-MM-yyyy');
      rawDate = Utilities.formatDate(row[6],tz,'yyyy-MM-dd');
    }

    let questions = [];
    let results = {};
    try { questions = JSON.parse(row[7] || '[]'); } catch(e) {}
    try { results = JSON.parse(row[8] || '{}'); } catch(e) {}

    const permissions = getTestPermissions_(String(row[0]));
    return {
      id:String(row[0]),
      datumIngevoerd:row[1] instanceof Date ? Utilities.formatDate(row[1],tz,'dd-MM-yyyy') : String(row[1]),
      docent:String(row[2]),
      klas:String(row[3]),
      subject:String(row[4]),
      testName:String(row[5]),
      testDate:date,
      testDateRaw:rawDate,
      questions:questions,
      results:results,
      permissions:permissions
    };
  }).filter(Boolean).reverse();
}

function getStudentTests(token) {
  const session = requireRole_(token, 'student');
  const memberships = getStudentMemberships_(session.username, session.name);
  const all = [];
  const seenTests = {};

  memberships.forEach(membership => {
    getAllTestsForClass_(membership.className, membership.teacherName).forEach(test => {
      if (seenTests[test.id]) return;
      seenTests[test.id] = true;

      const permission = hasTestPermission_(test.id, session.username);
      const ownResult = test.results[session.username] || test.results[session.name] || {};
      const total = calculateTotal_(test.questions, ownResult);

      const item = {
        id:test.id,
        klas:test.klas,
        docent:test.docent,
        subject:test.subject,
        testName:test.testName,
        testDate:test.testDate,
        total:total.total,
        max:total.max,
        percentage:total.percentage,
        grade:total.grade,
        canViewTest:permission
      };

      // Vraaginhoud wordt alleen naar de leerling gestuurd als er toestemming is.
      if (permission) item.questions = test.questions;
      all.push(item);
    });
  });

  return all.sort((a,b) => String(b.testDate).localeCompare(String(a.testDate)));
}

function getStudentTestDetails(token, testId) {
  const session = requireRole_(token, 'student');
  if (!hasTestPermission_(testId, session.username)) {
    return {success:false, message:'De docent heeft je geen toestemming gegeven om deze toets in te zien.'};
  }

  const test = getTestById_(testId);
  if (!test) return {success:false, message:'Toets niet gevonden.'};

  if (!studentIsInClass_(session.username, session.name, test.klas, test.docent)) {
    return {success:false, message:'Je hebt geen toegang tot deze toets.'};
  }

  const ownResult = test.results[session.username] || test.results[session.name] || {};
  return {
    success:true,
    id:test.id,
    testName:test.testName,
    subject:test.subject,
    testDate:test.testDate,
    questions:test.questions,
    score:calculateTotal_(test.questions, ownResult)
  };
}

function getAllTestsForClass_(className, teacherName) {
  const sheet = getSheet_('Toetsen Overzicht');
  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';

  return data.map(row => {
    if (String(row[3]).trim() !== String(className).trim()) return null;
    if (teacherName && String(row[2]).trim() !== String(teacherName).trim()) return null;

    let date = String(row[6] || '');
    if (row[6] instanceof Date) date = Utilities.formatDate(row[6],tz,'dd-MM-yyyy');

    let questions = [];
    let results = {};
    try { questions = JSON.parse(row[7] || '[]'); } catch(e) {}
    try { results = JSON.parse(row[8] || '{}'); } catch(e) {}

    return {
      id:String(row[0]),
      docent:String(row[2]),
      klas:String(row[3]),
      subject:String(row[4]),
      testName:String(row[5]),
      testDate:date,
      questions:questions,
      results:results
    };
  }).filter(Boolean);
}

function getTestById_(testId) {
  const sheet = getSheet_('Toetsen Overzicht');
  if (sheet.getLastRow() <= 1) return null;

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Europe/Amsterdam';

  for (const row of data) {
    if (String(row[0]) !== String(testId)) continue;

    let date = String(row[6] || '');
    if (row[6] instanceof Date) date = Utilities.formatDate(row[6],tz,'dd-MM-yyyy');

    let questions = [];
    let results = {};
    try { questions = JSON.parse(row[7] || '[]'); } catch(e) {}
    try { results = JSON.parse(row[8] || '{}'); } catch(e) {}

    return {
      id:String(row[0]),
      docent:String(row[2]),
      klas:String(row[3]),
      subject:String(row[4]),
      testName:String(row[5]),
      testDate:date,
      questions:questions,
      results:results
    };
  }
  return null;
}

function deleteTest(token, testId) {
  try {
    const session = requireRole_(token, 'teacher');
    const sheet = getSheet_('Toetsen Overzicht');
    if (sheet.getLastRow() <= 1) return {success:false, message:'Geen toetsen gevonden.'};

    const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
    for (let i=0; i<data.length; i++) {
      if (String(data[i][0]) === String(testId)) {
        if (String(data[i][2]).trim() !== session.name) {
          return {success:false, message:'Je mag alleen je eigen toetsen verwijderen.'};
        }
        sheet.deleteRow(i+2);
        deleteTestPermissions_(testId);
        SpreadsheetApp.flush();
        return {success:true, message:'Toets verwijderd.'};
      }
    }
    return {success:false, message:'Toets ID niet gevonden.'};
  } catch (err) {
    return {success:false, message:'Fout bij verwijderen toets: ' + err.message};
  }
}

function saveTestPermissions_(testId, teacherName, className, usernames) {
  const sheet = getSheet_('Toets Toestemming');
  const members = getClassMembers_(teacherName, className);
  const allowed = {};
  members.forEach(s => allowed[s.username.toLowerCase()] = s);

  // Verwijder oude toestemmingen voor deze toets.
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
    for (let i=data.length-1; i>=0; i--) {
      if (String(data[i][0]) === String(testId)) sheet.deleteRow(i+2);
    }
  }

  const unique = {};
  const rows = [];
  (Array.isArray(usernames) ? usernames : []).forEach(u => {
    const account = allowed[String(u).toLowerCase()];
    if (account && !unique[account.username.toLowerCase()]) {
      unique[account.username.toLowerCase()] = true;
      rows.push([testId,account.username,'JA']);
    }
  });

  if (rows.length) sheet.getRange(sheet.getLastRow()+1,1,rows.length,3).setValues(rows);
}

function getTestPermissions_(testId) {
  const sheet = getSheet_('Toets Toestemming');
  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
  return data
    .filter(r => String(r[0]) === String(testId) && String(r[2]).toUpperCase() === 'JA')
    .map(r => String(r[1]));
}

function hasTestPermission_(testId, username) {
  return getTestPermissions_(testId).some(u =>
    String(u).toLowerCase() === String(username).toLowerCase()
  );
}

function deleteTestPermissions_(testId) {
  const sheet = getSheet_('Toets Toestemming');
  if (sheet.getLastRow() <= 1) return;

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
  for (let i=data.length-1; i>=0; i--) {
    if (String(data[i][0]) === String(testId)) sheet.deleteRow(i+2);
  }
}

/* =========================
   LEERLING HELPERS
   ========================= */

function getStudentMemberships_(username, name) {
  const sheet = getSheet_('Klassen & Leerlingen');
  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  const result = [];
  const seen = {};

  data.forEach(row => {
    const teacherName = String(row[0]).trim();
    const className = String(row[1]).trim();
    const rowName = String(row[2]).trim();
    const rowUser = String(row[3]).trim();

    if (!teacherName || !className) return;

    const isMember =
      rowUser.toLowerCase() === String(username).toLowerCase() ||
      (!rowUser && rowName.toLowerCase() === String(name).toLowerCase());

    if (!isMember) return;

    const key = teacherName.toLowerCase() + '|' + className.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      result.push({teacherName:teacherName, className:className});
    }
  });

  return result;
}

function getStudentClassNames_(username, name) {
  const sheet = getSheet_('Klassen & Leerlingen');
  if (sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  const result = new Set();

  data.forEach(row => {
    const className = String(row[1]).trim();
    const rowName = String(row[2]).trim();
    const rowUser = String(row[3]).trim();

    if (!className) return;
    if (
      rowUser.toLowerCase() === String(username).toLowerCase() ||
      (!rowUser && rowName.toLowerCase() === String(name).toLowerCase())
    ) {
      result.add(className);
    }
  });

  return Array.from(result);
}

function studentIsInClass_(username, name, className, teacherName) {
  const sheet = getSheet_('Klassen & Leerlingen');
  if (sheet.getLastRow() <= 1) return false;

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  return data.some(row => {
    return String(row[0]).trim() === String(teacherName).trim() &&
      String(row[1]).trim() === String(className).trim() &&
      (
        String(row[3]).trim().toLowerCase() === String(username).trim().toLowerCase() ||
        (!String(row[3]).trim() &&
         String(row[2]).trim().toLowerCase() === String(name).trim().toLowerCase())
      );
  });
}

/* =========================
   CIJFER / UTILITIES
   ========================= */

function calculateTotal_(questions, scores) {
  questions = Array.isArray(questions) ? questions : [];
  scores = scores || {};

  let total = 0;
  let max = 0;

  questions.forEach((q, idx) => {
    const qid = String(q.id || ('q_' + idx));
    const qMax = Number(q.maxPoints || 0);
    max += qMax;

    const value = scores[qid];
    if (value !== '' && value !== null && value !== undefined && !isNaN(Number(value))) {
      total += Number(value);
    }
  });

  total = Math.max(0, total);
  const percentage = max > 0 ? Math.max(0, Math.min(100, total / max * 100)) : 0;

  // Indicatief Nederlands 1-10 cijfer: 0% = 1, 100% = 10.
  const grade = max > 0 ? Math.round((1 + 9 * (percentage / 100)) * 10) / 10 : null;

  return {total:total, max:max, percentage:Math.round(percentage*10)/10, grade:grade};
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g,'').substring(0,12).toUpperCase();
}

// Oude functies met de oude signatuur worden bewust niet gebruikt door de nieuwe HTML.
// Ze zijn hierboven vervangen door token-gebaseerde functies voor rol- en toegangscontrole.
