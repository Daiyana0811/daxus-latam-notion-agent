require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { Client } = require('@notionhq/client');
const {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  TextRun,
  AlignmentType
} = require('docx');

const TRANSCRIPTION_PROPERTY_NAME = 'Transcripcion';
const APOSTILLA_PROPERTY_NAME = 'Apostilla';
const DOCX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const VIDEO_FILE_PATTERN = /\.(mp4|mov|m4v|webm|mkv)$/i;

const SHAREPOINT_SITE_BASE_URL = (process.env.SHAREPOINT_SITE_BASE_URL || 'https://zakidatasas.sharepoint.com/sites/general').replace(/\/$/, '');
const SHAREPOINT_COURSES_SERVER_RELATIVE = process.env.SHAREPOINT_COURSES_SERVER_RELATIVE ||
  '/sites/general/Documentos compartidos/1. COMUNICACIONES/1.CURSOS';
const SHAREPOINT_DISCOVERY_DEPTH = Number(process.env.SHAREPOINT_DISCOVERY_DEPTH || 6);
const SHAREPOINT_DISCOVERY_LIMIT = Number(process.env.SHAREPOINT_DISCOVERY_LIMIT || 2000);
const SHAREPOINT_EDITED_FOLDER_DEPTH = Number(process.env.SHAREPOINT_EDITED_FOLDER_DEPTH || 6);

const notion = new Client({ auth: process.env.NOTION_API_KEY });
const collator = new Intl.Collator('es', { numeric: true, sensitivity: 'base' });

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function sanitizeFileName(value) {
  return String(value || 'documento')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120);
}

function getSharePointUser() {
  return process.env.SHAREPOINT_EMAIL || process.env.MICROSOFT_EMAIL || '';
}

function getSharePointPassword() {
  return process.env.SHAREPOINT_PASSWORD || process.env.MICROSOFT_PASSWORD || '';
}

function getTitle(page) {
  const titleProperty = Object.values(page.properties).find(prop => prop.type === 'title');
  return (titleProperty?.title || []).map(item => item.plain_text).join('').trim();
}

function isFilesPropertyEmpty(property) {
  if (!property || property.type !== 'files') {
    return false;
  }

  return property.files.length === 0;
}

function getFilesCount(property) {
  if (!property || property.type !== 'files') {
    return 0;
  }

  return property.files.length;
}

function escapeODataStringLiteral(value) {
  return String(value || '').replace(/'/g, "''");
}

function sharePointApiUrl(serverRelativeUrl, collection, select) {
  const quotedPath = `'${escapeODataStringLiteral(serverRelativeUrl)}'`;
  const encodedPath = encodeURIComponent(quotedPath);
  return `${SHAREPOINT_SITE_BASE_URL}/_api/web/GetFolderByServerRelativePath(decodedurl=@folder)/${collection}?@folder=${encodedPath}&$select=${encodeURIComponent(select)}`;
}

function streamUrlFor(serverRelativeUrl) {
  return `${SHAREPOINT_SITE_BASE_URL}/_layouts/15/stream.aspx?id=${encodeURIComponent(serverRelativeUrl)}&referrer=StreamWebApp.Web&referrerScenario=AddressBarCopied.view`;
}

function parseVtt(vttText) {
  return String(vttText || '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line =>
      line &&
      !/^WEBVTT/i.test(line) &&
      !/^NOTE\b/i.test(line) &&
      !/^\d+$/.test(line) &&
      !/-->/i.test(line)
    )
    .map(line => line.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseTranscriptTimeToSeconds(label) {
  const value = normalizeText(label);
  let total = 0;

  const hours = value.match(/(\d+)\s*hora/);
  if (hours) {
    total += Number(hours[1]) * 3600;
  }

  const minutes = value.match(/(\d+)\s*minuto/);
  if (minutes) {
    total += Number(minutes[1]) * 60;
  }

  const seconds = value.match(/(\d+)\s*segundo/);
  if (seconds) {
    total += Number(seconds[1]);
  }

  return total;
}

function groupLessonsByModule(lessons) {
  const groups = [];
  const indexesByModule = new Map();

  for (const lesson of lessons) {
    const moduleTitle = lesson.moduleTitle || 'Sin modulo';
    if (!indexesByModule.has(moduleTitle)) {
      indexesByModule.set(moduleTitle, groups.length);
      groups.push({ title: moduleTitle, lessons: [] });
    }

    groups[indexesByModule.get(moduleTitle)].lessons.push(lesson);
  }

  return groups;
}

async function getDataSourceId(databaseId) {
  const database = await notion.databases.retrieve({ database_id: databaseId });
  return database.data_sources?.[0]?.id || null;
}

async function findNotionCourse(courseName) {
  const courses = await listNotionCourses();
  return courses.find(course => normalizeText(course.title) === normalizeText(courseName)) || null;
}

async function listNotionCourses() {
  const databaseId = process.env.NOTION_DATABASE_ID;
  const dataSourceId = await getDataSourceId(databaseId);
  const courses = [];
  let cursor;

  do {
    const response = dataSourceId
      ? await notion.dataSources.query({
          data_source_id: dataSourceId,
          page_size: 100,
          start_cursor: cursor
        })
      : await notion.databases.query({
          database_id: databaseId,
          page_size: 100,
          start_cursor: cursor
        });

    for (const page of response.results) {
      if (page.archived || page.in_trash) {
        continue;
      }

      const title = getTitle(page);
      if (!title) {
        continue;
      }

      const apostillaProperty = page.properties[APOSTILLA_PROPERTY_NAME];
      const transcriptionProperty = page.properties[TRANSCRIPTION_PROPERTY_NAME];

      courses.push({
        id: page.id,
        title,
        url: page.url,
        apostillaEmpty: isFilesPropertyEmpty(apostillaProperty),
        apostillaFilesCount: getFilesCount(apostillaProperty),
        transcriptionEmpty: isFilesPropertyEmpty(transcriptionProperty),
        transcriptionFilesCount: getFilesCount(transcriptionProperty),
        transcriptionPropertyType: transcriptionProperty?.type || null
      });
    }

    cursor = response.has_more ? response.next_cursor : undefined;
  } while (cursor);

  return courses;
}

async function createBrowserContext(browser) {
  if (process.env.SHAREPOINT_STORAGE_STATE) {
    return browser.newContext({ acceptDownloads: true, storageState: process.env.SHAREPOINT_STORAGE_STATE });
  }

  if (process.env.SHAREPOINT_STORAGE_STATE_B64) {
    const storagePath = path.join(process.cwd(), '.sharepoint-storage-state.json');
    await fs.promises.writeFile(
      storagePath,
      Buffer.from(process.env.SHAREPOINT_STORAGE_STATE_B64, 'base64').toString('utf8')
    );
    return browser.newContext({ acceptDownloads: true, storageState: storagePath });
  }

  return browser.newContext({ acceptDownloads: true });
}

async function signInSharePoint(page) {
  await page.goto(SHAREPOINT_SITE_BASE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(2000);

  if (!/login\.microsoftonline\.com|login\.live\.com/i.test(page.url())) {
    return;
  }

  const email = getSharePointUser();
  const password = getSharePointPassword();

  if (!email || !password) {
    throw new Error('SharePoint pide inicio de sesion. Configura SHAREPOINT_EMAIL y SHAREPOINT_PASSWORD, o SHAREPOINT_STORAGE_STATE_B64 si la cuenta usa MFA.');
  }

  const emailInput = page.locator('input[type="email"], input[name="loginfmt"]').first();
  if (await emailInput.count()) {
    await emailInput.fill(email);
    await Promise.all([
      page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {}),
      page.keyboard.press('Enter')
    ]);
  }

  const passwordInput = page.locator('input[type="password"], input[name="passwd"]').first();
  await passwordInput.waitFor({ state: 'visible', timeout: 30000 });
  await passwordInput.fill(password);
  await Promise.all([
    page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => {}),
    page.keyboard.press('Enter')
  ]);

  const staySignedInNo = page.locator('input[type="button"][value="No"], button:has-text("No")').first();
  const staySignedInYes = page.locator('input[type="submit"][value="Sí"], input[type="submit"][value="Yes"], button:has-text("Sí"), button:has-text("Yes")').first();
  await Promise.race([
    staySignedInNo.waitFor({ state: 'visible', timeout: 8000 }).then(() => staySignedInNo.click()).catch(() => {}),
    staySignedInYes.waitFor({ state: 'visible', timeout: 8000 }).then(() => staySignedInYes.click()).catch(() => {}),
    page.waitForURL(url => !/login\.microsoftonline\.com|login\.live\.com/i.test(String(url)), { timeout: 8000 }).catch(() => {})
  ]);

  await page.goto(SHAREPOINT_SITE_BASE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(2000);

  if (/login\.microsoftonline\.com|login\.live\.com/i.test(page.url())) {
    throw new Error('No se pudo completar el login de SharePoint. Si hay MFA, usa SHAREPOINT_STORAGE_STATE_B64.');
  }
}

async function sharePointRestGet(page, url) {
  const acceptHeaders = [
    'application/json;odata=nometadata',
    'application/json;odata=minimalmetadata',
    'application/json;odata=verbose',
    'application/json'
  ];
  let lastError;

  for (const accept of acceptHeaders) {
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const response = await page.request.get(url, {
        headers: { Accept: accept },
        timeout: 60000
      });

      if (response.ok()) {
        const json = await response.json();
        if (json.d?.results) {
          return { value: json.d.results };
        }
        if (json.d) {
          return json.d;
        }
        return json;
      }

      const body = await response.text().catch(() => '');
      lastError = new Error(`SharePoint REST fallo ${response.status()} ${response.statusText()}: ${body.slice(0, 500)}`);

      if (![406, 429, 500, 502, 503, 504].includes(response.status())) {
        throw lastError;
      }

      await page.waitForTimeout(1000 * (attempt + 1));
    }
  }

  throw lastError;
}

async function listSharePointFolders(page, serverRelativeUrl) {
  const json = await sharePointRestGet(page, sharePointApiUrl(serverRelativeUrl, 'Folders', 'Name,ServerRelativeUrl'));
  return (json.value || [])
    .filter(item => item.Name && !item.Name.startsWith('_'))
    .map(item => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl
    }))
    .sort((a, b) => collator.compare(a.name, b.name));
}

async function listSharePointFiles(page, serverRelativeUrl) {
  const json = await sharePointRestGet(page, sharePointApiUrl(serverRelativeUrl, 'Files', 'Name,ServerRelativeUrl,Length,TimeLastModified'));
  return (json.value || [])
    .filter(item => item.Name)
    .map(item => ({
      name: item.Name,
      serverRelativeUrl: item.ServerRelativeUrl,
      length: Number(item.Length || 0),
      timeLastModified: item.TimeLastModified
    }))
    .sort((a, b) => collator.compare(a.name, b.name));
}

async function findSharePointCourseFolder(page, courseName) {
  const target = normalizeText(courseName);
  const queue = [{ name: path.posix.basename(SHAREPOINT_COURSES_SERVER_RELATIVE), serverRelativeUrl: SHAREPOINT_COURSES_SERVER_RELATIVE, depth: 0 }];
  const candidates = [];
  let visited = 0;

  while (queue.length > 0) {
    const current = queue.shift();
    visited += 1;

    if (visited > SHAREPOINT_DISCOVERY_LIMIT) {
      break;
    }

    const currentName = normalizeText(current.name);
    const fuzzyMatchAllowed = currentName.length >= 6 && target.length >= 6;
    if (current.depth > 0 && (
      currentName === target ||
      (fuzzyMatchAllowed && (currentName.includes(target) || target.includes(currentName)))
    )) {
      candidates.push(current);
    }

    if (current.depth >= SHAREPOINT_DISCOVERY_DEPTH) {
      continue;
    }

    const folders = await listSharePointFolders(page, current.serverRelativeUrl);
    for (const folder of folders) {
      queue.push({ ...folder, depth: current.depth + 1 });
    }
  }

  if (candidates.length === 0) {
    throw new Error(`No encontre una carpeta de SharePoint para el curso "${courseName}" dentro de ${SHAREPOINT_COURSES_SERVER_RELATIVE}.`);
  }

  candidates.sort((a, b) => {
    const aName = normalizeText(a.name);
    const bName = normalizeText(b.name);
    const aExact = aName === target ? 0 : 1;
    const bExact = bName === target ? 0 : 1;
    return aExact - bExact || a.depth - b.depth || collator.compare(a.name, b.name);
  });

  return candidates[0];
}

function getModuleTitleFromEditedFolder(courseFolder, editedFolder) {
  const relativePath = editedFolder.serverRelativeUrl
    .replace(courseFolder.serverRelativeUrl, '')
    .split('/')
    .map(part => decodeURIComponent(part).trim())
    .filter(Boolean);

  const editedIndex = relativePath.findIndex(part => ['editado', 'editados'].includes(normalizeText(part)));
  const parentParts = relativePath.slice(0, editedIndex > -1 ? editedIndex : relativePath.length);
  const moduleParts = parentParts.filter(part => !['videos', 'video'].includes(normalizeText(part)));

  return moduleParts.at(-1) || parentParts.at(-1) || 'Sin modulo';
}

async function findEditedFolders(page, courseFolder) {
  const editedFolders = [];
  const queue = [{ ...courseFolder, depth: 0 }];
  let visited = 0;

  while (queue.length > 0) {
    const current = queue.shift();
    visited += 1;

    if (visited > SHAREPOINT_DISCOVERY_LIMIT || current.depth > SHAREPOINT_EDITED_FOLDER_DEPTH) {
      continue;
    }

    const folders = await listSharePointFolders(page, current.serverRelativeUrl);
    for (const folder of folders) {
      const folderName = normalizeText(folder.name);

      if (['editado', 'editados'].includes(folderName)) {
        editedFolders.push({
          ...folder,
          moduleTitle: getModuleTitleFromEditedFolder(courseFolder, folder)
        });
        continue;
      }

      if (folderName.includes('material')) {
        continue;
      }

      queue.push({ ...folder, depth: current.depth + 1 });
    }
  }

  return editedFolders.sort((a, b) => collator.compare(a.serverRelativeUrl, b.serverRelativeUrl));
}

async function extractSharePointLessons(page, courseFolder) {
  const editedFolders = await findEditedFolders(page, courseFolder);
  const lessons = [];

  for (const editedFolder of editedFolders) {
    const files = await listSharePointFiles(page, editedFolder.serverRelativeUrl);
    const videoFiles = files.filter(file => VIDEO_FILE_PATTERN.test(file.name));

    for (const file of videoFiles) {
      lessons.push({
        moduleTitle: editedFolder.moduleTitle,
        title: file.name.replace(VIDEO_FILE_PATTERN, '').replace(/\s+/g, ' ').trim(),
        fileName: file.name,
        serverRelativeUrl: file.serverRelativeUrl,
        streamUrl: streamUrlFor(file.serverRelativeUrl),
        isVideo: true
      });
    }
  }

  if (lessons.length === 0) {
    throw new Error(`No encontre archivos de video dentro de carpetas "editado/editados" en ${courseFolder.serverRelativeUrl}. Carpetas encontradas: ${editedFolders.length}.`);
  }

  return lessons;
}

async function ensureTranscriptPanel(page) {
  for (let attempt = 0; attempt < 25; attempt += 1) {
    const status = await page.evaluate(() => {
      const body = document.body.innerText || '';
      const entries = Array.from(document.querySelectorAll('[role="group"] [role="listitem"], [role="group"] li'))
        .map(el => el.textContent?.trim())
        .filter(Boolean);

      return {
        hasEntries: entries.length > 0,
        hasTranscriptPanel: /Transcripci[oó]n|Transcript/i.test(body),
        hasReadButton: /Leer transcripci[oó]n|Transcript/i.test(body)
      };
    });

    if (status.hasEntries) {
      return;
    }

    if (status.hasReadButton) {
      const readButton = page.getByRole('button', { name: /Leer transcripci[oó]n|Transcript/i }).first();
      if (await readButton.count()) {
        await readButton.click({ timeout: 5000 }).catch(() => {});
      }
    }

    await page.waitForTimeout(1000);
  }
}

async function downloadTranscriptVtt(page, lesson, outputRoot, index) {
  await page.goto(lesson.streamUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await ensureTranscriptPanel(page);

  const panel = page.getByRole('complementary', { name: /Transcripci[oó]n|Transcript/i }).first();
  const downloadButton = panel.getByRole('button', { name: /Descargar|Download/i }).first();
  await downloadButton.waitFor({ state: 'visible', timeout: 15000 });
  await downloadButton.click();

  const vttMenuItem = page.getByRole('menuitem', { name: /\.vtt/i }).first();
  await vttMenuItem.waitFor({ state: 'visible', timeout: 10000 });

  const download = await Promise.all([
    page.waitForEvent('download', { timeout: 30000 }),
    vttMenuItem.click()
  ]).then(([downloadResult]) => downloadResult);

  const vttPath = path.join(outputRoot, `${String(index + 1).padStart(2, '0')}-${sanitizeFileName(lesson.title)}.vtt`);
  await download.saveAs(vttPath);

  const transcript = parseVtt(await fs.promises.readFile(vttPath, 'utf8'));
  if (!transcript) {
    throw new Error(`La transcripcion VTT descargada esta vacia para "${lesson.title}".`);
  }

  return transcript;
}

async function getVisibleTranscriptEntries(page) {
  return page.evaluate(() => Array.from(document.querySelectorAll('[role="group"]')).map(group => {
    const time = group.getAttribute('aria-label') || '';
    const item = group.querySelector('[role="listitem"], li');
    const text = item?.textContent?.replace(/\s+/g, ' ').trim() || '';
    return { time, text };
  }).filter(entry => entry.text && /minuto|segundo|hora|hour|minute|second/i.test(entry.time)));
}

async function collectTranscriptFromPanel(page) {
  const seen = new Map();
  let stagnant = 0;
  let previousMax = -1;
  const duration = await page.evaluate(() => {
    const video = document.querySelector('video');
    return video && Number.isFinite(video.duration) ? video.duration : null;
  });

  for (let step = 0; step < 120; step += 1) {
    const entries = await getVisibleTranscriptEntries(page);
    for (const entry of entries) {
      const seconds = parseTranscriptTimeToSeconds(entry.time);
      const key = `${seconds}|${entry.text.slice(0, 100)}`;
      if (!seen.has(key)) {
        seen.set(key, { ...entry, seconds });
      }
    }

    const maxSeconds = Math.max(-1, ...Array.from(seen.values()).map(entry => entry.seconds));
    if (duration && maxSeconds >= duration - 20) {
      break;
    }

    if (maxSeconds <= previousMax) {
      stagnant += 1;
    } else {
      stagnant = 0;
    }

    if (stagnant >= 8) {
      break;
    }

    previousMax = Math.max(previousMax, maxSeconds);

    const scrolled = await page.evaluate(() => {
      const scroller = Array.from(document.querySelectorAll('*'))
        .find(el => el.scrollHeight > el.clientHeight + 200 && /Generado por Microsoft|Generated by Microsoft/i.test(el.textContent || ''));

      if (!scroller) {
        return false;
      }

      scroller.scrollTop += 900;
      return true;
    }).catch(() => false);

    if (!scrolled) {
      await page.mouse.wheel(0, 900);
    }

    await page.waitForTimeout(300);
  }

  const ordered = Array.from(seen.values()).sort((a, b) => a.seconds - b.seconds);
  return ordered.map(entry => entry.text).join(' ').replace(/\s+/g, ' ').trim();
}

async function extractTranscriptFromStreamPanel(page, lesson) {
  await page.goto(lesson.streamUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await ensureTranscriptPanel(page);

  const transcript = await collectTranscriptFromPanel(page);
  if (!transcript) {
    throw new Error(`No pude leer la transcripcion del panel de Stream para "${lesson.title}".`);
  }

  return transcript;
}

async function extractSharePointTranscript(page, lesson, outputRoot, index) {
  try {
    return await downloadTranscriptVtt(page, lesson, outputRoot, index);
  } catch (error) {
    console.log(`  No pude descargar VTT (${error.message}). Intentare leer el panel de Stream.`);
    return extractTranscriptFromStreamPanel(page, lesson);
  }
}

async function uploadTranscriptionToNotion(pageId, filePath) {
  const fileName = path.basename(filePath);
  const fileBlob = new Blob([await fs.promises.readFile(filePath)], { type: DOCX_CONTENT_TYPE });

  console.log(`Creando carga de archivo en Notion: ${fileName}`);
  const upload = await notion.fileUploads.create({
    mode: 'single_part',
    filename: fileName,
    content_type: DOCX_CONTENT_TYPE
  });

  const sentUpload = await notion.fileUploads.send({
    file_upload_id: upload.id,
    file: {
      filename: fileName,
      data: fileBlob
    }
  });

  if (sentUpload.status !== 'uploaded') {
    throw new Error(`Notion no marco el archivo como uploaded. Estado recibido: ${sentUpload.status}`);
  }

  await notion.pages.update({
    page_id: pageId,
    properties: {
      [TRANSCRIPTION_PROPERTY_NAME]: {
        type: 'files',
        files: [
          {
            type: 'file_upload',
            file_upload: {
              id: upload.id
            }
          }
        ]
      }
    }
  });

  return upload.id;
}

function createTranscriptDocument(courseName, lessons, outputPath) {
  const children = [
    new Paragraph({
      children: [
        new TextRun({
          text: `Transcripcion del curso: ${courseName}`,
          size: 52,
          font: 'Arial'
        })
      ],
      spacing: { after: 120 }
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Generado el ${new Date().toLocaleString('es-CO', { timeZone: 'America/Bogota' })} desde transcripciones automaticas de Microsoft Stream.`,
          italics: true,
          size: 22,
          font: 'Arial'
        })
      ],
      spacing: { after: 320 }
    })
  ];

  const modules = groupLessonsByModule(lessons);
  let globalLessonIndex = 0;

  for (const [moduleIndex, module] of modules.entries()) {
    const moduleTitle = /^modulo|^m[oó]dulo/i.test(module.title)
      ? module.title
      : `Modulo ${moduleIndex + 1}: ${module.title}`;

    children.push(new Paragraph({
      text: moduleTitle,
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 360, after: 160 }
    }));

    for (const lesson of module.lessons) {
      globalLessonIndex += 1;
      children.push(new Paragraph({
        text: `${globalLessonIndex}. ${lesson.title}`,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 220, after: 120 }
      }));

      if (!lesson.transcript) {
        children.push(new Paragraph({
          children: [new TextRun({ text: 'No se pudo obtener transcripcion automatica para esta clase.', italics: true })],
          spacing: { after: 160 }
        }));
        continue;
      }

      const chunks = lesson.transcript.match(/.{1,1800}(?:\s|$)/g) || [lesson.transcript];
      for (const chunk of chunks) {
        children.push(new Paragraph({
          children: [new TextRun({ text: chunk.trim(), font: 'Arial', size: 22 })],
          alignment: AlignmentType.LEFT,
          spacing: { after: 160 },
          lineSpacing: 276
        }));
      }
    }
  }

  const document = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440
          }
        }
      },
      children
    }]
  });

  return Packer.toBuffer(document).then(buffer => fs.promises.writeFile(outputPath, buffer));
}

async function cleanupIntermediateFiles(outputRoot, outputDocx) {
  const removableExtensions = new Set(['.wav', '.mp3', '.txt', '.json', '.vtt']);
  const entries = await fs.promises.readdir(outputRoot, { withFileTypes: true });
  let removedCount = 0;
  let removedBytes = 0;

  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }

    const filePath = path.join(outputRoot, entry.name);
    if (path.resolve(filePath) === path.resolve(outputDocx)) {
      continue;
    }

    const shouldRemove = removableExtensions.has(path.extname(entry.name).toLowerCase());
    if (!shouldRemove) {
      continue;
    }

    const stat = await fs.promises.stat(filePath);
    await fs.promises.unlink(filePath);
    removedCount += 1;
    removedBytes += stat.size;
  }

  console.log(`Limpieza local completada: ${removedCount} archivos eliminados (${(removedBytes / 1024 / 1024).toFixed(1)} MB).`);
}

function validateEnv() {
  const required = ['NOTION_API_KEY', 'NOTION_DATABASE_ID'];
  const missing = required.filter(name => !process.env[name]);
  const hasSharePointLogin = Boolean(getSharePointUser() && getSharePointPassword());
  const hasSharePointStorage = Boolean(process.env.SHAREPOINT_STORAGE_STATE || process.env.SHAREPOINT_STORAGE_STATE_B64);

  if (!hasSharePointLogin && !hasSharePointStorage) {
    missing.push('SHAREPOINT_EMAIL/SHAREPOINT_PASSWORD o SHAREPOINT_STORAGE_STATE_B64');
  }

  if (missing.length > 0) {
    throw new Error(`Faltan variables de entorno requeridas: ${missing.join(', ')}`);
  }
}

async function processNotionCourse(page, notionCourse, options) {
  const {
    metadataOnly,
    overwrite,
    keepIntermediate,
    uploadNotion,
    lessonLimit
  } = options;

  const courseName = notionCourse.title;
  const outputRoot = path.join(process.cwd(), 'transcriptions', sanitizeFileName(courseName));
  fs.mkdirSync(outputRoot, { recursive: true });

  console.log(`\n=== ${courseName} ===`);
  console.log(`Notion: ${notionCourse.title}`);
  console.log(`Apostilla vacia: ${notionCourse.apostillaEmpty ? 'si' : 'no'}`);
  console.log(`Transcripcion vacia: ${notionCourse.transcriptionEmpty ? 'si' : 'no'}`);
  console.log(`Propiedad objetivo: ${TRANSCRIPTION_PROPERTY_NAME} (${notionCourse.transcriptionPropertyType || 'no encontrada'})`);

  if (!notionCourse.apostillaEmpty) {
    console.log('Omitido: Apostilla no esta vacia.');
    return { status: 'skipped', reason: 'apostilla_not_empty' };
  }

  if (notionCourse.transcriptionPropertyType !== 'files') {
    throw new Error(`La propiedad "${TRANSCRIPTION_PROPERTY_NAME}" debe ser de tipo files.`);
  }

  let courseFolder;
  try {
    courseFolder = await findSharePointCourseFolder(page, notionCourse.title);
  } catch (error) {
    console.log(`Omitido: no encontre carpeta de SharePoint para este curso (${error.message || error}).`);
    return { status: 'skipped', reason: 'sharepoint_folder_not_found' };
  }
  console.log(`Carpeta SharePoint: ${courseFolder.serverRelativeUrl}`);

  let lessons;
  try {
    lessons = await extractSharePointLessons(page, courseFolder);
  } catch (error) {
    console.log(`Omitido: no encontre videos editados para este curso (${error.message || error}).`);
    return { status: 'skipped', reason: 'sharepoint_videos_not_found' };
  }
  if (lessonLimit > 0) {
    lessons = lessons.slice(0, lessonLimit);
  }

  console.log(`Clases editadas encontradas: ${lessons.length}`);

  const metadataPath = path.join(outputRoot, 'metadata.json');
  await fs.promises.writeFile(metadataPath, JSON.stringify({
    course: notionCourse,
    sharePointCourseFolder: courseFolder,
    lessons: lessons.map(lesson => ({
      moduleTitle: lesson.moduleTitle,
      title: lesson.title,
      fileName: lesson.fileName,
      streamUrl: lesson.streamUrl,
      serverRelativeUrl: lesson.serverRelativeUrl
    }))
  }, null, 2));
  console.log(`Metadata local guardada: ${metadataPath}`);

  if (metadataOnly) {
    return { status: 'metadata', lessons: lessons.length };
  }

  for (const [index, lesson] of lessons.entries()) {
    const baseName = `${String(index + 1).padStart(2, '0')}-${sanitizeFileName(lesson.title)}`;
    const transcriptPath = path.join(outputRoot, `${baseName}.sharepoint.txt`);

    if (!overwrite && fs.existsSync(transcriptPath)) {
      lesson.transcript = await fs.promises.readFile(transcriptPath, 'utf8');
      console.log(`[${index + 1}/${lessons.length}] Transcripcion reutilizada: ${lesson.title}`);
      continue;
    }

    console.log(`[${index + 1}/${lessons.length}] Transcripcion SharePoint/Stream: ${lesson.title}`);
    try {
      lesson.transcript = await extractSharePointTranscript(page, lesson, outputRoot, index);
      await fs.promises.writeFile(transcriptPath, lesson.transcript, 'utf8');
    } catch (error) {
      lesson.transcript = '';
      lesson.transcriptError = error.message || String(error);
      console.warn(`  Sin transcripcion para esta clase: ${lesson.transcriptError}`);
      await fs.promises.writeFile(`${transcriptPath}.error.txt`, lesson.transcriptError, 'utf8');
    }
  }

  const outputDocx = path.join(outputRoot, `${sanitizeFileName(courseName)} - transcripcion.docx`);
  await createTranscriptDocument(courseName, lessons, outputDocx);
  console.log(`Documento Word generado localmente: ${outputDocx}`);

  let fileUploadId = null;
  if (uploadNotion) {
    fileUploadId = await uploadTranscriptionToNotion(notionCourse.id, outputDocx);
    console.log(`Documento cargado en Notion (${TRANSCRIPTION_PROPERTY_NAME}): ${fileUploadId}`);
  } else {
    console.log('No se subio nada a Notion. Usa --upload-notion para cargar el Word final.');
  }

  if (!keepIntermediate) {
    await cleanupIntermediateFiles(outputRoot, outputDocx);
  }

  return {
    status: 'processed',
    fileUploadId,
    lessons: lessons.length,
    videoLessons: lessons.length,
    outputDocx
  };
}

async function main() {
  const args = process.argv.slice(2);
  const allCourses = args.includes('--all');
  const metadataOnly = args.includes('--metadata-only');
  const overwrite = args.includes('--overwrite');
  const keepIntermediate = args.includes('--keep-intermediate');
  const uploadNotion = args.includes('--upload-notion');
  const includeExistingTranscriptions = args.includes('--include-existing-transcriptions');
  const limitArg = args.find(arg => arg.startsWith('--limit='));
  const courseLimitArg = args.find(arg => arg.startsWith('--course-limit='));
  const lessonLimit = limitArg ? Number(limitArg.split('=')[1]) : 0;
  const courseLimit = courseLimitArg ? Number(courseLimitArg.split('=')[1]) : 0;
  const courseName = args.filter(arg => !arg.startsWith('--')).join(' ') || 'Liderazgo Personal';

  validateEnv();

  let targetCourses;

  if (allCourses) {
    const allNotionCourses = await listNotionCourses();
    const emptyApostillaCourses = allNotionCourses.filter(course => course.apostillaEmpty);
    targetCourses = emptyApostillaCourses.filter(course =>
      overwrite || includeExistingTranscriptions || course.transcriptionEmpty
    );

    if (courseLimit > 0) {
      targetCourses = targetCourses.slice(0, courseLimit);
    }

    console.log(`Cursos en Notion: ${allNotionCourses.length}`);
    console.log(`Cursos con Apostilla vacia: ${emptyApostillaCourses.length}`);
    console.log(`Cursos pendientes de Transcripcion: ${targetCourses.length}`);
  } else {
    console.log(`Buscando curso en Notion: ${courseName}`);
    const notionCourse = await findNotionCourse(courseName);
    if (!notionCourse) {
      throw new Error(`No encontre "${courseName}" en la base de Notion.`);
    }
    targetCourses = [notionCourse];
  }

  if (targetCourses.length === 0) {
    console.log('No hay cursos pendientes para transcribir.');
    return;
  }

  const browser = await chromium.launch({
    headless: process.env.PLAYWRIGHT_HEADFUL !== 'true'
  });
  const context = await createBrowserContext(browser);
  const page = await context.newPage();
  const results = [];

  try {
    console.log('Entrando a SharePoint/Stream...');
    await signInSharePoint(page);

    for (const [index, notionCourse] of targetCourses.entries()) {
      console.log(`\nCurso ${index + 1}/${targetCourses.length}`);
      try {
        const result = await processNotionCourse(page, notionCourse, {
          metadataOnly,
          overwrite,
          keepIntermediate,
          uploadNotion,
          lessonLimit
        });
        results.push({ course: notionCourse.title, ...result });
      } catch (error) {
        console.error(`Error procesando "${notionCourse.title}":`, error.message || error);
        results.push({ course: notionCourse.title, status: 'failed', error: error.message || String(error) });
      }
    }
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }

  const processed = results.filter(result => result.status === 'processed').length;
  const skipped = results.filter(result => result.status === 'skipped').length;
  const failed = results.filter(result => result.status === 'failed');

  console.log('\nResumen de transcripcion');
  console.log(`Procesados: ${processed}`);
  console.log(`Omitidos: ${skipped}`);
  console.log(`Fallidos: ${failed.length}`);

  if (failed.length > 0) {
    for (const failure of failed) {
      console.log(`- ${failure.course}: ${failure.error}`);
    }
    throw new Error(`Fallaron ${failed.length} cursos durante la transcripcion.`);
  }
}

if (require.main === module) {
  main().catch(error => {
    console.error('Error generando transcripcion:', error);
    process.exit(1);
  });
}
