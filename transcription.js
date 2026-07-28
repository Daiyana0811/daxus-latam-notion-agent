require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('playwright');
const { Client } = require('@notionhq/client');
const OpenAI = require('openai');
const ffmpegPath = require('ffmpeg-static');
const {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  TextRun,
  AlignmentType
} = require('docx');

const LOGIN_URL = 'https://miembro.daxus.com/users/sign_in';
const COURSE_LIST_URL = 'https://miembro.daxus.com/?browse=available';
const TRANSCRIPTION_PROPERTY_NAME = 'Transcripcion';
const APOSTILLA_PROPERTY_NAME = 'Apostilla';
const OPENAI_TRANSCRIPTION_MODEL = process.env.OPENAI_TRANSCRIPTION_MODEL || 'gpt-4o-transcribe';

const notion = new Client({ auth: process.env.NOTION_API_KEY });

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

async function getDataSourceId(databaseId) {
  const database = await notion.databases.retrieve({ database_id: databaseId });
  return database.data_sources?.[0]?.id || null;
}

async function findNotionCourse(courseName) {
  const databaseId = process.env.NOTION_DATABASE_ID;
  const dataSourceId = await getDataSourceId(databaseId);
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
      if (normalizeText(title) === normalizeText(courseName)) {
        return {
          id: page.id,
          title,
          url: page.url,
          apostillaEmpty: isFilesPropertyEmpty(page.properties[APOSTILLA_PROPERTY_NAME]),
          transcriptionPropertyType: page.properties[TRANSCRIPTION_PROPERTY_NAME]?.type || null
        };
      }
    }

    cursor = response.has_more ? response.next_cursor : undefined;
  } while (cursor);

  return null;
}

async function signIn(page) {
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.fill('input[type="email"]', process.env.DAXUS_EMAIL);
  await page.fill('input[type="password"]', process.env.DAXUS_PASSWORD);

  await Promise.all([
    page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(() => {}),
    page.keyboard.press('Enter')
  ]);

  if (page.url().includes('/onboarding')) {
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(() => {}),
      page.getByRole('button', { name: 'Estoy de acuerdo', exact: true }).click()
    ]);
  }
}

async function findDaxusCourseUrl(page, courseName) {
  await page.goto(COURSE_LIST_URL, { waitUntil: 'networkidle', timeout: 30000 });
  const target = normalizeText(courseName);

  const links = await page.evaluate(() => Array.from(document.querySelectorAll('a')).map(link => ({
    href: link.href,
    text: (link.innerText || link.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim(),
    imageAlt: link.querySelector('img')?.getAttribute('alt') || '',
    hasImage: Boolean(link.querySelector('img'))
  })));

  const match = links.find(link => {
    const combined = normalizeText(`${link.text} ${link.imageAlt} ${link.href}`);
    return link.hasImage && combined.includes(target.replace(/\s+/g, '-'));
  }) || links.find(link => normalizeText(`${link.text} ${link.imageAlt} ${link.href}`).includes(target));

  if (!match) {
    throw new Error(`No encontre el curso "${courseName}" en Daxus.`);
  }

  return match.href;
}

async function extractLessons(page, courseUrl) {
  await page.goto(courseUrl, { waitUntil: 'networkidle', timeout: 30000 });

  return page.evaluate(() => {
    const lessons = [];
    const headings = document.querySelectorAll('h4');

    for (const heading of headings) {
      const spans = heading.querySelectorAll('span');
      let moduleTitle = '';

      for (const span of spans) {
        const text = span.innerText.trim();
        if (text && !span.querySelector('svg')) {
          moduleTitle = text;
          break;
        }
      }

      if (!moduleTitle) {
        moduleTitle = heading.innerText.trim();
      }

      let moduleNumber = '';
      const parentDiv = heading.closest('div');
      if (parentDiv) {
        const descriptionNode = parentDiv.querySelector('p.section__description');
        if (descriptionNode) {
          moduleNumber = descriptionNode.innerText.trim();
        }
      }

      const fullModuleTitle = moduleNumber ? `${moduleNumber} - ${moduleTitle}` : moduleTitle;
      let sectionDiv = null;
      let element = heading.closest('div[class]');

      while (element) {
        const next = element.nextElementSibling;
        if (next && next.id && next.id.startsWith('inner_section')) {
          sectionDiv = next;
          break;
        }

        element = element.parentElement;
        if (!element || element.tagName.toLowerCase() === 'body') {
          break;
        }
      }

      if (!sectionDiv) {
        continue;
      }

      for (const link of sectionDiv.querySelectorAll('a.lesson__title')) {
        const title = (link.innerText || '').replace(/\s+/g, ' ').trim();
        if (title && link.href) {
          lessons.push({
            moduleTitle: fullModuleTitle || 'Sin modulo',
            title,
            url: link.href
          });
        }
      }
    }

    if (lessons.length > 0) {
      return lessons;
    }

    return Array.from(document.querySelectorAll('a.lesson__title')).map(link => ({
      moduleTitle: 'Sin modulo',
      title: (link.innerText || '').replace(/\s+/g, ' ').trim(),
      url: link.href
    })).filter(lesson => lesson.title && lesson.url);
  });
}

async function extractVideoPlaylist(page, lesson) {
  const mediaRequests = [];
  const onRequest = request => {
    const url = request.url();
    if (/\.m3u8(\?|$)/i.test(url)) {
      mediaRequests.push(url);
    }
  };

  page.on('request', onRequest);
  try {
    await page.goto(lesson.url, { waitUntil: 'networkidle', timeout: 45000 });
    await page.waitForTimeout(5000);

    const pageData = await page.evaluate(() => ({
      title: document.querySelector('h1,h2,h3')?.innerText?.replace(/\s+/g, ' ').trim() || '',
      iframeSrc: document.querySelector('iframe[src*="pandavideo"]')?.src || ''
    }));

    const uniqueRequests = Array.from(new Set(mediaRequests));
    const selectedPlaylist =
      uniqueRequests.find(url => /360p\/video\.m3u8$/i.test(url)) ||
      uniqueRequests.find(url => /480p\/video\.m3u8$/i.test(url)) ||
      uniqueRequests.find(url => /720p\/video\.m3u8$/i.test(url)) ||
      uniqueRequests.find(url => /1080p\/video\.m3u8$/i.test(url)) ||
      uniqueRequests.find(url => /playlist\.m3u8$/i.test(url)) ||
      uniqueRequests.find(url => /360p\/video\.m3u8/i.test(url) && !/token=/i.test(url)) ||
      uniqueRequests.find(url => /video\.m3u8/i.test(url) && !/token=/i.test(url)) ||
      '';

    return {
      ...lesson,
      pageTitle: pageData.title || lesson.title,
      iframeSrc: pageData.iframeSrc,
      playlistUrl: selectedPlaylist,
      isVideo: Boolean(pageData.iframeSrc && selectedPlaylist)
    };
  } finally {
    page.off('request', onRequest);
  }
}

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      ...options
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', chunk => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', chunk => {
      stderr += chunk.toString();
    });

    child.on('error', reject);
    child.on('close', code => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      reject(new Error(`${command} exited with code ${code}\n${stderr || stdout}`));
    });
  });
}

async function extractAudio(lesson, outputPath) {
  const headers = [
    'Referer: https://miembro.daxus.com/',
    'User-Agent: Mozilla/5.0'
  ].join('\r\n');

  await runProcess(ffmpegPath, [
    '-y',
    '-headers',
    headers,
    '-i',
    lesson.playlistUrl,
    '-vn',
    '-ac',
    '1',
    '-ar',
    '16000',
    '-acodec',
    'libmp3lame',
    '-b:a',
    '64k',
    outputPath
  ]);
}

function createOpenAiTranscriber() {
  const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  return async audioPath => {
    const result = await client.audio.transcriptions.create({
      file: fs.createReadStream(audioPath),
      model: OPENAI_TRANSCRIPTION_MODEL,
      language: 'es'
    });

    return String(result.text || '').replace(/\s+/g, ' ').trim();
  };
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
          text: `Generado localmente el ${new Date().toLocaleString('es-CO', { timeZone: 'America/Bogota' })}. No subido a Notion.`,
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
    children.push(new Paragraph({
      text: `Modulo ${moduleIndex + 1}: ${module.title}`,
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

      if (!lesson.isVideo) {
        children.push(new Paragraph({
          children: [new TextRun({ text: 'Contenido sin video detectado; se omitio la transcripcion.', italics: true })],
          spacing: { after: 160 }
        }));
        continue;
      }

      if (!lesson.transcript) {
        children.push(new Paragraph({
          children: [new TextRun({ text: 'No se pudo generar transcripcion para esta leccion.', italics: true })],
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
  const removableExtensions = new Set(['.wav', '.mp3', '.txt', '.json']);
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

function validateEnv(metadataOnly) {
  const required = ['DAXUS_EMAIL', 'DAXUS_PASSWORD', 'NOTION_API_KEY', 'NOTION_DATABASE_ID'];
  const missing = required.filter(name => !process.env[name]);

  if (!metadataOnly && !process.env.OPENAI_API_KEY) {
    missing.push('OPENAI_API_KEY');
  }

  if (missing.length > 0) {
    throw new Error(`Faltan variables de entorno requeridas: ${missing.join(', ')}`);
  }
}

async function main() {
  const args = process.argv.slice(2);
  const metadataOnly = args.includes('--metadata-only');
  const overwrite = args.includes('--overwrite');
  const keepIntermediate = args.includes('--keep-intermediate');
  const limitArg = args.find(arg => arg.startsWith('--limit='));
  const limit = limitArg ? Number(limitArg.split('=')[1]) : 0;
  const courseName = args.filter(arg => !arg.startsWith('--'))[0] || 'Liderazgo Personal';

  validateEnv(metadataOnly);

  const outputRoot = path.join(process.cwd(), 'transcriptions', sanitizeFileName(courseName));
  fs.mkdirSync(outputRoot, { recursive: true });

  console.log(`Buscando curso en Notion: ${courseName}`);
  const notionCourse = await findNotionCourse(courseName);
  if (!notionCourse) {
    throw new Error(`No encontre "${courseName}" en la base de Notion.`);
  }

  console.log(`Notion: ${notionCourse.title}`);
  console.log(`Apostilla vacia: ${notionCourse.apostillaEmpty ? 'si' : 'no'}`);
  console.log(`Propiedad local objetivo: ${TRANSCRIPTION_PROPERTY_NAME} (${notionCourse.transcriptionPropertyType || 'no encontrada'})`);

  if (!notionCourse.apostillaEmpty) {
    console.log('La propiedad Apostilla no esta vacia. Para esta prueba no se generara transcripcion.');
    return;
  }

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    console.log('Entrando a Daxus LATAM...');
    await signIn(page);

    const courseUrl = await findDaxusCourseUrl(page, notionCourse.title);
    console.log(`Curso Daxus: ${courseUrl}`);

    let lessons = await extractLessons(page, courseUrl);
    if (limit > 0) {
      lessons = lessons.slice(0, limit);
    }

    console.log(`Lecciones encontradas para la prueba: ${lessons.length}`);
    const videoLessons = [];

    for (const [index, lesson] of lessons.entries()) {
      console.log(`[${index + 1}/${lessons.length}] Detectando video: ${lesson.title}`);
      const videoLesson = await extractVideoPlaylist(page, lesson);
      videoLessons.push(videoLesson);
      console.log(videoLesson.isVideo ? '  video detectado' : '  sin video detectado');
    }

    const metadataPath = path.join(outputRoot, 'metadata.json');
    await fs.promises.writeFile(metadataPath, JSON.stringify({
      course: notionCourse,
      daxusCourseUrl: courseUrl,
      lessons: videoLessons.map(lesson => ({
        moduleTitle: lesson.moduleTitle,
        title: lesson.title,
        url: lesson.url,
        isVideo: lesson.isVideo,
        iframeSrc: lesson.iframeSrc,
        hasPlaylist: Boolean(lesson.playlistUrl)
      }))
    }, null, 2));
    console.log(`Metadata local guardada: ${metadataPath}`);

    if (metadataOnly) {
      return;
    }

    let transcriber;
    const needsTranscription = videoLessons.some((lesson, index) => {
      if (!lesson.isVideo) {
        return false;
      }

      const baseName = `${String(index + 1).padStart(2, '0')}-${sanitizeFileName(lesson.title)}`;
      const transcriptPath = path.join(outputRoot, `${baseName}.openai.txt`);
      return overwrite || !fs.existsSync(transcriptPath);
    });

    if (needsTranscription) {
      console.log(`Proveedor de transcripcion: OpenAI (${OPENAI_TRANSCRIPTION_MODEL})`);
      transcriber = createOpenAiTranscriber();
    } else {
      console.log('Todas las transcripciones temporales ya existen. Regenerare solo el Word final.');
    }

    for (const [index, lesson] of videoLessons.entries()) {
      if (!lesson.isVideo) {
        continue;
      }

      const baseName = `${String(index + 1).padStart(2, '0')}-${sanitizeFileName(lesson.title)}`;
      const audioPath = path.join(outputRoot, `${baseName}.mp3`);
      const transcriptPath = path.join(outputRoot, `${baseName}.openai.txt`);

      if (overwrite || !fs.existsSync(audioPath)) {
        console.log(`[${index + 1}/${videoLessons.length}] Extrayendo audio: ${lesson.title}`);
        await extractAudio(lesson, audioPath);
      }

      if (!overwrite && fs.existsSync(transcriptPath)) {
        lesson.transcript = await fs.promises.readFile(transcriptPath, 'utf8');
      } else {
        console.log(`[${index + 1}/${videoLessons.length}] Transcribiendo: ${lesson.title}`);
        lesson.transcript = await transcriber(audioPath);
        await fs.promises.writeFile(transcriptPath, lesson.transcript, 'utf8');
      }
    }

    const outputDocx = path.join(outputRoot, `${sanitizeFileName(courseName)} - transcripcion.docx`);
    await createTranscriptDocument(courseName, videoLessons, outputDocx);
    console.log(`Documento Word generado localmente: ${outputDocx}`);
    if (!keepIntermediate) {
      await cleanupIntermediateFiles(outputRoot, outputDocx);
    }
    console.log('No se subio nada a Notion.');
  } finally {
    await browser.close().catch(() => {});
  }
}

if (require.main === module) {
  main().catch(error => {
    console.error('Error generando transcripcion local:', error);
    process.exit(1);
  });
}
