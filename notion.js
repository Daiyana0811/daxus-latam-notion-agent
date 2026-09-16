require('dotenv').config();

const { Client } = require('@notionhq/client');

const notionApiKey = process.env.NOTION_API_KEY;
const databaseId = process.env.NOTION_DATABASE_ID;
const CATEGORY_PROPERTY_NAME = 'Categoria';
const MEDIA_PROPERTY_NAME = 'Archivos y multimedia';
const COVER_FILE_NAME = 'Portada';

if (!notionApiKey) {
  throw new Error('Missing required environment variable: NOTION_API_KEY');
}

if (!databaseId) {
  throw new Error('Missing required environment variable: NOTION_DATABASE_ID');
}

const notion = new Client({ auth: notionApiKey });
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
let cachedDataSourceId;

function sanitizeFileName(value) {
  return String(value || 'cover')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'cover';
}

function extensionFromContentType(contentType) {
  const normalized = String(contentType || '').split(';')[0].trim().toLowerCase();

  if (normalized === 'image/png') return '.png';
  if (normalized === 'image/webp') return '.webp';
  if (normalized === 'image/gif') return '.gif';
  if (normalized === 'image/svg+xml') return '.svg';

  return '.jpg';
}

async function downloadCoverImage(course) {
  const response = await fetch(course.coverUrl);

  if (!response.ok) {
    throw new Error(`HTTP ${response.status} descargando portada`);
  }

  const contentType = response.headers.get('content-type') || 'image/jpeg';

  if (!contentType.toLowerCase().startsWith('image/')) {
    throw new Error(`La portada no parece ser imagen (${contentType})`);
  }

  const buffer = Buffer.from(await response.arrayBuffer());
  const filename = `${sanitizeFileName(course.title)} - portada${extensionFromContentType(contentType)}`;

  return { buffer, contentType, filename };
}

async function uploadImageBuffer({ buffer, contentType, filename }) {
  const fileBlob = new Blob([buffer], { type: contentType });
  const upload = await notion.fileUploads.create({
    mode: 'single_part',
    filename,
    content_type: contentType
  });

  const sentUpload = await notion.fileUploads.send({
    file_upload_id: upload.id,
    file: {
      filename,
      data: fileBlob
    }
  });

  if (sentUpload.status !== 'uploaded') {
    throw new Error(`Notion no marco la portada como uploaded. Estado recibido: ${sentUpload.status}`);
  }

  return upload.id;
}

async function uploadCoverImage(course, includePageCover) {
  if (!course.coverUrl || !course.coverUrl.startsWith('http')) {
    return null;
  }

  try {
    const image = await downloadCoverImage(course);
    const mediaUploadId = await uploadImageBuffer(image);
    const coverUploadId = includePageCover ? await uploadImageBuffer(image) : null;

    return { mediaUploadId, coverUploadId };
  } catch (error) {
    console.warn(`No pude subir portada como archivo para "${course.title}". Uso URL externa como respaldo: ${error.message || error}`);
    return null;
  }
}

function normalizeSelectName(value) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  return text.length > 100 ? text.slice(0, 100).trim() : text;
}

function getCategoryOptions(courses) {
  return Array.from(new Set(courses.map(course => normalizeSelectName(course.category)).filter(Boolean)))
    .map(name => ({ name, color: 'default' }));
}

async function getDataSourceId() {
  if (cachedDataSourceId !== undefined) {
    return cachedDataSourceId;
  }

  const database = await notion.databases.retrieve({ database_id: databaseId });

  if (database.data_sources && database.data_sources.length > 0) {
    cachedDataSourceId = database.data_sources[0].id;
    return cachedDataSourceId;
  }

  cachedDataSourceId = null;
  return cachedDataSourceId;
}

async function retrieveSchema() {
  const dataSourceId = await getDataSourceId();

  if (dataSourceId) {
    return notion.dataSources.retrieve({ data_source_id: dataSourceId });
  }

  return notion.databases.retrieve({ database_id: databaseId });
}

async function updateSchemaProperty(propertyConfig) {
  const dataSourceId = await getDataSourceId();

  if (dataSourceId) {
    await notion.dataSources.update({
      data_source_id: dataSourceId,
      properties: propertyConfig
    });
    return;
  }

  await notion.databases.update({
    database_id: databaseId,
    properties: propertyConfig
  });
}

async function ensureCategoryProperty(courses) {
  const database = await retrieveSchema();
  const existingProperty = database.properties[CATEGORY_PROPERTY_NAME];

  if (existingProperty) {
    if (existingProperty.type !== 'select') {
      throw new Error(`The Notion property "${CATEGORY_PROPERTY_NAME}" must be a select property.`);
    }

    return;
  }

  const options = getCategoryOptions(courses);
  await updateSchemaProperty({
    [CATEGORY_PROPERTY_NAME]: {
      select: options.length > 0 ? { options } : {}
    }
  });

  console.log(`Created Notion select property: ${CATEGORY_PROPERTY_NAME}`);
}

function getPageTitle(page) {
  const titlePropKey = Object.keys(page.properties).find(key => page.properties[key].type === 'title');
  const title = titlePropKey && page.properties[titlePropKey].title[0];
  return title ? title.plain_text.trim() : '';
}

async function listExistingCourses() {
  const dataSourceId = await getDataSourceId();

  if (dataSourceId) {
    return listExistingCoursesFromDataSource(dataSourceId);
  }

  return listExistingCoursesFromDatabaseSearch();
}

async function listExistingCoursesFromDataSource(dataSourceId) {
  const existingCourses = {};
  let cursor;

  do {
    const queryResults = await notion.dataSources.query({
      data_source_id: dataSourceId,
      page_size: 100,
      start_cursor: cursor
    });

    for (const page of queryResults.results) {
      if (page.archived || page.in_trash) {
        continue;
      }

      const title = getPageTitle(page);

      if (title) {
        existingCourses[title] = {
          id: page.id,
          mediaFiles: getMediaFiles(page)
        };
      }
    }

    cursor = queryResults.has_more ? queryResults.next_cursor : undefined;
  } while (cursor);

  return existingCourses;
}

async function listExistingCoursesFromDatabaseSearch() {
  const existingCourses = {};
  let cursor;

  do {
    const searchResults = await notion.search({
      filter: { property: 'object', value: 'page' },
      sort: { direction: 'descending', timestamp: 'last_edited_time' },
      page_size: 100,
      start_cursor: cursor
    });

    for (const page of searchResults.results) {
      if (page.archived || page.in_trash) {
        continue;
      }

      const parentDatabaseId = page.parent && page.parent.database_id;
      if (!parentDatabaseId || parentDatabaseId.replace(/-/g, '') !== databaseId.replace(/-/g, '')) {
        continue;
      }

      const title = getPageTitle(page);

      if (title) {
        existingCourses[title] = {
          id: page.id,
          mediaFiles: getMediaFiles(page)
        };
      }
    }

    cursor = searchResults.has_more ? searchResults.next_cursor : undefined;
  } while (cursor);

  return existingCourses;
}

function getMediaFiles(page) {
  const property = page.properties && page.properties[MEDIA_PROPERTY_NAME];
  return property && property.type === 'files' ? property.files : [];
}

function shouldUploadCover(existingCourse) {
  if (!existingCourse) {
    return true;
  }

  const mediaFiles = existingCourse.mediaFiles || [];
  return mediaFiles.length === 0 || mediaFiles.some(file => file.type !== 'file');
}

function buildCoverProperty(course, coverUploads, shouldSetFallback) {
  if (coverUploads && coverUploads.mediaUploadId) {
    return {
      files: [
        {
          type: 'file_upload',
          name: COVER_FILE_NAME,
          file_upload: { id: coverUploads.mediaUploadId }
        }
      ]
    };
  }

  if (shouldSetFallback && course.coverUrl && course.coverUrl.startsWith('http')) {
    return {
      files: [
        {
          type: 'external',
          name: COVER_FILE_NAME,
          external: { url: course.coverUrl }
        }
      ]
    };
  }

  return null;
}

function buildPageCover(course, coverUploads) {
  if (coverUploads && coverUploads.coverUploadId) {
    return {
      type: 'file_upload',
      file_upload: { id: coverUploads.coverUploadId }
    };
  }

  if (course.coverUrl && course.coverUrl.startsWith('http')) {
    return { type: 'external', external: { url: course.coverUrl } };
  }

  return null;
}

function buildProperties(course, coverUploads, shouldSetCoverFallback = false) {
  const category = normalizeSelectName(course.category);
  const properties = {
    'Nombre del recurso': {
      title: [
        {
          text: { content: course.title }
        }
      ]
    },
    Seleccionar: {
      select: { name: 'Niveles' }
    }
  };

  if (category) {
    properties[CATEGORY_PROPERTY_NAME] = {
      select: { name: category }
    };
  }

  const coverProperty = buildCoverProperty(course, coverUploads, shouldSetCoverFallback);
  if (coverProperty) {
    properties[MEDIA_PROPERTY_NAME] = coverProperty;
  }

  return properties;
}

function buildBlocks(course) {
  const blocks = [
    {
      object: 'block',
      type: 'callout',
      callout: {
        rich_text: [
          {
            type: 'text',
            text: { content: 'Duracion del curso: ' },
            annotations: { bold: true }
          },
          {
            type: 'text',
            text: { content: course.duration || 'Duracion no especificada' }
          }
        ],
        icon: {
          type: 'emoji',
          emoji: '⏱️'
        },
        color: 'gray_background'
      }
    },
    {
      object: 'block',
      type: 'paragraph',
      paragraph: {
        rich_text: [
          {
            type: 'text',
            text: { content: course.description || 'Sin descripcion detallada.' }
          }
        ]
      }
    }
  ];

  if (!course.modules || course.modules.length === 0) {
    return blocks;
  }

  for (const module of course.modules) {
    const moduleTitle = typeof module === 'object' ? module.title : module;
    const lessons = typeof module === 'object' && module.lessons ? module.lessons : [];

    blocks.push({
      object: 'block',
      type: 'toggle',
      toggle: {
        rich_text: [
          {
            type: 'text',
            text: { content: moduleTitle },
            annotations: { bold: true }
          }
        ],
        color: 'gray_background',
        children: lessons.slice(0, 50).map(lesson => ({
          object: 'block',
          type: 'bulleted_list_item',
          bulleted_list_item: {
            rich_text: [
              {
                type: 'text',
                text: { content: lesson }
              }
            ]
          }
        }))
      }
    });
  }

  return blocks;
}

async function updateCoursePage(pageId, course, properties, blocks, coverUploads, shouldSetCoverFallback) {
  const payload = {
    page_id: pageId,
    properties
  };

  const pageCover = coverUploads || shouldSetCoverFallback ? buildPageCover(course, coverUploads) : null;
  if (pageCover) {
    payload.cover = pageCover;
  }

  await notion.pages.update(payload);

  const existingBlocks = await notion.blocks.children.list({ block_id: pageId });
  for (const block of existingBlocks.results) {
    await notion.blocks.delete({ block_id: block.id }).catch(() => {});
    await sleep(350);
  }

  if (blocks.length > 0) {
    await notion.blocks.children.append({
      block_id: pageId,
      children: blocks
    });
    await sleep(350);
  }
}

async function createCoursePage(properties, blocks, course, coverUploads) {
  const dataSourceId = await getDataSourceId();
  const parent = dataSourceId ? { data_source_id: dataSourceId } : { type: 'database_id', database_id: databaseId };
  const payload = {
    parent,
    properties,
    children: blocks
  };

  const pageCover = buildPageCover(course, coverUploads);
  if (pageCover) {
    payload.cover = pageCover;
  }

  await notion.pages.create(payload);
}

async function syncCoursesWithDatabase(courses) {
  try {
    console.log(`Syncing ${courses.length} courses with Notion database...`);
    await ensureCategoryProperty(courses);
    const existingCourses = await listExistingCourses();
    console.log(`Existing courses in Notion: ${Object.keys(existingCourses).length}`);

    for (const course of courses) {
      const existingCourse = existingCourses[course.title];
      const needsCoverUpload = shouldUploadCover(existingCourse);
      const coverUploads = needsCoverUpload ? await uploadCoverImage(course, true) : null;
      const shouldSetCoverFallback = needsCoverUpload && !coverUploads;
      const properties = buildProperties(course, coverUploads, shouldSetCoverFallback);
      const blocks = buildBlocks(course);

      if (existingCourse) {
        const pageId = existingCourse.id;
        console.log(`[UPDATE] Course already exists: ${course.title} (${pageId})`);
        await updateCoursePage(pageId, course, properties, blocks, coverUploads, shouldSetCoverFallback);
      } else {
        console.log(`[CREATE] New course detected: ${course.title}`);
        await createCoursePage(properties, blocks, course, coverUploads);
      }
    }

    console.log('Notion sync completed.');
  } catch (error) {
    console.error('Error syncing Notion database:', error.body || error.message);
    throw error;
  }
}

module.exports = { syncCoursesWithDatabase };
