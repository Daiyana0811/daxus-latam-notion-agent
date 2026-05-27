require('dotenv').config();

const { Client } = require('@notionhq/client');

const notionApiKey = process.env.NOTION_API_KEY;
const databaseId = process.env.NOTION_DATABASE_ID;
const CATEGORY_PROPERTY_NAME = 'Categoria';

if (!notionApiKey) {
  throw new Error('Missing required environment variable: NOTION_API_KEY');
}

if (!databaseId) {
  throw new Error('Missing required environment variable: NOTION_DATABASE_ID');
}

const notion = new Client({ auth: notionApiKey });
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

function normalizeSelectName(value) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  return text.length > 100 ? text.slice(0, 100).trim() : text;
}

function getCategoryOptions(courses) {
  return Array.from(new Set(courses.map(course => normalizeSelectName(course.category)).filter(Boolean)))
    .map(name => ({ name, color: 'default' }));
}

async function ensureCategoryProperty(courses) {
  const database = await notion.databases.retrieve({ database_id: databaseId });
  const existingProperty = database.properties[CATEGORY_PROPERTY_NAME];

  if (existingProperty) {
    if (existingProperty.type !== 'select') {
      throw new Error(`The Notion property "${CATEGORY_PROPERTY_NAME}" must be a select property.`);
    }

    return;
  }

  const options = getCategoryOptions(courses);
  await notion.databases.update({
    database_id: databaseId,
    properties: {
      [CATEGORY_PROPERTY_NAME]: {
        select: options.length > 0 ? { options } : {}
      }
    }
  });

  console.log(`Created Notion select property: ${CATEGORY_PROPERTY_NAME}`);
}

async function listExistingCourses() {
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
      const parentDatabaseId = page.parent && page.parent.database_id;
      if (!parentDatabaseId || parentDatabaseId.replace(/-/g, '') !== databaseId.replace(/-/g, '')) {
        continue;
      }

      const titlePropKey = Object.keys(page.properties).find(key => page.properties[key].type === 'title');
      const title = titlePropKey && page.properties[titlePropKey].title[0];

      if (title) {
        existingCourses[title.plain_text.trim()] = page.id;
      }
    }

    cursor = searchResults.has_more ? searchResults.next_cursor : undefined;
  } while (cursor);

  return existingCourses;
}

function buildProperties(course) {
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

  if (course.coverUrl && course.coverUrl.startsWith('http')) {
    properties['Archivos y multimedia'] = {
      files: [
        {
          type: 'external',
          name: 'Portada',
          external: { url: course.coverUrl }
        }
      ]
    };
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

async function updateCoursePage(pageId, course, properties, blocks) {
  const payload = {
    page_id: pageId,
    properties
  };

  if (course.coverUrl && course.coverUrl.startsWith('http')) {
    payload.cover = { type: 'external', external: { url: course.coverUrl } };
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

async function createCoursePage(properties, blocks) {
  await notion.pages.create({
    parent: { type: 'database_id', database_id: databaseId },
    properties,
    children: blocks
  });
}

async function syncCoursesWithDatabase(courses) {
  try {
    console.log(`Syncing ${courses.length} courses with Notion database...`);
    await ensureCategoryProperty(courses);
    const existingCourses = await listExistingCourses();
    console.log(`Existing courses in Notion: ${Object.keys(existingCourses).length}`);

    for (const course of courses) {
      const properties = buildProperties(course);
      const blocks = buildBlocks(course);

      if (existingCourses[course.title]) {
        const pageId = existingCourses[course.title];
        console.log(`[UPDATE] Course already exists: ${course.title} (${pageId})`);
        await updateCoursePage(pageId, course, properties, blocks);
      } else {
        console.log(`[CREATE] New course detected: ${course.title}`);
        await createCoursePage(properties, blocks);
      }
    }

    console.log('Notion sync completed.');
  } catch (error) {
    console.error('Error syncing Notion database:', error.body || error.message);
    throw error;
  }
}

module.exports = { syncCoursesWithDatabase };
