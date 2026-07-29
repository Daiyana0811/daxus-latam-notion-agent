require('dotenv').config();

const fs = require('fs');
const { Client } = require('@notionhq/client');

const APOSTILLA_PROPERTY_NAME = 'Apostilla';
const TRANSCRIPTION_PROPERTY_NAME = 'Transcripcion';

const notion = new Client({ auth: process.env.NOTION_API_KEY });

function isFilesPropertyEmpty(property) {
  return Boolean(property && property.type === 'files' && property.files.length === 0);
}

async function getDataSourceId(databaseId) {
  const database = await notion.databases.retrieve({ database_id: databaseId });
  return database.data_sources?.[0]?.id || null;
}

async function main() {
  const databaseId = process.env.NOTION_DATABASE_ID;
  if (!process.env.NOTION_API_KEY || !databaseId) {
    throw new Error('Missing required environment variables: NOTION_API_KEY, NOTION_DATABASE_ID');
  }

  const dataSourceId = await getDataSourceId(databaseId);
  let cursor;
  let pending = 0;
  let completed = 0;

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

      const apostillaEmpty = isFilesPropertyEmpty(page.properties[APOSTILLA_PROPERTY_NAME]);
      const transcriptionEmpty = isFilesPropertyEmpty(page.properties[TRANSCRIPTION_PROPERTY_NAME]);

      if (apostillaEmpty && transcriptionEmpty) {
        pending += 1;
      } else if (apostillaEmpty && !transcriptionEmpty) {
        completed += 1;
      }
    }

    cursor = response.has_more ? response.next_cursor : undefined;
  } while (cursor);

  console.log(`Pending transcriptions: ${pending}`);
  console.log(`Completed transcriptions: ${completed}`);

  if (process.env.GITHUB_OUTPUT) {
    fs.appendFileSync(process.env.GITHUB_OUTPUT, `pending=${pending}\ncompleted=${completed}\n`);
  }
}

main().catch(error => {
  console.error('Error counting pending transcriptions:', error);
  process.exit(1);
});
