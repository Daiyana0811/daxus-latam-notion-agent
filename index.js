require('dotenv').config();

const { scrapeCourses } = require('./scraper');
const { syncCoursesWithDatabase } = require('./notion');

const REQUIRED_ENV = [
  'DAXUS_EMAIL',
  'DAXUS_PASSWORD',
  'NOTION_API_KEY',
  'NOTION_DATABASE_ID'
];

function validateEnv() {
  const missing = REQUIRED_ENV.filter(name => !process.env[name]);
  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }
}

async function main() {
  try {
    validateEnv();
    console.log('Starting Daxus LATAM -> Notion Database Sync agent...');

    const courses = await scrapeCourses();

    if (!courses || courses.length === 0) {
      console.log('No courses were found to sync.');
      return;
    }

    await syncCoursesWithDatabase(courses);

    console.log('\n=========================================');
    console.log('Sync process completed successfully.');
    console.log('=========================================');
  } catch (error) {
    console.error('Fatal error while running the agent:', error);
    process.exit(1);
  }
}

main();
