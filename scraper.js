require('dotenv').config();

const { chromium } = require('playwright');

const LOGIN_URL = 'https://miembro.daxus.com/users/sign_in';
const COURSE_LIST_URL = 'https://miembro.daxus.com/?browse=available';
const FALLBACK_CATEGORY_URL = 'https://miembro.daxus.com/categories/48527-comienza-por-aqui/courses';

function normalizeText(value, fallback = '') {
  return (value || fallback).trim();
}

function uniqueCards(cards) {
  const unique = [];
  const indexesByUrl = new Map();

  for (const card of cards) {
    if (!card.url || card.url.endsWith('/dashboard')) {
      continue;
    }

    if (!indexesByUrl.has(card.url)) {
      indexesByUrl.set(card.url, unique.length);
      unique.push(card);
      continue;
    }

    const existingIndex = indexesByUrl.get(card.url);
    if (!unique[existingIndex].category && card.category) {
      unique[existingIndex] = { ...unique[existingIndex], category: card.category };
    }
  }

  return unique;
}

async function extractCourseCards(page) {
  const cards = await page.evaluate(() => {
    function cleanCategoryText(value) {
      return (value || '')
        .replace(/\s+/g, ' ')
        .replace(/\s*Ver todo.*$/i, '')
        .trim();
    }

    function isCategoryCandidate(value) {
      if (!value || value.length > 80) {
        return false;
      }

      return !/^(buscar|catalogo|categorias|cursos|inicio|mi aprendizaje|ver todo)$/i.test(value);
    }

    const courseLinks = Array.from(document.querySelectorAll('a')).filter(link => {
      const image = link.querySelector('img');
      const isCourseLink = /\/course[s]?\//.test(link.href) || /\/\d+-/.test(link.href);
      return image && isCourseLink && !link.href.includes('/categories');
    });

    const courseLinkSet = new Set(courseLinks);
    const extractedCards = [];
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
    let currentCategory = '';
    let element = walker.currentNode;

    while (element) {
      if (courseLinkSet.has(element)) {
        const image = element.querySelector('img');
        extractedCards.push({
          url: element.href,
          coverUrl: image && image.src ? image.src : '',
          category: currentCategory
        });
      } else if (!courseLinkSet.has(element.closest('a'))) {
        const tagName = element.tagName.toLowerCase();
        const role = (element.getAttribute('role') || '').toLowerCase();
        const className = String(element.className || '').toLowerCase();
        const looksLikeHeading =
          /^h[1-4]$/.test(tagName) ||
          role === 'heading' ||
          /category|collection|heading|section|shelf|title/.test(className);

        if (looksLikeHeading) {
          const category = cleanCategoryText(element.innerText);
          if (isCategoryCandidate(category)) {
            currentCategory = category;
          }
        }
      }

      element = walker.nextNode();
    }

    return extractedCards;
  });

  return uniqueCards(cards);
}

async function extractFallbackCourseCards(page) {
  const cards = await page.evaluate(() => {
    const pageHeading = document.querySelector('h1, h2, h3, [role="heading"]');
    const category = pageHeading ? pageHeading.innerText.replace(/\s+/g, ' ').trim() : 'Comienza por aqui';

    return Array.from(document.querySelectorAll('a'))
      .filter(link => link.querySelector('img') && link.href.includes('/course'))
      .map(link => ({
        url: link.href,
        coverUrl: link.querySelector('img').src,
        category
      }));
  });

  return uniqueCards(cards);
}

async function scrapeCourses() {
  console.log('Starting browser for course extraction...');
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    console.log('Opening Daxus LATAM login...');
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });

    console.log('Signing in...');
    await page.fill('input[type="email"]', process.env.DAXUS_EMAIL);
    await page.fill('input[type="password"]', process.env.DAXUS_PASSWORD);

    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle', timeout: 30000 }).catch(() => {}),
      page.keyboard.press('Enter')
    ]);

    console.log('Session started.');

    await page.goto(COURSE_LIST_URL, { waitUntil: 'networkidle', timeout: 30000 });

    console.log('Extracting course links...');
    let courseCards = await extractCourseCards(page);

    if (courseCards.length === 0) {
      console.log('No courses found in the general view. Trying fallback category...');
      await page.goto(FALLBACK_CATEGORY_URL, { waitUntil: 'networkidle', timeout: 30000 });
      courseCards = await extractFallbackCourseCards(page);
    }

    const categories = Array.from(new Set(courseCards.map(card => card.category).filter(Boolean)));
    console.log(`Found ${courseCards.length} unique course cards across ${categories.length} categories.`);
    if (categories.length > 0) {
      console.log(`Categories detected: ${categories.join(', ')}`);
    }

    const courses = [];

    for (const card of courseCards) {
      console.log(`Analyzing course: ${card.url}`);
      await page.goto(card.url, { waitUntil: 'networkidle', timeout: 30000 }).catch(() => {});

      let title = 'Untitled course';
      let description = 'No detailed description.';
      let duration = 'Duration not specified';
      let modules = [];

      try {
        const titleEl = await page.$('h1, h2, h3.text-xl');
        if (titleEl) {
          title = await titleEl.innerText();
        }

        const descEl = await page.$('.description, .course-description, p:not(.module-title):not(.lesson-title)');
        if (descEl) {
          description = await descEl.innerText();
        }

        const durationEl = await page.$('.sidebar__duration span');
        if (durationEl) {
          duration = await durationEl.innerText();
        }

        modules = await page.evaluate(() => {
          const results = [];
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
            const lessons = [];
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

            if (sectionDiv) {
              const lessonLinks = sectionDiv.querySelectorAll('a.lesson__title');
              for (const link of lessonLinks) {
                const lessonTitle = link.innerText.trim();
                if (lessonTitle) {
                  lessons.push(lessonTitle);
                }
              }
            }

            if (moduleTitle && moduleTitle.length < 100) {
              results.push({ title: fullModuleTitle, lessons });
            }
          }

          return results;
        });
      } catch (error) {
        console.error(`Error extracting data from ${card.url}:`, error.message);
      }

      const seenModules = new Set();
      modules = modules.filter(module => {
        if (seenModules.has(module.title)) {
          return false;
        }
        seenModules.add(module.title);
        return true;
      });

      courses.push({
        title: normalizeText(title, 'Untitled course'),
        description: normalizeText(description, 'No detailed description.'),
        duration: normalizeText(duration, 'Duration not specified'),
        category: normalizeText(card.category, 'Sin categoria'),
        coverUrl: card.coverUrl,
        url: card.url,
        modules
      });
    }

    console.log('Course extraction finished. Courses extracted:', courses.length);
    return courses;
  } catch (error) {
    console.error('Critical scraping error:', error.message);
    throw error;
  } finally {
    await browser.close().catch(() => {});
  }
}

if (require.main === module) {
  scrapeCourses().then(console.log).catch(error => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = { scrapeCourses };
