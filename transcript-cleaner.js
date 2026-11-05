const VTT_TIMING_PATTERN = /^\s*(?:\d{2}:)?\d{2}:\d{2}[.,]\d{3}\s+-->\s+(?:\d{2}:)?\d{2}:\d{2}[.,]\d{3}(?:\s|$)/;

const HTML_ENTITIES = {
  amp: '&',
  apos: "'",
  gt: '>',
  lt: '<',
  nbsp: ' ',
  quot: '"',
  lrm: '',
  rlm: ''
};

function decodeHtmlEntities(value) {
  return String(value || '').replace(/&(#(?:x[0-9a-f]+|\d+)|[a-z]+);/gi, (match, entity) => {
    if (entity[0] === '#') {
      const hexadecimal = entity[1]?.toLowerCase() === 'x';
      const codePoint = Number.parseInt(entity.slice(hexadecimal ? 2 : 1), hexadecimal ? 16 : 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
    }

    return Object.prototype.hasOwnProperty.call(HTML_ENTITIES, entity.toLowerCase())
      ? HTML_ENTITIES[entity.toLowerCase()]
      : match;
  });
}

function cleanTranscriptText(value) {
  return decodeHtmlEntities(String(value || '')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]*>/g, ' '))
    .replace(/\b[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}-\d+\b/gi, ' ')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, ' ')
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseVtt(vttText) {
  const blocks = String(vttText || '')
    .replace(/^\uFEFF/, '')
    .replace(/\r\n?/g, '\n')
    .split(/\n\s*\n/);
  const cues = [];

  for (const block of blocks) {
    const lines = block.split('\n').map(line => line.trim());
    const timingIndex = lines.findIndex(line => VTT_TIMING_PATTERN.test(line));

    // Only timed cues contain spoken transcript. This deliberately excludes
    // WEBVTT headers, cue IDs, NOTE, STYLE, REGION, JSON and Stream metadata.
    if (timingIndex < 0) {
      continue;
    }

    const cueText = cleanTranscriptText(lines.slice(timingIndex + 1).join(' '));
    if (cueText && cueText !== cues[cues.length - 1]) {
      cues.push(cueText);
    }
  }

  return cues.join(' ').replace(/\s+/g, ' ').trim();
}

module.exports = {
  cleanTranscriptText,
  decodeHtmlEntities,
  parseVtt
};
