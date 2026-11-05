const test = require('node:test');
const assert = require('node:assert/strict');
const { cleanTranscriptText, parseVtt } = require('../transcript-cleaner');

test('parseVtt keeps only timed spoken cues', () => {
  const source = `\uFEFFWEBVTT

NOTE Microsoft Stream metadata
{"language":"es-CO","generated":true}

STYLE
::cue { color: white; }

REGION
id:fred
width:40%

60043f6e-2ff5-407c-8e15-e260117bb504-1
00:00:01.000 --> 00:00:04.000 position:10% align:start
<v Docente>Hola &amp; bienvenidos al curso.</v>

cue-2
00:00:05.000 --> 00:00:08.000
Hoy veremos &lt;conceptos&gt; importantes.
`;

  assert.equal(
    parseVtt(source),
    'Hola & bienvenidos al curso. Hoy veremos <conceptos> importantes.'
  );
});

test('cleanTranscriptText removes markup and invisible control characters', () => {
  assert.equal(
    cleanTranscriptText('<c.colorE5E5E5>Texto&nbsp;normal</c> 60043f6e-2ff5-407c-8e15-e260117bb504-1\u200B\u0007'),
    'Texto normal'
  );
});

test('parseVtt does not duplicate consecutive identical cues', () => {
  const source = `WEBVTT

00:00:01.000 --> 00:00:02.000
Una frase.

00:00:02.000 --> 00:00:03.000
Una frase.
`;

  assert.equal(parseVtt(source), 'Una frase.');
});
