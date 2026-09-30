'use strict';
// Codex Cloud task helper: uses the task's existing model, never starts a nested API call.
// Only the bundled synthetic fixture is accepted here. No credentials or live database access.
const fs = require('node:fs');
const path = require('node:path');
const interpretation = require('./interpretation.js');
const worker = require('./worker.cjs');
const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'synthetic.json'), 'utf8'));
const request = fixture.request;
const command = process.argv[2];
if (command === 'prompt') {
  process.stdout.write(interpretation.buildPrompt(request) + '\n\nOutput schema:\n' + JSON.stringify(interpretation.schema) + '\n');
} else if (command === 'evaluate') {
  // The cloud task produces condition JSON from the prompt. Pipe that JSON through stdin.
  const conditions = JSON.parse(fs.readFileSync(0, 'utf8'));
  const response = interpretation.makeResponse(request, conditions);
  const output = { response };
  if (response.status === 'ready') output.draft = worker.calculate(request, response, fixture.snapshot,
    { conditionsConfirmed: true, packConfirmed: true, analysisId: fixture.snapshot.analysisId }, fixture.nowMs);
  process.stdout.write(JSON.stringify(output, null, 2) + '\n');
} else {
  console.error('Usage: node ai-order/cloud-task.cjs prompt | evaluate (condition JSON on stdin)');
  process.exitCode = 1;
}
