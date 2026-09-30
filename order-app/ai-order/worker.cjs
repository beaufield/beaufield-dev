'use strict';
const interpretation = require('./interpretation.js');
const core = require('./core.generated.cjs');

async function interpret(request, provider) {
  const safe = interpretation.validateRequest(request);
  if (!provider || provider.usage !== 'chatgpt-plan' || typeof provider.complete !== 'function')
    throw new Error('A ChatGPT plan provider is required; there is no API-key fallback.');
  const conditions = await provider.complete(interpretation.buildPrompt(safe), interpretation.schema);
  return interpretation.makeResponse(safe, conditions);
}

function calculate(request, response, snapshot, confirmation, nowMs = Date.now()) {
  const accepted = interpretation.acceptResponse(request, response);
  if (accepted.status !== 'ready') throw new Error('Clarification or unsupported conditions remain.');
  if (!snapshot || !confirmation || confirmation.analysisId !== snapshot.analysisId ||
      confirmation.conditionsConfirmed !== true)
    throw new Error('Conditions and the analysis snapshot must be confirmed before calculation.');
  // Interpretation never contains these booleans. Only the user confirmation supplies them.
  return core.aiOrderCalculateSnapshot({ ...accepted.conditions, packConfirmed: confirmation.packConfirmed === true }, snapshot, nowMs);
}
module.exports = { interpret, calculate };
