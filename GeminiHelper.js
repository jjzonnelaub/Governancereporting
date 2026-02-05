/**
 * PLACEHOLDER: This is where your actual summarization logic goes.
 * It could be an API call to Gemini, OpenAI, or another service.
 * @param {string} data The context or data to summarize.
 * @param {string} prompt The prompt guiding the summary.
 * @returns {string} The generated summary.
 */
function summarize(data, prompt, model, generationArgs) {
  if (!data || data.trim() === '') {
    return 'N/A';
  }
  return callGemini(data, prompt, model, generationArgs);
}

function callGemini(data, promptTemplate, model, generationArgs) {
  // Input validation
  if (!data || data.trim() === '' || data === 'N/A') {
    return 'N/A';
  }

  model = model || 'gemini-2.0-flash';

  // Default generation config optimized for summarization
  var config = generationArgs && Object.keys(generationArgs).length > 0
    ? generationArgs
    : { temperature: 0.3, maxOutputTokens: 200 };

  // Garbage input check — skip API call for non-meaningful input
  var words = data.trim().split(/\s+/);
  var hasLetters = /[a-zA-Z]/.test(data);

  if (words.length < 2 || !hasLetters) {
    return 'N/A';
  }

  // Short input is already concise enough — return as-is
  if (words.length < 5) {
    return data.trim();
  }

  // Retrieve API key
  var API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    console.error('❌ GEMINI_API_KEY not found in Script Properties');
    return 'AI key not configured.';
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + API_KEY;

  // Concatenate prompt and data with clear delimiter
  var fullPrompt = promptTemplate + '\n\n---\n\n' + data;

  var payload = {
    "contents": [{ "parts": [{ "text": fullPrompt }] }],
    "generationConfig": config
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      var responseData = JSON.parse(responseBody);
      var content = responseData.candidates
        && responseData.candidates[0]
        && responseData.candidates[0].content
        && responseData.candidates[0].content.parts
        && responseData.candidates[0].content.parts[0]
        && responseData.candidates[0].content.parts[0].text;

      return content ? content.trim() : 'Response not available.';
    } else {
      console.error('❌ Gemini API Status: ' + responseCode + ', Response: ' + responseBody);
      return 'AI response failed.';
    }
  } catch (e) {
    console.error('❌ Failed to call Gemini API: ' + e.toString());
    return 'Error connecting to AI.';
  }
}


/**
 * Summarizes an array of texts in a single API call using UrlFetchApp.
 * @param {string[]} texts An array of strings to summarize.
 * @param {string} prompt The summarization instruction for each item.
 * @returns {string[]} An array of summaries in the same order as the input.
 */
function batchSummarize(texts, prompt, options) {
  if (!texts || texts.length === 0) {
    return [];
  }

  var chunkSize = (options && options.chunkSize) || 20;
  var temperature = (options && options.temperature) || 0.3;
  var maxTokens = (options && options.maxOutputTokens) || 4096;
  var model = (options && options.model) || 'gemini-2.0-flash';

  // Chunk large batches to avoid token limit truncation
  if (texts.length > chunkSize) {
    console.log('📊 Batch too large (' + texts.length + '), chunking into groups of ' + chunkSize);
    var allSummaries = [];
    for (var i = 0; i < texts.length; i += chunkSize) {
      var chunk = texts.slice(i, i + chunkSize);
      var chunkResults = batchSummarize(chunk, prompt, {
        chunkSize: chunkSize,
        temperature: temperature,
        maxOutputTokens: maxTokens,
        model: model
      });
      allSummaries = allSummaries.concat(chunkResults);

      // Rate limiting between chunks
      if (i + chunkSize < texts.length) {
        Utilities.sleep(300);
      }
    }
    return allSummaries;
  }

  // Get API key
  var API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    console.error('❌ GEMINI_API_KEY not found in Script Properties');
    return texts.map(function() { return 'AI key not configured.'; });
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + API_KEY;

  // Build batch prompt
  var itemsBlock = texts.map(function(t, i) {
    return '\n=== ITEM ' + (i + 1) + ' OF ' + texts.length + ' ===\n' + t;
  }).join('\n');

  var batchPrompt = prompt + '\n\n'
    + 'RESPONSE RULES:\n'
    + '1. Process EXACTLY ' + texts.length + ' items below\n'
    + '2. Return EXACTLY ' + texts.length + ' summaries as a JSON array\n'
    + '3. Format: ["summary one", "summary two", ...]\n'
    + '4. Use single quotes inside summaries, never double quotes\n'
    + '5. Return ONLY the JSON array — no markdown, no explanation\n'
    + '\n---\n'
    + itemsBlock;

  var payload = {
    "contents": [{ "parts": [{ "text": batchPrompt }] }],
    "generationConfig": {
      "temperature": temperature,
      "maxOutputTokens": maxTokens
    }
  };

  var options_req = {
    'method': 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch(url, options_req);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode !== 200) {
      console.error('❌ Gemini API Status: ' + responseCode);
      return texts.map(function() { return 'AI response failed'; });
    }

    var responseData = JSON.parse(responseBody);
    var rawResponse = responseData.candidates
      && responseData.candidates[0]
      && responseData.candidates[0].content
      && responseData.candidates[0].content.parts
      && responseData.candidates[0].content.parts[0]
      && responseData.candidates[0].content.parts[0].text || '[]';

    // Clean response
    var cleaned = rawResponse.trim();
    cleaned = cleaned.replace(/```json\n?|\n?```/g, '');
    cleaned = cleaned.trim();

    // Extract JSON array if surrounded by extra text
    var arrayMatch = cleaned.match(/\[[\s\S]*\]/);
    if (arrayMatch) {
      cleaned = arrayMatch[0];
    }

    // Parse with recovery
    var summaries;
    try {
      summaries = JSON.parse(cleaned);
    } catch (parseError) {
      console.error('⚠️ JSON parse failed, attempting recovery: ' + parseError.message);

      try {
        // Replace internal double quotes with single quotes
        var fixed = cleaned.replace(/(?<=":?\s*"[^"]*)"(?=[^"]*")/g, "'");
        summaries = JSON.parse(fixed);
        console.log('✓ JSON recovery successful');
      } catch (secondError) {
        console.error('❌ JSON recovery failed: ' + secondError.message);
        return texts.map(function() { return 'Error: AI response format invalid'; });
      }
    }

    // Validate array length
    if (!Array.isArray(summaries)) {
      console.error('❌ Response was not an array');
      return texts.map(function() { return 'Error: AI response format invalid'; });
    }

    if (summaries.length === texts.length) {
      console.log('✓ Batch successful: ' + summaries.length + ' summaries');
      return summaries;
    }

    if (summaries.length > texts.length) {
      console.warn('⚠️ Truncating ' + (summaries.length - texts.length) + ' extra summaries');
      return summaries.slice(0, texts.length);
    }

    // Pad if too few
    console.error('⚠️ Short batch: expected ' + texts.length + ', got ' + summaries.length);
    while (summaries.length < texts.length) {
      summaries.push('Error: AI response incomplete');
    }
    return summaries;

  } catch (e) {
    console.error('❌ Batch API call failed: ' + e.toString());
    return texts.map(function() { return 'Error: Connecting to AI'; });
  }
}
function callGeminiAPIWithFallback(prompt, fallback, model) {
  var MAX_RETRIES = 2;
  var attempt = 0;

  while (attempt <= MAX_RETRIES) {
    try {
      // Pass empty string as data since prompt is self-contained
      var response = callGemini(prompt, '', model, {
        temperature: 0.3,
        maxOutputTokens: 200
      });

      if (response && typeof response === 'string') {
        var cleaned = response.trim();

        // Strip formatting artifacts
        cleaned = cleaned.replace(/^["']|["']$/g, '');
        cleaned = cleaned.replace(/^```.*\n?|\n?```$/g, '');
        cleaned = cleaned.replace(/^\*\*|^\*|\*\*$|\*$/g, '');

        // Validate: no bullets, reasonable length, meaningful content
        var hasBullets = /[•\-]\s/.test(cleaned) || /^\d+\.\s/m.test(cleaned);
        var wordCount = cleaned.split(/\s+/).length;

        if (cleaned.length > 10 && cleaned.length < 1200 && !hasBullets && wordCount <= 60) {
          if (attempt > 0) {
            console.log('✓ AI response accepted on attempt ' + (attempt + 1));
          }
          return cleaned;
        }

        console.warn('⚠️ Invalid AI response (attempt ' + (attempt + 1) + '): '
          + 'length=' + cleaned.length + ', words=' + wordCount + ', bullets=' + hasBullets);
      }
    } catch (error) {
      console.error('❌ AI call failed (attempt ' + (attempt + 1) + '): ' + error.message);
    }

    attempt++;
    if (attempt <= MAX_RETRIES) {
      Utilities.sleep(500);
    }
  }

  console.warn('⚠️ All AI attempts failed, using fallback');
  return fallback;
}
