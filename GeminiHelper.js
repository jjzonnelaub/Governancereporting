/**
 * PLACEHOLDER: This is where your actual summarization logic goes.
 * It could be an API call to Gemini, OpenAI, or another service.
 * @param {string} data The context or data to summarize.
 * @param {string} prompt The prompt guiding the summary.
 * @returns {string} The generated summary.
 */
function summarize(data, prompt) {
  if (!data || data.trim() === '') {
    return 'no data'; // Return empty if there's no data to summarize
  }
  // Replace this with your actual API call.
  // For example: return GeminiApp.generateContent([prompt, data]);
  console.log(`Called summarize with prompt: "${prompt}"`);
  return callGemini(data, prompt);
}

function callGemini(data, promptTemplate, model, generationArgs = {}) {

  // Input Data Validation
  if (!data || data.trim() === '' || data === 'N/A') {
    return 'N/A';
  }

  // Set a default model if one is not provided 
  if (!model) {
    model = 'gemini-2.0-flash'; // Default to the most cost-effective model.
    //Logger.log(`Model not specified, defaulting to ${model}`);
  }

  // --- Garbage Input Data Check to optimize costs ---
  const minWordCount = 2; // Minimum number of words to be considered valid
  const wordCount = 10; // Word count 
  const words = data.trim().split(/\s+/); // Split by one or more spaces
  const hasLetters = /[a-zA-Z]/.test(data); // Check if the string contains any letters

  if (words.length < minWordCount || !hasLetters) {
    //Logger.log(`Skipping API call: Input is considered garbage (words: ${words.length}, hasLetters: ${hasLetters}).`);
    return "Input too short or invalid."; // Return a clear message instead of 'N/A'
  }

  if( words.length < wordCount){
    return data;
  }

  // Retrieve GEMINI_API_KEY
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    Logger.log("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return "AI key not configured.";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${API_KEY}`;
  //const fullPrompt = promptTemplate.replace('{{data}}', data);
  
  const newPrompt = promptTemplate.concat(data);

  Logger.log('--- Calling Gemini API ---');
  Logger.log(`Model: ${model}`);
  Logger.log(`Prompt: ${newPrompt}`);
  
  const payload = {
    "contents": [{
      "parts": [{
        "text": newPrompt
      }]
    }],
    "generationConfig": generationArgs
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const responseData = JSON.parse(responseBody);
      const content = responseData.candidates?.[0]?.content?.parts?.[0]?.text || "Response not available.";
      
      //Logger.log(`Gemini Response: ${content.trim()}`);
     // Logger.log('--------------------------');
      
      return content.trim();
    } else {
      Logger.log(`ERROR - Gemini API Status: ${responseCode}, Response: ${responseBody}`);
      return "AI response failed.";
    }
  } catch (e) {
    Logger.log("FATAL ERROR: Failed to call Gemini API: " + e.toString());
    return "Error connecting to AI.";
  }
}


/**
 * Summarizes an array of texts in a single API call using UrlFetchApp.
 * @param {string[]} texts An array of strings to summarize.
 * @param {string} prompt The summarization instruction for each item.
 * @returns {string[]} An array of summaries in the same order as the input.
 */
function batchSummarize(texts, prompt) {
  if (!texts || texts.length === 0) {
    return [];
  }

  const model = 'gemini-2.0-flash';

  // 1. --- Get API Key ---
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) {
    console.error("ERROR: GEMINI_API_KEY not found in Script Properties.");
    return texts.map(() => "AI key not configured.");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${API_KEY}`;

  // 2. --- Construct the batch prompt with ULTRA-STRICT JSON rules ---
  const batchPrompt = `
${prompt}

CRITICAL JSON FORMATTING RULES:
1. You will receive EXACTLY ${texts.length} items to process
2. You MUST return EXACTLY ${texts.length} summaries - no more, no less
3. Return ONLY a valid JSON array: ["summary1", "summary2", ...]
4. NEVER use double quotes (") inside summaries - use single quotes (') instead
5. If a summary needs quotes, use single quotes: 'The system is on track'
6. Your ENTIRE response must be ONLY the JSON array - no explanations, no markdown
7. Do NOT wrap response in markdown code blocks
8. Process each item as ONE summary - do not split items

Items to process:
${texts.map((t, i) => `\n=== ITEM ${i + 1} OF ${texts.length} ===\n${t}`).join('\n')}

RESPONSE FORMAT (EXACT):
["summary 1 here", "summary 2 here", ..., "summary ${texts.length} here"]

REMEMBER: 
- Return EXACTLY ${texts.length} summaries
- NO double quotes inside summaries
- ONLY the JSON array, nothing else`;

  // 3. --- Prepare Payload and Options ---
  const payload = {
    "contents": [{
      "parts": [{
        "text": batchPrompt
      }]
    }],
    "generationConfig": {
      "temperature": 0.7,
      "maxOutputTokens": 4096
    }
  };
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    'payload': JSON.stringify(payload)
  };

  try {
    // 4. --- Make the API Call ---
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const responseData = JSON.parse(responseBody);
      const rawResponse = responseData.candidates?.[0]?.content?.parts?.[0]?.text || "[]";

      // 5. --- Clean and parse the JSON with robust error handling ---
      let cleanedResponse = rawResponse.trim();
      
      // Remove markdown code blocks
      cleanedResponse = cleanedResponse.replace(/```json\n?|\n?```/g, '');
      cleanedResponse = cleanedResponse.trim();
      
      // Find the JSON array (in case there's extra text)
      const arrayMatch = cleanedResponse.match(/\[[\s\S]*\]/);
      if (arrayMatch) {
        cleanedResponse = arrayMatch[0];
      }
      
      // Log the first 500 chars for debugging
      console.log('Raw AI response (first 500 chars):', rawResponse.substring(0, 500));
      console.log('Cleaned response (first 500 chars):', cleanedResponse.substring(0, 500));
      
      let summaries;
      try {
        summaries = JSON.parse(cleanedResponse);
      } catch (parseError) {
        console.error('JSON Parse Error:', parseError.message);
        console.error('Failed to parse:', cleanedResponse.substring(0, 500));
        
        // Attempt to fix common escaping issues
        try {
          // Try to fix unescaped quotes by replacing internal double quotes with single quotes
          let fixedResponse = cleanedResponse;
          
          // This regex attempts to fix quotes inside JSON strings
          // It's a heuristic approach - not perfect but handles common cases
          fixedResponse = fixedResponse.replace(/"([^"]*)"([^"]*)"([^"]*)"/g, (match, p1, p2, p3) => {
            // If we find quotes inside a string, replace with single quotes
            return `"${p1}'${p2}'${p3}"`;
          });
          
          console.log('Attempting to parse fixed response...');
          summaries = JSON.parse(fixedResponse);
          console.log('Successfully parsed after fixing!');
        } catch (secondError) {
          console.error('Still failed after attempted fix:', secondError.message);
          // Return error for all items
          return texts.map(() => "Error: AI response format invalid");
        }
      }

      // 6. --- Enhanced Validation with Better Error Handling ---
      if (Array.isArray(summaries)) {
        if (summaries.length === texts.length) {
          // Perfect match - return as-is
          console.log(`✓ Batch successful: ${summaries.length} summaries generated`);
          return summaries;
          
        } else if (summaries.length > texts.length) {
          // Too many summaries - truncate and warn
          const extras = summaries.length - texts.length;
          console.warn(`Truncating ${extras} extra summaries`);
          console.log('First few inputs:', texts.slice(0, 3));
          console.log('First few outputs:', summaries.slice(0, 3));
          return summaries.slice(0, texts.length);
          
        } else {
          // Too few summaries - pad with errors
          console.error(`Too few summaries: expected ${texts.length}, got ${summaries.length}`);
          console.log('First few inputs:', texts.slice(0, 3));
          console.log('All outputs:', summaries);
          
          // Pad with error messages for missing summaries
          const padded = [...summaries];
          while (padded.length < texts.length) {
            padded.push("Error: AI response incomplete");
          }
          return padded;
        }
      } else {
        console.error("AI response was not an array:", typeof summaries);
        return texts.map(() => "Error: AI response format invalid");
      }
      
    } else {
      console.error(`ERROR - Gemini API Status: ${responseCode}, Response: ${responseBody}`);
      return texts.map(() => "AI response failed");
    }
    
  } catch (e) {
    console.error("FATAL ERROR: Failed to call or parse Gemini batch API: " + e.toString());
    console.error("Stack trace:", e.stack);
    return texts.map(() => "Error: Connecting to AI");
  }
}
function callGeminiAPIWithFallback(prompt, fallback) {
  const MAX_RETRIES = 2;
  let attempt = 0;
  
  while (attempt <= MAX_RETRIES) {
    try {
      const response = callGeminiAPI(prompt);
      
      if (response && typeof response === 'string') {
        let cleaned = response.trim();
        
        // Basic cleanup
        cleaned = cleaned.replace(/^["']|["']$/g, '');
        cleaned = cleaned.replace(/^```.*\n?|\n?```$/g, '');
        cleaned = cleaned.replace(/^\*\*|^\*|\*\*$|\*$/g, '');
        
        // Validation
        const hasBullets = /[•\-]\s/.test(cleaned) || /^\d+\.\s/.test(cleaned);
        const wordCount = cleaned.split(/\s+/).length;
        
        if (cleaned.length > 10 && cleaned.length < 1200 && !hasBullets && wordCount <= 60) {

          return cleaned;
        }
        
        console.warn(`Invalid AI response (attempt ${attempt + 1}): length=${cleaned.length}, words=${wordCount}, bullets=${hasBullets}`);
      }
    } catch (error) {
      console.error(`Error calling AI (attempt ${attempt + 1}): ${error.message}`);
    }
    
    attempt++;
    if (attempt <= MAX_RETRIES) {
      Utilities.sleep(500);
    }
  }
  
  return fallback;
}