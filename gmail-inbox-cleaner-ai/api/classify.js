const { GoogleGenerativeAI } = require("@google/generative-ai");

const BATCH_SIZE = 200;

const SYSTEM_PROMPT = `You are an email triage assistant. You will receive a JSON array of email senders. For each sender, decide whether their emails should be DELETED or KEPT.

DELETE these types of emails:
- Marketing and promotional emails
- Newsletters and digests
- Notification spam (social media notifications, app notifications)
- Automated alerts that are not critical (e.g. "someone liked your post")
- Mailing lists the user likely doesn't read
- Recruitment/job spam
- Cold outreach and sales pitches

KEEP (do NOT select for deletion) these types of emails:
- Personal email addresses (e.g. john.doe@gmail.com, jane@yahoo.com) — real people
- Transactional emails (order confirmations, shipping updates, purchase receipts)
- Financial emails (bank statements, invoices, payment confirmations)
- Contracts, legal, or compliance emails
- Account security emails (password resets, 2FA codes, login alerts)
- Calendar invites and RSVPs
- Emails from employers or colleagues
- Government or healthcare communications
- Any email that seems individually important or time-sensitive

Respond with ONLY a JSON object in this exact format, no markdown fencing, no extra text:
{"delete":["email1@example.com","email2@example.com"]}

The "delete" array should contain ONLY the email addresses of senders whose emails should be deleted. If no emails should be deleted, return {"delete":[]}.`;

function buildUserPrompt(senders) {
  const rows = senders.map(
    (s) =>
      `${s.email} | ${s.name || "(no name)"} | count: ${s.count} | subject: ${s.subject || "(none)"}`
  );
  return `Classify these email senders:\n\n${rows.join("\n")}`;
}

function parseGeminiResponse(text) {
  // Strip markdown code fences if present
  let cleaned = text.trim();
  if (cleaned.startsWith("```")) {
    cleaned = cleaned.replace(/^```(?:json)?\s*/, "").replace(/\s*```$/, "");
  }

  // Try direct JSON parse first
  try {
    const parsed = JSON.parse(cleaned);
    if (parsed && Array.isArray(parsed.delete)) {
      return parsed.delete.filter((e) => typeof e === "string");
    }
  } catch (_) {
    // Fall through to extraction attempts
  }

  // Try to find a JSON object anywhere in the text
  const objectMatch = cleaned.match(/\{[\s\S]*\}/);
  if (objectMatch) {
    try {
      const parsed = JSON.parse(objectMatch[0]);
      if (parsed && Array.isArray(parsed.delete)) {
        return parsed.delete.filter((e) => typeof e === "string");
      }
    } catch (_) {
      // Fall through
    }
  }

  // Try to find a JSON array (in case LLM returned just the array)
  const arrayMatch = cleaned.match(/\[[\s\S]*\]/);
  if (arrayMatch) {
    try {
      const parsed = JSON.parse(arrayMatch[0]);
      if (Array.isArray(parsed)) {
        return parsed.filter((e) => typeof e === "string");
      }
    } catch (_) {
      // Fall through
    }
  }

  // Last resort: extract anything that looks like an email address after "delete"
  const deleteSection = cleaned.toLowerCase().includes("delete")
    ? cleaned.slice(cleaned.toLowerCase().indexOf("delete"))
    : cleaned;
  const emailPattern = /[\w.+-]+@[\w.-]+\.\w+/g;
  const emails = deleteSection.match(emailPattern);
  return emails || [];
}

async function classifyBatch(model, senders) {
  const result = await model.generateContent(buildUserPrompt(senders));
  const text = result.response.text();
  return parseGeminiResponse(text);
}

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function sendJson(res, status, body) {
  setCors(res);
  res.setHeader("Content-Type", "application/json");
  res.statusCode = status;
  res.end(JSON.stringify(body));
}

module.exports = async function handler(req, res) {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    setCors(res);
    res.statusCode = 204;
    return res.end();
  }

  if (req.method !== "POST") {
    return sendJson(res, 405, { error: "Method not allowed" });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return sendJson(res, 500, { error: "GEMINI_API_KEY not configured" });
  }

  const body = req.body || {};
  const senders = body.senders;

  if (!Array.isArray(senders) || senders.length === 0) {
    return sendJson(res, 400, {
      error: "Request body must contain a non-empty 'senders' array",
    });
  }

  // Validate sender objects
  const validSenders = senders.filter(
    (s) => s && typeof s.email === "string" && s.email.includes("@")
  );
  if (validSenders.length === 0) {
    return sendJson(res, 400, {
      error: "No valid sender objects found in array",
    });
  }

  try {
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({
      model: "gemini-2.0-flash-lite",
      systemInstruction: SYSTEM_PROMPT,
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 8192,
      },
    });

    // Batch if necessary to stay within token limits
    const batches = [];
    for (let i = 0; i < validSenders.length; i += BATCH_SIZE) {
      batches.push(validSenders.slice(i, i + BATCH_SIZE));
    }

    const allDeleteEmails = new Set();

    // Process batches sequentially to avoid rate limits
    for (const batch of batches) {
      const deleteEmails = await classifyBatch(model, batch);

      // Only include emails that were actually in the input
      const inputEmails = new Set(batch.map((s) => s.email.toLowerCase()));
      for (const email of deleteEmails) {
        if (inputEmails.has(email.toLowerCase())) {
          allDeleteEmails.add(email.toLowerCase());
        }
      }
    }

    // Return emails using the original casing from the input
    const emailCaseMap = new Map(
      validSenders.map((s) => [s.email.toLowerCase(), s.email])
    );
    const deleteList = [...allDeleteEmails].map(
      (e) => emailCaseMap.get(e) || e
    );

    return sendJson(res, 200, { delete: deleteList });
  } catch (err) {
    console.error("Gemini API error:", err);
    return sendJson(res, 502, {
      error: "AI classification failed",
      detail: err.message,
    });
  }
};
