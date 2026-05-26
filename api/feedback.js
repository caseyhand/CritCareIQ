async function callClaude(key, body, timeoutMs = 25000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`Anthropic ${res.status}: ${err?.error?.message || res.statusText}`);
    }
    return res.json();
  } finally {
    clearTimeout(timer);
  }
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { name, answers } = req.body;
  if (!name || !answers) return res.status(400).json({ error: "Missing name or answers" });

  const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
  if (!ANTHROPIC_KEY) return res.status(500).json({ error: "API key not configured" });

  const summary = `Physician: ${name}
Q1 - CC billing confidence (1-5): ${answers.q1 ?? "not answered"}
Q2 - Believes required for CC billing: ${(answers.q2 || []).join(", ") || "not answered"}
Q3 - Cases they typically bill as CC: ${(answers.q3 || []).join(", ") || "not answered"}
Q4 - Main reason for not billing CC: ${answers.q4 || "not answered"}
Q5 - EPIC CC note familiarity (1-5): ${answers.q5 ?? "not answered"}
Q6 - What would help most: ${(answers.q6 || []).join(", ") || "not answered"}
Q7 - Open comments: ${answers.q7 || "none"}`;

  try {
    const [feedbackData, tagData] = await Promise.all([
      callClaude(ANTHROPIC_KEY, {
        model: "claude-sonnet-4-5",
        max_tokens: 1000,
        system: `You are Casey, an emergency medicine AMD running a critical care billing improvement program at Mercy Fitzgerald ED. You've just received a physician's survey responses about CC documentation. Write a brief, warm, peer-to-peer personalized feedback message directly to this physician.
Format your response as:
1. One sentence acknowledging their specific situation (reference their actual answers — confidence score, main barrier, etc.)
2. "YOUR ACTION ITEMS:" followed by exactly 3 bullet points — specific, practical, non-punitive things they should do based on their gaps
3. One closing sentence of encouragement
Keep it under 150 words total. Never mention billing revenue. Never be punitive. Write as a colleague, not an administrator. Use plain text, no markdown headers.`,
        messages: [{ role: "user", content: `Here are the survey responses:\n\n${summary}\n\nGenerate personalized feedback and action items for this physician.` }],
      }),
      callClaude(ANTHROPIC_KEY, {
        model: "claude-sonnet-4-5",
        max_tokens: 200,
        system: `You are analyzing physician survey responses for a CC billing improvement program. Return ONLY a JSON object with these fields:
- tier: "high_performer" | "developing" | "needs_attention"
- primary_gap: one short phrase (max 6 words) describing the main issue
- recommended_action: one of these exact strings: "send_dot_phrases" | "schedule_chart_review" | "send_email_series" | "reinforce_and_share" | "no_action_needed"
- priority: "high" | "medium" | "low"
Return only valid JSON, no other text.`,
        messages: [{ role: "user", content: summary }],
      }),
    ]);

    const feedback = feedbackData.content?.[0]?.text || "";
    let tag = {};
    try { tag = JSON.parse(tagData.content?.[0]?.text || "{}"); } catch {}

    const actionSummary = tag.recommended_action
      ? `${tag.tier} | ${tag.primary_gap} | ${tag.recommended_action} | priority: ${tag.priority}`
      : "";

    return res.status(200).json({ feedback, actionSummary });
  } catch (e) {
    const msg = e.name === "AbortError" ? "Feedback generation timed out" : (e.message || "Generation failed");
    console.error("feedback.js error:", msg);
    return res.status(500).json({ error: msg });
  }
}
