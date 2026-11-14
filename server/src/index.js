/**
 * وسيط الذكاء الاصطناعي (Cloudflare Worker)
 * ==========================================
 * يحمل مفتاح Gemini سرًّا (GEMINI_KEY) ويمرّر الطلبات إلى Google، فلا يحتاج
 * المستخدم إلى مفتاح. يحدّ الاستعمال لكل جهاز ولكل IP لمنع استنزاف المفتاح.
 *
 * الطلب من التطبيق يطابق مسار Gemini تمامًا:
 *   POST /v1beta/models/<model>:generateContent
 * مع ترويسة x-goog-api-key تحمل رمز التطبيق (APP_TOKEN).
 */

const GEMINI_BASE = "https://generativelanguage.googleapis.com";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "content-type, x-goog-api-key, x-app-token, x-device-id",
};

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...CORS },
  });
}

/** عدّاد نافذة ثابتة في KV: يعيد true إن تجاوز الحدّ. */
async function overLimit(kv, key, limit, ttl) {
  const cur = parseInt((await kv.get(key)) || "0", 10);
  if (cur >= limit) return true;
  await kv.put(key, String(cur + 1), { expirationTtl: ttl });
  return false;
}

const PATH_RE = /^\/v1beta\/models\/[A-Za-z0-9._-]+:(generateContent|countTokens)$/;

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS });
    }
    if (url.pathname === "/" || url.pathname === "/health") {
      return json({ ok: true, service: "jadhadha-ai" });
    }
    if (request.method !== "POST") {
      return json({ error: "method_not_allowed" }, 405);
    }
    if (!PATH_RE.test(url.pathname)) {
      return json({ error: "not_found" }, 404);
    }
    if (!env.GEMINI_KEY) {
      return json({ error: "server_not_configured" }, 500);
    }

    // رمز التطبيق (اختياري): يردع الاستعمال العشوائي. يقبل x-app-token أو x-goog-api-key.
    if (env.APP_TOKEN) {
      const tok = request.headers.get("x-app-token") || request.headers.get("x-goog-api-key") || "";
      if (tok !== env.APP_TOKEN) {
        return json({ error: "unauthorized" }, 401);
      }
    }

    // تحديد المعدّل اليومي: لكل جهاز ولكل IP.
    const ip = request.headers.get("cf-connecting-ip") || "0.0.0.0";
    const device = (request.headers.get("x-device-id") || "").trim();
    const day = new Date().toISOString().slice(0, 10);
    const perDevice = parseInt(env.DEVICE_DAILY_LIMIT || "40", 10);
    const perIp = parseInt(env.IP_DAILY_LIMIT || "300", 10);

    if (env.RATE_KV) {
      if (device && (await overLimit(env.RATE_KV, `d:${day}:${device}`, perDevice, 172800))) {
        return json({ error: "device_daily_limit_reached", limit: perDevice }, 429);
      }
      if (await overLimit(env.RATE_KV, `i:${day}:${ip}`, perIp, 172800)) {
        return json({ error: "ip_daily_limit_reached", limit: perIp }, 429);
      }
    }

    let upstream;
    try {
      upstream = await fetch(GEMINI_BASE + url.pathname + url.search, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-goog-api-key": env.GEMINI_KEY,
        },
        body: request.body,
      });
    } catch (e) {
      return json({ error: "upstream_unreachable", detail: String(e) }, 502);
    }

    return new Response(upstream.body, {
      status: upstream.status,
      headers: { "Content-Type": "application/json; charset=utf-8", ...CORS },
    });
  },
};
