# وسيط الذكاء الاصطناعي (Cloudflare Worker)

عامل صغير يحمل مفتاح Gemini سرًّا ويمرّر الطلبات إليه، فلا يحتاج مستخدم التطبيق
إلى أي مفتاح. يحدّ الاستعمال يوميًا لكل جهاز ولكل IP.

> **منشور حاليًا على:** `https://jadhadha-ai.jadhadha.workers.dev`
> (الحساب `55b5c1cab5600bbb60e47f99cbe36f95`، مخزن KV: `ee8c13a74dd749a386b2e20e0dc8e636`)

## المتطلبات

- حساب Cloudflare مجاني.
- Node.js و`wrangler` (أداة Cloudflare).

## النشر خطوة بخطوة

```bash
cd server
npm i -g wrangler                 # مرة واحدة
wrangler login                    # أو: export CLOUDFLARE_API_TOKEN=...

# 1) أنشئ مخزن العدّادات وانسخ المعرّف الناتج إلى wrangler.toml
wrangler kv namespace create RATE_KV

# 2) ضع الأسرار (لن تُحفظ في الملفات)
wrangler secret put GEMINI_KEY    # ألصق مفتاح Gemini
wrangler secret put APP_TOKEN     # رمز عشوائي يشاركه التطبيق

# 3) انشر
wrangler deploy
```

بعد النشر يطبع wrangler عنوان العامل، مثل:

```
https://jadhadha-ai.<account>.workers.dev
```

ضع هذا العنوان في التطبيق ضمن `AI_PROXY_BASE`، ونفس `APP_TOKEN` في `AI_PROXY_TOKEN`
(في `MainActivity.kt`)، ثم أعد بناء التطبيق.

## الاختبار

```bash
curl -sS -X POST \
  -H "x-goog-api-key: $APP_TOKEN" -H "Content-Type: application/json" \
  -d '{"contents":[{"parts":[{"text":"قل: مرحبا"}]}]}' \
  "https://jadhadha-ai.<account>.workers.dev/v1beta/models/gemini-3.5-flash:generateContent"
```

## ملاحظات

- الحدّ الافتراضي: 40 طلبًا/يوم لكل جهاز، و300 لكل IP. عدّلهما في `wrangler.toml`.
- `APP_TOKEN` اختياري لكن مُستحسن؛ بدونه يصبح الرابط مفتوحًا (يبقى محدودًا بالمعدّل).
- المفتاح لا يُسجَّل ولا يُرسل للتطبيق أبدًا؛ يبقى داخل العامل.
- لتدوير المفتاح لاحقًا: `wrangler secret put GEMINI_KEY` بلا إعادة نشر.
