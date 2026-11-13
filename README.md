# الجذاذة النموذجية 2.0

تطبيق أندرويد يفرّغ الجذاذات المصدرية (docx) في القوالب النموذجية، بثلاثة مكوّنات:
القراءة، الدرس اللغوي، التعبير والإنشاء.

- يشغّل سكريبتات Python نفسها (python-docx) داخل التطبيق عبر Chaquopy، فالناتج مطابق للسكريبت.
- القوالب مضمَّنة في `app/src/main/assets/models/`.
- لا مفتاح داخل الكود؛ مفتاح Gemini اختياري ويُدخل من الإعدادات، ويُخزَّن مشفّراً في Android Keystore.
- الكود المصدري موجود مباشرة في المستودع (لم يعد داخل `jadhadha-pro.zip`).

## البناء محلياً

```bash
./gradlew assembleDebug
```

لبناء نسخة إصدار موقّعة محلياً، أنشئ ملف `keystore.properties` في جذر المشروع
(وهو مُستثنى من Git) بالشكل التالي:

```properties
storeFile=/absolute/path/to/release.jks
storePassword=...
keyAlias=jadhadha
keyPassword=...
```

أو مرّر المتغيرات: `JADHADHA_STORE_FILE`، `JADHADHA_STORE_PASSWORD`،
`JADHADHA_KEY_ALIAS`، `JADHADHA_KEY_PASSWORD`.

> **مهم:** تم تدوير مفتاح التوقيع القديم. المستخدمون الذين ثبّتوا نسخة سابقة
> يحتاجون إلى إلغاء تثبيتها قبل تثبيت النسخة الجديدة، لأن توقيع التطبيق تغيّر.

## أسرار GitHub Actions

أضف هذه الأسرار في إعدادات المستودع قبل تشغيل سير العمل:

| السر | الوصف |
|---|---|
| `ANDROID_KEYSTORE_BASE64` | ملف `release.jks` الجديد بترميز Base64 |
| `ANDROID_KEYSTORE_PASSWORD` | كلمة مرور المخزن |
| `ANDROID_KEY_ALIAS` | اسم المفتاح، غالباً `jadhadha` |
| `ANDROID_KEY_PASSWORD` | كلمة مرور المفتاح |

لإنتاج قيمة Base64:

```bash
base64 -w0 release.jks > release.jks.b64
```

## الاختبارات

```bash
PYTHONPATH=app/src/main/python python3 -m unittest discover -s tests -v
```

## الاستعمال من الحاسوب

```bash
PYTHONPATH=app/src/main/python python3 app/src/main/python/bridge.py \
  reading app/src/main/assets/models/reading.docx source.docx out
```

## ملاحظات الأمان

- لا تُرفع ملفات `*.jks` أو `*.keystore` أو `keystore.properties` أو `.env`.
- يفحص سير العمل المستودع بحثاً عن التسريبات وملفات التوقيع قبل البناء.
- جلب الروابط يقبل `https` فقط، ويمنع العناوين المحلية والشبكات الداخلية.
- النص المستورد يُعامل كبيانات لا كتعليمات، ويُغلَّف بوسم `<source_text>`.
