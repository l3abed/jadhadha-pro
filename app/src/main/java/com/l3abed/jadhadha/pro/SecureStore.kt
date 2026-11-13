package com.l3abed.jadhadha.pro

import android.content.Context
import android.security.keystore.KeyGenParameterSpec
import android.security.keystore.KeyProperties
import android.util.Base64
import java.security.KeyStore
import javax.crypto.Cipher
import javax.crypto.KeyGenerator
import javax.crypto.SecretKey
import javax.crypto.spec.GCMParameterSpec

/**
 * تخزين مفتاح Gemini مشفّراً بمفتاح داخل Android Keystore.
 *
 * لا يُحفظ المفتاح نصاً في SharedPreferences. عند أول تشغيل تُرحَّل القيمة
 * القديمة (إن وُجدت) ثم تُمحى من التخزين العادي.
 */
object SecureStore {
    private const val ANDROID_KEYSTORE = "AndroidKeyStore"
    private const val KEY_ALIAS = "jadhadha_gemini_key"
    private const val PREFS = "settings_secure"
    private const val CIPHER_KEY = "gemini_api_key"
    private const val LEGACY_PREFS = "settings"
    private const val LEGACY_KEY = "api_key"
    private const val TRANSFORMATION = "AES/GCM/NoPadding"
    private const val TAG_BITS = 128

    private fun prefs(ctx: Context) = ctx.getSharedPreferences(PREFS, Context.MODE_PRIVATE)

    private fun key(): SecretKey {
        val ks = KeyStore.getInstance(ANDROID_KEYSTORE).apply { load(null) }
        try {
            (ks.getEntry(KEY_ALIAS, null) as? KeyStore.SecretKeyEntry)?.let { return it.secretKey }
        } catch (_: Exception) {
            // مفتاح قديم تالف: نحذفه وننشئ بديلاً.
            try { ks.deleteEntry(KEY_ALIAS) } catch (_: Exception) { }
        }
        val gen = KeyGenerator.getInstance(KeyProperties.KEY_ALGORITHM_AES, ANDROID_KEYSTORE)
        gen.init(
            KeyGenParameterSpec.Builder(
                KEY_ALIAS,
                KeyProperties.PURPOSE_ENCRYPT or KeyProperties.PURPOSE_DECRYPT,
            )
                .setBlockModes(KeyProperties.BLOCK_MODE_GCM)
                .setEncryptionPaddings(KeyProperties.ENCRYPTION_PADDING_NONE)
                .build(),
        )
        return gen.generateKey()
    }

    fun load(ctx: Context): String {
        val stored = prefs(ctx).getString(CIPHER_KEY, "") ?: ""
        if (stored.isNotBlank()) {
            try {
                val raw = Base64.decode(stored, Base64.NO_WRAP)
                val iv = raw.copyOfRange(0, 12)
                val body = raw.copyOfRange(12, raw.size)
                val cipher = Cipher.getInstance(TRANSFORMATION)
                cipher.init(Cipher.DECRYPT_MODE, key(), GCMParameterSpec(TAG_BITS, iv))
                return String(cipher.doFinal(body), Charsets.UTF_8)
            } catch (_: Exception) {
                // مفتاح Keystore تالف أو أُلغي: نبدأ من قيمة فارغة بدل الانهيار.
                prefs(ctx).edit().remove(CIPHER_KEY).apply()
            }
        }
        // ترحيل المفتاح القديم المخزَّن نصاً، ثم محوه.
        val legacy = ctx.getSharedPreferences(LEGACY_PREFS, Context.MODE_PRIVATE)
            .getString(LEGACY_KEY, "") ?: ""
        if (legacy.isNotBlank()) {
            save(ctx, legacy)
            ctx.getSharedPreferences(LEGACY_PREFS, Context.MODE_PRIVATE)
                .edit().remove(LEGACY_KEY).apply()
            return legacy
        }
        return ""
    }

    fun save(ctx: Context, value: String) {
        if (value.isBlank()) {
            prefs(ctx).edit().remove(CIPHER_KEY).apply()
            return
        }
        val cipher = Cipher.getInstance(TRANSFORMATION)
        cipher.init(Cipher.ENCRYPT_MODE, key())
        val body = cipher.doFinal(value.toByteArray(Charsets.UTF_8))
        val raw = cipher.iv + body
        prefs(ctx).edit().putString(CIPHER_KEY, Base64.encodeToString(raw, Base64.NO_WRAP)).apply()
        ctx.getSharedPreferences(LEGACY_PREFS, Context.MODE_PRIVATE)
            .edit().remove(LEGACY_KEY).apply()
    }
}
