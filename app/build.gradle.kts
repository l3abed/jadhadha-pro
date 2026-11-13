import java.util.Properties
import org.gradle.api.GradleException

plugins {
  id("com.android.application")
  id("org.jetbrains.kotlin.plugin.compose")
  id("com.chaquo.python")
}

// بيانات التوقيع لا تُخزَّن في المستودع: keystore.properties محلياً (مُستثنى من Git)
// أو متغيرات البيئة في GitHub Actions.
val keystoreProps = Properties().apply {
  val f = rootProject.file("keystore.properties")
  if (f.exists()) f.inputStream().use { load(it) }
}
fun signingValue(prop: String, env: String): String? =
  (keystoreProps.getProperty(prop) ?: System.getenv(env))?.takeIf { it.isNotBlank() }

val releaseStorePath = signingValue("storeFile", "JADHADHA_STORE_FILE")
val releaseStorePassword = signingValue("storePassword", "JADHADHA_STORE_PASSWORD")
val releaseKeyAlias = signingValue("keyAlias", "JADHADHA_KEY_ALIAS")
val releaseKeyPassword = signingValue("keyPassword", "JADHADHA_KEY_PASSWORD")
val hasReleaseSigning = listOf(releaseStorePath, releaseStorePassword, releaseKeyAlias, releaseKeyPassword).all { it != null }
val ciVersionCode = (findProperty("versionCode") as String?)?.toIntOrNull() ?: 2300
val ciVersionName = (findProperty("versionName") as String?)?.takeIf { it.isNotBlank() } ?: "2.0.0-dev"

if (gradle.startParameter.taskNames.any { it.contains("Release", ignoreCase = true) } && !hasReleaseSigning) {
  throw GradleException(
    "بيانات توقيع الإصدار غير متوفرة. أضف keystore.properties محلياً أو متغيرات " +
    "JADHADHA_STORE_FILE / JADHADHA_STORE_PASSWORD / JADHADHA_KEY_ALIAS / JADHADHA_KEY_PASSWORD."
  )
}

android {
  namespace = "com.l3abed.jadhadha.pro"
  compileSdk = 36

  defaultConfig {
    applicationId = "com.l3abed.jadhadha.pro"
    minSdk = 24
    targetSdk = 36
    versionCode = ciVersionCode
    versionName = ciVersionName

    ndk {
      // Python 3.12+ يدعم 64 بت فقط
      abiFilters += listOf("arm64-v8a", "x86_64")
    }
  }

  signingConfigs {
    create("release") {
      if (hasReleaseSigning) {
        storeFile = file(releaseStorePath!!)
        storePassword = releaseStorePassword
        keyAlias = releaseKeyAlias
        keyPassword = releaseKeyPassword
      }
    }
  }

  buildTypes {
    release {
      isMinifyEnabled = false
      signingConfig = if (hasReleaseSigning) signingConfigs.getByName("release") else null
    }
    debug {
      // يستعمل مفتاح debug الافتراضي؛ لا نستخدم مفتاح الإصدار للتطوير.
    }
  }

  compileOptions {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
  }

  buildFeatures {
    compose = true
    buildConfig = true
  }
}

chaquopy {
  defaultConfig {
    version = "3.13"
    pip {
      install("python-docx==1.2.0")
    }
  }
}

dependencies {
  implementation(platform("androidx.compose:compose-bom:2024.09.00"))
  implementation("androidx.activity:activity-compose:1.10.1")
  implementation("androidx.compose.material3:material3")
  implementation("androidx.compose.ui:ui")
  implementation("androidx.core:core-ktx:1.15.0")
  implementation("androidx.documentfile:documentfile:1.0.1")
  implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.10.2")
}
