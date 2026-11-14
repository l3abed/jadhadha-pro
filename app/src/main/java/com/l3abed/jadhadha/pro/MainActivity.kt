package com.l3abed.jadhadha.pro

import android.content.ActivityNotFoundException
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.provider.OpenableColumns
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.SystemBarStyle
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxHeight
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.heightIn
import androidx.compose.foundation.layout.imePadding
import androidx.compose.foundation.layout.navigationBarsPadding
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.statusBarsPadding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.ButtonDefaults
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.LinearProgressIndicator
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Surface
import androidx.compose.material3.Switch
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.runtime.CompositionLocalProvider
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableIntStateOf
import androidx.compose.runtime.mutableStateMapOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalLayoutDirection
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.PasswordVisualTransformation
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.LayoutDirection
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties
import androidx.core.content.FileProvider
import androidx.documentfile.provider.DocumentFile
import com.chaquo.python.Python
import java.io.File
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.json.JSONObject

private val APP_VERSION: String = BuildConfig.VERSION_NAME
private const val DOCX_MIME =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

private const val ROOT_NAME = "جذاذات الثالثة إعدادي"
private const val GEMINI_KEY_URL = "https://aistudio.google.com/app/apikey"
private val COMP_FOLDERS = listOf("القراءة", "الدرس اللغوي", "التعبير والإنشاء")
private val UNIT_FOLDERS = listOf("الأولى", "الثانية", "الثالثة", "الرابعة", "الخامسة", "السادسة")
    .map { "الوحدة $it" }

// ───── الهوية البصرية: ورق دافئ، أخضر زمردي، ذهب ─────
// لوحة الألوان: نهارية خضراء، وليلية زرقاء (المتغيرات تُحدَّث عند تبديل الوضع)
private var Emerald = Color(0xFF1B5E20)
private var EmeraldDark = Color(0xFF0C3512)
private var EmeraldSoft = Color(0xFFE4EFE2)
private var Gold = Color(0xFFB8860B)
private var Paper = Color(0xFFFAF7F0)
private var CardBg = Color.White
private var Ink = Color(0xFF1E2A22)
private var Muted = Color(0xFF5E6A62)
private var Line = Color(0xFFDDD7C6)
private var ErrorRed = Color(0xFFB3261E)
private var ErrorSoft = Color(0xFFFBE9E7)
private val DarkMode = mutableStateOf(false)

private fun setPalette(dark: Boolean) {
    if (dark) {
        Emerald = Color(0xFF4A9EFF)
        EmeraldDark = Color(0xFF0B2A52)
        EmeraldSoft = Color(0xFF1B2F4D)
        Gold = Color(0xFFE0B252)
        Paper = Color(0xFF0E141C)
        CardBg = Color(0xFF18212D)
        Ink = Color(0xFFE6EDF5)
        Muted = Color(0xFF9AA8B8)
        Line = Color(0xFF334357)
        ErrorRed = Color(0xFFFF8A80)
        ErrorSoft = Color(0xFF3A1F20)
    } else {
        Emerald = Color(0xFF1B5E20)
        EmeraldDark = Color(0xFF0C3512)
        EmeraldSoft = Color(0xFFE4EFE2)
        Gold = Color(0xFFB8860B)
        Paper = Color(0xFFFAF7F0)
        CardBg = Color.White
        Ink = Color(0xFF1E2A22)
        Muted = Color(0xFF5E6A62)
        Line = Color(0xFFDDD7C6)
        ErrorRed = Color(0xFFB3261E)
        ErrorSoft = Color(0xFFFBE9E7)
    }
}

private data class Comp(val id: String, val label: String)

private val COMPS = listOf(
    Comp("reading", "القراءة"),
    Comp("language", "الدرس اللغوي"),
    Comp("expression", "التعبير والإنشاء"),
)

private data class TabState(
    val sourceName: String = "",
    val sourcePath: String = "",
    val busy: Boolean = false,
    val message: String = "",
    val ok: Boolean = false,
    val outPath: String = "",
    val basePath: String = "",          // نسخة أساسية بلا نمط لإعادة الأنماط دون إعادة التفريغ
    val unitChoice: String = "",        // الوحدة المختارة يدوياً عند غيابها في الترويسة
    val log: String = "",
    val folderComp: String = "",
    val folderUnit: String = "",
    val saveMsg: String = "",
    val activity: String = "",
    val url: String = "",
    val importText: String = "",
    val stage: String = "",
    val importMsg: String = "",
)

private fun TabState.toJson(): String = JSONObject().apply {
    put("sourceName", sourceName)
    put("sourcePath", sourcePath)
    put("message", message)
    put("ok", ok)
    put("outPath", outPath)
    put("basePath", basePath)
    put("unitChoice", unitChoice)
    put("log", log)
    put("folderComp", folderComp)
    put("folderUnit", folderUnit)
    put("saveMsg", saveMsg)
    put("activity", activity)
    put("url", url)
    put("importText", importText)
    put("stage", stage)
    put("importMsg", importMsg)
}.toString()

private fun tabStateFromJson(raw: String?): TabState {
    if (raw.isNullOrBlank()) return TabState()
    return runCatching {
        val j = JSONObject(raw)
        TabState(
            sourceName = j.optString("sourceName"),
            sourcePath = j.optString("sourcePath"),
            busy = false,                       // لا نستعيد عملية كانت جارية عند موت النشاط
            message = j.optString("message"),
            ok = j.optBoolean("ok"),
            outPath = j.optString("outPath"),
            basePath = j.optString("basePath"),
            unitChoice = j.optString("unitChoice"),
            log = j.optString("log"),
            folderComp = j.optString("folderComp"),
            folderUnit = j.optString("folderUnit"),
            saveMsg = j.optString("saveMsg"),
            activity = j.optString("activity"),
            url = j.optString("url"),
            importText = j.optString("importText"),
            stage = j.optString("stage"),
            importMsg = j.optString("importMsg"),
        )
    }.getOrDefault(TabState())
}

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge(statusBarStyle = SystemBarStyle.dark(android.graphics.Color.TRANSPARENT))
        DarkMode.value = getSharedPreferences("settings", Context.MODE_PRIVATE).getBoolean("dark", false)
        setContent {
            JadhadhaTheme {
                CompositionLocalProvider(LocalLayoutDirection provides LayoutDirection.Rtl) {
                    Surface(modifier = Modifier.fillMaxSize(), color = Paper) { JadhadhaApp() }
                }
            }
        }
    }
}

@Composable
private fun JadhadhaTheme(content: @Composable () -> Unit) {
    val dark = DarkMode.value
    setPalette(dark)
    val scheme = if (dark) {
        darkColorScheme(
            primary = Emerald, onPrimary = Color.White,
            secondary = Gold, onSecondary = Color.White,
            background = Paper, onBackground = Ink,
            surface = CardBg, onSurface = Ink, error = ErrorRed,
        )
    } else {
        lightColorScheme(
            primary = Emerald, onPrimary = Color.White,
            secondary = Gold, onSecondary = Color.White,
            background = Paper, onBackground = Ink,
            surface = CardBg, onSurface = Ink, error = ErrorRed,
        )
    }
    MaterialTheme(colorScheme = scheme, content = content)
}

@Composable
private fun JadhadhaApp() {
    val ctx = LocalContext.current
    val scope = rememberCoroutineScope()
    val prefs = remember { ctx.getSharedPreferences("settings", Context.MODE_PRIVATE) }

    var apiKey by remember { mutableStateOf(SecureStore.load(ctx)) }
    var hTeacher by remember { mutableStateOf(prefs.getString("h_teacher", "") ?: "") }
    var hSchool by remember { mutableStateOf(prefs.getString("h_school", "") ?: "") }
    var hLevel by remember { mutableStateOf(prefs.getString("h_level", "") ?: "") }
    var styleKey by remember { mutableStateOf(prefs.getString("style", "classic") ?: "classic") }
    var aiEnabled by remember { mutableStateOf(prefs.getBoolean("ai_on", true)) }
    var showSettings by rememberSaveable { mutableStateOf(false) }
    var showStyleDialog by rememberSaveable { mutableStateOf(false) }
    var headerDirty by remember { mutableStateOf(false) }
    var showLog by rememberSaveable { mutableStateOf(false) }
    var tab by rememberSaveable { mutableIntStateOf(0) }
    var pendingComp by rememberSaveable { mutableStateOf(COMPS[0].id) }
    val states = remember {
        mutableStateMapOf<String, TabState>().apply {
            COMPS.forEach { c ->
                prefs.getString("state_${c.id}", null)?.let { put(c.id, tabStateFromJson(it)) }
            }
        }
    }
    var treeUri by remember { mutableStateOf(prefs.getString("tree_uri", null)) }
    var storageMsg by remember { mutableStateOf("") }

    fun update(id: String, f: (TabState) -> TabState) {
        val next = f(states[id] ?: TabState())
        states[id] = next
        prefs.edit().putString("state_$id", next.toJson()).apply()
    }

    val picker = rememberLauncherForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? ->
        val id = pendingComp
        if (uri != null) {
            scope.launch {
                val res = withContext(Dispatchers.IO) { runCatching { importSource(ctx, uri, id) } }
                res.onSuccess { pair ->
                    update(id) { TabState(sourceName = pair.first, sourcePath = pair.second) }
                }
                res.onFailure { e ->
                    update(id) { it.copy(ok = false, message = "تعذّر فتح الملف: ${e.message}") }
                }
            }
        }
    }

    val treePicker = rememberLauncherForActivityResult(ActivityResultContracts.OpenDocumentTree()) { uri: Uri? ->
        if (uri != null) {
            try {
                ctx.contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION,
                )
            } catch (e: Exception) {
                // نتابع حتى إن لم يمكن حفظ الإذن بشكل دائم
            }
            prefs.edit().putString("tree_uri", uri.toString()).apply()
            treeUri = uri.toString()
            storageMsg = "جارٍ إنشاء المجلدات…"
            scope.launch {
                val res = withContext(Dispatchers.IO) { runCatching { ensureTree(ctx, uri) } }
                storageMsg = if (res.isSuccess) {
                    "تمّ إنشاء «$ROOT_NAME» بثلاثة مجلدات للمكوّنات، وفي كل منها ستة مجلدات للوحدات ✔"
                } else {
                    "تعذّر إنشاء المجلدات: ${res.exceptionOrNull()?.message}"
                }
            }
        }
    }

    fun chooseUnit(id: String, unit: String) {
        val st = states[id] ?: return
        scope.launch {
            withContext(Dispatchers.IO) {
                runCatching {
                    Python.getInstance().getModule("bridge")
                        .callAttr("apply_header_json", st.outPath, JSONObject().put("unit", unit).toString())
                }
            }
            update(id) { it.copy(folderUnit = "الوحدة $unit", unitChoice = unit) }
        }
    }

    fun fetchLinkText(id: String) {
        val st = states[id] ?: return
        if (st.url.isBlank()) return
        update(id) { it.copy(importMsg = "جارٍ جلب النص من الرابط…") }
        scope.launch {
            val res = withContext(Dispatchers.IO) {
                runCatching { Python.getInstance().getModule("bridge").callAttr("fetch_link", st.url.trim()).toString() }
            }
            res.onSuccess { raw ->
                val j = JSONObject(raw)
                if (j.optBoolean("ok")) {
                    update(id) { it.copy(importText = j.optString("text"), importMsg = "جُلب النص، راجعه ثم أدرجه ✔") }
                } else {
                    update(id) { it.copy(importMsg = j.optString("error")) }
                }
            }
            res.onFailure { e -> update(id) { it.copy(importMsg = "خطأ: ${e.message}") } }
        }
    }

    fun transcribeImport(id: String) {
        val st = states[id] ?: return
        if (st.busy) return
        if (st.importText.isBlank()) {
            update(id) { it.copy(importMsg = "الصق النص أو اجلبه من الرابط أولاً.") }
            return
        }
        if (id == "expression" && st.activity.isBlank()) {
            update(id) { it.copy(importMsg = "اختر نشاط التعبير أولاً (في بطاقة التوليد).") }
            return
        }
        update(id) { it.copy(importMsg = "جارٍ تحويل النص وتفريغه في القالب…") }
        val key = if (aiEnabled) apiKey.trim() else ""
        scope.launch {
            val res = withContext(Dispatchers.IO) {
                runCatching {
                    val model = ensureModel(ctx, id)
                    val outDir = File(ctx.filesDir, "import/$id").apply { mkdirs() }
                    Python.getInstance().getModule("bridge").callAttr(
                        "transcribe_text", id, st.importText, st.activity, key,
                        outDir.absolutePath, model.absolutePath, styleKey, headerJson(hTeacher, hSchool, hLevel),
                    ).toString()
                }
            }
            res.onSuccess { raw ->
                val j = JSONObject(raw)
                if (j.optBoolean("ok")) {
                    update(id) {
                        it.copy(
                            ok = true, outPath = j.optString("path"), basePath = j.optString("base"), message = "",
                            folderComp = j.optJSONObject("target")?.optString("component") ?: "",
                            folderUnit = j.optJSONObject("target")?.optString("unit") ?: "",
                            importMsg = "فُرِّغ النص في القالب ✔",
                            log = j.optString("log"),
                        )
                    }
                } else {
                    update(id) { it.copy(importMsg = "تعذّر التفريغ: " + j.optString("error")) }
                }
            }
            res.onFailure { e -> update(id) { it.copy(importMsg = "خطأ: ${e.message}") } }
        }
    }

    fun saveComp(id: String) {
        val st = states[id] ?: return
        val uriStr = treeUri
        if (uriStr == null) {
            update(id) { it.copy(saveMsg = "اختر أولاً مكان التخزين من «الإعدادات».") }
            showSettings = true
            return
        }
        scope.launch {
            val res = withContext(Dispatchers.IO) {
                runCatching { storeDocx(ctx, Uri.parse(uriStr), st.outPath, st.folderComp, st.folderUnit) }
            }
            if (res.isSuccess) {
                update(id) { it.copy(saveMsg = "خُزّنت في: ${res.getOrNull()} ✔") }
            } else {
                update(id) { it.copy(saveMsg = "تعذّر التخزين: ${res.exceptionOrNull()?.message}") }
            }
        }
    }

    fun runComp(id: String) {
        val st = states[id] ?: return
        if (st.sourcePath.isEmpty() || st.busy) return
        update(id) {
            it.copy(busy = true, ok = false, message = "جارٍ التفريغ… قد يستغرق ثوانيَ.", outPath = "", log = "", saveMsg = "")
        }
        scope.launch {
            val res = withContext(Dispatchers.Default) {
                runCatching { runBridge(ctx, id, st.sourcePath, st.sourceName, if (aiEnabled) apiKey.trim() else "", headerJson(hTeacher, hSchool, hLevel), styleKey) }
            }
            res.onSuccess { j ->
                val ok = j.optBoolean("ok")
                update(id) {
                    it.copy(
                        busy = false,
                        ok = ok,
                        outPath = j.optString("path"),
                        basePath = j.optString("base"),
                        message = if (ok) "تمّ إنجاز الجذاذة ✔" else "تعذّر التفريغ: " + j.optString("error"),
                        log = j.optString("log"),
                        folderComp = j.optJSONObject("target")?.optString("component") ?: "",
                        folderUnit = j.optJSONObject("target")?.optString("unit") ?: "",
                    )
                }
            }
            res.onFailure { e ->
                update(id) { it.copy(busy = false, ok = false, message = "خطأ: ${e.message}") }
            }
        }
    }

    /** يغيّر نمط الجذاذة من النسخة الأساسية المحفوظة، بلا إعادة تفريغ ولا استهلاك للمفتاح. */
    fun restyleCurrent(
        id: String,
        style: String,
        note: String = "تمّ تغيير النمط من النسخة المحفوظة دون إعادة التفريغ ✔",
    ) {
        val cur = states[id] ?: return
        val base = cur.basePath
        if (cur.busy || base.isEmpty() || !File(base).exists()) return
        update(id) { it.copy(message = "جارٍ تطبيق النمط على الجذاذة المحفوظة…") }
        scope.launch {
            val res = withContext(Dispatchers.IO) {
                runCatching {
                    Python.getInstance().getModule("bridge").callAttr(
                        "restyle", base, cur.outPath, headerJson(hTeacher, hSchool, hLevel, cur.unitChoice), style,
                    ).toString()
                }
            }
            res.onSuccess { raw ->
                val j = JSONObject(raw)
                if (j.optBoolean("ok")) {
                    update(id) {
                        it.copy(
                            ok = true,
                            outPath = j.optString("path"),
                            basePath = j.optString("base").ifEmpty { base },
                            message = note,
                            folderComp = j.optJSONObject("target")?.optString("component")?.takeIf { s -> s.isNotEmpty() } ?: it.folderComp,
                            folderUnit = j.optJSONObject("target")?.optString("unit")?.takeIf { s -> s.isNotEmpty() } ?: it.folderUnit,
                            log = j.optString("log"),
                        )
                    }
                } else {
                    update(id) { it.copy(message = "تعذّر تطبيق النمط: " + j.optString("error")) }
                }
            }
            res.onFailure { e -> update(id) { it.copy(message = "خطأ: ${e.message}") } }
        }
    }

    val comp = COMPS[tab]
    val st = states[comp.id] ?: TabState()

    Column(Modifier.fillMaxSize()) {
        Hero(showSettings) { showSettings = true }

        Column(
            Modifier
                .weight(1f)
                .verticalScroll(rememberScrollState())
                .navigationBarsPadding()
                .imePadding()
                .padding(horizontal = 16.dp, vertical = 16.dp),
            verticalArrangement = Arrangement.spacedBy(14.dp),
        ) {
            ComponentSelector(tab) { tab = it }

            SoftCard {
                StepTitle("١", "اختر الجذاذة المصدر")
                OutlinedButton(
                    onClick = {
                        pendingComp = comp.id
                        picker.launch(arrayOf("*/*"))
                    },
                    enabled = !st.busy,
                    shape = RoundedCornerShape(14.dp),
                    border = BorderStroke(1.5.dp, Emerald),
                    colors = ButtonDefaults.outlinedButtonColors(contentColor = Emerald),
                    modifier = Modifier.fillMaxWidth().height(52.dp),
                ) {
                    Text(
                        if (st.sourceName.isEmpty()) "اختيار ملف docx" else "✔  " + st.sourceName,
                        maxLines = 1,
                        fontWeight = FontWeight.Medium,
                    )
                }

                StepTitle("٢", "اختر النمط")
                OutlinedButton(
                    onClick = { showStyleDialog = true },
                    enabled = !st.busy,
                    shape = RoundedCornerShape(14.dp),
                    border = BorderStroke(1.5.dp, Emerald),
                    colors = ButtonDefaults.outlinedButtonColors(contentColor = Emerald),
                    modifier = Modifier.fillMaxWidth().height(52.dp),
                ) {
                    Text("النمط: " + styleLabel(styleKey), fontWeight = FontWeight.Bold)
                }
                StepTitle("٣", "فرّغها في القالب النموذجي")
                SolidButton(
                    text = "تفريغ الجذاذة",
                    color = Emerald,
                    enabled = st.sourcePath.isNotEmpty() && !st.busy,
                    onClick = { runComp(comp.id) },
                    modifier = Modifier.fillMaxWidth(),
                )

                if (st.busy) {
                    LinearProgressIndicator(
                        modifier = Modifier.fillMaxWidth().clip(RoundedCornerShape(50)),
                        color = Emerald,
                        trackColor = EmeraldSoft,
                    )
                }
                if (st.message.isNotEmpty()) {
                    StatusBox(st.message, if (st.busy) 0 else if (st.ok) 1 else 2)
                }
            }

            if (st.ok && st.outPath.isNotEmpty()) {
                SoftCard {
                    StepTitle("٤", "افتح الجذاذة أو خزّنها")
                    Row(horizontalArrangement = Arrangement.spacedBy(10.dp)) {
                        SolidButton(
                            text = "فتح",
                            color = Emerald,
                            onClick = { openOrShare(ctx, st.outPath, view = true) },
                            modifier = Modifier.weight(1f),
                        )
                        OutlinedButton(
                            onClick = { openOrShare(ctx, st.outPath, view = false) },
                            shape = RoundedCornerShape(14.dp),
                            border = BorderStroke(1.5.dp, Emerald),
                            colors = ButtonDefaults.outlinedButtonColors(contentColor = Emerald),
                            modifier = Modifier.weight(1f).height(52.dp),
                        ) { Text("مشاركة / حفظ", fontWeight = FontWeight.Bold) }
                    }
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        InfoChip("المكوّن: " + st.folderComp.ifEmpty { "—" })
                        InfoChip("الوحدة: " + st.folderUnit.ifEmpty { "—" })
                    }
                    if (st.folderUnit.isEmpty()) {
                        Text("لم أجد الوحدة في الترويسة، اختر الوحدة:", fontSize = 13.sp, color = Muted)
                        UnitPicker { unit -> chooseUnit(comp.id, unit) }
                    }
                    SolidButton(
                        text = "تخزين في «$ROOT_NAME»",
                        color = Gold,
                        onClick = { saveComp(comp.id) },
                        modifier = Modifier.fillMaxWidth(),
                    )
                    if (st.saveMsg.isNotEmpty()) {
                        StatusBox(st.saveMsg, if (st.saveMsg.contains("✔")) 1 else 2)
                    }
                }
            }

            run {
                SoftCard {
                    StepTitle("✦", "توليد الجذاذة من نص أو رابط")
                    Text(
                        "ألصق نص التحضير أو اجلب نص الرابط، فيكمل الذكاء الاصطناعي بقية الجذاذة حوله، ثم تُفرَّغ في القالب المختار. الترويسة (الدرس والمقرر والمستوى) تُستخرج من النص.",
                        fontSize = 13.sp, color = Muted, lineHeight = 20.sp,
                    )
                    if (comp.id == "expression") {
                        Text("النشاط:", fontSize = 13.sp, color = Muted)
                        listOf("الاكتساب", "التطبيق", "الإنتاج", "التصحيح").chunked(2).forEach { row ->
                            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                                row.forEach { act ->
                                    val on = st.activity == act
                                    Box(
                                        Modifier.weight(1f).clip(RoundedCornerShape(12.dp))
                                            .background(if (on) Emerald else CardBg)
                                            .border(1.dp, if (on) Emerald else Line, RoundedCornerShape(12.dp))
                                            .clickable { update(comp.id) { it.copy(activity = act) } }
                                            .padding(vertical = 10.dp),
                                        contentAlignment = Alignment.Center,
                                    ) { Text(act, color = if (on) Color.White else Ink, fontSize = 13.sp) }
                                }
                            }
                        }
                    }
                    OutlinedTextField(
                        value = st.url, onValueChange = { v -> update(comp.id) { it.copy(url = v) } },
                        singleLine = true, label = { Text("رابط التحضير") }, shape = RoundedCornerShape(14.dp),
                        modifier = Modifier.fillMaxWidth(),
                    )
                    OutlinedButton(
                        onClick = { fetchLinkText(comp.id) },
                        enabled = st.url.isNotBlank() && !st.busy,
                        shape = RoundedCornerShape(14.dp),
                        border = BorderStroke(1.5.dp, Emerald),
                        colors = ButtonDefaults.outlinedButtonColors(contentColor = Emerald),
                        modifier = Modifier.fillMaxWidth().height(48.dp),
                    ) { Text("جلب النص من الرابط") }
                    OutlinedTextField(
                        value = st.importText, onValueChange = { v -> update(comp.id) { it.copy(importText = v) } },
                        minLines = 4, label = { Text("نص التحضير (لصق أو بعد الجلب)") },
                        shape = RoundedCornerShape(14.dp), modifier = Modifier.fillMaxWidth(),
                    )
                    SolidButton(
                        text = "توليد وتفريغ في القالب",
                        color = Emerald,
                        enabled = st.importText.isNotBlank() && !st.busy && (comp.id != "expression" || st.activity.isNotBlank()),
                        onClick = { transcribeImport(comp.id) },
                        modifier = Modifier.fillMaxWidth(),
                    )
                    if (st.importMsg.isNotEmpty()) StatusBox(st.importMsg, if (st.importMsg.contains("✔")) 1 else 2)
                }
            }

            if (st.log.isNotEmpty()) {
                TextButton(onClick = { showLog = !showLog }) {
                    Text(if (showLog) "إخفاء التقرير" else "عرض التقرير", color = Muted)
                }
                if (showLog) {
                    Text(st.log.takeLast(4000), fontFamily = FontFamily.Monospace, fontSize = 11.sp, color = Muted)
                }
            }

            Text(
                "الجذاذة النموذجية  ·  محمد العابد  ·  الإصدار $APP_VERSION",
                fontSize = 11.sp,
                color = Muted,
                textAlign = TextAlign.Center,
                modifier = Modifier.fillMaxWidth().padding(top = 8.dp, bottom = 12.dp),
            )
        }
    }

    if (showSettings) {
        SettingsWindow(
            dark = DarkMode.value,
            onDark = { on ->
                DarkMode.value = on
                prefs.edit().putBoolean("dark", on).apply()
            },
            teacher = hTeacher,
            school = hSchool,
            level = hLevel,
            onHeader = { t, sc, lv ->
                if (t != hTeacher || sc != hSchool || lv != hLevel) headerDirty = true
                hTeacher = t
                hSchool = sc
                hLevel = lv
                prefs.edit().putString("h_teacher", t).putString("h_school", sc).putString("h_level", lv).apply()
            },
            treeUri = treeUri,
            storageMsg = storageMsg,
            onPickStorage = { treePicker.launch(null) },
            apiKey = apiKey,
            aiOn = aiEnabled,
            onChangeKey = { v ->
                apiKey = v
                SecureStore.save(ctx, v)
            },
            onToggleAi = { on ->
                aiEnabled = on
                prefs.edit().putBoolean("ai_on", on).apply()
            },
            onClose = {
                showSettings = false
                if (headerDirty) {
                    headerDirty = false
                    restyleCurrent(
                        comp.id,
                        styleKey,
                        "تمّ تحديث الترويسة من النسخة المحفوظة دون إعادة التفريغ ✔",
                    )
                }
            },
        )
    }

    if (showStyleDialog) {
        StyleWindow(
            selected = styleKey,
            onPick = { k ->
                styleKey = k
                prefs.edit().putString("style", k).apply()
                restyleCurrent(comp.id, k)
            },
            onClose = { showStyleDialog = false },
        )
    }
}

@Composable
private fun Hero(settingsOpen: Boolean, onSettings: () -> Unit) {
    Box(
        Modifier
            .fillMaxWidth()
            .clip(RoundedCornerShape(bottomStart = 30.dp, bottomEnd = 30.dp))
            .background(Brush.verticalGradient(listOf(Emerald, EmeraldDark)))
            .statusBarsPadding()
            .padding(horizontal = 20.dp, vertical = 20.dp),
    ) {
        Row(verticalAlignment = Alignment.CenterVertically) {
            Column(Modifier.weight(1f)) {
                Text(
                    "الجذاذة النموذجية",
                    color = Color.White,
                    fontFamily = FontFamily.Serif,
                    fontWeight = FontWeight.Bold,
                    fontSize = 28.sp,
                )
                Spacer(Modifier.height(6.dp))
                Box(Modifier.width(44.dp).height(3.dp).clip(RoundedCornerShape(2.dp)).background(Gold))
                Spacer(Modifier.height(8.dp))
                Text(
                    "من المصدر إلى القالب النموذجي بلمسة واحدة",
                    color = Color(0xCCFFFFFF),
                    fontSize = 13.sp,
                )
            }
            Box(
                Modifier
                    .clip(RoundedCornerShape(50))
                    .border(1.dp, Color(0x80FFFFFF), RoundedCornerShape(50))
                    .clickable { onSettings() }
                    .padding(horizontal = 14.dp, vertical = 8.dp),
            ) {
                Text(if (settingsOpen) "إغلاق" else "الإعدادات", color = Color.White, fontSize = 13.sp)
            }
        }
    }
}

@Composable
private fun ComponentSelector(selected: Int, onSelect: (Int) -> Unit) {
    Row(Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.spacedBy(8.dp)) {
        COMPS.forEachIndexed { i, c ->
            val on = i == selected
            Box(
                Modifier
                    .weight(1f)
                    .height(48.dp)
                    .clip(RoundedCornerShape(14.dp))
                    .background(if (on) Emerald else CardBg)
                    .border(1.dp, if (on) Emerald else Line, RoundedCornerShape(14.dp))
                    .clickable { onSelect(i) },
                contentAlignment = Alignment.Center,
            ) {
                Text(
                    c.label,
                    color = if (on) Color.White else Ink,
                    fontWeight = if (on) FontWeight.Bold else FontWeight.Medium,
                    fontSize = 13.sp,
                    textAlign = TextAlign.Center,
                )
            }
        }
    }
}

@Composable
private fun SoftCard(content: @Composable () -> Unit) {
    Card(
        modifier = Modifier.fillMaxWidth(),
        shape = RoundedCornerShape(22.dp),
        colors = CardDefaults.cardColors(containerColor = CardBg),
        elevation = CardDefaults.cardElevation(defaultElevation = 3.dp),
    ) {
        Column(Modifier.padding(18.dp), verticalArrangement = Arrangement.spacedBy(12.dp)) {
            content()
        }
    }
}

private val STYLE_CHOICES = listOf(
    "classic" to "الأصلي",
    "emerald" to "زمردي ملوّن",
    "modern" to "عصري",
    "print_bw" to "أبيض وأسود",
    "royal" to "بنفسجي ملكي",
    "sunset" to "برتقالي دافئ",
    "ocean" to "أزرق محيطي",
    "rose" to "وردي",
    "slate" to "رمادي أنيق",
    "sepia" to "بني كلاسيكي",
)

/** يعيد التسمية العربية للنمط المختار، أو التسمية الأولى عند عدم العثور عليه. */
private fun styleLabel(key: String): String =
    STYLE_CHOICES.firstOrNull { it.first == key }?.second ?: STYLE_CHOICES.first().second

@Composable
private fun StylePicker(selected: String, onPick: (String) -> Unit) {
    Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
        STYLE_CHOICES.chunked(2).forEach { row ->
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                row.forEach { (key, label) ->
                    val on = key == selected
                    Box(
                        Modifier
                            .weight(1f)
                            .clip(RoundedCornerShape(12.dp))
                            .background(if (on) Emerald else CardBg)
                            .border(1.dp, if (on) Emerald else Line, RoundedCornerShape(12.dp))
                            .clickable { onPick(key) }
                            .padding(vertical = 12.dp),
                        contentAlignment = Alignment.Center,
                    ) {
                        Text(label, color = if (on) Color.White else Ink, fontSize = 13.sp,
                            fontWeight = if (on) FontWeight.Bold else FontWeight.Medium)
                    }
                }
            }
        }
    }
}

/**
 * نافذة فرعية موحّدة. حين تكون [fullHeight] صحيحة تمتدّ لمعظم الشاشة مع تمرير داخلي (للإعدادات)،
 * وإلا فتلتفّ حول محتواها بحدّ أقصى للارتفاع (لاختيار النمط).
 */
@Composable
private fun SubWindow(
    title: String,
    onClose: () -> Unit,
    fullHeight: Boolean,
    content: @Composable () -> Unit,
) {
    Dialog(onDismissRequest = onClose, properties = DialogProperties(usePlatformDefaultWidth = false)) {
        Surface(
            modifier = Modifier
                .fillMaxWidth(0.94f)
                .then(
                    if (fullHeight) Modifier.fillMaxHeight(0.92f)
                    else Modifier.heightIn(max = 560.dp)
                ),
            shape = RoundedCornerShape(24.dp),
            color = Paper,
            shadowElevation = 8.dp,
        ) {
            Column(
                modifier = if (fullHeight) Modifier.fillMaxSize().padding(18.dp)
                else Modifier.fillMaxWidth().padding(18.dp),
            ) {
                Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.fillMaxWidth()) {
                    Text(
                        title,
                        fontFamily = FontFamily.Serif,
                        fontWeight = FontWeight.Bold,
                        fontSize = 20.sp,
                        color = Ink,
                        modifier = Modifier.weight(1f),
                    )
                    TextButton(onClick = onClose) {
                        Text("إغلاق", color = Emerald, fontWeight = FontWeight.Bold)
                    }
                }
                Spacer(Modifier.height(4.dp))
                if (fullHeight) {
                    Column(
                        Modifier.weight(1f).verticalScroll(rememberScrollState()),
                        verticalArrangement = Arrangement.spacedBy(14.dp),
                    ) { content() }
                } else {
                    Column(
                        Modifier.fillMaxWidth().verticalScroll(rememberScrollState()),
                        verticalArrangement = Arrangement.spacedBy(14.dp),
                    ) { content() }
                }
            }
        }
    }
}

@Composable
private fun StyleWindow(selected: String, onPick: (String) -> Unit, onClose: () -> Unit) {
    SubWindow("اختر النمط", onClose, fullHeight = false) {
        Text(
            "تغيير النمط لا يعيد التفريغ ولا يستهلك مفتاح الذكاء الاصطناعي: يُنقل المحتوى المحفوظ إلى القالب الجديد.",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        StylePicker(selected) { key ->
            onPick(key)
            onClose()
        }
    }
}

@Composable
private fun StepTitle(num: String, text: String) {
    Row(verticalAlignment = Alignment.CenterVertically, horizontalArrangement = Arrangement.spacedBy(10.dp)) {
        Box(
            Modifier.size(28.dp).clip(CircleShape).background(EmeraldSoft),
            contentAlignment = Alignment.Center,
        ) { Text(num, color = Emerald, fontWeight = FontWeight.Bold, fontSize = 14.sp) }
        Text(text, fontWeight = FontWeight.Bold, fontSize = 15.sp, color = Ink)
    }
}

@Composable
private fun SolidButton(
    text: String,
    color: Color,
    onClick: () -> Unit,
    modifier: Modifier = Modifier,
    enabled: Boolean = true,
) {
    Button(
        onClick = onClick,
        enabled = enabled,
        shape = RoundedCornerShape(14.dp),
        colors = ButtonDefaults.buttonColors(
            containerColor = color,
            contentColor = Color.White,
            disabledContainerColor = Color(0xFFCCD3CD),
            disabledContentColor = Color.White,
        ),
        modifier = modifier.height(52.dp),
    ) { Text(text, fontWeight = FontWeight.Bold, fontSize = 15.sp) }
}

@Composable
private fun StatusBox(text: String, kind: Int) {
    val bg = when (kind) { 1 -> EmeraldSoft; 2 -> ErrorSoft; else -> Paper }
    val fg = when (kind) { 1 -> Emerald; 2 -> ErrorRed; else -> Muted }
    Box(Modifier.fillMaxWidth().clip(RoundedCornerShape(12.dp)).background(bg).padding(12.dp)) {
        Text(text, color = fg, fontSize = 14.sp, lineHeight = 21.sp)
    }
}

@Composable
private fun InfoChip(text: String) {
    Box(
        Modifier
            .clip(RoundedCornerShape(50))
            .background(EmeraldSoft)
            .padding(horizontal = 12.dp, vertical = 6.dp),
    ) { Text(text, color = Emerald, fontSize = 12.sp, fontWeight = FontWeight.Medium) }
}

@Composable
private fun HeaderCard(teacher: String, school: String, level: String, onChange: (String, String, String) -> Unit) {
    SoftCard {
        Text("بيانات الترويسة", fontFamily = FontFamily.Serif, fontWeight = FontWeight.Bold, fontSize = 18.sp)
        Text(
            "تُكتب هذه القيم في ترويسة كل جذاذة ناتجة. اترك الخانة فارغة لاعتماد ما في القالب أو الجذاذة.",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        OutlinedTextField(
            value = teacher,
            onValueChange = { onChange(it, school, level) },
            singleLine = true,
            label = { Text("الأستاذ(ة)") },
            shape = RoundedCornerShape(14.dp),
            modifier = Modifier.fillMaxWidth(),
        )
        OutlinedTextField(
            value = school,
            onValueChange = { onChange(teacher, it, level) },
            singleLine = true,
            label = { Text("المؤسسة") },
            shape = RoundedCornerShape(14.dp),
            modifier = Modifier.fillMaxWidth(),
        )
        OutlinedTextField(
            value = level,
            onValueChange = { onChange(teacher, school, it) },
            singleLine = true,
            label = { Text("الفئة المستهدفة (مثال: السنة الثالثة إعدادي)") },
            shape = RoundedCornerShape(14.dp),
            modifier = Modifier.fillMaxWidth(),
        )
    }
}

@Composable
private fun UnitPicker(onPick: (String) -> Unit) {
    val names = listOf("الأولى", "الثانية", "الثالثة", "الرابعة", "الخامسة", "السادسة")
    Column(verticalArrangement = Arrangement.spacedBy(8.dp)) {
        names.chunked(3).forEach { rowNames ->
            Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                rowNames.forEach { n ->
                    Box(
                        Modifier
                            .weight(1f)
                            .clip(RoundedCornerShape(50))
                            .border(1.dp, Emerald, RoundedCornerShape(50))
                            .clickable { onPick(n) }
                            .padding(vertical = 10.dp),
                        contentAlignment = Alignment.Center,
                    ) { Text(n, color = Emerald, fontSize = 13.sp, fontWeight = FontWeight.Medium) }
                }
            }
        }
    }
}

@Composable
private fun StorageCard(treeUri: String?, message: String, onPick: () -> Unit) {
    SoftCard {
        Text("مكان التخزين", fontFamily = FontFamily.Serif, fontWeight = FontWeight.Bold, fontSize = 18.sp)
        Text(
            "يُنشئ التطبيق في المكان الذي تختاره مجلد «$ROOT_NAME»، وفيه مجلدات المكوّنات الثلاثة، وفي كل منها ستة مجلدات للوحدات. اختر مثلاً «Documents» ثم «استخدام هذا المجلد».",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        val where = treeUri?.let { Uri.parse(it).lastPathSegment ?: it }
        InfoChip(if (where == null) "لم يُحدَّد مكان بعد" else "المكان: $where")
        if (message.isNotEmpty()) StatusBox(message, if (message.contains("✔")) 1 else 0)
        OutlinedButton(
            onClick = onPick,
            shape = RoundedCornerShape(14.dp),
            border = BorderStroke(1.5.dp, Emerald),
            colors = ButtonDefaults.outlinedButtonColors(contentColor = Emerald),
            modifier = Modifier.fillMaxWidth().height(52.dp),
        ) {
            Text(
                if (treeUri == null) "اختيار المكان وإنشاء المجلدات" else "تغيير المكان / إعادة إنشاء المجلدات",
                fontWeight = FontWeight.Bold,
            )
        }
    }
}

@Composable
private fun SettingsWindow(
    dark: Boolean,
    onDark: (Boolean) -> Unit,
    teacher: String,
    school: String,
    level: String,
    onHeader: (String, String, String) -> Unit,
    treeUri: String?,
    storageMsg: String,
    onPickStorage: () -> Unit,
    apiKey: String,
    aiOn: Boolean,
    onChangeKey: (String) -> Unit,
    onToggleAi: (Boolean) -> Unit,
    onClose: () -> Unit,
) {
    var showKeyHelp by remember { mutableStateOf(false) }
    SubWindow("الإعدادات", onClose, fullHeight = true) {
        SoftCard {
            Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.fillMaxWidth()) {
                Text("الوضع الليلي", fontFamily = FontFamily.Serif, fontWeight = FontWeight.Bold,
                    fontSize = 18.sp, modifier = Modifier.weight(1f))
                Switch(checked = dark, onCheckedChange = onDark)
            }
        }
        HeaderCard(teacher, school, level, onHeader)
        StorageCard(treeUri, storageMsg, onPickStorage)
        SettingsCard(apiKey, aiOn, onChangeKey, onToggleAi, onShowKeyHelp = { showKeyHelp = true })
    }
    if (showKeyHelp) KeyHelpWindow(onClose = { showKeyHelp = false })
}

@Composable
private fun SettingsCard(
    apiKey: String,
    aiOn: Boolean,
    onChange: (String) -> Unit,
    onToggle: (Boolean) -> Unit,
    onShowKeyHelp: () -> Unit,
) {
    SoftCard {
        Row(verticalAlignment = Alignment.CenterVertically, modifier = Modifier.fillMaxWidth()) {
            Text("الذكاء الاصطناعي", fontFamily = FontFamily.Serif, fontWeight = FontWeight.Bold, fontSize = 18.sp,
                modifier = Modifier.weight(1f))
            Switch(checked = aiOn, onCheckedChange = onToggle)
        }
        Text(
            if (aiOn) "مفعّل: يملأ الخانات الفارغة إن توفر المفتاح." else "متوقف: يعمل التفريغ دون ذكاء اصطناعي. المفتاح محفوظ ولم يُحذف.",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        Text("مفتاح Gemini", fontFamily = FontFamily.Serif, fontWeight = FontWeight.Bold, fontSize = 16.sp)
        Text(
            "بدون مفتاح يعمل التفريغ دون إنترنت. مع المفتاح يملأ الذكاء الاصطناعي الخانات الفارغة فقط، ولا يلمس نصاً من المصدر.",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        OutlinedTextField(
            value = apiKey,
            onValueChange = onChange,
            singleLine = true,
            label = { Text("المفتاح") },
            visualTransformation = PasswordVisualTransformation(),
            shape = RoundedCornerShape(14.dp),
            modifier = Modifier.fillMaxWidth(),
        )
        OutlinedButton(
            onClick = onShowKeyHelp,
            shape = RoundedCornerShape(14.dp),
            modifier = Modifier.fillMaxWidth(),
        ) {
            Text("كيف أحصل على المفتاح؟", fontWeight = FontWeight.Bold, color = Emerald)
        }
    }
}

/** نافذة شرح خطوات الحصول على مفتاح Gemini، مع زر يفتح صفحة Google AI Studio. */
@Composable
private fun KeyHelpWindow(onClose: () -> Unit) {
    val ctx = LocalContext.current
    SubWindow("كيف أحصل على مفتاح Gemini؟", onClose, fullHeight = false) {
        Text(
            "الذكاء الاصطناعي اختياري: بدون مفتاح يعمل التفريغ كاملًا ودون إنترنت. وإن رغبت في تعبئة الخانات الفارغة آليًا فاتبع الخطوات:",
            fontSize = 13.sp,
            color = Muted,
            lineHeight = 20.sp,
        )
        StepTitle("١", "افتح صفحة Google AI Studio")
        Text(
            "اضغط الزر بالأسفل، وسجّل الدخول بحساب Google العادي (نفس حساب Gmail).",
            fontSize = 13.sp, color = Muted, lineHeight = 20.sp,
        )
        StepTitle("٢", "أنشئ المفتاح")
        Text(
            "اضغط «Create API key»، واختر مشروعًا إن طُلب منك أو أنشئ مشروعًا جديدًا.",
            fontSize = 13.sp, color = Muted, lineHeight = 20.sp,
        )
        StepTitle("٣", "انسخ المفتاح")
        Text(
            "سيظهر مفتاح يبدأ بـ «AIza…» — انسخه.",
            fontSize = 13.sp, color = Muted, lineHeight = 20.sp,
        )
        StepTitle("٤", "ألصقه في التطبيق")
        Text(
            "ارجع إلى الإعدادات، ألصقه في خانة «المفتاح»، فيُحفظ داخل الجهاز ولا يُرسل إلى أي جهة غير Google.",
            fontSize = 13.sp, color = Muted, lineHeight = 20.sp,
        )
        Text(
            "المفتاح شخصي: لا تشاركه مع أحد. الحدّ المجاني في Google يكفي الاستعمال العادي.",
            fontSize = 12.sp, color = Muted, lineHeight = 18.sp,
        )
        SolidButton(
            "افتح صفحة المفتاح في المتصفح",
            Emerald,
            { openUrl(ctx, GEMINI_KEY_URL) },
            Modifier.fillMaxWidth(),
        )
    }
}

/** ينشئ شجرة المجلدات (الجذر ← 3 مكوّنات ← 6 وحدات لكل مكوّن) إن لم تكن موجودة. */
private fun ensureTree(ctx: Context, treeUri: Uri): DocumentFile {
    val base = DocumentFile.fromTreeUri(ctx, treeUri)
        ?: throw IllegalStateException("تعذّر الوصول إلى المكان المختار")
    val root = childDir(base, ROOT_NAME)
    for (c in COMP_FOLDERS) {
        val cd = childDir(root, c)
        for (u in UNIT_FOLDERS) childDir(cd, u)
    }
    return root
}

private fun childDir(parent: DocumentFile, name: String): DocumentFile {
    val found = parent.findFile(name)
    if (found != null && found.isDirectory) return found
    return parent.createDirectory(name)
        ?: throw IllegalStateException("تعذّر إنشاء المجلد «$name»")
}

/** يخزّن الملف في «الجذر/المكوّن/الوحدة»؛ يستبدل الملف إن وُجد بالاسم نفسه. يعيد المسار الظاهر. */
private fun storeDocx(ctx: Context, treeUri: Uri, srcPath: String, comp: String, unit: String): String {
    if (comp.isBlank()) throw IllegalStateException("تعذّر تحديد المكوّن من الترويسة")
    val root = ensureTree(ctx, treeUri)
    val cd = childDir(root, comp)
    val target = if (unit.isBlank()) cd else childDir(cd, unit)
    val src = File(srcPath)
    target.findFile(src.name)?.delete()
    val f = target.createFile(DOCX_MIME, src.name)
        ?: throw IllegalStateException("تعذّر إنشاء الملف")
    val out = ctx.contentResolver.openOutputStream(f.uri)
        ?: throw IllegalStateException("تعذّر الكتابة في الملف")
    out.use { o -> src.inputStream().use { it.copyTo(o) } }
    val where = if (unit.isBlank()) "$ROOT_NAME/$comp" else "$ROOT_NAME/$comp/$unit"
    return "$where/${f.name ?: src.name}"
}

/** ينسخ الملف المختار إلى مخزن التطبيق ويعيد (اسمه الظاهر، مساره). */
private fun importSource(ctx: Context, uri: Uri, id: String): Pair<String, String> {
    var name = "source.docx"
    ctx.contentResolver.query(uri, null, null, null, null)?.use { c ->
        val i = c.getColumnIndex(OpenableColumns.DISPLAY_NAME)
        if (i >= 0 && c.moveToFirst()) {
            name = c.getString(i) ?: name
        }
    }
    val dir = File(ctx.filesDir, "inputs").apply { mkdirs() }
    val dest = File(dir, "$id.docx")
    val input = ctx.contentResolver.openInputStream(uri)
        ?: throw IllegalStateException("لا يمكن قراءة الملف")
    input.use { src -> dest.outputStream().use { out -> src.copyTo(out) } }
    return Pair(name, dest.absolutePath)
}

/** ينسخ القالب المضمَّن من assets إلى مخزن التطبيق (Python يحتاج مساراً حقيقياً). */
private fun ensureModel(ctx: Context, comp: String): File {
    val dir = File(ctx.filesDir, "models").apply { mkdirs() }
    val dest = File(dir, "$comp.docx")
    ctx.assets.open("models/$comp.docx").use { src ->
        dest.outputStream().use { out -> src.copyTo(out) }
    }
    return dest
}

private fun headerJson(teacher: String, school: String, level: String, unit: String = ""): String =
    JSONObject()
        .put("teacher", teacher.trim())
        .put("school", school.trim())
        .put("level", level.trim())
        .apply { if (unit.isNotBlank()) put("unit", unit) }
        .toString()

private fun runBridge(
    ctx: Context,
    comp: String,
    sourcePath: String,
    displayName: String,
    apiKey: String,
    headerJson: String,
    style: String,
): JSONObject {
    val model = ensureModel(ctx, comp)
    val outDir = File(ctx.filesDir, "out/$comp").apply { mkdirs() }
    val outName = "جذاذة_" + displayName.substringBeforeLast('.') + ".docx"
    val py = Python.getInstance()
    val json = py.getModule("bridge")
        .callAttr("run", comp, model.absolutePath, sourcePath, outDir.absolutePath, apiKey, outName, headerJson, style)
        .toString()
    return JSONObject(json)
}

private fun openUrl(ctx: Context, url: String) {
    try {
        ctx.startActivity(Intent(Intent.ACTION_VIEW, Uri.parse(url)))
    } catch (e: Exception) {
        Toast.makeText(ctx, "تعذّر فتح المتصفح: ${e.message}", Toast.LENGTH_LONG).show()
    }
}

private fun openOrShare(ctx: Context, path: String, view: Boolean) {
    try {
        val uri = FileProvider.getUriForFile(ctx, "${ctx.packageName}.fileprovider", File(path))
        val intent = if (view) {
            Intent(Intent.ACTION_VIEW).setDataAndType(uri, DOCX_MIME)
        } else {
            Intent(Intent.ACTION_SEND).setType(DOCX_MIME).putExtra(Intent.EXTRA_STREAM, uri)
        }
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        ctx.startActivity(if (view) intent else Intent.createChooser(intent, "مشاركة الجذاذة"))
    } catch (e: ActivityNotFoundException) {
        Toast.makeText(ctx, "لا يوجد تطبيق يفتح ملفات Word. جرّب «مشاركة / حفظ».", Toast.LENGTH_LONG).show()
    } catch (e: Exception) {
        Toast.makeText(ctx, "تعذّر: ${e.message}", Toast.LENGTH_LONG).show()
    }
}
