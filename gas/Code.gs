/**
 * Монтажный план — мост между Google Диском и приложением.
 *
 * Работает под служебным Google-аккаунтом, которому рук-про дал доступ к таблице (хватит «Читатель»).
 * Приложение спрашивает:  …/exec?id=<ID файла>&since=<время последней версии>
 *   — если таблица не менялась, ответ короткий: {unchanged:true}
 *   — если менялась — вся таблица в формате .xlsx (base64) и время изменения.
 * Подходит и «настоящая» Google Таблица, и файл .xlsx/.xlsm, который открывают в Google Таблицах.
 */

// ID файлов, которые мост может отдавать (часть ссылки между /d/ и /edit).
// Пусто = любые файлы, к которым у служебного аккаунта есть доступ.
const ALLOWED = ["11SofZ2n57scQpGyay51Vsle1KVxrkB3M"];   // ЛУКОЙЛ 35

// Пароль к проекту. В коде его НЕТ (код лежит в открытом репозитории) — он хранится в настройках скрипта:
// редактор Apps Script → ⚙ «Настройки проекта» → «Свойства скрипта» → «Добавить свойство»:
//   PASS_<ID таблицы>  — пароль для одной таблицы, например PASS_11SofZ2n57scQpGyay51Vsle1KVxrkB3M
//   PASS               — общий пароль для всех таблиц, у которых нет своего
// Нет ни того, ни другого — таблица открывается без пароля. Новую версию развёртывать не нужно: свойство действует сразу.
// Приложение спрашивает пароль один раз и запоминает на устройстве; сменили пароль — у всех снова спросит.
function passFor(id) {
  const pr = PropertiesService.getScriptProperties();
  return pr.getProperty("PASS_" + id) || pr.getProperty("PASS") || "";
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const id = String(p.id || "").trim();
    if (!id) return out({ ok: true, msg: "Монтажный план: мост работает" });
    if (ALLOWED.length && ALLOWED.indexOf(id) < 0) return out({ error: "Эта таблица не разрешена в скрипте (ALLOWED)" });

    const cache = CacheService.getScriptCache();
    const pass = passFor(id);
    if (pass) {
      const key = String(p.key || "");
      if (key !== pass) return out({ error: key ? "Неверный пароль" : "Нужен пароль", auth: key ? "wrong" : "need" });
    }

    // Время изменения кэшируем на 10 секунд, чтобы много открытых телефонов не дёргали Диск
    let meta = cache.get("m:" + id);
    if (meta) meta = JSON.parse(meta);
    else {
      const f = DriveApp.getFileById(id);
      meta = { modified: f.getLastUpdated().toISOString(), name: f.getName(), mime: f.getMimeType() };
      cache.put("m:" + id, JSON.stringify(meta), 10);
    }
    if (p.since && p.since === meta.modified) return out({ unchanged: true, modified: meta.modified, name: meta.name });

    const blob = getXlsx(id, meta.mime);
    return out({ modified: meta.modified, name: meta.name, mime: meta.mime, data: Utilities.base64Encode(blob.getBytes()) });
  } catch (err) {
    const m = String((err && err.message) || err);
    if (/not found|не найден|access|доступ/i.test(m)) return out({ error: "У служебного аккаунта нет доступа к этому файлу" });
    return out({ error: m });
  }
}

function getXlsx(id, mime) {
  if (mime === MimeType.GOOGLE_SHEETS) {
    // настоящая Google Таблица — выгружаем её как .xlsx
    const r = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx", {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) throw new Error("Google не отдал таблицу (код " + r.getResponseCode() + ")");
    return r.getBlob();
  }
  // файл Excel на Диске — отдаём как есть
  return DriveApp.getFileById(id).getBlob();
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Запусти один раз вручную (кнопка «Выполнить»), чтобы Google спросил разрешения.
 *  Подставь ID таблицы — в журнале появится имя файла и время изменения. */
function test() {
  const ID = "11SofZ2n57scQpGyay51Vsle1KVxrkB3M";   // ЛУКОЙЛ 35
  const res = JSON.parse(doGet({ parameter: { id: ID } }).getContent());
  Logger.log(res.error || (res.name + " · изменён " + res.modified + " · " + Math.round((res.data || "").length * 0.75 / 1024) + " КБ"));
  UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true }); // чтобы сразу выдать разрешение на внешние запросы
}
