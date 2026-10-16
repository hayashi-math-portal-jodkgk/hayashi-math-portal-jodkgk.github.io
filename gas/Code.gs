// はやし数学ポータル（Google Apps Script 版）
//
// 学校の端末では github.io が教育委員会のフィルタで「不確定サイト」として止められるので，
// 計算王と同じ Apps Script から同じページを出す。ページの元は今までどおり GitHub（SITE）。
// ・10 分ごと（トリガー）に GitHub の index.html を読み，変わっていたら取りこむ
// ・プリントの画像は Google ドライブのフォルダ（FOLDER_NAME）に写しを置き，そこから見せる（変わった画像だけ置きかえる）
// ・画像だけを直したとき（index.html が同じとき）も，1 時間ごとに全部の画像を確かめる
// 使い方は gas/README.md。

var SITE = 'https://hayashi-math-portal-jodkgk.github.io/';
var TITLE = 'はやし数学ポータル';
var FOLDER_NAME = 'はやし数学ポータル（画像の写し）';
// ドライブの画像を表示する URL。{id} にファイルの ID が入る
var IMG_URL = 'https://lh3.googleusercontent.com/d/{id}=w1600';
var FULL_CHECK_MS = 60 * 60 * 1000;   // 画像を全部確かめる間隔
var TIME_LIMIT_MS = 4.5 * 60 * 1000;  // 1 回の実行はここまで（Apps Script は 6 分で止まる）。残りは次の回
var BATCH = 20;                        // 画像を一度に読む数
var CHUNK = 25000;                     // キャッシュ 1 つに入れる文字数（1 つ 100KB まで）

// ---- ページを出す ----
function doGet() {
  var page = loadPage_();
  var out = page ? HtmlService.createHtmlOutput(page)
    : HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">準備中です。数分たってから開きなおしてください。</p>');
  return out.setTitle(TITLE).addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function loadPage_() {
  var page = cacheGet_('page');
  if (page) return page;
  var id = PropertiesService.getScriptProperties().getProperty('pageId');
  if (!id) return null;
  try { page = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); } catch (e) { return null; }
  cachePut_('page', page);
  return page;
}

// ---- はじめに 1 回だけ実行する：10 分ごとのトリガーを作って，1 回目の取りこみをする ----
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(10).create();
  sync();
}

// ---- GitHub から取りこむ（トリガーで 10 分ごと） ----
// 手で実行する用：画像も全部すぐに確かめる
function syncNow() {
  PropertiesService.getScriptProperties().deleteProperty('fullAt');
  sync();
}

function sync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { sync_(); } finally { lock.releaseLock(); }
}

function sync_() {
  var start = Date.now();
  var props = PropertiesService.getScriptProperties();
  var res = UrlFetchApp.fetch(SITE + '?t=' + start, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('GitHub のページを読めませんでした（' + res.getResponseCode() + '）');
  var html = res.getContentText('UTF-8');
  var hash = md5_(html);
  var last = +(props.getProperty('fullAt') || 0);
  if (props.getProperty('hash') === hash && props.getProperty('done') === '1' && start - last < FULL_CHECK_MS) return;
  if (props.getProperty('hash') !== hash) { props.setProperties({ hash: hash, done: '0' }); props.deleteProperty('next'); }

  var m = html.match(/<script id="site-data" type="application\/json">([\s\S]*?)<\/script>/);
  if (!m) throw new Error('ページの中に site-data が見つかりません');
  var paths = imagePaths_(JSON.parse(m[1]));
  var folder = folder_();
  var map = readJson_('mapId') || {};
  var changed = 0;

  // 画像：前と同じ（304 か，中身が同じ）ならそのまま。変わったものだけ新しく置き，古いものはごみ箱へ
  var from = +(props.getProperty('next') || 0);
  if (from >= paths.length) from = 0;
  for (var i = from; i < paths.length; i += BATCH) {
    if (Date.now() - start > TIME_LIMIT_MS) {
      writeJson_('mapId', map);
      props.setProperty('next', String(i));
      Logger.log('時間切れ：' + i + '/' + paths.length + ' まで。続きは次の回');
      return;
    }
    var batch = paths.slice(i, i + BATCH);
    var reqs = batch.map(function (p) {
      var o = { url: SITE + p + '?t=' + start, muteHttpExceptions: true };
      if (map[p] && map[p].etag) o.headers = { 'If-None-Match': map[p].etag };
      return o;
    });
    UrlFetchApp.fetchAll(reqs).forEach(function (r, j) {
      var p = batch[j], code = r.getResponseCode(), old = map[p];
      if (code === 304 && old) return;
      if (code !== 200) { Logger.log('読めない画像：' + p + '（' + code + '）'); return; }
      var etag = r.getHeaders()['ETag'] || r.getHeaders()['etag'] || '';
      var blob = r.getBlob(), sum = md5_(blob.getBytes());
      if (old && old.md5 === sum) { old.etag = etag; return; }
      var f = folder.createFile(blob.setName(p.replace(/^p\//, '').replace(/\//g, '_')).setContentType('image/webp'));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      if (old) trash_(old.id);
      map[p] = { id: f.getId(), etag: etag, md5: sum };
      changed++;
    });
  }
  props.deleteProperty('next');

  // ページから消えた画像はごみ箱へ
  var keep = {};
  paths.forEach(function (p) { keep[p] = 1; });
  Object.keys(map).forEach(function (p) { if (!keep[p]) { trash_(map[p].id); delete map[p]; changed++; } });
  writeJson_('mapId', map);

  // ページ：画像の場所（window.GAS）を入れて保存する
  var cfg = { site: SITE, url: IMG_URL, img: {} };
  paths.forEach(function (p) { if (map[p]) cfg.img[p] = map[p].id; });
  var tag = '<script>window.GAS = ' + JSON.stringify(cfg).replace(/</g, '\\u003c') + ';</script>\n';
  var page = html.replace('<script id="site-data"', tag + '<script id="site-data"');
  var pageId = props.getProperty('pageId'), saved = false;
  if (pageId) { try { DriveApp.getFileById(pageId).setContent(page); saved = true; } catch (e) {} }
  if (!saved) props.setProperty('pageId', folder.createFile('page.html', page, MimeType.HTML).getId());
  cachePut_('page', page);
  props.setProperties({ done: '1', fullAt: String(start) });
  Logger.log('取りこみ完了：画像 ' + paths.length + ' 枚（置きかえ・削除 ' + changed + '）');
}

// site-data から画像のパス（p/<単元>/<no><q|a>.webp など）を全部あげる
function imagePaths_(data) {
  var out = [];
  data.grades.forEach(function (g) {
    g.units.forEach(function (u) {
      (u.items || []).forEach(function (it) {
        ['q', 'a'].forEach(function (k) {
          var n = (it.pages && it.pages[k]) || 0;
          if (!n) return;
          if (u.sprite) out.push(u.img + '/' + it.no + k + '.webp');
          else for (var s = 1; s <= n; s++) out.push(u.img + '/' + it.no + k + s + '.webp');
        });
      });
    });
  });
  return out.filter(function (p, i) { return out.indexOf(p) === i; });
}

// ---- 道具 ----
function folder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('folderId');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var folder = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('folderId', folder.getId());
  return folder;
}

function readJson_(key) {
  var id = PropertiesService.getScriptProperties().getProperty(key);
  if (!id) return null;
  try { return JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')); } catch (e) { return null; }
}

function writeJson_(key, obj) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(key), text = JSON.stringify(obj);
  if (id) { try { DriveApp.getFileById(id).setContent(text); return; } catch (e) {} }
  props.setProperty(key, folder_().createFile(key.replace(/Id$/, '') + '.json', text, MimeType.PLAIN_TEXT).getId());
}

function trash_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }

function md5_(v) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, v).map(function (b) {
    return ('0' + (b & 255).toString(16)).slice(-2);
  }).join('');
}

// CacheService は 1 つ 100KB までなので，分けて入れる（6 時間）
function cachePut_(key, str) {
  var o = {}, n = Math.ceil(str.length / CHUNK);
  for (var i = 0; i < n; i++) o[key + '_' + i] = str.substr(i * CHUNK, CHUNK);
  o[key] = String(n);
  try { CacheService.getScriptCache().putAll(o, 21600); } catch (e) {}
}

function cacheGet_(key) {
  var c = CacheService.getScriptCache(), n = +(c.get(key) || 0);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var got = c.getAll(keys), parts = [];
  for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; parts.push(got[keys[j]]); }
  return parts.join('');
}
