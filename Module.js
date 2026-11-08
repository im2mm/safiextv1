/**
 * PILOT Extension: battery_status
 *
 * Warehouse Sensor Status inside PILOT. Same logic as the Al Safi web report
 * (alsafi-temp-report, lib/im2m.ts + app/battery/page.tsx):
 *
 *   - Battery data comes from the raw device payload (triggerdata) of the latest packets:
 *       BatteryVoltage<dec>=3.630, BatteryVoltageStatus<dec>=Normal, Temperature<dec>=5.1,
 *       Humidity<dec>=60, RSSI<dec>=121
 *     <dec> is the wireless probe id in decimal; the id printed on the probe is the same
 *     number in hex (2183475330 -> 82253082).
 *   - A probe is listed only if the object's sensor list has a Temperature<dec> or
 *     Humidity<dec> sensor for it (average and status sensors are ignored). The Temperature
 *     sensor's name wins; a "BTS-<hex>-" prefix is removed.
 *   - Status is decided by VOLTAGE: >= 3.3 V Normal, < 3.3 V Low. A probe whose latest
 *     voltage is older than 3 h (or missing) is "No reading", not Low. The platform's
 *     BatteryVoltageStatus text is shown next to it when it differs.
 *   - Duration: Now (= last 24 h), 6, 12, 48 or 72 hours. Trace pages of 100 rows, newest
 *     first, stopping as soon as every probe has a voltage (max 2000 rows).
 *   - Click a sensor row for its readings in the last 24 hours; "All sensors" shows every
 *     configured sensor of the object with its latest value; Excel export with the
 *     Battery voltage / Gateways / All sensors sheets.
 *
 * Layout (nested tabs):
 *   Left navigation tab "Battery Status" (Pilot.utils.LeftBarPanel) with two inner tabs:
 *     - Warehouses : region -> object tree with normal/total badges; click an object to open it
 *     - Settings   : duration and auto refresh
 *   Main panel (Ext.tab.Panel) with three tabs:
 *     - Overview   : summary tiles + every sensor grouped by object, filter, Excel
 *     - Alerts     : sensors that are Low or No reading
 *     - Warehouses : nested vertical tab per object with gateway info and All sensors
 *
 * Data sources (read-only, legacy PILOT backend endpoints of the running web app,
 * not part of the documented /api/v3 surface):
 *   - ax/tree.php?vehs=1&checked=1&node=src            object tree
 *   - ax/sensors/sensors.php  (POST agent_id)          sensor list (cached 30 min)
 *   - ax/sensors_trace.php?cmd=get&agent_id=..&ts=..&te=..&page=..&start=..&limit=100&sort=..
 *
 * All classes live in this single file so the extension does not depend on
 * Ext.Loader resolving extra Store.battery_status.* files.
 */

(function () {
    var EXT = 'battery_status';
    var VERSION = '1.3.3'; // bump on every release; shown in the console, Settings tab and doc/

    var TZ = 'Asia/Riyadh';          // account local time used by the platform
    var TZ_OFFSET_S = 3 * 3600;      // Riyadh has no DST; used for Excel date cells
    var LOW_VOLTAGE = 3.3;           // below this = low battery
    var STALE_HOURS = 3;             // no reading for longer than this = "No reading" (not low)
    var HISTORY_HOURS = 24;          // fixed window for a sensor's previous readings
    var CONCURRENCY = 3;
    var WINDOWS = [['now', 'Now'], [6, 'Last 6 hours'], [12, 'Last 12 hours'], [48, 'Last 48 hours'], [72, 'Last 72 hours']];

    function t(text) {
        return text; // PILOT's l() translated some labels into Russian; the reports are English like the site
    }

    function enc(value) {
        return Ext.String.htmlEncode(value === null || value === undefined ? '' : String(value));
    }

    /** Decode HTML entities and repair UTF-8 text that was double-encoded as Latin-1 (e.g. "Â°C"). */
    function cleanText(v) {
        var s = (v === null || v === undefined || v === 'null') ? '' : String(v);
        if (/[ÂÃ][\u0080-ÿ]/.test(s)) {
            try { s = decodeURIComponent(escape(s)); } catch (e) { /* keep as is */ }
        }
        return Ext.String.htmlDecode(s).replace(/&#039;|&apos;/g, "'").replace(/&nbsp;/g, ' ').trim();
    }

    function storageGet(key, fallback) {
        try {
            var raw = window.localStorage.getItem(EXT + '.' + key);
            return raw === null ? fallback : JSON.parse(raw);
        } catch (e) {
            return fallback;
        }
    }

    function storageSet(key, value) {
        try {
            window.localStorage.setItem(EXT + '.' + key, JSON.stringify(value));
        } catch (e) {
            // Storage can be blocked (private mode); settings then last for this session only.
        }
    }

    var STATUS = {
        low:    { rank: 2, label: 'Low',        icon: 'fa fa-battery-quarter', cls: 'bs-alert',  tile: 'bs-tile-alert' },
        nodata: { rank: 1, label: 'No reading', icon: 'fa fa-signal',          cls: 'bs-nodata', tile: 'bs-tile-nodata' },
        normal: { rank: 0, label: 'Normal',     icon: 'fa fa-battery-full',    cls: 'bs-ok',     tile: 'bs-tile-ok' }
    };

    function statusOf(v) {
        return v === null ? 'nodata' : (v >= LOW_VOLTAGE ? 'normal' : 'low');
    }

    function statusPill(status) {
        var meta = STATUS[status] || STATUS.nodata;
        return '<span class="bs-pill ' + meta.cls + '"><i class="' + meta.icon + '"></i> ' + enc(t(meta.label)) + '</span>';
    }

    function num(v) {
        if (v === null || v === undefined || v === '') {
            return null;
        }
        var n = parseFloat(String(v).replace(',', '.'));
        return isFinite(n) ? n : null;
    }

    function fixed(v, d) {
        return v === null || v === undefined ? '—' : Number(v).toFixed(d === undefined ? 1 : d);
    }

    function toHex(dec) {
        try {
            if (typeof BigInt === 'function') {
                return BigInt(dec).toString(16).toUpperCase();
            }
        } catch (e) { /* fall through */ }
        var n = Number(dec);
        return isFinite(n) && n > 0 ? n.toString(16).toUpperCase() : String(dec);
    }

    function riyadhParts(ms) {
        var p = {};
        new Intl.DateTimeFormat('en-GB', {
            timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit',
            hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
        }).formatToParts(new Date(ms)).forEach(function (x) { p[x.type] = x.value; });
        return p;
    }

    /** Sensors Tracing date format: dd.mm.yyyy HH:MM in the account time zone. */
    function traceDate(ms) {
        var p = riyadhParts(ms);
        return p.day + '.' + p.month + '.' + p.year + ' ' + p.hour + ':' + p.minute;
    }

    function fmtTs(sec) {
        if (!sec) {
            return '—';
        }
        return new Intl.DateTimeFormat('en-GB', {
            timeZone: TZ, day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
        }).format(new Date(sec * 1000));
    }

    function ago(sec) {
        if (!sec) {
            return '';
        }
        var m = Math.round((Date.now() / 1000 - sec) / 60), h;
        if (m < 1) { return t('just now'); }
        if (m < 60) { return m + ' ' + t('min ago'); }
        h = m / 60;
        return h < 48 ? h.toFixed(h < 10 ? 1 : 0) + ' ' + t('h ago') : Math.round(h / 24) + ' ' + t('d ago');
    }

    function tsHtml(sec) {
        return enc(fmtTs(sec)) + (sec ? ' <span class="bs-muted">' + enc(ago(sec)) + '</span>' : '');
    }

    function byName(a, b) {
        return String(a).localeCompare(String(b), undefined, { numeric: true });
    }

    function siteOf(group) {
        var parts = String(group || '').split(' / ');
        return parts[parts.length - 1] || '';
    }

    // ---------------------------------------------------------------------
    // Shared helpers: external libraries, file download, Al Safi API client.
    // ---------------------------------------------------------------------
    var LIB = {
        excel: 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
        leafletJs: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js',
        leafletCss: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.css',
        geomanJs: 'https://cdn.jsdelivr.net/npm/@geoman-io/leaflet-geoman-free@2.20.4/dist/leaflet-geoman.js',
        geomanCss: 'https://cdn.jsdelivr.net/npm/@geoman-io/leaflet-geoman-free@2.20.4/dist/leaflet-geoman.css',
        togeojson: 'https://cdn.jsdelivr.net/npm/@tmcw/togeojson@7.1.3/dist/togeojson.umd.js'
    };
    var API_DEFAULT = 'https://alsafi-temp-report.vercel.app'; // Al Safi report backend (Vercel)
    var scriptPromises = {};

    /** Loads a UMD script as a global (an AMD loader on the page is hidden while it runs). */
    function loadScript(url) {
        if (!scriptPromises[url]) {
            scriptPromises[url] = new Promise(function (resolve, reject) {
                var amd = window.define, el = document.createElement('script');
                window.define = undefined;
                el.src = url;
                el.async = true;
                el.onload = function () { window.define = amd; resolve(); };
                el.onerror = function () { window.define = amd; delete scriptPromises[url]; reject(new Error('Could not load ' + url)); };
                document.head.appendChild(el);
            });
        }
        return scriptPromises[url];
    }

    function loadCss(url) {
        if (document.querySelector('link[href="' + url + '"]')) {
            return;
        }
        var el = document.createElement('link');
        el.rel = 'stylesheet';
        el.href = url;
        document.head.appendChild(el);
    }

    function ensureExcel() {
        if (window.ExcelJS) {
            return Promise.resolve(window.ExcelJS);
        }
        return loadScript(LIB.excel).then(function () {
            if (!window.ExcelJS) { throw new Error('Excel library did not load.'); }
            return window.ExcelJS;
        });
    }

    function saveFile(data, name, type) {
        var a = document.createElement('a'),
            url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type: type }));
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
    }

    function stampNow(tz) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: tz || TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
            .format(new Date()).replace(', ', ' ').replace(':', '-');
    }

    function errorText(e) {
        return (e && e.message) || String(e);
    }

    /**
     * Client for the Al Safi report API (the former web site, now the extension's backend).
     * Uses the same password as the site; the token is kept in this browser.
     */
    var Api = {
        base: API_DEFAULT,
        token: null,

        init: function () {
            var me = this;
            me.token = storageGet('apiToken', null);
            try {
                if (window.Pilot && Pilot.Extensions && Pilot.Extensions.getSettings) {
                    Pilot.Extensions.getSettings(EXT).then(function (s) {
                        if (s && s.api_url) { me.base = String(s.api_url).replace(/\/+$/, ''); }
                    }, function () { /* keep default */ });
                }
            } catch (e) { /* keep default */ }
        },

        raw: function (path, opts) {
            var headers = { 'content-type': 'application/json' };
            if (this.token) { headers.authorization = 'Bearer ' + this.token; }
            return fetch(this.base + path, {
                method: opts.method || (opts.json !== undefined ? 'POST' : 'GET'),
                headers: headers,
                body: opts.json !== undefined ? JSON.stringify(opts.json) : undefined
            });
        },

        parse: function (res) {
            return res.json().catch(function () { return {}; }).then(function (d) {
                if (!res.ok) { throw new Error(d.error || ('Request failed (' + res.status + ')')); }
                return d;
            });
        },

        /** opts: {method, json, silent}. silent = never show the sign-in window. */
        call: function (path, opts) {
            var me = this;
            opts = opts || {};
            if (!me.token && opts.silent) {
                return Promise.reject(new Error('Not signed in'));
            }
            return me.raw(path, opts).then(function (res) {
                if (res.status !== 401) { return me.parse(res); }
                me.setToken(null);
                if (opts.silent) { throw new Error('Not signed in'); }
                return me.login().then(function () { return me.raw(path, opts); }).then(function (r2) { return me.parse(r2); });
            }, function () {
                throw new Error(t('Cannot reach the report server') + ' (' + me.base + '). ' + t('If this keeps happening, the backend update (CORS for PILOT) is not deployed yet.'));
            });
        },

        setToken: function (tok) {
            this.token = tok;
            storageSet('apiToken', tok);
        },

        login: function () {
            var me = this;
            if (me.loginPromise) { return me.loginPromise; }
            me.loginPromise = new Promise(function (resolve, reject) {
                var settled = false, win, field, msg;
                function submit() {
                    var pw = field.getValue();
                    if (!pw) { return; }
                    win.setLoading(t('Signing in…'));
                    me.raw('/api/login', { json: { password: pw } }).then(function (res) { return me.parse(res); }).then(function (d) {
                        if (!d.token) { throw new Error(t('The report server is not updated for the extension yet (no token).')); }
                        me.setToken(d.token);
                        settled = true;
                        win.close();
                        resolve();
                    }).catch(function (e) {
                        win.setLoading(false);
                        msg.setHtml('<div class="bs-error">' + enc(errorText(e)) + '</div>');
                    });
                }
                field = Ext.create('Ext.form.field.Text', {
                    inputType: 'password', fieldLabel: t('Password'), labelAlign: 'top', anchor: '100%',
                    listeners: { specialkey: function (f, e) { if (e.getKey() === e.ENTER) { submit(); } } }
                });
                msg = Ext.create('Ext.Component', { html: '' });
                win = Ext.create('Ext.window.Window', {
                    title: t('Sign in to the Al Safi reports'),
                    modal: true, width: 360, bodyPadding: 14, layout: 'anchor',
                    items: [
                        { xtype: 'component', cls: 'bs-muted', html: enc(t('Use the same password as the Al Safi report site.')) + '<br>' + enc(me.base) },
                        field, msg
                    ],
                    buttons: [
                        { text: t('Cancel'), handler: function () { win.close(); } },
                        { text: t('Sign in'), handler: submit }
                    ],
                    listeners: {
                        close: function () {
                            me.loginPromise = null;
                            if (!settled) { reject(new Error(t('Sign-in cancelled'))); }
                        }
                    }
                });
                win.show();
                Ext.defer(function () { field.focus(); }, 100);
            });
            return me.loginPromise;
        }
    };

    /** Shared Excel helpers (same look as the site's ExcelJS exports). */
    function xlHead(row) {
        row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        row.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B7A75' } }; });
    }

    function xlFill(row, argb) {
        row.eachCell({ includeEmpty: true }, function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; });
    }

    // ---------------------------------------------------------------------
    // Sensor list rules (same as the web report)
    // ---------------------------------------------------------------------
    var PROBE_KEY = /^(BatteryVoltageStatus|BatteryVoltage|Temperature|Humidity|RSSI)(\d{6,})$/;

    /** Calculated "average" sensors (e.g. AVERAGE TEMP = (t1+t2+t3)/3) are not real probes. */
    function isAverageSensor(s) {
        var text = cleanText(s.info) + ' ' + cleanText(s.sensordefaultdesc),
            formula = String(s.formula === null || s.formula === undefined ? '' : s.formula);
        // also catches misspellings used on the platform: "Avreage", "Avarage", "Averge", "Avrage"
        return /\bav[a-z]{1,5}ge\b|average|\bavg\b/i.test(text) || /\)\s*\+.*\/\s*\d/.test(formula);
    }

    /** Status/selector sensors ("Battery Status 3", BatteryVoltageStatus<id> fields) are ignored. */
    function isStatusSensor(s) {
        return /status/i.test(cleanText(s.info) + ' ' + cleanText(s.sensordefaultdesc)) ||
            /^BatteryVoltageStatus\d+$/.test(String(s.sensorfieldname || ''));
    }

    /** Probes to list: decimal id -> name. Temperature sensor beats Humidity; formula sensors last. */
    function probeNames(sensors) {
        var best = {};
        Ext.Array.each(sensors, function (s) {
            var field = String(s.sensorfieldname || ''),
                m = PROBE_KEY.exec(field),
                base, name, r;
            if (!m) { return; }
            base = { Temperature: 0, Humidity: 2 }[m[1]];
            if (base === undefined) { return; }
            name = cleanText(s.info);
            if (!name || name === field) { return; }
            name = name.replace(/^BTS-[0-9A-F]+-/i, '');
            r = base + (s.formula && String(s.formula).trim() ? 1 : 0);
            if (!best[m[2]] || r < best[m[2]].r) {
                best[m[2]] = { name: name, r: r };
            }
        });
        var out = {};
        Ext.Object.each(best, function (id, v) { out[id] = v.name; });
        return out;
    }

    function parseTrigger(s) {
        var out = {};
        Ext.Array.each(String(s === null || s === undefined ? '' : s).split(/,\s*/), function (part) {
            var i = part.indexOf('=');
            if (i > 0) {
                out[part.slice(0, i).trim()] = part.slice(i + 1).trim();
            }
        });
        return out;
    }

    function tsOf(r) {
        return Number(r.ts_date !== undefined && r.ts_date !== null ? r.ts_date : r.ts) || 0;
    }

    // ---------------------------------------------------------------------
    // Data service.
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.Data', {
        mixins: ['Ext.mixin.Observable'],

        configTtlMs: 30 * 60 * 1000,      // sensor list cache lifetime

        constructor: function (config) {
            this.mixins.observable.constructor.call(this, config);
            this.objects = [];            // [{agentId, name, group, region, path, imei}]
            this.configCache = {};        // agentId -> {ts, sensors}
            this.results = {};            // agentId -> {obj, data?, error?}
            this.store = Ext.create('Ext.data.Store', {
                fields: [
                    'key', 'agentId', 'objectName', 'group', 'region', 'groupKey', 'idDec', 'idHex', 'name',
                    'status', 'platformStatus',
                    { name: 'voltage', type: 'auto' },
                    { name: 'temperature', type: 'auto' },
                    { name: 'humidity', type: 'auto' },
                    { name: 'rssi', type: 'auto' },
                    { name: 'ts', type: 'auto' },
                    { name: 'rank', type: 'int' },
                    { name: 'voltSort', type: 'float' }
                ],
                // Same order as the web report: Low -> No reading -> Normal, lowest voltage first.
                sorters: [
                    { property: 'rank', direction: 'DESC' },
                    { property: 'voltSort', direction: 'ASC' },
                    { property: 'objectName', direction: 'ASC' },
                    { property: 'name', direction: 'ASC' }
                ]
            });
        },

        axUrl: function (path) {
            // PILOT runs under <host>/backend/; base_url points there.
            var base = this.axBase || ((typeof window.base_url === 'string' && window.base_url) ? window.base_url : '/backend/');
            if (base.charAt(base.length - 1) !== '/') {
                base += '/';
            }
            return base + 'ax/' + path;
        },

        request: function (path, params, method, callback) {
            Ext.Ajax.request({
                url: this.axUrl(path),
                method: method,
                params: params,
                timeout: 60000,
                success: function (resp) {
                    var json = Ext.decode(resp.responseText, true);
                    callback(json, json !== null && json !== undefined, resp.status);
                },
                failure: function (resp) {
                    callback(null, false, resp.status);
                }
            });
        },

        windowHours: function (win) {
            return win === 'now' ? 24 : Math.min(72, Math.max(1, Number(win) || 24));
        },

        // ---- loading -----------------------------------------------------

        load: function (win) {
            var me = this,
                params = { vehs: 1, checked: 1, node: 'src' };

            if (me.loading) {
                return;
            }
            me.loading = true;
            me.cancelled = false;
            me.win = win === 'now' ? 'now' : me.windowHours(win);
            me.hours = me.windowHours(me.win);
            me.fireEvent('beforeload', me);

            me.request('tree.php', params, 'GET', function onTree(tree, ok, status) {
                if (!Ext.isArray(tree) && !me.triedTreeFallback) {
                    // Some installations serve the legacy endpoints from the site root.
                    me.triedTreeFallback = true;
                    me.axBase = '/';
                    me.request('tree.php', params, 'GET', onTree);
                    return;
                }
                if (!Ext.isArray(tree)) {
                    me.loading = false;
                    me.fireEvent('loaderror', me, t('Could not load the object list') + ' (HTTP ' + status + ')');
                    return;
                }
                me.objects = me.collectObjects(tree);
                me.loadObjects();
            });
        },

        stop: function () {
            this.cancelled = true;
        },

        collectObjects: function (tree) {
            var out = [], seen = {};
            function walk(nodes, path) {
                Ext.Array.each(nodes || [], function (n) {
                    if (!n) { return; }
                    var name = cleanText(n.name !== undefined ? n.name : n.text), agentId, group;
                    if (n.leaf) {
                        agentId = String(n.agentid || n.vehid || n.id || '');
                        if (!agentId || seen[agentId]) { return; }
                        seen[agentId] = true;
                        group = path.join(' / ') || cleanText(n.group);
                        out.push({
                            agentId: agentId,
                            name: name || agentId,
                            group: group,
                            region: siteOf(group),
                            path: path.slice(),
                            imei: n.uniqid || ''
                        });
                    } else if (n.children && n.children.length) {
                        walk(n.children, name ? path.concat([name]) : path);
                    }
                });
            }
            walk(tree, []);
            return out;
        },

        loadObjects: function () {
            var me = this,
                queue = me.objects.slice(),
                results = {},
                running = 0,
                done = 0,
                finished = false,
                total = queue.length;

            if (!total) {
                me.finish(results);
                return;
            }

            function next() {
                if (!queue.length || me.cancelled) {
                    if (running === 0 && !finished) {
                        finished = true;
                        me.finish(results);
                    }
                    return;
                }
                var obj = queue.shift();
                running += 1;
                me.objectBattery(obj, me.hours, function (data, error) {
                    results[obj.agentId] = error ? { obj: obj, error: error } : { obj: obj, data: data };
                    running -= 1;
                    done += 1;
                    me.progress = { done: done, total: total };
                    me.fireEvent('progress', me, done, total);
                    if (done < total && !me.cancelled) { me.finish(results, true); } // show each object as soon as it is read
                    Ext.defer(next, 150); // be gentle with the tracking server
                });
            }

            for (var i = 0; i < CONCURRENCY; i += 1) {
                next();
            }
        },

        listSensors: function (agentId, callback) {
            var me = this,
                cached = me.configCache[agentId];

            if (cached && (Date.now() - cached.ts) < me.configTtlMs) {
                callback(cached.sensors);
                return;
            }
            me.request('sensors/sensors.php', { agent_id: agentId }, 'POST', function (json) {
                var list = Ext.isArray(json) ? json : (json && Ext.isArray(json.items) ? json.items :
                    (json && Ext.isArray(json.data) ? json.data : null));
                if (list) {
                    me.configCache[agentId] = { ts: Date.now(), sensors: list };
                }
                callback(list);
            });
        },

        /**
         * Trace rows newest first; stops once enough(rows) returns true.
         * limit=100 is the largest page size the server honours.
         */
        traceRows: function (agentId, fromMs, toMs, enough, maxRows, callback) {
            var me = this,
                rows = [],
                start = 0,
                page = 1;

            function fetch() {
                me.request('sensors_trace.php', {
                    cmd: 'get',
                    agent_id: agentId,
                    ts: traceDate(fromMs),
                    te: traceDate(toMs),
                    page: page,
                    start: start,
                    limit: 100,
                    sort: Ext.encode([{ property: 'ts_date', direction: 'DESC' }])
                }, 'GET', function (json, ok, status) {
                    if (!json || !Ext.isArray(json.items)) {
                        callback(rows.length ? rows : null, (json && json.msg) || ('HTTP ' + status));
                        return;
                    }
                    var items = json.items;
                    rows.push.apply(rows, items);
                    start += items.length;
                    page += 1;
                    if (!items.length || start >= Number(json.totalCount || 0) || enough(rows) || rows.length >= maxRows || me.cancelled) {
                        callback(rows, null);
                        return;
                    }
                    fetch();
                });
            }
            fetch();
        },

        objectBattery: function (obj, hours, callback) {
            var me = this;

            me.listSensors(obj.agentId, function (list) {
                if (!list) {
                    callback(null, t('Could not load the sensor list'));
                    return;
                }
                var sensors = Ext.Array.filter(list, function (s) { return !isAverageSensor(s) && !isStatusSensor(s); }),
                    names = probeNames(sensors),
                    expected = Ext.Object.getKeys(names),
                    now = Date.now(),
                    seen = {},
                    scanned = 0;

                // Objects without wireless probes: one page is enough for the gateway and the
                // latest sensor values (the web report reads more, but shows no probes for them either).
                me.traceRows(obj.agentId, now - hours * 3600000, now + 5 * 60000, function (rs) {
                    for (; scanned < rs.length; scanned += 1) {
                        Ext.Object.each(parseTrigger(rs[scanned].triggerdata), function (k) {
                            var m = /^BatteryVoltage(\d{6,})$/.exec(k);
                            if (m) { seen[m[1]] = true; }
                        });
                    }
                    return expected.length > 0 && Ext.Array.every(expected, function (id) { return seen[id]; });
                }, expected.length ? 2000 : 100, function (rows, error) {
                    if (!rows) {
                        callback(null, t('Sensors Tracing failed') + ': ' + error);
                        return;
                    }
                    callback(me.buildObject(obj, hours, sensors, names, rows), null);
                });
            });
        },

        buildObject: function (obj, hours, sensors, names, rows) {
            var probes = {},
                gwBattery = null,
                gwPower = null,
                now = Date.now(),
                staleBefore = now / 1000 - STALE_HOURS * 3600,
                latest = {};

            rows.sort(function (a, b) { return tsOf(b) - tsOf(a); }); // newest first, even if the server ignored the sort

            Ext.Object.each(names, function (id, name) {
                probes[id] = {
                    idDec: id, idHex: toHex(id), name: name || (t('Probe') + ' ' + toHex(id)),
                    voltage: null, status: 'nodata', platformStatus: null,
                    temperature: null, humidity: null, rssi: null, ts: null
                };
            });

            Ext.Array.each(rows, function (r) {
                var ts = tsOf(r),
                    kv = parseTrigger(r.triggerdata);
                if (gwBattery === null && kv.hasOwnProperty('BatteryVoltage')) { gwBattery = num(kv.BatteryVoltage); }
                if (gwPower === null && kv.hasOwnProperty('PowerVoltage')) { gwPower = num(kv.PowerVoltage); }
                Ext.Object.each(kv, function (k, v) {
                    var m = PROBE_KEY.exec(k),
                        p = m && probes[m[2]]; // only probes registered in the object's sensor list
                    if (!p) { return; }
                    if (m[1] === 'BatteryVoltage') {
                        if (p.voltage === null) { p.voltage = num(v); p.ts = ts; }
                    } else if (m[1] === 'BatteryVoltageStatus') {
                        if (p.platformStatus === null) { p.platformStatus = v; }
                    } else if (m[1] === 'Temperature') {
                        if (p.temperature === null) { p.temperature = num(v); }
                    } else if (m[1] === 'Humidity') {
                        if (p.humidity === null) { p.humidity = num(v); }
                    } else if (m[1] === 'RSSI') {
                        if (p.rssi === null) { p.rssi = num(v); }
                    }
                });
            });

            Ext.Object.each(probes, function (id, p) {
                p.status = (p.ts === null || p.ts < staleBefore) ? 'nodata' : statusOf(p.voltage);
            });

            // Latest value of every configured sensor (numeric keys in trace rows are sensor ids).
            Ext.Array.each(rows, function (r) {
                Ext.Array.each(sensors, function (s) {
                    var id = String(s.sensorid), v = r[id];
                    if (v === null || v === undefined || v === '' || latest[id]) { return; }
                    latest[id] = { v: cleanText(v), ts: tsOf(r) };
                });
            });

            var readings = Ext.Array.map(sensors, function (s) {
                var id = String(s.sensorid), l = latest[id];
                return {
                    sensorid: id, name: cleanText(s.info) || id, field: cleanText(s.sensorfieldname),
                    unit: cleanText(s.measureunit), group: cleanText(s.sensorgroupname),
                    type: cleanText(s.sensortype), func: cleanText(s.sensorinterpret),
                    value: l ? l.v : null, ts: l ? l.ts : null
                };
            }).sort(function (a, b) { return byName(a.name, b.name); });

            return {
                agentid: obj.agentId, windowHours: hours, rows: rows.length,
                lastPacket: rows[0] ? tsOf(rows[0]) : null,
                gateway: { batteryVoltage: gwBattery, powerVoltage: gwPower },
                probes: Ext.Object.getValues(probes).sort(function (a, b) { return byName(a.name, b.name); }),
                sensors: readings
            };
        },

        /** Every packet in the last `hours` that contains this probe, newest first. */
        probeHistory: function (agentId, idDec, hours, callback) {
            var now = Date.now();
            this.traceRows(agentId, now - hours * 3600000, now + 5 * 60000, function () { return false; }, 4000, function (rows, error) {
                if (!rows) {
                    callback(null, error);
                    return;
                }
                var readings = [];
                Ext.Array.each(rows, function (r) {
                    var kv = parseTrigger(r.triggerdata), voltage;
                    if (!Ext.Array.some(['BatteryVoltage', 'BatteryVoltageStatus', 'Temperature', 'Humidity', 'RSSI'], function (k) {
                        return kv.hasOwnProperty(k + idDec);
                    })) {
                        return;
                    }
                    voltage = num(kv['BatteryVoltage' + idDec]);
                    readings.push({
                        ts: tsOf(r), voltage: voltage, status: statusOf(voltage),
                        platformStatus: kv.hasOwnProperty('BatteryVoltageStatus' + idDec) ? kv['BatteryVoltageStatus' + idDec] : null,
                        temperature: num(kv['Temperature' + idDec]), humidity: num(kv['Humidity' + idDec]), rssi: num(kv['RSSI' + idDec])
                    });
                });
                readings.sort(function (a, b) { return b.ts - a.ts; });
                callback(readings, null);
            });
        },

        finish: function (results, partial) {
            var me = this, rows = [], errors = 0;

            me.results = results;
            Ext.Object.each(results, function (agentId, r) {
                if (r.error) {
                    errors += 1;
                    return;
                }
                var probes = r.data.probes,
                    low = 0, nodata = 0;
                Ext.Array.each(probes, function (p) {
                    if (p.status === 'low') { low += 1; }
                    if (p.status === 'nodata') { nodata += 1; }
                });
                // Objects with most Low, then most No reading, come first (as in the web report).
                var groupKey = Ext.String.leftPad(String(999 - low), 3, '0') + Ext.String.leftPad(String(999 - nodata), 3, '0') +
                    r.obj.name.toLowerCase() + '#' + agentId;
                Ext.Array.each(probes, function (p) {
                    rows.push({
                        key: agentId + ':' + p.idDec,
                        agentId: agentId,
                        objectName: r.obj.name,
                        group: r.obj.group,
                        region: r.obj.region,
                        groupKey: groupKey,
                        idDec: p.idDec,
                        idHex: p.idHex,
                        name: p.name,
                        status: p.status,
                        platformStatus: p.platformStatus,
                        voltage: p.voltage,
                        temperature: p.temperature,
                        humidity: p.humidity,
                        rssi: p.rssi,
                        ts: p.ts,
                        rank: STATUS[p.status].rank,
                        voltSort: p.voltage === null ? 99 : p.voltage
                    });
                });
            });

            me.store.loadData(rows);
            if (!partial) {
                me.loading = false;
                me.progress = null;
                me.lastLoad = new Date();
            }
            me.fireEvent('load', me, rows, errors, !!partial);
        },

        /** Objects that have at least one listed probe (others are hidden from the tree and tabs). */
        getTagObjects: function () {
            var results = this.results;
            return Ext.Array.filter(this.objects, function (o) {
                var r = results[o.agentId];
                return r && r.data && r.data.probes.length;
            });
        },

        getFailed: function () {
            var out = [];
            Ext.Object.each(this.results, function (id, r) {
                if (r.error) { out.push(r); }
            });
            return out;
        },

        getSummary: function () {
            var s = { total: 0, normal: 0, low: 0, nodata: 0, objects: Ext.Object.getSize(this.results) };
            this.store.each(function (r) {
                s.total += 1;
                s[r.get('status')] += 1;
            });
            return s;
        },

        getObjectStats: function (agentId) {
            var stats = { total: 0, normal: 0, low: 0, nodata: 0, worst: 'nodata', any: false };
            this.store.each(function (r) {
                if (r.get('agentId') !== agentId) {
                    return;
                }
                stats.total += 1;
                stats[r.get('status')] += 1;
                if (!stats.any || STATUS[r.get('status')].rank > STATUS[stats.worst].rank) {
                    stats.worst = r.get('status');
                }
                stats.any = true;
            });
            return stats;
        }
    });

    // ---------------------------------------------------------------------
    // Battery Status Excel export: per-sheet column choice, order and sorting
    // (port of app/battery/exportConfig.ts + ExportOptions.tsx of the site).
    // The settings are the shared default stored on the server
    // (/api/battery/export-settings); without sign-in they are kept in this browser.
    // ---------------------------------------------------------------------
    var BX = (function () {
        var REPORT_NAME = 'Warehouse Sensor Status',
            LABEL = { normal: 'Normal', low: 'Low', nodata: 'No reading' },
            ORDER = { low: 0, nodata: 1, normal: 2 },
            DATE_FMT = 'dd mmm yyyy hh:mm';

        function xlDate(sec) { return sec ? new Date((sec + TZ_OFFSET_S) * 1000) : null; }
        function numOrText(v) {
            if (v === null || v === undefined || String(v).trim() === '') { return null; }
            var n = Number(v);
            return isFinite(n) ? n : v;
        }
        function col(key, header, width, value, sort, defaultOn, numFmt) {
            return { key: key, header: header, width: width, value: value, sort: sort || value, defaultOn: defaultOn !== false, numFmt: numFmt };
        }
        function objCols(get, on) {
            on = on || {};
            return [
                col('obj', 'Object', 28, function (r) { return get(r).name; }),
                col('site', 'Site', 22, function (r) { return get(r).region; }, null, on.site !== undefined ? on.site : true),
                col('group', 'Group path', 40, function (r) { return get(r).group; }, null, !!on.group),
                col('agent', 'Agent ID', 10, function (r) { return get(r).agentId; }, function (r) { return Number(get(r).agentId) || get(r).agentId; }, !!on.agent),
                col('imei', 'Gateway IMEI', 18, function (r) { return get(r).imei || ''; }, null, !!on.imei)
            ];
        }
        function count(r, s) {
            return r.data ? Ext.Array.filter(r.data.probes, function (p) { return !s || p.status === s; }).length : null;
        }
        var objOf = function (r) { return r.obj; };

        var SHEETS = [
            {
                key: 'battery', title: 'Battery voltage', hint: 'one row per sensor',
                cols: objCols(objOf).concat([
                    col('name', 'Sensor', 44, function (r) { return r.name; }),
                    col('hex', 'Sensor ID', 12, function (r) { return r.idHex; }),
                    col('dec', 'Sensor ID (decimal)', 16, function (r) { return r.idDec; }, null, false),
                    col('v', 'Voltage (V)', 12, function (r) { return r.voltage; }, null, true, '0.000'),
                    col('status', 'Status', 12, function (r) { return LABEL[r.status]; }, function (r) { return ORDER[r.status]; }),
                    col('pstatus', 'Battery Status', 15, function (r) { return r.platformStatus || ''; }, function (r) { return r.platformStatus; }),
                    col('t', 'Temp (°C)', 10, function (r) { return r.temperature; }, null, true, '0.0'),
                    col('h', 'Humidity (%)', 12, function (r) { return r.humidity; }, null, true, '0.0'),
                    col('rssi', 'Signal (RSSI)', 12, function (r) { return r.rssi; }),
                    col('ts', 'Last reading', 18, function (r) { return xlDate(r.ts); }, function (r) { return r.ts; }, true, DATE_FMT)
                ])
            },
            {
                key: 'gateways', title: 'Gateways', hint: 'one row per object',
                cols: objCols(objOf, { imei: true }).concat([
                    col('gwb', 'Gateway battery (V)', 18, function (r) { return r.data ? r.data.gateway.batteryVoltage : null; }, null, true, '0.00'),
                    col('pwr', 'Power supply (V)', 16, function (r) { return r.data ? r.data.gateway.powerVoltage : null; }, null, true, '0.00'),
                    col('last', 'Last packet', 18, function (r) { return xlDate(r.data && r.data.lastPacket); }, function (r) { return r.data && r.data.lastPacket; }, true, DATE_FMT),
                    col('n', 'Sensors', 9, function (r) { return count(r); }),
                    col('low', 'Low', 7, function (r) { return count(r, 'low'); }),
                    col('nodata', 'No reading', 11, function (r) { return count(r, 'nodata'); }),
                    col('normal', 'Normal', 9, function (r) { return count(r, 'normal'); }),
                    col('err', 'Error', 40, function (r) { return r.error || ''; }, function (r) { return r.error; })
                ])
            },
            {
                key: 'sensors', title: 'All sensors', hint: 'every configured sensor with its latest value',
                cols: objCols(objOf, { site: false }).concat([
                    col('name', 'Sensor', 44, function (r) { return r.name; }),
                    col('field', 'Field', 30, function (r) { return r.field; }),
                    col('sgroup', 'Sensor group', 22, function (r) { return r.group; }),
                    col('type', 'Type', 12, function (r) { return r.type; }),
                    col('func', 'Function', 22, function (r) { return r.func; }, null, false),
                    col('unit', 'Unit', 8, function (r) { return r.unit; }),
                    col('value', 'Latest value', 14, function (r) { return numOrText(r.value); }),
                    col('ts', 'Last reading', 18, function (r) { return xlDate(r.ts); }, function (r) { return r.ts; }, true, DATE_FMT),
                    col('id', 'Sensor ID', 10, function (r) { return r.sensorid; }, function (r) { return Number(r.sensorid) || r.sensorid; })
                ])
            }
        ];
        function sheetDef(k) { return Ext.Array.findBy(SHEETS, function (s) { return s.key === k; }); }

        var DEFAULT_SORT = {
            battery: [{ key: 'obj', dir: 'asc' }, { key: 'status', dir: 'asc' }],
            gateways: [{ key: 'obj', dir: 'asc' }, { key: '', dir: 'asc' }],
            sensors: [{ key: 'obj', dir: 'asc' }, { key: 'name', dir: 'asc' }]
        };
        function defaultSheet(k) {
            return {
                on: true,
                cols: Ext.Array.map(sheetDef(k).cols, function (c) { return { key: c.key, on: c.defaultOn }; }),
                sort: Ext.Array.map(DEFAULT_SORT[k], function (s) { return { key: s.key, dir: s.dir }; })
            };
        }
        function defaultCfg() {
            return { sheets: { battery: defaultSheet('battery'), gateways: defaultSheet('gateways'), sensors: defaultSheet('sensors') }, onlyFiltered: false };
        }

        /** Accept a saved config (any older shape) and make it valid against the current column definitions. */
        function normalize(raw) {
            var d = defaultCfg(),
                r = (raw && typeof raw === 'object') ? raw : {},
                sheetsIn = r.sheets || (r.cols ? { battery: { cols: r.cols, sort: r.sort }, gateways: { on: r.gateways !== false }, sensors: { on: r.allSensors !== false } } : {}),
                sheets = {};
            Ext.Array.each(SHEETS, function (s) {
                var inp = sheetsIn[s.key] || {},
                    known = {},
                    cols = [], seen = {}, missing, sort;
                Ext.Array.each(s.cols, function (c) { known[c.key] = true; });
                Ext.Array.each(Ext.isArray(inp.cols) ? inp.cols : [], function (c) {
                    if (c && known[c.key] && !seen[c.key]) { seen[c.key] = true; cols.push({ key: c.key, on: !!c.on }); }
                });
                missing = Ext.Array.filter(d.sheets[s.key].cols, function (c) { return !seen[c.key]; }); // columns added in a later version
                sort = Ext.Array.map([0, 1], function (i) {
                    var x = Ext.isArray(inp.sort) ? inp.sort[i] : null;
                    if (!x) { return d.sheets[s.key].sort[i] || { key: '', dir: 'asc' }; }
                    return { key: x.key && known[x.key] ? x.key : '', dir: x.dir === 'desc' ? 'desc' : 'asc' };
                });
                sheets[s.key] = { on: inp.on !== false, cols: cols.length ? cols.concat(missing) : d.sheets[s.key].cols, sort: sort };
            });
            return { sheets: sheets, onlyFiltered: !!r.onlyFiltered, updatedAt: typeof r.updatedAt === 'string' ? r.updatedAt : undefined };
        }

        function compareBy(cols, sort) {
            var byKey = {};
            Ext.Array.each(cols, function (c) { byKey[c.key] = c; });
            return function (a, b) {
                for (var i = 0; i < sort.length; i += 1) {
                    var c = byKey[sort[i].key];
                    if (!c) { continue; }
                    var x = c.sort(a), y = c.sort(b),
                        xe = x === null || x === undefined || x === '',
                        ye = y === null || y === undefined || y === '';
                    if (xe && ye) { continue; }
                    if (xe) { return 1; } // empty values always last
                    if (ye) { return -1; }
                    var r = (typeof x === 'number' && typeof y === 'number') ? x - y :
                        String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' });
                    if (r) { return sort[i].dir === 'desc' ? -r : r; }
                }
                return 0;
            };
        }

        function buildWorkbook(ExcelJS, cfg, data) {
            var wb = new ExcelJS.Workbook();
            wb.title = REPORT_NAME;
            wb.creator = 'Al Safi';
            Ext.Array.each(SHEETS, function (s) {
                var sc = cfg.sheets[s.key], byKey = {}, cols, ws, rows;
                if (!sc.on) { return; }
                Ext.Array.each(s.cols, function (c) { byKey[c.key] = c; });
                cols = Ext.Array.filter(Ext.Array.map(Ext.Array.filter(sc.cols, function (c) { return c.on; }), function (c) { return byKey[c.key]; }), Boolean);
                if (!cols.length) { return; }
                ws = wb.addWorksheet(s.title);
                ws.columns = Ext.Array.map(cols, function (c) { return { header: c.header, key: c.key, width: c.width }; });
                xlHead(ws.getRow(1));
                ws.views = [{ state: 'frozen', ySplit: 1 }];
                rows = data[s.key].slice().sort(compareBy(s.cols, sc.sort));
                Ext.Array.each(rows, function (r) {
                    var obj = {}, row, st, argb;
                    Ext.Array.each(cols, function (c) { var v = c.value(r); obj[c.key] = v === undefined ? null : v; });
                    row = ws.addRow(obj);
                    Ext.Array.each(cols, function (c) { if (c.numFmt) { row.getCell(c.key).numFmt = c.numFmt; } });
                    st = s.key === 'battery' ? r.status : (s.key === 'gateways' && r.error ? 'low' : null);
                    argb = st === 'low' ? 'FFFDECEC' : (st === 'nodata' ? 'FFF1F3F5' : null);
                    if (argb) { xlFill(row, argb); }
                });
                ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
            });
            if (!wb.worksheets.length) {
                throw new Error('Every sheet is switched off or has no columns — enable at least one in Export options.');
            }
            return wb.xlsx.writeBuffer();
        }

        // ---- shared default settings -------------------------------------
        var state = { cfg: normalize(storageGet('exportCfg', null)), saveState: 'idle', saveError: '', loaded: false, listeners: [] };

        function notify() { Ext.defer(function () { Ext.Array.each(state.listeners.slice(), function (fn) { fn(); }); }, 1); }

        function load() {
            if (state.loaded) { return; }
            state.loaded = true;
            Api.call('/api/battery/export-settings', { silent: true }).then(function (r) {
                if (r.settings && !state.dirty) { state.cfg = normalize(r.settings); notify(); }
            }, function () { /* not signed in: keep this browser's copy */ });
        }

        function set(fn) {
            state.dirty = true;
            state.cfg = fn(state.cfg);
            storageSet('exportCfg', state.cfg);
            state.saveState = 'saving';
            notify();
            clearTimeout(state.timer);
            state.timer = setTimeout(function () {
                var body = Ext.apply({}, state.cfg);
                delete body.updatedAt;
                Api.call('/api/battery/export-settings', { json: body, silent: true }).then(function (r) {
                    state.saveState = 'saved';
                    state.saveError = '';
                    state.cfg.updatedAt = r.settings && r.settings.updatedAt;
                    notify();
                }, function (e) {
                    state.saveState = Api.token ? 'error' : 'local';
                    state.saveError = errorText(e);
                    notify();
                });
            }, 600);
        }

        return {
            REPORT_NAME: REPORT_NAME, SHEETS: SHEETS, sheetDef: sheetDef, defaultSheet: defaultSheet, defaultCfg: defaultCfg,
            buildWorkbook: buildWorkbook, state: state, load: load, set: set
        };
    }());

    /** Export options window (port of the site's ExportOptions panel). */
    Ext.define('Store.battery_status.view.ExportOptions', {
        extend: 'Ext.window.Window',
        title: 'Export options',
        width: 760,
        height: 560,
        modal: true,
        layout: 'fit',
        scrollable: false,

        initComponent: function () {
            var me = this;
            me.tab = 'battery';
            me.title = t('Export options');
            me.width = Math.min(760, Ext.getBody().getViewSize().width - 30);
            me.height = Math.min(600, Ext.getBody().getViewSize().height - 30);
            me.body0 = Ext.create('Ext.container.Container', { scrollable: true, padding: 12, layout: 'anchor' });
            me.items = [me.body0];
            me.saveLabel = Ext.create('Ext.toolbar.TextItem', { text: '' });
            me.dockedItems = [{
                xtype: 'toolbar', dock: 'top',
                items: Ext.Array.map(BX.SHEETS, function (s) {
                    return { text: t(s.title), sheetKey: s.key, enableToggle: true, toggleGroup: 'bs-sheets', allowDepress: false, pressed: s.key === 'battery',
                        handler: function (b) { me.tab = b.sheetKey; me.refresh(); } };
                }).concat(['->', me.saveLabel])
            }];
            me.buttons = [
                { text: t('Reset all sheets'), handler: function () { BX.set(function () { return BX.defaultCfg(); }); } },
                '->',
                { text: t('Close'), handler: function () { me.close(); } },
                { text: t('Download Excel'), iconCls: 'fa fa-file-excel', itemId: 'dl', handler: function () { me.onDownload(); } }
            ];
            me.listener = function () { if (!me.destroyed && !me.isDestroyed) { me.refresh(); } };
            BX.state.listeners.push(me.listener);
            me.on('destroy', function () { Ext.Array.remove(BX.state.listeners, me.listener); });
            me.callParent(arguments);
            me.on('afterrender', me.refresh, me);
        },

        refresh: function () {
            var me = this,
                cfg = BX.state.cfg,
                def = BX.sheetDef(me.tab),
                sc = cfg.sheets[me.tab],
                byKey = {},
                visible = Ext.Array.filter(sc.cols, function (c) { return c.on; }).length,
                sheetsOn = Ext.Array.filter(BX.SHEETS, function (s) { return cfg.sheets[s.key].on && Ext.Array.some(cfg.sheets[s.key].cols, function (c) { return c.on; }); }).length,
                ss = BX.state.saveState,
                items = [];

            Ext.Array.each(def.cols, function (c) { byKey[c.key] = c; });
            function setSheet(fn) {
                BX.set(function (c) {
                    var n = Ext.apply({}, c);
                    n.sheets = Ext.apply({}, c.sheets);
                    n.sheets[me.tab] = fn(c.sheets[me.tab]);
                    return n;
                });
            }
            function setSort(level, patch) {
                setSheet(function (s) {
                    var sort = Ext.Array.map([0, 1], function (i) { return Ext.apply({}, s.sort[i] || { key: '', dir: 'asc' }); });
                    Ext.apply(sort[level], patch);
                    return Ext.apply({}, { sort: sort }, s);
                });
            }
            function move(i, d) {
                setSheet(function (s) {
                    var j = i + d, cols = s.cols.slice(), tmp;
                    if (j < 0 || j >= cols.length) { return s; }
                    tmp = cols[i]; cols[i] = cols[j]; cols[j] = tmp;
                    return Ext.apply({}, { cols: cols }, s);
                });
            }

            me.saveLabel.setText('<span class="bs-muted">' + enc(
                ss === 'saving' ? t('Saving as default…') :
                ss === 'saved' ? '✓ ' + t('Saved as default for everyone') :
                ss === 'error' ? t('Not saved') + ': ' + BX.state.saveError :
                ss === 'local' ? t('Saved in this browser (sign in to save for everyone)') :
                t('Changes are saved as the default')) + '</span>');

            Ext.Array.each(me.down('toolbar[dock=top]').query('button'), function (b) {
                var on = cfg.sheets[b.sheetKey].on;
                b.setText(t(BX.sheetDef(b.sheetKey).title) + (on ? '' : ' (' + t('off') + ')'));
            });

            items.push({
                xtype: 'checkbox', boxLabel: t('Include the') + ' “' + t(def.title) + '” ' + t('sheet') + ' <span class="bs-muted">— ' + enc(t(def.hint)) + '</span>',
                checked: sc.on,
                listeners: { change: function (f, v) { setSheet(function (s) { return Ext.apply({}, { on: v }, s); }); } }
            });

            var colRows = Ext.Array.map(sc.cols, function (c, i) {
                var d = byKey[c.key];
                return {
                    xtype: 'container', layout: { type: 'hbox', align: 'middle' }, cls: c.on ? '' : 'bs-off',
                    items: [
                        { xtype: 'checkbox', boxLabel: enc(d.header), checked: c.on, disabled: !sc.on, flex: 1,
                            listeners: { change: function (f, v) { setSheet(function (s) {
                                return Ext.apply({}, { cols: Ext.Array.map(s.cols, function (y) { return y.key === c.key ? { key: y.key, on: v } : y; }) }, s);
                            }); } } },
                        { xtype: 'button', text: '▲', disabled: !sc.on || i === 0, margin: '0 2', handler: function () { move(i, -1); } },
                        { xtype: 'button', text: '▼', disabled: !sc.on || i === sc.cols.length - 1, handler: function () { move(i, 1); } }
                    ]
                };
            });

            var sortRows = Ext.Array.map([0, 1], function (level) {
                var s = sc.sort[level] || { key: '', dir: 'asc' };
                return {
                    xtype: 'container', layout: 'hbox', margin: '0 0 6 0',
                    items: [
                        { xtype: 'component', html: '<span class="bs-muted">' + (level === 0 ? t('First') : t('Then')) + '</span>', width: 50, margin: '6 0 0 0' },
                        { xtype: 'combobox', editable: false, queryMode: 'local', value: s.key, disabled: !sc.on, flex: 1,
                            store: [['', '— ' + t('none') + ' —']].concat(Ext.Array.map(def.cols, function (c) { return [c.key, c.header]; })),
                            listeners: { select: function (f, r) { setSort(level, { key: f.getValue() || '' }); } } },
                        { xtype: 'combobox', editable: false, queryMode: 'local', value: s.dir, width: 170, margin: '0 0 0 6', disabled: !sc.on || !s.key,
                            store: [['asc', t('Ascending (A→Z, 0→9)')], ['desc', t('Descending (Z→A, 9→0)')]],
                            listeners: { select: function (f, r) { setSort(level, { dir: f.getValue() }); } } }
                    ]
                };
            });

            var right = [
                { xtype: 'component', html: '<b>' + enc(t('Sort rows by')) + '</b>', margin: '0 0 6 0' }
            ].concat(sortRows).concat([
                { xtype: 'component', cls: 'bs-muted', html: (me.tab === 'battery' ? enc(t('Status sorts Low → No reading → Normal (ascending).')) + ' ' : '') +
                    enc(t('Empty values always go last. Sorting by a hidden column works too.')) }
            ]);
            if (me.tab === 'battery') {
                right.push({ xtype: 'component', html: '<b>' + enc(t('Rows')) + '</b>', margin: '14 0 4 0' });
                right.push({
                    xtype: 'checkbox', checked: cfg.onlyFiltered, disabled: !sc.on,
                    boxLabel: enc(t('Only rows matching the on-screen filter')) + ' (' + me.counts().filtered + ' ' + t('of') + ' ' + me.counts().total + ')',
                    listeners: { change: function (f, v) { BX.set(function (c) { return Ext.apply({}, { onlyFiltered: v }, c); }); } }
                });
            }
            right.push({
                xtype: 'container', layout: 'hbox', margin: '14 0 0 0',
                items: [
                    { xtype: 'button', text: t('Show all'), disabled: !sc.on, margin: '0 6 0 0',
                        handler: function () { setSheet(function (s) { return Ext.apply({}, { cols: Ext.Array.map(s.cols, function (y) { return { key: y.key, on: true }; }) }, s); }); } },
                    { xtype: 'button', text: t('Reset this sheet'), disabled: !sc.on,
                        handler: function () { setSheet(function () { return Ext.apply(BX.defaultSheet(me.tab), { on: sc.on }); }); } }
                ]
            });
            if (!sheetsOn) {
                right.push({ xtype: 'component', html: '<span style="color:#b3261e">' + enc(t('All sheets are off — switch at least one on.')) + '</span>' });
            }
            if (cfg.updatedAt) {
                right.push({ xtype: 'component', cls: 'bs-muted', margin: '8 0 0 0',
                    html: enc(t('Default last changed')) + ' ' + enc(new Date(cfg.updatedAt).toLocaleString('en-GB', { timeZone: TZ, day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })) + '.' });
            }

            items.push({
                xtype: 'container', layout: { type: 'hbox', align: 'begin' }, style: { opacity: sc.on ? 1 : 0.5 },
                items: [
                    { xtype: 'container', flex: 1, margin: '8 16 0 0', items: [
                        { xtype: 'component', html: '<b>' + enc(t('Columns')) + '</b> <span class="bs-muted">(' + visible + ' ' + t('of') + ' ' + sc.cols.length + ' ' +
                            enc(t('shown) — tick to show, arrows to reorder')) + '</span>', margin: '0 0 6 0' }
                    ].concat(colRows) },
                    { xtype: 'container', flex: 1, margin: '8 0 0 0', items: right }
                ]
            });

            var dl = me.down('#dl');
            if (dl) { dl.setDisabled(!sheetsOn); }
            me.body0.suspendLayouts();
            me.body0.removeAll();
            me.body0.add(items);
            me.body0.resumeLayouts(true);
        },

        counts: function () {
            return this.getCounts ? this.getCounts() : { filtered: 0, total: 0 };
        }
    });

    // ---------------------------------------------------------------------
    // Shared grid for sensor rows (same columns as the web report).
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.view.SensorGrid', {
        extend: 'Ext.grid.Panel',
        cls: 'bs-grid',
        hideObject: false,

        initComponent: function () {
            var me = this;

            me.viewConfig = Ext.apply({
                emptyText: '<div class="bs-empty">' + enc(me.emptyMessage || t('No sensors found')) + '</div>',
                deferEmptyText: false,
                stripeRows: true,
                getRowClass: function (r) { return 'bs-row-' + r.get('status'); }
            }, me.viewConfig || {});

            me.columns = [
                {
                    text: t('Sensor'), dataIndex: 'name', flex: 2, minWidth: 180,
                    renderer: function (v) { return '<i class="fa fa-chart-line bs-muted"></i> ' + enc(v); }
                },
                {
                    text: t('Sensor ID'), dataIndex: 'idHex', width: 105,
                    renderer: function (v, meta, r) {
                        meta.tdAttr = 'data-qtip="' + enc(enc(t('Decimal') + ': ' + r.get('idDec'))) + '"';
                        return '<code>' + enc(v) + '</code>';
                    }
                },
                { text: t('Object'), dataIndex: 'objectName', flex: 1, minWidth: 150, hidden: me.hideObject, renderer: enc },
                { text: t('Site'), dataIndex: 'region', width: 120, hidden: me.hideObject, renderer: enc },
                {
                    text: t('Voltage'), dataIndex: 'voltSort', width: 100, align: 'right',
                    renderer: function (v, meta, r) {
                        var volt = r.get('voltage');
                        if (volt === null || volt === undefined) { return '—'; }
                        if (r.get('status') === 'nodata') {
                            meta.tdAttr = 'data-qtip="' + enc(t('Last known voltage — too old to count')) + '"';
                            return '<span class="bs-muted">' + volt.toFixed(3) + ' V<br>' + enc(t('last known')) + '</span>';
                        }
                        return '<b>' + volt.toFixed(3) + ' V</b>';
                    }
                },
                {
                    text: t('Status'), dataIndex: 'rank', width: 150,
                    renderer: function (v, meta, r) {
                        var ps = r.get('platformStatus'), label = STATUS[r.get('status')].label;
                        return statusPill(r.get('status')) +
                            (ps && ps.toLowerCase() !== label.toLowerCase() ? '<div class="bs-muted">' + enc(t('Battery Status')) + ': ' + enc(ps) + '</div>' : '');
                    }
                },
                { text: t('Temp °C'), dataIndex: 'temperature', width: 80, align: 'right', renderer: function (v) { return fixed(v); } },
                { text: t('Hum %'), dataIndex: 'humidity', width: 75, align: 'right', renderer: function (v) { return fixed(v); } },
                { text: t('Signal'), dataIndex: 'rssi', width: 65, align: 'right', hidden: true, renderer: function (v) { return v === null ? '—' : enc(v); } },
                { text: t('Last reading'), dataIndex: 'ts', width: 165, renderer: tsHtml }
            ];

            me.callParent(arguments);

            me.on('itemclick', function (view, record) {
                if (me.module) {
                    me.module.showHistory(record);
                }
            });
        }
    });

    // ---------------------------------------------------------------------
    // Main panel: nested tabs Overview / Alerts / Warehouses(per object).
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.view.MainPanel', {
        extend: 'Ext.tab.Panel',
        cls: 'bs-main',

        initComponent: function () {
            var me = this,
                source = me.service.store;

            me.statusFilter = 'all';
            me.searchText = '';

            me.overviewStore = Ext.create('Ext.data.ChainedStore', {
                source: source,
                groupField: 'groupKey'
            });

            me.alertStore = Ext.create('Ext.data.ChainedStore', {
                source: source,
                filters: [function (r) { return r.get('status') !== 'normal'; }]
            });

            me.summaryCmp = Ext.create('Ext.Component', {
                cls: 'bs-summary',
                html: '<div class="bs-muted">' + enc(t('Loading battery status...')) + '</div>',
                listeners: {
                    click: {
                        element: 'el',
                        fn: function (e) {
                            var tile = e.getTarget('[data-status]');
                            if (tile) {
                                var s = tile.getAttribute('data-status');
                                me.setStatusFilter(me.statusFilter === s ? 'all' : s);
                            }
                        }
                    }
                }
            });

            me.overviewGrid = Ext.create('Store.battery_status.view.SensorGrid', {
                store: me.overviewStore,
                module: me.module,
                flex: 1,
                features: [{
                    ftype: 'grouping',
                    startCollapsed: false,
                    groupHeaderTpl: [
                        '{[this.title(values)]}',
                        {
                            title: function (values) {
                                var rec = values.rows && values.rows[0] || values.children && values.children[0];
                                if (!rec) { return enc(values.name); }
                                var stats = me.service.getObjectStats(rec.get('agentId')),
                                    res = me.service.results[rec.get('agentId')],
                                    gw = res && res.data ? res.data.gateway.batteryVoltage : null,
                                    last = res && res.data ? res.data.lastPacket : null;
                                return enc(rec.get('objectName')) + ' <span class="bs-muted">· ' + enc(rec.get('region')) + '</span> ' +
                                    (stats.low ? '<span class="bs-pill bs-alert">' + stats.low + ' ' + enc(t('low')) + '</span> ' : '') +
                                    (stats.nodata ? '<span class="bs-pill bs-nodata">' + stats.nodata + ' ' + enc(t('no reading')) + '</span> ' : '') +
                                    (stats.normal ? '<span class="bs-pill bs-ok">' + stats.normal + ' ' + enc(t('normal')) + '</span> ' : '') +
                                    ' <span class="bs-muted">' + enc(t('Gateway')) + ' ' + fixed(gw, 2) + ' V · ' + tsHtml(last) + '</span>';
                            }
                        }
                    ]
                }],
                emptyMessage: t('Nothing matches the current filter.')
            });

            me.alertsGrid = Ext.create('Store.battery_status.view.SensorGrid', {
                store: me.alertStore,
                module: me.module,
                emptyMessage: t('Every sensor is Normal.')
            });

            me.objectTabs = Ext.create('Ext.tab.Panel', {
                cls: 'bs-object-tabs',
                tabPosition: 'left',
                tabRotation: 0,
                tabBar: { width: 230 },
                items: []
            });

            me.items = [
                {
                    title: t('Overview'),
                    iconCls: 'fa fa-gauge-high',
                    itemId: 'overview',
                    layout: { type: 'vbox', align: 'stretch' },
                    tbar: [
                        {
                            xtype: 'textfield',
                            itemId: 'search',
                            emptyText: t('Filter results…'),
                            width: 220,
                            listeners: {
                                change: function (field, value) { me.searchText = String(value || '').toLowerCase(); me.applyFilters(); },
                                buffer: 250
                            }
                        },
                        {
                            xtype: 'combobox',
                            itemId: 'statusFilter',
                            width: 140,
                            editable: false,
                            queryMode: 'local',
                            value: 'all',
                            store: [['all', t('All statuses')], ['low', t('Low')], ['nodata', t('No reading')], ['normal', t('Normal')]],
                            listeners: {
                                change: function (field, value) { me.setStatusFilter(value, true); }
                            }
                        },
                        { text: t('Expand all'), handler: function () { me.overviewGrid.getView().getFeature(0).expandAll(); } },
                        { text: t('Collapse all'), handler: function () { me.overviewGrid.getView().getFeature(0).collapseAll(); } },
                        '->',
                        { xtype: 'tbtext', itemId: 'lastLoad', cls: 'bs-muted', text: '' },
                        { text: t('Stop'), itemId: 'stopBtn', iconCls: 'fa fa-stop', hidden: true, handler: function () { me.service.stop(); } },
                        { text: t('Fetch battery status'), itemId: 'fetchBtn', iconCls: 'fa fa-rotate', handler: me.refresh, scope: me },
                        { text: t('Print'), iconCls: 'fa fa-print', handler: function () { window.print(); } },
                        { text: '⚙ ' + t('Export options'), handler: me.openExportOptions, scope: me },
                        { text: t('Download Excel'), iconCls: 'fa fa-file-excel', handler: me.download, scope: me }
                    ],
                    items: [me.summaryCmp, me.overviewGrid]
                },
                {
                    title: t('Alerts'),
                    iconCls: 'fa fa-triangle-exclamation',
                    itemId: 'alerts',
                    layout: 'fit',
                    items: [me.alertsGrid]
                },
                {
                    title: t('Warehouses'),
                    iconCls: 'fa fa-warehouse',
                    itemId: 'warehouses',
                    layout: 'fit',
                    items: [me.objectTabs]
                }
            ];

            me.callParent(arguments);

            me.service.on({
                beforeload: function () {
                    me.setBusy(true);
                    me.renderSummary();
                },
                progress: function () { me.renderSummary(); },
                load: me.onDataLoad,
                loaderror: function (data, msg) {
                    me.setBusy(false);
                    me.summaryCmp.setHtml('<div class="bs-error">' + enc(msg) + '</div>');
                },
                scope: me
            });
        },

        setBusy: function (busy) {
            var fetch = this.down('#fetchBtn'), stop = this.down('#stopBtn');
            if (fetch) { fetch.setDisabled(busy); fetch.setText(busy ? t('Fetching…') : t('Fetch battery status')); }
            if (stop) { stop.setHidden(!busy); }
        },

        refresh: function () {
            if (this.module) {
                this.module.reload();
            }
        },

        /** Rows for the three sheets, from the last fetch (battery rows optionally only the filtered ones). */
        exportData: function (onlyFiltered) {
            var results = Ext.Object.getValues(this.service.results),
                keep = null, battery = [], sensors = [];
            if (onlyFiltered) {
                keep = {};
                this.overviewStore.each(function (r) { keep[r.get('key')] = true; });
            }
            Ext.Array.each(results, function (r) {
                if (!r.data) { return; }
                Ext.Array.each(r.data.probes, function (p) {
                    if (!keep || keep[r.obj.agentId + ':' + p.idDec]) { battery.push(Ext.apply({ obj: r.obj }, p)); }
                });
                Ext.Array.each(r.data.sensors, function (s) { sensors.push(Ext.apply({ obj: r.obj }, s)); });
            });
            return { battery: battery, gateways: results, sensors: sensors };
        },

        counts: function () {
            return { filtered: this.overviewStore.getCount(), total: this.service.store.getCount() };
        },

        openExportOptions: function () {
            var me = this;
            BX.load();
            Ext.create('Store.battery_status.view.ExportOptions', {
                getCounts: function () { return me.counts(); },
                onDownload: function () { me.download(); }
            }).show();
        },

        download: function () {
            var me = this, cfg = BX.state.cfg;
            if (!me.service.lastLoad) {
                Ext.Msg.alert(t('Battery Status'), t('Fetch the battery status first.'));
                return;
            }
            me.setLoading(t('Building Excel…'));
            ensureExcel().then(function (ExcelJS) {
                return BX.buildWorkbook(ExcelJS, cfg, me.exportData(cfg.onlyFiltered));
            }).then(function (buf) {
                me.setLoading(false);
                saveFile(buf, BX.REPORT_NAME + ' ' + stampNow() + '.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            }).catch(function (e) {
                me.setLoading(false);
                Ext.Msg.alert(t('Download Excel'), enc(errorText(e)));
            });
        },

        setStatusFilter: function (value, fromCombo) {
            this.statusFilter = value || 'all';
            if (!fromCombo) {
                var combo = this.down('#statusFilter');
                if (combo) { combo.setValue(this.statusFilter); }
            }
            this.applyFilters();
            this.renderSummary();
        },

        applyFilters: function () {
            var status = this.statusFilter, q = this.searchText;
            this.overviewStore.clearFilter();
            if (status !== 'all' || q) {
                this.overviewStore.filterBy(function (r) {
                    return (status === 'all' || r.get('status') === status) &&
                        (!q || (r.get('objectName') + ' ' + r.get('group') + ' ' + r.get('name') + ' ' + r.get('idHex') + ' ' + r.get('idDec'))
                            .toLowerCase().indexOf(q) !== -1);
                });
            }
        },

        onDataLoad: function (data, rows, errors, partial) {
            var lastLoad = this.down('#lastLoad'),
                alertsTab = this.child('#alerts'),
                s = data.getSummary(),
                alerts = s.low + s.nodata;

            this.setBusy(!!partial);
            this.applyFilters();
            this.renderSummary();

            if (lastLoad && !partial) {
                lastLoad.setText(t('fetched') + ' ' + Ext.Date.format(new Date(), 'H:i') + ' · ' +
                    (data.win === 'now' ? t('Now') : t('last') + ' ' + data.win + ' h'));
            }
            if (alertsTab && alertsTab.tab) {
                alertsTab.setTitle(t('Alerts') + (alerts ? ' <span class="bs-count">' + alerts + '</span>' : ''));
            }

            this.syncObjectTabs();
        },

        renderSummary: function () {
            var me = this, data = me.service, s, failed, prog = data.progress,
                progHtml = data.loading ? '<div class="bs-progress"><div style="width:' + (prog ? Math.round(100 * prog.done / prog.total) : 0) + '%"></div></div>' +
                    '<div class="bs-muted bs-note">' + (prog ? prog.done + ' / ' + prog.total + ' ' + enc(t('objects read')) + ' — ' : '') + enc(t('Loading battery status… results appear as each object is read.')) + '</div>' : '';
            me.setEmptyTexts(data.loading);
            if (!data.lastLoad && !data.store.getCount()) {
                me.summaryCmp.setHtml(progHtml || '<div class="bs-muted">' + enc(t('Loading objects…')) + '</div>');
                return;
            }
            s = data.getSummary();
            failed = data.getFailed();
            me.summaryCmp.setHtml(progHtml +
                '<div class="bs-tiles">' +
                me.tile(t('Low battery') + ' (&lt; ' + LOW_VOLTAGE + ' V)', s.low, 'low') +
                me.tile(t('No reading') + ' (&gt; ' + STALE_HOURS + ' h)', s.nodata, 'nodata') +
                me.tile(t('Normal') + ' (≥ ' + LOW_VOLTAGE + ' V)', s.normal, 'normal') +
                '<div class="bs-tile bs-tile-total"><i class="fa fa-warehouse"></i><div class="bs-tile-value">' + s.total + ' / ' + s.objects +
                '</div><div class="bs-tile-label">' + enc(t('Sensors / objects')) + '</div></div>' +
                '</div>' +
                '<div class="bs-muted bs-note">' + enc(t('Latest battery voltage of every wireless sensor.')) + ' ' +
                LOW_VOLTAGE.toFixed(1) + ' V ' + enc(t('and above is normal; below')) + ' ' + LOW_VOLTAGE.toFixed(1) + ' V ' + enc(t('is low.')) + ' ' +
                enc(t('A sensor with no reading for more than')) + ' ' + STALE_HOURS + ' h ' + enc(t('counts as No reading, not low.')) + ' ' +
                enc(t('Click a sensor for its readings in the last 24 hours.')) + '</div>' +
                (failed.length ? '<div class="bs-error">' + failed.length + ' ' + enc(t('object(s) failed')) + ': ' +
                    Ext.Array.map(failed.slice(0, 5), function (f) { return enc(f.obj.name + ': ' + f.error); }).join('<br>') +
                    (failed.length > 5 ? '<br>…' : '') + '</div>' : '')
            );
        },

        /** While loading, the empty grids say so instead of "Nothing matches". */
        setEmptyTexts: function (loading) {
            Ext.Array.each([[this.overviewGrid, t('Nothing matches the current filter.')], [this.alertsGrid, t('Every sensor is Normal.')]], function (x) {
                var view = x[0].getView(), txt = '<div class="bs-empty">' + enc(loading ? t('Loading… results appear as each object is read.') : x[1]) + '</div>';
                if (view.emptyText !== txt) { view.emptyText = txt; if (view.rendered && !x[0].getStore().getCount()) { view.refresh(); } }
            });
        },

        // Label is pre-escaped by the caller (it contains &lt; / &gt;).
        tile: function (label, value, status) {
            return '<div class="bs-tile ' + STATUS[status].tile + (this.statusFilter === status ? ' bs-tile-on' : '') +
                '" data-status="' + status + '" style="cursor:pointer"><i class="' + STATUS[status].icon + '"></i>' +
                '<div class="bs-tile-value">' + enc(value) + '</div>' +
                '<div class="bs-tile-label">' + label + '</div></div>';
        },

        // One nested tab per object (warehouse) with its own grid and gateway info.
        syncObjectTabs: function () {
            var me = this,
                data = me.service,
                activeId = me.objectTabs.getActiveTab() ? me.objectTabs.getActiveTab().agentId : null;

            me.tabsById = me.tabsById || {};

            Ext.Array.each(data.getTagObjects(), function (obj) {
                var stats = data.getObjectStats(obj.agentId),
                    res = data.results[obj.agentId],
                    gw = res.data.gateway,
                    info = enc(obj.group) + '  <span class="bs-muted">' + enc(t('Gateway battery')) + ': ' + fixed(gw.batteryVoltage, 2) + ' V · ' +
                        enc(t('Power supply')) + ': ' + fixed(gw.powerVoltage, 2) + ' V · ' + enc(t('Last packet')) + ': ' + tsHtml(res.data.lastPacket) + '</span>',
                    tab = me.tabsById[obj.agentId],
                    title = enc(obj.name) + ' <span class="bs-tab-badge ' + STATUS[stats.worst].cls + '">' +
                        (stats.total ? stats.normal + '/' + stats.total : '-') + '</span>';

                if (!tab) {
                    tab = me.objectTabs.add({
                        title: title,
                        agentId: obj.agentId,
                        layout: 'fit',
                        items: [Ext.create('Store.battery_status.view.SensorGrid', {
                            hideObject: true,
                            module: me.module,
                            store: Ext.create('Ext.data.ChainedStore', {
                                source: data.store,
                                filters: [{ property: 'agentId', value: obj.agentId, exactMatch: true }]
                            }),
                            title: info,
                            tools: [{
                                xtype: 'button',
                                text: t('All sensors'),
                                iconCls: 'fa fa-list',
                                handler: function () { me.module.showAllSensors(obj.agentId); }
                            }],
                            emptyMessage: t('No sensors on this object.')
                        })]
                    });
                    me.tabsById[obj.agentId] = tab;
                } else {
                    tab.setTitle(title);
                    tab.down('grid').setTitle(info);
                }
                tab.setIconCls(STATUS[stats.worst].icon + ' ' + STATUS[stats.worst].cls);
            });

            if (activeId && me.tabsById[activeId]) {
                me.objectTabs.setActiveTab(me.tabsById[activeId]);
            } else if (!me.objectTabs.getActiveTab() && me.objectTabs.items.getCount()) {
                me.objectTabs.setActiveTab(0);
            }
        },

        showObject: function (agentId) {
            var tab = this.tabsById && this.tabsById[agentId];
            this.setActiveTab(this.child('#warehouses'));
            if (tab) {
                this.objectTabs.setActiveTab(tab);
            }
        },

        showRegion: function (regionName) {
            var field = this.down('#search');
            this.setActiveTab(this.child('#overview'));
            if (field) {
                field.setValue(regionName);
            }
        }
    });

    // ---------------------------------------------------------------------
    // Left navigation: region -> object tree with status badges.
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.view.ObjectTree', {
        extend: 'Ext.tree.Panel',
        cls: 'bs-tree',
        rootVisible: false,
        useArrows: true,

        initComponent: function () {
            var me = this;

            me.store = Ext.create('Ext.data.TreeStore', {
                fields: ['name', 'agentId', 'badge', 'worst'],
                root: { expanded: true, children: [] }
            });

            me.columns = [
                { xtype: 'treecolumn', dataIndex: 'name', flex: 1, renderer: enc },
                {
                    dataIndex: 'badge', width: 70, align: 'right',
                    renderer: function (v, meta, r) {
                        return v ? '<span class="bs-tab-badge ' + (STATUS[r.get('worst')] || STATUS.nodata).cls + '">' + enc(v) + '</span>' : '';
                    }
                }
            ];
            me.hideHeaders = true;

            me.tools = [{
                xtype: 'button',
                iconCls: 'fa fa-rotate',
                tooltip: t('Fetch battery status'),
                handler: function () { me.module.reload(); }
            }];

            me.listeners = {
                itemclick: function (view, record) {
                    if (record.get('agentId')) {
                        me.mainPanel.showObject(record.get('agentId'));
                    } else if (record.get('name')) {
                        me.mainPanel.showRegion(record.get('name'));
                    }
                }
            };

            me.callParent(arguments);
            me.service.on('load', me.rebuild, me);
        },

        // Flat list of objects (no site/region grouping), sorted by name.
        rebuild: function (data) {
            var children = Ext.Array.map(data.getTagObjects(), function (obj) {
                var stats = data.getObjectStats(obj.agentId);
                return {
                    name: obj.name,
                    agentId: obj.agentId,
                    leaf: true,
                    iconCls: STATUS[stats.worst].icon + ' ' + STATUS[stats.worst].cls,
                    badge: stats.total ? stats.normal + '/' + stats.total : '-',
                    worst: stats.worst
                };
            }).sort(function (a, b) { return byName(a.name, b.name); });

            this.getStore().setRoot({ expanded: true, children: children });
        }
    });

    // ---------------------------------------------------------------------
    // Left navigation: settings.
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.view.Settings', {
        extend: 'Ext.form.Panel',
        cls: 'bs-settings',
        bodyPadding: 12,
        scrollable: true,

        initComponent: function () {
            var me = this;

            me.items = [
                {
                    xtype: 'combobox',
                    fieldLabel: t('Duration'),
                    labelAlign: 'top',
                    anchor: '100%',
                    editable: false,
                    queryMode: 'local',
                    store: Ext.Array.map(WINDOWS, function (w) { return [String(w[0]), t(w[1])]; }),
                    value: String(me.module.settings.window),
                    listeners: {
                        change: function (field, value) {
                            me.module.applySettings({ window: value === 'now' ? 'now' : Number(value) });
                        }
                    }
                },
                {
                    xtype: 'combobox',
                    fieldLabel: t('Auto refresh'),
                    labelAlign: 'top',
                    anchor: '100%',
                    editable: false,
                    queryMode: 'local',
                    store: [[0, t('Off')], [600, t('Every 10 minutes')], [1800, t('Every 30 minutes')], [3600, t('Every hour')]],
                    value: me.module.settings.refreshSec,
                    listeners: {
                        change: function (field, value) { me.module.applySettings({ refreshSec: value }); }
                    }
                },
                {
                    xtype: 'component',
                    cls: 'bs-muted',
                    html: '<p>' + enc(t('Now = the latest reading within the last 24 hours.')) + '</p>' +
                        '<p>' + enc(t('Status by battery voltage:')) + ' ≥ ' + LOW_VOLTAGE + ' V ' + enc(t('Normal')) + ', &lt; ' + LOW_VOLTAGE + ' V ' +
                        enc(t('Low')) + '. ' + enc(t('No voltage for more than')) + ' ' + STALE_HOURS + ' h = ' + enc(t('No reading')) + '.</p>' +
                        '<p>' + enc(t('Values are read from the raw packets (triggerdata) of Sensors Tracing. A sensor is listed when the object has a Temperature<id> or Humidity<id> sensor for it.')) + '</p>' +
                        '<p>' + enc(t('Settings are saved in this browser only.')) + '</p>' +
                        '<p>' + enc(t('Version')) + ' ' + enc(VERSION) + '</p>'
                }
            ];

            me.callParent(arguments);
        }
    });

    // =====================================================================
    // Vehicle Temperature Report (port of the site's app/page.tsx).
    // Data: /api/config, /api/vehicles, /api/geozones, /api/report (one call per vehicle).
    // =====================================================================
    function hasPlate(v) {
        return !!v.plate && v.plate !== v.id && v.plate.trim() !== '' && !/^\d{12,}$/.test(v.plate.trim());
    }
    function todayIn(tz) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
    }
    function fmtTime(ms, tz, withDate) {
        if (ms === null || ms === undefined) { return '—'; }
        var o = { timeZone: tz, hour: '2-digit', minute: '2-digit', hourCycle: 'h23' };
        if (withDate) { o.day = '2-digit'; o.month = 'short'; }
        return new Intl.DateTimeFormat('en-GB', o).format(new Date(ms));
    }
    function dayKey(ms, tz) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(ms));
    }
    function fmtTemp(v) {
        return v === null || v === undefined || !isFinite(v) ? '—' : Number(v).toFixed(1);
    }
    function dur(m) {
        return m === null || m === undefined ? '—' : Math.floor(m / 60) + 'h ' + Ext.String.leftPad(String(m % 60), 2, '0') + 'm';
    }

    /** Left panel: period, options and vehicle picker. */
    Ext.define('Store.battery_status.temp.Nav', {
        extend: 'Ext.panel.Panel',
        layout: { type: 'vbox', align: 'stretch' },

        initComponent: function () {
            var me = this;
            me.store = Ext.create('Ext.data.Store', { fields: ['id', 'plate', { name: 'sel', type: 'boolean' }, { name: 'plated', type: 'boolean' }] });
            me.items = [
                {
                    xtype: 'form', bodyPadding: '8 10', border: false, layout: 'anchor',
                    defaults: { anchor: '100%', labelAlign: 'top' },
                    items: [
                        { xtype: 'fieldcontainer', fieldLabel: t('From'), layout: 'hbox', items: [
                            { xtype: 'datefield', itemId: 'fromDate', format: 'Y-m-d', flex: 1 },
                            { xtype: 'timefield', itemId: 'fromTime', format: 'H:i', increment: 30, width: 90, margin: '0 0 0 4', value: '00:00' }
                        ] },
                        { xtype: 'fieldcontainer', fieldLabel: t('To'), layout: 'hbox', items: [
                            { xtype: 'datefield', itemId: 'toDate', format: 'Y-m-d', flex: 1 },
                            { xtype: 'timefield', itemId: 'toTime', format: 'H:i', increment: 30, width: 90, margin: '0 0 0 4', value: '23:59' }
                        ] },
                        { xtype: 'combobox', itemId: 'interval', fieldLabel: t('Sampling interval'), editable: false, queryMode: 'local', value: 10,
                            store: Ext.Array.map([5, 10, 30, 60], function (i) { return [i, i === 60 ? t('1 hour') : i + ' ' + t('minutes')]; }) },
                        { xtype: 'checkbox', itemId: 'ignition', boxLabel: t('Only readings with ignition ON'), checked: true },
                        { xtype: 'numberfield', itemId: 'minTrip', fieldLabel: t('Minimum trip duration (minutes)'), minValue: 0, maxValue: 240, value: 10, allowDecimals: false },
                        { xtype: 'container', layout: 'hbox', items: [
                            { xtype: 'button', itemId: 'run', text: t('Generate report'), iconCls: 'fa fa-play', flex: 1, handler: function () { me.main.run(); } },
                            { xtype: 'button', itemId: 'stop', text: t('Stop'), hidden: true, margin: '0 0 0 6', handler: function () { me.main.cancel = true; } }
                        ] },
                        { xtype: 'component', itemId: 'zoneWarn', hidden: true, margin: '6 0 0 0',
                            html: '<span style="color:#b3261e">' + enc(t('No geozones yet — add them in the Geozones tab.')) + '</span>' }
                    ]
                },
                {
                    xtype: 'grid', itemId: 'vehGrid', flex: 1, store: me.store, hideHeaders: true,
                    viewConfig: { emptyText: '<div class="bs-empty">' + enc(t('Loading vehicles…')) + '</div>', deferEmptyText: false, markDirty: false },
                    tbar: [
                        { xtype: 'textfield', emptyText: t('Search plate…'), flex: 1, listeners: { change: function () { me.applyFilter(); }, buffer: 200 }, itemId: 'search' }
                    ],
                    dockedItems: [{
                        xtype: 'toolbar', dock: 'top', items: [
                            { xtype: 'checkbox', itemId: 'noPlate', boxLabel: t('Show devices without plate'), hidden: true, listeners: { change: function () { me.applyFilter(); } } }
                        ]
                    }, {
                        xtype: 'toolbar', dock: 'bottom', items: [
                            { xtype: 'tbtext', itemId: 'count', text: '' }, '->',
                            { text: t('Select shown'), handler: function () { me.setShown(true); } },
                            { text: t('Clear shown'), handler: function () { me.setShown(false); } }
                        ]
                    }],
                    columns: [
                        { xtype: 'checkcolumn', dataIndex: 'sel', width: 34, listeners: { checkchange: function () { me.updateCount(); } } },
                        { dataIndex: 'plate', flex: 1, renderer: enc },
                        { dataIndex: 'id', width: 90, renderer: function (v, m, r) { return v !== r.get('plate') ? '<span class="bs-muted">' + enc(v) + '</span>' : ''; } }
                    ]
                }
            ];
            me.callParent(arguments);
        },

        setVehicles: function (list, ignitionMapped) {
            var plated = 0;
            this.store.loadData(Ext.Array.map(list, function (v) {
                var p = hasPlate(v);
                if (p) { plated += 1; }
                return { id: v.id, plate: v.plate, sel: p, plated: p };
            }));
            this.plated = plated;
            var np = this.down('#noPlate');
            np.setBoxLabel(t('Show devices without plate') + ' (' + (list.length - plated) + ')');
            np.setHidden(list.length <= plated);
            if (!ignitionMapped) {
                this.down('#ignition').setValue(false);
                this.down('#ignition').setDisabled(true);
            }
            this.setEmptyText(list.length ? '' : '<div class="bs-empty">' + enc(t('No vehicles.')) + '</div>');
            this.applyFilter();
        },

        setEmptyText: function (html) {
            var view = this.down('#vehGrid').getView();
            view.emptyText = html;
            if (view.rendered) { view.refresh(); }
        },

        /** Loading failed: say so in the list (instead of "Loading vehicles…") with a retry link. */
        showError: function (msg) {
            var me = this;
            me.setEmptyText('<div class="bs-error" style="margin:8px">' + enc(msg) + '<br><a href="#" class="bs-retry">' + enc(t('Try again')) + '</a></div>');
            var el = me.down('#vehGrid').getView().getEl();
            if (el && !me.retryBound) {
                me.retryBound = true;
                el.on('click', function (e) {
                    if (e.getTarget('.bs-retry')) {
                        e.preventDefault();
                        me.setEmptyText('<div class="bs-empty">' + enc(t('Loading vehicles…')) + '</div>');
                        me.main.ensureLoaded();
                    }
                });
            }
        },

        applyFilter: function () {
            var q = String(this.down('#search').getValue() || '').toLowerCase(),
                all = this.down('#noPlate').getValue();
            this.store.clearFilter();
            this.store.filterBy(function (r) {
                return (all || r.get('plated')) && (!q || r.get('plate').toLowerCase().indexOf(q) !== -1 || r.get('id').indexOf(q) !== -1);
            });
            this.updateCount();
        },

        setShown: function (on) {
            this.store.each(function (r) { r.set('sel', on); });
            this.updateCount();
        },

        selected: function () {
            var out = [];
            (this.store.getData().getSource() || this.store.getData()).each(function (r) {
                if (r.get('sel')) { out.push({ id: r.get('id'), plate: r.get('plate') }); }
            });
            return out;
        },

        updateCount: function () {
            this.down('#count').setText(this.selected().length + ' ' + t('selected') + ' · ' + (this.plated || 0) + ' ' + t('with plate number'));
        },

        setDefaults: function (tz) {
            var today = Ext.Date.parse(todayIn(tz), 'Y-m-d');
            this.down('#fromDate').setValue(today);
            this.down('#toDate').setValue(today);
        },

        params: function () {
            function join(d, tm) {
                if (!d) { return ''; }
                return Ext.Date.format(d, 'Y-m-d') + 'T' + (tm ? Ext.Date.format(tm, 'H:i') : '00:00');
            }
            return {
                from: join(this.down('#fromDate').getValue(), this.down('#fromTime').getValue()),
                to: join(this.down('#toDate').getValue(), this.down('#toTime').getValue()),
                interval: Number(this.down('#interval').getValue()) || 10,
                ignitionOnly: !!this.down('#ignition').getValue(),
                minTrip: Math.max(0, Number(this.down('#minTrip').getValue()) || 0)
            };
        },

        setRunning: function (running) {
            this.down('#run').setDisabled(running);
            this.down('#run').setText(running ? t('Generating…') : t('Generate report'));
            this.down('#stop').setHidden(!running);
        }
    });

    /** Main panel: notice, stats, daily table per vehicle, Excel. */
    Ext.define('Store.battery_status.temp.Main', {
        extend: 'Ext.panel.Panel',
        layout: 'fit',
        cls: 'bs-main bs-report',

        initComponent: function () {
            var me = this;
            me.results = [];
            me.errors = [];
            me.open = {};
            me.hideEmpty = true;
            me.hideNoSensor = true;
            me.resultSearch = '';
            me.body0 = Ext.create('Ext.Component', {
                scrollable: true, padding: 12,
                html: '<div class="bs-muted">' + enc(t('Pick a period and vehicles on the left, then press Generate report.')) + '</div>',
                listeners: {
                    click: {
                        element: 'el',
                        fn: function (e) {
                            var tr = e.getTarget('tr[data-key]'), btn = e.getTarget('[data-act]');
                            if (btn) {
                                var act = btn.getAttribute('data-act');
                                if (act === 'excel') { me.exportExcel(); }
                                if (act === 'signin') { me.loaded = false; me.ensureLoaded(); }
                                return;
                            }
                            if (tr) {
                                var k = tr.getAttribute('data-key');
                                if (me.open[k]) { delete me.open[k]; } else { me.open[k] = true; }
                                me.render();
                            }
                        }
                    }
                }
            });
            me.items = [me.body0];
            me.tbar = [
                { xtype: 'textfield', emptyText: t('Search plate in results…'), width: 210, listeners: { change: function (f, v) { me.resultSearch = String(v || '').toLowerCase(); me.render(); }, buffer: 200 } },
                { xtype: 'checkbox', boxLabel: t('Hide vehicles without trips'), checked: true, margin: '0 0 0 8', listeners: { change: function (f, v) { me.hideEmpty = v; me.render(); } } },
                { xtype: 'checkbox', itemId: 'hideNoSensor', boxLabel: t('Hide vehicles without temperature sensor'), checked: true, hidden: true, margin: '0 0 0 8',
                    listeners: { change: function (f, v) { me.hideNoSensor = v; me.render(); } } },
                '->',
                { text: t('Download Excel'), iconCls: 'fa fa-file-excel', itemId: 'excel', disabled: true, handler: function () { me.exportExcel(); } },
                { text: t('Print / PDF'), iconCls: 'fa fa-print', handler: function () { window.print(); } }
            ];
            me.callParent(arguments);
        },

        ensureLoaded: function () {
            var me = this;
            if (me.loaded) { return; }
            me.loaded = true;
            me.notice = null;
            me.loadErr = '';
            me.render();
            Api.call('/api/config').then(function (c) {
                me.settings = c.config.settings;
                me.nav.setDefaults(me.tz());
                return Api.call('/api/vehicles');
            }).then(function (v) {
                me.nav.setVehicles(v.vehicles || [], v.ignitionMapped);
                return Api.call('/api/geozones').then(function (g) {
                    me.zoneCount = Ext.Array.filter(g.zones || [], function (z) { return z.active !== false; }).length;
                }, function () { me.zoneCount = null; });
            }).then(function () {
                me.nav.down('#zoneWarn').setHidden(me.zoneCount !== 0);
                me.render();
            }).catch(function (e) {
                me.loaded = false;
                me.loadErr = errorText(e);
                me.nav.showError(me.loadErr);
                me.render();
            });
        },

        tz: function () { return (this.settings && this.settings.reportTimezone) || TZ; },

        run: function () {
            var me = this, nav = me.nav, p = nav.params(), list = nav.selected(), out, next = 0, done = 0, errs = [];
            if (!me.settings) { me.ensureLoaded(); return; }
            if (!p.from || !p.to || p.from >= p.to) { me.notice = { kind: 'err', text: t('Invalid time range: "From" must be before "To".') }; me.render(); return; }
            if (!list.length) { me.notice = { kind: 'warn', text: t('Select at least one vehicle.') }; me.render(); return; }
            if (me.zoneCount === 0) {
                me.notice = { kind: 'err', text: t('No geozones defined yet. Add the ASD sites in the Geozones tab (import a KML/GeoJSON file or draw them), then generate the report.') };
                me.render();
                return;
            }
            me.notice = null;
            me.cancel = false;
            me.running = true;
            me.results = [];
            me.errors = [];
            me.open = {};
            me.progress = { done: 0, total: list.length };
            nav.setRunning(true);
            me.render();
            out = new Array(list.length);

            function worker() {
                if (next >= list.length || me.cancel) { return Promise.resolve(); }
                var k = next++, v = list[k];
                return Api.call('/api/report', { json: { deviceId: v.id, plate: v.plate, from: p.from, to: p.to, intervalMin: p.interval, ignitionOnly: p.ignitionOnly, minTripMinutes: p.minTrip } })
                    .then(function (r) { out[k] = r; }, function (e) { errs.push({ plate: v.plate, error: errorText(e) }); })
                    .then(function () {
                        done += 1;
                        me.progress = { done: done, total: list.length };
                        me.results = Ext.Array.filter(out, Boolean);
                        me.render();
                        return worker();
                    });
            }

            Promise.all([worker(), worker(), worker()]).then(function () {
                var okList = Ext.Array.filter(out, Boolean),
                    withTrips = Ext.Array.filter(okList, function (r) { return r.days.length; }).length,
                    ns = Ext.Array.filter(okList, function (r) { return r.noSensor; }).length;
                me.errors = errs;
                me.ran = p;
                me.open = {};
                me.running = false;
                nav.setRunning(false);
                if (!okList.length) {
                    me.notice = { kind: 'err', text: t('Report failed') + ': ' + ((errs[0] && errs[0].error) || t('no data returned')) };
                } else if (!withTrips) {
                    me.notice = { kind: 'warn', text: t('Report finished — no completed geozone trips were found for') + ' ' + okList.length + ' ' + t('vehicle(s) in this period.') };
                } else {
                    me.notice = { kind: 'ok', text: t('Report ready') + ' ✓ — ' + withTrips + ' ' + t('vehicle(s) with trips') + ', ' + okList.length + ' ' + t('processed') +
                        (ns ? ', ' + ns + ' ' + t('without temperature sensor (excluded from averages)') : '') + (errs.length ? ', ' + errs.length + ' ' + t('failed') : '') + '. ' +
                        t('Use “Download Excel” to save it.'), excel: true };
                }
                me.render();
            });
        },

        visible: function () {
            var me = this;
            return Ext.Array.filter(me.results, function (r) {
                return (!me.hideEmpty || r.days.length) && !(me.hideNoSensor && r.noSensor) && (!me.resultSearch || r.plate.toLowerCase().indexOf(me.resultSearch) !== -1);
            });
        },

        totals: function (visible) {
            var sum = 0, n = 0, days = 0, longs = 0, readings = 0;
            Ext.Array.each(visible, function (r) {
                if (r.noSensor) { return; }
                sum += r.sum;
                n += r.samples;
                days += r.days.length;
                readings += r.readings || 0;
                longs += Ext.Array.filter(r.days, function (d) { return d.long; }).length;
            });
            return { avg: n ? sum / n : null, n: n, days: days, longs: longs, readings: readings,
                withTrips: Ext.Array.filter(visible, function (r) { return r.days.length && !r.noSensor; }).length };
        },

        // ---- one definition of the report rows, used by BOTH the table and the Excel export ----
        longLabel: function () { return t('Long trip') + ' > ' + ((this.settings && this.settings.maxTripHours) || 24) + 'h'; },
        zoneLabel: function (d) {
            var tr = d.trips[0];
            return !tr.entryZone || tr.entryZone === tr.exitZone ? tr.exitZone : tr.exitZone + ' → ' + tr.entryZone;
        },
        vehLabel: function (r) {
            return r.plate + (r.noSensor ? ' — ' + t('No temperature sensor') : '') + ' · ' + r.days.length + ' ' + t('day(s)') + ' · ' + t('vehicle average for the period');
        },
        dayCells: function (r, d) {
            var me = this, tr = d.trips[0], tz = me.tz();
            function twd(ms) { return ms === null || ms === undefined ? '—' : fmtTime(ms, tz, dayKey(ms, tz) !== d.day); }
            return {
                plate: r.plate, day: d.day, exit: twd(tr.exitTime), entry: twd(tr.entryTime), duration: dur(tr.durationMin),
                running: d.runningMin ? dur(d.runningMin) : '—', idle: d.idleMin ? dur(d.idleMin) : '—',
                stopped: d.stoppedMin ? dur(d.stoppedMin) : '—', inactive: d.inactiveMin ? dur(d.inactiveMin) : '—',
                zone: me.zoneLabel(d), avg: d.long ? '—' : fmtTemp(d.avg), min: fmtTemp(d.min), max: fmtTemp(d.max),
                readings: d.samples || (tr.readings || []).length || 0
            };
        },
        totalLabel: function (totals) {
            return t('Average temperature for the entire selected period (mean of') + ' ' + totals.n + ' ' + t('daily averages)');
        },
        COLS: ['Plate Number', 'Date', 'Geozone Exit', 'Geozone Entry', 'Duration', 'Running', 'Idle', 'Stopped', 'Inactive', 'Geozone', 'Avg °C', 'Min °C', 'Max °C', 'Readings'],

        render: function () {
            var me = this, html = [], visible = me.visible(), totals = me.totals(visible), tz = me.tz(),
                noSensor = Ext.Array.filter(me.results, function (r) { return r.noSensor; }),
                empty = Ext.Array.filter(me.results, function (r) { return !r.days.length; }),
                hns = me.down('#hideNoSensor');

            if (hns) { hns.setHidden(!noSensor.length); hns.setBoxLabel(t('Hide vehicles without temperature sensor') + ' (' + noSensor.length + ')'); }
            me.down('#excel').setDisabled(me.running || !Ext.Array.some(me.results, function (r) { return r.days.length; }));

            html.push('<h2 class="bs-h">' + enc(t('Vehicle Temperature Report')) + '</h2><div class="bs-muted bs-note">' +
                enc(t('Temperature counted from the moment a vehicle exits a geozone until it enters a geozone again. Trips longer than')) + ' ' +
                ((me.settings && me.settings.maxTripHours) || 24) + ' h ' + enc(t('are treated as long-distance and excluded.')) + '</div>');
            if (me.loadErr) {
                html.push('<div class="bs-error">' + enc(me.loadErr) + ' <button data-act="signin">' + enc(t('Try again')) + '</button></div>');
            }
            if (me.progress && me.progress.total) {
                html.push('<div class="bs-progress" style="margin-top:8px"><div style="width:' + Math.round(100 * me.progress.done / me.progress.total) + '%"></div></div>' +
                    '<div class="bs-muted bs-note">' + me.progress.done + ' / ' + me.progress.total + ' ' + enc(t('vehicles processed')) +
                    (me.running ? ' — ' + enc(t('first run of a period takes longer, later runs come from cache')) : '') + '</div>');
            }
            if (me.notice) {
                html.push('<div class="bs-notice bs-notice-' + me.notice.kind + '">' + enc(me.notice.text) +
                    (me.notice.excel && !me.running ? ' <button data-act="excel">⬇ ' + enc(t('Download Excel')) + '</button>' : '') + '</div>');
            }
            if (me.errors.length && me.results.length) {
                var uniq = Ext.Array.unique(Ext.Array.pluck(me.errors, 'error'));
                html.push('<div class="bs-error">' + (uniq.length === 1 ? me.errors.length + ' ' + enc(t('vehicle(s) failed')) + ': ' + enc(me.errors[0].error) :
                    Ext.Array.map(me.errors.slice(0, 5), function (e) { return enc(e.plate + ': ' + e.error); }).join('<br>') + (me.errors.length > 5 ? '<br>…' : '')) + '</div>');
            }

            if (me.results.length) {
                html.push('<div class="bs-tiles" style="margin:12px 0">' +
                    '<div class="bs-tile bs-tile-total"><i class="fa fa-temperature-half"></i><div class="bs-tile-value">' + fmtTemp(totals.avg) + ' °C</div><div class="bs-tile-label">' + enc(t('Period average temperature')) + '</div></div>' +
                    '<div class="bs-tile"><i class="fa fa-truck"></i><div class="bs-tile-value">' + totals.withTrips + ' / ' + me.results.length + '</div><div class="bs-tile-label">' + enc(t('Vehicles with trips')) + '</div></div>' +
                    '<div class="bs-tile"><i class="fa fa-calendar-day"></i><div class="bs-tile-value">' + totals.days + '</div><div class="bs-tile-label">' + enc(t('Vehicle-days')) + '</div></div>' +
                    '<div class="bs-tile"><i class="fa fa-list-ol"></i><div class="bs-tile-value">' + totals.readings.toLocaleString() + '</div><div class="bs-tile-label">' + enc(t('Readings used')) + '</div></div>' +
                    '<div class="bs-tile"><i class="fa fa-route"></i><div class="bs-tile-value">' + totals.longs + '</div><div class="bs-tile-label">' + enc(t('Long trips flagged')) + '</div></div>' +
                    '</div>');
                if (me.ran) {
                    html.push('<div class="bs-muted">' + enc(t('Daily temperature per vehicle')) + ' · ' + enc(me.ran.from.replace('T', ' ')) + ' → ' + enc(me.ran.to.replace('T', ' ')) + ' · ' +
                        enc(t('every')) + ' ' + (me.ran.interval === 60 ? '1 h' : me.ran.interval + ' min') + ' · ' + enc(me.ran.ignitionOnly ? t('ignition ON only') : t('all readings')) + '</div>');
                }
                html.push('<div class="bs-muted" style="margin:4px 0 6px">' + enc(t('Click a row to see the temperature readings used for that day.')) + '</div>');
                html.push('<table class="bs-table"><thead><tr>' + Ext.Array.map(me.COLS, function (h, i) {
                    return '<th' + (Ext.Array.contains([4, 5, 6, 7, 8, 10, 11, 12, 13], i) ? ' class="bs-num"' : '') + '>' + enc(t(h)) + '</th>';
                }).join('') + '</tr></thead><tbody>');
                Ext.Array.each(visible, function (r) {
                    html.push('<tr class="bs-veh"><td colspan="10">' + enc(me.vehLabel(r)) + '</td><td class="bs-num"><b>' + fmtTemp(r.avg) + '</b></td><td colspan="2"></td><td class="bs-num">' + (r.readings || 0) + '</td></tr>');
                    Ext.Array.each(r.days, function (d) {
                        var c = me.dayCells(r, d), key = r.deviceId + '|' + d.day, rd = d.trips[0].readings || [], isOpen = !!me.open[key];
                        html.push('<tr' + (rd.length ? ' data-key="' + enc(key) + '" style="cursor:pointer"' : '') + (d.long ? ' class="bs-long"' : '') + '>' +
                            '<td>' + enc(c.plate) + '</td><td>' + enc(c.day) + '</td><td>' + enc(c.exit) + '</td><td>' + enc(c.entry) + '</td>' +
                            '<td class="bs-num">' + c.duration + '</td><td class="bs-num">' + c.running + '</td><td class="bs-num">' + c.idle + '</td><td class="bs-num">' + c.stopped + '</td><td class="bs-num">' + c.inactive + '</td>' +
                            '<td>' + enc(c.zone) + (d.long ? ' <span class="bs-pill bs-nodata">' + enc(me.longLabel()) + '</span>' : '') + '</td>' +
                            '<td class="bs-num"><b>' + c.avg + '</b></td><td class="bs-num">' + c.min + '</td><td class="bs-num">' + c.max + '</td>' +
                            '<td class="bs-num">' + (rd.length ? '<span class="bs-link">' + (isOpen ? '▾' : '▸') + ' ' + c.readings + '</span>' : c.readings) + '</td></tr>');
                        if (isOpen) {
                            html.push('<tr class="bs-sub"><td colspan="14"><div class="bs-readings">' + Ext.Array.map(rd, function (x) {
                                return '<span class="bs-reading"><b>' + x[1].toFixed(1) + '°</b> ' + enc(fmtTime(x[0], tz)) + '</span>';
                            }).join('') + '</div></td></tr>');
                        }
                    });
                });
                html.push('</tbody><tfoot><tr><td colspan="10">' + enc(me.totalLabel(totals)) + '</td><td class="bs-num"><b>' + fmtTemp(totals.avg) + '</b></td><td colspan="2"></td><td class="bs-num">' + totals.readings + '</td></tr></tfoot></table>');
                if (me.hideEmpty && empty.length) {
                    html.push('<div class="bs-muted bs-note">' + empty.length + ' ' + enc(t('vehicle(s) had no completed trips in this period')) + ': ' +
                        enc(Ext.Array.pluck(empty.slice(0, 30), 'plate').join(', ')) + (empty.length > 30 ? '…' : '') + '</div>');
                }
            }
            me.body0.setHtml(html.join(''));
        },

        exportExcel: function () {
            var me = this, ran = me.ran, tz = me.tz(), visible = me.visible(), totals = me.totals(visible);
            if (!ran) { return; }
            me.setLoading(t('Building Excel…'));
            ensureExcel().then(function (ExcelJS) {
                var wb = new ExcelJS.Workbook();
                wb.creator = 'Temperature Report';
                function head(ws, title) {
                    ws.addRow([title]).font = { bold: true, size: 14 };
                    ws.addRow(['Period: ' + ran.from.replace('T', ' ') + ' to ' + ran.to.replace('T', ' ') + '   ·   Sampling: every ' + ran.interval + ' min   ·   Ignition: ' +
                        (ran.ignitionOnly ? 'ON only' : 'all readings') + '   ·   Trips shorter than ' + ran.minTrip + ' min ignored']);
                    ws.addRow(['Generated: ' + new Date().toLocaleString('en-GB', { timeZone: tz }) + ' (' + tz + ')']);
                    ws.addRow([]);
                }
                // temperature strings like "10.4" become real numbers shown with one decimal — identical to the screen
                function xv(s) { return typeof s === 'string' && /^-?\d+(\.\d+)?$/.test(s) ? +s : s; }
                function tempFmt(row, cols) { Ext.Array.each(cols, function (i) { row.getCell(i).numFmt = '0.0'; }); }

                // Sheet 1: exactly the table on screen
                var ws = wb.addWorksheet('Daily Detail', { views: [{ state: 'frozen', ySplit: 5 }] });
                head(ws, 'Al Safi — Vehicle Temperature Report');
                xlHead(ws.addRow(me.COLS));
                Ext.Array.each(visible, function (r) {
                    var vr = ws.addRow([me.vehLabel(r), '', '', '', '', '', '', '', '', '', xv(fmtTemp(r.avg)), '', '', r.readings || 0]);
                    ws.mergeCells(vr.number, 1, vr.number, 10);
                    vr.font = { bold: true };
                    xlFill(vr, 'FFEEF3F6');
                    tempFmt(vr, [11]);
                    Ext.Array.each(r.days, function (d) {
                        var c = me.dayCells(r, d),
                            row = ws.addRow([c.plate, c.day, c.exit, c.entry, c.duration, c.running, c.idle, c.stopped, c.inactive,
                                d.long ? c.zone + '  [' + me.longLabel() + ']' : c.zone, xv(c.avg), xv(c.min), xv(c.max), c.readings]);
                        tempFmt(row, [11, 12, 13]);
                        if (d.long) { xlFill(row, 'FFFFF4E5'); }
                    });
                });
                var ft = ws.addRow([me.totalLabel(totals), '', '', '', '', '', '', '', '', '', xv(fmtTemp(totals.avg)), '', '', totals.readings]);
                ws.mergeCells(ft.number, 1, ft.number, 10);
                ft.font = { bold: true };
                xlFill(ft, 'FFE2F3F2');
                tempFmt(ft, [11]);
                ws.columns.forEach(function (c, i) { c.width = [16, 12, 15, 15, 11, 11, 11, 11, 11, 38, 9, 9, 9, 10][i] || 14; });
                Ext.Array.each([5, 6, 7, 8, 9, 11, 12, 13, 14], function (i) { ws.getColumn(i).alignment = { horizontal: 'right' }; });

                // Sheet 2: averages
                var wa = wb.addWorksheet('Averages', { views: [{ state: 'frozen', ySplit: 5 }] });
                head(wa, 'Al Safi — Average temperature per vehicle (mean of daily averages)');
                xlHead(wa.addRow(['Plate Number', 'Days', 'Average Temp (°C)', 'Lowest daily avg (°C)', 'Highest daily avg (°C)', 'Long trips flagged', 'Readings']));
                Ext.Array.each(visible, function (r) {
                    var avgs = Ext.Array.filter(Ext.Array.pluck(r.days, 'avg'), function (v) { return v !== null && v !== undefined; }),
                        row = wa.addRow([r.plate + (r.noSensor ? ' (no temperature sensor)' : ''), r.days.length, xv(fmtTemp(r.avg)),
                            avgs.length ? xv(fmtTemp(Math.min.apply(Math, avgs))) : '—', avgs.length ? xv(fmtTemp(Math.max.apply(Math, avgs))) : '—',
                            Ext.Array.filter(r.days, function (d) { return d.long; }).length, r.readings || 0]);
                    tempFmt(row, [3, 4, 5]);
                });
                wa.addRow([]);
                var tot = wa.addRow(['Average temperature for the entire selected period', totals.n, xv(fmtTemp(totals.avg)), '', '', totals.longs, totals.readings]);
                tempFmt(tot, [3]);
                tot.font = { bold: true };
                tot.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2F3F2' } }; });
                wa.columns.forEach(function (c, i) { c.width = [44, 8, 18, 20, 20, 18, 10][i] || 14; });

                // Sheet 3: the readings shown when a row is clicked
                var wr = wb.addWorksheet('Readings');
                head(wr, 'Temperature readings used in the averages');
                xlHead(wr.addRow(['Plate Number', 'Date', 'Time', 'Temp (°C)']));
                Ext.Array.each(visible, function (r) {
                    Ext.Array.each(r.days, function (d) {
                        Ext.Array.each(d.trips[0].readings || [], function (x) {
                            tempFmt(wr.addRow([r.plate, d.day, fmtTime(x[0], tz), xv(x[1].toFixed(1))]), [4]);
                        });
                    });
                });
                wr.columns.forEach(function (c) { c.width = 16; });
                return wb.xlsx.writeBuffer();
            }).then(function (buf) {
                me.setLoading(false);
                saveFile(buf, 'temperature-report_' + ran.from.replace(/[T:]/g, '-') + '_' + ran.to.replace(/[T:]/g, '-') + '.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            }).catch(function (e) {
                me.setLoading(false);
                Ext.Msg.alert(t('Download Excel'), enc(errorText(e)));
            });
        }
    });

    // =====================================================================
    // Geozones editor (port of the site's app/geozones/page.tsx + GeozoneMap.tsx).
    // Client sites used by the Temperature Report, stored by the backend (/api/geozones).
    // =====================================================================
    function gzUid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }
    function makeZone(name, g, source) { return Ext.apply({ id: gzUid(), name: name, active: true, source: source }, g); }

    /** GeoJSON FeatureCollection -> zones (KML/GeoJSON uploads). */
    function zonesFromGeoJSON(fc, defaultRadius) {
        var out = [];
        defaultRadius = defaultRadius || 200;
        Ext.Array.each((fc && fc.features) || [], function (f) {
            var g = f.geometry, props = f.properties || {}, ring, c, r,
                name = String(props.name || props.Name || props.title || props.NAME || ('Zone ' + (out.length + 1)));
            if (!g) { return; }
            if (g.type === 'Polygon' || g.type === 'MultiPolygon') {
                ring = g.type === 'Polygon' ? g.coordinates[0] : g.coordinates[0][0];
                out.push(makeZone(name, { type: 'polygon', points: Ext.Array.map(ring, function (p) { return [p[1], p[0]]; }) }, 'file'));
            } else if (g.type === 'Point') {
                c = g.coordinates;
                r = Number(props.radius || props.Radius || defaultRadius) || defaultRadius;
                out.push(makeZone(name, { type: 'circle', lat: c[1], lng: c[0], radius: r }, 'file'));
            }
        });
        return out;
    }

    function zonesToGeoJSON(zones) {
        return {
            type: 'FeatureCollection',
            features: Ext.Array.map(zones, function (z) {
                return z.type === 'circle' ?
                    { type: 'Feature', properties: { name: z.name, radius: z.radius, active: z.active !== false }, geometry: { type: 'Point', coordinates: [z.lng, z.lat] } } :
                    { type: 'Feature', properties: { name: z.name, active: z.active !== false },
                        geometry: { type: 'Polygon', coordinates: [Ext.Array.map(z.points.concat([z.points[0]]), function (p) { return [p[1], p[0]]; })] } };
            })
        };
    }

    function ensureLeaflet() {
        loadCss(LIB.geomanCss);
        var leaflet = window.L ? Promise.resolve() : (loadCss(LIB.leafletCss), loadScript(LIB.leafletJs));
        return leaflet.then(function () {
            return window.L.PM ? null : loadScript(LIB.geomanJs);
        }).then(function () {
            if (!window.L || !window.L.PM) { throw new Error('Map editor did not load.'); }
            return window.L;
        });
    }

    /** Shared editor state used by the left panel and the map. */
    Ext.define('Store.battery_status.geo.State', {
        mixins: ['Ext.mixin.Observable'],
        constructor: function (config) {
            this.mixins.observable.constructor.call(this, config);
            this.zones = [];
            this.preview = [];
            this.sel = null;
            this.dirty = false;
        },
        update: function (zones) { this.zones = zones; this.dirty = true; this.fireEvent('change', this); },
        replace: function (zones) { this.zones = zones; this.dirty = false; this.fireEvent('change', this); },
        select: function (id, focus) { this.sel = id; this.fireEvent('change', this); if (focus) { this.fireEvent('focus', this, id); } },
        setPreview: function (p) { this.preview = p; this.fireEvent('preview', this); },
        selected: function () { var id = this.sel; return Ext.Array.findBy(this.zones, function (z) { return z.id === id; }); },
        patch: function (id, fn) { this.update(Ext.Array.map(this.zones, function (z) { return z.id === id ? fn(Ext.apply({}, z)) : z; })); },
        remove: function (id) {
            this.zones = Ext.Array.filter(this.zones, function (z) { return z.id !== id; });
            if (this.sel === id) { this.sel = null; }
            this.dirty = true;
            this.fireEvent('change', this);
        }
    });

    /** Left panel: save/import/export, selected zone, list. */
    Ext.define('Store.battery_status.geo.Nav', {
        extend: 'Ext.panel.Panel',
        layout: { type: 'vbox', align: 'stretch' },

        initComponent: function () {
            var me = this, st = me.gstate;
            me.store = Ext.create('Ext.data.Store', { fields: ['id', 'name', 'type', 'radius', 'active', 'color'], sorters: [{ property: 'name', direction: 'ASC' }] });
            me.items = [
                {
                    xtype: 'container', padding: '8 10', layout: 'anchor',
                    items: [
                        { xtype: 'container', layout: 'hbox', items: [
                            { xtype: 'button', itemId: 'save', text: t('Saved'), disabled: true, iconCls: 'fa fa-floppy-disk', handler: function () { me.save(); } },
                            { xtype: 'button', text: t('Import…'), margin: '0 0 0 6', handler: function () { me.openImport(); } },
                            { xtype: 'button', text: t('Export'), margin: '0 0 0 6', handler: function () { me.exportGeo(); } }
                        ] },
                        { xtype: 'component', itemId: 'hint', cls: 'bs-muted', margin: '6 0 0 0',
                            html: enc(t('Draw with the tools on the map (circle / polygon / rectangle), edit with the pencil, move with the hand, delete with the bin.')) },
                        { xtype: 'component', itemId: 'msg', margin: '6 0 0 0', html: '' }
                    ]
                },
                {
                    xtype: 'form', itemId: 'selForm', hidden: true, bodyPadding: '6 10', title: t('Selected geozone'), layout: 'anchor',
                    defaults: { anchor: '100%', labelAlign: 'top' },
                    items: [
                        { xtype: 'textfield', itemId: 'name', fieldLabel: t('Name'), listeners: { change: function (f, v) { if (!me.syncing && st.sel) { st.patch(st.sel, function (z) { z.name = v; return z; }); } } } },
                        { xtype: 'numberfield', itemId: 'radius', fieldLabel: t('Radius (m)'), minValue: 10,
                            listeners: { change: function (f, v) { if (!me.syncing && st.sel) { st.patch(st.sel, function (z) { if (z.type === 'circle') { z.radius = Math.max(10, Number(v) || 10); } return z; }); } } } },
                        { xtype: 'checkbox', itemId: 'active', boxLabel: t('Active (used in report)'),
                            listeners: { change: function (f, v) { if (!me.syncing && st.sel) { st.patch(st.sel, function (z) { z.active = v; return z; }); } } } },
                        { xtype: 'component', itemId: 'info', cls: 'bs-muted' },
                        { xtype: 'container', layout: 'hbox', margin: '6 0 0 0', items: [
                            { xtype: 'button', text: t('Delete'), iconCls: 'fa fa-trash', handler: function () { if (st.sel) { st.remove(st.sel); } } },
                            { xtype: 'button', text: t('Close'), margin: '0 0 0 6', handler: function () { st.select(null); } }
                        ] }
                    ]
                },
                {
                    xtype: 'grid', itemId: 'list', flex: 1, store: me.store, hideHeaders: true, title: '0 ' + t('geozones'),
                    viewConfig: { emptyText: '<div class="bs-empty">' + enc(t('No geozones yet — import the ASD sites or draw them on the map.')) + '</div>', deferEmptyText: false },
                    tbar: [{ xtype: 'textfield', emptyText: t('Search…'), flex: 1, listeners: { change: function (f, v) {
                        var q = String(v || '').toLowerCase();
                        me.store.clearFilter();
                        if (q) { me.store.filterBy(function (r) { return r.get('name').toLowerCase().indexOf(q) !== -1; }); }
                    }, buffer: 200 } }],
                    columns: [
                        { dataIndex: 'name', flex: 1, renderer: function (v, m, r) {
                            return '<span class="bs-sw" style="background:' + (r.get('active') === false ? '#c4ccd4' : (r.get('color') || '#0b7a75')) + '"></span> ' + enc(v);
                        } },
                        { dataIndex: 'radius', width: 70, align: 'right', renderer: function (v, m, r) { return '<span class="bs-muted">' + (r.get('type') === 'circle' ? v + ' m' : 'poly') + '</span>'; } }
                    ],
                    listeners: { itemclick: function (v, r) { st.select(r.get('id'), true); } }
                }
            ];
            me.callParent(arguments);
            st.on('change', me.sync, me);
        },

        message: function (kind, text) {
            this.down('#msg').setHtml(text ? '<div class="bs-notice bs-notice-' + kind + '">' + enc(text) + '</div>' : '');
        },

        sync: function () {
            var me = this, st = me.gstate, z = st.selected(), save = me.down('#save'), form = me.down('#selForm'), list = me.down('#list');
            save.setDisabled(!st.dirty || me.busy);
            save.setText(st.dirty ? t('Save changes') : t('Saved'));
            me.store.loadData(Ext.Array.map(st.zones, function (x) { return { id: x.id, name: x.name, type: x.type, radius: x.radius, active: x.active, color: x.color }; }));
            list.setTitle(st.zones.length + ' ' + t('geozones'));
            form.setHidden(!z);
            if (z) {
                me.syncing = true;
                if (me.down('#name').getValue() !== z.name) { me.down('#name').setValue(z.name); }
                me.down('#radius').setHidden(z.type !== 'circle');
                if (z.type === 'circle' && me.down('#radius').getValue() !== z.radius) { me.down('#radius').setValue(z.radius); }
                me.down('#active').setValue(z.active !== false);
                me.down('#info').setHtml(enc(z.type === 'circle' ? 'Circle · ' + z.lat.toFixed(5) + ', ' + z.lng.toFixed(5) : 'Polygon · ' + z.points.length + ' points') +
                    (z.source ? ' · ' + enc(t('from')) + ' ' + enc(z.source) : ''));
                me.syncing = false;
                var rec = me.store.getById ? me.store.findRecord('id', z.id, 0, false, true, true) : null;
                if (rec) { list.getSelectionModel().select(rec, false, true); }
            } else {
                list.getSelectionModel().deselectAll(true);
            }
        },

        save: function () {
            var me = this, st = me.gstate;
            me.busy = true;
            me.sync();
            Api.call('/api/geozones', { method: 'PUT', json: { zones: st.zones } }).then(function (r) {
                me.busy = false;
                st.replace(r.zones);
                me.message('ok', t('Saved') + ' ' + r.zones.length + ' ' + t('geozones') + '.');
            }, function (e) {
                me.busy = false;
                me.sync();
                me.message('err', errorText(e));
            });
        },

        exportGeo: function () {
            if (!this.gstate.zones.length) { return; }
            saveFile(JSON.stringify(zonesToGeoJSON(this.gstate.zones), null, 1), 'geozones.geojson', 'application/geo+json');
        },

        mergeIn: function (newZones) {
            var st = this.gstate, byName = {}, order = [], added = 0, replaced = 0;
            Ext.Array.each(st.zones, function (z) { var k = z.name.trim().toLowerCase(); byName[k] = z; order.push(k); });
            Ext.Array.each(newZones, function (z) {
                var k = z.name.trim().toLowerCase(), old = byName[k];
                if (old) { byName[k] = Ext.apply({}, { id: old.id, color: old.color, active: old.active }, z); replaced += 1; }
                else { byName[k] = z; order.push(k); added += 1; }
            });
            st.update(Ext.Array.map(order, function (k) { return byName[k]; }));
            st.setPreview([]);
            this.message('warn', added + ' ' + t('added') + ', ' + replaced + ' ' + t('updated') + '. ' + t("Don't forget to click Save."));
        },

        openImport: function () {
            var me = this, st = me.gstate, tables = [], cands = [], imp = { table: '', nameCol: '', areaCol: '', areaType: '', latCol: '', lngCol: '', radiusCol: '', prefix: 'ASD', order: 'auto' },
                win, form, addBtn, msg;

            function cols() { var tb = Ext.Array.findBy(tables, function (x) { return x.name === imp.table; }); return tb ? tb.columns : []; }
            function colStore() { return [['', '—']].concat(Ext.Array.map(cols(), function (c) { return [c.name, c.name + ' (' + c.type + ')']; })); }
            function colSel(key, label) {
                return { xtype: 'combobox', fieldLabel: t(label), editable: false, queryMode: 'local', value: imp[key] || '', store: colStore(),
                    listeners: { select: function (f) {
                        imp[key] = f.getValue() || null;
                        if (key === 'areaCol') { var c = Ext.Array.findBy(cols(), function (x) { return x.name === f.getValue(); }); imp.areaType = c ? c.type : null; }
                    } } };
            }
            function refresh() {
                var items = [
                    { xtype: 'component', cls: 'bs-muted', html: enc(t('From the platform database (sites whose name starts with the prefix) or from a KML / GeoJSON file exported from the platform.')) },
                    { xtype: 'combobox', fieldLabel: t('Table'), editable: false, queryMode: 'local', value: imp.table,
                        store: [['', '— ' + t('choose') + ' —']].concat(Ext.Array.map(cands, function (c) { return [c.table, c.table + ' (' + t('suggested') + ')']; }))
                            .concat(Ext.Array.map(Ext.Array.filter(tables, function (tb) { return !Ext.Array.some(cands, function (c) { return c.table === tb.name; }); }), function (tb) { return [tb.name, tb.name]; })),
                        listeners: { select: function (f) {
                            var c = Ext.Array.findBy(cands, function (x) { return x.table === f.getValue(); });
                            Ext.apply(imp, c || { nameCol: '', areaCol: '', latCol: '', lngCol: '', radiusCol: '' }, {});
                            imp.table = f.getValue();
                            refresh();
                        } } }
                ];
                if (!cands.length && tables.length) {
                    items.push({ xtype: 'component', cls: 'bs-muted', html: enc(t('No geofence-like table detected automatically — pick one, or use a file.')) });
                }
                if (imp.table) {
                    items.push(colSel('nameCol', 'Name column'), colSel('areaCol', 'Geometry column (WKT / JSON / points)'),
                        colSel('latCol', 'Center lat'), colSel('lngCol', 'Center lng'), colSel('radiusCol', 'Radius column (m)'),
                        { xtype: 'textfield', fieldLabel: t('Name prefix'), value: imp.prefix, listeners: { change: function (f, v) { imp.prefix = v; } } },
                        { xtype: 'combobox', fieldLabel: t('Coordinate order'), editable: false, queryMode: 'local', value: imp.order,
                            store: [['auto', t('Auto')], ['latlng', 'lat, lng'], ['lnglat', 'lng, lat']], listeners: { select: function (f) { imp.order = f.getValue(); } } },
                        { xtype: 'button', text: t('Preview from database'), anchor: '50%', handler: preview });
                }
                items.push({ xtype: 'filefield', fieldLabel: t('…or upload KML / GeoJSON'), buttonText: t('Choose file…'),
                    listeners: { afterrender: function (f) { f.fileInputEl.dom.setAttribute('accept', '.kml,.geojson,.json'); }, change: function (f) { var file = f.fileInputEl.dom.files[0]; if (file) { onFile(file); } } } });
                form.removeAll();
                form.add(items);
            }
            function preview() {
                if (!imp.nameCol) { return; }
                win.setLoading(true);
                Api.call('/api/geozones/import-db', { json: imp }).then(function (r) {
                    win.setLoading(false);
                    st.setPreview(r.zones);
                    showResult(r.zones.length ? 'ok' : 'warn', t('Found') + ' ' + r.total + ' ' + t('rows') + ', ' + r.zones.length + ' ' + t('geozones parsed') +
                        (r.failed.length ? ', ' + r.failed.length + ' ' + t('could not be parsed') + ' (' + r.failed.slice(0, 5).join(', ') + ')' : '') + '. ' +
                        t('Blue dashed shapes on the map are the preview — check they sit on the real sites, then click “Add to my geozones”.'));
                }, function (e) { win.setLoading(false); showResult('err', errorText(e)); });
            }
            function onFile(file) {
                file.text().then(function (text) {
                    var isKml = /\.kml$/i.test(file.name) || text.trim().charAt(0) === '<';
                    return (isKml ? loadScript(LIB.togeojson).then(function () {
                        return window.toGeoJSON.kml(new DOMParser().parseFromString(text, 'text/xml'));
                    }) : Promise.resolve(JSON.parse(text))).then(function (fc) {
                        var z = zonesFromGeoJSON(fc);
                        if (!z.length) { throw new Error(t('No polygons or points found in the file.')); }
                        st.setPreview(z);
                        showResult('ok', z.length + ' ' + t('geozones read from') + ' ' + file.name + '. ' + t('Check the preview, then click “Add to my geozones”.'));
                    });
                }).catch(function (e) { showResult('err', t('Could not read file') + ': ' + errorText(e)); });
            }
            function showResult(kind, text) {
                msg.setHtml('<div class="bs-notice bs-notice-' + kind + '">' + enc(text) + '</div>');
                addBtn.setText(t('Add') + ' ' + st.preview.length + ' ' + t('to my geozones'));
                addBtn.setHidden(!st.preview.length);
            }

            form = Ext.create('Ext.form.Panel', { border: false, layout: 'anchor', defaults: { anchor: '100%', labelAlign: 'top' }, bodyPadding: 0 });
            msg = Ext.create('Ext.Component', { html: '' });
            addBtn = Ext.create('Ext.button.Button', { hidden: true, text: '', handler: function () { me.mergeIn(st.preview); win.close(); } });
            win = Ext.create('Ext.window.Window', {
                title: t('Import geozones'), width: 420, maxHeight: Ext.getBody().getViewSize().height - 40, scrollable: true, bodyPadding: 12, layout: 'anchor',
                items: [form, msg], buttons: [{ text: t('Close'), handler: function () { win.close(); } }, addBtn],
                listeners: { close: function () { if (!addBtn.isHidden()) { st.setPreview([]); } } }
            });
            win.show();
            refresh();
            win.setLoading(t('Reading database tables…'));
            Api.call('/api/schema').then(function (s) {
                win.setLoading(false);
                if (!s.ok) { throw new Error(s.error); }
                tables = s.tables || [];
                cands = s.geofences || [];
                if (cands[0]) { Ext.apply(imp, cands[0]); }
                refresh();
            }).catch(function (e) { win.setLoading(false); showResult('err', errorText(e)); refresh(); });
        }
    });

    /** Main panel: Leaflet map with Geoman drawing tools. */
    Ext.define('Store.battery_status.geo.Map', {
        extend: 'Ext.panel.Panel',
        cls: 'bs-main bs-geomap',
        layout: 'fit',

        initComponent: function () {
            var me = this;
            me.html = '<div class="bs-mapdiv" style="width:100%;height:100%"></div>';
            me.callParent(arguments);
            me.on('resize', function () { if (me.map) { me.map.invalidateSize(); } });
            me.gstate.on('change', me.draw, me);
            me.gstate.on('preview', me.drawPreview, me);
            me.gstate.on('focus', me.focusZone, me);
        },

        ensureLoaded: function () {
            var me = this, st = me.gstate;
            if (me.loaded) { return; }
            me.loaded = true;
            me.setLoading(t('Loading map…'));
            ensureLeaflet().then(function () {
                me.setLoading(false);
                if (me.rendered) { me.initMap(); } else { me.on('afterrender', me.initMap, me, { single: true }); }
                return Api.call('/api/geozones');
            }).then(function (r) {
                st.replace(r.zones || []);
                if (r.storage) { me.nav.down('#hint').setHtml(me.nav.down('#hint').initialConfig.html + ' ' + enc(t('Storage')) + ': <b>' + enc(r.storage) + '</b>'); }
            }).catch(function (e) {
                me.loaded = false;
                me.setLoading(false);
                me.nav.message('err', errorText(e));
            });
        },

        geom: function (layer) {
            var L = window.L, c, ll, ring;
            if (layer instanceof L.Circle) {
                c = layer.getLatLng();
                return { type: 'circle', lat: +c.lat.toFixed(7), lng: +c.lng.toFixed(7), radius: Math.round(layer.getRadius()) };
            }
            if (layer instanceof L.Polygon) {
                ll = layer.getLatLngs();
                ring = Ext.isArray(ll[0]) ? ll[0] : ll;
                return { type: 'polygon', points: Ext.Array.map(ring, function (p) { return [+p.lat.toFixed(7), +p.lng.toFixed(7)]; }) };
            }
            return null;
        },

        initMap: function () {
            var me = this, L = window.L, el = me.body.dom.querySelector('.bs-mapdiv'), m, streets, sat;
            if (me.map) {
                // Map already built (e.g. loading the zones is being retried): just refresh its size.
                me.map.invalidateSize();
                return;
            }
            m = L.map(el, { center: [24.7136, 46.6753], zoom: 6 });
            streets = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© OpenStreetMap' });
            sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 19, attribution: 'Esri World Imagery' });
            streets.addTo(m);
            L.control.layers({ Streets: streets, Satellite: sat }, {}, { position: 'topright' }).addTo(m);
            me.group = L.featureGroup().addTo(m);
            me.prevGroup = L.layerGroup().addTo(m);
            me.layers = {};
            m.pm.addControls({
                position: 'topleft', drawMarker: false, drawCircleMarker: false, drawPolyline: false, drawText: false,
                cutPolygon: false, rotateMode: false, drawRectangle: true, drawPolygon: true, drawCircle: true,
                editMode: true, dragMode: true, removalMode: true
            });
            m.pm.setGlobalOptions({ snappable: true, pathOptions: { color: '#0b7a75' } });
            m.on('pm:create', function (e) {
                var g = me.geom(e.layer), st = me.gstate, z;
                m.removeLayer(e.layer);
                if (!g) { return; }
                z = makeZone('ASD New Site ' + (st.zones.length + 1), g, 'manual');
                st.update(st.zones.concat([z]));
                st.select(z.id);
                me.nav.message('warn', t('New geozone added — rename it in the panel and click Save.'));
            });
            me.map = m;
            Ext.defer(function () { m.invalidateSize(); }, 200);
            me.draw();
        },

        draw: function () {
            var me = this, L = window.L, m = me.map, g = me.group, st = me.gstate;
            if (!m || !g) { return; }
            g.clearLayers();
            me.layers = {};
            Ext.Array.each(st.zones, function (z) {
                var sel = z.id === st.sel,
                    color = z.active === false ? '#8796a5' : (z.color || '#0b7a75'),
                    style = { color: sel ? '#d9480f' : color, weight: sel ? 3 : 2, fillOpacity: sel ? 0.28 : 0.15, dashArray: z.active === false ? '6 6' : null },
                    layer = z.type === 'circle' ? L.circle([z.lat, z.lng], Ext.apply({ radius: z.radius }, style)) : L.polygon(z.points, style);
                function changed() { var ng = me.geom(layer); if (ng) { st.patch(z.id, function (x) { return Ext.apply(x, ng); }); } }
                layer.bindTooltip(z.name, { sticky: true });
                layer.on('click', function () { st.select(z.id); });
                layer.on('pm:edit', changed);
                layer.on('pm:dragend', changed);
                layer.on('pm:remove', function () { st.remove(z.id); });
                layer.addTo(g);
                me.layers[z.id] = layer;
            });
            if (!me.fitted && st.zones.length) {
                me.fitted = true;
                var b = g.getBounds();
                if (b.isValid()) { m.fitBounds(b.pad(0.1)); }
            }
        },

        drawPreview: function () {
            var me = this, L = window.L, pg = me.prevGroup, m = me.map, fg, b;
            if (!pg || !m) { return; }
            pg.clearLayers();
            if (!me.gstate.preview.length) { return; }
            fg = L.featureGroup();
            Ext.Array.each(me.gstate.preview, function (z) {
                var s = { color: '#1f6fd1', weight: 2, dashArray: '4 4', fillOpacity: 0.1, pmIgnore: true },
                    l = z.type === 'circle' ? L.circle([z.lat, z.lng], Ext.apply({ radius: z.radius }, s)) : L.polygon(z.points, s);
                l.bindTooltip(z.name);
                fg.addLayer(l);
            });
            fg.addTo(pg);
            b = fg.getBounds();
            if (b.isValid()) { m.fitBounds(b.pad(0.1)); }
        },

        focusZone: function (st, id) {
            var l = this.layers && this.layers[id];
            if (l && this.map) { this.map.fitBounds(l.getBounds().pad(0.6), { maxZoom: 16 }); }
        }
    });

    // =====================================================================
    // Vehicles Battery Status (port of the site's app/vehicles-battery/page.tsx).
    // Latest packet of every vehicle from today's log on gps.im2m.ws (via the backend).
    // =====================================================================
    function fmtMs(ms) {
        if (!ms) { return '—'; }
        return new Intl.DateTimeFormat('en-GB', { timeZone: TZ, day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).format(new Date(ms));
    }
    function agoMs(ms) { return ms ? ago(ms / 1000) : ''; }
    /** IO voltages arrive in millivolts ("26542" -> 26.54 V). */
    function ioVolts(v) {
        var n = Number(v);
        if (v === null || v === undefined || v === '' || !isFinite(n)) { return null; }
        return n > 100 ? n / 1000 : n;
    }

    Ext.define('Store.battery_status.veh.Nav', {
        extend: 'Ext.panel.Panel',
        bodyPadding: '8 10',
        layout: 'anchor',
        initComponent: function () {
            var me = this;
            me.items = [
                { xtype: 'textfield', anchor: '100%', emptyText: t('Search plate / name / IMEI / branch / status…'),
                    listeners: { change: function (f, v) { me.main.search = String(v || '').trim().toLowerCase(); me.main.render(); }, buffer: 200 } },
                { xtype: 'component', itemId: 'info', cls: 'bs-muted', margin: '4 0 8 0', html: '' },
                { xtype: 'button', anchor: '100%', itemId: 'fetch', text: t('Fetch latest packets'), iconCls: 'fa fa-download', handler: function () { me.main.run(); } },
                { xtype: 'button', anchor: '100%', itemId: 'stop', text: t('Stop'), hidden: true, margin: '6 0 0 0', handler: function () { me.main.cancel = true; } },
                { xtype: 'button', anchor: '100%', itemId: 'reload', text: '↻ ' + t('Reload list'), margin: '6 0 0 0', handler: function () { me.main.loadList(); } },
                { xtype: 'component', itemId: 'progress', margin: '10 0 0 0', html: '' },
                { xtype: 'component', cls: 'bs-muted', margin: '10 0 0 0', html: enc(t('Latest packet of every vehicle from today\'s log on gps.im2m.ws. Click a row for the full packet, or Log for today\'s packets.')) }
            ];
            me.callParent(arguments);
        }
    });

    Ext.define('Store.battery_status.veh.Main', {
        extend: 'Ext.panel.Panel',
        cls: 'bs-main bs-report',
        layout: 'fit',

        initComponent: function () {
            var me = this;
            me.vehicles = [];
            me.results = {};
            me.open = {};
            me.logs = {};
            me.search = '';
            me.body0 = Ext.create('Ext.Component', {
                scrollable: true, padding: 12, html: '',
                listeners: { click: { element: 'el', fn: function (e) { me.onClick(e); } } }
            });
            me.items = [me.body0];
            me.callParent(arguments);
        },

        ensureLoaded: function () {
            if (!this.loaded) { this.loadList(); }
        },

        loadList: function () {
            var me = this;
            me.loaded = true;
            me.loading = true;
            me.loadErr = '';
            me.render();
            Api.call('/api/gps/vehicles').then(function (r) {
                me.vehicles = r.vehicles || [];
            }, function (e) {
                me.loaded = false;
                me.loadErr = errorText(e);
            }).then(function () { me.loading = false; me.render(); });
        },

        shown: function () {
            var q = this.search;
            return q ? Ext.Array.filter(this.vehicles, function (v) {
                return [v.name, v.branch, v.imei, v.id, v.status, v.location].join(' ').toLowerCase().indexOf(q) !== -1;
            }) : this.vehicles;
        },

        run: function () {
            var me = this, list = me.shown(), i = 0;
            if (!list.length || me.running) { return; }
            me.cancel = false;
            me.running = true;
            me.progress = { done: 0, total: list.length };
            me.render();
            function worker() {
                if (i >= list.length || me.cancel) { return Promise.resolve(); }
                var v = list[i++];
                return Api.call('/api/gps/packet', { json: { imei: v.imei } }).then(function (p) { me.results[v.imei] = { packet: p }; },
                    function (e) { me.results[v.imei] = { error: errorText(e) }; }).then(function () {
                    me.progress.done += 1;
                    me.render();
                    return worker();
                });
            }
            var workers = [];
            for (var k = 0; k < 6; k += 1) { workers.push(worker()); }
            Promise.all(workers).then(function () { me.running = false; me.render(); });
        },

        onClick: function (e) {
            var me = this, btn = e.getTarget('[data-act]'), row, imei;
            if (btn) {
                var act = btn.getAttribute('data-act');
                imei = btn.getAttribute('data-imei');
                if (act === 'log') {
                    if (me.logs[imei] && me.logs[imei].open) { me.logs[imei].open = false; } else { me.loadLog(imei); }
                } else if (act === 'logrefresh') {
                    me.loadLog(imei);
                } else if (act === 'retry') {
                    me.loadList();
                    return;
                }
                me.render();
                return;
            }
            row = e.getTarget('tr[data-raw]');
            if (row) {
                imei = row.getAttribute('data-imei');
                var idx = Number(row.getAttribute('data-raw'));
                me.logs[imei].raw = me.logs[imei].raw === idx ? null : idx;
                me.render();
                return;
            }
            row = e.getTarget('tr[data-imei]');
            if (row) {
                imei = row.getAttribute('data-imei');
                if (me.results[imei] && me.results[imei].packet) {
                    me.open[imei] = !me.open[imei];
                    me.render();
                }
            }
        },

        loadLog: function (imei) {
            var me = this, st = me.logs[imei] = { open: true, loading: true, raw: null };
            Api.call('/api/gps/log', { json: { imei: imei } }).then(function (r) { st.rows = r.rows || []; }, function (e) { st.error = errorText(e); })
                .then(function () { st.loading = false; me.render(); });
        },

        bleCell: function (ble, pending) {
            if (pending || !ble) { return '<span class="bs-muted">—</span>'; }
            if (ble.pct === null || ble.pct === undefined) {
                if (ble.lastRaw === null || ble.lastRaw === undefined) { return '<span class="bs-muted" title="' + enc(t("No BLE Battery field in today's packets")) + '">' + enc(t('no sensor')) + '</span>'; }
                return Number(ble.lastRaw) === 255 ? '<span class="bs-pill bs-nodata" title="255 = no reading">' + enc(t('No reading')) + '</span>' :
                    '<span class="bs-pill bs-alert" title="' + enc(t('Raw value') + ' ' + ble.lastRaw + ' ' + t('is outside 0–100')) + '">' + enc(t('Invalid')) + ' (' + enc(ble.lastRaw) + ')</span>';
            }
            var latestValid = ble.lastRaw !== null && ble.lastRaw !== undefined && Number(ble.lastRaw) === ble.pct;
            return '<span' + (ble.lastRaw !== null && !latestValid ? ' title="' + enc(t('Newest packet had') + ' ' + ble.lastRaw + '; ' + t('showing the last valid reading')) + '"' : '') + '>' +
                '<b' + (ble.pct < 20 ? ' style="color:#b3261e"' : '') + '>' + ble.pct + '%</b>' + (!latestValid ? '<div class="bs-muted">' + enc(t('at')) + ' ' + enc(fmtMs(ble.ms)) + '</div>' : '') + '</span>';
        },

        logHtml: function (imei) {
            var st = this.logs[imei], rows = st.rows || [], fixedIo = { 'External Voltage': 1, 'Battery Voltage': 1, 'Ignition': 1 }, extra = [], seen = {}, html = [];
            Ext.Array.each(rows, function (r) { Ext.Object.each(r.io || {}, function (k) { if (!fixedIo[k] && !seen[k]) { seen[k] = true; extra.push(k); } }); });
            html.push('<div class="bs-logbar"><b>' + enc(t("Today's log")) + ' — <code>' + enc(imei) + '</code></b>' + (rows.length ? ' <span class="bs-muted">· ' + rows.length + ' ' + enc(t('packets, newest first')) + '</span>' : '') +
                ' <button data-act="logrefresh" data-imei="' + enc(imei) + '"' + (st.loading ? ' disabled' : '') + '>↻ ' + enc(t('Refresh')) + '</button></div>');
            if (st.loading) { html.push('<div class="bs-muted">' + enc(t('Loading log…')) + '</div>'); }
            if (st.error) { html.push('<div class="bs-error">' + enc(st.error) + '</div>'); }
            if (!st.loading && !st.error && !rows.length) { html.push('<div class="bs-muted">' + enc(t('No packets today.')) + '</div>'); }
            if (rows.length) {
                html.push('<div style="max-height:420px;overflow:auto"><table class="bs-table" style="white-space:nowrap"><thead><tr><th>#</th><th>' + enc(t('Device time')) + '</th><th>' + enc(t('Received')) +
                    '</th><th>Lat, Lng</th><th class="bs-num">' + enc(t('Speed')) + '</th><th class="bs-num">' + enc(t('Sats')) + '</th><th>' + enc(t('Ignition')) + '</th><th class="bs-num">' +
                    enc(t('External V')) + '</th><th class="bs-num">' + enc(t('Battery V')) + '</th>' + Ext.Array.map(extra, function (k) { return '<th class="bs-num">' + enc(k) + '</th>'; }).join('') + '</tr></thead><tbody>');
                Ext.Array.each(rows, function (r, i) {
                    var ext = ioVolts(r.io['External Voltage']), bat = ioVolts(r.io['Battery Voltage']);
                    html.push('<tr data-raw="' + i + '" data-imei="' + enc(imei) + '" style="cursor:pointer" title="' + enc(t('Show raw packet')) + '"><td class="bs-muted">' + (rows.length - i) + '</td><td>' + enc(fmtMs(r.actMs)) +
                        '</td><td class="bs-muted">' + enc(fmtMs(r.receMs)) + '</td><td>' + enc(r.latt) + ', ' + enc(r.long) + '</td><td class="bs-num">' + enc(r.speed === null ? '—' : r.speed) +
                        '</td><td class="bs-num">' + enc(r.satellites === null ? '—' : r.satellites) + '</td><td>' + (r.io.Ignition === undefined ? '—' : (r.io.Ignition === '1' ? 'On' : 'Off')) +
                        '</td><td class="bs-num">' + (ext === null ? '—' : ext.toFixed(2)) + '</td><td class="bs-num">' + (bat === null ? '—' : bat.toFixed(3)) + '</td>' +
                        Ext.Array.map(extra, function (k) { return '<td class="bs-num">' + enc(r.io[k] === undefined ? '—' : r.io[k]) + '</td>'; }).join('') + '</tr>');
                    if (st.raw === i) { html.push('<tr class="bs-sub"><td colspan="' + (9 + extra.length) + '"><pre class="bs-pre">' + enc(r.raw) + '</pre></td></tr>'); }
                });
                html.push('</tbody></table></div>');
            }
            return html.join('');
        },

        render: function () {
            var me = this, shown = me.shown(), html = [], failed = 0, info = me.nav && me.nav.down('#info');
            Ext.Object.each(me.results, function (k, r) { if (r.error) { failed += 1; } });
            if (me.nav) {
                info.setHtml(me.loadErr && !me.loading ? '<div class="bs-error">' + enc(me.loadErr) + '</div>' : enc(me.loading ? t('Signing in to gps.im2m.ws and loading vehicles…') : shown.length + ' / ' + me.vehicles.length + ' ' + t('vehicles')));
                me.nav.down('#fetch').setText(t('Fetch latest packets') + ' (' + shown.length + ')');
                me.nav.down('#fetch').setDisabled(me.loading || me.running || !shown.length);
                me.nav.down('#fetch').setHidden(!!me.running);
                me.nav.down('#stop').setHidden(!me.running);
                me.nav.down('#reload').setDisabled(me.loading || me.running);
                me.nav.down('#progress').setHtml(me.progress && me.progress.total ? '<div class="bs-progress"><div style="width:' + Math.round(100 * me.progress.done / me.progress.total) + '%"></div></div>' +
                    '<div class="bs-muted bs-note">' + me.progress.done + ' / ' + me.progress.total + ' ' + enc(t('vehicles fetched')) + (failed ? ' · ' + failed + ' ' + enc(t('failed')) : '') + '</div>' : '');
            }
            html.push('<h2 class="bs-h">' + enc(t('Vehicles Battery Status')) + '</h2><div class="bs-muted bs-note">' + enc(t("Latest packet of every vehicle from today's log on gps.im2m.ws.")) + '</div>');
            if (me.loadErr) { html.push('<div class="bs-error">' + enc(me.loadErr) + ' <button data-act="retry">' + enc(t('Try again')) + '</button></div>'); }
            html.push('<table class="bs-table" style="margin-top:8px"><thead><tr><th>#</th><th>' + enc(t('Plate / Name')) + '</th><th>IMEI</th><th>' + enc(t('Status')) + '</th><th>' + enc(t('Ignition')) +
                '</th><th>' + enc(t('Last connection')) + '</th><th class="bs-num">' + enc(t('Battery %')) + '</th><th>' + enc(t('Latest packet (device time)')) + '</th><th class="bs-num">' +
                enc(t('External V')) + '</th><th class="bs-num">' + enc(t('Battery V')) + '</th><th class="bs-num">' + enc(t('Log rows today')) + '</th><th>' + enc(t('Log')) + '</th></tr></thead><tbody>');
            Ext.Array.each(shown, function (v, idx) {
                var r = me.results[v.imei], p = r && r.packet, isOpen = !!me.open[v.imei], io = (p && p.io) || {},
                    ext = ioVolts(io['External Voltage']), bat = ioVolts(io['Battery Voltage']),
                    lastConn = Math.max((p && p.receMs) || 0, v.lastPacketMs || 0) || null, mid, logOpen = me.logs[v.imei] && me.logs[v.imei].open;
                if (r && r.error) {
                    mid = '<td colspan="4"><span class="bs-pill bs-alert" title="' + enc(r.error) + '">' + enc(t('error')) + '</span> <span class="bs-muted">' + enc(r.error) + '</span></td>';
                } else if (p) {
                    mid = p.rows ? '<td>' + enc(fmtMs(p.actMs)) + ' <span class="bs-muted">' + enc(agoMs(p.actMs)) + '</span></td><td class="bs-num"><b>' + (ext === null ? '—' : ext.toFixed(2) + ' V') +
                        '</b></td><td class="bs-num"><b>' + (bat === null ? '—' : bat.toFixed(3) + ' V') + '</b></td><td class="bs-num">' + p.rows + '</td>' :
                        '<td colspan="4" class="bs-muted">' + enc(t('No packets today (last on platform') + ': ' + fmtMs(v.lastPacketMs) + ')') + '</td>';
                } else {
                    mid = '<td colspan="4" class="bs-muted">' + enc(me.running && !r ? '…' : t('Press Fetch latest packets')) + '</td>';
                }
                html.push('<tr data-imei="' + enc(v.imei) + '"' + (p ? ' style="cursor:pointer"' : '') + '><td class="bs-muted">' + (idx + 1) + '</td><td><span class="bs-muted" style="display:inline-block;width:14px">' +
                    (p ? (isOpen ? '▾' : '▸') : '') + '</span><b>' + enc(v.name || '—') + '</b>' + (v.branch ? '<div class="bs-muted" style="padding-left:14px">' + enc(v.branch) + '</div>' : '') +
                    '</td><td><code>' + enc(v.imei) + '</code><div class="bs-muted">ID ' + enc(v.id) + '</div></td><td>' + enc(v.status) + '</td><td>' +
                    enc(io.Ignition !== undefined ? (io.Ignition === '1' ? 'On' : 'Off') : v.ign) + '</td><td>' + enc(fmtMs(lastConn)) + '<div class="bs-muted">' + enc(agoMs(lastConn)) + '</div></td><td class="bs-num">' +
                    me.bleCell(p && p.ble, !p) + '</td>' + mid + '<td><button data-act="log" data-imei="' + enc(v.imei) + '">' + enc(logOpen ? t('Hide log') : t('Log')) + '</button></td></tr>');
                if (isOpen && p) {
                    html.push('<tr class="bs-sub"><td colspan="12" style="padding-left:28px"><div style="margin:4px 0 6px">' + enc(t('Position')) + ': <b>' + enc(p.latt) + ', ' + enc(p.long) + '</b> · ' +
                        enc(t('Device time')) + ': <b>' + enc(p.actTime) + '</b> · ' + enc(t('Received')) + ': <b>' + enc(p.receTime) + '</b></div><div class="bs-readings">' +
                        Ext.Array.map(Ext.Object.getKeys(io), function (k) { return '<span class="bs-reading">' + enc(k) + ': <b>' + enc(io[k]) + '</b></span>'; }).join('') +
                        '</div><pre class="bs-pre">' + enc(p.raw) + '</pre></td></tr>');
                }
                if (logOpen) { html.push('<tr class="bs-sub"><td colspan="12" style="padding-left:28px">' + me.logHtml(v.imei) + '</td></tr>'); }
            });
            if (!me.loading && !shown.length) { html.push('<tr><td colspan="12" class="bs-muted">' + enc(t('No vehicles.')) + '</td></tr>'); }
            html.push('</tbody></table>');
            me.body0.setHtml(html.join(''));
        }
    });

    // ---------------------------------------------------------------------
    // Entry point.
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.Module', {
        extend: 'Ext.Component',
        extensionName: 'battery_status',

        initModule: function () {
            var me = this;

            if (window.console) {
                console.log('battery_status: initModule() v' + VERSION);
            }

            if (!window.skeleton || !skeleton.navigation || !skeleton.navigation.add) {
                if (window.console) { console.warn('battery_status: skeleton.navigation not found'); }
                return;
            }

            var frame = skeleton.mapframe || skeleton.map_frame;
            if (!frame || !frame.add) {
                if (window.console) { console.warn('battery_status: skeleton.mapframe not found'); }
                return;
            }

            try {
                me.buildUi(frame);
            } catch (e) {
                // Surface the real cause instead of failing silently inside PILOT's loader.
                if (window.console) { console.error('battery_status: failed to build the UI', e); }
            }
        },

        buildUi: function (frame) {
            var me = this,
                win = storageGet('window', 'now');

            me.loadStyles();

            me.settings = {
                refreshSec: storageGet('refreshSec2', 0),
                window: (win === 'now' || Ext.Array.contains([6, 12, 48, 72], Number(win))) ? (win === 'now' ? 'now' : Number(win)) : 'now'
            };

            me.service = Ext.create('Store.battery_status.Data', {});

            me.mainPanel = Ext.create('Store.battery_status.view.MainPanel', {
                service: me.service,
                module: me
            });

            var LeftTabClass = Ext.ClassManager.get('Pilot.utils.LeftBarPanel') ? 'Pilot.utils.LeftBarPanel' : 'Ext.tab.Panel';

            // cls replaces LeftBarPanel's own classes, so keep the native ones
            // (same as the working temp_humidity_report tab) or the rail tab is not styled/shown.
            me.navTab = Ext.create(LeftTabClass, {
                title: t('Battery Status'),
                tooltip: t('Battery Status'),
                iconCls: 'fa fa-battery-three-quarters',
                iconAlign: 'top',
                minimized: true,
                cls: 'tab_clipped leftbarpanel bs-nav',
                bodyCls: 'left_top_round',
                items: [
                    Ext.create('Store.battery_status.view.ObjectTree', {
                        title: t('Warehouses'),
                        iconCls: 'fa fa-warehouse',
                        service: me.service,
                        mainPanel: me.mainPanel,
                        module: me
                    }),
                    Ext.create('Store.battery_status.view.Settings', {
                        title: t('Settings'),
                        iconCls: 'fa fa-sliders',
                        module: me
                    })
                ]
            });

            me.navTab.map_frame = me.mainPanel;

            skeleton.navigation.add(me.navTab);
            frame.add(me.mainPanel);

            // Load once the panel is first shown, then refresh only when asked (or on the chosen interval).
            me.mainPanel.on('activate', me.ensureLoaded, me);
            me.mainPanel.on('show', me.ensureLoaded, me);
            me.navTab.on('activate', me.ensureLoaded, me);

            me.startTimer();

            // The other Al Safi reports, each with its own icon on PILOT's left bar.
            // They use the Al Safi report backend (Vercel), so a sign-in is asked the first time.
            Api.init();
            me.addReportTab(frame, LeftTabClass, {
                title: t('Temperature Report'), iconCls: 'fa fa-temperature-half',
                nav: function (main) { return Ext.create('Store.battery_status.temp.Nav', { title: t('Report'), iconCls: 'fa fa-sliders', main: main }); },
                main: function () { return Ext.create('Store.battery_status.temp.Main', {}); }
            });
            var gstate = Ext.create('Store.battery_status.geo.State', {});
            me.addReportTab(frame, LeftTabClass, {
                title: t('Geozones'), iconCls: 'fa fa-draw-polygon',
                nav: function (main) { return Ext.create('Store.battery_status.geo.Nav', { title: t('Geozones'), iconCls: 'fa fa-list', gstate: gstate, main: main }); },
                main: function () { return Ext.create('Store.battery_status.geo.Map', { gstate: gstate }); }
            });
            me.addReportTab(frame, LeftTabClass, {
                title: t('Vehicles Battery'), iconCls: 'fa fa-car-battery',
                nav: function (main) { return Ext.create('Store.battery_status.veh.Nav', { title: t('Vehicles'), iconCls: 'fa fa-truck', main: main }); },
                main: function () { return Ext.create('Store.battery_status.veh.Main', {}); }
            });
        },

        /** One left-bar tab + paired main panel; the report loads the first time it is opened. */
        addReportTab: function (frame, LeftTabClass, def) {
            try {
                var main = def.main(),
                    nav = def.nav(main),
                    tab = Ext.create(LeftTabClass, {
                        title: def.title, tooltip: def.title, iconCls: def.iconCls, iconAlign: 'top', minimized: true,
                        cls: 'tab_clipped leftbarpanel bs-nav', bodyCls: 'left_top_round', items: [nav]
                    });
                main.nav = nav;
                tab.map_frame = main;
                skeleton.navigation.add(tab);
                frame.add(main);
                var load = function () { main.ensureLoaded(); };
                main.on('activate', load);
                main.on('show', load);
                tab.on('activate', load);
            } catch (e) {
                if (window.console) { console.error('battery_status: failed to add the ' + def.title + ' tab', e); }
            }
        },

        ensureLoaded: function () {
            if (!this.service.lastLoad && !this.service.loading) {
                this.reload();
            }
        },

        reload: function () {
            this.service.load(this.settings.window);
        },

        applySettings: function (changes) {
            Ext.apply(this.settings, changes);
            if (Ext.isDefined(changes.refreshSec)) {
                storageSet('refreshSec2', changes.refreshSec);
                this.startTimer();
            }
            if (Ext.isDefined(changes.window)) {
                storageSet('window', changes.window);
                if (this.service.lastLoad) {
                    this.reload();
                }
            }
        },

        startTimer: function () {
            var me = this;
            if (me.timer) {
                clearInterval(me.timer);
                me.timer = null;
            }
            if (me.settings.refreshSec > 0) {
                me.timer = setInterval(function () {
                    // Skip background polling while the browser tab is hidden.
                    if (!document.hidden && me.service.lastLoad) {
                        me.reload();
                    }
                }, me.settings.refreshSec * 1000);
            }
        },

        /** Readings of one sensor in the last 24 hours (newest first), with min/max and a trend line. */
        showHistory: function (record) {
            var me = this,
                name = record.get('name'),
                hex = record.get('idHex'),
                body = Ext.create('Ext.Component', {
                    scrollable: true,
                    padding: 10,
                    html: '<div class="bs-muted">' + enc(t('Loading readings…')) + '</div>'
                }),
                win = Ext.create('Ext.window.Window', {
                    title: enc(name) + ' <code>' + enc(hex) + '</code> — ' + enc(t('readings in the last')) + ' ' + HISTORY_HOURS + ' ' + enc(t('hours')),
                    width: Math.min(820, Ext.getBody().getViewSize().width - 40),
                    height: Math.min(560, Ext.getBody().getViewSize().height - 40),
                    layout: 'fit',
                    modal: true,
                    items: [body]
                });

            win.show();
            me.service.probeHistory(record.get('agentId'), record.get('idDec'), HISTORY_HOURS, function (rs, error) {
                if (win.destroyed || win.isDestroyed) {
                    return;
                }
                if (!rs) {
                    body.setHtml('<div class="bs-error">' + enc(error) + '</div>');
                    return;
                }
                if (!rs.length) {
                    body.setHtml('<div class="bs-muted">' + enc(t('No readings from this sensor in the last')) + ' ' + HISTORY_HOURS + ' ' + enc(t('hours')) + '.</div>');
                    return;
                }
                var volts = Ext.Array.filter(Ext.Array.pluck(rs, 'voltage'), function (v) { return v !== null; }),
                    lows = Ext.Array.filter(rs, function (r) { return r.status === 'low'; }).length,
                    html = '<div class="bs-muted" style="margin-bottom:6px">' + rs.length + ' ' + enc(t('readings')) +
                        (volts.length ? ' · ' + enc(t('Voltage min')) + ' <b>' + Math.min.apply(Math, volts).toFixed(3) + '</b> · ' +
                            enc(t('max')) + ' <b>' + Math.max.apply(Math, volts).toFixed(3) + '</b> V' : '') + ' · ' +
                        (lows ? '<span style="color:#b3261e">' + lows + ' ' + enc(t('below')) + ' ' + LOW_VOLTAGE + ' V</span>' : enc(t('all')) + ' ≥ ' + LOW_VOLTAGE + ' V') +
                        '</div>' +
                        (volts.length > 1 ? me.spark(Ext.Array.pluck(rs.slice().reverse(), 'voltage')) : '') +
                        '<table class="bs-table"><thead><tr><th>' + enc(t('Time')) + '</th><th class="bs-num">' + enc(t('Voltage')) + '</th><th>' +
                        enc(t('Status')) + '</th><th>' + enc(t('Battery Status')) + '</th><th class="bs-num">' + enc(t('Temp °C')) + '</th><th class="bs-num">' +
                        enc(t('Hum %')) + '</th><th class="bs-num">' + enc(t('Signal')) + '</th></tr></thead><tbody>' +
                        Ext.Array.map(rs, function (r) {
                            return '<tr class="bs-row-' + r.status + '"><td>' + enc(fmtTs(r.ts)) + '</td><td class="bs-num">' +
                                (r.voltage === null ? '—' : r.voltage.toFixed(3)) + '</td><td>' + statusPill(r.status) + '</td><td class="bs-muted">' +
                                enc(r.platformStatus === null ? '—' : r.platformStatus) + '</td><td class="bs-num">' + fixed(r.temperature) +
                                '</td><td class="bs-num">' + fixed(r.humidity) + '</td><td class="bs-num">' + (r.rssi === null ? '—' : r.rssi) + '</td></tr>';
                        }).join('') + '</tbody></table>';
                body.setHtml(html);
            });
        },

        /** Tiny voltage line (oldest -> newest) with the low-voltage threshold dashed. */
        spark: function (values) {
            var W = 600, H = 60, P = 4,
                vs = Ext.Array.filter(values, function (v) { return v !== null; }),
                lo = Math.min.apply(Math, vs.concat([LOW_VOLTAGE])) - 0.05,
                hi = Math.max.apply(Math, vs.concat([LOW_VOLTAGE])) + 0.05,
                x = function (i) { return P + (i * (W - 2 * P)) / Math.max(1, values.length - 1); },
                y = function (v) { return H - P - ((v - lo) * (H - 2 * P)) / (hi - lo || 1); },
                d = Ext.Array.map(values, function (v, i) {
                    return v === null ? '' : ((i && values[i - 1] !== null ? 'L' : 'M') + x(i).toFixed(1) + ',' + y(v).toFixed(1));
                }).join('');
            return '<svg viewBox="0 0 ' + W + ' ' + H + '" style="width:100%;max-width:600px;height:60px;display:block;margin-bottom:6px">' +
                '<line x1="' + P + '" x2="' + (W - P) + '" y1="' + y(LOW_VOLTAGE) + '" y2="' + y(LOW_VOLTAGE) +
                '" stroke="#b3261e" stroke-dasharray="4 4" stroke-width="1" opacity="0.6"/>' +
                '<path d="' + d + '" fill="none" stroke="#2f6db5" stroke-width="2"/></svg>';
        },

        /** Every configured sensor of an object with its latest value, plus gateway info. */
        showAllSensors: function (agentId) {
            var res = this.service.results[agentId];
            if (!res || !res.data) {
                return;
            }
            var data = res.data,
                html = '<div style="margin-bottom:8px">' + enc(t('Gateway battery')) + ': <b>' + fixed(data.gateway.batteryVoltage, 2) + ' V</b> · ' +
                    enc(t('Power supply')) + ': <b>' + fixed(data.gateway.powerVoltage, 2) + ' V</b> · ' +
                    enc(t('Last packet')) + ': <b>' + enc(fmtTs(data.lastPacket)) + '</b> <span class="bs-muted">' + enc(ago(data.lastPacket)) + '</span></div>' +
                    (data.sensors.length ? '<table class="bs-table"><thead><tr><th>' + enc(t('Sensor')) + '</th><th>' + enc(t('Field')) + '</th><th>' +
                        enc(t('Group')) + '</th><th class="bs-num">' + enc(t('Latest value')) + '</th><th>' + enc(t('Last reading')) + '</th><th>ID</th></tr></thead><tbody>' +
                        Ext.Array.map(data.sensors, function (s) {
                            return '<tr><td>' + enc(s.name) + '</td><td class="bs-muted">' + enc(s.field) + '</td><td>' + enc(s.group) +
                                '</td><td class="bs-num">' + (s.value === null ? '—' : enc(s.value + (s.unit ? ' ' + s.unit : ''))) +
                                '</td><td>' + enc(fmtTs(s.ts)) + '</td><td class="bs-muted">' + enc(s.sensorid) + '</td></tr>';
                        }).join('') + '</tbody></table>' : '<div class="bs-muted">' + enc(t('No sensors configured.')) + '</div>');

            Ext.create('Ext.window.Window', {
                title: enc(res.obj.name) + ' — ' + enc(t('All sensors')),
                width: Math.min(900, Ext.getBody().getViewSize().width - 40),
                height: Math.min(600, Ext.getBody().getViewSize().height - 40),
                layout: 'fit',
                modal: true,
                items: [{ xtype: 'component', scrollable: true, padding: 10, html: html }]
            }).show();
        },

        loadStyles: function () {
            if (document.getElementById('battery_status-css')) {
                return;
            }
            var css = document.createElement('link');
            css.id = 'battery_status-css';
            css.setAttribute('rel', 'stylesheet');
            css.setAttribute('type', 'text/css');
            css.setAttribute('href', '/store/battery_status/style.css?v=' + VERSION);
            document.head.appendChild(css);
        }
    });
}());
