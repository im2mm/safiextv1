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
    var VERSION = '1.2.0'; // bump on every release; shown in the console, Settings tab and doc/

    var TZ = 'Asia/Riyadh';          // account local time used by the platform
    var TZ_OFFSET_S = 3 * 3600;      // Riyadh has no DST; used for Excel date cells
    var LOW_VOLTAGE = 3.3;           // below this = low battery
    var STALE_HOURS = 3;             // no reading for longer than this = "No reading" (not low)
    var HISTORY_HOURS = 24;          // fixed window for a sensor's previous readings
    var CONCURRENCY = 3;
    var WINDOWS = [['now', 'Now'], [6, 'Last 6 hours'], [12, 'Last 12 hours'], [48, 'Last 48 hours'], [72, 'Last 72 hours']];

    function t(text) {
        return (typeof window.l === 'function') ? window.l(text) : text;
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
                    me.fireEvent('progress', me, done, total);
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

        finish: function (results) {
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
            me.loading = false;
            me.lastLoad = new Date();
            me.fireEvent('load', me, rows, errors);
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
        },

        // ---- Excel export (SpreadsheetML, no library needed) --------------

        exportExcel: function () {
            var me = this,
                results = Ext.Object.getValues(me.results),
                styles = {},
                styleXml = [],
                ORDER = { low: 0, nodata: 1, normal: 2 };

            function styleId(fmt, fill) {
                var key = (fmt || '') + '|' + (fill || '');
                if (!styles[key]) {
                    styles[key] = 's' + (Ext.Object.getSize(styles) + 1);
                    styleXml.push('<Style ss:ID="' + styles[key] + '">' +
                        (fmt ? '<NumberFormat ss:Format="' + enc(fmt) + '"/>' : '') +
                        (fill ? '<Interior ss:Color="' + fill + '" ss:Pattern="Solid"/>' : '') + '</Style>');
                }
                return styles[key];
            }
            function xlDate(sec) {
                return sec ? { date: new Date((sec + TZ_OFFSET_S) * 1000).toISOString().slice(0, 19) } : null;
            }
            function cell(v, fmt, fill) {
                var sid = styleId(v && v.date ? 'dd mmm yyyy hh:mm' : fmt, fill),
                    attr = ' ss:StyleID="' + sid + '"';
                if (v === null || v === undefined || v === '') {
                    return '<Cell' + attr + '/>';
                }
                if (v.date) {
                    return '<Cell' + attr + '><Data ss:Type="DateTime">' + v.date + '</Data></Cell>';
                }
                if (typeof v === 'number') {
                    return '<Cell' + attr + '><Data ss:Type="Number">' + v + '</Data></Cell>';
                }
                return '<Cell' + attr + '><Data ss:Type="String">' + enc(v) + '</Data></Cell>';
            }
            function sheet(name, cols, rows, fillOf) {
                var xml = ['<Worksheet ss:Name="' + enc(name) + '"><Table>'];
                Ext.Array.each(cols, function (c) { xml.push('<Column ss:Width="' + (c.w * 6) + '"/>'); });
                xml.push('<Row>' + Ext.Array.map(cols, function (c) {
                    return '<Cell ss:StyleID="h"><Data ss:Type="String">' + enc(c.h) + '</Data></Cell>';
                }).join('') + '</Row>');
                Ext.Array.each(rows, function (r) {
                    var fill = fillOf ? fillOf(r) : null;
                    xml.push('<Row>' + Ext.Array.map(cols, function (c) { return cell(c.v(r), c.f, fill); }).join('') + '</Row>');
                });
                xml.push('</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/>' +
                    '<SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane></WorksheetOptions>' +
                    '<AutoFilter x:Range="R1C1:R1C' + cols.length + '" xmlns="urn:schemas-microsoft-com:office:excel"/></Worksheet>');
                return xml.join('');
            }
            function numOrText(v) {
                if (v === null || String(v).trim() === '') { return null; }
                var n = Number(v);
                return isFinite(n) ? n : v;
            }
            function count(r, s) {
                return r.data ? Ext.Array.filter(r.data.probes, function (p) { return !s || p.status === s; }).length : null;
            }

            var battery = [], sensors = [];
            Ext.Array.each(results, function (r) {
                if (!r.data) { return; }
                Ext.Array.each(r.data.probes, function (p) { battery.push(Ext.apply({ obj: r.obj }, p)); });
                Ext.Array.each(r.data.sensors, function (s) { sensors.push(Ext.apply({ obj: r.obj }, s)); });
            });
            battery.sort(function (a, b) { return byName(a.obj.name, b.obj.name) || ORDER[a.status] - ORDER[b.status]; });
            results.sort(function (a, b) { return byName(a.obj.name, b.obj.name); });
            sensors.sort(function (a, b) { return byName(a.obj.name, b.obj.name) || byName(a.name, b.name); });

            var xml = '<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>' +
                '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
                'xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
            var body =
                sheet('Battery voltage', [
                    { h: 'Object', w: 28, v: function (r) { return r.obj.name; } },
                    { h: 'Site', w: 22, v: function (r) { return r.obj.region; } },
                    { h: 'Sensor', w: 44, v: function (r) { return r.name; } },
                    { h: 'Sensor ID', w: 12, v: function (r) { return r.idHex; } },
                    { h: 'Voltage (V)', w: 12, f: '0.000', v: function (r) { return r.voltage; } },
                    { h: 'Status', w: 12, v: function (r) { return STATUS[r.status].label; } },
                    { h: 'Battery Status', w: 15, v: function (r) { return r.platformStatus || ''; } },
                    { h: 'Temp (°C)', w: 10, f: '0.0', v: function (r) { return r.temperature; } },
                    { h: 'Humidity (%)', w: 12, f: '0.0', v: function (r) { return r.humidity; } },
                    { h: 'Signal (RSSI)', w: 12, v: function (r) { return r.rssi; } },
                    { h: 'Last reading', w: 18, v: function (r) { return xlDate(r.ts); } }
                ], battery, function (r) { return r.status === 'low' ? '#FDECEC' : (r.status === 'nodata' ? '#F1F3F5' : null); }) +
                sheet('Gateways', [
                    { h: 'Object', w: 28, v: function (r) { return r.obj.name; } },
                    { h: 'Site', w: 22, v: function (r) { return r.obj.region; } },
                    { h: 'Gateway IMEI', w: 18, v: function (r) { return r.obj.imei || ''; } },
                    { h: 'Gateway battery (V)', w: 18, f: '0.00', v: function (r) { return r.data ? r.data.gateway.batteryVoltage : null; } },
                    { h: 'Power supply (V)', w: 16, f: '0.00', v: function (r) { return r.data ? r.data.gateway.powerVoltage : null; } },
                    { h: 'Last packet', w: 18, v: function (r) { return xlDate(r.data && r.data.lastPacket); } },
                    { h: 'Sensors', w: 9, v: function (r) { return count(r); } },
                    { h: 'Low', w: 7, v: function (r) { return count(r, 'low'); } },
                    { h: 'No reading', w: 11, v: function (r) { return count(r, 'nodata'); } },
                    { h: 'Normal', w: 9, v: function (r) { return count(r, 'normal'); } },
                    { h: 'Error', w: 40, v: function (r) { return r.error || ''; } }
                ], results, function (r) { return r.error ? '#FDECEC' : null; }) +
                sheet('All sensors', [
                    { h: 'Object', w: 28, v: function (r) { return r.obj.name; } },
                    { h: 'Sensor', w: 44, v: function (r) { return r.name; } },
                    { h: 'Field', w: 30, v: function (r) { return r.field; } },
                    { h: 'Sensor group', w: 22, v: function (r) { return r.group; } },
                    { h: 'Type', w: 12, v: function (r) { return r.type; } },
                    { h: 'Unit', w: 8, v: function (r) { return r.unit; } },
                    { h: 'Latest value', w: 14, v: function (r) { return numOrText(r.value); } },
                    { h: 'Last reading', w: 18, v: function (r) { return xlDate(r.ts); } },
                    { h: 'Sensor ID', w: 10, v: function (r) { return Number(r.sensorid) || r.sensorid; } }
                ], sensors);

            xml += '<Styles><Style ss:ID="h"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0B7A75" ss:Pattern="Solid"/></Style>' +
                styleXml.join('') + '</Styles>' + body + '</Workbook>';

            var p = riyadhParts(Date.now()),
                a = document.createElement('a'),
                url = URL.createObjectURL(new Blob([xml], { type: 'application/vnd.ms-excel' }));
            a.href = url;
            a.download = 'Warehouse Sensor Status ' + p.year + '-' + p.month + '-' + p.day + ' ' + p.hour + '-' + p.minute + '.xls';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            Ext.defer(function () { URL.revokeObjectURL(url); }, 1000);
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
                    me.summaryCmp.setHtml('<div class="bs-muted">' + enc(t('Loading objects…')) + '</div>');
                },
                progress: function (data, done, total) {
                    me.summaryCmp.setHtml('<div class="bs-progress"><div style="width:' + Math.round(100 * done / total) + '%"></div></div>' +
                        '<div class="bs-muted bs-note">' + done + ' / ' + total + ' ' + enc(t('objects processed')) + '</div>');
                },
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

        download: function () {
            if (!this.service.lastLoad) {
                Ext.Msg.alert(t('Battery Status'), t('Fetch the battery status first.'));
                return;
            }
            this.service.exportExcel();
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

        onDataLoad: function (data, rows, errors) {
            var lastLoad = this.down('#lastLoad'),
                alertsTab = this.child('#alerts'),
                s = data.getSummary(),
                alerts = s.low + s.nodata;

            this.setBusy(false);
            this.applyFilters();
            this.renderSummary();

            if (lastLoad) {
                lastLoad.setText(t('fetched') + ' ' + Ext.Date.format(new Date(), 'H:i') + ' · ' +
                    (data.win === 'now' ? t('Now') : t('last') + ' ' + data.win + ' h'));
            }
            if (alertsTab && alertsTab.tab) {
                alertsTab.setTitle(t('Alerts') + (alerts ? ' <span class="bs-count">' + alerts + '</span>' : ''));
            }

            this.syncObjectTabs();
        },

        renderSummary: function () {
            var me = this, data = me.service, s, failed;
            if (!data.lastLoad) {
                return;
            }
            s = data.getSummary();
            failed = data.getFailed();
            me.summaryCmp.setHtml(
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

        rebuild: function (data) {
            var regions = {},
                order = [],
                children = [];

            Ext.Array.each(data.getTagObjects(), function (obj) {
                var key = obj.region || t('Ungrouped'),
                    stats = data.getObjectStats(obj.agentId);
                if (!regions[key]) {
                    regions[key] = [];
                    order.push(key);
                }
                regions[key].push({
                    name: obj.name,
                    agentId: obj.agentId,
                    leaf: true,
                    iconCls: STATUS[stats.worst].icon + ' ' + STATUS[stats.worst].cls,
                    badge: stats.total ? stats.normal + '/' + stats.total : '-',
                    worst: stats.worst
                });
            });

            Ext.Array.each(order, function (key) {
                var worst = 'normal', ok = 0, total = 0;
                Ext.Array.each(regions[key], function (leaf) {
                    var stats = data.getObjectStats(leaf.agentId);
                    ok += stats.normal;
                    total += stats.total;
                    if (STATUS[stats.worst].rank > STATUS[worst].rank) {
                        worst = stats.worst;
                    }
                });
                children.push({
                    name: key,
                    expanded: true,
                    iconCls: 'fa fa-folder',
                    badge: total ? ok + '/' + total : '',
                    worst: worst,
                    children: regions[key]
                });
            });

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
