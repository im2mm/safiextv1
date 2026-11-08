/**
 * PILOT Extension: battery_status
 *
 * Shows the live battery status of every BLE tag (temperature/humidity sensor)
 * on objects with tags. The status text (BatteryVoltageStatus<tag id> = "Normal")
 * is read from the raw packets of Sensors Tracing, because the BTS-<hex>-<name>
 * Selector sensors get a TEXT value and show no data in the normal sensor list.
 *
 * Layout (nested tabs):
 *   Left navigation tab "Battery Status" (Pilot.utils.LeftBarPanel) with two inner tabs:
 *     - Warehouses : region -> object tree with ok/total badges; click an object to open it
 *     - Settings   : auto-refresh interval and sensor-tracing look-back window
 *   Main panel (Ext.tab.Panel) with three tabs:
 *     - Overview   : summary tiles + every battery sensor grouped by object
 *     - Alerts     : tags that are not Normal: alert, not reporting or unknown text
 *     - Warehouses : nested vertical tab per object with that object's battery sensors
 *
 * Data sources (read-only, see the Data class below):
 *   - ax/tree.php?vehs=1&state=1                    object tree
 *   - ax/sensors/sensors.php?agent_id=<id>          sensor config (names, field names)
 *   - ax/sensors_trace.php?cmd=get&agent_id=<id>... raw packets (Sensors Tracing); the battery
 *                                                    status TEXT is read from "triggerdata"
 *   These are legacy PILOT backend endpoints observed in the running web app, not part
 *   of the documented /api/v3 surface.
 *
 * All classes live in this single file so the extension does not depend on
 * Ext.Loader resolving extra Store.battery_status.* files.
 */

(function () {
    var EXT = 'battery_status';

    function t(text) {
        return (typeof window.l === 'function') ? window.l(text) : text;
    }

    function enc(value) {
        return Ext.String.htmlEncode(value === null || value === undefined ? '' : String(value));
    }

    function dec(value) {
        return Ext.String.htmlDecode(value === null || value === undefined ? '' : String(value));
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

    // Status meta: order is used to pick the "worst" status of an object.
    var STATUS = {
        alert:   { rank: 4, label: 'Alert',     icon: 'fa fa-battery-quarter',       cls: 'bs-alert' },
        nodata:  { rank: 3, label: 'Not reporting', icon: 'fa fa-signal',            cls: 'bs-nodata' },
        unknown: { rank: 2, label: 'Unknown',   icon: 'fa fa-circle-question',       cls: 'bs-unknown' },
        ok:      { rank: 0, label: 'Normal',    icon: 'fa fa-battery-full',          cls: 'bs-ok' }
    };

    function statusPill(status, text) {
        var meta = STATUS[status] || STATUS.unknown;
        return '<span class="bs-pill ' + meta.cls + '"><i class="' + meta.icon + '"></i> ' +
            enc(text || t(meta.label)) + '</span>';
    }

    function formatTs(ts) {
        ts = Number(ts || 0);
        if (!ts) {
            return '';
        }
        var date = new Date(ts * 1000),
            diff = Math.max(0, Math.round(Date.now() / 1000 - ts)),
            ago;

        if (diff < 3600) {
            ago = Math.round(diff / 60) + ' ' + t('min ago');
        } else if (diff < 86400) {
            ago = Math.round(diff / 3600) + ' ' + t('h ago');
        } else {
            ago = Math.round(diff / 86400) + ' ' + t('d ago');
        }

        return enc(Ext.Date.format(date, 'd.m.Y H:i')) +
            ' <span class="bs-muted">(' + enc(ago) + ')</span>';
    }

    // ---------------------------------------------------------------------
    // Data service.
    //
    // Battery status sensors are "Selector" sensors fed by a TEXT field
    // (BatteryVoltageStatus<tag id> = "Normal"), so /api/v3/vehicles/status
    // usually returns nothing for them. The raw packet data from Sensors
    // Tracing still carries the text, so this service reads it from there:
    //
    //   1. ax/tree.php?vehs=1&state=1                 object tree
    //   2. ax/sensors/sensors.php?agent_id=<id>       sensor config per object
    //      (names, field names) - cached, refreshed every 30 minutes
    //   3. ax/sensors_trace.php?cmd=get&agent_id=<id>&ts=..&te=..
    //      raw packets of the last N hours; "triggerdata" is parsed and the
    //      newest value of every BatteryVoltageStatus<dec>, BatteryVoltage<dec>,
    //      Temperature<dec>, Humidity<dec>, RSSI<dec> key is kept per tag.
    //
    // Only objects that have BLE tag fields (Temperature<dec>, BatteryVoltage*<dec>)
    // are traced, a few at a time, so trucks without tags cost one config call only.
    // ---------------------------------------------------------------------
    var TAG_FIELD_RE = /^(BatteryVoltageStatus|BatteryVoltage|Temperature|Humidity)(\d{6,})$/;
    var BATTERY_FIELD_RE = /^(BatteryVoltageStatus|BatteryVoltage)(\d{6,})$/;
    var PAIR_RE = /([A-Za-z]+?)(\d{6,})=([^,]*)/g;

    function decToHex(decId) {
        var n = Number(decId);
        return isFinite(n) && n > 0 ? n.toString(16).toUpperCase() : '';
    }

    Ext.define('Store.battery_status.Data', {
        mixins: ['Ext.mixin.Observable'],

        parallel: 3,                      // trace requests running at the same time
        configTtlMs: 30 * 60 * 1000,      // sensor config cache lifetime
        traceLimit: 2000,

        constructor: function (config) {
            this.mixins.observable.constructor.call(this, config);
            this.objects = [];            // [{agentId, name, region, path}]
            this.configCache = {};        // agentId -> {ts, sensors}
            this.objectInfo = {};         // agentId -> {gatewayBattery, powerVoltage, lastPacket, traced}
            this.store = Ext.create('Ext.data.Store', {
                fields: [
                    'key', 'agentId', 'objectName', 'region', 'tagDec', 'hexId', 'location',
                    'btsName', 'configured', 'statusText', 'level', 'temperature', 'humidity',
                    'rssi', 'status',
                    { name: 'lastTs', type: 'int' },
                    { name: 'rank', type: 'int' }
                ],
                sorters: [{ property: 'rank', direction: 'DESC' }, { property: 'location', direction: 'ASC' }]
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

        classify: function (text) {
            var value = String(text || '').trim().toLowerCase();

            if (!value) {
                return 'unknown';
            }
            // Check negative words first: "Not normal" also contains "normal".
            if (/not\s*normal|abnormal|low|critical|replace|bad|empty|fail|error|alarm|dead|weak|fault/.test(value)) {
                return 'alert';
            }
            if (/normal|^ok$|good|full|high|charged|healthy/.test(value)) {
                return 'ok';
            }
            return 'unknown';
        },

        parseName: function (name) {
            // BTS-<hex sensor id>-<location name>
            var match = /^BTS-([0-9A-F]+)-(.*)$/i.exec(name);
            if (match) {
                return { hexId: match[1].toUpperCase(), location: match[2].trim() };
            }
            return { hexId: '', location: name };
        },

        collectObjects: function (nodes, path) {
            Ext.Array.each(nodes || [], function (node) {
                if (!node) {
                    return;
                }
                var name = dec(node.name || ''),
                    agentId = node.agentid || node.vehid || node.id;

                if (node.leaf || (node.agentid && !node.children)) {
                    this.objects.push({
                        agentId: Number(agentId),
                        name: name,
                        region: path.length ? path[path.length - 1] : '',
                        path: path.slice()
                    });
                } else if (Ext.isArray(node.children)) {
                    this.collectObjects(node.children, path.concat([name.trim()]));
                }
            }, this);
        },

        // ---- loading -----------------------------------------------------

        load: function (hoursBack) {
            var me = this;

            if (me.loading) {
                return;
            }
            me.loading = true;
            me.hoursBack = Math.min(24, Math.max(1, Number(hoursBack) || 3));
            me.fireEvent('beforeload', me);

            me.requestJson(me.axUrl('tree.php'), { vehs: 1, state: 1 }, function onTree(groups, ok, status) {
                if (!ok && !me.triedTreeFallback) {
                    // Some installations serve the legacy endpoints from the site root.
                    me.triedTreeFallback = true;
                    me.axBase = '/';
                    me.requestJson('/ax/tree.php', { vehs: 1, state: 1 }, onTree);
                    return;
                }
                if (!ok) {
                    me.loading = false;
                    me.fireEvent('loaderror', me, t('Could not load the object list') + ' (HTTP ' + status + ')');
                    return;
                }
                me.objects = [];
                me.collectObjects(Ext.isArray(groups) ? groups : (groups && groups.children) || [], []);
                me.loadObjects();
            });
        },

        requestJson: function (url, params, callback) {
            Ext.Ajax.request({
                url: url,
                method: 'GET',
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

        // Runs loadObject() for every object with a small worker pool.
        loadObjects: function () {
            var me = this,
                queue = me.objects.slice(),
                rows = [],
                errors = 0,
                running = 0,
                done = 0,
                total = queue.length;

            if (!total) {
                me.finish([], 0);
                return;
            }

            function next() {
                if (!queue.length) {
                    if (running === 0) {
                        me.finish(rows, errors);
                    }
                    return;
                }
                var obj = queue.shift();
                running += 1;
                me.loadObject(obj, function (objRows, failed) {
                    running -= 1;
                    done += 1;
                    if (failed) {
                        errors += 1;
                    }
                    rows.push.apply(rows, objRows);
                    me.fireEvent('progress', me, done, total);
                    next();
                });
            }

            for (var i = 0; i < me.parallel; i += 1) {
                next();
            }
        },

        loadObject: function (obj, done) {
            var me = this,
                cached = me.configCache[obj.agentId];

            if (cached && (Date.now() - cached.ts) < me.configTtlMs) {
                me.traceObject(obj, cached.sensors, done);
                return;
            }

            me.requestJson(me.axUrl('sensors/sensors.php'), { agent_id: obj.agentId }, function (json, ok) {
                var sensors = Ext.isArray(json) ? json : (json && Ext.isArray(json.data) ? json.data : null);
                if (!ok || !sensors) {
                    done([], true);
                    return;
                }
                me.configCache[obj.agentId] = { ts: Date.now(), sensors: sensors };
                me.traceObject(obj, sensors, done);
            });
        },

        // Builds the tag list from the sensor config, then reads the raw trace.
        traceObject: function (obj, sensors, done) {
            var me = this,
                tags = me.tagsFromConfig(sensors),
                end = new Date(),
                start = Ext.Date.add(end, Ext.Date.HOUR, -me.hoursBack);

            if (!tags.hasTagFields) {
                // No BLE tag on this object: nothing to show, no trace request.
                done([], false);
                return;
            }

            me.requestJson(me.axUrl('sensors_trace.php'), {
                cmd: 'get',
                agent_id: obj.agentId,
                ts: Ext.Date.format(start, 'Y-m-d\\TH:i:s'),
                te: Ext.Date.format(end, 'Y-m-d\\TH:i:s'),
                page: 1,
                start: 0,
                limit: me.traceLimit
            }, function (json, ok) {
                var items = json && Ext.isArray(json.items) ? json.items : [],
                    latest = me.parseTrace(items);

                me.objectInfo[obj.agentId] = Ext.apply({ traced: ok }, latest.gateway);
                done(me.buildRows(obj, tags, latest.tags), !ok);
            });
        },

        tagsFromConfig: function (sensors) {
            var result = { hasTagFields: false, byDec: {} };

            Ext.Array.each(sensors, function (s) {
                var field = String(s.sensorfieldname || ''),
                    match = TAG_FIELD_RE.exec(field),
                    info = dec(s.info || '').trim(),
                    tag;

                if (!match) {
                    return;
                }
                result.hasTagFields = true;
                tag = result.byDec[match[2]] = result.byDec[match[2]] || { battery: false };

                if (match[1] === 'BatteryVoltageStatus' && /^BTS-/i.test(info)) {
                    tag.btsName = info;
                }
                if (BATTERY_FIELD_RE.test(field)) {
                    tag.battery = true;
                    tag.batteryName = tag.batteryName || info;
                }
                // Prefer the plain temperature sensor over formula ones (e.g. "AVERAGE TEMP").
                if (match[1] === 'Temperature' && (!tag.tempName || (tag.tempIsFormula && !s.formula))) {
                    tag.tempName = info;
                    tag.tempIsFormula = !!s.formula;
                }
            });
            return result;
        },

        // triggerdata: "Key<dec>=value, Key=value, ..."; items are newest first,
        // but every value is compared by packet time to be safe.
        parseTrace: function (items) {
            var tags = {},
                gateway = {};

            Ext.Array.each(items, function (item) {
                var ts = Number(item.ts || item.ts_date || 0),
                    text = String(item.triggerdata || ''),
                    m;

                if (!gateway.lastPacket || ts > gateway.lastPacket) {
                    var gb = /(?:^|,\s*)BatteryVoltage=([^,]*)/.exec(text),
                        pv = /(?:^|,\s*)PowerVoltage=([^,]*)/.exec(text);
                    gateway.lastPacket = ts;
                    if (gb) { gateway.gatewayBattery = gb[1].trim(); }
                    if (pv) { gateway.powerVoltage = pv[1].trim(); }
                }

                PAIR_RE.lastIndex = 0;
                while ((m = PAIR_RE.exec(text)) !== null) {
                    var key = m[1], id = m[2], value = m[3].trim(),
                        tag = tags[id] = tags[id] || {},
                        slot = tag[key];

                    if (value === '') {
                        continue;
                    }
                    if (!slot || ts > slot.ts) {
                        tag[key] = { value: value, ts: ts };
                    }
                }
            });

            return { tags: tags, gateway: gateway };
        },

        buildRows: function (obj, config, trace) {
            var me = this,
                rows = [],
                ids = {};

            // Every tag that has a battery sensor configured, plus every tag that
            // reported a battery value in the trace window.
            Ext.Object.each(config.byDec, function (id, tag) {
                if (tag.battery) { ids[id] = true; }
            });
            Ext.Object.each(trace, function (id, tag) {
                if (tag.BatteryVoltageStatus || tag.BatteryVoltage) { ids[id] = true; }
            });

            Ext.Object.each(ids, function (id) {
                var cfg = config.byDec[id] || {},
                    raw = trace[id] || {},
                    hex = decToHex(id),
                    parsed = cfg.btsName ? me.parseName(cfg.btsName) : null,
                    location = (parsed && parsed.location) || cfg.tempName || cfg.batteryName || (t('Tag') + ' ' + hex),
                    statusText = raw.BatteryVoltageStatus ? raw.BatteryVoltageStatus.value : '',
                    level = raw.BatteryVoltage ? raw.BatteryVoltage.value : '',
                    lastTs = Math.max(
                        raw.BatteryVoltageStatus ? raw.BatteryVoltageStatus.ts : 0,
                        raw.BatteryVoltage ? raw.BatteryVoltage.ts : 0,
                        raw.Temperature ? raw.Temperature.ts : 0
                    ),
                    status;

                if (!statusText && !level) {
                    status = 'nodata';
                } else if (statusText) {
                    status = me.classify(statusText);
                } else {
                    status = 'unknown';
                }

                rows.push({
                    key: obj.agentId + ':' + id,
                    agentId: obj.agentId,
                    objectName: obj.name,
                    region: obj.region,
                    tagDec: id,
                    hexId: hex,
                    location: location,
                    btsName: cfg.btsName || '',
                    configured: !!cfg.btsName,
                    statusText: statusText,
                    level: level,
                    temperature: raw.Temperature ? raw.Temperature.value : '',
                    humidity: raw.Humidity ? raw.Humidity.value : '',
                    rssi: raw.RSSI ? raw.RSSI.value : '',
                    status: status,
                    lastTs: lastTs,
                    rank: STATUS[status].rank
                });
            });

            return rows;
        },

        finish: function (rows, errors) {
            this.store.loadData(rows);
            this.loading = false;
            this.lastLoad = new Date();
            this.fireEvent('load', this, rows, errors);
        },

        // Objects that have at least one battery row (others are hidden from the UI).
        getTagObjects: function () {
            var seen = {};
            this.store.each(function (r) { seen[r.get('agentId')] = true; });
            return Ext.Array.filter(this.objects, function (o) { return seen[o.agentId]; });
        },

        getSummary: function () {
            var s = { total: 0, ok: 0, alert: 0, nodata: 0, unknown: 0, objects: 0, objectsWithData: 0 },
                all = {},
                reporting = {};

            this.store.each(function (r) {
                s.total += 1;
                s[r.get('status')] = (s[r.get('status')] || 0) + 1;
                all[r.get('agentId')] = true;
                if (r.get('status') !== 'nodata') {
                    reporting[r.get('agentId')] = true;
                }
            });
            s.objects = Ext.Object.getSize(all);
            s.objectsWithData = Ext.Object.getSize(reporting);
            return s;
        },

        getObjectStats: function (agentId) {
            var stats = { total: 0, ok: 0, worst: 'nodata', any: false };
            this.store.each(function (r) {
                if (r.get('agentId') !== agentId) {
                    return;
                }
                stats.total += 1;
                if (r.get('status') === 'ok') {
                    stats.ok += 1;
                }
                if (!stats.any || STATUS[r.get('status')].rank > STATUS[stats.worst].rank) {
                    stats.worst = r.get('status');
                }
                stats.any = true;
            });
            return stats;
        }
    });

    // ---------------------------------------------------------------------
    // Shared grid for battery sensor rows.
    // ---------------------------------------------------------------------
    Ext.define('Store.battery_status.view.SensorGrid', {
        extend: 'Ext.grid.Panel',
        cls: 'bs-grid',
        hideObject: false,

        initComponent: function () {
            var me = this;

            me.viewConfig = Ext.apply({
                emptyText: '<div class="bs-empty">' + enc(me.emptyMessage || t('No battery tags found')) + '</div>',
                deferEmptyText: false,
                stripeRows: true
            }, me.viewConfig || {});

            me.columns = [
                {
                    text: t('Battery'), dataIndex: 'rank', width: 150,
                    renderer: function (v, meta, r) {
                        return statusPill(r.get('status'), r.get('statusText'));
                    }
                },
                {
                    text: t('Sensor location'), dataIndex: 'location', flex: 2, minWidth: 200,
                    renderer: function (v, meta, r) {
                        var tip = r.get('btsName') || t('No BTS- sensor configured for this tag');
                        meta.tdAttr = 'data-qtip="' + enc(enc(tip)) + '"';
                        return enc(v) + (r.get('configured') ? '' : ' <span class="bs-muted">(' + enc(t('no BTS sensor')) + ')</span>');
                    }
                },
                { text: t('Sensor ID (hex)'), dataIndex: 'hexId', width: 115, renderer: enc },
                { text: t('Object'), dataIndex: 'objectName', flex: 1, minWidth: 150, hidden: me.hideObject, renderer: enc },
                { text: t('Region'), dataIndex: 'region', width: 115, hidden: me.hideObject, renderer: enc },
                { text: t('Level'), dataIndex: 'level', width: 65, align: 'center', renderer: enc },
                {
                    text: t('Temp.'), dataIndex: 'temperature', width: 75, align: 'right',
                    renderer: function (v) { return v === '' ? '' : enc(v) + ' &deg;C'; }
                },
                {
                    text: t('Humidity'), dataIndex: 'humidity', width: 80, align: 'right',
                    renderer: function (v) { return v === '' ? '' : enc(v) + ' %'; }
                },
                { text: t('RSSI'), dataIndex: 'rssi', width: 60, align: 'right', hidden: true, renderer: enc },
                { text: t('Last packet'), dataIndex: 'lastTs', width: 185, renderer: formatTs },
                { text: t('Tag ID (dec)'), dataIndex: 'tagDec', width: 110, hidden: true, renderer: enc }
            ];

            me.callParent(arguments);
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

            me.overviewStore = Ext.create('Ext.data.ChainedStore', {
                source: source,
                groupField: 'objectName'
            });

            me.alertStore = Ext.create('Ext.data.ChainedStore', {
                source: source,
                filters: [function (r) { return r.get('status') !== 'ok'; }]
            });

            me.summaryCmp = Ext.create('Ext.Component', {
                cls: 'bs-summary',
                html: '<div class="bs-muted">' + enc(t('Loading battery status...')) + '</div>'
            });

            me.overviewGrid = Ext.create('Store.battery_status.view.SensorGrid', {
                store: me.overviewStore,
                flex: 1,
                features: [{
                    ftype: 'grouping',
                    groupHeaderTpl: '{name} ({rows.length})',
                    startCollapsed: false
                }],
                emptyMessage: t('No battery sensor values reported yet')
            });

            me.alertsGrid = Ext.create('Store.battery_status.view.SensorGrid', {
                store: me.alertStore,
                emptyMessage: t('All reporting battery sensors are Normal')
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
                            emptyText: t('Search sensor or object'),
                            width: 260,
                            listeners: {
                                change: me.onSearch,
                                buffer: 250,
                                scope: me
                            }
                        },
                        '->',
                        { xtype: 'tbtext', itemId: 'lastLoad', cls: 'bs-muted', text: '' },
                        { text: t('Refresh'), iconCls: 'fa fa-rotate', handler: me.refresh, scope: me }
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
                beforeload: function () { me.setLoading(t('Loading battery status...')); },
                progress: function (data, done, total) {
                    me.setLoading(t('Reading sensor tracing') + ' ' + done + ' / ' + total);
                },
                load: me.onDataLoad,
                loaderror: function (data, msg) {
                    me.setLoading(false);
                    me.summaryCmp.setHtml('<div class="bs-error">' + enc(msg) + '</div>');
                },
                scope: me
            });
        },

        refresh: function () {
            if (this.module) {
                this.module.reload();
            }
        },

        onSearch: function (field, value) {
            var text = String(value || '').toLowerCase();
            this.overviewStore.clearFilter();
            if (text) {
                this.overviewStore.filterBy(function (r) {
                    return (r.get('sensorName') + ' ' + r.get('objectName') + ' ' + r.get('region') + ' ' + r.get('hexId'))
                        .toLowerCase().indexOf(text) !== -1;
                });
            }
        },

        onDataLoad: function (data, rows, errors) {
            var s = data.getSummary(),
                lastLoad = this.down('#lastLoad'),
                alertsTab = this.child('#alerts');

            this.setLoading(false);

            this.summaryCmp.setHtml(
                '<div class="bs-tiles">' +
                this.tile(t('Battery tags'), s.total, 'bs-tile-total', 'fa fa-car-battery') +
                this.tile(t('Normal'), s.ok || 0, 'bs-tile-ok', STATUS.ok.icon) +
                this.tile(t('Alert'), s.alert || 0, 'bs-tile-alert', STATUS.alert.icon) +
                this.tile(t('Not reporting'), s.nodata || 0, 'bs-tile-nodata', STATUS.nodata.icon) +
                this.tile(t('Unknown'), s.unknown || 0, 'bs-tile-unknown', STATUS.unknown.icon) +
                this.tile(t('Objects reporting'), s.objectsWithData + ' / ' + s.objects, 'bs-tile-total', 'fa fa-warehouse') +
                '</div>' +
                '<div class="bs-muted bs-note">' + enc(t('Values come from the raw packets (Sensors Tracing) of the last')) + ' ' +
                enc(data.hoursBack) + ' ' + enc(t('hours. "Not reporting" means the tag sent no battery value in that window.')) + '</div>' +
                (errors ? '<div class="bs-error">' + enc(t('Some objects could not be loaded. Showing partial data.')) + '</div>' : '')
            );

            if (lastLoad) {
                lastLoad.setText(t('Updated') + ' ' + Ext.Date.format(new Date(), 'H:i:s'));
            }

            if (alertsTab && alertsTab.tab) {
                var alerts = (s.alert || 0) + (s.nodata || 0) + (s.unknown || 0);
                alertsTab.setTitle(t('Alerts') + (alerts ? ' <span class="bs-count">' + alerts + '</span>' : ''));
            }

            this.syncObjectTabs();
        },

        tile: function (label, value, cls, icon) {
            return '<div class="bs-tile ' + cls + '"><i class="' + icon + '"></i>' +
                '<div class="bs-tile-value">' + enc(value) + '</div>' +
                '<div class="bs-tile-label">' + enc(label) + '</div></div>';
        },

        // One nested tab per object (warehouse), with its own filtered grid.
        syncObjectTabs: function () {
            var me = this,
                data = me.service,
                activeId = me.objectTabs.getActiveTab() ? me.objectTabs.getActiveTab().agentId : null;

            me.tabsById = me.tabsById || {};

            Ext.Array.each(data.getTagObjects(), function (obj) {
                var stats = data.getObjectStats(obj.agentId),
                    info = data.objectInfo[obj.agentId] || {},
                    gridTitle = enc(obj.path.concat([obj.name]).join('  /  ')) +
                        (info.gatewayBattery ? '  <span class="bs-muted">' + enc(t('Gateway battery')) + ': ' + enc(info.gatewayBattery) + ' V' +
                            (info.powerVoltage ? ', ' + enc(t('power')) + ': ' + enc(info.powerVoltage) + ' V' : '') + '</span>' : ''),
                    tab = me.tabsById[obj.agentId],
                    title = enc(obj.name) + ' <span class="bs-tab-badge ' + STATUS[stats.worst].cls + '">' +
                        (stats.total ? stats.ok + '/' + stats.total : '-') + '</span>';

                if (!tab) {
                    tab = me.objectTabs.add({
                        title: title,
                        agentId: obj.agentId,
                        layout: 'fit',
                        items: [Ext.create('Store.battery_status.view.SensorGrid', {
                            hideObject: true,
                            store: Ext.create('Ext.data.ChainedStore', {
                                source: data.store,
                                filters: [{ property: 'agentId', value: obj.agentId, exactMatch: true }]
                            }),
                            title: gridTitle,
                            emptyMessage: t('No battery tags on this object.')
                        })]
                    });
                    me.tabsById[obj.agentId] = tab;
                } else {
                    tab.setTitle(title);
                    tab.down('grid').setTitle(gridTitle);
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
            var field = this.down('textfield');
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
                tooltip: t('Refresh'),
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
                var key = obj.region || t('Ungrouped');
                if (!regions[key]) {
                    regions[key] = [];
                    order.push(key);
                }
                var stats = data.getObjectStats(obj.agentId);
                regions[key].push({
                    name: obj.name,
                    agentId: obj.agentId,
                    leaf: true,
                    iconCls: STATUS[stats.worst].icon + ' ' + STATUS[stats.worst].cls,
                    badge: stats.total ? stats.ok + '/' + stats.total : '-',
                    worst: stats.worst
                });
            });

            Ext.Array.each(order, function (key) {
                var worst = 'nodata', ok = 0, total = 0;
                Ext.Array.each(regions[key], function (leaf) {
                    var stats = data.getObjectStats(leaf.agentId);
                    ok += stats.ok;
                    total += stats.total;
                    if (worst === 'nodata' || STATUS[stats.worst].rank > STATUS[worst].rank) {
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
                    fieldLabel: t('Auto refresh'),
                    labelAlign: 'top',
                    anchor: '100%',
                    editable: false,
                    queryMode: 'local',
                    store: [[0, t('Off')], [120, t('Every 2 minutes')], [300, t('Every 5 minutes')], [600, t('Every 10 minutes')], [1800, t('Every 30 minutes')]],
                    value: me.module.settings.refreshSec,
                    listeners: {
                        change: function (field, value) { me.module.applySettings({ refreshSec: value }); }
                    }
                },
                {
                    xtype: 'numberfield',
                    fieldLabel: t('Look back in sensor tracing (hours, 1-24)'),
                    labelAlign: 'top',
                    anchor: '100%',
                    minValue: 1,
                    maxValue: 24,
                    allowDecimals: false,
                    value: me.module.settings.hoursBack,
                    listeners: {
                        change: { fn: function (field, value) {
                            if (field.isValid()) {
                                me.module.applySettings({ hoursBack: Number(value) || 3 });
                            }
                        }, buffer: 600 }
                    }
                },
                {
                    xtype: 'component',
                    cls: 'bs-muted',
                    html: '<p>' + enc(t('The battery status is read from the raw packet data (triggerdata) of Sensors Tracing, field BatteryVoltageStatus<tag id>. A tag that sent no battery value in this window is shown as Not reporting.')) + '</p>' +
                        '<p>' + enc(t('Names come from the BTS-<hex>-<name> sensors, or from the Temperature<tag id> sensor when no BTS sensor exists.')) + '</p>' +
                        '<p>' + enc(t('Settings are saved in this browser only.')) + '</p>'
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
                console.log('battery_status: initModule()');
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
            var me = this;

            me.loadStyles();

            me.settings = {
                refreshSec: storageGet('refreshSec', 300),
                hoursBack: storageGet('hoursBack', 3)
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

            // Load once the panel is first shown, then keep refreshing on the chosen interval.
            me.mainPanel.on('activate', me.ensureLoaded, me);
            me.mainPanel.on('show', me.ensureLoaded, me);
            me.navTab.on('activate', me.ensureLoaded, me);
            Ext.defer(me.ensureLoaded, 1500, me);

            me.startTimer();
        },

        ensureLoaded: function () {
            if (!this.service.lastLoad && !this.service.loading) {
                this.reload();
            }
        },

        reload: function () {
            this.service.load(this.settings.hoursBack);
        },

        applySettings: function (changes) {
            Ext.apply(this.settings, changes);
            Ext.Object.each(changes, function (key, value) { storageSet(key, value); });
            if (Ext.isDefined(changes.refreshSec)) {
                this.startTimer();
            }
            if (Ext.isDefined(changes.hoursBack) && this.service.lastLoad) {
                this.reload();
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

        loadStyles: function () {
            if (document.getElementById('battery_status-css')) {
                return;
            }
            var css = document.createElement('link');
            css.id = 'battery_status-css';
            css.setAttribute('rel', 'stylesheet');
            css.setAttribute('type', 'text/css');
            css.setAttribute('href', '/store/battery_status/style.css');
            document.head.appendChild(css);
        }
    });
}());
