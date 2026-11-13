"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SharePointEventsService = void 0;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
var SharePointEventsService = /** @class */ (function () {
    function SharePointEventsService(context) {
        this.context = context;
    }
    /*
     * =====================================================
     * EVENTS LADEN
     * =====================================================
     */
    SharePointEventsService.prototype.getEvents = function (sourceMode, selectedListId, selectedSites, count, showPast) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var currentWeb, sources, sites, nested, groups;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        currentWeb = this.context.pageContext.web.absoluteUrl;
                        sources = [];
                        if (!(sourceMode === 'currentList')) return [3 /*break*/, 1];
                        if (!selectedListId) {
                            throw new Error('Bitte in den Webpart-Eigenschaften eine Ereignisliste auswählen.');
                        }
                        sources = [
                            {
                                webUrl: currentWeb,
                                listId: selectedListId
                            }
                        ];
                        return [3 /*break*/, 5];
                    case 1:
                        if (!(sourceMode === 'currentSite')) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.getEventLists(currentWeb)];
                    case 2:
                        sources =
                            _a.sent();
                        return [3 /*break*/, 5];
                    case 3:
                        sites = selectedSites || [];
                        if (sites.length === 0) {
                            throw new Error('Bitte in den Webpart-Eigenschaften mindestens eine Site auswählen.');
                        }
                        return [4 /*yield*/, Promise.all(sites.map(function (site) {
                                return _this.getEventLists(site.url || '');
                            }))];
                    case 4:
                        nested = _a.sent();
                        sources =
                            nested.reduce(function (all, part) {
                                return all.concat(part);
                            }, []);
                        _a.label = 5;
                    case 5:
                        /*
                         * Keine Ereignisliste gefunden
                         */
                        if (sources.length === 0) {
                            throw new Error('In der gewählten Quelle wurde keine Ereignisliste gefunden.');
                        }
                        return [4 /*yield*/, Promise.all(sources.map(function (source) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                                var web;
                                return tslib_1.__generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.getJson("".concat(source.webUrl, "/_api/web?$select=Title"))];
                                        case 1:
                                            web = _a.sent();
                                            return [2 /*return*/, this.getEventsFromList(source.webUrl, source.listId, web.Title || source.webUrl, count, showPast)];
                                    }
                                });
                            }); }))];
                    case 6:
                        groups = _a.sent();
                        /*
                         * Events aller Sites zusammenführen,
                         * chronologisch sortieren und begrenzen.
                         */
                        return [2 /*return*/, groups
                                .reduce(function (all, part) {
                                return all.concat(part);
                            }, [])
                                .sort(function (a, b) {
                                return a.start.getTime() -
                                    b.start.getTime();
                            })
                                .slice(0, Math.max(1, Math.min(count, 20)))];
                }
            });
        });
    };
    /*
     * =====================================================
     * URL FÜR "EREIGNIS HINZUFÜGEN"
     *
     * webUrl = Site, auf der das Event erstellt werden soll
     * listId = Ereignisliste dieser Site
     *
     * Ergebnis:
     * /_layouts/15/Event.aspx?ListGuid=...&Mode=Edit
     * =====================================================
     */
    SharePointEventsService.prototype.getAddEventUrl = function (webUrl, listId) {
        if (!webUrl ||
            !listId) {
            return '';
        }
        /*
         * Eventuelle abschliessende Slashes entfernen.
         */
        var normalizedWebUrl = webUrl.replace(/\/$/, '');
        return ("".concat(normalizedWebUrl, "/_layouts/15/Event.aspx") +
            "?ListGuid=".concat(encodeURIComponent(listId)) +
            "&Mode=Edit");
    };
    /*
     * =====================================================
     * EREIGNISLISTE DER AKTUELLEN SITE ERMITTELN
     *
     * Diese Methode ist für "+ Ereignis hinzufügen".
     *
     * DarkEvents.tsx übergibt:
     *
     * props.context.pageContext.web.absoluteUrl
     *
     * Dadurch wird NICHT die ausgewählte Anzeige-Site
     * verwendet, sondern immer die Site, auf welcher
     * das Webpart eingebettet ist.
     * =====================================================
     */
    SharePointEventsService.prototype.getSiteAddEventUrl = function (webUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var eventLists, eventList;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!webUrl) {
                            return [2 /*return*/, ''];
                        }
                        return [4 /*yield*/, this.getEventLists(webUrl)];
                    case 1:
                        eventLists = _a.sent();
                        if (eventLists.length === 0) {
                            return [2 /*return*/, ''];
                        }
                        eventList = eventLists[0];
                        return [2 /*return*/, this.getAddEventUrl(eventList.webUrl, eventList.listId)];
                }
            });
        });
    };
    /*
     * =====================================================
     * EREIGNISLISTEN EINER SHAREPOINT-SITE SUCHEN
     *
     * BaseTemplate 106 = Ereignisliste
     * =====================================================
     */
    SharePointEventsService.prototype.getEventLists = function (webUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, data;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!webUrl) {
                            return [2 /*return*/, []];
                        }
                        url = "".concat(webUrl, "/_api/web/lists") +
                            "?$select=Id,Title" +
                            "&$filter=BaseTemplate eq 106 and Hidden eq false";
                        return [4 /*yield*/, this.getJson(url)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, data.value.map(function (list) { return ({
                                webUrl: webUrl,
                                listId: list.Id
                            }); })];
                }
            });
        });
    };
    /*
     * =====================================================
     * EVENTS EINER BESTIMMTEN LISTE LADEN
     * =====================================================
     */
    SharePointEventsService.prototype.getEventsFromList = function (webUrl, listId, siteName, count, showPast) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var now, filter, select, url, rows;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        now = new Date().toISOString();
                        filter = showPast
                            ? ''
                            : "&$filter=EndDate ge datetime'".concat(now, "'");
                        select = 'ID,' +
                            'Title,' +
                            'Description,' +
                            'EventDate,' +
                            'EndDate,' +
                            'Location,' +
                            'Category,' +
                            'fAllDayEvent,' +
                            'BannerUrl,' +
                            'FileRef';
                        url = "".concat(webUrl, "/_api/web/lists(guid'").concat(listId, "')/items") +
                            "?$select=".concat(select) +
                            "".concat(filter) +
                            "&$orderby=EventDate asc" +
                            "&$top=".concat(Math.max(1, Math.min(count, 20)));
                        return [4 /*yield*/, this.getJson(url)];
                    case 1:
                        rows = _a.sent();
                        return [2 /*return*/, rows.value.map(function (row) {
                                return _this.mapRow(row, webUrl, listId, siteName);
                            })];
                }
            });
        });
    };
    /*
     * =====================================================
     * SHAREPOINT REST REQUEST
     * =====================================================
     */
    SharePointEventsService.prototype.getJson = function (url) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, body;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                            headers: {
                                Accept: 'application/json;odata.metadata=none'
                            }
                        })];
                    case 1:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, response.text()];
                    case 2:
                        body = _a.sent();
                        throw new Error("SharePoint REST ".concat(response.status, ": ") +
                            "".concat(body || response.statusText));
                    case 3: return [2 /*return*/, response.json()];
                }
            });
        });
    };
    /*
     * =====================================================
     * SHAREPOINT-EVENT IN UNSER MODELL UMWANDELN
     * =====================================================
     */
    SharePointEventsService.prototype.mapRow = function (row, webUrl, listId, siteName) {
        /*
         * Event-Banner auflösen
         */
        var imageUrl = this.getImageUrl(row.BannerUrl, webUrl);
        /*
         * URL zur SharePoint-Eventdetailseite.
         *
         * Klick auf die Event-Kachel navigiert dorthin.
         */
        var eventUrl = "".concat(webUrl, "/_layouts/15/Event.aspx") +
            "?ListGuid=".concat(encodeURIComponent(listId)) +
            "&ItemId=".concat(row.ID);
        return {
            id: row.ID,
            title: row.Title ||
                'Ohne Titel',
            description: row.Description ||
                '',
            start: new Date(row.EventDate),
            end: new Date(row.EndDate),
            location: row.Location ||
                '',
            category: row.Category ||
                '',
            siteName: siteName,
            allDay: row.fAllDayEvent === true ||
                row.fAllDayEvent === 1,
            imageUrl: imageUrl,
            eventUrl: eventUrl
        };
    };
    /*
     * =====================================================
     * EVENT-BANNER ERMITTELN
     * =====================================================
     */
    SharePointEventsService.prototype.getImageUrl = function (value, webUrl) {
        /*
         * Kein Bild
         */
        if (!value) {
            return '';
        }
        /*
         * SharePoint kann das URL-Feld
         * als Objekt liefern.
         */
        if (typeof value === 'object' &&
            value.Url) {
            return this.toAbsoluteUrl(value.Url, webUrl);
        }
        if (typeof value !== 'string') {
            return '';
        }
        var trimmed = value.trim();
        if (!trimmed) {
            return '';
        }
        /*
         * SharePoint kann Bannerinformationen
         * als JSON-String zurückgeben.
         */
        try {
            var parsed = JSON.parse(trimmed);
            return this.toAbsoluteUrl(parsed.serverRelativeUrl ||
                parsed.url ||
                '', webUrl);
        }
        catch (_a) {
            /*
             * Fallback für normale URLs
             * beziehungsweise SharePoint-Werte.
             */
            return this.toAbsoluteUrl(trimmed
                .split(',')[0]
                .trim(), webUrl);
        }
    };
    /*
     * =====================================================
     * SERVERRELATIVE URL
     * IN ABSOLUTE URL UMWANDELN
     * =====================================================
     */
    SharePointEventsService.prototype.toAbsoluteUrl = function (url, webUrl) {
        if (!url) {
            return '';
        }
        /*
         * Bereits eine vollständige URL
         */
        if (/^https?:\/\//i.test(url)) {
            return url;
        }
        /*
         * Aus /sites/... eine vollständige
         * SharePoint-URL erzeugen.
         */
        return ("".concat(new URL(webUrl).origin) +
            "".concat(url.startsWith('/') ? '' : '/') +
            "".concat(url));
    };
    return SharePointEventsService;
}());
exports.SharePointEventsService = SharePointEventsService;
//# sourceMappingURL=SharePointEventsService.js.map