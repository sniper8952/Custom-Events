"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("react");
var SharePointEventsService_1 = require("../services/SharePointEventsService");
var EventCard_1 = tslib_1.__importDefault(require("./EventCard"));
var DarkEvents_module_scss_1 = tslib_1.__importDefault(require("./DarkEvents.module.scss"));
var DarkEvents = function (props) {
    var _a;
    var _b = (0, react_1.useState)([]), events = _b[0], setEvents = _b[1];
    var _c = (0, react_1.useState)(true), loading = _c[0], setLoading = _c[1];
    var _d = (0, react_1.useState)(''), error = _d[0], setError = _d[1];
    var _e = (0, react_1.useState)(0), activePage = _e[0], setActivePage = _e[1];
    var _f = (0, react_1.useState)(1), pageCount = _f[0], setPageCount = _f[1];
    var trackRef = (0, react_1.useRef)(null);
    var service = (0, react_1.useMemo)(function () { return new SharePointEventsService_1.SharePointEventsService(props.context); }, [props.context]);
    (0, react_1.useEffect)(function () {
        var active = true;
        setLoading(true);
        setError('');
        service.getEvents(props.sourceMode, props.selectedListId, props.selectedSites, props.numberOfEvents, props.showPastEvents)
            .then(function (items) { if (active) {
            setEvents(items);
            setLoading(false);
        } })
            .catch(function (reason) {
            if (active) {
                setError(reason instanceof Error ? reason.message : 'Ereignisse konnten nicht geladen werden.');
                setLoading(false);
            }
        });
        return function () { active = false; };
    }, [service, props.sourceMode, props.selectedListId, props.selectedSites, props.numberOfEvents, props.showPastEvents]);
    var updatePagination = function () {
        var track = trackRef.current;
        if (!track)
            return;
        var pages = Math.max(1, Math.ceil(track.scrollWidth / track.clientWidth));
        setPageCount(pages);
        var page = Math.min(pages - 1, Math.round(track.scrollLeft / Math.max(1, track.clientWidth)));
        setActivePage(page);
    };
    (0, react_1.useEffect)(function () {
        var track = trackRef.current;
        if (!track)
            return;
        updatePagination();
        var onResize = function () { return updatePagination(); };
        window.addEventListener('resize', onResize);
        return function () { return window.removeEventListener('resize', onResize); };
    }, [events]);
    var goToPage = function (page) {
        var track = trackRef.current;
        if (!track)
            return;
        var nextPage = Math.max(0, Math.min(pageCount - 1, page));
        track.scrollTo({ left: nextPage * track.clientWidth, behavior: 'smooth' });
        setActivePage(nextPage);
    };
    return (React.createElement("section", { className: DarkEvents_module_scss_1.default.root, style: (_a = {},
            _a['--accent'] = props.accentColor,
            _a['--root-bg'] = props.backgroundColor,
            _a['--card-bg'] = props.cardBackgroundColor,
            _a['--text'] = props.textColor,
            _a['--secondary-text'] = props.secondaryTextColor,
            _a['--card-size'] = "".concat(Math.max(50, Math.min(200, Number(props.cardSizePercent) || 100)) / 100),
            _a['--image-height'] = "".concat(Math.max(50, Math.min(200, Number(props.imageHeightPercent) || 100)) / 100),
            _a) },
        React.createElement("div", { className: DarkEvents_module_scss_1.default.header },
            React.createElement("h2", null, "Events")),
        loading && React.createElement("div", { className: DarkEvents_module_scss_1.default.status }, "Ereignisse werden geladen \u2026"),
        !loading && error && React.createElement("div", { className: DarkEvents_module_scss_1.default.error },
            React.createElement("strong", null, "Fehler beim Laden"),
            React.createElement("br", null),
            error),
        !loading && !error && events.length === 0 && React.createElement("div", { className: DarkEvents_module_scss_1.default.status }, "Keine passenden Ereignisse gefunden."),
        !loading && !error && events.length > 0 && (React.createElement("div", { className: DarkEvents_module_scss_1.default.carousel },
            React.createElement("button", { type: "button", className: "".concat(DarkEvents_module_scss_1.default.arrow, " ").concat(DarkEvents_module_scss_1.default.arrowLeft), onClick: function () { return goToPage(activePage - 1); }, "aria-label": "Vorherige Ereignisse", disabled: activePage === 0 }, "\u2039"),
            React.createElement("div", { className: DarkEvents_module_scss_1.default.track, ref: trackRef, onScroll: updatePagination }, events.map(function (event) { return React.createElement(EventCard_1.default, { key: "".concat(event.siteName, "-").concat(event.id), event: event, accentColor: props.accentColor, fallbackImageUrl: props.fallbackImageUrl, showSiteName: props.showSiteName, showCategory: props.showCategory, imageFocusX: Math.max(0, Math.min(100, Number(props.imageFocusXPercent) || 50)), imageFocusY: Math.max(0, Math.min(100, Number(props.imageFocusYPercent) || 50)) }); })),
            React.createElement("button", { type: "button", className: "".concat(DarkEvents_module_scss_1.default.arrow, " ").concat(DarkEvents_module_scss_1.default.arrowRight), onClick: function () { return goToPage(activePage + 1); }, "aria-label": "N\u00E4chste Ereignisse", disabled: activePage >= pageCount - 1 }, "\u203A"),
            pageCount > 1 && React.createElement("div", { className: DarkEvents_module_scss_1.default.dots, "aria-label": "Ereignisseiten" }, Array.from({ length: pageCount }, function (_unused, index) { return (React.createElement("button", { key: index, type: "button", className: index === activePage ? DarkEvents_module_scss_1.default.dotActive : DarkEvents_module_scss_1.default.dot, onClick: function () { return goToPage(index); }, "aria-label": "Seite ".concat(index + 1), "aria-current": index === activePage ? 'true' : undefined })); }))))));
};
exports.default = DarkEvents;
//# sourceMappingURL=DarkEvents%20-%20Backup.js.map