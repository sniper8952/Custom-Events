import { __awaiter, __generator } from "tslib";
import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { SharePointEventsService } from '../services/SharePointEventsService';
import EventCard from './EventCard';
import styles from './DarkEvents.module.scss';
var DarkEvents = function (props) {
    var _a;
    var _b = useState([]), events = _b[0], setEvents = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    var _d = useState(''), error = _d[0], setError = _d[1];
    var _e = useState(0), activePage = _e[0], setActivePage = _e[1];
    var _f = useState(1), pageCount = _f[0], setPageCount = _f[1];
    var _g = useState(''), addEventUrl = _g[0], setAddEventUrl = _g[1];
    var trackRef = useRef(null);
    var service = useMemo(function () {
        return new SharePointEventsService(props.context);
    }, [props.context]);
    /*
     * ===================================================
     * EVENTS UND ADD-EVENT-URL LADEN
     * ===================================================
     */
    useEffect(function () {
        var active = true;
        setLoading(true);
        setError('');
        setAddEventUrl('');
        var loadAddEventUrl = function () { return __awaiter(void 0, void 0, void 0, function () {
            var currentWebUrl, url, reason_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        currentWebUrl = props.context.pageContext.web.absoluteUrl;
                        return [4 /*yield*/, service.getSiteAddEventUrl(currentWebUrl)];
                    case 1:
                        url = _a.sent();
                        if (active) {
                            setAddEventUrl(url);
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        reason_1 = _a.sent();
                        console.error('Die URL für "Ereignis hinzufügen" ' +
                            'konnte nicht ermittelt werden:', reason_1);
                        if (active) {
                            setAddEventUrl('');
                        }
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        void loadAddEventUrl();
        service
            .getEvents(props.sourceMode, props.selectedListId, props.selectedSites, props.numberOfEvents, props.showPastEvents)
            .then(function (items) {
            if (!active) {
                return;
            }
            setEvents(items);
            setActivePage(0);
            setLoading(false);
        })
            .catch(function (reason) {
            if (!active) {
                return;
            }
            setEvents([]);
            setError(reason instanceof Error
                ? reason.message
                : 'Ereignisse konnten nicht geladen werden.');
            setLoading(false);
        });
        return function () {
            active = false;
        };
    }, [
        service,
        props.context,
        props.sourceMode,
        props.selectedListId,
        props.selectedSites,
        props.numberOfEvents,
        props.showPastEvents
    ]);
    /*
     * ===================================================
     * PAGINIERUNG
     * ===================================================
     */
    var updatePagination = function () {
        var track = trackRef.current;
        if (!track) {
            return;
        }
        var clientWidth = Math.max(1, track.clientWidth);
        var pages = Math.max(1, Math.ceil(track.scrollWidth /
            clientWidth));
        var page = Math.min(pages - 1, Math.round(track.scrollLeft /
            clientWidth));
        setPageCount(pages);
        setActivePage(page);
    };
    useEffect(function () {
        var track = trackRef.current;
        if (!track) {
            return;
        }
        var animationFrame = window.requestAnimationFrame(updatePagination);
        var onResize = function () {
            updatePagination();
        };
        window.addEventListener('resize', onResize);
        return function () {
            window.cancelAnimationFrame(animationFrame);
            window.removeEventListener('resize', onResize);
        };
    }, [events]);
    /*
     * ===================================================
     * NAVIGATION
     * ===================================================
     */
    var goToPage = function (page) {
        var track = trackRef.current;
        if (!track) {
            return;
        }
        var nextPage = Math.max(0, Math.min(pageCount - 1, page));
        track.scrollTo({
            left: nextPage *
                track.clientWidth,
            behavior: 'smooth'
        });
        setActivePage(nextPage);
    };
    /*
     * ===================================================
     * DARSTELLUNGSWERTE
     * ===================================================
     */
    var cardSize = Math.max(50, Math.min(200, Number(props.cardSizePercent) || 100)) / 100;
    var imageHeight = Math.max(50, Math.min(200, Number(props.imageHeightPercent) || 100)) / 100;
    var imageFocusX = Math.max(0, Math.min(100, Number(props.imageFocusXPercent) || 50));
    var imageFocusY = Math.max(0, Math.min(100, Number(props.imageFocusYPercent) || 50));
    /*
     * ===================================================
     * RENDERING
     * ===================================================
     */
    return (React.createElement("section", { className: styles.root, style: (_a = {},
            _a['--root-bg'] = props.backgroundColor,
            _a['--card-bg'] = props.cardBackgroundColor,
            _a['--text'] = '#FFFFFF',
            _a['--accent'] = props.siteTitleColor,
            _a['--webpart-title-color'] = props.webPartTitleColor,
            _a['--add-event-color'] = props.addEventColor,
            _a['--add-event-disabled-color'] = props.addEventDisabledColor,
            _a['--category-color'] = props.categoryColor,
            _a['--site-title-color'] = props.siteTitleColor,
            _a['--event-title-color'] = props.eventTitleColor,
            _a['--event-date-color'] = props.eventDateColor,
            _a['--location-color'] = props.locationColor,
            _a['--calendar-icon-color'] = props.calendarIconColor,
            _a['--pagination-color'] = props.paginationColor,
            _a['--arrow-background-color'] = props.arrowBackgroundColor,
            _a['--card-size'] = "".concat(cardSize),
            _a['--image-height'] = "".concat(imageHeight),
            _a) },
        props.showWebPartTitle && (React.createElement("div", { className: styles.header },
            React.createElement("h2", null, props.webPartTitle))),
        addEventUrl && (React.createElement("div", { className: styles.addEventRow }, props.canAddEvents
            ? (React.createElement("a", { className: styles.addEvent, href: addEventUrl },
                React.createElement("span", { className: styles.addEventIcon, "aria-hidden": "true" }, "+"),
                React.createElement("span", null, "Ereignis hinzuf\u00FCgen")))
            : (React.createElement("span", { className: "".concat(styles.addEvent, " ") +
                    "".concat(styles.addEventDisabled), "aria-disabled": "true", title: "Sie verf\u00FCgen nicht \u00FCber die Berechtigung, Ereignisse hinzuzuf\u00FCgen." },
                React.createElement("span", { className: styles.addEventIcon, "aria-hidden": "true" }, "+"),
                React.createElement("span", null, "Ereignis hinzuf\u00FCgen"))))),
        loading && (React.createElement("div", { className: styles.status }, "Ereignisse werden geladen \u2026")),
        !loading &&
            error && (React.createElement("div", { className: styles.error },
            React.createElement("strong", null, "Fehler beim Laden"),
            React.createElement("br", null),
            error)),
        !loading &&
            !error &&
            events.length === 0 && (React.createElement("div", { className: styles.status }, "Keine passenden Ereignisse gefunden.")),
        !loading &&
            !error &&
            events.length > 0 && (React.createElement("div", { className: styles.carousel },
            React.createElement("button", { type: "button", className: "".concat(styles.arrow, " ") +
                    "".concat(styles.arrowLeft), onClick: function () {
                    return goToPage(activePage - 1);
                }, "aria-label": "Vorherige Ereignisse", disabled: activePage === 0 },
                React.createElement("span", { className: styles.arrowIcon, "aria-hidden": "true" }, "\u2039")),
            React.createElement("div", { className: styles.track, ref: trackRef, onScroll: updatePagination }, events.map(function (event) { return (React.createElement(EventCard, { key: "".concat(event.siteName, "-").concat(event.id), event: event, fallbackImageUrl: props.fallbackImageUrl, showSiteName: props.showSiteName, showCategory: props.showCategory, imageFocusX: imageFocusX, imageFocusY: imageFocusY })); })),
            React.createElement("button", { type: "button", className: "".concat(styles.arrow, " ") +
                    "".concat(styles.arrowRight), onClick: function () {
                    return goToPage(activePage + 1);
                }, "aria-label": "N\u00E4chste Ereignisse", disabled: activePage >=
                    pageCount - 1 },
                React.createElement("span", { className: styles.arrowIcon, "aria-hidden": "true" }, "\u203A")),
            pageCount > 1 && (React.createElement("div", { className: styles.dots, "aria-label": "Ereignisseiten" }, Array.from({
                length: pageCount
            }, function (_unused, index) { return (React.createElement("button", { key: index, type: "button", className: index === activePage
                    ? styles.dotActive
                    : styles.dot, onClick: function () {
                    return goToPage(index);
                }, "aria-label": "Seite ".concat(index + 1), "aria-current": index === activePage
                    ? 'true'
                    : undefined })); })))))));
};
export default DarkEvents;
//# sourceMappingURL=DarkEvents.js.map