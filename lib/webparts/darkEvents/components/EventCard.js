import { __spreadArray } from "tslib";
import * as React from 'react';
import styles from './DarkEvents.module.scss';
/*
 * =====================================================
 * DATUMSFORMATIERUNG
 * =====================================================
 */
var dateFormatter = new Intl.DateTimeFormat('de-CH', {
    weekday: 'short',
    day: '2-digit',
    month: 'short'
});
var timeFormatter = new Intl.DateTimeFormat('de-CH', {
    hour: '2-digit',
    minute: '2-digit'
});
var monthFormatter = new Intl.DateTimeFormat('de-CH', {
    month: 'short'
});
/*
 * =====================================================
 * KALENDER + ICON
 * =====================================================
 */
var CalendarPlusIcon = function () { return (React.createElement("svg", { className: styles.calendarPlusIcon, viewBox: "0 0 24 24", "aria-hidden": "true", focusable: "false" },
    React.createElement("path", { d: "\n          M7 2v3\n          M17 2v3\n          M3.5 8.5h17\n          M5.5 4h13\n          a2 2 0 0 1 2 2\n          v13\n          a2 2 0 0 1-2 2\n          h-13\n          a2 2 0 0 1-2-2\n          V6\n          a2 2 0 0 1 2-2Z\n        " }),
    React.createElement("path", { d: "\n          M12 11.5v6\n          M9 14.5h6\n        " }))); };
/*
 * =====================================================
 * ICS-WERTE ESCAPEN
 * =====================================================
 */
var escapeIcs = function (value) {
    return value
        .replace(/\\/g, '\\\\')
        .replace(/\r?\n/g, '\\n')
        .replace(/,/g, '\\,')
        .replace(/;/g, '\\;');
};
/*
 * =====================================================
 * HTML → TEXT
 *
 * Wird für DESCRIPTION in der ICS verwendet.
 * =====================================================
 */
var htmlToText = function (value) {
    var element = document.createElement('div');
    element.innerHTML =
        value || '';
    return (element.textContent ||
        element.innerText ||
        '');
};
/*
 * =====================================================
 * DATUM FÜR ICS
 * =====================================================
 */
var toIcsDate = function (value) {
    return value
        .toISOString()
        .replace(/[-:]/g, '')
        .replace(/\.\d{3}Z$/, 'Z');
};
/*
 * =====================================================
 * DATEINAME FÜR ICS
 * =====================================================
 */
var safeFileName = function (value) {
    var cleaned = value
        .replace(/[\\/:*?"<>|]/g, '-')
        .trim();
    return ("".concat(cleaned || 'event', ".ics"));
};
/*
 * =====================================================
 * HTML FÜR ICS ESCAPEN
 * =====================================================
 */
var escapeHtmlForIcs = function (value) {
    return value
        .replace(/\\/g, '\\\\')
        .replace(/\r?\n/g, '')
        .replace(/,/g, '\\,')
        .replace(/;/g, '\\;');
};
/*
 * =====================================================
 * ICS HERUNTERLADEN
 * =====================================================
 */
var downloadIcs = function (event) {
    /*
     * Beschreibung als normaler Text.
     */
    var plainDescription = htmlToText(event.description).trim();
    /*
     * Originales HTML für Outlook.
     */
    var htmlDescription = event.description || '';
    var lines = __spreadArray(__spreadArray([
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Agrisano//Dark Events//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'BEGIN:VEVENT',
        /*
         * Eindeutige ID
         */
        "UID:agrisano-".concat(event.siteName, "-").concat(event.id, "@sharepoint"),
        /*
         * Erstellungszeit
         */
        "DTSTAMP:".concat(toIcsDate(new Date())),
        /*
         * Eventzeit
         */
        "DTSTART:".concat(toIcsDate(event.start)),
        "DTEND:".concat(toIcsDate(event.end)),
        /*
         * Eventtitel
         */
        "SUMMARY:".concat(escapeIcs(event.title))
    ], (event.category
        ? [
            "CATEGORIES:".concat(escapeIcs(event.category))
        ]
        : []), true), [
        /*
         * Standort
         */
        "LOCATION:".concat(escapeIcs(event.location || '')),
        /*
         * Beschreibung als Text
         */
        "DESCRIPTION:".concat(escapeIcs(plainDescription)),
        /*
         * Beschreibung als HTML.
         *
         * Outlook kann dadurch die Formatierung
         * aus "Über dieses Ereignis" übernehmen.
         */
        "X-ALT-DESC;FMTTYPE=text/html:".concat(escapeHtmlForIcs(htmlDescription)),
        'END:VEVENT',
        'END:VCALENDAR'
    ], false);
    var blob = new Blob([
        lines.join('\r\n')
    ], {
        type: 'text/calendar;charset=utf-8'
    });
    var url = URL.createObjectURL(blob);
    var anchor = document.createElement('a');
    anchor.href =
        url;
    anchor.download =
        safeFileName(event.title);
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
};
/*
 * =====================================================
 * EVENT CARD
 * =====================================================
 */
var EventCard = function (_a) {
    /*
     * ===================================================
     * MEHRTÄGIGE / GANZTÄGIGE EVENTS
     *
     * Bei SharePoint ist das Enddatum eines ganztägigen
     * Events exklusiv.
     *
     * Beispiel:
     *
     * Start:
     * 26.11.2026
     *
     * SharePoint-Enddatum:
     * 01.12.2026
     *
     * Sichtbare Anzeige:
     * 26.–30.11.2026
     * ===================================================
     */
    var event = _a.event, fallbackImageUrl = _a.fallbackImageUrl, showSiteName = _a.showSiteName, showCategory = _a.showCategory, imageFocusX = _a.imageFocusX, imageFocusY = _a.imageFocusY;
    var displayEnd = new Date(event.end);
    if (event.allDay) {
        displayEnd.setDate(displayEnd.getDate() - 1);
    }
    var sameDay = event.start.toDateString() ===
        displayEnd.toDateString();
    /*
     * ===================================================
     * BANNER
     * ===================================================
     */
    var image = event.imageUrl ||
        fallbackImageUrl;
    /*
     * ===================================================
     * DATUM / ZEIT UNTEN
     * ===================================================
     */
    var timeText = event.allDay
        ? sameDay
            ? "".concat(dateFormatter.format(event.start), ", Ganzt\u00E4gig")
            : "".concat(dateFormatter.format(event.start), " \u2013 ").concat(dateFormatter.format(displayEnd))
        : sameDay
            ? "".concat(dateFormatter.format(event.start), ", ").concat(timeFormatter.format(event.start))
            : "".concat(dateFormatter.format(event.start), " \u2013 ").concat(dateFormatter.format(displayEnd));
    /*
     * ===================================================
     * GROSSES DATUM AUF DEM BANNER
     * ===================================================
     */
    var dayText = sameDay
        ? String(event.start.getDate())
        : "".concat(event.start.getDate(), "\u2013").concat(displayEnd.getDate());
    /*
     * ===================================================
     * EVENT ÖFFNEN
     * ===================================================
     */
    var openEvent = function () {
        window.location.href =
            event.eventUrl;
    };
    /*
     * ===================================================
     * TASTATUR-NAVIGATION
     * ===================================================
     */
    var onCardKeyDown = function (keyboardEvent) {
        if (keyboardEvent.key ===
            'Enter' ||
            keyboardEvent.key ===
                ' ') {
            keyboardEvent.preventDefault();
            openEvent();
        }
    };
    /*
     * ===================================================
     * KALENDER-KLICK
     *
     * stopPropagation verhindert, dass gleichzeitig
     * die Eventdetailseite geöffnet wird.
     * ===================================================
     */
    var onCalendarClick = function (mouseEvent) {
        mouseEvent.stopPropagation();
        downloadIcs(event);
    };
    /*
     * ===================================================
     * RENDERING
     * ===================================================
     */
    return (React.createElement("article", { className: styles.card, role: "link", tabIndex: 0, "aria-label": "".concat(event.title, " \u00F6ffnen"), onClick: openEvent, onKeyDown: onCardKeyDown },
        React.createElement("div", { className: styles.imageLink },
            image
                ? (React.createElement("img", { className: styles.image, src: image, alt: "", loading: "lazy", style: {
                        objectPosition: "".concat(imageFocusX, "% ").concat(imageFocusY, "%")
                    } }))
                : (React.createElement("div", { className: styles.imageFallback })),
            React.createElement("div", { className: styles.imageShade }),
            React.createElement("div", { className: styles.dateBadge },
                React.createElement("span", null, monthFormatter
                    .format(event.start)
                    .replace('.', '')
                    .toUpperCase()),
                React.createElement("strong", null, dayText))),
        React.createElement("div", { className: styles.content },
            React.createElement("div", { className: "".concat(styles.category, " ") +
                    "".concat(!showCategory
                        ? styles.hiddenRow
                        : '') }, event.category ||
                '\u00A0'),
            React.createElement("div", { className: "".concat(styles.siteName, " ") +
                    "".concat(!showSiteName
                        ? styles.hiddenRow
                        : '') }, event.siteName ||
                '\u00A0'),
            React.createElement("div", { className: styles.title }, event.title),
            React.createElement("div", { className: styles.spacer }),
            React.createElement("div", { className: styles.meta }, timeText),
            React.createElement("div", { className: styles.location }, event.location ||
                '\u00A0')),
        React.createElement("div", { className: styles.footer },
            React.createElement("button", { type: "button", className: styles.calendarLink, title: "Als ICS herunterladen", "aria-label": "".concat(event.title, " als ICS herunterladen"), onClick: onCalendarClick },
                React.createElement(CalendarPlusIcon, null)))));
};
export default EventCard;
//# sourceMappingURL=EventCard.js.map