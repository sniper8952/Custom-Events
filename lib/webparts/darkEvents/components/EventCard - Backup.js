import { __spreadArray } from "tslib";
import * as React from 'react';
import styles from './DarkEvents.module.scss';
var dateFormatter = new Intl.DateTimeFormat('de-CH', { weekday: 'short', day: '2-digit', month: 'short' });
var timeFormatter = new Intl.DateTimeFormat('de-CH', { hour: '2-digit', minute: '2-digit' });
var monthFormatter = new Intl.DateTimeFormat('de-CH', { month: 'short' });
var CalendarPlusIcon = function () { return (React.createElement("svg", { className: styles.calendarPlusIcon, viewBox: "0 0 24 24", "aria-hidden": "true", focusable: "false" },
    React.createElement("path", { d: "M7 2v3M17 2v3M3.5 8.5h17M5.5 4h13a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Z" }),
    React.createElement("path", { d: "M12 11.5v6M9 14.5h6" }))); };
var escapeIcs = function (value) { return value
    .replace(/\\/g, '\\\\')
    .replace(/\r?\n/g, '\\n')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;'); };
var htmlToText = function (value) {
    var element = document.createElement('div');
    element.innerHTML = value || '';
    return element.textContent || element.innerText || '';
};
var toIcsDate = function (value) { return value.toISOString()
    .replace(/[-:]/g, '')
    .replace(/\.\d{3}Z$/, 'Z'); };
var safeFileName = function (value) {
    var cleaned = value.replace(/[\\/:*?"<>|]/g, '-').trim();
    return "".concat(cleaned || 'event', ".ics");
};
var downloadIcs = function (event) {
    var plainDescription = htmlToText(event.description).trim();
    var htmlDescription = event.description || '';
    var escapeHtmlForIcs = function (value) {
        return value
            .replace(/\\/g, '\\\\')
            .replace(/\r?\n/g, '')
            .replace(/,/g, '\\,')
            .replace(/;/g, '\\;');
    };
    var lines = __spreadArray(__spreadArray([
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Agrisano//Dark Events//DE',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'BEGIN:VEVENT',
        "UID:agrisano-".concat(event.siteName, "-").concat(event.id, "@sharepoint"),
        "DTSTAMP:".concat(toIcsDate(new Date())),
        "DTSTART:".concat(toIcsDate(event.start)),
        "DTEND:".concat(toIcsDate(event.end)),
        "SUMMARY:".concat(escapeIcs(event.title))
    ], (event.category
        ? ["CATEGORIES:".concat(escapeIcs(event.category))]
        : []), true), [
        "LOCATION:".concat(escapeIcs(event.location || '')),
        "DESCRIPTION:".concat(escapeIcs(plainDescription)),
        "X-ALT-DESC;FMTTYPE=text/html:".concat(escapeHtmlForIcs(htmlDescription)),
        'END:VEVENT',
        'END:VCALENDAR'
    ], false);
    var blob = new Blob([lines.join('\r\n')], {
        type: 'text/calendar;charset=utf-8'
    });
    var url = URL.createObjectURL(blob);
    var anchor = document.createElement('a');
    anchor.href = url;
    anchor.download =
        safeFileName(event.title);
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
};
var EventCard = function (_a) {
    var event = _a.event, accentColor = _a.accentColor, fallbackImageUrl = _a.fallbackImageUrl, showSiteName = _a.showSiteName, showCategory = _a.showCategory, imageFocusX = _a.imageFocusX, imageFocusY = _a.imageFocusY;
    var sameDay = event.start.toDateString() === event.end.toDateString();
    var image = event.imageUrl || fallbackImageUrl;
    var timeText = event.allDay ? 'Ganztägig' : sameDay
        ? "".concat(dateFormatter.format(event.start), ", ").concat(timeFormatter.format(event.start))
        : "".concat(dateFormatter.format(event.start), " \u2013 ").concat(dateFormatter.format(event.end));
    var dayText = sameDay ? String(event.start.getDate()) : "".concat(event.start.getDate(), "\u2013").concat(event.end.getDate());
    var openEvent = function () { window.location.href = event.eventUrl; };
    var onCardKeyDown = function (keyboardEvent) {
        if (keyboardEvent.key === 'Enter' || keyboardEvent.key === ' ') {
            keyboardEvent.preventDefault();
            openEvent();
        }
    };
    var onCalendarClick = function (mouseEvent) {
        mouseEvent.stopPropagation();
        downloadIcs(event);
    };
    return (React.createElement("article", { className: styles.card, role: "link", tabIndex: 0, "aria-label": "".concat(event.title, " \u00F6ffnen"), onClick: openEvent, onKeyDown: onCardKeyDown },
        React.createElement("div", { className: styles.imageLink },
            image ? React.createElement("img", { className: styles.image, src: image, alt: "", loading: "lazy", style: { objectPosition: "".concat(imageFocusX, "% ").concat(imageFocusY, "%") } }) : React.createElement("div", { className: styles.imageFallback }),
            React.createElement("div", { className: styles.imageShade }),
            React.createElement("div", { className: styles.dateBadge },
                React.createElement("span", null, monthFormatter.format(event.start).replace('.', '').toUpperCase()),
                React.createElement("strong", null, dayText))),
        React.createElement("div", { className: styles.content },
            showCategory && event.category && React.createElement("div", { className: styles.category, style: { color: accentColor } }, event.category),
            showSiteName && event.siteName && React.createElement("div", { className: styles.siteName }, event.siteName),
            React.createElement("div", { className: styles.title }, event.title),
            React.createElement("div", { className: styles.spacer }),
            React.createElement("div", { className: styles.meta }, timeText),
            event.location && React.createElement("div", { className: styles.location }, event.location)),
        React.createElement("div", { className: styles.footer },
            React.createElement("button", { type: "button", className: styles.calendarLink, title: "Als ICS herunterladen", "aria-label": "".concat(event.title, " als ICS herunterladen"), onClick: onCalendarClick },
                React.createElement(CalendarPlusIcon, null)))));
};
export default EventCard;
//# sourceMappingURL=EventCard%20-%20Backup.js.map