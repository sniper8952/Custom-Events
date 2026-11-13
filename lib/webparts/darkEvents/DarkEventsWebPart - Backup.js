import { __awaiter, __extends, __generator } from "tslib";
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider } from '@microsoft/sp-component-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import DarkEvents from './components/DarkEvents';
var DarkEventsWebPart = /** @class */ (function (_super) {
    __extends(DarkEventsWebPart, _super);
    function DarkEventsWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.handleThemeChanged = function (_args) { return _this.render(); };
        return _this;
    }
    DarkEventsWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
                        this.themeProvider.themeChangedEvent.add(this, this.handleThemeChanged);
                        return [2 /*return*/];
                }
            });
        });
    };
    DarkEventsWebPart.prototype.render = function () {
        if (!this.root)
            this.root = createRoot(this.domElement);
        this.root.render(React.createElement(DarkEvents, {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            context: this.context,
            sourceMode: this.properties.sourceMode || 'currentList',
            selectedListId: this.properties.selectedListId || '',
            selectedSites: this.properties.selectedSites || [],
            numberOfEvents: this.properties.numberOfEvents || 8,
            showPastEvents: !!this.properties.showPastEvents,
            showSiteName: this.properties.showSiteName !== false,
            showCategory: this.properties.showCategory !== false,
            cardSizePercent: this.properties.cardSizePercent || '100',
            imageHeightPercent: this.properties.imageHeightPercent || '100',
            imageFocusXPercent: this.properties.imageFocusXPercent || '50',
            imageFocusYPercent: this.properties.imageFocusYPercent || '50',
            backgroundColor: this.properties.backgroundColor || '#000000',
            cardBackgroundColor: this.properties.cardBackgroundColor || '#111111',
            textColor: this.properties.textColor || '#ffffff',
            secondaryTextColor: this.properties.secondaryTextColor || '#d6d6d6',
            accentColor: this.properties.accentColor || '#b7c800',
            fallbackImageUrl: this.properties.fallbackImageUrl || ''
        }));
    };
    DarkEventsWebPart.prototype.onDispose = function () {
        var _a, _b;
        (_a = this.themeProvider) === null || _a === void 0 ? void 0 : _a.themeChangedEvent.remove(this, this.handleThemeChanged);
        (_b = this.root) === null || _b === void 0 ? void 0 : _b.unmount();
        this.root = undefined;
    };
    DarkEventsWebPart.prototype.getPropertyPaneConfiguration = function () {
        var sourceFields = [
            PropertyPaneDropdown('sourceMode', {
                label: 'Quelle',
                options: [
                    { key: 'currentList', text: 'Liste der Ereignisse auf dieser Site' },
                    { key: 'currentSite', text: 'Diese Site' },
                    { key: 'selectedSites', text: 'Sites auswählen' }
                ]
            })
        ];
        if ((this.properties.sourceMode || 'currentList') === 'currentList') {
            sourceFields.push(PropertyFieldListPicker('selectedListId', {
                label: 'Ereignisliste', selectedList: this.properties.selectedListId,
                baseTemplate: 106, includeHidden: false, orderBy: PropertyFieldListPickerOrderBy.Title,
                disabled: false, multiSelect: false,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties, // eslint-disable-next-line @typescript-eslint/no-explicit-any
                context: this.context,
                onGetErrorMessage: undefined, deferredValidationTime: 0,
                key: 'eventListPicker'
            }));
        }
        if (this.properties.sourceMode === 'selectedSites') {
            sourceFields.push(PropertyFieldSitePicker('selectedSites', {
                label: 'Sites auswählen', initialSites: this.properties.selectedSites || [],
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                context: this.context, deferredValidationTime: 500, multiSelect: true,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties, key: 'eventSitePicker'
            }));
        }
        return { pages: [{ header: { description: 'Quelle, Darstellung und Farben' }, groups: [
                        { groupName: 'Datenquelle', groupFields: sourceFields },
                        { groupName: 'Anzeige', groupFields: [
                                PropertyPaneSlider('numberOfEvents', { label: 'Anzahl Ereignisse', min: 1, max: 20, step: 1 }),
                                PropertyPaneToggle('showPastEvents', { label: 'Vergangene Ereignisse anzeigen', onText: 'Ja', offText: 'Nein' }),
                                PropertyPaneToggle('showSiteName', { label: 'SharePoint-Sitename anzeigen', onText: 'Ein', offText: 'Aus' }),
                                PropertyPaneToggle('showCategory', { label: 'Kategorie anzeigen', onText: 'Ein', offText: 'Aus' }),
                                PropertyPaneTextField('cardSizePercent', { label: 'Kachelgrösse in % (50–200)' }),
                                PropertyPaneTextField('imageHeightPercent', { label: 'Bildhöhe in % (50–200)' }),
                                PropertyPaneTextField('imageFocusXPercent', { label: 'Banner-Fokus horizontal in % (0–100)' }),
                                PropertyPaneTextField('imageFocusYPercent', { label: 'Banner-Fokus vertikal in % (0–100)' })
                            ] },
                        { groupName: 'Farben', groupFields: [
                                PropertyPaneTextField('backgroundColor', { label: 'Webpart-Hintergrund' }),
                                PropertyPaneTextField('cardBackgroundColor', { label: 'Karten-Hintergrund' }),
                                PropertyPaneTextField('textColor', { label: 'Schriftfarbe' }),
                                PropertyPaneTextField('secondaryTextColor', { label: 'Sekundäre Schriftfarbe' }),
                                PropertyPaneTextField('accentColor', { label: 'Akzentfarbe' }),
                                PropertyPaneTextField('fallbackImageUrl', { label: 'Fallback-Bild-URL (optional)' })
                            ] }
                    ] }] };
    };
    return DarkEventsWebPart;
}(BaseClientSideWebPart));
export default DarkEventsWebPart;
//# sourceMappingURL=DarkEventsWebPart%20-%20Backup.js.map