"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var client_1 = require("react-dom/client");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_component_base_1 = require("@microsoft/sp-component-base");
var sp_page_context_1 = require("@microsoft/sp-page-context");
var PropertyFieldListPicker_1 = require("@pnp/spfx-property-controls/lib/PropertyFieldListPicker");
var PropertyFieldSitePicker_1 = require("@pnp/spfx-property-controls/lib/PropertyFieldSitePicker");
var DarkEvents_1 = tslib_1.__importDefault(require("./components/DarkEvents"));
var DarkEventsWebPart = /** @class */ (function (_super) {
    tslib_1.__extends(DarkEventsWebPart, _super);
    function DarkEventsWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.canManageWebPart = false;
        _this.canAddEvents = false;
        _this.handleThemeChanged = function (_args) {
            _this.render();
        };
        return _this;
    }
    DarkEventsWebPart.prototype.onInit = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        this.canManageWebPart =
                            this.context.pageContext.web.permissions.hasPermission(sp_page_context_1.SPPermission.manageWeb);
                        this.canAddEvents =
                            this.context.pageContext.web.permissions.hasPermission(sp_page_context_1.SPPermission.addListItems);
                        this.themeProvider =
                            this.context.serviceScope.consume(sp_component_base_1.ThemeProvider.serviceKey);
                        this.themeProvider.themeChangedEvent.add(this, this.handleThemeChanged);
                        return [2 /*return*/];
                }
            });
        });
    };
    DarkEventsWebPart.prototype.render = function () {
        if (!this.root) {
            this.root =
                (0, client_1.createRoot)(this.domElement);
        }
        this.root.render(React.createElement(DarkEvents_1.default, {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            context: this.context,
            sourceMode: this.properties.sourceMode ||
                'currentList',
            selectedListId: this.properties.selectedListId ||
                '',
            selectedSites: this.properties.selectedSites ||
                [],
            canAddEvents: this.canAddEvents,
            numberOfEvents: this.properties.numberOfEvents ||
                8,
            showPastEvents: !!this.properties.showPastEvents,
            showWebPartTitle: this.properties.showWebPartTitle !== false,
            webPartTitle: this.properties.webPartTitle ||
                'Events',
            showSiteName: this.properties.showSiteName !== false,
            showCategory: this.properties.showCategory !== false,
            cardSizePercent: this.properties.cardSizePercent ||
                '100',
            imageHeightPercent: this.properties.imageHeightPercent ||
                '100',
            imageFocusXPercent: this.properties.imageFocusXPercent ||
                '50',
            imageFocusYPercent: this.properties.imageFocusYPercent ||
                '50',
            backgroundColor: this.properties.backgroundColor ||
                'transparent',
            cardBackgroundColor: this.properties.cardBackgroundColor ||
                '#1F1F1F',
            webPartTitleColor: this.properties.webPartTitleColor ||
                '#1F1F1F',
            addEventColor: this.properties.addEventColor ||
                '#000000',
            addEventDisabledColor: this.properties.addEventDisabledColor ||
                '#D3D3D3',
            categoryColor: this.properties.categoryColor ||
                '#FFFFFF',
            siteTitleColor: this.properties.siteTitleColor ||
                '#9EA500',
            eventTitleColor: this.properties.eventTitleColor ||
                '#FFFFFF',
            eventDateColor: this.properties.eventDateColor ||
                '#FFFFFF',
            locationColor: this.properties.locationColor ||
                '#9EA500',
            calendarIconColor: this.properties.calendarIconColor ||
                '#FFFFFF',
            paginationColor: this.properties.paginationColor ||
                '#000000',
            arrowBackgroundColor: this.properties.arrowBackgroundColor ||
                '#282828',
            fallbackImageUrl: this.properties.fallbackImageUrl ||
                ''
        }));
    };
    DarkEventsWebPart.prototype.onDispose = function () {
        var _a, _b;
        (_a = this.themeProvider) === null || _a === void 0 ? void 0 : _a.themeChangedEvent.remove(this, this.handleThemeChanged);
        (_b = this.root) === null || _b === void 0 ? void 0 : _b.unmount();
        this.root =
            undefined;
    };
    DarkEventsWebPart.prototype.getPropertyPaneConfiguration = function () {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        var sourceFields = [
            (0, sp_property_pane_1.PropertyPaneDropdown)('sourceMode', {
                label: 'Quelle',
                options: [
                    {
                        key: 'currentList',
                        text: 'Liste der Ereignisse auf dieser Site'
                    },
                    {
                        key: 'currentSite',
                        text: 'Diese Site'
                    },
                    {
                        key: 'selectedSites',
                        text: 'Sites auswählen'
                    }
                ]
            })
        ];
        if ((this.properties.sourceMode ||
            'currentList') ===
            'currentList') {
            sourceFields.push((0, PropertyFieldListPicker_1.PropertyFieldListPicker)('selectedListId', {
                label: 'Ereignisliste',
                selectedList: this.properties.selectedListId,
                baseTemplate: 106,
                includeHidden: false,
                orderBy: PropertyFieldListPicker_1.PropertyFieldListPickerOrderBy.Title,
                disabled: false,
                multiSelect: false,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                context: this.context,
                onGetErrorMessage: undefined,
                deferredValidationTime: 0,
                key: 'eventListPicker'
            }));
        }
        if (this.properties.sourceMode ===
            'selectedSites') {
            sourceFields.push((0, PropertyFieldSitePicker_1.PropertyFieldSitePicker)('selectedSites', {
                label: 'Sites auswählen',
                initialSites: this.properties.selectedSites ||
                    [],
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                context: this.context,
                deferredValidationTime: 500,
                multiSelect: true,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                key: 'eventSitePicker'
            }));
        }
        var groups = [
            {
                groupName: 'Datenquelle',
                groupFields: sourceFields
            }
        ];
        if (this.canManageWebPart) {
            /*
             * =================================================
             * ANZEIGE
             * =================================================
             */
            groups.push({
                groupName: 'Anzeige',
                groupFields: [
                    (0, sp_property_pane_1.PropertyPaneToggle)('showWebPartTitle', {
                        label: 'Webpart-Titel anzeigen',
                        onText: 'Ein',
                        offText: 'Aus'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('webPartTitle', {
                        label: 'Webpart-Titel',
                        description: 'Beispiel: Events, Veranstaltungen oder Termine',
                        disabled: this.properties.showWebPartTitle === false
                    }),
                    (0, sp_property_pane_1.PropertyPaneSlider)('numberOfEvents', {
                        label: 'Anzahl Ereignisse',
                        min: 1,
                        max: 20,
                        step: 1
                    }),
                    (0, sp_property_pane_1.PropertyPaneToggle)('showPastEvents', {
                        label: 'Vergangene Ereignisse anzeigen',
                        onText: 'Ja',
                        offText: 'Nein'
                    }),
                    (0, sp_property_pane_1.PropertyPaneToggle)('showSiteName', {
                        label: 'SharePoint-Sitename anzeigen',
                        onText: 'Ein',
                        offText: 'Aus'
                    }),
                    (0, sp_property_pane_1.PropertyPaneToggle)('showCategory', {
                        label: 'Kategorie anzeigen',
                        onText: 'Ein',
                        offText: 'Aus'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('cardSizePercent', {
                        label: 'Kachelgrösse in % (50–200)'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('imageHeightPercent', {
                        label: 'Bildhöhe in % (50–200)'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('imageFocusXPercent', {
                        label: 'Banner-Fokus horizontal in % (0–100)'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('imageFocusYPercent', {
                        label: 'Banner-Fokus vertikal in % (0–100)'
                    })
                ]
            });
            /*
             * =================================================
             * FARBEN
             * =================================================
             */
            groups.push({
                groupName: 'Farben',
                groupFields: [
                    (0, sp_property_pane_1.PropertyPaneTextField)('backgroundColor', {
                        label: 'Webpart-Hintergrund',
                        description: 'Standard: transparent'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('cardBackgroundColor', {
                        label: 'Karten-Hintergrund',
                        description: 'Standard: #1F1F1F'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('webPartTitleColor', {
                        label: 'Webpart-Titel',
                        description: 'Standard: #1F1F1F'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('addEventColor', {
                        label: '"Ereignis hinzufügen"',
                        description: 'Standard: #000000'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('addEventDisabledColor', {
                        label: '"Ereignis hinzufügen" deaktiviert',
                        description: 'Standard: #D3D3D3'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('categoryColor', {
                        label: 'Kategorie',
                        description: 'Standard: #FFFFFF'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('siteTitleColor', {
                        label: 'Seitentitel',
                        description: 'Standard: #9EA500'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('eventTitleColor', {
                        label: 'Eventtitel',
                        description: 'Standard: #FFFFFF'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('eventDateColor', {
                        label: 'Eventdatum',
                        description: 'Standard: #FFFFFF'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('locationColor', {
                        label: 'Standort',
                        description: 'Standard: #9EA500'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('calendarIconColor', {
                        label: 'Kalender +',
                        description: 'Standard: #FFFFFF'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('paginationColor', {
                        label: 'Navigationspunkte',
                        description: 'Standard: #000000'
                    }),
                    /*
                     * NEU
                     */
                    (0, sp_property_pane_1.PropertyPaneTextField)('arrowBackgroundColor', {
                        label: 'Pfeil-Hintergrund',
                        description: 'Standard: #282828'
                    }),
                    (0, sp_property_pane_1.PropertyPaneTextField)('fallbackImageUrl', {
                        label: 'Fallback-Bild-URL (optional)'
                    })
                ]
            });
        }
        return {
            pages: [
                {
                    header: {
                        description: this.canManageWebPart
                            ? 'Quelle, Anzeige und Farben'
                            : 'Ereignisquelle'
                    },
                    groups: groups
                }
            ]
        };
    };
    return DarkEventsWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = DarkEventsWebPart;
//# sourceMappingURL=DarkEventsWebPart.js.map