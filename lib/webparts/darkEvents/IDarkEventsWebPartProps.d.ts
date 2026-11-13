import { IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
export type EventSourceMode = 'currentList' | 'currentSite' | 'selectedSites';
export interface IDarkEventsWebPartProps {
    sourceMode: EventSourceMode;
    selectedListId: string;
    selectedSites: IPropertyFieldSite[];
    numberOfEvents: number;
    showPastEvents: boolean;
    showWebPartTitle: boolean;
    webPartTitle: string;
    showSiteName: boolean;
    showCategory: boolean;
    cardSizePercent: string;
    imageHeightPercent: string;
    imageFocusXPercent: string;
    imageFocusYPercent: string;
    backgroundColor: string;
    cardBackgroundColor: string;
    webPartTitleColor: string;
    addEventColor: string;
    addEventDisabledColor: string;
    categoryColor: string;
    siteTitleColor: string;
    eventTitleColor: string;
    eventDateColor: string;
    locationColor: string;
    calendarIconColor: string;
    paginationColor: string;
    arrowBackgroundColor: string;
    fallbackImageUrl: string;
}
//# sourceMappingURL=IDarkEventsWebPartProps.d.ts.map