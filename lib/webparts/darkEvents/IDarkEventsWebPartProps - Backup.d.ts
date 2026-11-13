import { IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
export type EventSourceMode = 'currentList' | 'currentSite' | 'selectedSites';
export interface IDarkEventsWebPartProps {
    sourceMode: EventSourceMode;
    selectedListId: string;
    selectedSites: IPropertyFieldSite[];
    numberOfEvents: number;
    showPastEvents: boolean;
    showSiteName: boolean;
    showCategory: boolean;
    cardSizePercent: string;
    imageHeightPercent: string;
    imageFocusXPercent: string;
    imageFocusYPercent: string;
    backgroundColor: string;
    cardBackgroundColor: string;
    textColor: string;
    secondaryTextColor: string;
    accentColor: string;
    fallbackImageUrl: string;
}
//# sourceMappingURL=IDarkEventsWebPartProps%20-%20Backup.d.ts.map