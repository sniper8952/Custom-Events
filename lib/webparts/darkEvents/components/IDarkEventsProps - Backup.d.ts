import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { EventSourceMode } from '../IDarkEventsWebPartProps';
export interface IDarkEventsProps {
    context: WebPartContext;
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
//# sourceMappingURL=IDarkEventsProps%20-%20Backup.d.ts.map