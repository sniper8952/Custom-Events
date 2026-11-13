import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { EventSourceMode } from '../IDarkEventsWebPartProps';
import { IEventItem } from '../models/IEventItem';
export declare class SharePointEventsService {
    private readonly context;
    constructor(context: WebPartContext);
    getEvents(sourceMode: EventSourceMode, selectedListId: string, selectedSites: IPropertyFieldSite[], count: number, showPast: boolean): Promise<IEventItem[]>;
    getAddEventUrl(webUrl: string, listId: string): string;
    getSiteAddEventUrl(webUrl: string): Promise<string>;
    private getEventLists;
    private getEventsFromList;
    private getJson;
    private mapRow;
    private getImageUrl;
    private toAbsoluteUrl;
}
//# sourceMappingURL=SharePointEventsService.d.ts.map