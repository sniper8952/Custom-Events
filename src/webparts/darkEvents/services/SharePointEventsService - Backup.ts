import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPropertyFieldSite } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { EventSourceMode } from '../IDarkEventsWebPartProps';
import { IEventItem, ISharePointEventRow } from '../models/IEventItem';

interface IRowsResponse { value: ISharePointEventRow[]; }
interface IListRow { Id: string; Title: string; }
interface IListsResponse { value: IListRow[]; }
interface IWebResponse { Title: string; }

export class SharePointEventsService {
  public constructor(private readonly context: WebPartContext) {}

  public async getEvents(
    sourceMode: EventSourceMode,
    selectedListId: string,
    selectedSites: IPropertyFieldSite[],
    count: number,
    showPast: boolean
  ): Promise<IEventItem[]> {
    const currentWeb = this.context.pageContext.web.absoluteUrl;
    let sources: Array<{ webUrl: string; listId: string }> = [];

    if (sourceMode === 'currentList') {
      if (!selectedListId) throw new Error('Bitte in den Webpart-Eigenschaften eine Ereignisliste auswählen.');
      sources = [{ webUrl: currentWeb, listId: selectedListId }];
    } else if (sourceMode === 'currentSite') {
      sources = await this.getEventLists(currentWeb);
    } else {
      const sites = selectedSites || [];
      if (sites.length === 0) throw new Error('Bitte in den Webpart-Eigenschaften mindestens eine Site auswählen.');
      const nested = await Promise.all(sites.map(site => this.getEventLists(site.url || '')));
      sources = nested.reduce((all, part) => all.concat(part), []);
    }

    if (sources.length === 0) throw new Error('In der gewählten Quelle wurde keine Ereignisliste gefunden.');
    const groups = await Promise.all(sources.map(async source => {
      const web = await this.getJson<IWebResponse>(`${source.webUrl}/_api/web?$select=Title`);
      return this.getEventsFromList(source.webUrl, source.listId, web.Title || source.webUrl, count, showPast);
    }));
    return groups.reduce((all, part) => all.concat(part), [])
      .sort((a, b) => a.start.getTime() - b.start.getTime())
      .slice(0, Math.max(1, Math.min(count, 20)));
  }

  private async getEventLists(webUrl: string): Promise<Array<{ webUrl: string; listId: string }>> {
    if (!webUrl) return [];
    const url = `${webUrl}/_api/web/lists?$select=Id,Title&$filter=BaseTemplate eq 106 and Hidden eq false`;
    const data = await this.getJson<IListsResponse>(url);
    return data.value.map(list => ({ webUrl, listId: list.Id }));
  }

  private async getEventsFromList(webUrl: string, listId: string, siteName: string, count: number, showPast: boolean): Promise<IEventItem[]> {
    const now = new Date().toISOString();
    const filter = showPast ? '' : `&$filter=EndDate ge datetime'${now}'`;
    const select =
  'ID,Title,Description,EventDate,EndDate,Location,Category,' +
  'fAllDayEvent,BannerUrl,FileRef';
    const url = `${webUrl}/_api/web/lists(guid'${listId}')/items?$select=${select}${filter}` +
      `&$orderby=EventDate asc&$top=${Math.max(1, Math.min(count, 20))}`;
    const rows = await this.getJson<IRowsResponse>(url);
    return rows.value.map(row => this.mapRow(row, webUrl, listId, siteName));
  }

  private async getJson<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url, SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata.metadata=none' } }
    );
    if (!response.ok) {
      const body = await response.text();
      throw new Error(`SharePoint REST ${response.status}: ${body || response.statusText}`);
    }
    return response.json() as Promise<T>;
  }

private mapRow(
  row: ISharePointEventRow,
  webUrl: string,
  listId: string,
  siteName: string
): IEventItem {

  const imageUrl =
    this.getImageUrl(row.BannerUrl, webUrl);

  const eventUrl =
    `${webUrl}/_layouts/15/Event.aspx` +
    `?ListGuid=${encodeURIComponent(listId)}` +
    `&ItemId=${row.ID}`;

return {
  id: row.ID,
  title: row.Title || 'Ohne Titel',
  description: row.Description || '',
  start: new Date(row.EventDate),
  end: new Date(row.EndDate),
  location: row.Location || '',
  category: row.Category || '',
  siteName,
  allDay:
    row.fAllDayEvent === true ||
    row.fAllDayEvent === 1,
  imageUrl,
  eventUrl
};
}

  private getImageUrl(value: ISharePointEventRow['BannerUrl'], webUrl: string): string {
    if (!value) return '';
    if (typeof value === 'object' && value.Url) return this.toAbsoluteUrl(value.Url, webUrl);
    if (typeof value !== 'string') return '';
    const trimmed = value.trim();
    if (!trimmed) return '';
    try {
      const parsed = JSON.parse(trimmed) as { serverRelativeUrl?: string; url?: string };
      return this.toAbsoluteUrl(parsed.serverRelativeUrl || parsed.url || '', webUrl);
    } catch {
      return this.toAbsoluteUrl(trimmed.split(',')[0].trim(), webUrl);
    }
  }

  private toAbsoluteUrl(url: string, webUrl: string): string {
    if (!url) return '';
    if (/^https?:\/\//i.test(url)) return url;
    return `${new URL(webUrl).origin}${url.startsWith('/') ? '' : '/'}${url}`;
  }
}
