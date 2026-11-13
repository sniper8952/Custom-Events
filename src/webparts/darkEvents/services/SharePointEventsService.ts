import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  IPropertyFieldSite
} from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';

import {
  EventSourceMode
} from '../IDarkEventsWebPartProps';

import {
  IEventItem,
  ISharePointEventRow
} from '../models/IEventItem';


interface IRowsResponse {
  value: ISharePointEventRow[];
}

interface IListRow {
  Id: string;
  Title: string;
}

interface IListsResponse {
  value: IListRow[];
}

interface IWebResponse {
  Title: string;
}


export class SharePointEventsService {

  public constructor(
    private readonly context: WebPartContext
  ) {}


  /*
   * =====================================================
   * EVENTS LADEN
   * =====================================================
   */

  public async getEvents(
    sourceMode: EventSourceMode,
    selectedListId: string,
    selectedSites: IPropertyFieldSite[],
    count: number,
    showPast: boolean
  ): Promise<IEventItem[]> {

    const currentWeb =
      this.context.pageContext.web.absoluteUrl;

    let sources: Array<{
      webUrl: string;
      listId: string;
    }> = [];


    /*
     * Eine bestimmte Ereignisliste
     */
    if (sourceMode === 'currentList') {

      if (!selectedListId) {
        throw new Error(
          'Bitte in den Webpart-Eigenschaften eine Ereignisliste auswählen.'
        );
      }

      sources = [
        {
          webUrl: currentWeb,
          listId: selectedListId
        }
      ];

    }


    /*
     * Alle Ereignislisten der aktuellen Website
     */
    else if (sourceMode === 'currentSite') {

      sources =
        await this.getEventLists(
          currentWeb
        );

    }


    /*
     * Ereignislisten der ausgewählten Websites
     */
    else {

      const sites =
        selectedSites || [];

      if (sites.length === 0) {

        throw new Error(
          'Bitte in den Webpart-Eigenschaften mindestens eine Site auswählen.'
        );

      }


      const nested =
        await Promise.all(

          sites.map(
            site =>
              this.getEventLists(
                site.url || ''
              )
          )

        );


      sources =
        nested.reduce(
          (all, part) =>
            all.concat(part),
          []
        );

    }


    /*
     * Keine Ereignisliste gefunden
     */
    if (sources.length === 0) {

      throw new Error(
        'In der gewählten Quelle wurde keine Ereignisliste gefunden.'
      );

    }


    /*
     * Events aus den gefundenen Listen laden
     */
    const groups =
      await Promise.all(

        sources.map(
          async source => {

            /*
             * SharePoint-Sitename laden
             */
            const web =
              await this.getJson<IWebResponse>(
                `${source.webUrl}/_api/web?$select=Title`
              );


            return this.getEventsFromList(
              source.webUrl,
              source.listId,
              web.Title || source.webUrl,
              count,
              showPast
            );

          }
        )

      );


    /*
     * Events aller Sites zusammenführen,
     * chronologisch sortieren und begrenzen.
     */
    return groups
      .reduce(
        (all, part) =>
          all.concat(part),
        []
      )
      .sort(
        (a, b) =>
          a.start.getTime() -
          b.start.getTime()
      )
      .slice(
        0,
        Math.max(
          1,
          Math.min(
            count,
            20
          )
        )
      );

  }


  /*
   * =====================================================
   * URL FÜR "EREIGNIS HINZUFÜGEN"
   *
   * webUrl = Site, auf der das Event erstellt werden soll
   * listId = Ereignisliste dieser Site
   *
   * Ergebnis:
   * /_layouts/15/Event.aspx?ListGuid=...&Mode=Edit
   * =====================================================
   */

  public getAddEventUrl(
    webUrl: string,
    listId: string
  ): string {

    if (
      !webUrl ||
      !listId
    ) {

      return '';

    }


    /*
     * Eventuelle abschliessende Slashes entfernen.
     */
    const normalizedWebUrl =
      webUrl.replace(
        /\/$/,
        ''
      );


    return (
      `${normalizedWebUrl}/_layouts/15/Event.aspx` +
      `?ListGuid=${encodeURIComponent(listId)}` +
      `&Mode=Edit`
    );

  }


  /*
   * =====================================================
   * EREIGNISLISTE DER AKTUELLEN SITE ERMITTELN
   *
   * Diese Methode ist für "+ Ereignis hinzufügen".
   *
   * DarkEvents.tsx übergibt:
   *
   * props.context.pageContext.web.absoluteUrl
   *
   * Dadurch wird NICHT die ausgewählte Anzeige-Site
   * verwendet, sondern immer die Site, auf welcher
   * das Webpart eingebettet ist.
   * =====================================================
   */

  public async getSiteAddEventUrl(
    webUrl: string
  ): Promise<string> {

    if (!webUrl) {

      return '';

    }


    /*
     * Sichtbare Ereignislisten dieser Site suchen
     */
    const eventLists =
      await this.getEventLists(
        webUrl
      );


    if (
      eventLists.length === 0
    ) {

      return '';

    }


    /*
     * Erste gefundene sichtbare
     * Ereignisliste verwenden.
     */
    const eventList =
      eventLists[0];


    return this.getAddEventUrl(
      eventList.webUrl,
      eventList.listId
    );

  }


  /*
   * =====================================================
   * EREIGNISLISTEN EINER SHAREPOINT-SITE SUCHEN
   *
   * BaseTemplate 106 = Ereignisliste
   * =====================================================
   */

  private async getEventLists(
    webUrl: string
  ): Promise<Array<{
    webUrl: string;
    listId: string;
  }>> {

    if (!webUrl) {

      return [];

    }


    const url =
      `${webUrl}/_api/web/lists` +
      `?$select=Id,Title` +
      `&$filter=BaseTemplate eq 106 and Hidden eq false`;


    const data =
      await this.getJson<IListsResponse>(
        url
      );


    return data.value.map(
      list => ({

        webUrl,

        listId:
          list.Id

      })
    );

  }


  /*
   * =====================================================
   * EVENTS EINER BESTIMMTEN LISTE LADEN
   * =====================================================
   */

  private async getEventsFromList(
    webUrl: string,
    listId: string,
    siteName: string,
    count: number,
    showPast: boolean
  ): Promise<IEventItem[]> {

    const now =
      new Date().toISOString();


    /*
     * Standardmässig nur Events anzeigen,
     * deren Ende noch nicht überschritten ist.
     */
    const filter =
      showPast
        ? ''
        : `&$filter=EndDate ge datetime'${now}'`;


    const select =
      'ID,' +
      'Title,' +
      'Description,' +
      'EventDate,' +
      'EndDate,' +
      'Location,' +
      'Category,' +
      'fAllDayEvent,' +
      'BannerUrl,' +
      'FileRef';


    const url =
      `${webUrl}/_api/web/lists(guid'${listId}')/items` +
      `?$select=${select}` +
      `${filter}` +
      `&$orderby=EventDate asc` +
      `&$top=${Math.max(
        1,
        Math.min(
          count,
          20
        )
      )}`;


    const rows =
      await this.getJson<IRowsResponse>(
        url
      );


    return rows.value.map(
      row =>
        this.mapRow(
          row,
          webUrl,
          listId,
          siteName
        )
    );

  }


  /*
   * =====================================================
   * SHAREPOINT REST REQUEST
   * =====================================================
   */

  private async getJson<T>(
    url: string
  ): Promise<T> {

    const response:
      SPHttpClientResponse =
        await this.context.spHttpClient.get(

          url,

          SPHttpClient.configurations.v1,

          {
            headers: {

              Accept:
                'application/json;odata.metadata=none'

            }
          }

        );


    if (!response.ok) {

      const body =
        await response.text();


      throw new Error(

        `SharePoint REST ${response.status}: ` +
        `${body || response.statusText}`

      );

    }


    return response.json() as Promise<T>;

  }


  /*
   * =====================================================
   * SHAREPOINT-EVENT IN UNSER MODELL UMWANDELN
   * =====================================================
   */

  private mapRow(
    row: ISharePointEventRow,
    webUrl: string,
    listId: string,
    siteName: string
  ): IEventItem {

    /*
     * Event-Banner auflösen
     */
    const imageUrl =
      this.getImageUrl(
        row.BannerUrl,
        webUrl
      );


    /*
     * URL zur SharePoint-Eventdetailseite.
     *
     * Klick auf die Event-Kachel navigiert dorthin.
     */
    const eventUrl =
      `${webUrl}/_layouts/15/Event.aspx` +
      `?ListGuid=${encodeURIComponent(listId)}` +
      `&ItemId=${row.ID}`;


    return {

      id:
        row.ID,


      title:
        row.Title ||
        'Ohne Titel',


      description:
        row.Description ||
        '',


      start:
        new Date(
          row.EventDate
        ),


      end:
        new Date(
          row.EndDate
        ),


      location:
        row.Location ||
        '',


      category:
        row.Category ||
        '',


      siteName,


      allDay:
        row.fAllDayEvent === true ||
        row.fAllDayEvent === 1,


      imageUrl,


      eventUrl

    };

  }


  /*
   * =====================================================
   * EVENT-BANNER ERMITTELN
   * =====================================================
   */

  private getImageUrl(
    value:
      ISharePointEventRow['BannerUrl'],
    webUrl: string
  ): string {

    /*
     * Kein Bild
     */
    if (!value) {

      return '';

    }


    /*
     * SharePoint kann das URL-Feld
     * als Objekt liefern.
     */
    if (
      typeof value === 'object' &&
      value.Url
    ) {

      return this.toAbsoluteUrl(
        value.Url,
        webUrl
      );

    }


    if (
      typeof value !== 'string'
    ) {

      return '';

    }


    const trimmed =
      value.trim();


    if (!trimmed) {

      return '';

    }


    /*
     * SharePoint kann Bannerinformationen
     * als JSON-String zurückgeben.
     */
    try {

      const parsed =
        JSON.parse(
          trimmed
        ) as {

          serverRelativeUrl?: string;

          url?: string;

        };


      return this.toAbsoluteUrl(

        parsed.serverRelativeUrl ||
        parsed.url ||
        '',

        webUrl

      );

    }

    catch {

      /*
       * Fallback für normale URLs
       * beziehungsweise SharePoint-Werte.
       */
      return this.toAbsoluteUrl(

        trimmed
          .split(',')[0]
          .trim(),

        webUrl

      );

    }

  }


  /*
   * =====================================================
   * SERVERRELATIVE URL
   * IN ABSOLUTE URL UMWANDELN
   * =====================================================
   */

  private toAbsoluteUrl(
    url: string,
    webUrl: string
  ): string {

    if (!url) {

      return '';

    }


    /*
     * Bereits eine vollständige URL
     */
    if (
      /^https?:\/\//i.test(
        url
      )
    ) {

      return url;

    }


    /*
     * Aus /sites/... eine vollständige
     * SharePoint-URL erzeugen.
     */
    return (

      `${new URL(webUrl).origin}` +

      `${url.startsWith('/') ? '' : '/'}` +

      `${url}`

    );

  }

}