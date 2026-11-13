import * as React from 'react';

import {
  createRoot,
  Root
} from 'react-dom/client';

import {
  IPropertyPaneConfiguration,
  IPropertyPaneGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  ThemeProvider,
  ThemeChangedEventArgs
} from '@microsoft/sp-component-base';

import {
  SPPermission
} from '@microsoft/sp-page-context';

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import {
  PropertyFieldSitePicker
} from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';

import DarkEvents from './components/DarkEvents';

import {
  IDarkEventsWebPartProps
} from './IDarkEventsWebPartProps';


export default class DarkEventsWebPart
  extends BaseClientSideWebPart<IDarkEventsWebPartProps> {

  private root:
    Root | undefined;

  private themeProvider:
    ThemeProvider | undefined;

  private canManageWebPart:
    boolean = false;

  private canAddEvents:
    boolean = false;


  public async onInit():
    Promise<void> {

    await super.onInit();


    this.canManageWebPart =
      this.context.pageContext.web.permissions.hasPermission(
        SPPermission.manageWeb
      );


    this.canAddEvents =
      this.context.pageContext.web.permissions.hasPermission(
        SPPermission.addListItems
      );


    this.themeProvider =
      this.context.serviceScope.consume(
        ThemeProvider.serviceKey
      );


    this.themeProvider.themeChangedEvent.add(
      this,
      this.handleThemeChanged
    );

  }


  public render():
    void {

    if (!this.root) {

      this.root =
        createRoot(
          this.domElement
        );

    }


    this.root.render(

      React.createElement(
        DarkEvents,
        {

          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          context:
            this.context as any,


          sourceMode:
            this.properties.sourceMode ||
            'currentList',

          selectedListId:
            this.properties.selectedListId ||
            '',

          selectedSites:
            this.properties.selectedSites ||
            [],


          canAddEvents:
            this.canAddEvents,


          numberOfEvents:
            this.properties.numberOfEvents ||
            8,

          showPastEvents:
            !!this.properties.showPastEvents,

          showWebPartTitle:
            this.properties.showWebPartTitle !== false,

          webPartTitle:
            this.properties.webPartTitle ||
            'Events',

          showSiteName:
            this.properties.showSiteName !== false,

          showCategory:
            this.properties.showCategory !== false,


          cardSizePercent:
            this.properties.cardSizePercent ||
            '100',

          imageHeightPercent:
            this.properties.imageHeightPercent ||
            '100',

          imageFocusXPercent:
            this.properties.imageFocusXPercent ||
            '50',

          imageFocusYPercent:
            this.properties.imageFocusYPercent ||
            '50',


          backgroundColor:
            this.properties.backgroundColor ||
            'transparent',

          cardBackgroundColor:
            this.properties.cardBackgroundColor ||
            '#1F1F1F',


          webPartTitleColor:
            this.properties.webPartTitleColor ||
            '#1F1F1F',

          addEventColor:
            this.properties.addEventColor ||
            '#000000',

          addEventDisabledColor:
            this.properties.addEventDisabledColor ||
            '#D3D3D3',

          categoryColor:
            this.properties.categoryColor ||
            '#FFFFFF',

          siteTitleColor:
            this.properties.siteTitleColor ||
            '#9EA500',

          eventTitleColor:
            this.properties.eventTitleColor ||
            '#FFFFFF',

          eventDateColor:
            this.properties.eventDateColor ||
            '#FFFFFF',

          locationColor:
            this.properties.locationColor ||
            '#9EA500',

          calendarIconColor:
            this.properties.calendarIconColor ||
            '#FFFFFF',

          paginationColor:
            this.properties.paginationColor ||
            '#000000',

          arrowBackgroundColor:
            this.properties.arrowBackgroundColor ||
            '#282828',


          fallbackImageUrl:
            this.properties.fallbackImageUrl ||
            ''

        }
      )

    );

  }


  protected onDispose():
    void {

    this.themeProvider
      ?.themeChangedEvent
      .remove(
        this,
        this.handleThemeChanged
      );


    this.root?.unmount();

    this.root =
      undefined;

  }


  private handleThemeChanged =
    (_args: ThemeChangedEventArgs):
      void => {

      this.render();

    };


  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {


    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sourceFields: any[] = [

      PropertyPaneDropdown(
        'sourceMode',
        {

          label:
            'Quelle',

          options: [

            {
              key:
                'currentList',

              text:
                'Liste der Ereignisse auf dieser Site'
            },

            {
              key:
                'currentSite',

              text:
                'Diese Site'
            },

            {
              key:
                'selectedSites',

              text:
                'Sites auswählen'
            }

          ]

        }
      )

    ];


    if (
      (
        this.properties.sourceMode ||
        'currentList'
      ) ===
      'currentList'
    ) {

      sourceFields.push(

        PropertyFieldListPicker(
          'selectedListId',
          {

            label:
              'Ereignisliste',

            selectedList:
              this.properties.selectedListId,

            baseTemplate:
              106,

            includeHidden:
              false,

            orderBy:
              PropertyFieldListPickerOrderBy.Title,

            disabled:
              false,

            multiSelect:
              false,

            onPropertyChange:
              this.onPropertyPaneFieldChanged.bind(
                this
              ),

            properties:
              this.properties,

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            context:
              this.context as any,

            onGetErrorMessage:
              undefined,

            deferredValidationTime:
              0,

            key:
              'eventListPicker'

          }
        )

      );

    }


    if (
      this.properties.sourceMode ===
      'selectedSites'
    ) {

      sourceFields.push(

        PropertyFieldSitePicker(
          'selectedSites',
          {

            label:
              'Sites auswählen',

            initialSites:
              this.properties.selectedSites ||
              [],

            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            context:
              this.context as any,

            deferredValidationTime:
              500,

            multiSelect:
              true,

            onPropertyChange:
              this.onPropertyPaneFieldChanged.bind(
                this
              ),

            properties:
              this.properties,

            key:
              'eventSitePicker'

          }
        )

      );

    }


    const groups:
      IPropertyPaneGroup[] = [

        {

          groupName:
            'Datenquelle',

          groupFields:
            sourceFields

        }

      ];


    if (
      this.canManageWebPart
    ) {


      /*
       * =================================================
       * ANZEIGE
       * =================================================
       */

      groups.push({

        groupName:
          'Anzeige',

        groupFields: [


          PropertyPaneToggle(
            'showWebPartTitle',
            {

              label:
                'Webpart-Titel anzeigen',

              onText:
                'Ein',

              offText:
                'Aus'

            }
          ),


          PropertyPaneTextField(
            'webPartTitle',
            {

              label:
                'Webpart-Titel',

              description:
                'Beispiel: Events, Veranstaltungen oder Termine',

              disabled:
                this.properties.showWebPartTitle === false

            }
          ),


          PropertyPaneSlider(
            'numberOfEvents',
            {

              label:
                'Anzahl Ereignisse',

              min:
                1,

              max:
                20,

              step:
                1

            }
          ),


          PropertyPaneToggle(
            'showPastEvents',
            {

              label:
                'Vergangene Ereignisse anzeigen',

              onText:
                'Ja',

              offText:
                'Nein'

            }
          ),


          PropertyPaneToggle(
            'showSiteName',
            {

              label:
                'SharePoint-Sitename anzeigen',

              onText:
                'Ein',

              offText:
                'Aus'

            }
          ),


          PropertyPaneToggle(
            'showCategory',
            {

              label:
                'Kategorie anzeigen',

              onText:
                'Ein',

              offText:
                'Aus'

            }
          ),


          PropertyPaneTextField(
            'cardSizePercent',
            {

              label:
                'Kachelgrösse in % (50–200)'

            }
          ),


          PropertyPaneTextField(
            'imageHeightPercent',
            {

              label:
                'Bildhöhe in % (50–200)'

            }
          ),


          PropertyPaneTextField(
            'imageFocusXPercent',
            {

              label:
                'Banner-Fokus horizontal in % (0–100)'

            }
          ),


          PropertyPaneTextField(
            'imageFocusYPercent',
            {

              label:
                'Banner-Fokus vertikal in % (0–100)'

            }
          )

        ]

      });


      /*
       * =================================================
       * FARBEN
       * =================================================
       */

      groups.push({

        groupName:
          'Farben',

        groupFields: [


          PropertyPaneTextField(
            'backgroundColor',
            {

              label:
                'Webpart-Hintergrund',

              description:
                'Standard: transparent'

            }
          ),


          PropertyPaneTextField(
            'cardBackgroundColor',
            {

              label:
                'Karten-Hintergrund',

              description:
                'Standard: #1F1F1F'

            }
          ),


          PropertyPaneTextField(
            'webPartTitleColor',
            {

              label:
                'Webpart-Titel',

              description:
                'Standard: #1F1F1F'

            }
          ),


          PropertyPaneTextField(
            'addEventColor',
            {

              label:
                '"Ereignis hinzufügen"',

              description:
                'Standard: #000000'

            }
          ),


          PropertyPaneTextField(
            'addEventDisabledColor',
            {

              label:
                '"Ereignis hinzufügen" deaktiviert',

              description:
                'Standard: #D3D3D3'

            }
          ),


          PropertyPaneTextField(
            'categoryColor',
            {

              label:
                'Kategorie',

              description:
                'Standard: #FFFFFF'

            }
          ),


          PropertyPaneTextField(
            'siteTitleColor',
            {

              label:
                'Seitentitel',

              description:
                'Standard: #9EA500'

            }
          ),


          PropertyPaneTextField(
            'eventTitleColor',
            {

              label:
                'Eventtitel',

              description:
                'Standard: #FFFFFF'

            }
          ),


          PropertyPaneTextField(
            'eventDateColor',
            {

              label:
                'Eventdatum',

              description:
                'Standard: #FFFFFF'

            }
          ),


          PropertyPaneTextField(
            'locationColor',
            {

              label:
                'Standort',

              description:
                'Standard: #9EA500'

            }
          ),


          PropertyPaneTextField(
            'calendarIconColor',
            {

              label:
                'Kalender +',

              description:
                'Standard: #FFFFFF'

            }
          ),


          PropertyPaneTextField(
            'paginationColor',
            {

              label:
                'Navigationspunkte',

              description:
                'Standard: #000000'

            }
          ),


          /*
           * NEU
           */

          PropertyPaneTextField(
            'arrowBackgroundColor',
            {

              label:
                'Pfeil-Hintergrund',

              description:
                'Standard: #282828'

            }
          ),


          PropertyPaneTextField(
            'fallbackImageUrl',
            {

              label:
                'Fallback-Bild-URL (optional)'

            }
          )

        ]

      });

    }


    return {

      pages: [

        {

          header: {

            description:
              this.canManageWebPart

                ? 'Quelle, Anzeige und Farben'

                : 'Ereignisquelle'

          },

          groups

        }

      ]

    };

  }

}