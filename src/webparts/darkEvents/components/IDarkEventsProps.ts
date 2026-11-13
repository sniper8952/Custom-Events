import {
  WebPartContext
} from '@microsoft/sp-webpart-base';

import {
  IPropertyFieldSite
} from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';

import {
  EventSourceMode
} from '../IDarkEventsWebPartProps';


export interface IDarkEventsProps {

  context:
    WebPartContext;


  sourceMode:
    EventSourceMode;

  selectedListId:
    string;

  selectedSites:
    IPropertyFieldSite[];


  canAddEvents:
    boolean;


  numberOfEvents:
    number;

  showPastEvents:
    boolean;

  showWebPartTitle:
    boolean;

  webPartTitle:
    string;

  showSiteName:
    boolean;

  showCategory:
    boolean;


  cardSizePercent:
    string;

  imageHeightPercent:
    string;

  imageFocusXPercent:
    string;

  imageFocusYPercent:
    string;


  backgroundColor:
    string;

  cardBackgroundColor:
    string;


  webPartTitleColor:
    string;

  addEventColor:
    string;

  addEventDisabledColor:
    string;

  categoryColor:
    string;

  siteTitleColor:
    string;

  eventTitleColor:
    string;

  eventDateColor:
    string;

  locationColor:
    string;

  calendarIconColor:
    string;

  paginationColor:
    string;

  arrowBackgroundColor:
    string;


  fallbackImageUrl:
    string;

}