import * as React from 'react';

import {
  IEventItem
} from '../models/IEventItem';

import styles from './DarkEvents.module.scss';


interface IEventCardProps {

  event:
    IEventItem;

  fallbackImageUrl:
    string;

  showSiteName:
    boolean;

  showCategory:
    boolean;

  imageFocusX:
    number;

  imageFocusY:
    number;

}


/*
 * =====================================================
 * DATUMSFORMATIERUNG
 * =====================================================
 */

const dateFormatter =
  new Intl.DateTimeFormat(
    'de-CH',
    {
      weekday: 'short',
      day: '2-digit',
      month: 'short'
    }
  );


const timeFormatter =
  new Intl.DateTimeFormat(
    'de-CH',
    {
      hour: '2-digit',
      minute: '2-digit'
    }
  );


const monthFormatter =
  new Intl.DateTimeFormat(
    'de-CH',
    {
      month: 'short'
    }
  );


/*
 * =====================================================
 * KALENDER + ICON
 * =====================================================
 */

const CalendarPlusIcon:
  React.FC = () => (

    <svg

      className={
        styles.calendarPlusIcon
      }

      viewBox="0 0 24 24"

      aria-hidden="true"

      focusable="false"

    >

      <path
        d="
          M7 2v3
          M17 2v3
          M3.5 8.5h17
          M5.5 4h13
          a2 2 0 0 1 2 2
          v13
          a2 2 0 0 1-2 2
          h-13
          a2 2 0 0 1-2-2
          V6
          a2 2 0 0 1 2-2Z
        "
      />

      <path
        d="
          M12 11.5v6
          M9 14.5h6
        "
      />

    </svg>

  );


/*
 * =====================================================
 * ICS-WERTE ESCAPEN
 * =====================================================
 */

const escapeIcs =
  (value: string): string =>

    value
      .replace(
        /\\/g,
        '\\\\'
      )
      .replace(
        /\r?\n/g,
        '\\n'
      )
      .replace(
        /,/g,
        '\\,'
      )
      .replace(
        /;/g,
        '\\;'
      );


/*
 * =====================================================
 * HTML → TEXT
 *
 * Wird für DESCRIPTION in der ICS verwendet.
 * =====================================================
 */

const htmlToText =
  (value: string): string => {

    const element =
      document.createElement(
        'div'
      );

    element.innerHTML =
      value || '';

    return (
      element.textContent ||
      element.innerText ||
      ''
    );

  };


/*
 * =====================================================
 * DATUM FÜR ICS
 * =====================================================
 */

const toIcsDate =
  (value: Date): string =>

    value
      .toISOString()
      .replace(
        /[-:]/g,
        ''
      )
      .replace(
        /\.\d{3}Z$/,
        'Z'
      );


/*
 * =====================================================
 * DATEINAME FÜR ICS
 * =====================================================
 */

const safeFileName =
  (value: string): string => {

    const cleaned =
      value
        .replace(
          /[\\/:*?"<>|]/g,
          '-'
        )
        .trim();

    return (
      `${cleaned || 'event'}.ics`
    );

  };


/*
 * =====================================================
 * HTML FÜR ICS ESCAPEN
 * =====================================================
 */

const escapeHtmlForIcs =
  (value: string): string =>

    value
      .replace(
        /\\/g,
        '\\\\'
      )
      .replace(
        /\r?\n/g,
        ''
      )
      .replace(
        /,/g,
        '\\,'
      )
      .replace(
        /;/g,
        '\\;'
      );


/*
 * =====================================================
 * ICS HERUNTERLADEN
 * =====================================================
 */

const downloadIcs =
  (event: IEventItem): void => {

    /*
     * Beschreibung als normaler Text.
     */
    const plainDescription =
      htmlToText(
        event.description
      ).trim();


    /*
     * Originales HTML für Outlook.
     */
    const htmlDescription =
      event.description || '';


    const lines = [

      'BEGIN:VCALENDAR',

      'VERSION:2.0',

      'PRODID:-//Agrisano//Dark Events//DE',

      'CALSCALE:GREGORIAN',

      'METHOD:PUBLISH',

      'BEGIN:VEVENT',


      /*
       * Eindeutige ID
       */

      `UID:agrisano-${event.siteName}-${event.id}@sharepoint`,


      /*
       * Erstellungszeit
       */

      `DTSTAMP:${toIcsDate(
        new Date()
      )}`,


      /*
       * Eventzeit
       */

      `DTSTART:${toIcsDate(
        event.start
      )}`,

      `DTEND:${toIcsDate(
        event.end
      )}`,


      /*
       * Eventtitel
       */

      `SUMMARY:${escapeIcs(
        event.title
      )}`,


      /*
       * SharePoint-Kategorie
       */

      ...(
        event.category
          ? [
              `CATEGORIES:${escapeIcs(
                event.category
              )}`
            ]
          : []
      ),


      /*
       * Standort
       */

      `LOCATION:${escapeIcs(
        event.location || ''
      )}`,


      /*
       * Beschreibung als Text
       */

      `DESCRIPTION:${escapeIcs(
        plainDescription
      )}`,


      /*
       * Beschreibung als HTML.
       *
       * Outlook kann dadurch die Formatierung
       * aus "Über dieses Ereignis" übernehmen.
       */

      `X-ALT-DESC;FMTTYPE=text/html:${escapeHtmlForIcs(
        htmlDescription
      )}`,


      'END:VEVENT',

      'END:VCALENDAR'

    ];


    const blob =
      new Blob(
        [
          lines.join(
            '\r\n'
          )
        ],
        {
          type:
            'text/calendar;charset=utf-8'
        }
      );


    const url =
      URL.createObjectURL(
        blob
      );


    const anchor =
      document.createElement(
        'a'
      );


    anchor.href =
      url;


    anchor.download =
      safeFileName(
        event.title
      );


    document.body.appendChild(
      anchor
    );


    anchor.click();


    anchor.remove();


    URL.revokeObjectURL(
      url
    );

  };


/*
 * =====================================================
 * EVENT CARD
 * =====================================================
 */

const EventCard:
  React.FC<IEventCardProps> = ({

    event,

    fallbackImageUrl,

    showSiteName,

    showCategory,

    imageFocusX,

    imageFocusY

  }) => {


    /*
     * ===================================================
     * MEHRTÄGIGE / GANZTÄGIGE EVENTS
     *
     * Bei SharePoint ist das Enddatum eines ganztägigen
     * Events exklusiv.
     *
     * Beispiel:
     *
     * Start:
     * 26.11.2026
     *
     * SharePoint-Enddatum:
     * 01.12.2026
     *
     * Sichtbare Anzeige:
     * 26.–30.11.2026
     * ===================================================
     */

    const displayEnd =
      new Date(
        event.end
      );


    if (
      event.allDay
    ) {

      displayEnd.setDate(
        displayEnd.getDate() - 1
      );

    }


    const sameDay =
      event.start.toDateString() ===
      displayEnd.toDateString();


    /*
     * ===================================================
     * BANNER
     * ===================================================
     */

    const image =
      event.imageUrl ||
      fallbackImageUrl;


    /*
     * ===================================================
     * DATUM / ZEIT UNTEN
     * ===================================================
     */

    const timeText =
      event.allDay

        ? sameDay

          ? `${dateFormatter.format(
              event.start
            )}, Ganztägig`

          : `${dateFormatter.format(
              event.start
            )} – ${dateFormatter.format(
              displayEnd
            )}`

        : sameDay

          ? `${dateFormatter.format(
              event.start
            )}, ${timeFormatter.format(
              event.start
            )}`

          : `${dateFormatter.format(
              event.start
            )} – ${dateFormatter.format(
              displayEnd
            )}`;


    /*
     * ===================================================
     * GROSSES DATUM AUF DEM BANNER
     * ===================================================
     */

    const dayText =
      sameDay

        ? String(
            event.start.getDate()
          )

        : `${
            event.start.getDate()
          }–${
            displayEnd.getDate()
          }`;


    /*
     * ===================================================
     * EVENT ÖFFNEN
     * ===================================================
     */

    const openEvent =
      (): void => {

        window.location.href =
          event.eventUrl;

      };


    /*
     * ===================================================
     * TASTATUR-NAVIGATION
     * ===================================================
     */

    const onCardKeyDown =
      (
        keyboardEvent:
          React.KeyboardEvent<HTMLElement>
      ): void => {

        if (
          keyboardEvent.key ===
            'Enter' ||
          keyboardEvent.key ===
            ' '
        ) {

          keyboardEvent.preventDefault();

          openEvent();

        }

      };


    /*
     * ===================================================
     * KALENDER-KLICK
     *
     * stopPropagation verhindert, dass gleichzeitig
     * die Eventdetailseite geöffnet wird.
     * ===================================================
     */

    const onCalendarClick =
      (
        mouseEvent:
          React.MouseEvent<HTMLButtonElement>
      ): void => {

        mouseEvent.stopPropagation();

        downloadIcs(
          event
        );

      };


    /*
     * ===================================================
     * RENDERING
     * ===================================================
     */

    return (

      <article

        className={
          styles.card
        }

        role="link"

        tabIndex={0}

        aria-label={
          `${event.title} öffnen`
        }

        onClick={
          openEvent
        }

        onKeyDown={
          onCardKeyDown
        }

      >


        {/*
         * =================================================
         * BANNER
         * =================================================
         */}

        <div
          className={
            styles.imageLink
          }
        >


          {image

            ? (

              <img

                className={
                  styles.image
                }

                src={
                  image
                }

                alt=""

                loading="lazy"

                style={{
                  objectPosition:
                    `${imageFocusX}% ${imageFocusY}%`
                }}

              />

            )

            : (

              <div
                className={
                  styles.imageFallback
                }
              />

            )}


          <div
            className={
              styles.imageShade
            }
          />


          {/*
           * Datum auf Banner
           */}

          <div
            className={
              styles.dateBadge
            }
          >

            <span>

              {
                monthFormatter
                  .format(
                    event.start
                  )
                  .replace(
                    '.',
                    ''
                  )
                  .toUpperCase()
              }

            </span>


            <strong>
              {dayText}
            </strong>

          </div>

        </div>


        {/*
         * =================================================
         * KARTENINHALT
         *
         * Jede Zeile besitzt im SCSS einen festen Platz.
         * Dadurch stehen die Informationen in allen
         * Eventkarten auf derselben Höhe.
         * =================================================
         */}

        <div
          className={
            styles.content
          }
        >


          {/*
           * Kategorie
           *
           * Bei deaktiviertem Toggle bleibt der Platz
           * mit visibility:hidden reserviert.
           */}

          <div

            className={
              `${styles.category} ` +
              `${
                !showCategory
                  ? styles.hiddenRow
                  : ''
              }`
            }

          >

            {
              event.category ||
              '\u00A0'
            }

          </div>


          {/*
           * SharePoint-Sitename
           */}

          <div

            className={
              `${styles.siteName} ` +
              `${
                !showSiteName
                  ? styles.hiddenRow
                  : ''
              }`
            }

          >

            {
              event.siteName ||
              '\u00A0'
            }

          </div>


          {/*
           * Eventtitel
           */}

          <div
            className={
              styles.title
            }
          >

            {event.title}

          </div>


          {/*
           * Flexibler Zwischenraum
           */}

          <div
            className={
              styles.spacer
            }
          />


          {/*
           * Eventdatum
           */}

          <div
            className={
              styles.meta
            }
          >

            {timeText}

          </div>


          {/*
           * Standort
           *
           * Auch ohne Standort wird eine Zeile
           * reserviert.
           */}

          <div
            className={
              styles.location
            }
          >

            {
              event.location ||
              '\u00A0'
            }

          </div>

        </div>


        {/*
         * =================================================
         * FOOTER / ICS-DOWNLOAD
         * =================================================
         */}

        <div
          className={
            styles.footer
          }
        >

          <button

            type="button"

            className={
              styles.calendarLink
            }

            title=
              "Als ICS herunterladen"

            aria-label={
              `${event.title} als ICS herunterladen`
            }

            onClick={
              onCalendarClick
            }

          >

            <CalendarPlusIcon />

          </button>

        </div>

      </article>

    );

  };


export default EventCard;