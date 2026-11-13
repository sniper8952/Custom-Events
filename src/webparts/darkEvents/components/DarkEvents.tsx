import * as React from 'react';

import {
  useEffect,
  useMemo,
  useRef,
  useState
} from 'react';

import { IDarkEventsProps } from './IDarkEventsProps';
import { IEventItem } from '../models/IEventItem';
import { SharePointEventsService } from '../services/SharePointEventsService';
import EventCard from './EventCard';
import styles from './DarkEvents.module.scss';


const DarkEvents:
  React.FC<IDarkEventsProps> =
  props => {

    const [events, setEvents] =
      useState<IEventItem[]>([]);

    const [loading, setLoading] =
      useState<boolean>(true);

    const [error, setError] =
      useState<string>('');

    const [activePage, setActivePage] =
      useState<number>(0);

    const [pageCount, setPageCount] =
      useState<number>(1);

    const [addEventUrl, setAddEventUrl] =
      useState<string>('');

    const trackRef =
      useRef<HTMLDivElement>(null);


    const service =
      useMemo(
        () =>
          new SharePointEventsService(
            props.context
          ),
        [props.context]
      );


    /*
     * ===================================================
     * EVENTS UND ADD-EVENT-URL LADEN
     * ===================================================
     */

    useEffect(() => {

      let active = true;

      setLoading(true);
      setError('');
      setAddEventUrl('');


      const loadAddEventUrl =
        async (): Promise<void> => {

          try {

            const currentWebUrl =
              props.context.pageContext.web.absoluteUrl;

            const url =
              await service.getSiteAddEventUrl(
                currentWebUrl
              );

            if (active) {
              setAddEventUrl(url);
            }

          } catch (reason: unknown) {

            console.error(
              'Die URL für "Ereignis hinzufügen" ' +
              'konnte nicht ermittelt werden:',
              reason
            );

            if (active) {
              setAddEventUrl('');
            }

          }

        };


      void loadAddEventUrl();


      service
        .getEvents(
          props.sourceMode,
          props.selectedListId,
          props.selectedSites,
          props.numberOfEvents,
          props.showPastEvents
        )
        .then(
          (items: IEventItem[]) => {

            if (!active) {
              return;
            }

            setEvents(items);
            setActivePage(0);
            setLoading(false);

          }
        )
        .catch(
          (reason: unknown) => {

            if (!active) {
              return;
            }

            setEvents([]);

            setError(
              reason instanceof Error
                ? reason.message
                : 'Ereignisse konnten nicht geladen werden.'
            );

            setLoading(false);

          }
        );


      return () => {
        active = false;
      };

    }, [
      service,
      props.context,
      props.sourceMode,
      props.selectedListId,
      props.selectedSites,
      props.numberOfEvents,
      props.showPastEvents
    ]);


    /*
     * ===================================================
     * PAGINIERUNG
     * ===================================================
     */

    const updatePagination =
      (): void => {

        const track =
          trackRef.current;

        if (!track) {
          return;
        }


        const clientWidth =
          Math.max(
            1,
            track.clientWidth
          );


        const pages =
          Math.max(
            1,
            Math.ceil(
              track.scrollWidth /
              clientWidth
            )
          );


        const page =
          Math.min(
            pages - 1,
            Math.round(
              track.scrollLeft /
              clientWidth
            )
          );


        setPageCount(pages);
        setActivePage(page);

      };


    useEffect(() => {

      const track =
        trackRef.current;

      if (!track) {
        return;
      }


      const animationFrame =
        window.requestAnimationFrame(
          updatePagination
        );


      const onResize =
        (): void => {

          updatePagination();

        };


      window.addEventListener(
        'resize',
        onResize
      );


      return () => {

        window.cancelAnimationFrame(
          animationFrame
        );

        window.removeEventListener(
          'resize',
          onResize
        );

      };

    }, [events]);


    /*
     * ===================================================
     * NAVIGATION
     * ===================================================
     */

    const goToPage =
      (page: number): void => {

        const track =
          trackRef.current;

        if (!track) {
          return;
        }


        const nextPage =
          Math.max(
            0,
            Math.min(
              pageCount - 1,
              page
            )
          );


        track.scrollTo({

          left:
            nextPage *
            track.clientWidth,

          behavior:
            'smooth'

        });


        setActivePage(
          nextPage
        );

      };


    /*
     * ===================================================
     * DARSTELLUNGSWERTE
     * ===================================================
     */

    const cardSize =
      Math.max(
        50,
        Math.min(
          200,
          Number(
            props.cardSizePercent
          ) || 100
        )
      ) / 100;


    const imageHeight =
      Math.max(
        50,
        Math.min(
          200,
          Number(
            props.imageHeightPercent
          ) || 100
        )
      ) / 100;


    const imageFocusX =
      Math.max(
        0,
        Math.min(
          100,
          Number(
            props.imageFocusXPercent
          ) || 50
        )
      );


    const imageFocusY =
      Math.max(
        0,
        Math.min(
          100,
          Number(
            props.imageFocusYPercent
          ) || 50
        )
      );


    /*
     * ===================================================
     * RENDERING
     * ===================================================
     */

    return (

      <section

        className={styles.root}

        style={{

          ['--root-bg' as string]:
            props.backgroundColor,

          ['--card-bg' as string]:
            props.cardBackgroundColor,

          ['--text' as string]:
            '#FFFFFF',

          ['--accent' as string]:
            props.siteTitleColor,

          ['--webpart-title-color' as string]:
            props.webPartTitleColor,

          ['--add-event-color' as string]:
            props.addEventColor,

          ['--add-event-disabled-color' as string]:
            props.addEventDisabledColor,

          ['--category-color' as string]:
            props.categoryColor,

          ['--site-title-color' as string]:
            props.siteTitleColor,

          ['--event-title-color' as string]:
            props.eventTitleColor,

          ['--event-date-color' as string]:
            props.eventDateColor,

          ['--location-color' as string]:
            props.locationColor,

          ['--calendar-icon-color' as string]:
            props.calendarIconColor,

          ['--pagination-color' as string]:
            props.paginationColor,

          ['--arrow-background-color' as string]:
            props.arrowBackgroundColor,

          ['--card-size' as string]:
            `${cardSize}`,

          ['--image-height' as string]:
            `${imageHeight}`

        }}

      >


        {/*
         * WEBPART-TITEL
         */}

        {props.showWebPartTitle && (

          <div className={styles.header}>

            <h2>
              {props.webPartTitle}
            </h2>

          </div>

        )}


        {/*
         * EREIGNIS HINZUFÜGEN
         */}

        {addEventUrl && (

          <div className={styles.addEventRow}>

            {props.canAddEvents

              ? (

                <a
                  className={styles.addEvent}
                  href={addEventUrl}
                >

                  <span
                    className={styles.addEventIcon}
                    aria-hidden="true"
                  >
                    +
                  </span>

                  <span>
                    Ereignis hinzufügen
                  </span>

                </a>

              )

              : (

                <span

                  className={
                    `${styles.addEvent} ` +
                    `${styles.addEventDisabled}`
                  }

                  aria-disabled="true"

                  title="Sie verfügen nicht über die Berechtigung, Ereignisse hinzuzufügen."

                >

                  <span
                    className={styles.addEventIcon}
                    aria-hidden="true"
                  >
                    +
                  </span>

                  <span>
                    Ereignis hinzufügen
                  </span>

                </span>

              )}

          </div>

        )}


        {/*
         * LADESTATUS
         */}

        {loading && (

          <div className={styles.status}>
            Ereignisse werden geladen …
          </div>

        )}


        {/*
         * FEHLER
         */}

        {!loading &&
          error && (

            <div className={styles.error}>

              <strong>
                Fehler beim Laden
              </strong>

              <br />

              {error}

            </div>

          )}


        {/*
         * KEINE EVENTS
         */}

        {!loading &&
          !error &&
          events.length === 0 && (

            <div className={styles.status}>
              Keine passenden Ereignisse gefunden.
            </div>

          )}


        {/*
         * EVENT-CAROUSEL
         */}

        {!loading &&
          !error &&
          events.length > 0 && (

            <div className={styles.carousel}>


              {/*
               * LINKER PFEIL
               */}

              <button

                type="button"

                className={
                  `${styles.arrow} ` +
                  `${styles.arrowLeft}`
                }

                onClick={
                  () =>
                    goToPage(
                      activePage - 1
                    )
                }

                aria-label="Vorherige Ereignisse"

                disabled={
                  activePage === 0
                }

              >

                <span
                  className={styles.arrowIcon}
                  aria-hidden="true"
                >
                  ‹
                </span>

              </button>


              {/*
               * KACHELN
               */}

              <div

                className={styles.track}

                ref={trackRef}

                onScroll={updatePagination}

              >

                {events.map(
                  event => (

                    <EventCard

                      key={
                        `${event.siteName}-${event.id}`
                      }

                      event={event}

                      fallbackImageUrl={
                        props.fallbackImageUrl
                      }

                      showSiteName={
                        props.showSiteName
                      }

                      showCategory={
                        props.showCategory
                      }

                      imageFocusX={
                        imageFocusX
                      }

                      imageFocusY={
                        imageFocusY
                      }

                    />

                  )
                )}

              </div>


              {/*
               * RECHTER PFEIL
               */}

              <button

                type="button"

                className={
                  `${styles.arrow} ` +
                  `${styles.arrowRight}`
                }

                onClick={
                  () =>
                    goToPage(
                      activePage + 1
                    )
                }

                aria-label="Nächste Ereignisse"

                disabled={
                  activePage >=
                  pageCount - 1
                }

              >

                <span
                  className={styles.arrowIcon}
                  aria-hidden="true"
                >
                  ›
                </span>

              </button>


              {/*
               * NAVIGATIONSPUNKTE
               */}

              {pageCount > 1 && (

                <div
                  className={styles.dots}
                  aria-label="Ereignisseiten"
                >

                  {Array.from(
                    {
                      length:
                        pageCount
                    },
                    (
                      _unused,
                      index
                    ) => (

                      <button

                        key={index}

                        type="button"

                        className={
                          index === activePage
                            ? styles.dotActive
                            : styles.dot
                        }

                        onClick={
                          () =>
                            goToPage(
                              index
                            )
                        }

                        aria-label={
                          `Seite ${index + 1}`
                        }

                        aria-current={
                          index === activePage
                            ? 'true'
                            : undefined
                        }

                      />

                    )
                  )}

                </div>

              )}

            </div>

          )}

      </section>

    );

  };


export default DarkEvents;