import * as React from 'react';
import { ICalendarProps } from './ICalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FunctionComponent } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { useEffect, useState } from 'react';
import spservices from '../../../services/spservices';
import { Calendar as BigCalendar, momentLocalizer } from 'react-big-calendar';
import * as moment from 'moment';
import 'react-big-calendar/lib/css/react-big-calendar.css';

const Calendar: FunctionComponent<ICalendarProps> =
  ({ description, context }: { description: string, context: WebPartContext }) => {

    const [busy, setBusy] = useState(true)
    const [events, setEvents] = useState([])

    useEffect(() => {
      (async () => {
        await refreshEvents()
      })();
    }, [])

    const refreshEvents = async () => {
      try {
        setBusy(() => true);
        const spService = new spservices(context);
        const calendarEvents = await spService.getEvents();
        setEvents(() => calendarEvents);
        setBusy(() => false);
      }
      catch (error) {
        setBusy(() => false);
      }
    }

    const deleteEvent = async (id: number) => {
      try {
        const spService = new spservices(context);
        await spService.deleteEvent(id);
        const calendarEvents = await spService.getEvents();
        setEvents(() => calendarEvents);
      }
      catch (error) {
        setBusy(() => false);
      }
    }

    const localizer = momentLocalizer(moment)

    return (
      (busy)
        ? (
          <Spinner size={SpinnerSize.large}></Spinner>
        )
        : (
          <>
            <h1>{escape(description)}</h1>
            <BigCalendar
              defaultDate={moment().startOf('day').toDate()}
              localizer={localizer}
              events={events}
              views={{ day: true, week: true, month: true }}
              style={{ height: 500 }}
              components={{
                eventWrapper: ({ event, children }) => (
                  <div
                    onContextMenu={
                      e => {
                        alert(`Will delete ${event.title} (${event.id})!`);
                        deleteEvent(event.id);
                        e.preventDefault();
                      }
                    }
                  >
                    {children}
                  </div>
                )
              }}
            />
          </>
        )
    )
  }

export default Calendar