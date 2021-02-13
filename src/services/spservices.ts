import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/regional-settings";
import * as moment from 'moment';

export default class spservices {

    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async getLocalTime(date: string | Date): Promise<string> {
        try {
            const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(date);
            return localTime
        } catch (error) {
            return Promise.reject(error)
        }
    }

    public async getEvents(): Promise<any[]> {
        try {
            if (sp) {

                const items: any[] = await sp.web.lists.getByTitle("Calendar").items.get();

                const promises = items.map(item =>
                    (async (item) => {

                        let start:Date
                        let end:Date

                        if (item.fAllDayEvent) {

                            const startSplit = item.EventDate.split('-')
                            const endSplit = item.EndDate.split('-')

                            start = new Date(parseInt(startSplit[0]), parseInt(startSplit[1]) - 1, parseInt(startSplit[2]))
                            end = new Date(parseInt(endSplit[0]), parseInt(endSplit[1]) - 1, parseInt(endSplit[2]) + 1)

                        } else {

                            const start1 = await this.getLocalTime(item.EventDate)
                            start = new Date(start1)

                            const end1 = await this.getLocalTime(item.EndDate)
                            end = new Date(end1)

                        }

                        return ({
                            id: item.Id,
                            title: item.Title,
                            allDay: item.fAllDayEvent,
                            start: start,
                            end: end,
                        })

                    })(item))

                return await (Promise.all(promises))

            }
        } catch (error) {
            return Promise.reject(error)
        }
    }

    public async deleteEvent(id: number) {
        try {
            if (sp) {
                const list = sp.web.lists.getByTitle("Calendar");
                await list.items.getById(id).delete();
            }
        } catch (error) {
            return Promise.reject(error)
        }
    }

}