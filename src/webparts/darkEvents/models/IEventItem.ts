export interface IEventItem {
  id: number;
  title: string;
  description: string;
  start: Date;
  end: Date;
  location: string;
  category: string;
  siteName: string;
  allDay: boolean;
  imageUrl: string;
  eventUrl: string;
}

export interface ISharePointEventRow {
  ID: number;
  Title?: string;
  Description?: string;
  EventDate: string;
  EndDate: string;
  Location?: string;
  Category?: string;
  fAllDayEvent?: boolean | number;
  BannerUrl?: string | { Url?: string };
  FileRef?: string;
}
