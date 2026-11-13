import * as React from 'react';
import { IEventItem } from '../models/IEventItem';
interface IEventCardProps {
    event: IEventItem;
    accentColor: string;
    fallbackImageUrl: string;
    showSiteName: boolean;
    showCategory: boolean;
    imageFocusX: number;
    imageFocusY: number;
}
declare const EventCard: React.FC<IEventCardProps>;
export default EventCard;
//# sourceMappingURL=EventCard%20-%20Backup.d.ts.map