import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IDarkEventsWebPartProps } from './IDarkEventsWebPartProps';
export default class DarkEventsWebPart extends BaseClientSideWebPart<IDarkEventsWebPartProps> {
    private root;
    private themeProvider;
    onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    private handleThemeChanged;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=DarkEventsWebPart%20-%20Backup.d.ts.map