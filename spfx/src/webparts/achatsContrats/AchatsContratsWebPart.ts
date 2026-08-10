import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AchatsContratsWebPartStrings';
import App, { IAppProps } from './components/App';
import { getSp, IListsConfig } from './services/spClient';

export interface IAchatsContratsWebPartProps {
  conditionsListTitle: string;
  templatesListTitle: string;
  mappingsListTitle: string;
  generationsListTitle: string;
  documentLibraryTitle: string;
}

export default class AchatsContratsWebPart extends BaseClientSideWebPart<IAchatsContratsWebPartProps> {
  public render(): void {
    const config: IListsConfig = {
      conditionsListTitle: this.properties.conditionsListTitle || 'AchatsContrats_ConditionsVersions',
      templatesListTitle: this.properties.templatesListTitle || 'AchatsContrats_TemplatesContrats',
      mappingsListTitle: this.properties.mappingsListTitle || 'AchatsContrats_MappingsTemplate',
      generationsListTitle: this.properties.generationsListTitle || 'AchatsContrats_Generations',
      documentLibraryTitle: this.properties.documentLibraryTitle || 'AchatsContrats_Fichiers',
    };

    const element: React.ReactElement<IAppProps> = React.createElement(App, {
      sp: getSp(this.context),
      config,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('conditionsListTitle', {
                  label: strings.ConditionsListTitleFieldLabel,
                }),
                PropertyPaneTextField('templatesListTitle', {
                  label: strings.TemplatesListTitleFieldLabel,
                }),
                PropertyPaneTextField('mappingsListTitle', {
                  label: strings.MappingsListTitleFieldLabel,
                }),
                PropertyPaneTextField('generationsListTitle', {
                  label: strings.GenerationsListTitleFieldLabel,
                }),
                PropertyPaneTextField('documentLibraryTitle', {
                  label: strings.DocumentLibraryTitleFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
