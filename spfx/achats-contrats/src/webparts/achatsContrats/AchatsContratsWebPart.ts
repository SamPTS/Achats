import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AchatsContrats from './components/AchatsContrats';
import { IAchatsContratsProps } from './components/IAchatsContratsProps';

export interface IAchatsContratsWebPartProps {
  description: string;
}

export default class AchatsContratsWebPart extends BaseClientSideWebPart<IAchatsContratsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAchatsContratsProps> = React.createElement(AchatsContrats, {
      context: this.context,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Pas de panneau de propriétés : rien à configurer, tout est auto-provisionné au premier chargement.
  protected getPropertyPaneConfiguration(): { pages: [] } {
    return { pages: [] };
  }
}
