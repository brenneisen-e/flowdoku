/**
 * SPFx WebPart - DEX Event Experience Platform
 *
 * Wrapper-Klasse die das React-basierte Event-Management
 * als SharePoint WebPart bereitstellt.
 *
 * - Eike, Maerz 2026
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DexEventPlatform, { IDexEventPlatformProps } from './components/DexEventPlatform';

export default class DexEventPlatformWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element: React.ReactElement<IDexEventPlatformProps> = React.createElement(
      DexEventPlatform,
      { context: this.context }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
