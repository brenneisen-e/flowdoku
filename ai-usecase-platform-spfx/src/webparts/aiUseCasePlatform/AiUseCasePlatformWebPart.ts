/**
 * SPFx WebPart — AI Use Case Platform (Agentic Banking Demo Hub)
 *
 * Duenner Wrapper um die React-Anwendung, genau wie in DEX: Der WebPart
 * kennt nur den Kontext und rendert; alles Weitere passiert in
 * `components/AiUseCasePlatform`.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AiUseCasePlatform, { IAiUseCasePlatformProps } from './components/AiUseCasePlatform';

/** Das Webpart hat keine Eigenschaften im Property-Pane — alles steht in
 *  den SharePoint-Listen. `Record<string, never>` sagt genau das; `{}` hiesse
 *  laut ESLint „jeder nicht-nullische Wert" und damit gar nichts. */
export default class AiUseCasePlatformWebPart extends BaseClientSideWebPart<Record<string, never>> {

  public render(): void {
    const element: React.ReactElement<IAiUseCasePlatformProps> = React.createElement(
      AiUseCasePlatform,
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
