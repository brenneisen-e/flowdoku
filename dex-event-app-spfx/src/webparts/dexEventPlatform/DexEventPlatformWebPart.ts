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
import { APP_VERSION } from './version';

/**
 * v29.51 (Messpunkt): Zeit vom Seitenaufruf bis hierher.
 *
 * Dieser Wert läuft, BEVOR die App einen einzigen Request abgesetzt hat — er
 * misst also, was SharePoint für die Seite selbst braucht plus Download,
 * Parsen und Ausführen unseres Bundles. Genau dieser Anteil ist bei der
 * Ursachensuche zum „DEX lädt zehn Sekunden" bisher unsichtbar geblieben; alle
 * bisherigen Messungen (`[DEX][perf][boot]`) beginnen erst danach.
 */
const bootShellMs = (typeof performance !== 'undefined' && performance.now) ? Math.round(performance.now()) : -1;

export default class DexEventPlatformWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    // eslint-disable-next-line no-console
    console.log(`[DEX][perf][shell] v${APP_VERSION} · Bundle bereit nach ${bootShellMs} ms (Seite + Download + Parsen, noch kein Request)`);
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
