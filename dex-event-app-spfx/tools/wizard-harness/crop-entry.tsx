/* v31.9.8: Bildschirmfoto des Zuschnitt-Dialogs — er geht nur nach einem
 * echten Upload auf, ist im Wizard-Harness also nicht erreichbar. Hier wird
 * der ECHTE ImageCropModal gemountet; nur der Kinder-Block (die zwei Haken)
 * ist aus BasicsStep gespiegelt, weil er dort im Aufrufer steht.
 *
 * Nicht Teil des Produkt-Builds. `node crop-build.js && node crop-shot.js`.
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import ImageCropModal from '../../src/webparts/dexEventPlatform/components/ImageCropModal';
import { cx } from '../../src/webparts/dexEventPlatform/components/dexUi';

// Ein Querformat-Foto als data:-URL (die Zeichenroutine braucht ein echtes
// Bild, kein Platzhalter): schlichtes SVG mit Streifen, 1200 × 500.
const PHOTO = 'data:image/svg+xml;base64,' + btoa(`<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="500">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
<stop offset="0" stop-color="#5aa9e6"/><stop offset="1" stop-color="#86bc25"/></linearGradient></defs>
<rect width="1200" height="500" fill="url(#g)"/>
<circle cx="600" cy="250" r="120" fill="#ffffff" opacity="0.85"/>
<text x="600" y="270" font-family="Arial" font-size="64" text-anchor="middle" fill="#26890d">FOTO</text>
</svg>`);

const App = (): React.ReactElement => {
  const [photoForEmail, setPhotoForEmail] = React.useState(true);
  const [photoForOutlook, setPhotoForOutlook] = React.useState(true);
  const [imageDisplayOpen, setImageDisplayOpen] = React.useState(false);
  const isDe = true;
  return (
    <ImageCropModal open src={PHOTO} isDe={isDe} onClose={() => undefined} onApply={() => undefined}>
      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">Wo soll das Foto noch erscheinen?</div>
        {([
          { on: photoForEmail, set: setPhotoForEmail, title: 'E-Mails', desc: 'Als Kopfbild in den Mails zu diesem Event.', replaces: false },
          { on: photoForOutlook, set: setPhotoForOutlook, title: 'Outlook-Termin', desc: 'Als Kopfbild im Kalendereintrag.', replaces: true },
        ]).map(row => (
          <label key={row.title} className={cx('dex-ui-toggle-row', row.on && 'is-active')} style={{ marginBottom: 8 }}>
            <input type="checkbox" checked={row.on} onChange={e => row.set(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{row.title}</span>
              <span className="dex-ui-toggle-row-desc">
                {row.desc}
                {row.on && row.replaces && ' Ersetzt dein eigenes Kopfbild.'}
              </span>
            </span>
          </label>
        ))}
      </div>
      <div>
        <button type="button" className={cx('dex-ui-disclosure', imageDisplayOpen && 'is-open')} onClick={() => setImageDisplayOpen(o => !o)}>
          <span className="dex-ui-disclosure-chevron">›</span>
          Darstellung pro Ansicht anpassen
          <span className="dex-ui-disclosure-count">optional</span>
        </button>
      </div>
    </ImageCropModal>
  );
};

ReactDOM.render(<App />, document.getElementById('root'));
