/**
 * Rollenverwaltung.
 *
 * Der Unterschied zu einer naiven Rollenliste steht in `RoleContext`: Eine
 * Rolle wirkt nur mit Leserecht auf der Rollenliste, und ein
 * `addroleassignment`-POST ist noch keine Vergabe. Diese Seite MELDET
 * deshalb, wenn ein Recht nicht gesetzt werden konnte — statt „gespeichert"
 * zu sagen und die Person spaeter ihre Kachel vermissen zu lassen.
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import { Plus, Trash2, AlertCircle } from './Icons';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useDialog } from '../context/DialogContext';
import { UserRole } from '../types';

const ROLLEN: UserRole[] = ['Admin', 'Kurator', 'User'];

export default function RolePage(): React.ReactElement {
  ensureDexUiStyles();
  const { roles, isAdmin, isRolesLoading, rolesReadStatus, addRole, updateRole, removeRole, lastRightsMissing, refreshRoles } = useRoles();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { confirmDialog, showAlert } = useDialog();

  const [email, setEmail] = React.useState('');
  const [name, setName] = React.useState('');
  const [rolle, setRolle] = React.useState<UserRole>('Kurator');
  const [laeuft, setLaeuft] = React.useState(false);

  if (!isAdmin) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Nur für Admins', 'Admins only')}</div>
        <div className="dex-ui-empty-desc">{t('Rollen vergeben dürfen nur Admins der Plattform.', 'Only platform admins may assign roles.')}</div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('start')}>{t('Zur Übersicht', 'Back to overview')}</button>
      </div>
    );
  }

  /** Nach jedem Schreiben: melden, was an Rechten fehlt. Nie verschweigen. */
  function meldeRechte(erfolg: boolean, was: string): void {
    if (!erfolg) { showAlert(t(`${was} hat nicht geklappt.`, `${was} did not work.`), { variant: 'error' }); return; }
    const fehlt = lastRightsMissing();
    if (fehlt.length > 0) {
      showAlert(
        isDe
          ? `${was} ist eingetragen — aber diese Rechte konnten NICHT gesetzt werden: ${fehlt.join(', ')}. Solange sie fehlen, wirkt die Rolle nicht vollständig. Versuch es gleich noch einmal; bleibt es dabei, fehlt dir selbst ein Recht auf der Liste.`
          : `${was} is stored — but these rights could NOT be set: ${fehlt.join(', ')}. While they are missing the role does not fully apply. Try again shortly; if it persists, you lack a right on the list yourself.`,
        { variant: 'error' },
      );
      return;
    }
    showAlert(t(`${was} erledigt.`, `${was} done.`), { variant: 'success' });
  }

  async function hinzufuegen(): Promise<void> {
    const mail = email.trim();
    if (!mail || mail.indexOf('@') < 0) {
      showAlert(t('Bitte eine vollständige E-Mail-Adresse eintragen.', 'Please enter a complete email address.'), { variant: 'error' });
      return;
    }
    setLaeuft(true);
    try {
      const ok = await addRole(mail, name.trim() || mail, rolle);
      if (!ok && roles.some(r => r.userEmail.toLowerCase() === mail.toLowerCase())) {
        // `addRole` verweigert eine implizite Herabstufung — das ist Absicht
        // und muss erklaert werden, sonst wirkt es wie ein Fehler.
        showAlert(
          t('Diese Person hat bereits eine höhere Rolle. Eine Herabstufung geht bewusst nur über die Auswahl in der Zeile unten.',
            'This person already has a higher role. Downgrading deliberately only works via the dropdown in the row below.'),
          { variant: 'error' },
        );
        return;
      }
      meldeRechte(ok, t('Rolle vergeben', 'Assigning the role'));
      if (ok) { setEmail(''); setName(''); }
    } finally {
      setLaeuft(false);
    }
  }

  async function entfernen(id: number, mail: string): Promise<void> {
    const ja = await confirmDialog(
      t(`Rolle von ${mail} entfernen? Die Person verliert damit auch ihre Rechte auf den Listen der Plattform.`,
        `Remove the role of ${mail}? The person also loses their rights on the platform lists.`),
      { confirmLabel: t('Entfernen', 'Remove'), danger: true },
    );
    if (!ja) return;
    setLaeuft(true);
    try {
      meldeRechte(await removeRole(id), t('Entfernen', 'Removing'));
    } finally {
      setLaeuft(false);
    }
  }

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{t('Rollen', 'Roles')}</h1>
          <p className="dex-ui-page-head-meta">{roles.length} {t('Einträge', 'entries')}</p>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="dex-ui-textbtn" onClick={() => { void refreshRoles(); }}>{t('Neu laden', 'Reload')}</button>
        </div>
      </div>

      {rolesReadStatus === 'forbidden' && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginBottom: 16 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{t('Die Rollenliste ist für dich nicht lesbar. Was hier steht, ist deshalb unvollständig.', 'The roles list is not readable for you. What you see here is therefore incomplete.')}</span>
        </div>
      )}

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{t('Wer soll dazukommen?', 'Who should be added?')}</div>
        <div className="dex-ui-grid-3">
          <label className="dex-ui-field">
            <span className="dex-ui-label">{t('E-Mail-Adresse', 'Email address')}</span>
            <input id="rolle-mail" type="email" className="dex-ui-input" value={email} onChange={e => setEmail(e.target.value)} placeholder="vorname.nachname@deloitte.de" />
          </label>
          <label className="dex-ui-field">
            <span className="dex-ui-label">{t('Name', 'Name')} <span className="dex-ui-label-optional">{t('optional', 'optional')}</span></span>
            <input id="rolle-name" type="text" className="dex-ui-input" value={name} onChange={e => setName(e.target.value)} />
          </label>
          <label className="dex-ui-field">
            <span className="dex-ui-label">{t('Rolle', 'Role')}</span>
            <select id="rolle-wahl" className="dex-ui-select" value={rolle} onChange={e => setRolle(e.target.value as UserRole)}>
              {ROLLEN.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </label>
        </div>
        <button type="button" className="btn btn-primary" disabled={laeuft} onClick={() => { void hinzufuegen(); }}>
          <Plus size={15} /> {t('Rolle vergeben', 'Assign role')}
        </button>
        <p className="dex-ui-help" style={{ marginTop: 8 }}>
          {t('Admin: alles, inklusive dieser Seite. Kurator: Use Cases anlegen und pflegen. User: sehen und aufrufen — dafür braucht niemand einen Eintrag.',
            'Admin: everything, including this page. Curator: create and maintain use cases. User: view and launch — nobody needs an entry for that.')}
        </p>
      </div>

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{t('Vergebene Rollen', 'Assigned roles')}</div>
        {isRolesLoading ? (
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
        ) : roles.length === 0 ? (
          <p className="dex-ui-help" style={{ margin: 0 }}>{t('Noch niemand eingetragen.', 'Nobody assigned yet.')}</p>
        ) : (
          <div className="dex-ui-stack">
            {roles.map(r => (
              <div key={r.id} className="dex-ui-row dex-ui-row--bordered">
                <span className="dex-ui-row-main">
                  <span className="dex-ui-row-title dex-ui-row-title--wrap">
                    {r.userName || r.userEmail}
                    <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', r.role === 'Admin' ? 'dex-ui-pill--orange' : r.role === 'Kurator' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')}>{r.role}</span>
                  </span>
                  <span className="dex-ui-row-sub">{r.userEmail}{r.assignedBy ? ` · ${t('vergeben von', 'assigned by')} ${r.assignedBy}` : ''}</span>
                </span>
                <span className="dex-ui-row-actions">
                  <select
                    className="dex-ui-select dex-ui-select--sm"
                    value={r.role}
                    disabled={laeuft}
                    aria-label={t('Rolle ändern', 'Change role')}
                    onChange={e => {
                      const neu = e.target.value as UserRole;
                      setLaeuft(true);
                      // Kein `.finally` — die es5-lib dieses Projekts kennt es
                      // nicht. Das zweiarmige `then` deckt beide Ausgaenge ab.
                      void updateRole(r.id, neu).then(
                        ok => { meldeRechte(ok, t('Rolle ändern', 'Changing the role')); setLaeuft(false); },
                        () => { meldeRechte(false, t('Rolle ändern', 'Changing the role')); setLaeuft(false); },
                      );
                    }}
                  >
                    {ROLLEN.map(x => <option key={x} value={x}>{x}</option>)}
                  </select>
                  <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" disabled={laeuft} onClick={() => { void entfernen(r.id, r.userEmail); }} title={t('Entfernen', 'Remove')} aria-label={t('Entfernen', 'Remove')}>
                    <Trash2 size={15} />
                  </button>
                </span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
