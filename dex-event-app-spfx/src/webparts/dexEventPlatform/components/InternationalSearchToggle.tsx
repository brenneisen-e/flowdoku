import * as React from 'react';

interface Props {
  checked: boolean;
  onChange: (next: boolean) => void;
  isDe?: boolean;
  compact?: boolean;
  /** v24.3: Aktueller Sucheingabe-Wert. Wenn übergeben, erscheint der Toggle
   *  erst, sobald mindestens ein Zeichen eingetippt wurde (app-weit einheitlich).
   *  Ohne diese Prop bleibt das alte Verhalten (immer sichtbar). */
  query?: string;
}

const InternationalSearchToggle: React.FC<Props> = ({ checked, onChange, isDe = true, compact = false, query }) => {
  // v24.3: erst ab dem ersten eingetippten Zeichen anzeigen.
  if (query !== undefined && query.trim().length === 0) return null;
  const label = isDe
    ? 'Auch international suchen (@deloitte.com)'
    : 'Search internationally too (@deloitte.com)';
  const hint = isDe
    ? 'Standardmäßig wird nur in der deutschen Member-Firm (@deloitte.de) gesucht.'
    : 'By default only the German member firm (@deloitte.de) is searched.';
  return (
    <label
      title={hint}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 6,
        fontSize: compact ? 11 : 12,
        color: 'var(--dex-gray-600, #5a5a5a)',
        cursor: 'pointer',
        userSelect: 'none',
        marginTop: 4,
      }}
    >
      <input
        type="checkbox"
        checked={checked}
        onChange={e => onChange(e.target.checked)}
        style={{ cursor: 'pointer' }}
      />
      <span>{label}</span>
    </label>
  );
};

export default InternationalSearchToggle;
