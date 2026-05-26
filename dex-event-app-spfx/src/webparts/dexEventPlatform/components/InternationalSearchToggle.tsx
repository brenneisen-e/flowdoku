import * as React from 'react';

interface Props {
  checked: boolean;
  onChange: (next: boolean) => void;
  isDe?: boolean;
  compact?: boolean;
}

const InternationalSearchToggle: React.FC<Props> = ({ checked, onChange, isDe = true, compact = false }) => {
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
