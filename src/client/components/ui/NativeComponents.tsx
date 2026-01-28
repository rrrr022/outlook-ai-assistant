/**
 * Simple native UI components to replace Fluent UI
 * Reduces bundle size from 60MB to ~2MB
 */
import React from 'react';

// Simple Spinner component
export const Spinner: React.FC<{ size?: 'small' | 'medium' | 'large'; label?: string }> = ({ 
  size = 'medium', 
  label 
}) => {
  const sizeMap = { small: 16, medium: 24, large: 40 };
  const dimension = sizeMap[size];
  
  return (
    <div className="spinner-container" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
      <div 
        className="spinner"
        style={{
          width: dimension,
          height: dimension,
          border: '3px solid var(--bg-light, #1a1a24)',
          borderTop: '3px solid var(--neon-green, #39ff14)',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite',
        }}
      />
      {label && <span style={{ color: 'var(--text-secondary, #a0a0a0)', fontSize: '12px' }}>{label}</span>}
    </div>
  );
};

// Simple Button component
export const Button: React.FC<{
  children: React.ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  appearance?: 'primary' | 'secondary' | 'subtle';
  size?: 'small' | 'medium' | 'large';
  icon?: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
}> = ({ 
  children, 
  onClick, 
  disabled, 
  appearance = 'secondary',
  size = 'medium',
  icon,
  className = '',
  style
}) => {
  const baseClass = `native-btn native-btn-${appearance} native-btn-${size} ${className}`;
  
  return (
    <button 
      className={baseClass}
      onClick={onClick}
      disabled={disabled}
      style={style}
    >
      {icon && <span className="btn-icon">{icon}</span>}
      {children}
    </button>
  );
};

// Simple Input component
export const Input: React.FC<{
  value: string;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  placeholder?: string;
  type?: string;
  disabled?: boolean;
  className?: string;
  style?: React.CSSProperties;
  onKeyDown?: (e: React.KeyboardEvent) => void;
}> = ({ 
  value, 
  onChange, 
  placeholder, 
  type = 'text',
  disabled,
  className = '',
  style,
  onKeyDown
}) => {
  return (
    <input
      type={type}
      value={value}
      onChange={onChange}
      placeholder={placeholder}
      disabled={disabled}
      className={`native-input ${className}`}
      style={style}
      onKeyDown={onKeyDown}
    />
  );
};

// Simple Switch/Toggle component
export const Switch: React.FC<{
  checked: boolean;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  label?: string;
  disabled?: boolean;
}> = ({ checked, onChange, label, disabled }) => {
  return (
    <label className="native-switch">
      <input
        type="checkbox"
        checked={checked}
        onChange={onChange}
        disabled={disabled}
      />
      <span className="switch-slider"></span>
      {label && <span className="switch-label">{label}</span>}
    </label>
  );
};

// Simple Checkbox component
export const Checkbox: React.FC<{
  checked: boolean;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  label?: string;
  disabled?: boolean;
}> = ({ checked, onChange, label, disabled }) => {
  return (
    <label className="native-checkbox">
      <input
        type="checkbox"
        checked={checked}
        onChange={onChange}
        disabled={disabled}
      />
      <span className="checkbox-mark"></span>
      {label && <span className="checkbox-label">{label}</span>}
    </label>
  );
};

// Send icon SVG
export const SendIcon: React.FC<{ size?: number }> = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M2.01 21L23 12L2.01 3L2 10L17 12L2 14L2.01 21Z" fill="currentColor"/>
  </svg>
);
