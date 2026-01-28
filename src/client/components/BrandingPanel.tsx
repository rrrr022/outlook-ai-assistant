/**
 * Branding Settings Panel
 * UI for users to customize their document branding
 * 
 * Built by FreedomForged_AI
 */

import React, { useState, useEffect } from 'react';
import { brandingService, UserBranding, COLOR_PRESETS } from '../services/brandingService';

interface BrandingPanelProps {
  onClose: () => void;
  onSave: (branding: UserBranding) => void;
}

const BrandingPanel: React.FC<BrandingPanelProps> = ({ onClose, onSave }) => {
  const [branding, setBranding] = useState<UserBranding>(brandingService.load());
  const [activeTab, setActiveTab] = useState<'info' | 'colors' | 'logo' | 'preview'>('info');
  const [saving, setSaving] = useState(false);

  const handleChange = (field: keyof UserBranding, value: string | number | boolean) => {
    setBranding(prev => ({ ...prev, [field]: value }));
  };

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const base64 = await brandingService.imageToBase64(file);
      handleChange('logoUrl', base64);
    }
  };

  const handleSignatureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const base64 = await brandingService.imageToBase64(file);
      handleChange('signatureImage', base64);
    }
  };

  const applyPreset = (presetKey: keyof typeof COLOR_PRESETS) => {
    const preset = COLOR_PRESETS[presetKey];
    setBranding(prev => ({
      ...prev,
      primaryColor: preset.primary,
      secondaryColor: preset.secondary,
      accentColor: preset.accent,
    }));
  };

  const handleSave = () => {
    setSaving(true);
    const saved = brandingService.save(branding);
    setTimeout(() => {
      setSaving(false);
      onSave(saved);
    }, 500);
  };

  const handleReset = () => {
    if (confirm('Reset all branding to defaults?')) {
      brandingService.clear();
      setBranding(brandingService.defaults);
    }
  };

  return (
    <div className="branding-panel">
      <div className="branding-header">
        <h2>🎨 Document Branding</h2>
        <button className="close-btn" onClick={onClose}>×</button>
      </div>

      <div className="branding-tabs">
        <button className={activeTab === 'info' ? 'active' : ''} onClick={() => setActiveTab('info')}>
          📋 My Info
        </button>
        <button className={activeTab === 'colors' ? 'active' : ''} onClick={() => setActiveTab('colors')}>
          🎨 Colors
        </button>
        <button className={activeTab === 'logo' ? 'active' : ''} onClick={() => setActiveTab('logo')}>
          🖼️ Logo
        </button>
        <button className={activeTab === 'preview' ? 'active' : ''} onClick={() => setActiveTab('preview')}>
          👁️ Preview
        </button>
      </div>

      <div className="branding-content">
        {activeTab === 'info' && (
          <div className="tab-content">
            <h3>Company / Personal Information</h3>
            <p className="hint">This info will auto-populate in your documents</p>
            
            <div className="form-grid">
              <div className="form-group full-width">
                <label>Company Name</label>
                <input
                  type="text"
                  value={branding.companyName}
                  onChange={e => handleChange('companyName', e.target.value)}
                  placeholder="Your Company LLC"
                />
              </div>
              
              <div className="form-group">
                <label>Your Name</label>
                <input
                  type="text"
                  value={branding.contactName}
                  onChange={e => handleChange('contactName', e.target.value)}
                  placeholder="John Smith"
                />
              </div>
              
              <div className="form-group">
                <label>Title</label>
                <input
                  type="text"
                  value={branding.title}
                  onChange={e => handleChange('title', e.target.value)}
                  placeholder="Sales Manager"
                />
              </div>
              
              <div className="form-group">
                <label>Email</label>
                <input
                  type="email"
                  value={branding.email}
                  onChange={e => handleChange('email', e.target.value)}
                  placeholder="john@company.com"
                />
              </div>
              
              <div className="form-group">
                <label>Phone</label>
                <input
                  type="tel"
                  value={branding.phone}
                  onChange={e => handleChange('phone', e.target.value)}
                  placeholder="(555) 123-4567"
                />
              </div>
              
              <div className="form-group full-width">
                <label>Street Address</label>
                <input
                  type="text"
                  value={branding.address}
                  onChange={e => handleChange('address', e.target.value)}
                  placeholder="123 Business Ave, Suite 100"
                />
              </div>
              
              <div className="form-group">
                <label>City</label>
                <input
                  type="text"
                  value={branding.city}
                  onChange={e => handleChange('city', e.target.value)}
                  placeholder="New York"
                />
              </div>
              
              <div className="form-group">
                <label>State</label>
                <input
                  type="text"
                  value={branding.state}
                  onChange={e => handleChange('state', e.target.value)}
                  placeholder="NY"
                />
              </div>
              
              <div className="form-group">
                <label>ZIP Code</label>
                <input
                  type="text"
                  value={branding.zipCode}
                  onChange={e => handleChange('zipCode', e.target.value)}
                  placeholder="10001"
                />
              </div>
              
              <div className="form-group">
                <label>Country</label>
                <input
                  type="text"
                  value={branding.country}
                  onChange={e => handleChange('country', e.target.value)}
                  placeholder="USA"
                />
              </div>
              
              <div className="form-group full-width">
                <label>Website</label>
                <input
                  type="url"
                  value={branding.website}
                  onChange={e => handleChange('website', e.target.value)}
                  placeholder="https://yourcompany.com"
                />
              </div>
              
              <div className="form-group">
                <label>LinkedIn</label>
                <input
                  type="url"
                  value={branding.linkedIn}
                  onChange={e => handleChange('linkedIn', e.target.value)}
                  placeholder="linkedin.com/in/yourprofile"
                />
              </div>
              
              <div className="form-group">
                <label>Twitter/X</label>
                <input
                  type="text"
                  value={branding.twitter}
                  onChange={e => handleChange('twitter', e.target.value)}
                  placeholder="@yourhandle"
                />
              </div>
            </div>
          </div>
        )}

        {activeTab === 'colors' && (
          <div className="tab-content">
            <h3>Brand Colors</h3>
            <p className="hint">Choose your brand colors or pick a preset</p>
            
            <div className="color-presets">
              <h4>Quick Presets</h4>
              <div className="preset-grid">
                {Object.entries(COLOR_PRESETS).map(([key, preset]) => (
                  <button
                    key={key}
                    className="preset-btn"
                    onClick={() => applyPreset(key as keyof typeof COLOR_PRESETS)}
                    style={{ 
                      background: `linear-gradient(135deg, ${preset.primary} 0%, ${preset.secondary} 100%)`,
                      border: branding.primaryColor === preset.primary ? '3px solid #fff' : 'none'
                    }}
                    title={preset.name}
                  >
                    <span className="preset-name">{preset.name}</span>
                  </button>
                ))}
              </div>
            </div>
            
            <div className="custom-colors">
              <h4>Custom Colors</h4>
              <div className="color-grid">
                <div className="color-group">
                  <label>Primary Color</label>
                  <div className="color-input-wrapper">
                    <input
                      type="color"
                      value={branding.primaryColor}
                      onChange={e => handleChange('primaryColor', e.target.value)}
                    />
                    <input
                      type="text"
                      value={branding.primaryColor}
                      onChange={e => handleChange('primaryColor', e.target.value)}
                      placeholder="#2B579A"
                    />
                  </div>
                </div>
                
                <div className="color-group">
                  <label>Secondary Color</label>
                  <div className="color-input-wrapper">
                    <input
                      type="color"
                      value={branding.secondaryColor}
                      onChange={e => handleChange('secondaryColor', e.target.value)}
                    />
                    <input
                      type="text"
                      value={branding.secondaryColor}
                      onChange={e => handleChange('secondaryColor', e.target.value)}
                      placeholder="#1F4E79"
                    />
                  </div>
                </div>
                
                <div className="color-group">
                  <label>Accent Color</label>
                  <div className="color-input-wrapper">
                    <input
                      type="color"
                      value={branding.accentColor}
                      onChange={e => handleChange('accentColor', e.target.value)}
                    />
                    <input
                      type="text"
                      value={branding.accentColor}
                      onChange={e => handleChange('accentColor', e.target.value)}
                      placeholder="#5B9BD5"
                    />
                  </div>
                </div>
              </div>
            </div>
            
            <div className="header-style">
              <h4>Header Style</h4>
              <div className="style-options">
                {(['modern', 'classic', 'minimal', 'bold'] as const).map(style => (
                  <label key={style} className={`style-option ${branding.headerStyle === style ? 'selected' : ''}`}>
                    <input
                      type="radio"
                      name="headerStyle"
                      checked={branding.headerStyle === style}
                      onChange={() => handleChange('headerStyle', style)}
                    />
                    <span className="style-preview" data-style={style}>
                      {style.charAt(0).toUpperCase() + style.slice(1)}
                    </span>
                  </label>
                ))}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'logo' && (
          <div className="tab-content">
            <h3>Logo & Signature</h3>
            <p className="hint">Upload your logo to include in documents</p>
            
            <div className="logo-section">
              <h4>Company Logo</h4>
              <div className="upload-area">
                {branding.logoUrl ? (
                  <div className="logo-preview">
                    <img src={branding.logoUrl} alt="Logo" style={{ maxWidth: branding.logoWidth }} />
                    <button className="remove-btn" onClick={() => handleChange('logoUrl', '')}>Remove</button>
                  </div>
                ) : (
                  <label className="upload-label">
                    <input type="file" accept="image/*" onChange={handleLogoUpload} />
                    <span>📤 Click to upload logo</span>
                    <small>PNG, JPG, or SVG (max 500KB)</small>
                  </label>
                )}
              </div>
              
              {branding.logoUrl && (
                <div className="logo-size">
                  <label>Logo Width: {branding.logoWidth}px</label>
                  <input
                    type="range"
                    min="50"
                    max="300"
                    value={branding.logoWidth}
                    onChange={e => handleChange('logoWidth', parseInt(e.target.value))}
                  />
                </div>
              )}
            </div>
            
            <div className="signature-section">
              <h4>Email Signature</h4>
              <textarea
                value={branding.signatureText}
                onChange={e => handleChange('signatureText', e.target.value)}
                placeholder="Best regards,&#10;John Smith&#10;Sales Manager"
                rows={4}
              />
              
              <h4>Signature Image (Optional)</h4>
              <div className="upload-area small">
                {branding.signatureImage ? (
                  <div className="signature-preview">
                    <img src={branding.signatureImage} alt="Signature" />
                    <button className="remove-btn" onClick={() => handleChange('signatureImage', '')}>Remove</button>
                  </div>
                ) : (
                  <label className="upload-label">
                    <input type="file" accept="image/*" onChange={handleSignatureUpload} />
                    <span>📝 Upload signature image</span>
                  </label>
                )}
              </div>
            </div>
            
            <div className="watermark-section">
              <h4>Watermark</h4>
              <label className="checkbox-label">
                <input
                  type="checkbox"
                  checked={branding.includeWatermark}
                  onChange={e => handleChange('includeWatermark', e.target.checked)}
                />
                Include watermark in documents
              </label>
              {branding.includeWatermark && (
                <input
                  type="text"
                  value={branding.watermarkText}
                  onChange={e => handleChange('watermarkText', e.target.value)}
                  placeholder="CONFIDENTIAL"
                />
              )}
            </div>
          </div>
        )}

        {activeTab === 'preview' && (
          <div className="tab-content">
            <h3>Preview</h3>
            <p className="hint">See how your branding will look in documents</p>
            
            <div 
              className="document-preview"
              style={{ 
                borderTop: `4px solid ${branding.primaryColor}`,
                fontFamily: branding.fontFamily 
              }}
            >
              <div className="preview-header" style={{ borderBottom: `2px solid ${branding.primaryColor}` }}>
                {branding.logoUrl && (
                  <img src={branding.logoUrl} alt="Logo" style={{ width: branding.logoWidth, marginBottom: '10px' }} />
                )}
                <h1 style={{ color: branding.primaryColor }}>{branding.companyName || 'Your Company Name'}</h1>
                <p style={{ color: branding.secondaryColor }}>{branding.contactName || 'Your Name'} | {branding.title || 'Your Title'}</p>
              </div>
              
              <div className="preview-body">
                <h2 style={{ color: branding.secondaryColor, borderLeft: `4px solid ${branding.primaryColor}`, paddingLeft: '10px' }}>
                  Sample Section Header
                </h2>
                <p>This is how your document content will appear. <span style={{ background: `${branding.accentColor}40`, padding: '2px 6px', borderRadius: '4px' }}>Highlighted text</span> uses your accent color.</p>
                
                <table style={{ width: '100%', marginTop: '15px', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ background: branding.primaryColor, color: '#fff' }}>
                      <th style={{ padding: '10px', textAlign: 'left' }}>Item</th>
                      <th style={{ padding: '10px', textAlign: 'left' }}>Value</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td style={{ padding: '10px', borderBottom: '1px solid #eee' }}>Example Row</td>
                      <td style={{ padding: '10px', borderBottom: '1px solid #eee' }}>$1,000</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              
              <div className="preview-footer" style={{ borderTop: `1px solid ${branding.primaryColor}`, marginTop: '20px', paddingTop: '15px', fontSize: '12px', color: '#666' }}>
                <p>📧 {branding.email || 'email@company.com'} | 📞 {branding.phone || '(555) 123-4567'}</p>
                <p>🌐 {branding.website || 'www.yourcompany.com'}</p>
                {branding.signatureImage && <img src={branding.signatureImage} alt="Signature" style={{ maxWidth: '150px', marginTop: '10px' }} />}
              </div>
            </div>
          </div>
        )}
      </div>

      <div className="branding-actions">
        <button className="reset-btn" onClick={handleReset}>Reset to Defaults</button>
        <button className="save-btn" onClick={handleSave} disabled={saving}>
          {saving ? '💾 Saving...' : '💾 Save Branding'}
        </button>
      </div>
    </div>
  );
};

export default BrandingPanel;
