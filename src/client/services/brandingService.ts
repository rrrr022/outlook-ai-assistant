/**
 * Branding Service
 * Stores and manages user's custom branding, colors, logo, and contact info
 * Persists to localStorage for reuse across sessions
 * 
 * Built by FreedomForged_AI
 */

export interface UserBranding {
  // Company/Personal Info
  companyName: string;
  contactName: string;
  title: string;
  email: string;
  phone: string;
  address: string;
  city: string;
  state: string;
  zipCode: string;
  country: string;
  website: string;
  
  // Branding Colors
  primaryColor: string;      // Main brand color
  secondaryColor: string;    // Secondary/darker shade
  accentColor: string;       // Highlight color
  textColor: string;         // Primary text color
  backgroundColor: string;   // Document background
  
  // Logo
  logoUrl: string;           // Base64 or URL
  logoWidth: number;         // Logo width in pixels
  
  // Document Preferences
  fontFamily: string;
  headerStyle: 'modern' | 'classic' | 'minimal' | 'bold';
  includeWatermark: boolean;
  watermarkText: string;
  
  // Social Links
  linkedIn: string;
  twitter: string;
  
  // Signature
  signatureText: string;
  signatureImage: string;    // Base64 signature image
}

const DEFAULT_BRANDING: UserBranding = {
  companyName: '',
  contactName: '',
  title: '',
  email: '',
  phone: '',
  address: '',
  city: '',
  state: '',
  zipCode: '',
  country: '',
  website: '',
  
  primaryColor: '#2B579A',      // Professional blue
  secondaryColor: '#1F4E79',
  accentColor: '#5B9BD5',
  textColor: '#333333',
  backgroundColor: '#FFFFFF',
  
  logoUrl: '',
  logoWidth: 150,
  
  fontFamily: 'Segoe UI, Arial, sans-serif',
  headerStyle: 'modern',
  includeWatermark: false,
  watermarkText: '',
  
  linkedIn: '',
  twitter: '',
  
  signatureText: '',
  signatureImage: '',
};

const STORAGE_KEY = 'outlook-ai-branding';

/**
 * Load branding from localStorage
 */
export function loadBranding(): UserBranding {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      return { ...DEFAULT_BRANDING, ...JSON.parse(stored) };
    }
  } catch (e) {
    console.error('Failed to load branding:', e);
  }
  return { ...DEFAULT_BRANDING };
}

/**
 * Save branding to localStorage
 */
export function saveBranding(branding: Partial<UserBranding>): UserBranding {
  try {
    const current = loadBranding();
    const updated = { ...current, ...branding };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
    return updated;
  } catch (e) {
    console.error('Failed to save branding:', e);
    return loadBranding();
  }
}

/**
 * Clear all branding
 */
export function clearBranding(): void {
  localStorage.removeItem(STORAGE_KEY);
}

/**
 * Convert image file to base64
 */
export function imageToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/**
 * Get formatted address
 */
export function getFormattedAddress(branding: UserBranding): string {
  const parts = [
    branding.address,
    branding.city,
    branding.state && branding.zipCode ? `${branding.state} ${branding.zipCode}` : branding.state || branding.zipCode,
    branding.country,
  ].filter(Boolean);
  return parts.join(', ');
}

/**
 * Get formatted contact block
 */
export function getContactBlock(branding: UserBranding): string {
  const lines = [];
  if (branding.companyName) lines.push(branding.companyName);
  if (branding.contactName) lines.push(branding.contactName);
  if (branding.title) lines.push(branding.title);
  if (branding.email) lines.push(`📧 ${branding.email}`);
  if (branding.phone) lines.push(`📞 ${branding.phone}`);
  if (branding.website) lines.push(`🌐 ${branding.website}`);
  const address = getFormattedAddress(branding);
  if (address) lines.push(`📍 ${address}`);
  return lines.join('\n');
}

/**
 * Color presets for quick selection
 */
export const COLOR_PRESETS = {
  professional: { primary: '#2B579A', secondary: '#1F4E79', accent: '#5B9BD5', name: 'Professional Blue' },
  corporate: { primary: '#1E3A5F', secondary: '#152C4A', accent: '#4A90D9', name: 'Corporate Navy' },
  modern: { primary: '#6366F1', secondary: '#4F46E5', accent: '#A5B4FC', name: 'Modern Indigo' },
  elegant: { primary: '#374151', secondary: '#1F2937', accent: '#9CA3AF', name: 'Elegant Gray' },
  energetic: { primary: '#DC2626', secondary: '#B91C1C', accent: '#FCA5A5', name: 'Energetic Red' },
  nature: { primary: '#059669', secondary: '#047857', accent: '#6EE7B7', name: 'Nature Green' },
  creative: { primary: '#7C3AED', secondary: '#6D28D9', accent: '#C4B5FD', name: 'Creative Purple' },
  warm: { primary: '#D97706', secondary: '#B45309', accent: '#FCD34D', name: 'Warm Orange' },
  tech: { primary: '#0891B2', secondary: '#0E7490', accent: '#67E8F9', name: 'Tech Cyan' },
  luxury: { primary: '#B8860B', secondary: '#996515', accent: '#FFD700', name: 'Luxury Gold' },
  retro: { primary: '#39FF14', secondary: '#32CD32', accent: '#FFB000', name: 'Retro Neon' },
};

/**
 * Branding service singleton
 */
export const brandingService = {
  load: loadBranding,
  save: saveBranding,
  clear: clearBranding,
  imageToBase64,
  getFormattedAddress,
  getContactBlock,
  presets: COLOR_PRESETS,
  defaults: DEFAULT_BRANDING,
};

export default brandingService;
