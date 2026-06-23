import { describe, expect, it } from 'vitest';
import { officeThemeToTimelineTheme } from '../src/office-theme.ts';

describe('Office theme bridge', () => {
  it('prefers the Office isDarkTheme signal', () => {
    expect(officeThemeToTimelineTheme({ isDarkTheme: true })).toBe('dark');
    expect(officeThemeToTimelineTheme({ isDarkTheme: false })).toBe('light');
  });

  it('falls back to Office background luminance', () => {
    expect(officeThemeToTimelineTheme({ bodyBackgroundColor: '#202020' })).toBe('dark');
    expect(officeThemeToTimelineTheme({ bodyBackgroundColor: '#FAFAFA' })).toBe('light');
  });

  it('uses the foreground color when Office omits background colors', () => {
    expect(officeThemeToTimelineTheme({ bodyForegroundColor: '#FFFFFF' })).toBe('dark');
    expect(officeThemeToTimelineTheme({ bodyForegroundColor: '#111111' })).toBe('light');
  });
});
