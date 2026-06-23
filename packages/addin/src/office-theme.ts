import type { TimelineTheme } from './ui/contract.ts';

export interface OfficeThemeLike {
  readonly bodyBackgroundColor?: string;
  readonly bodyForegroundColor?: string;
  readonly controlBackgroundColor?: string;
  readonly controlForegroundColor?: string;
  readonly isDarkTheme?: boolean;
}

export interface OfficeReadyInfo {
  readonly host?: Office.HostType | string;
  readonly platform?: Office.PlatformType | string;
}

export interface OfficeLike {
  readonly HostType?: {
    readonly Excel?: Office.HostType | string;
  };
  readonly context?: {
    readonly officeTheme?: OfficeThemeLike;
  };
  readonly onReady?: () => Promise<OfficeReadyInfo>;
}

interface RgbColor {
  readonly r: number;
  readonly g: number;
  readonly b: number;
}

export function getGlobalOffice(): OfficeLike | undefined {
  return (globalThis as unknown as { Office?: OfficeLike }).Office;
}

export function officeThemeToTimelineTheme(theme?: OfficeThemeLike | null): TimelineTheme {
  if (theme?.isDarkTheme === true) {
    return 'dark';
  }

  if (theme?.isDarkTheme === false) {
    return 'light';
  }

  const background =
    parseHexColor(theme?.bodyBackgroundColor) ?? parseHexColor(theme?.controlBackgroundColor);

  if (background) {
    return relativeLuminance(background) < 0.45 ? 'dark' : 'light';
  }

  const foreground =
    parseHexColor(theme?.bodyForegroundColor) ?? parseHexColor(theme?.controlForegroundColor);

  return foreground && relativeLuminance(foreground) > 0.65 ? 'dark' : 'light';
}

export function getOfficeTimelineTheme(
  office: OfficeLike | undefined = getGlobalOffice(),
): TimelineTheme {
  return officeThemeToTimelineTheme(office?.context?.officeTheme);
}

function parseHexColor(value: string | undefined): RgbColor | undefined {
  const match = /^#?([0-9a-f]{6})$/i.exec(value ?? '');
  if (!match) {
    return undefined;
  }

  const hex = match[1];
  if (!hex) {
    return undefined;
  }

  return {
    r: Number.parseInt(hex.slice(0, 2), 16),
    g: Number.parseInt(hex.slice(2, 4), 16),
    b: Number.parseInt(hex.slice(4, 6), 16),
  };
}

function relativeLuminance({ r, g, b }: RgbColor): number {
  const linearize = (channel: number): number => {
    const srgb = channel / 255;
    return srgb <= 0.03928 ? srgb / 12.92 : ((srgb + 0.055) / 1.055) ** 2.4;
  };

  return 0.2126 * linearize(r) + 0.7152 * linearize(g) + 0.0722 * linearize(b);
}
