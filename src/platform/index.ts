import { isTauri } from '@tauri-apps/api/core';
import { tauriPlatform } from './tauri';
import { webPlatform } from './web';
import type { Platform } from './types';

export const platform: Platform = isTauri() ? tauriPlatform : webPlatform;
export type * from './types';
