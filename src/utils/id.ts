import { randomUUID } from 'crypto';

export const generateId = (): string => {
  try {
    return randomUUID();
  } catch {
    return `id_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
  }
};
