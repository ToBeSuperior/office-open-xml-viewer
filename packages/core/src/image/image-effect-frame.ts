import type { RgbaBuffer } from './duotone.js';

/** The authored image frame within a grown effect grid, in pixels. This is
 * decoded-image geometry, never a retained layout or parser mutation. */
export interface EffectFrame { x: number; y: number; width: number; height: number }
export interface EffectBuffer extends RgbaBuffer { frame?: EffectFrame }
const frames = new WeakMap<object, EffectFrame>();
export function imageEffectFrame(image: object): EffectFrame | undefined { return frames.get(image); }
export function setImageEffectFrame(image: object, frame: EffectFrame | undefined): void {
  if (frame) frames.set(image, frame);
}


/** Whether transformed pixels extend outside their authored picture frame. */
export function hasImageEffectOutset(image: object): boolean {
  const frame = frames.get(image);
  if (!frame) return false;
  const size = image as { width: number; height: number };
  return frame.x > 0 || frame.y > 0 || size.width > frame.x + frame.width || size.height > frame.y + frame.height;
}
