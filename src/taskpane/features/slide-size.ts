export interface SlideSize {
  width: number;
  height: number;
}

/**
 * PowerPoint exposes real slide dimensions only from PowerPointApi 1.10, above the floor this
 * add-in supports, so every consumer assumes the standard 16:9 canvas.
 */
export const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
