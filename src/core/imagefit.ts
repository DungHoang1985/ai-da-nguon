/** Tính chỗ đặt ảnh trên slide (hàm thuần, đơn vị điểm): chọn vùng theo vị trí được nhờ rồi đặt ảnh vừa vùng đó, giữ đúng tỷ lệ. */

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export type ImagePlace = "content" | "left" | "right" | "center" | "full";

export const IMAGE_PLACES: ImagePlace[] = ["content", "left", "right", "center", "full"];
/** Lề slide và khoảng dành cho tiêu đề (điểm) */
const MARGIN = 40;
const GAP = 20;

const num = (v: any): number | undefined => {
  const n = typeof v === "string" && v.trim() !== "" ? Number(v) : v;
  return typeof n === "number" && Number.isFinite(n) ? n : undefined;
};

/** Vùng dành cho ảnh theo vị trí. titleBottom = mép dưới của ô tiêu đề (không biết thì ước 22% chiều cao slide). */
export function placeRegion(place: ImagePlace, slide: { w: number; h: number }, titleBottom?: number): Box {
  const W = slide.w;
  const H = slide.h;
  if (place === "full") return { left: 0, top: 0, width: W, height: H };
  const tb = num(titleBottom);
  const top = Math.min(Math.max(tb !== undefined && tb > 0 ? tb + GAP / 2 : H * 0.22, MARGIN), H * 0.5);
  const region: Box = { left: MARGIN, top, width: W - 2 * MARGIN, height: H - top - MARGIN };
  if (place === "left" || place === "right") {
    const half = (region.width - GAP) / 2;
    return { ...region, width: half, left: place === "left" ? MARGIN : MARGIN + half + GAP };
  }
  if (place === "center") {
    const w = region.width * 0.6;
    return { ...region, width: w, left: (W - w) / 2 };
  }
  return region;
}

/** Đặt ảnh (rộng × cao, điểm ảnh) vừa một vùng, giữ tỷ lệ, canh giữa vùng. Vị trí "full" thì phủ kín chiều vừa hơn (vẫn nằm trong slide). */
export function fitInBox(img: { w: number; h: number }, box: Box): Box {
  const iw = img.w > 0 ? img.w : 1;
  const ih = img.h > 0 ? img.h : 1;
  const scale = Math.min(box.width / iw, box.height / ih);
  const width = iw * scale;
  const height = ih * scale;
  return { left: box.left + (box.width - width) / 2, top: box.top + (box.height - height) / 2, width, height };
}

/** Khung cuối cùng của ảnh. Người gọi cho đủ left/top/width/height thì dùng số đó làm vùng (ảnh vẫn giữ tỷ lệ bên trong vùng);
 *  chỉ cho width hoặc height thì tính chiều kia theo tỷ lệ. Kết quả luôn nằm trong slide. */
export function imageBox(
  img: { w: number; h: number },
  slide: { w: number; h: number },
  o: { place?: string; left?: any; top?: any; width?: any; height?: any; titleBottom?: number } = {}
): Box {
  const place = (IMAGE_PLACES as string[]).includes(String(o.place)) ? (o.place as ImagePlace) : "content";
  const region = placeRegion(place, slide, o.titleBottom);
  const ratio = (img.w > 0 ? img.w : 1) / (img.h > 0 ? img.h : 1);
  const L = num(o.left);
  const T = num(o.top);
  let Wd = num(o.width);
  let Ht = num(o.height);
  if (Wd !== undefined && Wd <= 0) Wd = undefined;
  if (Ht !== undefined && Ht <= 0) Ht = undefined;
  let box: Box;
  if (Wd !== undefined || Ht !== undefined) {
    const width = Wd ?? (Ht as number) * ratio;
    const height = Ht ?? (Wd as number) / ratio;
    const fitted = fitInBox(img, { left: 0, top: 0, width, height });
    box = { left: L ?? region.left + (region.width - fitted.width) / 2, top: T ?? region.top + (region.height - fitted.height) / 2, width: fitted.width, height: fitted.height };
  } else {
    box = fitInBox(img, region);
    if (L !== undefined) box.left = L;
    if (T !== undefined) box.top = T;
  }
  // Không để ảnh ra ngoài slide: thu nhỏ nếu to hơn slide, rồi kéo vào trong
  const shrink = Math.min(1, slide.w / box.width, slide.h / box.height);
  box.width *= shrink;
  box.height *= shrink;
  box.left = Math.min(Math.max(box.left, 0), slide.w - box.width);
  box.top = Math.min(Math.max(box.top, 0), slide.h - box.height);
  const r = (n: number) => Math.round(n * 100) / 100;
  return { left: r(box.left), top: r(box.top), width: r(box.width), height: r(box.height) };
}

const IMAGE_EXT = /\.(png|jpe?g|gif|webp)$/i;
export function isImageName(name: string): boolean {
  return IMAGE_EXT.test(String(name || ""));
}
