/**
 * 시트 위에 떠 있는 그림·도형(`xl/drawings/drawingN.xml`, DrawingML)을 읽는다.
 *
 * ExcelJS는 그림을 일부만 읽고(위치/크기/바이너리) 자르기·회전·절대위치 그림·도형·텍스트
 * 상자·묶음은 읽는 순간 버린다. 실제 업무 파일(스크린샷 위에 빨간 사각형·화살표·말풍선으로
 * "여기를 고쳐주세요"를 표시한 요청서)은 그 버려지는 쪽이 의미의 절반이라, 원본 XML을
 * 직접 읽는다. 셀 데이터는 여전히 ExcelJS 경로가 담당하고 이 모듈은 그림/도형만 책임진다.
 *
 * 경계: 이 모듈은 XML -> DrawingItem 모델 변환까지만 한다. 실제 화면 좌표 계산(열너비/행높이
 * 필요)은 src/grid/drawingLayout.ts, 그리기는 src/grid/DrawingLayer.tsx.
 *
 * 지원하지 않는 개체(차트·표·슬라이서 등 graphicFrame)는 조용히 버리지 않고 종류별 개수를
 * skipped로 돌려준다 — UI가 "표시하지 못한 개체가 있어요"라고 사용자에게 알릴 수 있게.
 */
import { resolveDrawingColor } from './drawingColor'
import type {
  CellPoint,
  DrawingAnchor,
  DrawingItem,
  DrawingNode,
  DrawingTransform,
  PictureNode,
  RelativeRect,
  ShapeLine,
  ShapeNode,
  ShapeText,
  TextParagraph,
  TextRun,
} from './drawingTypes'
import type { ThemeColors } from './themeColor'
import type { ZipArchive } from './zipReader'

const EMU_PER_PX = 9525
const ROTATION_UNIT = 60000
const PERCENT_UNIT = 100_000
const PX_PER_PT = 96 / 72
const FONT_SIZE_UNIT = 100
const DEFAULT_FONT_PT = 11
const DEFAULT_INSET_EMU = { l: 91440, t: 45720, r: 91440, b: 45720 }
/** 테마에 선 두께 목록이 없을 때 쓰는 Office 2007~2010 기본값(EMU) — lnRef idx 1,2,3. */
const FALLBACK_LINE_WIDTHS_EMU = [9525, 25400, 38100]
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
}

export interface MediaEntry {
  blob: Blob | null
  unsupportedFormat: string | null
}

export interface DrawingParseContext {
  theme: ThemeColors | null
  /** r:embed 관계 ID -> 이미 읽어 둔 그림 데이터. 모르는 ID면 undefined. */
  media: (rId: string) => MediaEntry | undefined
}

export interface SheetDrawings {
  items: DrawingItem[]
  /** 표시하지 못한 개체의 종류별 개수(예: { chart: 2 }). */
  skipped: Record<string, number>
}

// ---------- XML 도우미 (접두사에 의존하지 않고 localName으로만 찾는다) ----------

function kids(el: Element | null | undefined, name: string): Element[] {
  return el ? Array.from(el.children).filter((c) => c.localName === name) : []
}

function kid(el: Element | null | undefined, name: string): Element | null {
  return kids(el, name)[0] ?? null
}

/** `mc:AlternateContent`는 선택 가능한 내용(Choice)으로 풀어서 보여준다. */
function contentChildren(el: Element): Element[] {
  return Array.from(el.children).flatMap((c) => {
    if (c.localName !== 'AlternateContent') return [c]
    const chosen = kid(c, 'Choice') ?? kid(c, 'Fallback')
    return chosen ? contentChildren(chosen) : []
  })
}

function num(el: Element | null | undefined, attr: string): number {
  const n = Number(el?.getAttribute(attr) ?? 0)
  return Number.isFinite(n) ? n : 0
}

const emuToPx = (emu: number) => emu / EMU_PER_PX

// ---------- 변형 / 위치 ----------

interface Xfrm {
  x: number
  y: number
  w: number
  h: number
  chX: number
  chY: number
  chW: number
  chH: number
  rotation: number
  flipH: boolean
  flipV: boolean
}

function truthy(v: string | null): boolean {
  return v === '1' || v === 'true'
}

function parseXfrm(xfrm: Element | null): Xfrm | null {
  if (!xfrm) return null
  const off = kid(xfrm, 'off')
  const ext = kid(xfrm, 'ext')
  const chOff = kid(xfrm, 'chOff')
  const chExt = kid(xfrm, 'chExt')
  const rot = num(xfrm, 'rot') / ROTATION_UNIT
  return {
    x: num(off, 'x'),
    y: num(off, 'y'),
    w: num(ext, 'cx'),
    h: num(ext, 'cy'),
    chX: num(chOff, 'x'),
    chY: num(chOff, 'y'),
    chW: num(chExt, 'cx'),
    chH: num(chExt, 'cy'),
    rotation: ((rot % 360) + 360) % 360,
    flipH: truthy(xfrm.getAttribute('flipH')),
    flipV: truthy(xfrm.getAttribute('flipV')),
  }
}

function transformOf(xfrm: Xfrm | null): DrawingTransform {
  return { rotation: xfrm?.rotation ?? 0, flipH: xfrm?.flipH ?? false, flipV: xfrm?.flipV ?? false }
}

function childNumber(el: Element | null, name: string): number {
  const n = Number(kid(el, name)?.textContent ?? 0)
  return Number.isFinite(n) ? n : 0
}

function parseCellPoint(el: Element | null): CellPoint {
  return {
    col: childNumber(el, 'col'),
    row: childNumber(el, 'row'),
    colOffset: emuToPx(childNumber(el, 'colOff')),
    rowOffset: emuToPx(childNumber(el, 'rowOff')),
  }
}

// ---------- 도형 스타일 (채우기 / 선 / 글자) ----------

function lineWidthPx(theme: ThemeColors | null, styleIdx: number): number {
  const table = theme?.lineWidthsEmu ?? FALLBACK_LINE_WIDTHS_EMU
  return emuToPx(table[Math.min(Math.max(styleIdx, 1), table.length) - 1])
}

function resolveFill(spPr: Element | null, style: Element | null, theme: ThemeColors | null): string | null {
  if (kid(spPr, 'noFill')) return null
  const solid = kid(spPr, 'solidFill')
  if (solid) return resolveDrawingColor(solid, theme)
  const grad = kid(spPr, 'gradFill')
  if (grad) {
    // 그라데이션은 첫 번째 정지점 색 하나로 근사한다(의도적 단순화).
    return resolveDrawingColor(kid(kid(grad, 'gsLst'), 'gs'), theme)
  }
  if (kid(spPr, 'blipFill') || kid(spPr, 'pattFill')) return null

  const fillRef = kid(style, 'fillRef')
  return fillRef && num(fillRef, 'idx') > 0 ? resolveDrawingColor(fillRef, theme) : null
}

function resolveLine(spPr: Element | null, style: Element | null, theme: ThemeColors | null): ShapeLine | null {
  const ln = kid(spPr, 'ln')
  if (kid(ln, 'noFill')) return null

  const lnRef = kid(style, 'lnRef')
  const refIdx = num(lnRef, 'idx')
  const explicitFill = kid(ln, 'solidFill')

  // 선 색이 spPr에 없으면 스타일 참조(lnRef)가 정한다. 둘 다 없으면 선 없음.
  const color = explicitFill ? resolveDrawingColor(explicitFill, theme) : refIdx > 0 ? resolveDrawingColor(lnRef, theme) : null
  if (!color) return null

  const widthAttr = ln?.getAttribute('w')
  const width = widthAttr !== null && widthAttr !== undefined ? emuToPx(Number(widthAttr)) : lineWidthPx(theme, refIdx || 1)

  const dashVal = kid(ln, 'prstDash')?.getAttribute('val') ?? 'solid'
  const dash = dashVal === 'solid' ? 'solid' : dashVal === 'dot' || dashVal === 'sysDot' ? 'dotted' : 'dashed'
  return { color, width, dash }
}

function parseAdjust(spPr: Element | null): Record<string, number> {
  const out: Record<string, number> = {}
  const gds = kids(kid(kid(spPr, 'prstGeom'), 'avLst'), 'gd')
  for (const gd of gds) {
    const m = (gd.getAttribute('fmla') ?? '').match(/^val\s+(-?\d+)/)
    const name = gd.getAttribute('name')
    if (m && name) out[name] = Number(m[1])
  }
  return out
}

function parseText(txBody: Element | null, style: Element | null, theme: ThemeColors | null): ShapeText | null {
  if (!txBody) return null
  const bodyPr = kid(txBody, 'bodyPr')
  const defaultColor = resolveDrawingColor(kid(style, 'fontRef'), theme)

  const paragraphs: TextParagraph[] = kids(txBody, 'p').map((p) => {
    const algn = kid(p, 'pPr')?.getAttribute('algn')
    const runs: TextRun[] = []
    for (const child of Array.from(p.children)) {
      if (child.localName === 'br') {
        runs.push({ text: '\n', sizePx: DEFAULT_FONT_PT * PX_PER_PT, bold: false, italic: false, color: defaultColor })
        continue
      }
      if (child.localName !== 'r' && child.localName !== 'fld') continue
      const rPr = kid(child, 'rPr')
      const text = kid(child, 't')?.textContent ?? ''
      if (!text) continue
      const sz = rPr?.getAttribute('sz')
      runs.push({
        text,
        sizePx: ((sz ? Number(sz) / FONT_SIZE_UNIT : DEFAULT_FONT_PT) * PX_PER_PT),
        bold: truthy(rPr?.getAttribute('b') ?? null),
        italic: truthy(rPr?.getAttribute('i') ?? null),
        color: resolveDrawingColor(kid(rPr, 'solidFill'), theme) ?? defaultColor,
      })
    }
    return { align: algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left', runs }
  })

  if (!paragraphs.some((p) => p.runs.some((r) => r.text.trim() !== '' && r.text !== '\n'))) return null

  const anchor = bodyPr?.getAttribute('anchor')
  const inset = (attr: string, fallback: number) => emuToPx(bodyPr?.hasAttribute(attr) ? num(bodyPr, attr) : fallback)
  return {
    paragraphs,
    vAlign: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
    wrap: bodyPr?.getAttribute('wrap') !== 'none',
    insets: {
      l: inset('lIns', DEFAULT_INSET_EMU.l),
      t: inset('tIns', DEFAULT_INSET_EMU.t),
      r: inset('rIns', DEFAULT_INSET_EMU.r),
      b: inset('bIns', DEFAULT_INSET_EMU.b),
    },
  }
}

// ---------- 노드 파싱 ----------

interface ParsedNode {
  node: DrawingNode
  xfrm: Xfrm | null
}

function countSkipped(skipped: Record<string, number>, kind: string): void {
  skipped[kind] = (skipped[kind] ?? 0) + 1
}

function parsePicture(el: Element, ctx: DrawingParseContext, skipped: Record<string, number>): ParsedNode | null {
  const blipFill = kid(el, 'blipFill')
  const blip = kid(blipFill, 'blip')
  const rId = blip?.getAttributeNS(R_NS, 'embed') ?? blip?.getAttribute('r:embed')
  const media = rId ? ctx.media(rId) : undefined
  if (!media) {
    countSkipped(skipped, 'picture')
    return null
  }

  const src = kid(blipFill, 'srcRect')
  const crop = (attr: string) => Math.max(0, Math.min(1, num(src, attr) / PERCENT_UNIT))
  const spPr = kid(el, 'spPr')
  const xfrm = parseXfrm(kid(spPr, 'xfrm'))
  const cNvPr = kid(kid(el, 'nvPicPr'), 'cNvPr')

  const node: PictureNode = {
    kind: 'picture',
    blob: media.blob,
    unsupportedFormat: media.unsupportedFormat,
    crop: { l: crop('l'), t: crop('t'), r: crop('r'), b: crop('b') },
    alt: cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('name') || '',
    ...transformOf(xfrm),
  }
  return { node, xfrm }
}

function parseShape(el: Element, ctx: DrawingParseContext): ParsedNode {
  const spPr = kid(el, 'spPr')
  const style = kid(el, 'style')
  const xfrm = parseXfrm(kid(spPr, 'xfrm'))
  const node: ShapeNode = {
    kind: 'shape',
    preset: kid(spPr, 'prstGeom')?.getAttribute('prst') ?? (kid(spPr, 'custGeom') ? 'custGeom' : 'rect'),
    adjust: parseAdjust(spPr),
    fill: resolveFill(spPr, style, ctx.theme),
    line: resolveLine(spPr, style, ctx.theme),
    text: parseText(kid(el, 'txBody'), style, ctx.theme),
    ...transformOf(xfrm),
  }
  return { node, xfrm }
}

function parseGroup(el: Element, ctx: DrawingParseContext, skipped: Record<string, number>): ParsedNode {
  const xfrm = parseXfrm(kid(kid(el, 'grpSpPr'), 'xfrm'))
  const children: { rel: RelativeRect; node: DrawingNode }[] = []

  for (const child of contentChildren(el)) {
    const parsed = parseNode(child, ctx, skipped)
    if (!parsed) continue
    // 자식 좌표는 묶음의 "자식 좌표계"(chOff/chExt) 기준이라, 묶음 상자 대비 비율로 바꿔 둔다.
    const cw = xfrm?.chW || 1
    const ch = xfrm?.chH || 1
    const c = parsed.xfrm
    children.push({
      rel: c
        ? { x: (c.x - (xfrm?.chX ?? 0)) / cw, y: (c.y - (xfrm?.chY ?? 0)) / ch, w: c.w / cw, h: c.h / ch }
        : { x: 0, y: 0, w: 1, h: 1 },
      node: parsed.node,
    })
  }
  return { node: { kind: 'group', children, ...transformOf(xfrm) }, xfrm }
}

function parseNode(el: Element, ctx: DrawingParseContext, skipped: Record<string, number>): ParsedNode | null {
  switch (el.localName) {
    case 'pic':
      return parsePicture(el, ctx, skipped)
    case 'sp':
    case 'cxnSp':
      return parseShape(el, ctx)
    case 'grpSp':
      return parseGroup(el, ctx, skipped)
    case 'graphicFrame': {
      const uri = el.getElementsByTagNameNS('*', 'graphicData')[0]?.getAttribute('uri') ?? ''
      countSkipped(skipped, uri.endsWith('/chart') ? 'chart' : 'graphicFrame')
      return null
    }
    default:
      return null
  }
}

// ---------- 공개: drawing XML 한 개 파싱 ----------

export function parseDrawingXml(xml: string, ctx: DrawingParseContext): SheetDrawings {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror')[0]) throw new Error('drawing XML을 해석하지 못했어요')

  const items: DrawingItem[] = []
  const skipped: Record<string, number> = {}

  for (const anchorEl of contentChildren(doc.documentElement)) {
    const kind = anchorEl.localName
    if (kind !== 'twoCellAnchor' && kind !== 'oneCellAnchor' && kind !== 'absoluteAnchor') continue

    let anchor: DrawingAnchor
    if (kind === 'twoCellAnchor') {
      anchor = { kind: 'twoCell', from: parseCellPoint(kid(anchorEl, 'from')), to: parseCellPoint(kid(anchorEl, 'to')) }
    } else {
      const ext = kid(anchorEl, 'ext')
      const width = emuToPx(num(ext, 'cx'))
      const height = emuToPx(num(ext, 'cy'))
      if (kind === 'oneCellAnchor') {
        anchor = { kind: 'oneCell', from: parseCellPoint(kid(anchorEl, 'from')), width, height }
      } else {
        const pos = kid(anchorEl, 'pos')
        anchor = { kind: 'absolute', x: emuToPx(num(pos, 'x')), y: emuToPx(num(pos, 'y')), width, height }
      }
    }

    for (const content of contentChildren(anchorEl)) {
      const parsed = parseNode(content, ctx, skipped)
      if (parsed) items.push({ anchor, node: parsed.node })
    }
  }
  return { items, skipped }
}

// ---------- 공개: 통합문서 전체에서 시트별 그림/도형 읽기 ----------

interface Relationship {
  type: string
  target: string
}

/** `baseDir`(예: xl/worksheets) 기준 상대 경로 target을 ZIP 내부 절대 경로로. */
function resolveZipPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = baseDir ? baseDir.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.' && seg !== '') parts.push(seg)
  }
  return parts.join('/')
}

function dirname(path: string): string {
  const i = path.lastIndexOf('/')
  return i < 0 ? '' : path.slice(0, i)
}

function basename(path: string): string {
  return path.slice(path.lastIndexOf('/') + 1)
}

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror')[0]) throw new Error('XML을 해석하지 못했어요')
  return doc
}

async function readRelationships(zip: ZipArchive, partPath: string): Promise<Map<string, Relationship>> {
  const relsPath = `${dirname(partPath)}/_rels/${basename(partPath)}.rels`
  const xml = await zip.readText(relsPath)
  const out = new Map<string, Relationship>()
  if (!xml) return out
  for (const rel of Array.from(parseXml(xml).getElementsByTagName('Relationship'))) {
    const id = rel.getAttribute('Id')
    if (id) out.set(id, { type: rel.getAttribute('Type') ?? '', target: rel.getAttribute('Target') ?? '' })
  }
  return out
}

async function loadMedia(zip: ZipArchive, path: string, cache: Map<string, Promise<MediaEntry>>): Promise<MediaEntry> {
  const hit = cache.get(path)
  if (hit) return hit
  const promise = (async (): Promise<MediaEntry> => {
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
    const mime = IMAGE_MIME[ext]
    // 브라우저가 못 그리는 형식(EMF/WMF/TIFF/SVG 등)은 바이트를 읽지 않고 형식만 알려준다.
    if (!mime) return { blob: null, unsupportedFormat: ext }
    const bytes = await zip.read(path)
    if (!bytes) return { blob: null, unsupportedFormat: ext }
    return { blob: new Blob([bytes as BlobPart], { type: mime }), unsupportedFormat: null }
  })()
  cache.set(path, promise)
  return promise
}

/**
 * 시트 이름 -> 그 시트의 그림/도형. 시트 하나가 실패해도 다른 시트는 계속 읽는다(셀 데이터는
 * 이미 별도 경로로 읽혔으므로 그림 실패가 파일 열기 전체를 막으면 안 된다) — 실패는 warnings로
 * 돌려줘 UI가 사용자에게 알린다.
 */
export async function readWorkbookDrawings(
  zip: ZipArchive,
  theme: ThemeColors | null,
): Promise<{ bySheet: Map<string, SheetDrawings>; warnings: string[] }> {
  const bySheet = new Map<string, SheetDrawings>()
  const warnings: string[] = []

  const workbookXml = await zip.readText('xl/workbook.xml')
  if (!workbookXml) return { bySheet, warnings }
  const workbookRels = await readRelationships(zip, 'xl/workbook.xml')
  const mediaCache = new Map<string, Promise<MediaEntry>>()

  for (const sheetEl of Array.from(parseXml(workbookXml).getElementsByTagName('sheet'))) {
    const name = sheetEl.getAttribute('name') ?? ''
    const rId = sheetEl.getAttributeNS(R_NS, 'id') ?? sheetEl.getAttribute('r:id')
    const sheetTarget = rId ? workbookRels.get(rId)?.target : undefined
    if (!name || !sheetTarget) continue

    try {
      const sheetPath = resolveZipPath('xl', sheetTarget)
      const sheetRels = await readRelationships(zip, sheetPath)
      const drawingRel = [...sheetRels.values()].find((r) => r.type.endsWith('/drawing'))
      if (!drawingRel) continue

      const drawingPath = resolveZipPath(dirname(sheetPath), drawingRel.target)
      const drawingXml = await zip.readText(drawingPath)
      if (!drawingXml) continue

      const drawingRels = await readRelationships(zip, drawingPath)
      const mediaById = new Map<string, MediaEntry>()
      await Promise.all(
        [...drawingRels.entries()]
          .filter(([, r]) => r.type.endsWith('/image'))
          .map(async ([id, r]) => {
            mediaById.set(id, await loadMedia(zip, resolveZipPath(dirname(drawingPath), r.target), mediaCache))
          }),
      )

      const parsed = parseDrawingXml(drawingXml, { theme, media: (id) => mediaById.get(id) })
      if (parsed.items.length > 0 || Object.keys(parsed.skipped).length > 0) bySheet.set(name, parsed)
    } catch (err) {
      warnings.push(`"${name}" 시트의 그림/도형을 읽지 못했어요: ${err instanceof Error ? err.message : String(err)}`)
    }
  }
  return { bySheet, warnings }
}
