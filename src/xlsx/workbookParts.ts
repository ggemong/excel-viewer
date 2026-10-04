/**
 * xlsx(ZIP) 안 부품들의 경로를 찾아가는 공용 도구 — 통합문서의 시트 목록, 부품 간 관계(rels),
 * 상대 경로 해석. 그림/도형(drawing.ts)과 필터(filter.ts)가 모두 "시트 이름 -> 그 시트의 XML 경로"가
 * 필요해서 한 곳에 모았다(각자 복제하면 경로 규칙이 갈라진다).
 */
import type { ZipArchive } from './zipReader'

export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface Relationship {
  type: string
  target: string
}

export interface SheetPart {
  name: string
  /** ZIP 안의 시트 XML 경로(예: xl/worksheets/sheet3.xml) */
  path: string
}

/** `baseDir`(예: xl/worksheets) 기준 상대 경로 target을 ZIP 내부 절대 경로로. */
export function resolveZipPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = baseDir ? baseDir.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.' && seg !== '') parts.push(seg)
  }
  return parts.join('/')
}

export function dirname(path: string): string {
  const i = path.lastIndexOf('/')
  return i < 0 ? '' : path.slice(0, i)
}

export function basename(path: string): string {
  return path.slice(path.lastIndexOf('/') + 1)
}

export function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror')[0]) throw new Error('XML을 해석하지 못했어요')
  return doc
}

export async function readRelationships(zip: ZipArchive, partPath: string): Promise<Map<string, Relationship>> {
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


/**
 * 통합문서 테마(xl/theme/theme1.xml 등)의 XML 원문. 없으면 null.
 *
 * 왜 ExcelJS의 `workbook.model.themes`를 안 쓰나: `workbook.model`은 읽을 때마다 모든 시트의 모델을 통째로 새로
 * 만드는 getter라서, 테마 한 줄을 얻으려고 불러도 실제 파일(59시트)에서 메인 화면을 1.4초 멈췄다(운영 빌드 실측).
 */
export async function readThemeXml(zip: ZipArchive): Promise<string | null> {
  const rels = await readRelationships(zip, 'xl/workbook.xml')
  for (const rel of rels.values()) {
    if (rel.type.endsWith('/theme')) return zip.readText(resolveZipPath('xl', rel.target))
  }
  return null
}

/** workbook.xml 순서대로 (시트 이름, 시트 XML 경로) 목록. 숨김 시트도 포함한다(걸러내는 건 호출부 몫). */
export async function listSheetParts(zip: ZipArchive): Promise<SheetPart[]> {
  const workbookXml = await zip.readText('xl/workbook.xml')
  if (!workbookXml) return []
  const rels = await readRelationships(zip, 'xl/workbook.xml')

  const parts: SheetPart[] = []
  for (const el of Array.from(parseXml(workbookXml).getElementsByTagName('sheet'))) {
    const name = el.getAttribute('name') ?? ''
    const rId = el.getAttributeNS(R_NS, 'id') ?? el.getAttribute('r:id')
    const target = rId ? rels.get(rId)?.target : undefined
    if (name && target) parts.push({ name, path: resolveZipPath('xl', target) })
  }
  return parts
}
