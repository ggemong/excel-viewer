/**
 * "이 행의 높이는 사용자가 직접 정했다"(customHeight)는 정보를 원본 시트 XML에서 읽는다.
 *
 * 왜 필요한가: 줄바꿈 칸이 저장된 행높이보다 길면 행을 키워 다 보여줘야 하지만(글꼴이 파일의 맑은 고딕과
 * 달라 같은 글도 줄 수가 달라진다), 사용자가 일부러 정한 높이(제목 배너, 얇은 구분 행 등)까지 키우면
 * 모양이 망가진다. 둘을 가르는 표지가 customHeight인데 ExcelJS는 읽을 때 이 속성을 버린다.
 */
import type { ZipArchive } from './zipReader'

const ROW_TAG_RE = /<row\b[^>]*>/g
const CUSTOM_HEIGHT_RE = /\bcustomHeight="(?:1|true)"/
const ROW_NUMBER_RE = /\br="(\d+)"/

/**
 * 시트 XML에서 높이를 직접 정한 행 번호(1-based)들.
 *
 * @remarks DOM 파서를 쓰지 않고 `<row ...>` 여는 태그만 훑는다 — 큰 시트의 XML은 수십 MB가 될 수 있어서
 * 전체를 DOM으로 만들면 메모리와 시간이 크게 든다. 필요한 건 여는 태그의 속성뿐이다.
 *
 * @throws 시트 XML을 ZIP에서 찾지 못하면 — 조용히 "직접 정한 행 없음"으로 돌려주면 모든 행을 키워도 되는 것으로
 * 오해하게 되므로 호출부가 알 수 있게 던진다.
 */
export async function readCustomHeightRows(zip: ZipArchive, sheetPath: string): Promise<Set<number>> {
  const xml = await zip.readText(sheetPath)
  if (xml === null) throw new Error(`${sheetPath}를 찾지 못했어요`)
  return customHeightRowsOf(xml)
}

/** readCustomHeightRows의 순수 부분(시트 XML 문자열 -> 행 번호 집합). */
export function customHeightRowsOf(sheetXml: string): Set<number> {
  const rows = new Set<number>()
  for (const [tag] of sheetXml.matchAll(ROW_TAG_RE)) {
    if (!CUSTOM_HEIGHT_RE.test(tag)) continue
    const number = ROW_NUMBER_RE.exec(tag)?.[1]
    if (number) rows.add(Number(number))
  }
  return rows
}
