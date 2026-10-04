/**
 * 통합문서의 기본 글꼴 크기(pt)를 원본 styles.xml에서 읽는다.
 *
 * 왜 따로 읽나: 칸마다 글꼴 크기(font.size)는 ExcelJS가 주지만, 그 크기를 "기본 글꼴 대비 몇 배"로
 * 환산해 그리려면 기준(기본 글꼴 크기)이 필요하다. 열너비가 "기본 글꼴 숫자 폭" 단위라서 화면의 기본
 * 글자 크기를 기본 글꼴로 놓고 비율로 줄이고 늘려야 열너비와 글자 크기의 비례가 파일과 같아진다.
 * ExcelJS는 이 값을 공개 API로 주지 않는다.
 */
import type { ZipArchive } from './zipReader'
import { parseXml } from './workbookParts'

const STYLES_PATH = 'xl/styles.xml'

/**
 * 첫 번째 셀 서식(cellXfs[0], 모든 칸의 기본 서식 = "표준" 스타일)이 가리키는 글꼴의 크기.
 * 못 읽으면 null — 호출부는 "크기를 모른다"로 다뤄 글자 크기 조정을 하지 않는다(11pt 같은 값을
 * 추정해서 쓰면 파일마다 조용히 틀어진다).
 */
export async function readDefaultFontPt(zip: ZipArchive): Promise<number | null> {
  const xml = await zip.readText(STYLES_PATH)
  if (!xml) return null
  const doc = parseXml(xml)

  const firstXf = doc.getElementsByTagName('cellXfs')[0]?.getElementsByTagName('xf')[0]
  const fontId = Number(firstXf?.getAttribute('fontId') ?? 0)
  const fonts = doc.getElementsByTagName('fonts')[0]?.getElementsByTagName('font')
  const size = Number(fonts?.[Number.isInteger(fontId) ? fontId : 0]?.getElementsByTagName('sz')[0]?.getAttribute('val'))
  return Number.isFinite(size) && size > 0 ? size : null
}
