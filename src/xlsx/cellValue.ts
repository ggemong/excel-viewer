/**
 * ExcelJS가 돌려주는 셀 원시값(cell.value)을 화면 모델의 스칼라 값으로 푼다.
 *
 * ExcelJS는 셀 종류마다 값 모양이 다르다 — 숫자/문자/Date 외에 수식·공유 수식·오류값·
 * 서식 있는 글자(richText)·하이퍼링크가 전부 객체로 온다. 이 중 하나라도 놓치면
 * `String(obj)`로 떨어져 화면에 "[object Object]"가 찍힌다(실제 파일에서 공유 수식
 * 206칸이 이렇게 깨졌다). 그래서 알려진 모양은 전부 여기서 풀고, 모르는 모양은
 * "[object Object]" 대신 눈에 띄는 표식(UNSUPPORTED_VALUE)으로 돌려준다 — 조용히
 * 추정하거나 빈칸으로 숨기지 않고 문제가 있다는 걸 사용자도 볼 수 있게 하기 위해서다.
 */
import { dateToSerial } from './excelDate'
import type { CellModel } from './types'

/** 알려진 어떤 모양도 아닌 값이 올 때 화면에 찍는 표식. */
export const UNSUPPORTED_VALUE = '#UNSUPPORTED'


export interface InterpretedCell {
  value: CellModel['value']
  /** 수식 텍스트(공유 수식 "따라가는 칸"은 ExcelJS가 마스터 주소만 줘서 null). */
  formula: string | null
  /** 하이퍼링크 URL 원문 — 스킴 검증은 호출부(read.ts)가 한다. */
  hyperlink: string | null
}

export { dateToSerial }

function scalar(raw: unknown): CellModel['value'] {
  if (raw === null || raw === undefined) return null
  if (raw instanceof Date) return dateToSerial(raw)
  if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') return raw
  if (typeof raw === 'object') {
    const obj = raw as Record<string, unknown>
    if (typeof obj.error === 'string') return obj.error
    if (Array.isArray(obj.richText)) {
      return (obj.richText as { text?: string }[]).map((run) => run.text ?? '').join('')
    }
  }
  return UNSUPPORTED_VALUE
}

export function interpretCellValue(raw: unknown): InterpretedCell {
  if (raw && typeof raw === 'object' && !(raw instanceof Date)) {
    const obj = raw as Record<string, unknown>

    // 수식: 일반 수식은 formula, "채우기로 복사된" 공유 수식의 따라가는 칸은
    // sharedFormula(마스터 칸 주소)만 온다. 둘 다 화면에 쓸 건 result(캐시된 계산값)다.
    if ('formula' in obj || 'sharedFormula' in obj) {
      return {
        value: scalar(obj.result),
        formula: typeof obj.formula === 'string' ? obj.formula : null,
        hyperlink: null,
      }
    }

    // 외부(관계 ID 기반) 링크만 이 모양으로 온다 — 같은 통합문서 안 다른 시트/셀로
    // 가는 내부 링크는 ExcelJS가 애초에 이 모양으로 안 올려준다(알려진 한계).
    if ('hyperlink' in obj) {
      return {
        value: scalar(obj.text),
        formula: null,
        hyperlink: typeof obj.hyperlink === 'string' ? obj.hyperlink : null,
      }
    }
  }

  return { value: scalar(raw), formula: null, hyperlink: null }
}

/**
 * ExcelJS의 셀 메모(`cell.note`)를 글자 하나로 만든다. 메모는 문자열로 오거나, 실제 파일에서는
 * `{ texts: [{ font, text }, ...] }` 처럼 서식 조각 배열로 온다(굵게/색 조각마다 한 항목). 서식은
 * 화면에 쓰지 않으므로 글자만 이어 붙인다. 글자가 하나도 없으면 undefined.
 */
export function interpretNote(raw: unknown): string | undefined {
  if (typeof raw === 'string') return raw.trim() === '' ? undefined : raw
  if (raw && typeof raw === 'object' && Array.isArray((raw as { texts?: unknown }).texts)) {
    const text = ((raw as { texts: { text?: string }[] }).texts).map((t) => t.text ?? '').join('')
    return text.trim() === '' ? undefined : text
  }
  return undefined
}
