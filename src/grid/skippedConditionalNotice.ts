/**
 * 조건부서식 중 아직 못 칠하는 규칙이 있다는 걸 사용자에게 알리는 문구. 조용히 빠뜨리면 "원래 색이 없는
 * 시트인가?"와 "뷰어가 못 칠하나?"를 구분할 수 없고, 상태 색은 정보라서 특히 알려야 한다.
 */
const KIND_LABELS: Record<string, string> = {
  colorScale: '색조',
  dataBar: '데이터 막대',
  iconSet: '아이콘 집합',
  top10: '상위/하위 N',
  aboveAverage: '평균 기준',
  duplicateValues: '중복 값',
  uniqueValues: '고유 값',
  unsupportedFormula: '지원하지 않는 수식',
  borderOrNumberFormat: '테두리·표시 형식 변경',
}

export function skippedConditionalNotice(skipped: Record<string, number>): string | null {
  const parts = Object.entries(skipped)
    .filter(([, count]) => count > 0)
    .map(([kind, count]) => `${KIND_LABELS[kind] ?? kind} ${count}개`)
  return parts.length > 0 ? `일부 조건부서식은 아직 표시하지 못해요: ${parts.join(', ')}` : null
}
