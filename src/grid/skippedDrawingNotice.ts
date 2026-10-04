/**
 * 시트에 있지만 아직 못 그리는 개체(차트 등)를 사용자에게 알리는 문구를 만든다.
 * 조용히 빠뜨리면 사용자는 "파일에 원래 없었나?"와 "뷰어가 못 보여주나?"를 구분할 수 없다.
 */
const KIND_LABELS: Record<string, string> = {
  chart: '차트',
  picture: '읽을 수 없는 그림',
  graphicFrame: '표·슬라이서 등',
}

export function skippedDrawingNotice(skipped: Record<string, number>): string | null {
  const parts = Object.entries(skipped)
    .filter(([, count]) => count > 0)
    .map(([kind, count]) => `${KIND_LABELS[kind] ?? kind} ${count}개`)
  return parts.length > 0 ? `이 시트에는 아직 표시하지 못하는 개체가 있어요: ${parts.join(', ')}` : null
}
