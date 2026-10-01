/**
 * 사용자가 입력한 문자열이 "순수 숫자 리터럴"인지 판별한다. 저장 경로(patch.ts, 숫자면
 * <v>로 저장하고 아니면 inlineStr로 저장)와 편집 상태 관리(useWorkbookController.ts,
 * 화면에 표시할 값을 number로 둘지 string으로 둘지 결정)가 반드시 같은 기준으로 판정해야
 * 저장 후 재로드했을 때 편집 중 보이던 값과 실제 저장된 값이 어긋나지 않는다 — 두 곳에서
 * 각자 정규식을 들고 있으면 한쪽만 고쳤을 때 조용히 갈라질 수 있어 여기 하나로 모은다.
 */
export function isPlainNumberLiteral(value: string): boolean {
  return /^-?\d+(\.\d+)?$/.test(value.trim())
}
