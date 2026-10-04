/**
 * DrawingML/차트 XML을 읽을 때 쓰는 DOM 도우미. 접두사(`a:`, `c:`, `xdr:`)는 파일마다 다를 수 있어
 * 접두사에 의존하지 않고 localName으로만 찾는다. 그림·도형(drawing.ts)과 차트(chart.ts)가 함께 쓴다.
 */

/** `el`의 직계 자식 중 localName이 `name`인 것들. */
export function kids(el: Element | null | undefined, name: string): Element[] {
  return el ? Array.from(el.children).filter((c) => c.localName === name) : []
}

/** `el`의 직계 자식 중 localName이 `name`인 첫 번째. */
export function kid(el: Element | null | undefined, name: string): Element | null {
  return kids(el, name)[0] ?? null
}
