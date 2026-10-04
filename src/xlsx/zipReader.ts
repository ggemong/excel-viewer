/**
 * 최소 ZIP 리더 — xlsx(=ZIP) 안의 특정 파일 몇 개만 꺼내 읽는 용도다.
 *
 * 왜 필요한가: ExcelJS는 그림 위치·크기는 주지만 자르기(crop)·회전·절대위치 그림·도형·
 * 텍스트상자·묶음·필터 조건 같은 건 읽는 순간 버린다(소스에 파싱 코드 자체가 없다).
 * 그걸 읽으려면 원본 XML을 직접 열어야 해서 ZIP 접근이 필요하다.
 *
 * 왜 라이브러리가 아닌가: 압축 해제는 브라우저 내장 `DecompressionStream`이 해주고,
 * ZIP 컨테이너 구조(중앙 디렉터리)를 읽는 건 수십 줄이라 새 의존성을 들일 이유가 없다
 * (이전에 쓰던 fflate는 D-006에서 제거했다).
 *
 * 범위: ZIP64(4GB 초과)·암호화·분할 압축은 지원하지 않고 ZipUnsupportedError를 던진다.
 * 호출부는 이를 "그림/도형만 못 읽음"으로 격하해서 처리한다(셀 데이터는 ExcelJS 경로로
 * 이미 읽혔으므로 그것까지 막지 않는다).
 */

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_SIGNATURE = 0x02014b50
const LOCAL_SIGNATURE = 0x04034b50
/** EOCD는 파일 끝에서 (고정 22바이트 + 최대 65535바이트 주석) 안에 있다. */
const EOCD_MIN_SIZE = 22
const EOCD_MAX_SEARCH = EOCD_MIN_SIZE + 0xffff
const METHOD_STORED = 0
const METHOD_DEFLATE = 8
const ZIP64_MARKER = 0xffffffff
const ZIP64_COUNT_MARKER = 0xffff

export class ZipUnsupportedError extends Error {}

interface ZipEntry {
  method: number
  compressedSize: number
  localHeaderOffset: number
  flags: number
}

const utf8 = new TextDecoder('utf-8')

export class ZipArchive {
  private readonly bytes: Uint8Array
  private readonly view: DataView
  private readonly entries: Map<string, ZipEntry>

  private constructor(bytes: Uint8Array, view: DataView, entries: Map<string, ZipEntry>) {
    this.bytes = bytes
    this.view = view
    this.entries = entries
  }

  /** 중앙 디렉터리만 읽는다(압축 해제 없음) — 큰 파일에서도 빠르다. */
  static open(buffer: ArrayBuffer): ZipArchive {
    const bytes = new Uint8Array(buffer)
    const view = new DataView(buffer)

    let eocd = -1
    const searchFrom = Math.max(0, bytes.length - EOCD_MAX_SEARCH)
    for (let i = bytes.length - EOCD_MIN_SIZE; i >= searchFrom; i--) {
      if (view.getUint32(i, true) === EOCD_SIGNATURE) {
        eocd = i
        break
      }
    }
    if (eocd < 0) throw new ZipUnsupportedError('ZIP 끝 레코드를 찾지 못했어요(손상된 파일일 수 있어요)')

    const totalEntries = view.getUint16(eocd + 10, true)
    const centralOffset = view.getUint32(eocd + 16, true)
    if (totalEntries === ZIP64_COUNT_MARKER || centralOffset === ZIP64_MARKER) {
      throw new ZipUnsupportedError('ZIP64 형식은 지원하지 않아요')
    }

    const entries = new Map<string, ZipEntry>()
    let p = centralOffset
    for (let n = 0; n < totalEntries; n++) {
      if (view.getUint32(p, true) !== CENTRAL_SIGNATURE) {
        throw new ZipUnsupportedError('ZIP 중앙 디렉터리가 손상됐어요')
      }
      const flags = view.getUint16(p + 8, true)
      const method = view.getUint16(p + 10, true)
      const compressedSize = view.getUint32(p + 20, true)
      const nameLen = view.getUint16(p + 28, true)
      const extraLen = view.getUint16(p + 30, true)
      const commentLen = view.getUint16(p + 32, true)
      const localHeaderOffset = view.getUint32(p + 42, true)
      const name = utf8.decode(bytes.subarray(p + 46, p + 46 + nameLen))
      entries.set(name, { method, compressedSize, localHeaderOffset, flags })
      p += 46 + nameLen + extraLen + commentLen
    }

    return new ZipArchive(bytes, view, entries)
  }

  has(name: string): boolean {
    return this.entries.has(name)
  }

  names(): string[] {
    return [...this.entries.keys()]
  }

  /** 항목 하나의 압축을 풀어 원본 바이트를 돌려준다. 없는 항목이면 null. */
  async read(name: string): Promise<Uint8Array | null> {
    const entry = this.entries.get(name)
    if (!entry) return null
    if (entry.flags & 0x1) throw new ZipUnsupportedError(`암호화된 항목은 읽을 수 없어요: ${name}`)

    const p = entry.localHeaderOffset
    if (this.view.getUint32(p, true) !== LOCAL_SIGNATURE) {
      throw new ZipUnsupportedError(`ZIP 로컬 헤더가 손상됐어요: ${name}`)
    }
    const nameLen = this.view.getUint16(p + 26, true)
    const extraLen = this.view.getUint16(p + 28, true)
    const start = p + 30 + nameLen + extraLen
    const raw = this.bytes.subarray(start, start + entry.compressedSize)

    if (entry.method === METHOD_STORED) return raw
    if (entry.method === METHOD_DEFLATE) return inflateRaw(raw)
    throw new ZipUnsupportedError(`지원하지 않는 압축 방식(${entry.method}): ${name}`)
  }

  async readText(name: string): Promise<string | null> {
    const bytes = await this.read(name)
    return bytes ? utf8.decode(bytes) : null
  }
}

async function inflateRaw(raw: Uint8Array): Promise<Uint8Array> {
  const stream = new Response(raw as BodyInit).body!.pipeThrough(new DecompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}
