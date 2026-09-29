/**
 * 최소한의 ZIP 리더/라이터. xlsx 저장 경로가 "건드리지 않은 항목은 바이트 단위로
 * 원본과 동일"을 보장하려면 fflate의 고수준 unzipSync/zipSync(압축을 다시 검)를
 * 쓸 수 없어서, PKZIP 구조(EOCD/Central Directory/Local File Header)를 직접
 * 다룬다 (구현 계획 §2 참고).
 *
 * 지원 범위: 표준(비-ZIP64) ZIP, compression method 0(store)/8(deflate)만.
 * 그 외(ZIP64, 다른 압축 방식, 스트리밍 data descriptor)는 만나면 조용히
 * 무시하지 않고 에러를 던진다 — xlsx가 실제로 이런 걸 쓰는 경우는 사실상 없고,
 * "확인 안 된 걸 조용히 통과시키지 않는다"는 원칙을 저장 경로에 그대로 적용한다.
 */

const LOCAL_FILE_SIG = 0x04034b50
const CENTRAL_DIR_SIG = 0x02014b50
const EOCD_SIG = 0x06054b50

export interface ZipEntry {
  name: string
  compressionMethod: number
  crc32: number
  compressedSize: number
  uncompressedSize: number
  data: Uint8Array
  versionNeeded: number
  generalPurposeFlag: number
  modTime: number
  modDate: number
  versionMadeBy: number
  internalAttributes: number
  externalAttributes: number
  extraField: Uint8Array
}

const textEncoder = new TextEncoder()
const textDecoder = new TextDecoder()

export function parseZip(buffer: ArrayBuffer): ZipEntry[] {
  const bytes = new Uint8Array(buffer)
  const view = new DataView(buffer)

  const eocdOffset = findEocd(bytes, view)
  const totalEntries = view.getUint16(eocdOffset + 10, true)
  const cdSize = view.getUint32(eocdOffset + 12, true)
  const cdOffset = view.getUint32(eocdOffset + 16, true)

  if (cdOffset + cdSize > eocdOffset) {
    throw new Error('ZIP central directory 범위가 파일 크기와 맞지 않아요 (ZIP64이거나 손상된 파일일 수 있어요).')
  }

  const entries: ZipEntry[] = []
  let pos = cdOffset

  for (let i = 0; i < totalEntries; i++) {
    if (view.getUint32(pos, true) !== CENTRAL_DIR_SIG) {
      throw new Error(`central directory 레코드 ${i}의 시그니처가 올바르지 않아요.`)
    }
    const versionMadeBy = view.getUint16(pos + 4, true)
    const versionNeeded = view.getUint16(pos + 6, true)
    const generalPurposeFlag = view.getUint16(pos + 8, true)
    const compressionMethod = view.getUint16(pos + 10, true)
    const modTime = view.getUint16(pos + 12, true)
    const modDate = view.getUint16(pos + 14, true)
    const crc = view.getUint32(pos + 16, true)
    const compressedSize = view.getUint32(pos + 20, true)
    const uncompressedSize = view.getUint32(pos + 24, true)
    const nameLen = view.getUint16(pos + 28, true)
    const extraLen = view.getUint16(pos + 30, true)
    const commentLen = view.getUint16(pos + 32, true)
    const internalAttributes = view.getUint16(pos + 36, true)
    const externalAttributes = view.getUint32(pos + 38, true)
    const localHeaderOffset = view.getUint32(pos + 42, true)

    if (compressionMethod !== 0 && compressionMethod !== 8) {
      throw new Error(`지원하지 않는 압축 방식(${compressionMethod})이 있는 zip이에요.`)
    }

    const nameBytes = bytes.subarray(pos + 46, pos + 46 + nameLen)
    const name = textDecoder.decode(nameBytes)

    const { dataOffset, extraField } = readLocalHeader(bytes, view, localHeaderOffset)
    const data = bytes.slice(dataOffset, dataOffset + compressedSize)

    entries.push({
      name,
      compressionMethod,
      crc32: crc,
      compressedSize,
      uncompressedSize,
      data,
      versionNeeded,
      generalPurposeFlag,
      modTime,
      modDate,
      versionMadeBy,
      internalAttributes,
      externalAttributes,
      extraField,
    })

    pos += 46 + nameLen + extraLen + commentLen
  }

  return entries
}

function readLocalHeader(
  bytes: Uint8Array,
  view: DataView,
  offset: number,
): { dataOffset: number; extraField: Uint8Array } {
  if (view.getUint32(offset, true) !== LOCAL_FILE_SIG) {
    throw new Error('local file header 시그니처가 올바르지 않아요.')
  }
  const nameLen = view.getUint16(offset + 26, true)
  const extraLen = view.getUint16(offset + 28, true)
  const extraField = bytes.slice(offset + 30 + nameLen, offset + 30 + nameLen + extraLen)
  return { dataOffset: offset + 30 + nameLen + extraLen, extraField }
}

function findEocd(bytes: Uint8Array, view: DataView): number {
  const maxCommentLen = 65535
  const searchStart = Math.max(0, bytes.length - 22 - maxCommentLen)
  for (let i = bytes.length - 22; i >= searchStart; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      return i
    }
  }
  throw new Error('ZIP end-of-central-directory를 못 찾았어요 — 올바른 xlsx 파일이 아닐 수 있어요.')
}

export function buildZip(entries: ZipEntry[]): Uint8Array {
  const nameBytesList = entries.map((e) => textEncoder.encode(e.name))

  let totalSize = 0
  for (let i = 0; i < entries.length; i++) {
    totalSize += 30 + nameBytesList[i].length + entries[i].extraField.length + entries[i].data.length
  }
  const cdStart = totalSize
  for (let i = 0; i < entries.length; i++) {
    totalSize += 46 + nameBytesList[i].length + entries[i].extraField.length
  }
  totalSize += 22 // EOCD, no comment

  const out = new Uint8Array(totalSize)
  const view = new DataView(out.buffer)
  const localOffsets: number[] = []

  let pos = 0
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]
    const nameBytes = nameBytesList[i]
    localOffsets.push(pos)

    view.setUint32(pos, LOCAL_FILE_SIG, true)
    view.setUint16(pos + 4, entry.versionNeeded, true)
    view.setUint16(pos + 6, entry.generalPurposeFlag, true)
    view.setUint16(pos + 8, entry.compressionMethod, true)
    view.setUint16(pos + 10, entry.modTime, true)
    view.setUint16(pos + 12, entry.modDate, true)
    view.setUint32(pos + 14, entry.crc32, true)
    view.setUint32(pos + 18, entry.compressedSize, true)
    view.setUint32(pos + 22, entry.uncompressedSize, true)
    view.setUint16(pos + 26, nameBytes.length, true)
    view.setUint16(pos + 28, entry.extraField.length, true)
    out.set(nameBytes, pos + 30)
    out.set(entry.extraField, pos + 30 + nameBytes.length)
    out.set(entry.data, pos + 30 + nameBytes.length + entry.extraField.length)

    pos += 30 + nameBytes.length + entry.extraField.length + entry.data.length
  }

  pos = cdStart
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]
    const nameBytes = nameBytesList[i]

    view.setUint32(pos, CENTRAL_DIR_SIG, true)
    view.setUint16(pos + 4, entry.versionMadeBy, true)
    view.setUint16(pos + 6, entry.versionNeeded, true)
    view.setUint16(pos + 8, entry.generalPurposeFlag, true)
    view.setUint16(pos + 10, entry.compressionMethod, true)
    view.setUint16(pos + 12, entry.modTime, true)
    view.setUint16(pos + 14, entry.modDate, true)
    view.setUint32(pos + 16, entry.crc32, true)
    view.setUint32(pos + 20, entry.compressedSize, true)
    view.setUint32(pos + 24, entry.uncompressedSize, true)
    view.setUint16(pos + 28, nameBytes.length, true)
    view.setUint16(pos + 30, entry.extraField.length, true)
    view.setUint16(pos + 32, 0, true) // comment length
    view.setUint16(pos + 34, 0, true) // disk number start
    view.setUint16(pos + 36, entry.internalAttributes, true)
    view.setUint32(pos + 38, entry.externalAttributes, true)
    view.setUint32(pos + 42, localOffsets[i], true)
    out.set(nameBytes, pos + 46)
    out.set(entry.extraField, pos + 46 + nameBytes.length)

    pos += 46 + nameBytes.length + entry.extraField.length
  }

  const cdSize = pos - cdStart
  view.setUint32(pos, EOCD_SIG, true)
  view.setUint16(pos + 4, 0, true)
  view.setUint16(pos + 6, 0, true)
  view.setUint16(pos + 8, entries.length, true)
  view.setUint16(pos + 10, entries.length, true)
  view.setUint32(pos + 12, cdSize, true)
  view.setUint32(pos + 16, cdStart, true)
  view.setUint16(pos + 20, 0, true) // comment length

  return out
}
