import { deflateSync, inflateSync } from 'fflate'
import { crc32 } from './crc32'

/** compressionMethod(0=store, 8=deflate)에 맞춰 압축 해제. */
export function decompressEntry(data: Uint8Array, compressionMethod: number): Uint8Array {
  if (compressionMethod === 0) return data
  if (compressionMethod === 8) return inflateSync(data)
  throw new Error(`지원하지 않는 압축 방식(${compressionMethod})이에요.`)
}

/** 원본과 같은 압축 방식으로 다시 압축하고, crc32/크기까지 계산해서 돌려준다. */
export function compressEntry(
  data: Uint8Array,
  compressionMethod: number,
): { data: Uint8Array; crc32: number; compressedSize: number; uncompressedSize: number } {
  const compressed = compressionMethod === 0 ? data : deflateSync(data, { level: 6 })
  return {
    data: compressed,
    crc32: crc32(data),
    compressedSize: compressed.length,
    uncompressedSize: data.length,
  }
}
