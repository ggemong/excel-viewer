import { deflateSync, inflateSync } from 'fflate'
import { crc32 } from './crc32'
import { ZIP_COMPRESSION_DEFLATE, ZIP_COMPRESSION_STORE } from './zip'

/** 원본 압축률과 파일 크기를 비슷하게 유지하는 선에서의 절충값(zlib 기본값과 동일) — 더 올려도 저장 속도만 늦어지고 이득은 작다. */
const DEFLATE_LEVEL = 6

/** compressionMethod에 맞춰 압축 해제. */
export function decompressEntry(data: Uint8Array, compressionMethod: number): Uint8Array {
  if (compressionMethod === ZIP_COMPRESSION_STORE) return data
  if (compressionMethod === ZIP_COMPRESSION_DEFLATE) return inflateSync(data)
  throw new Error(`지원하지 않는 압축 방식(${compressionMethod})이에요.`)
}

/** 원본과 같은 압축 방식으로 다시 압축하고, crc32/크기까지 계산해서 돌려준다. */
export function compressEntry(
  data: Uint8Array,
  compressionMethod: number,
): { data: Uint8Array; crc32: number; compressedSize: number; uncompressedSize: number } {
  const compressed = compressionMethod === ZIP_COMPRESSION_STORE ? data : deflateSync(data, { level: DEFLATE_LEVEL })
  return {
    data: compressed,
    crc32: crc32(data),
    compressedSize: compressed.length,
    uncompressedSize: data.length,
  }
}
