import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * 아이콘 선언이 가리키는 파일이 실제로 있는지 검사한다.
 *
 * 왜 필요한가: 배포 서버(Cloudflare Pages)는 없는 경로를 404가 아니라 앱의 첫 화면(HTML)으로 대신 내준다. 그래서 아이콘 파일이
 * 빠져도 오류가 나지 않고 "아이콘이 기본 모양으로 나온다"로만 드러난다(실제로 /favicon.ico가 없어서 윈도우 바로가기가 기본 아이콘이었다).
 */
const ROOT = process.cwd()
const PUBLIC_DIR = join(ROOT, 'public')
const INDEX_HTML = readFileSync(join(ROOT, 'index.html'), 'utf8')
const MANIFEST = JSON.parse(readFileSync(join(PUBLIC_DIR, 'manifest.webmanifest'), 'utf8')) as { icons: { src: string; sizes: string; type: string }[] }

const ICON_LINK_RE = /<link\s+rel="(?:icon|apple-touch-icon)"[^>]*\bhref="([^"]+)"/g
const iconHrefs = [...INDEX_HTML.matchAll(ICON_LINK_RE)].map((m) => m[1])
const onDisk = (webPath: string) => join(PUBLIC_DIR, webPath.replace(/^\//, ''))

describe('아이콘 파일', () => {
  it('index.html의 모든 아이콘 링크가 public에 실제로 있다', () => {
    expect(iconHrefs.length).toBeGreaterThanOrEqual(4)
    for (const href of iconHrefs) expect(existsSync(onDisk(href)), href).toBe(true)
  })

  it('웹 매니페스트의 모든 아이콘이 있고, 선언한 크기와 실제 PNG 크기가 같다', () => {
    for (const icon of MANIFEST.icons) {
      const file = onDisk(icon.src)
      expect(existsSync(file), icon.src).toBe(true)
      const png = readFileSync(file)
      expect(png.subarray(1, 4).toString('ascii'), icon.src).toBe('PNG')
      const [w, h] = icon.sizes.split('x').map(Number)
      // PNG: 8바이트 서명 뒤 IHDR 청크의 너비/높이(빅엔디언 4바이트씩)
      expect([png.readUInt32BE(16), png.readUInt32BE(20)], icon.src).toEqual([w, h])
    }
  })

  it('/favicon.ico가 진짜 ICO 파일이다(없으면 서버가 HTML을 내줘서 바로가기·고정 아이콘이 기본 모양이 된다)', () => {
    const ico = readFileSync(join(PUBLIC_DIR, 'favicon.ico'))
    // ICO 헤더: 예약 0, 종류 1(아이콘), 그리고 이미지 개수
    expect([ico.readUInt16LE(0), ico.readUInt16LE(2)]).toEqual([0, 1])
    expect(ico.readUInt16LE(4)).toBeGreaterThanOrEqual(1)
    expect(iconHrefs).toContain('/favicon.ico')
  })
})
