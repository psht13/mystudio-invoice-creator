import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserFileDownloader } from '../src/infrastructure/browser/browser-file-downloader'
import { FetchTemplateLoader } from '../src/infrastructure/browser/fetch-template-loader'

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

afterEach(() => {
  document.body.replaceChildren()
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('FetchTemplateLoader', () => {
  it('fetches the template path and returns the response array buffer', async () => {
    const template = new Uint8Array([1, 2, 3]).buffer
    const fetchMock = vi.fn<typeof fetch>()
    fetchMock.mockResolvedValue(new Response(template, { status: 200 }))
    vi.stubGlobal('fetch', fetchMock)

    await expect(new FetchTemplateLoader().loadTemplate('/templates/template-16th.xlsx')).resolves.toEqual(template)

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetchMock).toHaveBeenCalledWith('/templates/template-16th.xlsx')
  })

  it('throws the existing template load error when fetch returns a non-ok response', async () => {
    const fetchMock = vi.fn<typeof fetch>()
    fetchMock.mockResolvedValue(new Response(null, { status: 404 }))
    vi.stubGlobal('fetch', fetchMock)

    await expect(new FetchTemplateLoader().loadTemplate('/templates/template-1st.xlsx')).rejects.toThrow(
      'Failed to load template file: /templates/template-1st.xlsx',
    )
  })
})

describe('BrowserFileDownloader', () => {
  it('creates an XLSX blob, downloads through a temporary anchor, and revokes the object URL', async () => {
    const events: string[] = []
    const fileBuffer = new Uint8Array([4, 5, 6]).buffer
    const originalAppendChild = document.body.appendChild.bind(document.body)
    const originalRemove = Element.prototype.remove

    const createObjectURLMock = vi.fn((blob: Blob) => {
      events.push('createObjectURL')
      expect(blob.type).toBe(XLSX_MIME_TYPE)
      return 'blob:invoice'
    })
    const revokeObjectURLMock = vi.fn((url: string) => {
      events.push('revokeObjectURL')
      expect(url).toBe('blob:invoice')
    })
    Object.defineProperty(URL, 'createObjectURL', {
      configurable: true,
      value: createObjectURLMock,
    })
    Object.defineProperty(URL, 'revokeObjectURL', {
      configurable: true,
      value: revokeObjectURLMock,
    })

    const createElementSpy = vi.spyOn(document, 'createElement')
    const appendChildSpy = vi.spyOn(document.body, 'appendChild').mockImplementation(<T extends Node>(node: T) => {
      events.push('appendChild')
      return originalAppendChild(node) as T
    })
    const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function click() {
      events.push('click')
      expect(document.body.contains(this)).toBe(true)
    })
    const removeSpy = vi.spyOn(Element.prototype, 'remove').mockImplementation(function remove() {
      events.push('remove')
      originalRemove.call(this)
    })

    await new BrowserFileDownloader().downloadFile({
      fileName: 'Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx',
      buffer: fileBuffer,
    })

    const blob = createObjectURLMock.mock.calls[0][0]
    await expect(blob.arrayBuffer()).resolves.toEqual(fileBuffer)

    const link = appendChildSpy.mock.calls[0][0] as HTMLAnchorElement
    expect(createElementSpy).toHaveBeenCalledWith('a')
    expect(link.getAttribute('href')).toBe('blob:invoice')
    expect(link.download).toBe('Pavlo Yurchenko Time Tracking - 2_16_2026.xlsx')
    expect(clickSpy).toHaveBeenCalledTimes(1)
    expect(removeSpy).toHaveBeenCalledTimes(1)
    expect(document.body.contains(link)).toBe(false)
    expect(revokeObjectURLMock).toHaveBeenCalledWith('blob:invoice')
    expect(events).toEqual(['createObjectURL', 'appendChild', 'click', 'remove', 'revokeObjectURL'])
  })
})
