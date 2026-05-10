import type { DownloadFileParams, FileDownloader } from '../../application/ports/file-downloader'

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

export class BrowserFileDownloader implements FileDownloader {
  async downloadFile(params: DownloadFileParams): Promise<void> {
    const blob = new Blob([params.buffer], {
      type: XLSX_MIME_TYPE,
    })

    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = params.fileName
    document.body.appendChild(link)
    link.click()
    link.remove()
    URL.revokeObjectURL(url)
  }
}
