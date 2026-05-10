export type DownloadFileParams = {
  fileName: string
  buffer: ArrayBuffer
}

export interface FileDownloader {
  downloadFile(params: DownloadFileParams): Promise<void>
}
