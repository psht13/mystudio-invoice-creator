import type { TemplateLoader } from '../../application/ports/template-loader'

export class FetchTemplateLoader implements TemplateLoader {
  async loadTemplate(templatePath: string): Promise<ArrayBuffer> {
    const response = await fetch(templatePath)

    if (!response.ok) {
      throw new Error(`Failed to load template file: ${templatePath}`)
    }

    return response.arrayBuffer()
  }
}
