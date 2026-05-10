export interface TemplateLoader {
  loadTemplate(templatePath: string): Promise<ArrayBuffer>
}
