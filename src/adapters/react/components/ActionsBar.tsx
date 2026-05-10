type ActionsBarProps = {
  isGenerating: boolean
  onResetCurrentPeriod: () => void
  onGenerate: () => void
}

export function ActionsBar({ isGenerating, onResetCurrentPeriod, onGenerate }: ActionsBarProps) {
  return (
    <footer className="actions">
      <button type="button" className="ghost" onClick={onResetCurrentPeriod}>
        Reset period
      </button>
      <button type="button" className="primary" onClick={onGenerate} disabled={isGenerating}>
        {isGenerating ? 'Generating...' : 'Generate & Download XLSX'}
      </button>
    </footer>
  )
}
