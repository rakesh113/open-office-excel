interface StatusBarProps {
  filePath: string | null
  dirty: boolean
  message?: string | null
}

export function StatusBar({ filePath, dirty, message }: StatusBarProps): JSX.Element {
  const name = filePath ? filePath.split(/[/\\]/).pop() : 'Untitled.xlsx'
  return (
    <div className="status-bar">
      <span className="status-file" title={filePath ?? 'Untitled'}>
        {dirty ? '● ' : ''}
        {name}
      </span>
      <span className="status-message">{message ?? 'Ready'}</span>
    </div>
  )
}
