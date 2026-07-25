interface AppToolbarProps {
  onNew: () => void
  onOpen: () => void
  onSave: () => void
  onSaveAs: () => void
  onDatePicker: () => void
  onClearDatePicker: () => void
  onToggleFilter: () => void
  onClearFilter: () => void
  onFreezeHeader: () => void
  onFreezeAbove: () => void
  onUnfreeze: () => void
  dirty: boolean
  busy: boolean
}

export function AppToolbar({
  onNew,
  onOpen,
  onSave,
  onSaveAs,
  onDatePicker,
  onClearDatePicker,
  onToggleFilter,
  onClearFilter,
  onFreezeHeader,
  onFreezeAbove,
  onUnfreeze,
  dirty,
  busy
}: AppToolbarProps): JSX.Element {
  return (
    <div className="app-toolbar">
      <div className="brand">Open Office Excel</div>
      <div className="app-toolbar-actions">
        <button type="button" onClick={onNew} disabled={busy}>
          New
        </button>
        <button type="button" onClick={onOpen} disabled={busy}>
          Open
        </button>
        <button type="button" onClick={onSave} disabled={busy}>
          Save{dirty ? ' *' : ''}
        </button>
        <button type="button" onClick={onSaveAs} disabled={busy}>
          Save As
        </button>

        <span className="toolbar-sep" aria-hidden="true" />

        <button
          type="button"
          onClick={onToggleFilter}
          disabled={busy}
          title="Toggle AutoFilter on the selection (or data range)"
        >
          Filter
        </button>
        <button
          type="button"
          onClick={onClearFilter}
          disabled={busy}
          title="Remove AutoFilter from this sheet"
        >
          Clear filter
        </button>

        <span className="toolbar-sep" aria-hidden="true" />

        <button
          type="button"
          onClick={onFreezeHeader}
          disabled={busy}
          title="Freeze the first row as a header"
        >
          Freeze header
        </button>
        <button
          type="button"
          onClick={onFreezeAbove}
          disabled={busy}
          title="Freeze rows above the selected cell (Freeze to this row)"
        >
          Freeze to this row
        </button>
        <button type="button" onClick={onUnfreeze} disabled={busy} title="Unfreeze panes">
          Unfreeze
        </button>

        <span className="toolbar-sep" aria-hidden="true" />

        <button
          type="button"
          onClick={onDatePicker}
          disabled={busy}
          title="Enable calendar dropdown on the selected cells"
        >
          Date picker
        </button>
        <button
          type="button"
          onClick={onClearDatePicker}
          disabled={busy}
          title="Remove calendar dropdown from the selected cells"
        >
          Clear date
        </button>
      </div>
    </div>
  )
}
