export function StatusBar({ pages, platformKind }: { pages: number | null; platformKind: 'web' | 'desktop' }) {
  return (
    <footer className="statusbar">
      <span>{pages == null ? 'Carregando…' : `${pages} ${pages === 1 ? 'página' : 'páginas'}`}</span>
      <span className="spacer" />
      <span>{platformKind === 'desktop' ? 'Desktop' : 'Navegador'}</span>
    </footer>
  );
}
