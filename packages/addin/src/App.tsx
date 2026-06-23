import { TimelinePaneContainer } from './ui/TimelinePaneContainer.tsx';
import type { TimelineTheme } from './ui/contract.ts';
import type { TimelineDataSource } from './ui/data-source.ts';

export interface AppProps {
  readonly source?: TimelineDataSource;
  readonly theme?: TimelineTheme;
}

export function App({ source, theme = 'light' }: AppProps = {}): React.JSX.Element {
  return source ? (
    <TimelinePaneContainer source={source} theme={theme} />
  ) : (
    <TimelinePaneContainer theme={theme} />
  );
}
