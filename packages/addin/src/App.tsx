import { TimelinePaneContainer } from './ui/TimelinePaneContainer.tsx';
import type { TimelineDataSource } from './ui/data-source.ts';

export interface AppProps {
  readonly source?: TimelineDataSource;
}

export function App({ source }: AppProps = {}): React.JSX.Element {
  return source ? <TimelinePaneContainer source={source} /> : <TimelinePaneContainer />;
}
