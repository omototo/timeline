import { useCallback, useSyncExternalStore } from 'react';
import { FakeTimelineDataSource } from './data-source.ts';
import { sampleTimeline } from './sample-timeline.ts';
import { TimelinePane } from './TimelinePane.tsx';
import type { TimelineCommand, TimelineTheme } from './contract.ts';
import type { TimelineDataSource } from './data-source.ts';

const defaultDataSource = new FakeTimelineDataSource(sampleTimeline);

export interface TimelinePaneContainerProps {
  readonly source?: TimelineDataSource;
  readonly theme?: TimelineTheme;
}

export function TimelinePaneContainer({
  source = defaultDataSource,
  theme = 'light',
}: TimelinePaneContainerProps = {}): React.JSX.Element {
  const view = useSyncExternalStore(
    (listener) => source.subscribe(listener),
    () => source.getView(),
    () => source.getView(),
  );
  const dispatch = useCallback(
    (command: TimelineCommand) => {
      source.dispatch(command);
    },
    [source],
  );

  return <TimelinePane view={view} dispatch={dispatch} theme={theme} />;
}
