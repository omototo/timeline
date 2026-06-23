import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { App } from '../src/App.tsx';

describe('App', () => {
  it('mounts the fake-backed timeline task pane', () => {
    render(<App />);

    expect(screen.getByRole('heading', { name: 'Parametric Timeline' })).toBeInTheDocument();
    expect(screen.getByLabelText('Worksheet drill-down')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Switch to what-if branch' })).toBeInTheDocument();
  });

  it('applies the requested theme to the pane', () => {
    const { container } = render(<App theme="dark" />);

    expect(container.querySelector('.timeline-pane')).toHaveAttribute('data-theme', 'dark');
  });
});
