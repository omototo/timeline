import { describe, it, expect } from 'vitest';
import { render, screen } from '@testing-library/react';
import { App } from '../src/App.tsx';

describe('App', () => {
  it('renders the heading', () => {
    render(<App />);
    expect(screen.getByRole('heading', { name: 'Timeline' })).toBeInTheDocument();
  });

  it('renders the engine-computed delta for A1 by default', () => {
    render(<App />);
    expect(screen.getByText('A1: 1 → 2')).toBeInTheDocument();
  });

  it('renders "No change" when there is no delta', () => {
    render(<App delta={null} />);
    expect(screen.getByText('No change')).toBeInTheDocument();
  });

  it('renders ∅ placeholders for null before/after values', () => {
    render(<App delta={{ address: 'B2', before: null, after: null }} />);
    expect(screen.getByText('B2: ∅ → ∅')).toBeInTheDocument();
  });
});
