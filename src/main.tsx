import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { error: Error | null }> {
  state = { error: null as Error | null };
  static getDerivedStateFromError(error: Error) { return { error }; }
  componentDidCatch(error: Error, info: React.ErrorInfo) { console.error('UI crash:', error, info); }
  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="p-8 max-w-xl mx-auto mt-16 bg-white text-slate-800 rounded-2xl border border-rose-200 shadow">
        <h1 className="text-lg font-bold text-rose-700">Something went wrong on this screen</h1>
        <p className="text-sm mt-2 break-words">{this.state.error.message}</p>
        <button className="mt-4 px-4 py-2 rounded-xl bg-slate-900 text-white text-sm font-bold"
          onClick={() => this.setState({ error: null })}>Try again</button>
      </div>
    );
  }
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
