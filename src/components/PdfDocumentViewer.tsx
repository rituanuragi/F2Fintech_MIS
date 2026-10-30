import React, { useState, useEffect, useRef } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import {
  ChevronLeft,
  ChevronRight,
  ZoomIn,
  ZoomOut,
  RotateCw,
  Download,
  ExternalLink,
  Loader2,
  AlertTriangle,
  Maximize2,
  Minimize2,
  RefreshCw,
  FileText,
  Eye,
} from 'lucide-react';

// Configure PDF.js worker
try {
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    'pdfjs-dist/build/pdf.worker.min.mjs',
    import.meta.url
  ).toString();
} catch (e) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
}

interface PdfDocumentViewerProps {
  url: string;
  fileName?: string;
  title?: string;
  className?: string;
}

export const PdfDocumentViewer: React.FC<PdfDocumentViewerProps> = ({
  url,
  fileName = 'document.pdf',
  title = 'PDF Document',
  className = '',
}) => {
  const [numPages, setNumPages] = useState<number>(0);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [scale, setScale] = useState<number>(1.1);
  const [rotation, setRotation] = useState<number>(0);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [loadingProgress, setLoadingProgress] = useState<string>('Loading document...');
  const [error, setError] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'canvas' | 'embed' | 'gdocs'>('canvas');
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const pdfDocRef = useRef<any>(null);
  const renderTaskRef = useRef<any>(null);

  // Backend proxy URL that bypasses CORS and S3 iframe headers
  const proxyUrl = `/api/f2-tickets/proxy-document?url=${encodeURIComponent(url)}`;
  const downloadUrl = `/api/f2-tickets/proxy-document?url=${encodeURIComponent(url)}&download=1&filename=${encodeURIComponent(fileName)}`;
  const googleDocsViewerUrl = `https://docs.google.com/viewer?url=${encodeURIComponent(url)}&embedded=true`;

  // Load the PDF Document
  useEffect(() => {
    let isCancelled = false;
    setIsLoading(true);
    setError(null);
    setLoadingProgress('Connecting to document server...');

    const loadPdf = async () => {
      try {
        // Cancel any pending render
        if (renderTaskRef.current) {
          try {
            renderTaskRef.current.cancel();
          } catch {}
        }

        // Use our proxy URL to fetch arrayBuffer directly, bypassing CORS
        setLoadingProgress('Downloading PDF bytes...');
        const response = await fetch(proxyUrl);
        if (!response.ok) {
          throw new Error(`Failed to fetch PDF (HTTP ${response.status})`);
        }

        const arrayBuffer = await response.arrayBuffer();
        if (isCancelled) return;

        setLoadingProgress('Parsing PDF structures...');
        const loadingTask = pdfjsLib.getDocument({
          data: new Uint8Array(arrayBuffer),
          cMapUrl: `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/cmaps/`,
          cMapPacked: true,
        });

        const doc = await loadingTask.promise;
        if (isCancelled) return;

        pdfDocRef.current = doc;
        setNumPages(doc.numPages);
        setCurrentPage(1);
        setIsLoading(false);
      } catch (err: any) {
        if (isCancelled) return;
        console.warn('PDF.js loading notice:', err);
        setError(err?.message || 'Could not load PDF in canvas mode.');
        setIsLoading(false);
      }
    };

    loadPdf();

    return () => {
      isCancelled = true;
      if (renderTaskRef.current) {
        try {
          renderTaskRef.current.cancel();
        } catch {}
      }
    };
  }, [url, proxyUrl]);

  // Render the current page onto the canvas
  useEffect(() => {
    if (!pdfDocRef.current || viewMode !== 'canvas' || numPages === 0) return;

    let isCancelled = false;

    const renderPage = async () => {
      try {
        if (renderTaskRef.current) {
          try {
            await renderTaskRef.current.cancel();
          } catch {}
        }

        const page = await pdfDocRef.current.getPage(currentPage);
        if (isCancelled) return;

        const canvas = canvasRef.current;
        if (!canvas) return;

        const context = canvas.getContext('2d');
        if (!context) return;

        const viewport = page.getViewport({ scale, rotation });

        canvas.height = viewport.height;
        canvas.width = viewport.width;

        const renderContext = {
          canvasContext: context,
          viewport,
        };

        const task = page.render(renderContext);
        renderTaskRef.current = task;
        await task.promise;
      } catch (renderErr: any) {
        if (renderErr?.name !== 'RenderingCancelledException') {
          console.warn('Page render notice:', renderErr);
        }
      }
    };

    renderPage();

    return () => {
      isCancelled = true;
      if (renderTaskRef.current) {
        try {
          renderTaskRef.current.cancel();
        } catch {}
      }
    };
  }, [currentPage, scale, rotation, viewMode, numPages]);

  const handlePrevPage = () => {
    setCurrentPage((prev) => Math.max(1, prev - 1));
  };

  const handleNextPage = () => {
    setCurrentPage((prev) => Math.min(numPages, prev + 1));
  };

  const handleZoomIn = () => {
    setScale((prev) => Math.min(2.5, Number((prev + 0.2).toFixed(1))));
  };

  const handleZoomOut = () => {
    setScale((prev) => Math.max(0.5, Number((prev - 0.2).toFixed(1))));
  };

  const handleResetZoom = () => {
    setScale(1.1);
  };

  const handleRotate = () => {
    setRotation((prev) => (prev + 90) % 360);
  };

  const toggleFullscreen = () => {
    setIsFullscreen((prev) => !prev);
  };

  return (
    <div
      ref={containerRef}
      className={`flex flex-col bg-slate-900 rounded-2xl overflow-hidden shadow-xl border border-slate-700 text-white ${
        isFullscreen ? 'fixed inset-4 z-50 rounded-2xl' : 'w-full h-full min-h-[520px]'
      } ${className}`}
    >
      {/* Top Header / Control Toolbar */}
      <div className="bg-slate-800/95 border-b border-slate-700/80 px-3.5 py-2.5 flex flex-wrap items-center justify-between gap-2.5">
        {/* Left: View Mode selector */}
        <div className="flex items-center gap-1 bg-slate-900/80 p-0.5 rounded-xl border border-slate-700/60 text-xs">
          <button
            type="button"
            onClick={() => setViewMode('canvas')}
            className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer flex items-center gap-1.5 ${
              viewMode === 'canvas'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
            title="Render using built-in in-app reader (Works in all browsers)"
          >
            <Eye className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">In-App Reader</span>
          </button>

          <button
            type="button"
            onClick={() => setViewMode('embed')}
            className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer flex items-center gap-1.5 ${
              viewMode === 'embed'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
            title="Proxy Embed (Browser Native PDF Plugin)"
          >
            <FileText className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Native Embed</span>
          </button>

          <button
            type="button"
            onClick={() => setViewMode('gdocs')}
            className={`px-2.5 py-1 rounded-lg font-bold transition cursor-pointer flex items-center gap-1.5 ${
              viewMode === 'gdocs'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
            title="Google Docs Cloud Viewer fallback"
          >
            <span className="hidden sm:inline">Google Viewer</span>
          </button>
        </div>

        {/* Center: Pagination & Zoom (When in Canvas Mode) */}
        {viewMode === 'canvas' && numPages > 0 && (
          <div className="flex items-center gap-1.5 bg-slate-900/80 px-2 py-1 rounded-xl border border-slate-700/60">
            {/* Page Nav */}
            <button
              type="button"
              onClick={handlePrevPage}
              disabled={currentPage <= 1}
              className="p-1 rounded-lg hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-transparent text-slate-300 hover:text-white cursor-pointer"
              title="Previous Page"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>

            <span className="text-xs font-mono font-bold text-slate-200 px-1">
              {currentPage} / {numPages}
            </span>

            <button
              type="button"
              onClick={handleNextPage}
              disabled={currentPage >= numPages}
              className="p-1 rounded-lg hover:bg-slate-700 disabled:opacity-30 disabled:hover:bg-transparent text-slate-300 hover:text-white cursor-pointer"
              title="Next Page"
            >
              <ChevronRight className="w-4 h-4" />
            </button>

            <div className="w-px h-4 bg-slate-700 mx-1" />

            {/* Zoom Controls */}
            <button
              type="button"
              onClick={handleZoomOut}
              disabled={scale <= 0.6}
              className="p-1 rounded-lg hover:bg-slate-700 disabled:opacity-30 text-slate-300 hover:text-white cursor-pointer"
              title="Zoom Out"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onClick={handleResetZoom}
              className="text-[11px] font-mono font-bold text-slate-300 hover:text-white px-1 hover:underline cursor-pointer"
              title="Reset Zoom to 100%"
            >
              {Math.round(scale * 100)}%
            </button>

            <button
              type="button"
              onClick={handleZoomIn}
              disabled={scale >= 2.4}
              className="p-1 rounded-lg hover:bg-slate-700 disabled:opacity-30 text-slate-300 hover:text-white cursor-pointer"
              title="Zoom In"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>

            <div className="w-px h-4 bg-slate-700 mx-1" />

            {/* Rotate */}
            <button
              type="button"
              onClick={handleRotate}
              className="p-1 rounded-lg hover:bg-slate-700 text-slate-300 hover:text-white cursor-pointer"
              title="Rotate 90°"
            >
              <RotateCw className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Right: Actions (Download, Open, Fullscreen) */}
        <div className="flex items-center gap-1.5">
          <a
            href={downloadUrl}
            download={fileName}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-700/80 hover:bg-slate-600 text-slate-200 hover:text-white font-bold text-xs transition cursor-pointer"
            title="Download PDF directly to your device"
          >
            <Download className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Download</span>
          </a>

          <a
            href={proxyUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs transition cursor-pointer"
            title="Open clean PDF stream in a new browser tab"
          >
            <ExternalLink className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Open New Tab</span>
          </a>

          <button
            type="button"
            onClick={toggleFullscreen}
            className="p-1.5 rounded-xl bg-slate-700/80 hover:bg-slate-600 text-slate-300 hover:text-white cursor-pointer"
            title={isFullscreen ? 'Exit Fullscreen' : 'Fullscreen Viewer'}
          >
            {isFullscreen ? (
              <Minimize2 className="w-3.5 h-3.5" />
            ) : (
              <Maximize2 className="w-3.5 h-3.5" />
            )}
          </button>
        </div>
      </div>

      {/* Main Content Viewer Area */}
      <div className="flex-1 bg-slate-950 p-4 overflow-auto flex items-center justify-center relative min-h-[460px]">
        {/* Mode 1: In-App Canvas Reader */}
        {viewMode === 'canvas' && (
          <>
            {isLoading && (
              <div className="flex flex-col items-center justify-center gap-3 p-8 text-center text-slate-300">
                <Loader2 className="w-8 h-8 animate-spin text-blue-500" />
                <p className="text-xs font-bold text-slate-300">{loadingProgress}</p>
                <p className="text-[11px] text-slate-500 max-w-sm">
                  Rendering PDF in pure canvas mode to avoid browser iframe restrictions...
                </p>
              </div>
            )}

            {!isLoading && error && (
              <div className="p-6 max-w-md w-full bg-slate-900 border border-slate-800 rounded-2xl text-center space-y-3">
                <AlertTriangle className="w-8 h-8 text-amber-400 mx-auto" />
                <h5 className="font-extrabold text-sm text-slate-200">
                  Canvas Mode Notice
                </h5>
                <p className="text-xs text-slate-400">{error}</p>
                <div className="pt-2 flex flex-col sm:flex-row items-center justify-center gap-2">
                  <button
                    type="button"
                    onClick={() => setViewMode('embed')}
                    className="w-full sm:w-auto px-3.5 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs"
                  >
                    Switch to Native Embed
                  </button>
                  <button
                    type="button"
                    onClick={() => setViewMode('gdocs')}
                    className="w-full sm:w-auto px-3.5 py-1.5 rounded-xl bg-slate-700 hover:bg-slate-600 text-slate-200 font-bold text-xs"
                  >
                    Try Google Docs Viewer
                  </button>
                </div>
              </div>
            )}

            <div
              className={`flex justify-center items-center w-full transition-opacity duration-200 ${
                isLoading || error ? 'hidden' : 'block'
              }`}
            >
              <canvas
                ref={canvasRef}
                className="max-w-full rounded-xl shadow-2xl bg-white border border-slate-800"
              />
            </div>
          </>
        )}

        {/* Mode 2: Proxy Embed (Bypasses S3 CORS and frame headers) */}
        {viewMode === 'embed' && (
          <object
            data={proxyUrl}
            type="application/pdf"
            className="w-full h-full min-h-[580px] rounded-xl border border-slate-800 bg-white"
          >
            <div className="p-8 text-center text-slate-400 space-y-3">
              <FileText className="w-10 h-10 text-blue-400 mx-auto" />
              <p className="text-sm font-bold text-slate-200">
                Your browser plugin is blocking embedded PDF display.
              </p>
              <div className="flex items-center justify-center gap-2">
                <button
                  type="button"
                  onClick={() => setViewMode('canvas')}
                  className="px-4 py-2 rounded-xl bg-blue-600 text-white font-bold text-xs"
                >
                  Use In-App Reader (Canvas)
                </button>
                <a
                  href={proxyUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="px-4 py-2 rounded-xl bg-slate-700 text-slate-200 font-bold text-xs"
                >
                  Open in New Tab
                </a>
              </div>
            </div>
          </object>
        )}

        {/* Mode 3: Google Docs Viewer fallback */}
        {viewMode === 'gdocs' && (
          <iframe
            src={googleDocsViewerUrl}
            title={title}
            className="w-full h-full min-h-[580px] rounded-xl border border-slate-800 bg-white"
          />
        )}
      </div>

      {/* Bottom Status bar */}
      <div className="bg-slate-900 border-t border-slate-800 px-4 py-2 flex flex-wrap items-center justify-between text-[11px] text-slate-400 gap-2">
        <span className="font-mono truncate max-w-sm" title={fileName}>
          {fileName}
        </span>
        <div className="flex items-center gap-3">
          <span>Mode: <strong className="text-blue-400 uppercase">{viewMode}</strong></span>
          {numPages > 0 && <span>Total: {numPages} {numPages === 1 ? 'Page' : 'Pages'}</span>}
        </div>
      </div>
    </div>
  );
};
