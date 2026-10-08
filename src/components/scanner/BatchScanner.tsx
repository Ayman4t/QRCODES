"use client";

import {
  Activity,
  AlertTriangle,
  ArrowDownToLine,
  Camera,
  Check,
  CheckCircle2,
  CircleHelp,
  FileSpreadsheet,
  Focus,
  Import,
  LockKeyhole,
  Pause,
  Play,
  RefreshCw,
  ScanLine,
  ShieldCheck,
  Smartphone,
  Square,
  Upload,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { db, type ManifestItem, type ScanEvent } from "@/lib/db";

const BARCODE_FORMATS = ["code_128", "code_39", "ean_13", "ean_8", "qr_code", "upc_a", "upc_e"];
const DETECTION_INTERVAL_MS = 100;
const TRACK_TIMEOUT_MS = 850;

type Point = { x: number; y: number };
type DetectedCode = {
  format?: string;
  rawValue: string;
  cornerPoints?: Point[];
  boundingBox?: { x: number; y: number; width: number; height: number };
};
type Detector = { detect(source: HTMLVideoElement): Promise<DetectedCode[]> };
type DetectorClass = {
  new (options: { formats: string[] }): Detector;
  getSupportedFormats?: () => Promise<string[]>;
};
type ScanTrack = { x: number; y: number; seenAt: number };

declare global {
  interface Window {
    BarcodeDetector?: DetectorClass;
    barcodeDetectorPolyfill?: { BarcodeDetectorPolyfill?: DetectorClass };
    webkitAudioContext?: typeof AudioContext;
  }
}

function normalizeHeader(value: string): string {
  return value.trim().toLowerCase().replace(/[\s_-]+/g, "");
}

function extractManifest(file: File): Promise<ManifestItem[]> {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellText: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!firstSheet) throw new Error("The file does not contain a worksheet.");

    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { defval: "", raw: false });
    const items = new Map<string, ManifestItem>();

    for (const row of rows) {
      const values = new Map(Object.entries(row).map(([key, value]) => [normalizeHeader(key), value]));
      const rawBarcode = values.get("barcode") ?? values.get("code") ?? values.get("sku") ?? values.get("upc") ?? "";
      const barcode = String(rawBarcode).trim();
      if (!barcode) continue;

      const rawQty = values.get("quantity") ?? values.get("qty") ?? values.get("expectedqty") ?? 1;
      const quantity = Number(String(rawQty).trim() || "1");
      if (!Number.isSafeInteger(quantity) || quantity < 1) {
        throw new Error(`Invalid quantity for barcode ${barcode}. Use a positive whole number.`);
      }

      const rawName = values.get("name") ?? values.get("item") ?? values.get("description") ?? "";
      const name = String(rawName).trim() || "Unnamed item";
      const existing = items.get(barcode);
      if (existing) {
        existing.expectedQty += quantity;
        if (existing.name === "Unnamed item" && name !== "Unnamed item") existing.name = name;
      } else {
        items.set(barcode, { barcode, name, expectedQty: quantity, scannedQty: 0 });
      }
    }

    if (!items.size) throw new Error("No barcodes found. Add a Barcode, Code, or SKU column and try again.");
    return [...items.values()];
  });
}

function formatTime(timestamp: number): string {
  return new Intl.DateTimeFormat(undefined, { hour: "2-digit", minute: "2-digit", second: "2-digit" }).format(timestamp);
}

function centerOf(code: DetectedCode): Point | null {
  if (code.cornerPoints?.length) {
    return {
      x: code.cornerPoints.reduce((sum, point) => sum + point.x, 0) / code.cornerPoints.length,
      y: code.cornerPoints.reduce((sum, point) => sum + point.y, 0) / code.cornerPoints.length,
    };
  }
  if (code.boundingBox) return { x: code.boundingBox.x + code.boundingBox.width / 2, y: code.boundingBox.y + code.boundingBox.height / 2 };
  return null;
}

function drawCode(ctx: CanvasRenderingContext2D, code: DetectedCode, color: string, canvasW: number, canvasH: number, videoW: number, videoH: number) {
  const scale = Math.max(canvasW / videoW, canvasH / videoH);
  const offsetX = (canvasW - videoW * scale) / 2;
  const offsetY = (canvasH - videoH * scale) / 2;
  let points = code.cornerPoints;
  if ((!points || points.length < 4) && code.boundingBox) {
    const box = code.boundingBox;
    points = [
      { x: box.x, y: box.y }, { x: box.x + box.width, y: box.y },
      { x: box.x + box.width, y: box.y + box.height }, { x: box.x, y: box.y + box.height },
    ];
  }
  if (!points?.length) return;

  const mapped = points.map((point) => ({ x: point.x * scale + offsetX, y: point.y * scale + offsetY }));
  ctx.beginPath();
  ctx.moveTo(mapped[0].x, mapped[0].y);
  for (let index = 1; index < mapped.length; index += 1) ctx.lineTo(mapped[index].x, mapped[index].y);
  ctx.closePath();
  ctx.lineWidth = 2.5;
  ctx.strokeStyle = color;
  ctx.fillStyle = `${color}20`;
  ctx.fill();
  ctx.stroke();

  for (const point of mapped) {
    ctx.beginPath();
    ctx.arc(point.x, point.y, 3.2, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
  }

  const minX = Math.min(...mapped.map((point) => point.x));
  const minY = Math.min(...mapped.map((point) => point.y));
  const label = (code.rawValue.length > 24 ? `${code.rawValue.slice(0, 21)}…` : code.rawValue) || "Barcode";
  ctx.font = "600 11px ui-monospace, SFMono-Regular, Menlo, monospace";
  const labelWidth = Math.min(canvasW - 12, ctx.measureText(label).width + 16);
  const badgeX = Math.min(Math.max(6, minX), canvasW - labelWidth - 6);
  const badgeY = Math.max(8, minY - 24);
  ctx.fillStyle = color;
  ctx.fillRect(badgeX, badgeY, labelWidth, 19);
  ctx.fillStyle = "#081018";
  ctx.fillText(label, badgeX + 8, badgeY + 13, labelWidth - 12);
}

export default function BatchScanner() {
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const audioContextRef = useRef<AudioContext | null>(null);
  const tracksRef = useRef<Map<string, ScanTrack[]>>(new Map());
  const manifestRef = useRef<Map<string, ManifestItem>>(new Map());
  const fileInputRef = useRef<HTMLInputElement>(null);
  const installPromptRef = useRef<Event | null>(null);
  const [items, setItems] = useState<ManifestItem[]>([]);
  const [recentEvents, setRecentEvents] = useState<ScanEvent[]>([]);
  const [manifestFileName, setManifestFileName] = useState("");
  const [unexpectedCount, setUnexpectedCount] = useState(0);
  const [cameraActive, setCameraActive] = useState(false);
  const [isScanning, setIsScanning] = useState(false);
  const [isStarting, setIsStarting] = useState(false);
  const [engineName, setEngineName] = useState("Detector idle");
  const [toast, setToast] = useState("");
  const [canInstall, setCanInstall] = useState(false);
  const [isIOS, setIsIOS] = useState(false);

  const manifestMap = useMemo(() => new Map(items.map((item) => [item.barcode, item])), [items]);
  const expectedTotal = items.reduce((total, item) => total + item.expectedQty, 0);
  const scannedTotal = items.reduce((total, item) => total + item.scannedQty, 0);
  const remainingTotal = Math.max(0, expectedTotal - scannedTotal);
  const progress = expectedTotal ? Math.min(100, Math.round((scannedTotal / expectedTotal) * 100)) : 0;

  useEffect(() => { manifestRef.current = manifestMap; }, [manifestMap]);

  useEffect(() => {
    let cancelled = false;
    async function restoreLocalState() {
      try {
        const [storedItems, storedEvents, nameSetting] = await Promise.all([
          db.manifestItems.toArray(),
          db.scanEvents.orderBy("scannedAt").reverse().limit(20).toArray(),
          db.settings.get("manifestName"),
        ]);
        if (cancelled) return;
        setItems(storedItems);
        setRecentEvents(storedEvents);
        setUnexpectedCount(await db.scanEvents.filter((event) => !event.matched).count());
        setManifestFileName(nameSetting?.value ?? "");
      } catch (error) {
        console.error("Could not restore local BatchScan data", error);
        if (!cancelled) setToast("Local storage could not be opened in this browser.");
      }
    }
    void restoreLocalState();
    setIsIOS(/iphone|ipad|ipod/i.test(navigator.userAgent));
    if ("serviceWorker" in navigator) void navigator.serviceWorker.register("/sw.js").catch((error) => console.warn("PWA cache could not be registered", error));
    const onInstall = (event: Event) => { event.preventDefault(); installPromptRef.current = event; setCanInstall(true); };
    window.addEventListener("beforeinstallprompt", onInstall);
    return () => {
      cancelled = true;
      window.removeEventListener("beforeinstallprompt", onInstall);
    };
  }, []);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(""), 3600);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  useEffect(() => () => {
    streamRef.current?.getTracks().forEach((track) => track.stop());
    if (audioContextRef.current && audioContextRef.current.state !== "closed") void audioContextRef.current.close();
  }, []);

  useEffect(() => {
    if (!cameraActive || !videoRef.current || !streamRef.current) return;
    videoRef.current.srcObject = streamRef.current;
    void videoRef.current.play().catch((error) => {
      console.error("Camera preview could not start", error);
      setToast("The camera connected, but the preview could not start. Try reloading the page.");
    });
  }, [cameraActive]);

  const playBeep = useCallback(() => {
    try {
      const AudioContextClass = window.AudioContext ?? window.webkitAudioContext;
      if (!AudioContextClass) return;
      if (!audioContextRef.current) audioContextRef.current = new AudioContextClass();
      const context = audioContextRef.current;
      if (context.state === "suspended") void context.resume();
      const oscillator = context.createOscillator();
      const gain = context.createGain();
      oscillator.type = "sine";
      oscillator.frequency.setValueAtTime(880, context.currentTime);
      gain.gain.setValueAtTime(0.13, context.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, context.currentTime + 0.065);
      oscillator.connect(gain);
      gain.connect(context.destination);
      oscillator.start();
      oscillator.stop(context.currentTime + 0.07);
    } catch (error) {
      console.warn("Audio confirmation is unavailable", error);
    }
  }, []);

  const handleManifestUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = "";
    if (!file) return;
    if (items.length && !window.confirm("Importing a new manifest replaces the current manifest and scan session. Continue?")) return;
    try {
      const newItems = await extractManifest(file);
      await db.transaction("rw", db.manifestItems, db.scanEvents, db.settings, async () => {
        await db.manifestItems.clear();
        await db.scanEvents.clear();
        await db.manifestItems.bulkPut(newItems);
        await db.settings.put({ key: "manifestName", value: file.name });
      });
      setItems(newItems);
      manifestRef.current = new Map(newItems.map((item) => [item.barcode, item]));
      setRecentEvents([]);
      setUnexpectedCount(0);
      setManifestFileName(file.name);
      tracksRef.current.clear();
      setToast(`${newItems.length} unique manifest items imported.`);
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Could not read that manifest file.");
    }
  };

  const startCamera = async () => {
    if (!navigator.mediaDevices?.getUserMedia) {
      setToast("Camera access is unavailable here. Open this app on localhost or a secure HTTPS site.");
      return;
    }
    setIsStarting(true);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } },
        audio: false,
      });
      streamRef.current = stream;
      if (window.AudioContext && !audioContextRef.current) audioContextRef.current = new window.AudioContext();
      tracksRef.current.clear();
      setCameraActive(true);
      setIsScanning(true);
    } catch (error) {
      console.error("Camera initialization failed", error);
      setToast(error instanceof Error && error.name === "NotAllowedError" ? "Allow camera access in your browser settings, then try again." : "Could not start the camera. Check permissions and camera availability.");
    } finally {
      setIsStarting(false);
    }
  };

  const stopCamera = () => {
    streamRef.current?.getTracks().forEach((track) => track.stop());
    streamRef.current = null;
    if (videoRef.current) videoRef.current.srcObject = null;
    setCameraActive(false);
    setIsScanning(false);
    setEngineName("Detector idle");
    tracksRef.current.clear();
  };

  const exportSession = async () => {
    try {
      const events = await db.scanEvents.orderBy("scannedAt").toArray();
      if (!events.length) { setToast("There are no scans to export yet."); return; }
      const quote = (value: string) => `"${value.replaceAll('"', '""')}"`;
      const csv = ["Scanned at,Barcode,Item,Status", ...events.map((entry) => `${quote(new Date(entry.scannedAt).toISOString())},${quote(entry.barcode)},${quote(entry.name)},${entry.matched ? "MATCHED" : "UNEXPECTED"}`)].join("\r\n");
      const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `batchscan-${new Date().toISOString().slice(0, 10)}.csv`;
      anchor.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Could not export scan history", error);
      setToast("Could not export this scan session.");
    }
  };

  const resetSession = async () => {
    if (!items.length) { setToast("Import a manifest before resetting a session."); return; }
    if (!window.confirm("Clear all scan counts and unexpected scans for this manifest?")) return;
    const resetItems = items.map((item) => ({ ...item, scannedQty: 0 }));
    await db.transaction("rw", db.manifestItems, db.scanEvents, async () => {
      await db.manifestItems.bulkPut(resetItems);
      await db.scanEvents.clear();
    });
    setItems(resetItems);
    manifestRef.current = new Map(resetItems.map((item) => [item.barcode, item]));
    setRecentEvents([]);
    setUnexpectedCount(0);
    tracksRef.current.clear();
    setToast("Scan session reset.");
  };

  const installApp = async () => {
    const prompt = installPromptRef.current as (Event & { prompt?: () => Promise<void>; userChoice?: Promise<{ outcome: string }> }) | null;
    if (!prompt?.prompt) {
      setToast("On iPhone/iPad, use Share → Add to Home Screen in Safari.");
      return;
    }
    await prompt.prompt();
    installPromptRef.current = null;
    setCanInstall(false);
  };

  useEffect(() => {
    if (!cameraActive || !isScanning) return;
    let animationFrame = 0;
    let cancelled = false;
    let detector: Detector | null = null;
    let lastDetectionAt = 0;
    let isDetecting = false;
    let reportedDetectorError = false;

    const initialize = async () => {
      try {
        let usesCompatibilityEngine = false;
        if (!window.BarcodeDetector) {
          usesCompatibilityEngine = true;
          await new Promise<void>((resolve, reject) => {
            const existing = document.querySelector<HTMLScriptElement>("script[data-batchscan-detector]");
            const script = existing ?? document.createElement("script");
            const onLoad = () => {
              const polyfill = window.barcodeDetectorPolyfill?.BarcodeDetectorPolyfill;
              if (!polyfill) { reject(new Error("The compatibility engine did not load.")); return; }
              window.BarcodeDetector = polyfill;
              resolve();
            };
            if (window.barcodeDetectorPolyfill?.BarcodeDetectorPolyfill) { onLoad(); return; }
            script.addEventListener("load", onLoad, { once: true });
            script.addEventListener("error", () => reject(new Error("The compatibility engine could not be downloaded.")), { once: true });
            if (!existing) {
              script.src = "https://cdn.jsdelivr.net/npm/@undecaf/barcode-detector-polyfill@0.9.23/dist/index.js";
              script.async = true;
              script.dataset.batchscanDetector = "true";
              document.head.appendChild(script);
            }
          });
        }
        const DetectorConstructor = window.BarcodeDetector;
        if (!DetectorConstructor) throw new Error("No barcode detector is available.");
        const supported = DetectorConstructor.getSupportedFormats ? await DetectorConstructor.getSupportedFormats() : BARCODE_FORMATS;
        const formats = BARCODE_FORMATS.filter((format) => supported.includes(format));
        if (!formats.length) throw new Error("This browser has no support for the selected barcode formats.");
        detector = new DetectorConstructor({ formats });
        if (!cancelled) setEngineName(usesCompatibilityEngine ? "Compatibility engine" : "Device detector");
      } catch (error) {
        console.error("Barcode engine initialization failed", error);
        if (!cancelled) {
          setEngineName("Detector unavailable");
          setToast("Barcode scanning could not initialize. Check your connection for the iPhone compatibility engine.");
        }
        return;
      }

      const scanFrame = async (timestamp: number) => {
        if (cancelled) return;
        animationFrame = requestAnimationFrame(scanFrame);
        if (!detector || isDetecting || timestamp - lastDetectionAt < DETECTION_INTERVAL_MS) return;
        lastDetectionAt = timestamp;

        const video = videoRef.current;
        const canvas = canvasRef.current;
        if (!video || !canvas || video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA || !video.videoWidth || !video.videoHeight) return;
        const bounds = canvas.getBoundingClientRect();
        const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
        const canvasWidth = bounds.width;
        const canvasHeight = bounds.height;
        if (canvas.width !== Math.round(canvasWidth * pixelRatio) || canvas.height !== Math.round(canvasHeight * pixelRatio)) {
          canvas.width = Math.round(canvasWidth * pixelRatio);
          canvas.height = Math.round(canvasHeight * pixelRatio);
        }
        const context = canvas.getContext("2d");
        if (!context) return;
        context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
        context.clearRect(0, 0, canvasWidth, canvasHeight);

        isDetecting = true;
        try {
          const detected = await detector.detect(video);
          const now = Date.now();
          const grouped = new Map<string, DetectedCode[]>();
          for (const code of detected) {
            const value = String(code.rawValue ?? "").trim();
            if (!value) continue;
            const group = grouped.get(value) ?? [];
            group.push({ ...code, rawValue: value });
            grouped.set(value, group);
          }

          const updated = new Map(manifestRef.current);
          const newEvents: ScanEvent[] = [];
          const nextTracks = new Map<string, ScanTrack[]>();
          let shouldBeep = false;

          for (const [value, codes] of grouped) {
            const existingTracks = (tracksRef.current.get(value) ?? []).filter((track) => now - track.seenAt < TRACK_TIMEOUT_MS);
            const usedTracks = new Set<number>();
            const currentTracks: ScanTrack[] = [];
            const item = updated.get(value);

            for (const code of codes) {
              const center = centerOf(code);
              let nearestIndex = -1;
              let nearestDistance = Number.POSITIVE_INFINITY;
              if (center) {
                for (let index = 0; index < existingTracks.length; index += 1) {
                  if (usedTracks.has(index)) continue;
                  const track = existingTracks[index];
                  const distance = Math.hypot((track.x - center.x) / video.videoWidth, (track.y - center.y) / video.videoHeight);
                  if (distance < nearestDistance) { nearestDistance = distance; nearestIndex = index; }
                }
              }

              const isNewTrack = nearestIndex < 0 || nearestDistance > 0.16;
              if (!isNewTrack) usedTracks.add(nearestIndex);
              currentTracks.push({
                x: center?.x ?? existingTracks[nearestIndex]?.x ?? video.videoWidth / 2,
                y: center?.y ?? existingTracks[nearestIndex]?.y ?? video.videoHeight / 2,
                seenAt: now,
              });

              if (isNewTrack) {
                if (item && item.scannedQty < item.expectedQty) {
                  item.scannedQty += 1;
                  shouldBeep = true;
                  newEvents.push({ barcode: value, name: item.name, matched: true, scannedAt: now });
                } else if (!item) {
                  newEvents.push({ barcode: value, name: "Unexpected item", matched: false, scannedAt: now });
                }
              }
            }
            nextTracks.set(value, currentTracks);
          }

          for (const [value, oldTracks] of tracksRef.current) {
            if (grouped.has(value)) continue;
            const stillPresent = oldTracks.filter((track) => now - track.seenAt < TRACK_TIMEOUT_MS);
            if (stillPresent.length) nextTracks.set(value, stillPresent);
          }
          tracksRef.current = nextTracks;

          for (const code of detected) {
            const value = String(code.rawValue ?? "").trim();
            if (!value) continue;
            const item = updated.get(value);
            const color = !item ? "#ff6678" : item.scannedQty >= item.expectedQty ? "#f5b942" : "#39d98a";
            drawCode(context, code, color, canvasWidth, canvasHeight, video.videoWidth, video.videoHeight);
          }

          if (newEvents.length) {
            if (shouldBeep) {
              playBeep();
              if (navigator.vibrate) navigator.vibrate(35);
            }
            if (newEvents.some((event) => !event.matched)) setUnexpectedCount((count) => count + newEvents.filter((event) => !event.matched).length);
            setRecentEvents((events) => [...newEvents.slice().reverse(), ...events].slice(0, 20));
            const nextItems = [...updated.values()];
            manifestRef.current = updated;
            setItems(nextItems);
            void db.transaction("rw", db.manifestItems, db.scanEvents, async () => {
              await db.manifestItems.bulkPut(nextItems);
              await db.scanEvents.bulkAdd(newEvents);
            }).catch((error) => console.error("Could not save scan events locally", error));
          }
        } catch (error) {
          console.error("Barcode detection frame failed", error);
          if (!reportedDetectorError) {
            reportedDetectorError = true;
            setToast("A camera frame could not be scanned. The scanner will keep trying.");
          }
        } finally {
          isDetecting = false;
        }
      };

      animationFrame = requestAnimationFrame(scanFrame);
    };

    void initialize();
    return () => {
      cancelled = true;
      cancelAnimationFrame(animationFrame);
    };
  }, [cameraActive, isScanning, playBeep]);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark"><ScanLine size={21} strokeWidth={2.5} /></div>
          <div>
            <h1 className="brand-name">BatchScan <span style={{ color: "#8190a5", fontWeight: 500 }}>Studio</span></h1>
            <div className="brand-sub">MATRIX BARCODE SCANNER</div>
          </div>
        </div>
        <div className="top-actions">
          <div className={`engine-pill ${cameraActive && engineName !== "Detector unavailable" ? "ready" : ""}`}><span className="engine-dot" />{engineName}</div>
          <div className="privacy-pill"><LockKeyhole size={13} />All processing stays on this device</div>
          {(canInstall || isIOS) && <button className="icon-button" title="Install BatchScan" aria-label="Install BatchScan" onClick={() => void installApp()}><ArrowDownToLine size={16} /></button>}
        </div>
      </header>

      <main className="main-layout">
        <section className="scanner-column" aria-label="Barcode scanner">
          <div className="section-heading">
            <div className="section-label"><span className={`live-indicator ${cameraActive && isScanning ? "active" : ""}`} />Live camera</div>
            <div className="camera-hint">Center items inside the frame for best results</div>
          </div>

          <div className="camera-stage">
            {cameraActive ? (
              <>
                <video ref={videoRef} className="camera-video" autoPlay muted playsInline />
                <canvas ref={canvasRef} className="camera-canvas" aria-hidden="true" />
                <div className="focus-frame" />
                <div className="stage-topline">
                  <span className="camera-tag"><Camera size={13} /> BACK CAMERA</span>
                  <span className="camera-tag muted"><Activity size={13} /> {isScanning ? "SCANNING" : "PAUSED"}</span>
                </div>
                <div className="stage-bottomline">
                  <span className="camera-tag muted"><Focus size={13} /> Hold steady · avoid glare</span>
                  <span className="camera-tag muted">{items.length ? `${items.length} SKUs LOADED` : "NO MANIFEST"}</span>
                </div>
              </>
            ) : (
              <div className="camera-empty">
                <div className="camera-icon-wrap"><Camera size={29} strokeWidth={1.6} /></div>
                <h2>{isStarting ? "Connecting to camera…" : "Ready when you are"}</h2>
                <p>Start the camera to detect several barcodes in one frame. Import a manifest to verify items as you scan.</p>
                <button className="button primary" onClick={() => void startCamera()} disabled={isStarting}><Camera size={15} />{isStarting ? "Starting camera…" : "Start camera"}</button>
              </div>
            )}
          </div>

          <div className="stage-controls">
            {!cameraActive ? (
              <button className="button primary" onClick={() => void startCamera()} disabled={isStarting}><Camera size={15} />{isStarting ? "Starting…" : "Start camera"}</button>
            ) : (
              <>
                <button className="button primary" onClick={() => setIsScanning((value) => !value)}>{isScanning ? <Pause size={15} /> : <Play size={15} />}{isScanning ? "Pause scanning" : "Resume scanning"}</button>
                <button className="button" onClick={stopCamera}><Square size={13} />Stop camera</button>
              </>
            )}
            <button className="button" onClick={() => fileInputRef.current?.click()}><Upload size={14} />Import manifest</button>
            <button className="button" onClick={() => void exportSession()}><ArrowDownToLine size={14} />Export scans</button>
          </div>
        </section>

        <aside className="side-column" aria-label="Session details">
          <section className="card manifest-card">
            <div className="card-header">
              <h2 className="card-title"><FileSpreadsheet size={15} />Manifest</h2>
              <button className="text-link" onClick={() => fileInputRef.current?.click()}>{items.length ? "Replace" : "Import"}</button>
            </div>
            {items.length ? (
              <div className="manifest-loaded">
                <div className="manifest-file">
                  <div className="file-icon"><FileSpreadsheet size={17} /></div>
                  <div className="manifest-file-text"><strong title={manifestFileName}>{manifestFileName || "Active manifest"}</strong><span>{items.length} unique codes · {expectedTotal} units expected</span></div>
                  <CheckCircle2 size={17} color="#39d98a" />
                </div>
                <div className="manifest-progress"><span style={{ width: `${progress}%` }} /></div>
                <div className="progress-caption"><span>{progress}% complete</span><strong>{scannedTotal} / {expectedTotal} scanned</strong></div>
                <button className="text-link" style={{ marginTop: 12 }} onClick={() => void resetSession()}>Reset this scan session</button>
              </div>
            ) : (
              <div className="manifest-empty">
                <FileSpreadsheet size={24} strokeWidth={1.5} />
                <strong>No manifest loaded</strong>
                <p>Import an Excel or CSV file with Barcode, Quantity, and Name columns.</p>
                <button className="button" onClick={() => fileInputRef.current?.click()}><Import size={14} />Choose file</button>
              </div>
            )}
          </section>

          <section className="card">
            <div className="card-header"><h2 className="card-title"><Activity size={15} />Session totals</h2><button className="text-link" onClick={() => void exportSession()}>Export CSV</button></div>
            <div className="stats-grid">
              <div className="stat"><span className="stat-label"><FileSpreadsheet size={11} />Expected</span><strong className="stat-number">{expectedTotal}</strong></div>
              <div className="stat"><span className="stat-label"><Check size={11} />Verified</span><strong className="stat-number green">{scannedTotal}</strong></div>
              <div className="stat"><span className="stat-label"><AlertTriangle size={11} />Unknown</span><strong className="stat-number red">{unexpectedCount}</strong></div>
            </div>
            {items.length > 0 && <div className="progress-caption" style={{ padding: "0 14px 13px" }}><span>Remaining</span><strong style={{ color: "#f5b942" }}>{remainingTotal} units</strong></div>}
          </section>

          <section className="card">
            <div className="card-header"><h2 className="card-title"><ShieldCheck size={15} />Verification key</h2><CircleHelp size={14} color="#66768c" /></div>
            <div className="legend">
              <div className="legend-item"><span className="legend-key"><span className="legend-swatch green" />Matched · needs scan</span><span className="legend-note">Counted once</span></div>
              <div className="legend-item"><span className="legend-key"><span className="legend-swatch amber" />Quota reached</span><span className="legend-note">Already complete</span></div>
              <div className="legend-item"><span className="legend-key"><span className="legend-swatch red" />Unexpected code</span><span className="legend-note">Not in manifest</span></div>
            </div>
          </section>

          <section className="card recent-card">
            <div className="card-header"><h2 className="card-title"><RefreshCw size={14} />Recent scans</h2><span style={{ color: "#75849a", fontSize: 10 }}>{recentEvents.length ? `${recentEvents.length} events` : "Live"}</span></div>
            <div className="recent-list">
              {recentEvents.length ? recentEvents.slice(0, 8).map((entry, index) => (
                <div className="recent-row" key={`${entry.id ?? entry.scannedAt}-${entry.barcode}-${index}`}>
                  <span className={`recent-dot ${entry.matched ? "green" : "red"}`} />
                  <span className="recent-value" title={entry.barcode}>{entry.barcode}</span>
                  <span className="recent-time">{formatTime(entry.scannedAt)}</span>
                </div>
              )) : <div className="recent-empty">Scanned barcodes will appear here.</div>}
            </div>
          </section>
        </aside>
      </main>

      <footer className="footer-note"><span><LockKeyhole size={11} />Your manifest and scans are stored locally in this browser</span><span><Smartphone size={11} />Works best on a phone camera</span></footer>
      <input ref={fileInputRef} className="hidden-input" type="file" accept=".xlsx,.xls,.csv" onChange={(event) => void handleManifestUpload(event)} />
      {toast && <div className="toast" role="status" onClick={() => setToast("")}>{toast}<button aria-label="Dismiss" style={{ marginLeft: 12, border: 0, background: "none", color: "#9eabbc", cursor: "pointer" }}><X size={13} /></button></div>}
      <span className="sr-only" aria-live="polite">{scannedTotal} of {expectedTotal} expected items verified; {unexpectedCount} unexpected codes.</span>
    </div>
  );
}
