import { useEffect, useRef, useState, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import {
  createChart, IChartApi, ISeriesApi,
  CandlestickData, Time, ColorType,
} from 'lightweight-charts'
import { stockApi } from '../services/api'
import {
  formatMarketCap, formatNum, formatPct,
  formatVolume, pctColor,
} from '../utils/formatters'
import { formatDistanceToNow } from 'date-fns'

// ── Timeframe config ────────────────────────────────────────────────────────
interface TF { key: string; label: string; group: string }
const TIMEFRAMES: TF[] = [
  { key: '1m',  label: '1m',  group: 'intraday' },
  { key: '5m',  label: '5m',  group: 'intraday' },
  { key: '15m', label: '15m', group: 'intraday' },
  { key: '1h',  label: '1H',  group: 'intraday' },
  { key: '4h',  label: '4H',  group: 'intraday' },
  { key: '1D',  label: '1D',  group: 'swing' },
  { key: '1W',  label: '1W',  group: 'swing' },
  { key: '1M',  label: '1M',  group: 'swing' },
  { key: '1Y',  label: '1Y',  group: 'long' },
  { key: '5Y',  label: '5Y',  group: 'long' },
  { key: 'MAX', label: 'MAX', group: 'long' },
]

// ── Helpers ─────────────────────────────────────────────────────────────────
const fmt = (n: number | null | undefined, d = 2) =>
  n == null ? '—' : n.toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d })

const fmtPrice = (n: number | null | undefined) =>
  n == null ? '—' : `₹${fmt(n)}`

function MetricRow({ label, value, color }: { label: string; value: string; color?: string }) {
  return (
    <div className="flex justify-between items-baseline py-1.5 border-b border-brand-border last:border-0">
      <span className="text-xs text-slate-500">{label}</span>
      <span className={`text-xs font-mono font-medium ${color || 'text-slate-200'}`}>{value}</span>
    </div>
  )
}

function Card({ title, children, className = '' }: { title?: string; children: React.ReactNode; className?: string }) {
  return (
    <div className={`card ${className}`}>
      {title && <div className="text-xs text-slate-500 uppercase tracking-wider font-semibold mb-3">{title}</div>}
      {children}
    </div>
  )
}

// ── OI Panel ────────────────────────────────────────────────────────────────
function OIPanel({ oi }: { oi: any }) {
  if (!oi || !oi.available) {
    return (
      <Card title="Open Interest">
        <div className="text-slate-600 text-xs py-4 text-center">
          Not an F&O stock or OI data unavailable
        </div>
      </Card>
    )
  }

  const pcrColor = oi.pcr >= 1.2 ? 'text-up' : oi.pcr <= 0.8 ? 'text-down' : 'text-yellow-400'

  return (
    <Card title="Open Interest (F&O)">
      {/* Summary row */}
      <div className="grid grid-cols-3 gap-3 mb-4">
        <div className="text-center">
          <div className="text-xs text-slate-500 mb-1">Total Call OI</div>
          <div className="text-sm font-mono text-down font-semibold">{formatVolume(oi.total_call_oi)}</div>
        </div>
        <div className="text-center">
          <div className="text-xs text-slate-500 mb-1">PCR</div>
          <div className={`text-lg font-mono font-bold ${pcrColor}`}>{oi.pcr?.toFixed(2) ?? '—'}</div>
          <div className="text-xs text-slate-600">
            {oi.pcr >= 1.2 ? 'Bullish' : oi.pcr <= 0.8 ? 'Bearish' : 'Neutral'}
          </div>
        </div>
        <div className="text-center">
          <div className="text-xs text-slate-500 mb-1">Total Put OI</div>
          <div className="text-sm font-mono text-up font-semibold">{formatVolume(oi.total_put_oi)}</div>
        </div>
      </div>

      {/* Futures */}
      {oi.futures?.length > 0 && (
        <div className="mb-4">
          <div className="text-xs text-slate-500 mb-2">Futures OI</div>
          <div className="space-y-1">
            {oi.futures.map((f: any, i: number) => (
              <div key={i} className="flex justify-between text-xs bg-brand rounded px-2 py-1.5">
                <span className="text-slate-400">{f.expiry}</span>
                <span className="font-mono">{formatVolume(f.oi)}</span>
                <span className={`font-mono ${f.change_oi >= 0 ? 'text-up' : 'text-down'}`}>
                  {f.change_oi >= 0 ? '+' : ''}{formatVolume(f.change_oi)}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Top strikes */}
      {(oi.top_call_strikes?.length > 0 || oi.top_put_strikes?.length > 0) && (
        <div>
          <div className="text-xs text-slate-500 mb-2">
            High OI Strikes — {oi.near_expiry}
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <div className="text-xs text-down mb-1">Calls</div>
              {oi.top_call_strikes?.slice(0, 5).map((s: any, i: number) => (
                <div key={i} className="flex justify-between text-xs py-0.5">
                  <span className="font-mono text-slate-300">{s.strike}</span>
                  <span className="text-slate-500 font-mono">{formatVolume(s.oi || s.openInterest)}</span>
                </div>
              ))}
            </div>
            <div>
              <div className="text-xs text-up mb-1">Puts</div>
              {oi.top_put_strikes?.slice(0, 5).map((s: any, i: number) => (
                <div key={i} className="flex justify-between text-xs py-0.5">
                  <span className="font-mono text-slate-300">{s.strike}</span>
                  <span className="text-slate-500 font-mono">{formatVolume(s.oi || s.openInterest)}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </Card>
  )
}

// ── News Panel ───────────────────────────────────────────────────────────────
function NewsPanel({ articles }: { articles: any[] }) {
  const badge = (imp: string) => ({
    high:   'bg-red-900 text-red-300 border-red-700',
    medium: 'bg-yellow-900 text-yellow-300 border-yellow-700',
    low:    'bg-slate-800 text-slate-400 border-slate-600',
  }[imp] || 'bg-slate-800 text-slate-400 border-slate-600')

  return (
    <Card title={`News (${articles.length} articles)`}>
      {articles.length === 0 ? (
        <div className="text-slate-600 text-xs py-6 text-center">
          No news articles found for this stock.
        </div>
      ) : (
        <div className="space-y-2">
          {articles.map((a, i) => (
            <a key={i} href={a.url} target="_blank" rel="noopener noreferrer"
              className="flex gap-3 p-3 rounded-lg bg-brand hover:bg-slate-800 border border-brand-border transition-colors block">
              {/* Thumbnail */}
              {a.thumbnail && (
                <img src={a.thumbnail} alt="" className="w-16 h-12 object-cover rounded flex-shrink-0" onError={e => { (e.target as HTMLImageElement).style.display = 'none' }} />
              )}
              <div className="flex-1 min-w-0">
                <div className="flex items-start gap-2 mb-1">
                  <span className={`text-xs px-1.5 py-0.5 rounded border flex-shrink-0 font-semibold uppercase tracking-wider ${badge(a.importance)}`}>
                    {a.importance || 'news'}
                  </span>
                </div>
                <div className="text-sm text-slate-100 line-clamp-2 leading-snug">{a.title}</div>
                {a.summary && (
                  <div className="text-xs text-slate-500 mt-1 line-clamp-2">{a.summary}</div>
                )}
                <div className="text-xs text-slate-600 mt-1.5">
                  {a.source}
                  {a.published_at && (
                    <> · {(() => { try { return formatDistanceToNow(new Date(a.published_at), { addSuffix: true }) } catch { return '' } })()}</>
                  )}
                </div>
              </div>
            </a>
          ))}
        </div>
      )}
    </Card>
  )
}

// ── Main Component ───────────────────────────────────────────────────────────
export default function StockDetailPage() {
  const { symbol } = useParams<{ symbol: string }>()
  const navigate = useNavigate()

  const [summary, setSummary]     = useState<any>(null)
  const [news, setNews]           = useState<any[]>([])
  const [oi, setOi]               = useState<any>(null)
  const [timeframe, setTimeframe] = useState('1D')
  const [chartLoading, setChartLoading] = useState(false)
  const [error, setError]         = useState<string | null>(null)

  const chartRef     = useRef<HTMLDivElement>(null)
  const chartApiRef  = useRef<IChartApi | null>(null)
  const candleRef    = useRef<ISeriesApi<'Candlestick'> | null>(null)
  const volumeRef    = useRef<ISeriesApi<'Histogram'> | null>(null)

  // Load summary + news + OI once on mount
  useEffect(() => {
    if (!symbol) return
    setError(null)

    stockApi.summary(symbol)
      .then(setSummary)
      .catch(() => setError('Failed to load stock data'))

    stockApi.news(symbol, 7)
      .then(d => setNews(d.articles || []))
      .catch(() => setNews([]))

    stockApi.oi(symbol)
      .then(setOi)
      .catch(() => setOi({ available: false }))
  }, [symbol])

  // Build / update chart whenever symbol or timeframe changes
  useEffect(() => {
    if (!chartRef.current || !symbol) return
    setChartLoading(true)

    // Init chart on first render
    if (!chartApiRef.current) {
      const chart = createChart(chartRef.current, {
        layout: {
          background: { type: ColorType.Solid, color: '#0f172a' },
          textColor: '#94a3b8',
        },
        grid: {
          vertLines: { color: '#1e293b' },
          horzLines: { color: '#1e293b' },
        },
        crosshair: { mode: 1 },
        rightPriceScale: { borderColor: '#334155' },
        timeScale: { borderColor: '#334155', timeVisible: true, secondsVisible: false },
        width: chartRef.current.clientWidth,
        height: 360,
      })

      const candle = chart.addCandlestickSeries({
        upColor:        '#22c55e',
        downColor:      '#ef4444',
        borderUpColor:  '#22c55e',
        borderDownColor:'#ef4444',
        wickUpColor:    '#22c55e',
        wickDownColor:  '#ef4444',
      })

      const vol = chart.addHistogramSeries({
        priceFormat: { type: 'volume' },
        priceScaleId: 'vol',
        color: '#334155',
      })
      chart.priceScale('vol').applyOptions({
        scaleMargins: { top: 0.85, bottom: 0 },
      })

      chartApiRef.current = chart
      candleRef.current   = candle
      volumeRef.current   = vol

      const ro = new ResizeObserver(() => {
        if (chartRef.current && chartApiRef.current)
          chartApiRef.current.applyOptions({ width: chartRef.current.clientWidth })
      })
      ro.observe(chartRef.current)
    }

    // Fetch OHLCV for chosen timeframe
    stockApi.ohlcv(symbol, timeframe).then(data => {
      const raw = data.candles || []
      const isIntraday = ['1m','5m','15m','1h','4h'].includes(timeframe)

      const candles: CandlestickData[] = raw
        .filter((c: any) => c.open && c.close)
        .map((c: any) => ({
          time: (isIntraday
            ? Math.floor(new Date(c.date).getTime() / 1000)
            : c.date) as Time,
          open:  c.open,
          high:  c.high,
          low:   c.low,
          close: c.close,
        }))

      const volumes = raw
        .filter((c: any) => c.volume != null)
        .map((c: any) => ({
          time: (isIntraday
            ? Math.floor(new Date(c.date).getTime() / 1000)
            : c.date) as Time,
          value: c.volume,
          color: c.close >= c.open ? '#22c55e33' : '#ef444433',
        }))

      candleRef.current?.setData(candles)
      volumeRef.current?.setData(volumes)
      chartApiRef.current?.timeScale().fitContent()
      setChartLoading(false)
    }).catch(() => setChartLoading(false))

    return () => {
      // Do NOT destroy chart here — only update data
    }
  }, [symbol, timeframe])

  // Cleanup chart on page unmount
  useEffect(() => {
    return () => {
      chartApiRef.current?.remove()
      chartApiRef.current = null
      candleRef.current   = null
      volumeRef.current   = null
    }
  }, [symbol])

  if (!symbol) return null

  const s = summary
  const up   = (v: number | null) => v == null ? '' : v >= 0 ? 'text-up' : 'text-down'
  const rsiColor = (v: number | null) =>
    v == null ? '' : v > 70 ? 'text-down' : v < 30 ? 'text-up' : 'text-yellow-400'

  return (
    <div className="flex flex-col h-full overflow-y-auto bg-brand">
      {/* ── Top bar ── */}
      <div className="sticky top-0 z-20 flex items-center gap-3 px-4 py-2.5 bg-brand-card border-b border-brand-border">
        <button className="btn-ghost text-xs" onClick={() => navigate(-1)}>← Back</button>
        <span className="text-sm font-bold text-white">{s?.symbol ?? symbol}</span>
        {s?.exchange && (
          <span className={`text-xs px-2 py-0.5 rounded font-mono ${s.exchange === 'NSE' ? 'bg-blue-900 text-blue-300' : 'bg-purple-900 text-purple-300'}`}>
            {s.exchange}
          </span>
        )}
        {s?.sector && (
          <span className="text-xs bg-slate-800 border border-brand-border px-2 py-0.5 rounded text-slate-300">
            {s.sector}
          </span>
        )}
        <button
          className="ml-auto mr-3 px-3 py-1.5 bg-brand-accent/20 hover:bg-brand-accent/40 text-brand-accent text-xs rounded border border-brand-accent/30 font-semibold transition-colors"
          onClick={() => navigate(`/stock/${symbol}/valuation`)}
        >
          📊 Valuation Models
        </button>
        <div className="text-right">
          <span className="text-xl font-mono font-bold text-white">{fmtPrice(s?.last_price)}</span>
          {s?.change_pct != null && (
            <span className={`ml-2 text-sm font-mono ${up(s.change_pct)}`}>
              {s.change_pct >= 0 ? '+' : ''}{fmt(s.change_pct)}%
            </span>
          )}
        </div>
      </div>

      {error && (
        <div className="mx-4 mt-4 p-3 bg-red-900/30 border border-red-700 rounded text-red-300 text-sm">{error}</div>
      )}

      <div className="p-4 space-y-4">

        {/* ── Price strip ── */}
        {s && (
          <div className="grid grid-cols-5 gap-2">
            {[
              ['Open',    fmtPrice(s.open)],
              ['High',    fmtPrice(s.day_high),    s.day_high > s.last_price ? 'text-up' : ''],
              ['Low',     fmtPrice(s.day_low),     s.day_low < s.last_price ? 'text-down' : ''],
              ['Prev Close', fmtPrice(s.prev_close)],
              ['Volume',  formatVolume(s.volume)],
            ].map(([label, val, color]) => (
              <div key={label as string} className="card text-center py-2">
                <div className="text-xs text-slate-500 mb-0.5">{label}</div>
                <div className={`text-sm font-mono font-medium ${color || 'text-white'}`}>{val}</div>
              </div>
            ))}
          </div>
        )}

        {/* ── Chart ── */}
        <Card>
          {/* Timeframe selector */}
          <div className="flex items-center gap-1 flex-wrap mb-3">
            {['intraday', 'swing', 'long'].map(group => (
              <div key={group} className="flex gap-1 mr-2">
                {TIMEFRAMES.filter(t => t.group === group).map(tf => (
                  <button
                    key={tf.key}
                    onClick={() => setTimeframe(tf.key)}
                    className={`text-xs px-2.5 py-1 rounded transition-all font-mono ${
                      timeframe === tf.key
                        ? 'bg-brand-accent text-white shadow-sm'
                        : 'text-slate-400 hover:text-white hover:bg-slate-700'
                    }`}
                  >
                    {tf.label}
                  </button>
                ))}
                <div className="w-px bg-brand-border mx-1" />
              </div>
            ))}
            {chartLoading && <span className="text-xs text-slate-500 ml-2">Loading…</span>}
          </div>
          <div ref={chartRef} className="w-full" />
        </Card>

        {/* ── Three-column metrics ── */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">

          {/* Fundamentals */}
          <Card title="Fundamentals">
            <MetricRow label="Market Cap"       value={formatMarketCap(s?.market_cap)} />
            <MetricRow label="PE Ratio"         value={fmt(s?.pe_ratio)} />
            <MetricRow label="Forward PE"       value={fmt(s?.forward_pe)} />
            <MetricRow label="EPS"              value={s?.eps ? `₹${fmt(s.eps)}` : '—'} />
            <MetricRow label="Book Value"       value={s?.book_value ? `₹${fmt(s.book_value)}` : '—'} />
            <MetricRow label="Price / Book"     value={fmt(s?.price_to_book)} />
            <MetricRow label="Dividend Yield"   value={s?.dividend_yield ? `${(s.dividend_yield * 100).toFixed(2)}%` : '—'} />
            <MetricRow label="Beta"             value={fmt(s?.beta)} />
            <MetricRow label="52W High"         value={fmtPrice(s?.['52w_high'] ?? s?.year_high)} />
            <MetricRow label="52W Low"          value={fmtPrice(s?.['52w_low'] ?? s?.year_low)} />
          </Card>

          {/* Technical */}
          <Card title="Technical Indicators">
            <MetricRow label="RSI (14)"         value={fmt(s?.rsi_14, 1)}      color={rsiColor(s?.rsi_14)} />
            <MetricRow label="MACD Signal"      value={s?.macd != null && s?.macd_signal != null ? (s.macd > s.macd_signal ? '🟢 Bullish' : '🔴 Bearish') : '—'} />
            <MetricRow label="MACD"             value={fmt(s?.macd, 3)} />
            <MetricRow label="Signal Line"      value={fmt(s?.macd_signal, 3)} />
            <MetricRow label="SMA 20"           value={fmtPrice(s?.sma_20)}    color={s?.last_price > s?.sma_20 ? 'text-up' : 'text-down'} />
            <MetricRow label="SMA 50"           value={fmtPrice(s?.sma_50)}    color={s?.last_price > s?.sma_50 ? 'text-up' : 'text-down'} />
            <MetricRow label="SMA 200"          value={fmtPrice(s?.sma_200)}   color={s?.last_price > s?.sma_200 ? 'text-up' : 'text-down'} />
            <MetricRow label="EMA 20"           value={fmtPrice(s?.ema_20)} />
            <MetricRow label="Max Drawdown 52W" value={formatPct(s?.max_drawdown_52w)} color={up(s?.max_drawdown_52w)} />
            <MetricRow label="Avg Volume (20D)" value={formatVolume(s?.avg_volume_20d)} />
          </Card>

          {/* Returns */}
          <Card title="Performance">
            <MetricRow label="Today"            value={formatPct(s?.daily_return ?? s?.change_pct)}  color={up(s?.daily_return ?? s?.change_pct)} />
            <MetricRow label="5 Days"           value={formatPct(s?.return_5d)}   color={up(s?.return_5d)} />
            <MetricRow label="1 Month"          value={formatPct(s?.return_1m)}   color={up(s?.return_1m)} />
            <MetricRow label="3 Months"         value={formatPct(s?.return_3m)}   color={up(s?.return_3m)} />
            <MetricRow label="1 Year"           value={formatPct(s?.return_1y)}   color={up(s?.return_1y)} />
            <div className="mt-3 pt-2 border-t border-brand-border">
              <div className="text-xs text-slate-500 mb-2">Financials</div>
              <MetricRow label="ROE"              value={s?.return_on_equity ? `${(s.return_on_equity * 100).toFixed(1)}%` : '—'} />
              <MetricRow label="ROA"              value={s?.return_on_assets ? `${(s.return_on_assets * 100).toFixed(1)}%` : '—'} />
              <MetricRow label="Debt / Equity"    value={fmt(s?.debt_to_equity)} />
              <MetricRow label="Profit Margin"    value={s?.profit_margins ? `${(s.profit_margins * 100).toFixed(1)}%` : '—'} />
              <MetricRow label="Rev. Growth"      value={s?.revenue_growth ? `${(s.revenue_growth * 100).toFixed(1)}%` : '—'} />
            </div>
          </Card>
        </div>

        {/* ── OI + News side by side ── */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <OIPanel oi={oi} />
          <NewsPanel articles={news} />
        </div>

      </div>
    </div>
  )
}
