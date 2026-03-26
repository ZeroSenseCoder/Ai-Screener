/**
 * Numeric / boolean filter controls (PE, RSI, Volume, MA signals, etc.)
 * Shown as a collapsible drawer below the sector nav.
 */
import { useState } from 'react'
import { ScreenerRequest } from '../../services/api'

interface Props {
  filters: ScreenerRequest
  onChange: (f: ScreenerRequest) => void
  onReset: () => void
}

function Section({ title, children, defaultOpen = false }: {
  title: string; children: React.ReactNode; defaultOpen?: boolean
}) {
  const [open, setOpen] = useState(defaultOpen)
  return (
    <div className="border-b border-brand-border">
      <button
        className="flex items-center justify-between w-full px-3 py-2 text-xs text-slate-400 hover:text-slate-200 transition-colors"
        onClick={() => setOpen(o => !o)}
      >
        <span className="font-medium uppercase tracking-wider">{title}</span>
        <span className="text-slate-600">{open ? '▾' : '▸'}</span>
      </button>
      {open && <div className="px-3 pb-3">{children}</div>}
    </div>
  )
}

export default function FilterPanel({ filters, onChange, onReset }: Props) {
  const set = (key: keyof ScreenerRequest, val: any) =>
    onChange({ ...filters, [key]: val, page: 1 })

  const numInput = (label: string, minKey: keyof ScreenerRequest, maxKey: keyof ScreenerRequest) => (
    <div className="mb-2.5">
      <div className="text-slate-500 text-xs mb-1">{label}</div>
      <div className="flex gap-1.5">
        <input type="number" placeholder="Min" className="filter-input"
          value={(filters[minKey] as number) ?? ''}
          onChange={e => set(minKey, e.target.value ? +e.target.value : undefined)} />
        <input type="number" placeholder="Max" className="filter-input"
          value={(filters[maxKey] as number) ?? ''}
          onChange={e => set(maxKey, e.target.value ? +e.target.value : undefined)} />
      </div>
    </div>
  )

  const boolToggle = (label: string, key: keyof ScreenerRequest) => (
    <label key={key} className="flex items-center gap-2 cursor-pointer mb-1.5 group">
      <input type="checkbox" className="accent-blue-500 w-3 h-3"
        checked={!!filters[key]}
        onChange={e => set(key, e.target.checked)} />
      <span className="text-xs text-slate-400 group-hover:text-slate-200">{label}</span>
    </label>
  )

  // Quick preset button
  const preset = (label: string, active: boolean, onClick: () => void) => (
    <button
      key={label}
      onClick={onClick}
      className={`text-xs px-2 py-1 rounded border transition-all ${
        active
          ? 'bg-brand-accent border-brand-accent text-white'
          : 'border-brand-border text-slate-500 hover:border-brand-accent hover:text-brand-accent'
      }`}
    >
      {label}
    </button>
  )

  // Count active filters for the reset badge
  const activeCount = [
    filters.min_pe, filters.max_pe,
    filters.min_market_cap, filters.max_market_cap,
    filters.min_beta, filters.max_beta,
    filters.min_price, filters.max_price,
    filters.min_daily_return, filters.max_daily_return,
    filters.min_return_1m, filters.max_return_1m,
    filters.min_return_3m, (filters as any).max_return_3m,
    filters.min_return_1y, (filters as any).max_return_1y,
    filters.min_rsi, filters.max_rsi,
    filters.min_avg_volume,
    filters.max_drawdown_threshold,
    filters.price_above_sma20, filters.price_above_sma50,
    filters.price_above_sma200, filters.sma50_above_sma200,
    filters.macd_bullish, filters.macd_bearish,
  ].filter(Boolean).length

  return (
    <div className="flex flex-col">
      {/* Header */}
      <div className="flex items-center justify-between px-3 py-2 border-b border-brand-border">
        <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
          Filters {activeCount > 0 && <span className="ml-1 bg-brand-accent text-white px-1.5 py-0.5 rounded-full text-xs">{activeCount}</span>}
        </span>
        {activeCount > 0 && (
          <button className="text-xs text-slate-500 hover:text-red-400" onClick={onReset}>
            Reset
          </button>
        )}
      </div>

      {/* Exchange */}
      <div className="px-3 py-2 border-b border-brand-border">
        <div className="text-xs text-slate-500 mb-1.5">Exchange</div>
        <div className="flex gap-1.5">
          {['NSE', 'BSE'].map(ex => {
            const active = (filters.exchanges || ['NSE', 'BSE']).includes(ex)
            const toggle = () => {
              const curr = filters.exchanges || ['NSE', 'BSE']
              const next = curr.includes(ex) ? curr.filter(e => e !== ex) : [...curr, ex]
              onChange({ ...filters, exchanges: next, page: 1 })
            }
            return (
              <button key={ex} onClick={toggle}
                className={`flex-1 py-1.5 rounded border text-xs font-semibold transition-all ${
                  active
                    ? ex === 'NSE' ? 'bg-blue-900 border-blue-600 text-blue-200'
                                   : 'bg-purple-900 border-purple-600 text-purple-200'
                    : 'bg-brand border-brand-border text-slate-500 hover:border-slate-400'
                }`}>
                {ex}
              </button>
            )
          })}
        </div>
      </div>

      {/* Fundamentals */}
      <Section title="Fundamentals" defaultOpen={true}>
        {numInput('PE Ratio', 'min_pe', 'max_pe')}
        {numInput('Market Cap (Cr)', 'min_market_cap', 'max_market_cap')}
        <div className="flex flex-wrap gap-1 mb-2">
          {[
            { label: 'Small Cap', min: undefined, max: 500 },
            { label: 'Mid Cap',   min: 500,  max: 20000 },
            { label: 'Large Cap', min: 20000, max: undefined },
          ].map(p => preset(p.label,
            filters.min_market_cap === p.min && filters.max_market_cap === p.max,
            () => onChange({ ...filters, min_market_cap: p.min, max_market_cap: p.max, page: 1 })
          ))}
        </div>
        {numInput('Beta', 'min_beta', 'max_beta')}
      </Section>

      {/* Returns */}
      <Section title="Returns (%)">
        {numInput('Daily Return %', 'min_daily_return', 'max_daily_return')}
        {numInput('1M Return %', 'min_return_1m', 'max_return_1m')}
        {numInput('3M Return %', 'min_return_3m', 'max_return_3m')}
        {numInput('1Y Return %', 'min_return_1y', 'max_return_1y')}
      </Section>

      {/* RSI */}
      <Section title="RSI (14)">
        {numInput('RSI Range', 'min_rsi', 'max_rsi')}
        <div className="flex gap-1 flex-wrap mt-1">
          {preset('Oversold <30',   filters.max_rsi === 30 && !filters.min_rsi,      () => onChange({ ...filters, min_rsi: undefined, max_rsi: 30,  page: 1 }))}
          {preset('Neutral 30-70', filters.min_rsi === 30 && filters.max_rsi === 70, () => onChange({ ...filters, min_rsi: 30,         max_rsi: 70,  page: 1 }))}
          {preset('Overbought >70', filters.min_rsi === 70 && !filters.max_rsi,      () => onChange({ ...filters, min_rsi: 70,         max_rsi: undefined, page: 1 }))}
        </div>
      </Section>

      {/* Volume */}
      <Section title="Volume">
        <div className="text-slate-500 text-xs mb-1">Min Avg Volume (20D)</div>
        <input type="number" placeholder="e.g. 100000" className="filter-input mb-2"
          value={filters.min_avg_volume ?? ''}
          onChange={e => set('min_avg_volume', e.target.value ? +e.target.value : undefined)} />
        <div className="flex flex-wrap gap-1">
          {[
            { label: '>1L',   val: 100_000 },
            { label: '>5L',   val: 500_000 },
            { label: '>10L',  val: 1_000_000 },
            { label: '>50L',  val: 5_000_000 },
          ].map(p => preset(p.label, filters.min_avg_volume === p.val,
            () => set('min_avg_volume', filters.min_avg_volume === p.val ? undefined : p.val)
          ))}
        </div>
      </Section>

      {/* Max Drawdown */}
      <Section title="Max Drawdown (52W)">
        <div className="text-slate-500 text-xs mb-1">Loss worse than (%)</div>
        <input type="number" placeholder="e.g. -30" className="filter-input mb-2"
          value={filters.max_drawdown_threshold ?? ''}
          onChange={e => set('max_drawdown_threshold', e.target.value ? +e.target.value : undefined)} />
        <div className="flex gap-1 flex-wrap">
          {[{ label: '< -10%', val: -10 }, { label: '< -20%', val: -20 }, { label: '< -40%', val: -40 }].map(p => (
            <button key={p.label}
              className={`text-xs px-2 py-1 rounded border transition-all ${
                filters.max_drawdown_threshold === p.val
                  ? 'bg-red-900 border-red-700 text-red-300'
                  : 'border-brand-border text-slate-500 hover:border-red-700 hover:text-red-400'
              }`}
              onClick={() => set('max_drawdown_threshold', filters.max_drawdown_threshold === p.val ? undefined : p.val)}
            >
              {p.label}
            </button>
          ))}
        </div>
      </Section>

      {/* Price */}
      <Section title="Price (₹)">
        {numInput('Price Range', 'min_price', 'max_price')}
      </Section>

      {/* MA Signals */}
      <Section title="MA Signals">
        {boolToggle('Price > SMA 20', 'price_above_sma20')}
        {boolToggle('Price > SMA 50', 'price_above_sma50')}
        {boolToggle('Price > SMA 200 (uptrend)', 'price_above_sma200')}
        {boolToggle('SMA 50 > SMA 200 (golden zone)', 'sma50_above_sma200')}
      </Section>

      {/* MACD */}
      <Section title="MACD">
        {boolToggle('MACD Bullish (MACD > Signal)', 'macd_bullish')}
        {boolToggle('MACD Bearish', 'macd_bearish')}
      </Section>

      {/* Reset button at bottom */}
      <div className="px-3 py-3">
        <button className="w-full btn-ghost text-xs" onClick={onReset}>
          Reset All Filters
        </button>
      </div>
    </div>
  )
}
