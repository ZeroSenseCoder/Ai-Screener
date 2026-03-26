import { useEffect, useState } from 'react'
import { macroApi } from '../../services/api'
import { formatPct, pctColor } from '../../utils/formatters'

interface IndexItem { name: string; price: number | null; change_pct: number | null }
interface FiiDii { date: string; fii_net: number; dii_net: number }

export default function MacroDashboard() {
  const [data, setData] = useState<any>(null)

  useEffect(() => {
    macroApi.overview().then(setData).catch(() => {})
    const id = setInterval(() => macroApi.overview().then(setData).catch(() => {}), 60_000)
    return () => clearInterval(id)
  }, [])

  if (!data) return (
    <div className="bg-brand-card border-b border-brand-border px-4 py-2 text-slate-500 text-xs">
      Loading macro data...
    </div>
  )

  const indices: IndexItem[] = data.global_indices || []
  const forex: IndexItem[] = data.forex || []
  const commodities: IndexItem[] = data.commodities || []
  const vix = data.india_vix
  const fiiDii: FiiDii[] = data.fii_dii || []
  const latestFii = fiiDii[0]

  return (
    <div className="bg-brand-card border-b border-brand-border overflow-x-auto">
      <div className="flex items-center gap-6 px-4 py-2 min-w-max text-xs">
        {/* Indices */}
        {indices.map(idx => (
          <div key={idx.name} className="flex flex-col">
            <span className="text-slate-500">{idx.name}</span>
            <span className={`font-semibold ${pctColor(idx.change_pct)}`}>
              {idx.price?.toLocaleString('en-IN') ?? '—'}{' '}
              <span className="text-xs">({formatPct(idx.change_pct)})</span>
            </span>
          </div>
        ))}

        <div className="w-px h-8 bg-brand-border mx-1" />

        {/* Forex */}
        {forex.slice(0, 2).map(f => (
          <div key={f.name} className="flex flex-col">
            <span className="text-slate-500">{f.name}</span>
            <span className={`font-semibold ${pctColor(f.change_pct)}`}>
              {f.price?.toFixed(2) ?? '—'}
            </span>
          </div>
        ))}

        <div className="w-px h-8 bg-brand-border mx-1" />

        {/* Commodities */}
        {commodities.slice(0, 2).map(c => (
          <div key={c.name} className="flex flex-col">
            <span className="text-slate-500">{c.name}</span>
            <span className={`font-semibold ${pctColor(c.change_pct)}`}>
              ${c.price?.toFixed(2) ?? '—'} <span className="text-xs">({formatPct(c.change_pct)})</span>
            </span>
          </div>
        ))}

        <div className="w-px h-8 bg-brand-border mx-1" />

        {/* India VIX */}
        {vix && (
          <div className="flex flex-col">
            <span className="text-slate-500">India VIX</span>
            <span className="font-semibold text-yellow-400">{vix.price?.toFixed(2) ?? '—'}</span>
          </div>
        )}

        {/* FII/DII */}
        {latestFii && (
          <>
            <div className="w-px h-8 bg-brand-border mx-1" />
            <div className="flex flex-col">
              <span className="text-slate-500">FII Net ({latestFii.date})</span>
              <span className={`font-semibold ${pctColor(latestFii.fii_net)}`}>
                ₹{(latestFii.fii_net / 100).toFixed(0)}Cr
              </span>
            </div>
            <div className="flex flex-col">
              <span className="text-slate-500">DII Net</span>
              <span className={`font-semibold ${pctColor(latestFii.dii_net)}`}>
                ₹{(latestFii.dii_net / 100).toFixed(0)}Cr
              </span>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
